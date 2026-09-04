"""apply_edits live route: the whole batch inside ONE COM undo group
(V2_DESIGN 9.4 binding 2: one live session, one run_live call, so the
entire batch is a single Ctrl+Z step).

Flow: the server validated the batch against a DISK SNAPSHOT of the locked
file (ops/batch.validate_edits), producing index-based plans; this module
pre-checks live support for every op (raising BEFORE Word is touched),
then executes the ops in order against the open document, remapping
paragraph indices as inserts and deletes shift them.

Live-route limits, stated honestly rather than approximated:
- Anchors resolve against the last SAVED state (the Wave E caveat: unsaved
  changes in Word are invisible to the resolver). The replace op verifies
  the matched text against Word's own range before writing; every other
  index-addressed op runs the snapshot staleness guard (live_ops
  _stale_guard): when the open document is dirty, the live target must
  still match the snapshot text its anchor resolved to, else STALE_ANCHOR.
- Markdown lists and pipe tables in insert ops are file-mode only; the
  batch refuses up front with the close-the-file hint.

Atomicity (2026-09-03 stress report bug 2 - partial COM writes leaked
through refused batches): ALL validation now completes before ANY COM
write. Three layers, all inside the one serialized live session:
1. _preflight_conflicts: pure-Python simulation of the batch's index
   remapping, catching target-deleted-by-earlier-edit up front;
2. _preflight_live: every op's target checked against the LIVE document
   (bounds, staleness guards, replace-find presence, style existence,
   cell addressing, tracked-revision refusals) with zero writes;
3. rollback: a mid-execution failure (the residual: Word-state changes
   the preflight cannot foresee) reverts the applied portion via the
   batch's single undo group when grouping is active; when it is not,
   the error says PARTIALLY APPLIED honestly.

Bodies are composed from live_ops' session-level body factories (the same
code the standalone live tools run), plus two small local bodies for the
replace and set_style ops which have no standalone live tool.
"""

from __future__ import annotations

import re

from ..core.errors import (
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from . import live_ops as _lo
from .live import check_text_safe, run_live

_HEADING_RE = re.compile(r"[Hh]eading\s?([1-9])$")


# ------------------------------------------------------------ local bodies


def _replace_body(index: int, find: str, new: str, occurrence: int | None):
    def body(session):
        doc = session.doc
        paras = _lo._body_paragraphs(doc)
        if not 0 <= index < len(paras):
            raise TargetNotFound(
                f"paragraph index {index} out of range "
                f"({len(paras)} body paragraphs)"
            )
        p = paras[index]
        text = _lo._para_text(p)
        spans = []
        pos = 0
        while True:
            pos = text.find(find, pos)
            if pos < 0:
                break
            spans.append((pos, pos + len(find)))
            pos += len(find)
        if not spans:
            raise TargetNotFound(
                f"replace: {find!r} not found in the target paragraph "
                "(the open document may differ from the saved file the "
                "anchors were resolved against)"
            )
        if occurrence is not None:
            if occurrence > len(spans):
                raise TargetNotFound(
                    f"replace: occurrence {occurrence} out of range, "
                    f"{len(spans)} match(es)"
                )
            spans = [spans[occurrence - 1]]
        base = p.Range.Start
        for lo_, hi_ in reversed(spans):
            rng = doc.Range(base + lo_, base + hi_)
            if rng.Text != find:
                raise UnsupportedStructure(
                    "character offsets do not line up with Word's "
                    "positions in this paragraph (complex fields present); "
                    "close the document and use the file-based route"
                )
            _lo._assign_text(rng, new)
        return {"replaced": len(spans), "paragraph": index}

    return body


def _set_style_body(index: int, style: str, verify_text: str | None = None):
    def body(session):
        doc = session.doc
        paras = _lo._body_paragraphs(doc)
        if not 0 <= index < len(paras):
            raise TargetNotFound(
                f"paragraph index {index} out of range "
                f"({len(paras)} body paragraphs)"
            )
        _lo._stale_guard(session, paras, index, verify_text, "set_style")
        rng = doc.Range(paras[index].Range.Start, paras[index].Range.End)
        tried = [style]
        m = _HEADING_RE.fullmatch(style)
        if m and " " not in style:
            tried.append(f"Heading {m.group(1)}")  # id form -> display name
        for cand in tried:
            try:
                rng.Style = cand
                return {"styled": 1, "paragraph": index, "style": cand}
            except Exception:
                continue
        raise TargetNotFound(
            f"style {style!r} does not exist in this document"
        )

    return body


# ---------------------------------------------------------------- specs


def _live_spec(edit: dict, plan: dict, i: int) -> dict:
    """Pre-check one op's live support and pre-validate its params, BEFORE
    Word is touched. Factories from live_ops are pure until their body
    runs, so calling them here (with a placeholder index where needed)
    runs their parameter validation and nothing else."""
    op = plan["op"]
    spec: dict = {"op": op, "index": plan.get("index")}
    if op == "insert":
        items: list[dict] = []
        for seg in plan["segments"]:
            if seg["kind"] != "paragraphs":
                raise UnsupportedStructure(
                    f"edit {i} (insert): markdown "
                    f"{'lists' if seg['kind'] == 'list' else 'tables'} are "
                    "file-mode only; close the document in Word and resend "
                    "the batch. Nothing was applied."
                )
            for item in seg["items"]:
                spec_item = {"text": item["text"]}
                if "level" in item:
                    # heading_level applies the built-in style by its
                    # numeric constant (locale-independent)
                    spec_item["heading_level"] = item["level"]
                items.append(spec_item)
        spec["items"] = items
        spec["mode"] = plan["mode"]
        for item in items:
            check_text_safe(item["text"])
    elif op == "delete":
        spec["indices"] = plan["indices"]
    elif op == "set_cell":
        spec.update(
            {"table_index": plan["table_index"], "row": plan["row"],
             "col": plan["col"], "text": edit["text"]}
        )
        check_text_safe(edit["text"])
    elif op == "replace":
        check_text_safe(edit["text"])
        spec.update(
            {"find": edit["find"], "text": edit["text"],
             "occurrence": edit.get("occurrence")}
        )
    elif op == "set_text":
        _lo.replace_paragraph_text_body(0, edit["text"])  # validate only
        spec["text"] = edit["text"]
    elif op == "set_style":
        spec["style"] = edit["style"]
    elif op == "format":
        _lo.format_text_body(  # validate only (live key set differs)
            edit["formatting"], paragraph_index=0,
            find=edit.get("find"), occurrence=edit.get("occurrence") or 1,
        )
        spec.update(
            {"formatting": edit["formatting"], "find": edit.get("find"),
             "occurrence": edit.get("occurrence") or 1}
        )
    elif op == "set_paragraph_format":
        _lo.set_paragraph_format_body([0], edit["format"])  # validate only
        spec["format"] = edit["format"]
    return spec


# ------------------------------------------------------------- preflight


def _preflight_conflicts(specs: list[dict], n_paras: int) -> None:
    """Pure-Python simulation of the batch's index remapping: catches a
    later op targeting a paragraph an earlier op deletes, BEFORE any COM
    write. Mirrors the execution loop's pos[] arithmetic."""
    pos: list = list(range(n_paras))

    def cur(o, i):
        if o is None:
            return None
        c = pos[o] if 0 <= o < len(pos) else None
        if c is None and 0 <= o < len(pos):
            raise WordMcpError(
                f"edit {i}: its target paragraph is deleted by an earlier "
                "edit in this batch. Nothing was applied - keep deletes "
                "last or in their own batch."
            )
        return c

    for i, spec in enumerate(specs):
        op = spec["op"]
        if op == "insert":
            k = len(spec["items"])
            mode = spec["mode"]
            if mode == "start":
                pos = [None if c is None else c + k for c in pos]
            elif mode in ("after", "before"):
                c = cur(spec["index"], i)
                if c is None:
                    continue
                if mode == "after":
                    pos = [None if x is None else x + k if x > c else x
                           for x in pos]
                else:
                    pos = [None if x is None else x + k if x >= c else x
                           for x in pos]
        elif op == "delete":
            cs = sorted(
                c for c in (cur(o, i) for o in spec["indices"])
                if c is not None
            )
            for c in reversed(cs):
                pos = [
                    None if x == c
                    else (x - 1 if x is not None and x > c else x)
                    for x in pos
                ]
        else:
            cur(spec.get("index"), i)


def _preflight_live(session, specs: list[dict],
                    snap_texts: list[str] | None) -> None:
    """Validate every op against the LIVE document's batch-start state -
    bounds, staleness, find presence, style existence, cell addressing,
    revision refusals - with ZERO writes. A failure here means nothing
    was applied, guaranteed."""
    doc = session.doc
    paras = _lo._body_paragraphs(doc)

    def vt(o):
        if snap_texts is None or o is None or not 0 <= o < len(snap_texts):
            return None
        return snap_texts[o]

    def check_index(o, i, op):
        if not 0 <= o < len(paras):
            raise TargetNotFound(
                f"edit {i} ({op}): paragraph index {o} out of range "
                f"({len(paras)} body paragraphs). Nothing was applied."
            )
        _lo._stale_guard(session, paras, o, vt(o), f"apply_edits edit {i}")

    for i, spec in enumerate(specs):
        op = spec["op"]
        if op == "insert":
            if spec["mode"] in ("after", "before"):
                check_index(spec["index"], i, op)
        elif op == "delete":
            for o in spec["indices"]:
                check_index(o, i, op)
        elif op == "set_cell":
            t = spec["table_index"]
            if not 0 <= t < doc.Tables.Count:
                raise TargetNotFound(
                    f"edit {i} (set_cell): no table with index {t} "
                    f"({doc.Tables.Count} tables). Nothing was applied."
                )
            table = doc.Tables(t + 1)
            r, c = spec["row"], spec["col"]
            if not 0 <= r < table.Rows.Count:
                raise TargetNotFound(
                    f"edit {i} (set_cell): row {r} out of range (table "
                    f"has {table.Rows.Count}). Nothing was applied."
                )
            try:
                cell_count = table.Rows(r + 1).Cells.Count
            except Exception as exc:
                raise UnsupportedStructure(
                    f"edit {i} (set_cell): this table has vertically "
                    "merged cells, which Word's live row addressing "
                    "cannot handle - use file-based set_cells on the "
                    "closed file. Nothing was applied."
                ) from exc
            if not 0 <= c < cell_count:
                raise TargetNotFound(
                    f"edit {i} (set_cell): cell {c} out of range (row {r} "
                    f"has {cell_count}). Nothing was applied."
                )
        elif op == "replace":
            check_index(spec["index"], i, op)
            text = _lo._para_text(paras[spec["index"]])
            if spec["find"] not in text:
                raise TargetNotFound(
                    f"edit {i} (replace): {spec['find']!r} not found in "
                    "the target paragraph (the open document may differ "
                    "from the saved file the anchors were resolved "
                    "against). Nothing was applied."
                )
        elif op == "set_text":
            check_index(spec["index"], i, op)
            p = paras[spec["index"]]
            rng = doc.Range(
                p.Range.Start, max(p.Range.Start, p.Range.End - 1)
            )
            try:
                has_rev = bool(rng.Revisions.Count)
            except Exception as exc:
                raise WordMcpError(
                    f"edit {i} (set_text): Word did not answer while "
                    "checking tracked revisions on the target; refusing "
                    "fail-closed. Nothing was applied - retry."
                ) from exc
            if has_rev:
                raise UnsupportedStructure(
                    f"edit {i} (set_text): the target paragraph carries "
                    "tracked revisions; accept or reject them first. "
                    "Nothing was applied."
                )
        elif op == "set_style":
            check_index(spec["index"], i, op)
            style = spec["style"]
            tried = [style]
            m = _HEADING_RE.fullmatch(style)
            if m and " " not in style:
                tried.append(f"Heading {m.group(1)}")
            for cand in tried:
                try:
                    doc.Styles(cand)  # read-only existence probe
                    break
                except Exception:
                    continue
            else:
                raise TargetNotFound(
                    f"edit {i} (set_style): style {style!r} does not "
                    "exist in this document. Nothing was applied."
                )
        elif op in ("format", "set_paragraph_format"):
            check_index(spec["index"], i, op)


def _rollback(session) -> str:
    """Revert a part-applied batch via its single undo group. Returns the
    honest final-state note appended to the raised error."""
    if not session.undo_grouped:
        return (
            "PARTIALLY APPLIED: the batch was interrupted and undo "
            "grouping was unavailable (the document is not Word's active "
            "document) - review and undo the applied edits in Word."
        )
    try:
        undo = session.app.UndoRecord
        if undo.IsRecordingCustomRecord:
            undo.EndCustomRecord()
        session.doc.Undo()
        return (
            "The applied portion was ROLLED BACK via the batch undo "
            "group; nothing from this batch remains in the document."
        )
    except Exception:
        return (
            "PARTIALLY APPLIED: automatic rollback failed - one Ctrl+Z "
            "in Word removes the applied portion (single undo group)."
        )


# ----------------------------------------------------------------- driver


def apply_edits_live(
    path: str, edits: list[dict], plans: list[dict], n_paras: int,
    snap_texts: list[str] | None = None,
) -> dict:
    """Execute a pre-validated batch against the OPEN document in one undo
    group. plans come from ops/batch.validate_edits on the disk snapshot;
    n_paras is that snapshot's body paragraph count (the index space the
    plans speak). snap_texts (parallel to that index space) enables the
    staleness guard: when the open document has unsaved changes, each
    index-addressed op verifies its live target still matches the snapshot
    text its anchor resolved to, refusing (STALE_ANCHOR) on mismatch."""
    specs = [_live_spec(e, p, i) for i, (e, p) in enumerate(zip(edits, plans))]

    def body(session):
        # ALL validation before ANY COM write (stress report bug 2):
        _preflight_conflicts(specs, n_paras)
        _preflight_live(session, specs, snap_texts)
        # snapshot index -> current index (None = deleted by this batch)
        pos: list = list(range(n_paras))
        # original indices whose TEXT this batch already rewrote: their
        # snapshot text is legitimately stale, skip further verification
        rewritten: set[int] = set()

        def vt(o: int | None) -> str | None:
            """snapshot verify-text for original index o, or None."""
            if (snap_texts is None or o is None
                    or o in rewritten or not 0 <= o < len(snap_texts)):
                return None
            return snap_texts[o]

        def cur(o: int, i: int) -> int:
            c = pos[o] if 0 <= o < len(pos) else None
            if c is None:
                raise WordMcpError(
                    f"edit {i}: its target paragraph was deleted by an "
                    "earlier edit in this batch; the batch stops here "
                    "(one Ctrl+Z in Word undoes the applied portion). "
                    "Keep deletes last or in their own batch."
                )
            return c

        changed: dict[str, dict] = {}
        try:
            for i, spec in enumerate(specs):
                op = spec["op"]
                if op == "insert":
                    k = len(spec["items"])
                    mode = spec["mode"]
                    if mode == "end":
                        b = _lo.insert_paragraphs_body(spec["items"], at_end=True)
                        b(session)
                    elif mode == "start":
                        b = _lo.insert_paragraphs_body(
                            spec["items"], before_index=0
                        )
                        b(session)
                        pos = [None if c is None else c + k for c in pos]
                    else:
                        c = cur(spec["index"], i)
                        if mode == "after":
                            _lo.insert_paragraphs_body(
                                spec["items"], after_index=c,
                                verify_text=vt(spec["index"]),
                            )(session)
                            pos = [None if x is None else x + k if x > c else x
                                   for x in pos]
                        else:  # before
                            _lo.insert_paragraphs_body(
                                spec["items"], before_index=c,
                                verify_text=vt(spec["index"]),
                            )(session)
                            pos = [None if x is None else x + k if x >= c else x
                                   for x in pos]
                    changed[str(i)] = {
                        "inserted_paragraphs": k,
                        "note": ("anchors for live-inserted paragraphs appear "
                                 "on the next view after the document is saved"),
                    }
                elif op == "delete":
                    pairs = sorted((cur(o, i), o) for o in spec["indices"])
                    runs: list[list[tuple[int, int]]] = []
                    for c, o in pairs:
                        if runs and c == runs[-1][-1][0] + 1:
                            runs[-1].append((c, o))
                        else:
                            runs.append([(c, o)])
                    deleted = 0
                    for pair_run in reversed(runs):  # bottom-up: stay valid
                        run = [c for c, _o in pair_run]
                        _lo.delete_paragraphs_body(
                            run[0], run[-1],
                            verify_start_text=vt(pair_run[0][1]),
                            verify_end_text=vt(pair_run[-1][1]),
                        )(session)
                        width = run[-1] - run[0] + 1
                        pos = [
                            None if x is None or run[0] <= x <= run[-1]
                            else x - width if x > run[-1] else x
                            for x in pos
                        ]
                        deleted += width
                    changed[str(i)] = {"deleted": deleted}
                elif op == "set_cell":
                    changed[str(i)] = {
                        **_lo.set_cells_body(
                            spec["table_index"],
                            [{"row": spec["row"], "cell": spec["col"],
                              "text": spec["text"]}],
                        )(session),
                        "row": spec["row"], "col": spec["col"],
                    }
                elif op == "replace":
                    # _replace_body re-finds and re-verifies the text in the
                    # LIVE paragraph, so it needs no snapshot guard
                    changed[str(i)] = _replace_body(
                        cur(spec["index"], i), spec["find"], spec["text"],
                        spec["occurrence"],
                    )(session)
                    rewritten.add(spec["index"])
                elif op == "set_text":
                    changed[str(i)] = _lo.replace_paragraph_text_body(
                        cur(spec["index"], i), spec["text"],
                        verify_text=vt(spec["index"]),
                    )(session)
                    rewritten.add(spec["index"])
                elif op == "set_style":
                    changed[str(i)] = _set_style_body(
                        cur(spec["index"], i), spec["style"],
                        verify_text=vt(spec["index"]),
                    )(session)
                elif op == "format":
                    changed[str(i)] = _lo.format_text_body(
                        spec["formatting"],
                        paragraph_index=cur(spec["index"], i),
                        find=spec["find"], occurrence=spec["occurrence"],
                        verify_text=vt(spec["index"]),
                    )(session)
                else:  # set_paragraph_format
                    changed[str(i)] = _lo.set_paragraph_format_body(
                        [cur(spec["index"], i)], spec["format"],
                        verify_texts=[vt(spec["index"])],
                    )(session)
        except Exception as exc:
            # residual mid-execution failure (a Word-state change the
            # preflight could not foresee): revert via the batch's one
            # undo group, then re-raise with the honest final state
            note = _rollback(session)
            typed = exc
            if not isinstance(exc, WordMcpError):
                from .live import _classify
                mapped = _classify(exc)
                if mapped is not None:
                    typed = mapped
                    typed.__cause__ = exc
            if isinstance(typed, WordMcpError):
                typed.args = (f"{typed} {note}",)
            raise typed
        return {"applied": len(specs), "changed": changed, "warnings": []}

    return run_live(path, "apply edits", body)
