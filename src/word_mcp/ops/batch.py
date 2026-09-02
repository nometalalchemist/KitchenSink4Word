"""Batch operations: apply_edits (the v2 anchor-addressed batch layer,
V2_DESIGN Section 9.2) plus the v1 multi-file batch_apply engine it
absorbed.

apply_edits contract:
1. Validate first, mutate second. Every anchor and location in the batch
   resolves against the current package before ANYTHING mutates; any
   resolution failure refuses the WHOLE batch (STALE_ANCHOR when a view
   anchor went stale, else NOT_FOUND/BAD_PARAMS) listing every failed op
   index. The server's _edit wrapper gives the batch one lock, one backup
   rotation, and one validated save; an apply-time failure raises before
   the save, so the on-disk file is untouched either way.
2. Targets are held as ELEMENTS between validation and apply (the Phase 1
   lxml lesson: keep the proxy lists alive, compare with `is`, never key
   by id() over a discarded list), so earlier edits shifting indices never
   mis-target later ones. An op whose target a PRIOR op deleted fails at
   apply time (whole batch abandoned, nothing saved): resolution
   guarantees targets exist at batch START, not internal consistency;
   keep deletes last or in their own batch.
3. insert ops accept markdown (headings, lists, plain paragraphs, pipe
   tables) mapped to real styles; unrepresentable markdown refuses with
   UNSUPPORTED_CONTENT and a pointer to the fine-grained tool.
4. `changed` maps op index to its result; inserted paragraphs are stamped
   with fresh w14:paraId values and their new anchors reported, so a
   follow-up batch chains without re-viewing.

One DocxPackage, one backup, one atomic save per file also governs the v1
batch_apply engine below (unchanged v1 behavior).
"""

from __future__ import annotations

import random
import re as _re
from pathlib import Path

from lxml import etree

from ..core.errors import (
    StaleAnchor,
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from ..core.package import DocxPackage, qn
from ..core.safesave import write_lock
from . import furniture as _fu, structure as _sx, tables as _tb, text as _tx


def _set_header(pkg, p):
    return _fu.set_header_footer(pkg, "header", p.pop("text", ""), **p)


def _set_footer(pkg, p):
    return _fu.set_header_footer(pkg, "footer", p.pop("text", ""), **p)


def _add_watermark(pkg, p):
    return _fu.add_watermark(pkg, p.pop("text", "DRAFT"), **p)


# tool name (as exposed by the server) -> callable(pkg, params).
# Params dicts arrive pre-copied; adapters may pop from them freely.
_ALLOWED = {
    "search_and_replace": lambda pkg, p: _tx.search_and_replace(pkg, **p),
    "insert_paragraphs": lambda pkg, p: _tx.insert_paragraphs(pkg, **p),
    "delete_paragraphs": lambda pkg, p: _tx.delete_paragraphs(pkg, **p),
    "replace_paragraph_text": lambda pkg, p: _tx.replace_paragraph_text(pkg, **p),
    "format_text": lambda pkg, p: _tx.format_text(pkg, **p),
    "set_paragraph_format": lambda pkg, p: _tx.set_paragraph_format(pkg, **p),
    "apply_style": lambda pkg, p: _tx.apply_style(pkg, **p),
    "set_header": _set_header,
    "set_footer": _set_footer,
    "add_page_numbers": lambda pkg, p: _fu.add_page_numbers(pkg, **p),
    "set_page_number_format": lambda pkg, p: _fu.set_page_number_format(pkg, **p),
    "set_document_properties": lambda pkg, p: _sx.set_document_properties(pkg, **p),
    "set_cells": lambda pkg, p: _tb.set_cells(pkg, **p),
    "add_watermark": _add_watermark,
    "remove_watermark": lambda pkg, p: _fu.remove_watermark(pkg, **p),
}


def _validate_operations(operations: list[dict]) -> None:
    if not operations:
        raise WordMcpError("operations list is empty")
    for i, op in enumerate(operations):
        if not isinstance(op, dict) or "tool" not in op:
            raise WordMcpError(
                f"operation {i} must be a dict with 'tool' and optional 'params'"
            )
        if op["tool"] not in _ALLOWED:
            raise WordMcpError(
                f"tool {op['tool']!r} is not batchable; allowed: "
                f"{sorted(_ALLOWED)}"
            )
        params = op.get("params", {})
        if not isinstance(params, dict):
            raise WordMcpError(f"operation {i} ({op['tool']}): params must be a dict")


def batch_apply(
    file_paths: list[str],
    operations: list[dict],
    *,
    stop_on_error: bool = True,
    backup: bool = True,
) -> dict:
    """Apply the same operation list to many documents. Each operation:
    {'tool': <name>, 'params': {...}} using the tool's normal parameters
    minus file_path. Per file: all operations run in order, then ONE save
    (one backup); if any operation fails, that file is not saved and stays
    exactly as it was. stop_on_error=True additionally skips the remaining
    files — files already saved stay saved and are reported as such."""
    if not file_paths:
        raise WordMcpError("file_paths is empty")
    _validate_operations(operations)
    # sandbox check BEFORE the existence probe: a blocked path must refuse
    # identically whether or not the file exists (no existence oracle)
    from ..core.sandbox import check_path

    file_paths = [check_path(f, "batch edit") for f in file_paths]
    absent = [f for f in file_paths if not Path(f).is_file()]
    if absent:
        raise WordMcpError(
            "refusing to start (no file was touched): missing files "
            + ", ".join(absent)
        )

    files: list[dict] = []
    saved: list[str] = []
    failed: list[str] = []
    aborted = False
    for pos, file_path in enumerate(file_paths):
        if aborted:
            break
        entry: dict = {"path": file_path, "ok": False, "operations": []}
        files.append(entry)
        # Per-file write lock across the full read-modify-save cycle, same as
        # the single-file _edit path: a concurrent mutation of the same file
        # must not interleave with this one.
        with write_lock(file_path):
            try:
                pkg = DocxPackage(file_path)
            except Exception as exc:
                entry["error"] = f"{type(exc).__name__}: {exc}"
                failed.append(file_path)
                if stop_on_error:
                    aborted = True
                continue
            ops_ok = True
            for op in operations:
                tool = op["tool"]
                try:
                    result = _ALLOWED[tool](pkg, dict(op.get("params", {})))
                    entry["operations"].append(
                        {"tool": tool, "ok": True, "result": result}
                    )
                except TypeError as exc:
                    entry["operations"].append(
                        {"tool": tool, "ok": False,
                         "error": f"bad params for {tool}: {exc}"}
                    )
                    ops_ok = False
                    break
                except Exception as exc:
                    entry["operations"].append(
                        {"tool": tool, "ok": False,
                         "error": f"{type(exc).__name__}: {exc}"}
                    )
                    ops_ok = False
                    break
            if not ops_ok:
                entry["error"] = (
                    "an operation failed; this file was NOT saved and is unchanged"
                )
                failed.append(file_path)
                if stop_on_error:
                    aborted = True
                continue
            try:
                entry["saved"] = str(pkg.save(do_backup=backup))
                entry["ok"] = True
                saved.append(file_path)
            except Exception as exc:
                entry["error"] = (
                    f"save failed, file unchanged — {type(exc).__name__}: {exc}"
                )
                failed.append(file_path)
                if stop_on_error:
                    aborted = True

    not_attempted = file_paths[len(files):]
    result = {
        "files": files,
        "saved": saved,
        "failed": failed,
        "not_attempted": not_attempted,
    }
    if failed and saved:
        result["note"] = (
            f"{len(saved)} file(s) were already saved before the failure and "
            "keep their changes; failed and unattempted files are unchanged"
        )
    return result


# ======================================================================
# apply_edits: the v2 anchor-addressed batch layer (Section 9.2)
# ======================================================================

# op -> (accepted param keys beyond "op", required keys)
APPLY_OPS: dict[str, tuple[set, set]] = {
    "replace": ({"anchor", "find", "text", "occurrence"},
                {"anchor", "find", "text"}),
    "set_text": ({"anchor", "text"}, {"anchor", "text"}),
    "insert": ({"location", "markdown"}, {"location", "markdown"}),
    "delete": ({"anchor", "anchors"}, set()),
    "set_style": ({"anchor", "style"}, {"anchor", "style"}),
    "format": ({"anchor", "find", "formatting", "occurrence"},
               {"anchor", "formatting"}),
    "set_paragraph_format": ({"anchor", "format"}, {"anchor", "format"}),
    "set_cell": ({"anchor", "text"}, {"anchor", "text"}),
}


# ------------------------------------------------------------- markdown


_MD_HEADING = _re.compile(r"^(#{1,9})\s+(.*)$")
_MD_BULLET = _re.compile(r"^(\s*)[-*+]\s+(.*)$")
_MD_NUMBER = _re.compile(r"^(\s*)\d+[.)]\s+(.*)$")

_MD_UNSUPPORTED: tuple[tuple[str, str, str], ...] = (
    ("```", "fenced code block", "insert_paragraphs with a fixed-width "
     "style, or format_text"),
    ("![", "image", "insert_image"),
    (">", "block quote", "insert_paragraphs plus apply_style('Quote')"),
    ("<", "raw HTML", "the fine-grained insert_ tools"),
)


def parse_markdown(markdown: str) -> list[dict]:
    """Markdown -> ordered block segments for the insert op.

    Supported (Section 9.2 point 3): #-######### headings (styles Heading
    1-9), plain paragraphs (one non-blank line = one paragraph, matching
    the view's one-line-per-block projection), -/*/+ bullet and N. numbered
    lists (2-space indent per nesting level), and pipe tables (first row
    = header; a |---| separator row is skipped; a backslash-escaped pipe
    stays literal). Inline markdown (emphasis, links) is NOT interpreted
    and stays literal text. Anything else refuses with a pointer to the
    tool that can do it. Segments: {"kind": "paragraphs", "items": [{text,
    level?}]} | {"kind": "list", "list_kind": "bullet"|"number", "items":
    [(text, level)]} | {"kind": "table", "data": [[cells]]}.
    """
    if not isinstance(markdown, str) or not markdown.strip():
        raise WordMcpError("insert op needs non-empty markdown text")
    segments: list[dict] = []

    def para_seg() -> dict:
        if not segments or segments[-1]["kind"] != "paragraphs":
            segments.append({"kind": "paragraphs", "items": []})
        return segments[-1]

    def rule_break():
        # a blank line (or a block boundary) ends any open run so the next
        # list/paragraph starts fresh
        segments.append({"kind": "_break"})

    for raw in markdown.split("\n"):
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            rule_break()
            continue
        for prefix, what, pointer in _MD_UNSUPPORTED:
            if stripped.startswith(prefix):
                raise UnsupportedStructure(
                    f"markdown {what} is not representable by the insert "
                    f"op; use {pointer} instead. Nothing was applied."
                )
        if _re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", stripped):
            raise UnsupportedStructure(
                "markdown horizontal rules have no Word mapping; use "
                "insert_break or a bottom-bordered paragraph via "
                "set_paragraph_format. Nothing was applied."
            )
        m = _MD_HEADING.match(stripped)
        if m:
            para_seg()["items"].append(
                {"text": m.group(2).strip(), "level": len(m.group(1))}
            )
            rule_break()
            continue
        if stripped.startswith("|"):
            row = [
                c.strip().replace("\\|", "|")
                for c in _re.split(r"(?<!\\)\|", stripped.strip("|"))
            ]
            if all(_re.fullmatch(r":?-{2,}:?", c or "-") for c in row):
                continue  # separator row under the header
            if segments and segments[-1]["kind"] == "table":
                segments[-1]["data"].append(row)
            else:
                segments.append({"kind": "table", "data": [row]})
            continue
        m = _MD_BULLET.match(line) or None
        kind = "bullet" if m else None
        if m is None:
            m = _MD_NUMBER.match(line)
            kind = "number" if m else None
        if m:
            level = min(8, len(m.group(1)) // 2)
            item = (m.group(2).strip(), level)
            if segments and segments[-1]["kind"] == "list" \
                    and segments[-1]["list_kind"] == kind:
                segments[-1]["items"].append(item)
            else:
                segments.append(
                    {"kind": "list", "list_kind": kind, "items": [item]}
                )
            continue
        para_seg()["items"].append({"text": stripped})
    out = [s for s in segments if s["kind"] != "_break" and (
        s["kind"] != "paragraphs" or s["items"]
    )]
    if not out:
        raise WordMcpError("insert op markdown contained no content")
    return out


# ----------------------------------------------------------- validation


def _resolve_para_anchor(pkg: DocxPackage, edit: dict, op: str) -> dict:
    from . import view as _view

    info = _view.resolve_anchor(pkg, edit["anchor"])
    if info["kind"] != "paragraph":
        what = "table cell" if info["kind"] == "cell" else "table"
        exc = UnsupportedStructure(
            f"op {op!r} addresses body paragraphs, but anchor "
            f"{edit['anchor']!r} is a {what}"
            + ("; use the set_cell op or set_cells for cell text"
               if op in ("replace", "set_text")
               else f"; use delete_element(type='table', "
                    f"index={info['table_index']}) to delete the table"
               if op == "delete" and info["kind"] == "table"
               else "; use the table tools for tables")
        )
        if op == "delete":
            exc.hint_tools = ["delete_element"]
        raise exc
    return info


def _style_exists(pkg: DocxPackage, style: str) -> bool:
    from .read import list_styles

    styles = list_styles(pkg)
    if style in {s["id"] for s in styles}:
        return True
    if style in {s["name"] for s in styles if s["name"]}:
        return True
    return bool(_re.fullmatch(r"[Hh]eading\s?([1-9])", style))


def _validate_one(pkg: DocxPackage, edit: dict) -> dict:
    """Validate one edit and resolve its target against current state.
    Returns a plan consumed by _apply_one (elements for the file route,
    indices for the live route). Raises on any problem."""
    from . import view as _view
    from .read import body_items

    if not isinstance(edit, dict) or "op" not in edit:
        raise WordMcpError('each edit must be a dict with an "op" key')
    op = edit["op"]
    if op not in APPLY_OPS:
        exc = WordMcpError(
            f"op {op!r} is not an apply_edits op; supported: "
            f"{sorted(APPLY_OPS)}"
        )
        from ..packs import pack_of

        if isinstance(op, str) and pack_of(op) is not None:
            exc.hint_tools = [op]
        raise exc
    allowed, required = APPLY_OPS[op]
    extra = sorted(set(edit) - allowed - {"op"})
    if extra:
        raise WordMcpError(
            f"op {op!r} does not accept {extra}; allowed params: "
            f"{sorted(allowed)}"
        )
    missing = sorted(required - set(edit))
    if missing:
        raise WordMcpError(f"op {op!r} is missing required {missing}")

    if op == "insert":
        from ..core.locate import resolve_location

        segments = parse_markdown(edit["markdown"])
        r = resolve_location(pkg, edit["location"])
        if r.position == "replace":
            raise WordMcpError(
                "insert op position 'replace' is not meaningful; use the "
                "set_text op to replace a paragraph, or position "
                "before/after/start/end to insert"
            )
        mode = r.position  # before | after | start | end
        paras = [
            el for k, _i, el in body_items(pkg) if k == "paragraph"
        ]
        ref = None
        if mode not in ("end", "start"):
            if r.matched.get("implicit") or not paras:
                mode = "end"  # fresh document: everything lands at the end
            else:
                ref = paras[r.paragraph_index]
        return {"op": op, "kind": "insert", "mode": mode, "ref": ref,
                "index": r.paragraph_index, "segments": segments,
                "_keepalive": paras}

    if op == "delete":
        given = [k for k in ("anchor", "anchors") if k in edit]
        if len(given) != 1:
            raise WordMcpError(
                'delete op takes exactly one of "anchor" (one id) or '
                '"anchors" (a list of ids)'
            )
        ids = edit.get("anchors", [edit.get("anchor")])
        if not isinstance(ids, list) or not ids:
            raise WordMcpError('"anchors" must be a non-empty list of ids')
        infos = [
            _resolve_para_anchor(pkg, {"anchor": a}, op) for a in ids
        ]
        # Dedupe by resolved paragraph (adversarial 6a finding 1): the same
        # anchor listed twice (or two spellings of one anchor) otherwise
        # yields duplicate indices, and the bottom-up run deletion then
        # deletes a NEIGHBOR paragraph on the second pass. Deleting a
        # paragraph twice can only mean it once.
        seen_idx: set[int] = set()
        infos = [
            info for info in infos
            if not (info["paragraph_index"] in seen_idx
                    or seen_idx.add(info["paragraph_index"]))
        ]
        return {"op": op, "kind": "delete",
                "els": [i["el"] for i in infos],
                "indices": [i["paragraph_index"] for i in infos],
                "_keepalive": infos}

    if op == "set_cell":
        info = _view.resolve_anchor(pkg, edit["anchor"])
        if info["kind"] != "cell":
            raise WordMcpError(
                f"set_cell needs a CELL anchor (t:hex:rNcN, 1-based), got "
                f"{edit['anchor']!r}"
                + (" which is a whole table" if info["kind"] == "table"
                   else " which is a paragraph")
            )
        if not isinstance(edit["text"], str):
            raise WordMcpError("set_cell op: text must be a string")
        return {"op": op, "kind": "cell", "el": info["el"],
                "table_index": info["table_index"],
                "row": info["row"], "col": info["col"],
                "_keepalive": info}

    # remaining ops address one paragraph by anchor
    info = _resolve_para_anchor(pkg, edit, op)
    plan = {"op": op, "kind": "para", "el": info["el"],
            "index": info["paragraph_index"], "_keepalive": info}

    if op in ("replace", "set_text") and not isinstance(edit["text"], str):
        raise WordMcpError(f"op {op!r}: text must be a string")
    if op == "replace":
        find = edit["find"]
        if not isinstance(find, str) or not find:
            raise WordMcpError("replace op: find must be a non-empty string")
        count = info["text"].count(find)
        if count == 0:
            raise TargetNotFound(
                f"replace op: {find!r} does not occur in the anchored "
                f"paragraph (its text begins: {info['text'][:120]!r})"
            )
        occ = edit.get("occurrence")
        if occ is not None:
            if isinstance(occ, bool) or not isinstance(occ, int) or occ < 1:
                raise WordMcpError(
                    "replace op: occurrence is a 1-based integer"
                )
            if occ > count:
                raise TargetNotFound(
                    f"replace op: occurrence {occ} out of range, "
                    f"{find!r} occurs {count} time(s) in the paragraph"
                )
    elif op == "set_style":
        style = edit["style"]
        if not isinstance(style, str) or not style:
            raise WordMcpError("set_style op: style must be a style id/name")
        if not _style_exists(pkg, style):
            from .read import list_styles

            avail = sorted(
                s["id"] for s in list_styles(pkg)
                if s["type"] in (None, "paragraph") and s["id"]
            )
            raise TargetNotFound(
                f"style {style!r} not defined in this document; available "
                f"paragraph styles: {avail[:30]}"
            )
    elif op == "format":
        fmt = edit["formatting"]
        if not isinstance(fmt, dict) or not fmt:
            raise WordMcpError(
                "format op: formatting must be a non-empty dict"
            )
        _tx._check_keys(fmt, _tx._CHAR_FMT_KEYS, "character-formatting")
        find = edit.get("find")
        if find is not None:
            if not isinstance(find, str) or not find:
                raise WordMcpError(
                    "format op: find must be a non-empty string when given"
                )
            if find not in info["text"]:
                raise TargetNotFound(
                    f"format op: {find!r} does not occur in the anchored "
                    "paragraph"
                )
    elif op == "set_paragraph_format":
        fmt = edit["format"]
        if not isinstance(fmt, dict) or not fmt:
            raise WordMcpError(
                "set_paragraph_format op: format must be a non-empty dict"
            )
        _tx._check_keys(fmt, _tx._PARA_FMT_KEYS, "paragraph-formatting")
    return plan


def validate_edits(pkg: DocxPackage, edits) -> list[dict]:
    """Whole-batch validation (contract point 1): every failure collected,
    one aggregate refusal, nothing mutated. Returns the per-op plans."""
    if not isinstance(edits, list) or not edits:
        raise WordMcpError(
            'edits must be a non-empty list of {"op": ...} dicts'
        )
    plans: list[dict] = []
    failures: list[dict] = []
    hint_tools: list[str] = []
    stale = not_found = unsupported = False
    for i, edit in enumerate(edits):
        try:
            plans.append(_validate_one(pkg, edit))
            continue
        except StaleAnchor as exc:
            stale = True
            err_text = str(exc)
            src: BaseException = exc
        except TargetNotFound as exc:
            not_found = True
            err_text = str(exc)
            src = exc
        except UnsupportedStructure as exc:
            unsupported = True
            err_text = str(exc)
            src = exc
        except Exception as exc:
            err_text = (str(exc) if isinstance(exc, WordMcpError)
                        else f"{type(exc).__name__}: {exc}")
            src = exc
        failures.append(
            {"index": i,
             "op": edit.get("op") if isinstance(edit, dict) else None,
             "error": err_text}
        )
        for tool in getattr(src, "hint_tools", None) or []:
            if tool not in hint_tools:
                hint_tools.append(tool)
        plans.append({})
    if failures:
        err = WordMcpError(
            f"batch refused, nothing was applied: {len(failures)} of "
            f"{len(edits)} edits failed validation. Failures: {failures}. "
            "Fix or drop the failed edits (re-run get_document_view if "
            "anchors went stale) and resend the batch."
        )
        if stale:
            err.code = "STALE_ANCHOR"
        elif not_found:
            err.code = "NOT_FOUND"
        elif unsupported:
            err.code = "UNSUPPORTED_CONTENT"
        err.detail = {"failures": failures}
        if hint_tools:
            err.hint_tools = hint_tools
        raise err
    return plans


# ---------------------------------------------------------------- apply


def _current_index(pkg: DocxPackage, el, kind: str, op_index: int) -> int:
    """Recompute an element's CURRENT type-scoped body index. The scan list
    stays local (alive) and comparison is by identity."""
    from .read import body_items

    items = body_items(pkg)  # keepalive during comparison
    for k, idx, cand in items:
        if k == kind and cand is el:
            return idx
    raise WordMcpError(
        f"edit {op_index}: its target {kind} was removed by an earlier "
        "edit in this batch; the batch was ABANDONED and the file is "
        "unchanged. Keep deletes last or in their own batch."
    )


def _fresh_para_id(existing: set) -> str:
    while True:
        pid = f"{random.randrange(1, 0x80000000):08X}"  # MS-DOCX range
        if pid not in existing:
            existing.add(pid)
            return pid


def _stamp_new_paragraphs(els: list, existing: set) -> None:
    for el in els:
        if el.tag == qn("w:p"):
            targets = [el]
        elif el.tag == qn("w:tbl"):
            # cell paragraphs too: the table's anchor rides its first cell
            # paragraph's paraId, so inserted tables get durable t: anchors
            targets = list(el.iter(qn("w:p")))
        else:
            continue
        for p in targets:
            if p.get(qn("w14:paraId")) is None:
                p.set(qn("w14:paraId"), _fresh_para_id(existing))


def _build_insert_elements(pkg: DocxPackage, segments: list[dict]) -> list:
    from . import lists as _ls

    els: list = []
    for seg in segments:
        if seg["kind"] == "paragraphs":
            for item in seg["items"]:
                style = None
                if "level" in item:
                    style = _tx.ensure_heading_style(pkg, item["level"])
                els.append(_tx._make_paragraph(item["text"], style=style))
        elif seg["kind"] == "list":
            paras, _num = _ls.build_list_paragraphs(
                pkg, seg["items"], seg["list_kind"]
            )
            els.extend(paras)
        else:  # table
            els.append(_tb.build_table_element(seg["data"], header_row=True))
            # Word requires a paragraph between two tables and at body end
            # (the same spacer create_table adds).
            els.append(etree.Element(qn("w:p")))
    return els


def _apply_insert(pkg: DocxPackage, plan: dict, op_index: int) -> dict:
    els = _build_insert_elements(pkg, plan["segments"])
    body = pkg.body()
    mode = plan["mode"]
    if mode == "end":
        sectpr = body.find(qn("w:sectPr"))
        for el in els:
            if sectpr is not None:
                sectpr.addprevious(el)
            else:
                body.append(el)
    elif mode == "start":
        first = next(
            (c for c in body if c.tag in (qn("w:p"), qn("w:tbl"))), None
        )
        if first is None:
            for el in els:
                body.append(el)
        else:
            for el in els:
                first.addprevious(el)
    else:
        ref = plan["ref"]
        _current_index(pkg, ref, "paragraph", op_index)  # still in body?
        if mode == "after":
            for el in reversed(els):
                ref.addnext(el)
        else:  # before
            for el in els:
                ref.addprevious(el)
    # guard: an inserted table now adjacent to a pre-existing table would
    # merge with it in Word; keep a separator paragraph between them
    for el in els:
        if el.tag == qn("w:tbl"):
            prev = el.getprevious()
            if prev is not None and prev.tag == qn("w:tbl"):
                el.addprevious(etree.Element(qn("w:p")))
    pkg.mark_dirty()
    n_paras = sum(1 for el in els if el.tag == qn("w:p"))
    n_tables = sum(1 for el in els if el.tag == qn("w:tbl"))
    return {"inserted_paragraphs": n_paras, "inserted_tables": n_tables,
            "_new_els": els}


def _apply_one(pkg: DocxPackage, edit: dict, plan: dict, i: int) -> dict:
    from . import _runmap

    op = plan["op"]
    if op == "insert":
        return _apply_insert(pkg, plan, i)
    if op == "delete":
        indices = sorted({
            _current_index(pkg, el, "paragraph", i) for el in plan["els"]
        })
        # contiguous runs, deleted bottom-up so indices stay valid
        runs: list[list[int]] = []
        for idx in indices:
            if runs and idx == runs[-1][-1] + 1:
                runs[-1].append(idx)
            else:
                runs.append([idx])
        deleted = 0
        for run in reversed(runs):
            _tx.delete_paragraphs(pkg, run[0], run[-1])
            deleted += len(run)
        return {"deleted": deleted}
    if op == "set_cell":
        tindex = _current_index(pkg, plan["el"], "table", i)
        return {
            **_tb.set_cells(
                pkg, tindex,
                [{"row": plan["row"], "cell": plan["col"],
                  "text": edit["text"]}],
            ),
            "table": tindex, "row": plan["row"], "col": plan["col"],
        }

    idx = _current_index(pkg, plan["el"], "paragraph", i)
    if op == "set_text":
        return _tx.replace_paragraph_text(pkg, idx, edit["text"])
    if op == "set_style":
        return _tx.apply_style(pkg, [idx], edit["style"])
    if op == "format":
        return _tx.format_text(
            pkg, paragraph_index=idx, find=edit.get("find"),
            occurrence=edit.get("occurrence") or 1,
            formatting=edit["formatting"],
        )
    if op == "set_paragraph_format":
        return _tx.set_paragraph_format(pkg, [idx], edit["format"])
    # replace: paragraph-scoped find/replace over the run map, applied
    # right-to-left on one snapshot (the search_and_replace convention)
    p = plan["el"]
    text, segments = _runmap.build_map(p)
    find, new = edit["find"], edit["text"]
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
            f"edit {i} (replace): {edit['find']!r} no longer occurs in the "
            "anchored paragraph (an earlier edit in this batch changed "
            "it); the batch was ABANDONED and the file is unchanged"
        )
    occ = edit.get("occurrence")
    if occ is not None:
        if occ > len(spans):
            raise TargetNotFound(
                f"edit {i} (replace): occurrence {occ} out of range, "
                f"{len(spans)} match(es) in the paragraph; the batch was "
                "ABANDONED and the file is unchanged"
            )
        spans = [spans[occ - 1]]
    for s, e in reversed(spans):
        _runmap.replace_range(p, segments, s, e, new)
    pkg.mark_dirty()
    return {"replaced": len(spans), "paragraph": idx}


def apply_edits(pkg: DocxPackage, edits, atomic: bool = True) -> dict:
    """Module docstring has the contract. Returns
    {applied, changed: {op index -> result}, warnings}; the server's _edit
    wrapper supplies the single lock/backup/save around it."""
    if atomic is not True:
        raise WordMcpError(
            "only atomic=true is supported: the whole batch applies in one "
            "save or nothing does. Split into separate apply_edits calls "
            "for independent failure domains."
        )
    plans = validate_edits(pkg, edits)

    existing_pids = {
        p.get(qn("w14:paraId")).upper()
        for p in pkg.root().iter(qn("w:p"))
        if p.get(qn("w14:paraId"))
    }
    changed: dict[str, dict] = {}
    warnings: list[str] = []
    inserted_by_op: dict[str, list] = {}
    for i, (edit, plan) in enumerate(zip(edits, plans)):
        try:
            result = _apply_one(pkg, edit, plan, i)
        except WordMcpError:
            raise
        except Exception as exc:
            raise WordMcpError(
                f"edit {i} ({plan['op']}) failed during apply; the batch "
                f"was ABANDONED and the file is unchanged: "
                f"{type(exc).__name__}: {exc}"
            ) from exc
        new_els = result.pop("_new_els", None)
        if new_els:
            # durable anchors for chaining (contract point 4): stamp fresh
            # paraIds now, translate to display anchors after the batch
            _stamp_new_paragraphs(new_els, existing_pids)
            inserted_by_op[str(i)] = new_els
        warnings.extend(result.pop("warnings", []) or [])
        changed[str(i)] = result

    if inserted_by_op:
        from . import view as _view

        amap = _view.anchor_map(pkg)
        _keepalive = amap["items"]  # noqa: F841  (proxy identity safety)
        p_disp = _view.display_anchors(amap["paragraphs"])
        t_disp = _view.display_anchors(amap["tables"])
        by_p = {id(rec["el"]): p_disp[j]
                for j, rec in enumerate(amap["paragraphs"])}
        by_t = {id(rec["el"]): "t:" + t_disp[j]
                for j, rec in enumerate(amap["tables"])}
        for key, els in inserted_by_op.items():
            # els keeps the inserted proxies alive, so id() lookups against
            # the live map above stay valid (the Phase 1 lxml lesson)
            anchors = [by_p[id(el)] for el in els
                       if el.tag == qn("w:p") and id(el) in by_p]
            t_anchors = [by_t[id(el)] for el in els
                         if el.tag == qn("w:tbl") and id(el) in by_t]
            changed[key]["anchors"] = anchors
            if t_anchors:
                changed[key]["table_anchors"] = t_anchors

    return {"applied": len(edits), "changed": changed, "warnings": warnings}
