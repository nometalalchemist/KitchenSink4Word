"""View layer: get_document_view, the anchored markdown projection of a
document, plus resolve_anchor (the scheme's inverse) and stamp_anchors
(the explicit opt-in durability mutation).

Contract: get_document_view and resolve_anchor take a DocxPackage, never
touch disk, never call mark_dirty(), never import FastMCP or COM. The view
is DETERMINISTIC: two reads of identical bytes produce identical output.
stamp_anchors is the ONE mutating function here and only ever runs behind
the server's explicit stamp_anchors=True parameter (V2_DESIGN ruling 13.5:
reads never mutate; stamping requires the explicit opt-in).

THE ANCHOR SCHEME (ops/batch.py apply_edits and the core.locate anchor
selector resolve these; it must stay re-derivable from package content):

- Paragraph anchors ride w14:paraId, the durable per-paragraph id Word
  stamps on save. It survives edits elsewhere in the document, which is
  exactly the stability property positional indices lack. The anchor
  DIGEST is the paraId lowercased (8 hex chars); the DISPLAY form is the
  last 4 hex chars, extended to 8 (then the full digest) only where the
  short suffixes collide within the document.
- Paragraphs without a paraId (non-Word producers, python-docx output, and
  paragraphs our own file tools insert) get a VOLATILE fallback digest,
  sha1("p:<index>:<text>"), same display rules. Volatile anchors resolve
  identically but change with ANY edit; the view header says so, and the
  stamp_anchors opt-in writes real paraIds so subsequent views are stable.
  A paraId duplicated across paragraphs (corrupt copy-paste artifacts) is
  untrustworthy, so those paragraphs fall back to volatile digests too.
- Tables anchor on the paraId of their FIRST cell paragraph when present
  (durable), else a volatile sha1("t:<index>:<header text>"). Display form
  "t:<hex>"; cells are addressed "t:<hex>:rNcN" with 1-BASED row/column
  numbers in the view and in apply_edits set_cell ops (set_cells and
  get_table keep their 0-based row/cell coordinates).
- resolve_anchor matches the given hex as a digest SUFFIX (display anchors
  are suffixes). Zero matches raise StaleAnchor (the document changed, or
  the anchor was mistyped: re-run get_document_view); several matches
  raise AmbiguousTarget listing extended anchors.

Detail levels: "structure" = headings only, with per-section paragraph
counts; "text" (default) = every block, anchored, full prose; "full" adds
inline revision markers ({++ins++}, {--del--}), comment references ([cN]),
and an author legend in the header.
"""

from __future__ import annotations

import hashlib
import random
import re

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    StaleAnchor,
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from ..core.package import DocxPackage, qn
from . import read as _rd

_DETAILS = ("structure", "text", "full")
_NOTES_MODES = ("refs", "inline")

_SUFFIX_SHORT = 4
_SUFFIX_LONG = 8

# Tables degrade to a stub past these bounds (the projection refuses to
# misrepresent; it never silently truncates content without saying so).
_MAX_TABLE_COLS = 12
_MAX_CELL_CHARS = 200

_ANCHOR_RE = re.compile(
    r"^(?P<t>t:)?(?P<hex>[0-9a-fA-F]{3,40})(?::r(?P<row>\d+)c(?P<col>\d+))?$"
)


def _para_id(p: etree._Element) -> str | None:
    return p.get(qn("w14:paraId"))


def _volatile_digest(kind: str, index: int, text: str) -> str:
    return hashlib.sha1(f"{kind}:{index}:{text}".encode()).hexdigest()


def _table_dims(tbl: etree._Element) -> tuple[int, int]:
    rows = tbl.findall(qn("w:tr"))
    cols = max((len(tr.findall(qn("w:tc"))) for tr in rows), default=0)
    return len(rows), cols


def _first_table_para(tbl: etree._Element) -> etree._Element | None:
    for p in tbl.iter(qn("w:p")):
        return p
    return None


def anchor_map(pkg: DocxPackage) -> dict:
    """Every body block with its anchor digest.

    Returns {"paragraphs": [rec], "tables": [rec], "items": keepalive}.
    Each rec: {kind, index, el, digest, volatile}. CAUTION (Phase 1
    lesson): the "items" list MUST stay referenced for as long as the recs
    are used; lxml keeps at most one live proxy per node, and identity
    comparisons on recs' elements are only safe while that list is alive.
    """
    items = _rd.body_items(pkg)  # keepalive for every el below
    # A paraId shared by several paragraphs identifies nothing; ban dupes.
    seen: dict[str, int] = {}
    for kind, _idx, el in items:
        if kind != "paragraph":
            continue
        pid = _para_id(el)
        if pid:
            seen[pid.lower()] = seen.get(pid.lower(), 0) + 1
    paragraphs: list[dict] = []
    tables: list[dict] = []
    for kind, idx, el in items:
        if kind == "paragraph":
            pid = _para_id(el)
            if pid and seen.get(pid.lower(), 0) == 1:
                digest, volatile = pid.lower(), False
            else:
                digest = _volatile_digest("p", idx, _rd.paragraph_text(el))
                volatile = True
            paragraphs.append(
                {"kind": "paragraph", "index": idx, "el": el,
                 "digest": digest, "volatile": volatile}
            )
        else:
            first = _first_table_para(el)
            pid = _para_id(first) if first is not None else None
            if pid:
                digest, volatile = pid.lower(), False
            else:
                header = ""
                tr = el.find(qn("w:tr"))
                if tr is not None:
                    header = "|".join(
                        _rd.paragraph_text(p) for p in tr.iter(qn("w:p"))
                    )
                digest = _volatile_digest("t", idx, header)
                volatile = True
            tables.append(
                {"kind": "table", "index": idx, "el": el,
                 "digest": digest, "volatile": volatile}
            )
    return {"paragraphs": paragraphs, "tables": tables, "items": items}


def display_anchors(recs: list[dict]) -> dict[int, str]:
    """rec position -> display anchor: last 4 hex chars of the digest,
    extended (8, then full) only for colliding suffixes so the projection
    stays short AND unambiguous within its namespace."""
    out: dict[int, str] = {}
    by_short: dict[str, list[int]] = {}
    for i, rec in enumerate(recs):
        by_short.setdefault(rec["digest"][-_SUFFIX_SHORT:], []).append(i)
    for short, group in by_short.items():
        if len(group) == 1:
            out[group[0]] = short
            continue
        by_long: dict[str, list[int]] = {}
        for i in group:
            by_long.setdefault(recs[i]["digest"][-_SUFFIX_LONG:], []).append(i)
        for long_sfx, lgroup in by_long.items():
            for i in lgroup:
                out[i] = long_sfx if len(lgroup) == 1 else recs[i]["digest"]
    return out


# ------------------------------------------------------------ anchor inverse


def resolve_anchor(pkg: DocxPackage, anchor: str) -> dict:
    """Resolve a view anchor back to its target in the CURRENT document.

    Accepts "<hex>" (paragraph), "t:<hex>" (table), "t:<hex>:rNcN" (cell,
    1-based). Returns {"kind": "paragraph"|"table"|"cell", ...} with
    paragraph_index or table_index (+ 0-based row/col for cells), the
    matched element under "el", "volatile", and "map" (KEEP the returned
    dict referenced while using "el": the map's keepalive list rides in
    it). Stale anchors raise StaleAnchor; suffix collisions raise
    AmbiguousTarget with extended candidate anchors."""
    if not isinstance(anchor, str) or not anchor:
        raise WordMcpError(
            "anchor must be a get_document_view anchor id (string), like "
            '"a3f9", "t:19e4", or "t:19e4:r2c3"'
        )
    m = _ANCHOR_RE.match(anchor.strip())
    if m is None:
        raise WordMcpError(
            f"malformed anchor {anchor!r}: expected a hex anchor id from "
            'get_document_view, like "a3f9" (paragraph), "t:19e4" (table), '
            'or "t:19e4:r2c3" (cell, 1-based)'
        )
    is_table = bool(m.group("t"))
    hexpart = m.group("hex").lower()
    amap = anchor_map(pkg)
    ns = amap["tables"] if is_table else amap["paragraphs"]
    hits = [rec for rec in ns if rec["digest"].endswith(hexpart)]
    if not hits:
        what = "table" if is_table else "paragraph"
        raise StaleAnchor(
            f"anchor {anchor!r} matches no {what} in {pkg.path.name}: the "
            "document changed since the view was taken (or the anchor was "
            "mistyped). Re-run get_document_view and use fresh anchors."
        )
    if len(hits) > 1:
        displays = display_anchors(ns)
        pos = {id(rec): i for i, rec in enumerate(ns)}
        cands = [
            {
                "anchor": ("t:" if is_table else "")
                + displays[pos[id(rec)]],
                ("table_index" if is_table else "paragraph"): rec["index"],
            }
            for rec in hits
        ]
        raise AmbiguousTarget(
            f"anchor {anchor!r} matches {len(hits)} blocks; use a longer "
            f"anchor. Candidates: {cands}"
        )
    rec = hits[0]
    if not is_table:
        return {
            "kind": "paragraph",
            "paragraph_index": rec["index"],
            "el": rec["el"],
            "volatile": rec["volatile"],
            "text": _rd.paragraph_text(rec["el"]),
            "map": amap,
        }
    out = {
        "kind": "table",
        "table_index": rec["index"],
        "el": rec["el"],
        "volatile": rec["volatile"],
        "map": amap,
    }
    if m.group("row") is not None:
        row, col = int(m.group("row")) - 1, int(m.group("col")) - 1
        trs = rec["el"].findall(qn("w:tr"))
        if not 0 <= row < len(trs):
            raise TargetNotFound(
                f"anchor {anchor!r}: row {row + 1} out of range, the table "
                f"has {len(trs)} row(s) (rNcN is 1-based)"
            )
        tcs = trs[row].findall(qn("w:tc"))
        if not 0 <= col < len(tcs):
            raise TargetNotFound(
                f"anchor {anchor!r}: column {col + 1} out of range, row "
                f"{row + 1} has {len(tcs)} cell(s) (rNcN is 1-based)"
            )
        tcpr = tcs[col].find(qn("w:tcPr"))
        vm = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
        if vm is not None and vm.get(qn("w:val"), "continue") == "continue":
            raise UnsupportedStructure(
                f"anchor {anchor!r} addresses a vertically merged "
                "CONTINUATION cell; text written there is invisible in "
                "Word. Address the restart cell at the top of the merge."
            )
        out["kind"] = "cell"
        out["row"], out["col"] = row, col
    return out


# ------------------------------------------------------------------ stamping


def stamp_anchors(pkg: DocxPackage) -> dict:
    """Write a w14:paraId to every paragraph in document.xml that lacks
    one, making anchors durable. EXPLICIT OPT-IN MUTATION (ruling 13.5:
    reads never mutate); the server only calls this behind
    stamp_anchors=true, through the normal lock/backup/save cycle. Ids are
    random 8-hex values in Word's valid range, unique in the document."""
    existing: set[str] = set()
    root = pkg.root()
    for p in root.iter(qn("w:p")):
        pid = _para_id(p)
        if pid:
            existing.add(pid.upper())
    stamped = 0
    for p in root.iter(qn("w:p")):
        if _para_id(p):
            continue
        while True:
            # MS-DOCX: paraId is nonzero and below 0x80000000.
            pid = f"{random.randrange(1, 0x80000000):08X}"
            if pid not in existing:
                break
        existing.add(pid)
        p.set(qn("w14:paraId"), pid)
        stamped += 1
    if stamped:
        pkg.mark_dirty()
    return {"stamped": stamped}


# ----------------------------------------------------------------- rendering


def _pipe_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def _cell_texts(tr: etree._Element) -> list[str]:
    return [
        "\n".join(_rd.paragraph_text(p) for p in tc.findall(qn("w:p")))
        for tc in tr.findall(qn("w:tc"))
    ]


def _table_lines(tbl: etree._Element, anchor: str, index: int) -> list[str]:
    rows, cols = _table_dims(tbl)
    head = f"[t:{anchor}] table {rows}x{cols} (cells t:{anchor}:rNcN, 1-based)"
    trs = tbl.findall(qn("w:tr"))
    texts = [_cell_texts(tr) for tr in trs]
    nested = any(tc.find(qn("w:tbl")) is not None
                 for tr in trs for tc in tr.findall(qn("w:tc")))
    merged = tbl.find(f".//{qn('w:vMerge')}") is not None or (
        tbl.find(f".//{qn('w:gridSpan')}") is not None
    )
    wide = cols > _MAX_TABLE_COLS or any(
        len(c) > _MAX_CELL_CHARS for row in texts for c in row
    )
    if nested or merged or wide:
        why = ("nested tables" if nested
               else "merged cells" if merged else "wide content")
        lines = [head + f" [{why}: stub only, read with get_table "
                        f"index={index}]"]
        if texts:
            lines.append(
                "| " + " | ".join(
                    _pipe_cell(c)[:40] for c in texts[0]
                ) + " |"
            )
        return lines
    lines = [head]
    for row in texts:
        lines.append("| " + " | ".join(_pipe_cell(c) for c in row) + " |")
    return lines


def _marked_paragraph_text(
    p: etree._Element, comment_marks: bool = True
) -> str:
    """Paragraph text with inline revision markers and comment refs, for
    detail='full': w:ins content as {++...++}, w:del as {--...--},
    commentRangeEnd as [cN]. A projection for agents, not a fidelity
    claim."""
    parts: list[str] = []

    def runs_text(el, include_deleted=False) -> str:
        return "".join(
            _rd.run_text(r, include_deleted=include_deleted)
            for r in el.iter(qn("w:r"))
        )

    def walk(el):
        for child in el:
            tag = etree.QName(child).localname
            if tag == "r":
                parts.append(_rd.run_text(child))
            elif tag == "ins":
                text = runs_text(child)
                if text:
                    parts.append("{++" + text + "++}")
            elif tag in ("del", "moveFrom"):
                text = runs_text(child, include_deleted=True)
                if text:
                    parts.append("{--" + text + "--}")
            elif tag == "commentRangeEnd" and comment_marks:
                parts.append(f"[c{child.get(qn('w:id'))}]")
            elif tag in ("hyperlink", "smartTag", "sdt", "sdtContent"):
                walk(child)
    walk(p)
    return "".join(parts)


def _outline_entries(pkg: DocxPackage) -> list[dict]:
    from ..core import locate as _loc  # lazy: locate lazily imports ops

    return _loc._outline_entries(pkg)


def _scope_range(pkg: DocxPackage, scope, entries: list[dict]) -> tuple:
    """scope -> (start_para_index, end_para_index_exclusive|None)."""
    if scope is None:
        return 0, None
    if not isinstance(scope, dict) or len(scope) != 1 or not (
        "outline" in scope or "paragraphs" in scope
    ):
        raise WordMcpError(
            'scope takes {"outline": "3.2"} (that heading\'s section) or '
            '{"paragraphs": {"start": N, "end": M}} (end exclusive); omit '
            "it for the whole document"
        )
    if "paragraphs" in scope:
        spec = scope["paragraphs"]
        if not isinstance(spec, dict) or "start" not in spec:
            raise WordMcpError(
                'scope.paragraphs takes {"start": N, "end": M} with 0-based '
                "body paragraph indices, end exclusive (end optional)"
            )
        start = spec["start"]
        end = spec.get("end")
        if isinstance(start, bool) or not isinstance(start, int) or start < 0:
            raise WordMcpError("scope.paragraphs.start must be an int >= 0")
        if end is not None and (
            isinstance(end, bool) or not isinstance(end, int) or end <= start
        ):
            raise WordMcpError(
                "scope.paragraphs.end must be an int > start (end exclusive)"
            )
        return start, end
    path = scope["outline"]
    for i, e in enumerate(entries):
        if e["path"] == path:
            start = e["paragraph_index"]
            for later in entries[i + 1:]:
                if later["level"] <= e["level"]:
                    return start, later["paragraph_index"]
            return start, None
    listing = "; ".join(
        f"{e['path']} {e['text'][:40]!r}" for e in entries[:25]
    )
    raise TargetNotFound(
        f"no heading at outline path {path!r}; document outline: "
        + (listing if entries else "no headings detected")
    )


def get_document_view(
    pkg: DocxPackage,
    scope=None,
    detail: str = "text",
    include: dict | None = None,
) -> dict:
    """The anchored projection (module docstring has the scheme). PURE
    READ: never mutates the package."""
    if detail not in _DETAILS:
        raise WordMcpError(
            f"unknown detail {detail!r}; one of: {', '.join(_DETAILS)}"
        )
    include = include or {}
    unknown = sorted(set(include) - {"tables", "notes"})
    if unknown:
        raise WordMcpError(
            f"unknown include key(s) {unknown}; include takes "
            '{"tables": bool, "notes": "refs"|"inline"}'
        )
    tables_on = include.get("tables", True)
    notes_mode = include.get("notes", "refs")
    if notes_mode not in _NOTES_MODES:
        raise WordMcpError(
            f"include.notes must be one of {_NOTES_MODES}, got {notes_mode!r}"
        )

    amap = anchor_map(pkg)
    _keepalive = amap["items"]  # noqa: F841 (proxy identity, Phase 1 lesson)
    p_disp = display_anchors(amap["paragraphs"])
    t_disp = display_anchors(amap["tables"])
    p_anchor = {rec["index"]: p_disp[i]
                for i, rec in enumerate(amap["paragraphs"])}
    t_anchor = {rec["index"]: t_disp[i]
                for i, rec in enumerate(amap["tables"])}
    volatile_count = sum(
        1 for rec in amap["paragraphs"] + amap["tables"] if rec["volatile"]
    )
    total = len(amap["paragraphs"]) + len(amap["tables"])
    if volatile_count == 0:
        mode = "paraId"
    elif volatile_count == total:
        mode = "volatile"
    else:
        mode = "mixed"

    entries = _outline_entries(pkg)
    by_idx = {e["paragraph_index"]: e for e in entries}
    start, end = _scope_range(pkg, scope, entries)

    # ----- header
    name = pkg.path.name
    lines = [f"# {name} ({total} blocks, detail={detail}, anchors={mode})"]
    if detail != "structure":
        lines.append(
            "Anchors: [hex] paragraph, [t:hex] table, cells t:hex:rNcN "
            '(1-based). Address blocks as {"anchor": "hex"} in any '
            "location object or apply_edits op."
        )
    if volatile_count:
        lines.append(
            f"CAUTION: {volatile_count} of {total} anchors are VOLATILE "
            "(no w14:paraId); they change with any edit. Re-view after "
            "each batch, or pass stamp_anchors=true once (a mutation) to "
            "make them durable."
        )
    if scope is None and detail != "structure" and entries:
        lines.append(
            "Outline: " + " | ".join(
                f"{e['path']} {e['text']} [{p_anchor[e['paragraph_index']]}]"
                for e in entries
            )
        )

    # ----- comment / revision legend at detail=full
    if detail == "full":
        comments = _rd.get_comments(pkg)
        rev_authors = sorted(
            {r["author"] for r in _rd.get_tracked_changes(pkg) if r["author"]}
        )
        legend = []
        if rev_authors:
            legend.append(
                "revisions {++ins++}/{--del--} by: " + ", ".join(rev_authors)
            )
        if comments:
            legend.append(
                "comments [cN]: " + "; ".join(
                    f"c{c['id']}={c['author']}" for c in comments
                )
            )
        if legend:
            lines.append("Markers: " + ". ".join(legend))

    # ----- per-section paragraph counts for detail=structure
    section_paras: dict[int, int] = {}
    if detail == "structure":
        heading_indices = sorted(by_idx)
        n_paras = len(amap["paragraphs"])
        for i, hidx in enumerate(heading_indices):
            nxt = (heading_indices[i + 1]
                   if i + 1 < len(heading_indices) else n_paras)
            section_paras[hidx] = max(0, nxt - hidx - 1)

    # ----- blocks
    blocks = 0
    last_para = -1
    for kind, idx, el in amap["items"]:
        if kind == "paragraph":
            last_para = idx
            if idx < start or (end is not None and idx >= end):
                continue
            head = by_idx.get(idx)
            if detail == "structure":
                if head is None:
                    continue
                lines.append(
                    f"[{p_anchor[idx]}] "
                    + "#" * head["level"] + " "
                    + head["text"]
                    + f" ({section_paras.get(idx, 0)} paras)"
                )
                blocks += 1
                continue
            text = (
                _marked_paragraph_text(el)
                if detail == "full"
                else _rd.paragraph_text(el)
            )
            prefix = "#" * head["level"] + " " if head else ""
            lines.append(f"[{p_anchor[idx]}] {prefix}{text}".rstrip())
            blocks += 1
        else:
            if detail == "structure" or not tables_on:
                continue
            # a table rides with the paragraph it follows; leading tables
            # only appear in whole-document views
            if not (scope is None or (start <= last_para and
                                      (end is None or last_para < end))):
                continue
            lines.append("")
            lines.extend(_table_lines(el, t_anchor[idx], idx))
            blocks += 1

    # ----- structure mode with no headings: flat fallback, never emptiness
    if detail == "structure" and blocks == 0:
        n_paras = len(amap["paragraphs"])
        n_tables = len(amap["tables"])
        n_words = sum(
            len(_rd.paragraph_text(el).split())
            for kind, _i, el in amap["items"]
            if kind == "paragraph"
        )
        lines.append(
            "No headings detected (no Heading styles or outlineLvl); the "
            "document may use direct formatting for structure. Flat "
            f"structure: {n_paras} paragraphs, ~{n_words} words, "
            f"{n_tables} tables. Try get_outline with "
            "detect_formatted=true, or detail='text' for the full view."
        )
        out_note = (
            "no headings detected; flat structure counts reported instead"
        )
    else:
        out_note = None

    # ----- notes
    if detail != "structure" and notes_mode == "inline":
        notes = []
        for n in _rd.list_footnotes(pkg):
            notes.append(f"[fn:{n['id']}] {n['text']}")
        for n in _rd.list_endnotes(pkg):
            notes.append(f"[en:{n['id']}] {n['text']}")
        if notes:
            lines.append("")
            lines.append("Notes:")
            lines.extend(notes)

    out = {
        "view": "\n".join(lines),
        "blocks": blocks,
        "anchor_mode": mode,
        "detail": detail,
    }
    if out_note:
        out["note"] = out_note
    if volatile_count:
        out["volatile_anchors"] = volatile_count
    return out
