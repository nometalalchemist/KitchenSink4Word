"""Read layer: extraction of text, structure, tables, notes, comments, revisions.

All functions take a DocxPackage and return plain dict/list structures ready to
be serialized by the MCP layer. Indices returned here (paragraph index, table
index) are the stable addressing scheme the edit layers accept as targets:
body-level paragraphs and tables are numbered in document order, separately.
"""

from __future__ import annotations

from lxml import etree

from ..core.package import NSMAP, DocxPackage, qn

# ---------------------------------------------------------------- text helpers

_HEADING_STYLES = {f"Heading{i}": i for i in range(1, 10)}
_HEADING_STYLES.update({f"heading {i}": i for i in range(1, 10)})


def run_text(r: etree._Element, *, include_deleted: bool = False) -> str:
    """Text of one w:r, honoring tabs and breaks. Deleted text excluded by default."""
    parts: list[str] = []
    for child in r:
        tag = etree.QName(child).localname
        if tag == "t":
            parts.append(child.text or "")
        elif tag == "delText" and include_deleted:
            parts.append(child.text or "")
        elif tag == "tab":
            parts.append("\t")
        elif tag == "br":
            # Page/column breaks are layout, not text.
            if child.get(qn("w:type")) not in ("page", "column"):
                parts.append("\n")
        elif tag == "cr":
            parts.append("\n")
        elif tag == "noBreakHyphen":
            parts.append("-")
        elif tag == "footnoteReference":
            parts.append(f"[fn:{child.get(qn('w:id'))}]")
        elif tag == "endnoteReference":
            parts.append(f"[en:{child.get(qn('w:id'))}]")
    return "".join(parts)


def paragraph_text(p: etree._Element, *, include_deleted: bool = False) -> str:
    """Full visible text of a w:p, descending into w:ins/hyperlinks/smartTags.
    Text-box content is excluded — it is a separate story that Word stores
    TWICE (mc:Choice + mc:Fallback), so counting it here doubles it; use
    ops/textboxes.py for box text."""
    from ._runmap import _in_textbox

    parts: list[str] = []
    for r in p.iter(qn("w:r")):
        # Runs inside w:del hold w:delText only; run_text already skips those
        # unless include_deleted is set.
        if _in_textbox(r, p):
            continue
        parts.append(run_text(r, include_deleted=include_deleted))
    return "".join(parts)


def _style_id(p: etree._Element) -> str | None:
    pstyle = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    return pstyle.get(qn("w:val")) if pstyle is not None else None


def default_paragraph_style(pkg: DocxPackage) -> str:
    """The document's default paragraph style id (the style Word applies to
    paragraphs with no explicit pStyle) — the w:default paragraph style in
    styles.xml, 'Normal' when none is marked."""
    if pkg.has_part("word/styles.xml"):
        for style in pkg.root("word/styles.xml").findall(qn("w:style")):
            if (
                style.get(qn("w:default")) in ("1", "true")
                and style.get(qn("w:type")) == "paragraph"
            ):
                sid = style.get(qn("w:styleId"))
                if sid:
                    return sid
    return "Normal"


def _outline_level_detected(
    p: etree._Element, style_outline: dict[str, int]
) -> tuple[int | None, str | None]:
    """(heading level 1-9 | None, how it was detected).

    Detection order mirrors Word's effective-value resolution:
    1. direct w:outlineLvl in the paragraph's own pPr ("outline_level") —
       OOXML val 0-8 maps to level 1-9; val 9 is the EXPLICIT body-text
       value and OVERRIDES any style-supplied level (returns None);
    2. a built-in Heading style ("heading_style");
    3. w:outlineLvl carried by the paragraph's style or inherited through
       its basedOn chain ("outline_level") — the academic-template pattern
       (e.g. Normal-styled/Normal-based paragraphs with outlineLvl
       overrides), which Word's own navigation pane honors.
    """
    lvl = p.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
    if lvl is not None:
        val = int(lvl.get(qn("w:val")))
        if 0 <= val <= 8:
            return val + 1, "outline_level"
        return None, None  # val 9 = explicit body text
    sid = _style_id(p)
    if sid:
        if sid in _HEADING_STYLES:
            return _HEADING_STYLES[sid], "heading_style"
        if sid in style_outline:
            return style_outline[sid], "outline_level"
    return None, None


def _outline_level(p: etree._Element, style_outline: dict[str, int]) -> int | None:
    """Heading level 1-9, from direct outlineLvl, style name, or style's outlineLvl."""
    return _outline_level_detected(p, style_outline)[0]


def _style_outline_map(pkg: DocxPackage) -> dict[str, int]:
    """styleId -> heading level (1-9), from styles.xml outlineLvl definitions,
    resolved through each style's basedOn chain (a style whose ancestor
    carries outlineLvl inherits it, exactly as Word resolves the effective
    value). outlineLvl val 9 (explicit body text) yields no entry."""
    if not pkg.has_part("word/styles.xml"):
        return {}
    direct: dict[str, int] = {}
    based_on: dict[str, str] = {}
    for style in pkg.root("word/styles.xml").findall(qn("w:style")):
        if style.get(qn("w:type")) not in (None, "paragraph"):
            continue
        sid = style.get(qn("w:styleId"))
        if not sid:
            continue
        lvl = style.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
        if lvl is not None:
            direct[sid] = int(lvl.get(qn("w:val")))
        base = style.find(qn("w:basedOn"))
        if base is not None and base.get(qn("w:val")):
            based_on[sid] = base.get(qn("w:val"))
    out: dict[str, int] = {}
    for sid in set(direct) | set(based_on):
        cur: str | None = sid
        seen: set[str] = set()
        while cur is not None and cur not in seen:
            seen.add(cur)
            if cur in direct:
                if 0 <= direct[cur] <= 8:
                    out[sid] = direct[cur] + 1
                break  # val 9 = explicit body text: stops inheritance too
            cur = based_on.get(cur)
    return out


def body_items(pkg: DocxPackage) -> list[tuple[str, int, etree._Element]]:
    """Document-order list of ('paragraph'|'table', type-scoped index, element)."""
    items: list[tuple[str, int, etree._Element]] = []
    p_idx = t_idx = 0
    for child in pkg.body():
        tag = etree.QName(child).localname
        if tag == "p":
            items.append(("paragraph", p_idx, child))
            p_idx += 1
        elif tag == "tbl":
            items.append(("table", t_idx, child))
            t_idx += 1
    return items


# ---------------------------------------------------------------- top-level API


def get_document_info(pkg: DocxPackage) -> dict:
    paragraphs = tables = 0
    for kind, _, _ in body_items(pkg):
        if kind == "paragraph":
            paragraphs += 1
        else:
            tables += 1
    root = pkg.root()
    info = {
        "path": str(pkg.path),
        "paragraphs": paragraphs,
        "tables": tables,
        "sections": len(root.findall(f".//{qn('w:sectPr')}")),
        "footnotes": len(list_footnotes(pkg)),
        "endnotes": len(list_endnotes(pkg)),
        "comments": len(get_comments(pkg)),
        "revisions": revision_summary(pkg)["total"],
        "images": len(root.findall(f".//{qn('a:blip')}")),
        "parts": pkg.part_names(),
    }
    if pkg.has_part("docProps/core.xml"):
        core = pkg.root("docProps/core.xml")
        for tag, key in (
            ("{http://purl.org/dc/elements/1.1/}title", "title"),
            ("{http://purl.org/dc/elements/1.1/}creator", "author"),
            (
                "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy",
                "last_modified_by",
            ),
        ):
            el = core.find(tag)
            if el is not None and el.text:
                info[key] = el.text
    return info


def _toggle_on(rpr: etree._Element | None, tag: str) -> bool:
    if rpr is None:
        return False
    el = rpr.find(qn(tag))
    if el is None:
        return False
    return el.get(qn("w:val"), "1") not in ("0", "false", "none")


def _formatted_heading_level(p: etree._Element) -> int | None:
    """Heuristic heading level for a direct-formatted paragraph (the
    academic-template pattern: Normal style + bold/centered runs). Level 1
    for centered bold, 2 for bold, 3 for italic-only short paragraphs."""
    ppr = p.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return None  # list item, not a heading
    text = paragraph_text(p).strip()
    if not text or len(text) > 120 or len(text.split()) > 15:
        return None
    if text[-1] in ".,;":
        return None  # sentence-final punctuation: body prose
    runs = [
        r for r in p.iter(qn("w:r"))
        if run_text(r).strip()
    ]
    if not runs:
        return None
    all_bold = all(_toggle_on(r.find(qn("w:rPr")), "w:b") for r in runs)
    all_italic = all(_toggle_on(r.find(qn("w:rPr")), "w:i") for r in runs)
    if not all_bold and not all_italic:
        return None
    centered = False
    if ppr is not None:
        jc = ppr.find(qn("w:jc"))
        centered = jc is not None and jc.get(qn("w:val")) == "center"
    if all_bold:
        return 1 if centered else 2
    return 3


def get_outline(
    pkg: DocxPackage, *, detect_formatted: bool = False
) -> list[dict]:
    """Headings in document order. Detects BOTH heading systems: built-in
    Heading styles AND w:outlineLvl overrides (direct pPr or inherited
    through the style's basedOn chain — the pattern academic templates use
    on Normal-styled paragraphs, which Word's navigation pane honors).
    detected_via per entry: "heading_style" | "outline_level" |
    "formatting_heuristic" (with detect_formatted=True, short bold or
    italic direct-formatted paragraphs join the outline)."""
    style_outline = _style_outline_map(pkg)
    out = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph":
            continue
        level, via = _outline_level_detected(el, style_outline)
        if level is None and detect_formatted:
            level = _formatted_heading_level(el)
            via = "formatting_heuristic" if level is not None else None
        if level is not None:
            text = paragraph_text(el).strip()
            if text:
                out.append(
                    {
                        "paragraph_index": idx,
                        "level": level,
                        "text": text,
                        "detected_via": via,
                    }
                )
    return out


def get_outline_report(
    pkg: DocxPackage, *, detect_formatted: bool = False
) -> list | dict:
    """get_outline for the tool surface: same flat list when headings are
    found; when nothing is detected, a dict with an honest note plus flat
    structure counts instead of a bare empty list (field test,
    2026-09-03)."""
    out = get_outline(pkg, detect_formatted=detect_formatted)
    if out:
        return out
    n_paras = n_tables = n_words = 0
    for kind, _idx, el in body_items(pkg):
        if kind == "paragraph":
            n_paras += 1
            n_words += len(paragraph_text(el).split())
        else:
            n_tables += 1
    return {
        "headings": [],
        "note": (
            "0 headings detected via Heading styles or outlineLvl; the "
            "document may use direct formatting (bold/centered Normal "
            "paragraphs) for structure"
            + (
                ""
                if detect_formatted
                else ". Re-run with detect_formatted=true for a heuristic "
                "scan, or use set_paragraph_format's outline_level to tag "
                "the headings durably"
            )
        ),
        "structure": {
            "paragraphs": n_paras,
            "tables": n_tables,
            "approx_words": n_words,
        },
    }


def _textbox_entries(pkg: DocxPackage) -> list[dict]:
    """Per-box read entries via ops/textboxes.py, the one sanctioned reader.
    NEVER re-include w:txbxContent in the body walk instead: Word stores each
    modern box twice (mc:Choice + mc:Fallback) and the body walk would smear
    the text into the host paragraph doubled (the v1.5 doubled-text bug)."""
    from .textboxes import _enumerate_boxes

    out = []
    for i, b in enumerate(_enumerate_boxes(pkg)):
        paras = [paragraph_text(p) for p in b["element"].findall(qn("w:p"))]
        out.append(
            {
                "box_index": i,
                "part": b["part"],
                "shape_name": b["shape_name"],
                "anchor_paragraph_index": b["anchor_paragraph_index"],
                "text": "\n".join(paras),
            }
        )
    return out


def get_paragraphs(
    pkg: DocxPackage,
    start: int = 0,
    end: int | None = None,
    *,
    contains: str | None = None,
    include_sdt: bool = True,
    include_textboxes: bool = False,
) -> list[dict]:
    if start < 0:
        from ..core.errors import WordMcpError

        raise WordMcpError("start must be >= 0")
    style_outline = _style_outline_map(pkg)
    default_style = default_paragraph_style(pkg)
    out = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph" or idx < start or (end is not None and idx >= end):
            continue
        text = paragraph_text(el)
        if contains is not None and contains not in text:
            continue
        entry = {"index": idx, "text": text}
        # Always report the EFFECTIVE style: an absent pStyle means the
        # document's default paragraph style (usually Normal), and omitting
        # it made "explicitly Normal" and "no style info" indistinguishable.
        entry["style"] = _style_id(el) or default_style
        lvl = _outline_level(el, style_outline)
        if lvl is not None:
            entry["heading_level"] = lvl
        out.append(entry)
    # Block-level SDT content (TOC, caption lists, gallery bibliographies)
    # is not body-addressable (index None) but must be READABLE. Only listed
    # on a whole-document read — a slice request is about body indices, and
    # the live layer behaves the same way.
    if include_sdt and start == 0 and end is None:
        for sdt in pkg.body().iter(qn("w:sdt")):
            content = sdt.find(qn("w:sdtContent"))
            if content is None:
                continue
            for p in content.findall(qn("w:p")):
                text = paragraph_text(p)
                if contains is not None and contains not in text:
                    continue
                if not text.strip():
                    continue
                entry = {"index": None, "in_sdt": True, "text": text}
                entry["style"] = _style_id(p) or default_style
                out.append(entry)
    # Text-box content: clearly-labeled ADDITIONAL entries (source "textbox",
    # index None) read via ops/textboxes.py. Body paragraph indices are
    # unaffected by this flag; box_index is the address for set_textbox_text.
    if include_textboxes:
        for box in _textbox_entries(pkg):
            text = box["text"]
            if not text.strip():
                continue
            if contains is not None and contains not in text:
                continue
            out.append(
                {
                    "index": None,
                    "source": "textbox",
                    "box_index": box["box_index"],
                    "part": box["part"],
                    "shape_name": box["shape_name"],
                    "anchor_paragraph_index": box["anchor_paragraph_index"],
                    "text": text,
                }
            )
    return out


# -------------------------------------------------------------------- tables


def _cell_info(tc: etree._Element) -> dict:
    tcpr = tc.find(qn("w:tcPr"))
    grid_span = 1
    vmerge = None
    if tcpr is not None:
        gs = tcpr.find(qn("w:gridSpan"))
        if gs is not None:
            grid_span = int(gs.get(qn("w:val"), "1"))
        vm = tcpr.find(qn("w:vMerge"))
        if vm is not None:
            vmerge = vm.get(qn("w:val"), "continue")
    text = "\n".join(
        paragraph_text(p) for p in tc.findall(qn("w:p"))
    )
    return {"text": text, "grid_span": grid_span, "vmerge": vmerge}


def get_table(pkg: DocxPackage, table_index: int) -> dict:
    for kind, idx, el in body_items(pkg):
        if kind == "table" and idx == table_index:
            return _table_dict(el, table_index)
    from ..core.errors import TargetNotFound

    raise TargetNotFound(f"no body-level table with index {table_index}")


def _table_dict(tbl: etree._Element, index: int) -> dict:
    grid = [
        int(gc.get(qn("w:w"), "0"))
        for gc in tbl.findall(f"{qn('w:tblGrid')}/{qn('w:gridCol')}")
    ]
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        rows.append([_cell_info(tc) for tc in tr.findall(qn("w:tc"))])
    has_merges = any(
        c["grid_span"] > 1 or c["vmerge"] is not None for row in rows for c in row
    )
    return {
        "index": index,
        "rows": len(rows),
        "grid_columns": len(grid),
        "grid_widths_twips": grid,
        "has_merges": has_merges,
        "cells": rows,
    }


def list_tables(pkg: DocxPackage) -> list[dict]:
    out = []
    for kind, idx, el in body_items(pkg):
        if kind != "table":
            continue
        d = _table_dict(el, idx)
        preview = [c["text"][:40] for c in d["cells"][0]] if d["cells"] else []
        out.append(
            {
                "index": idx,
                "rows": d["rows"],
                "grid_columns": d["grid_columns"],
                "has_merges": d["has_merges"],
                "header_preview": preview,
            }
        )
    return out


# ------------------------------------------------------------- notes (read)

_SEPARATOR_TYPES = {"separator", "continuationSeparator", "continuationNotice"}


def _list_notes(pkg: DocxPackage, part: str, tag: str, ref_tag: str) -> list[dict]:
    if not pkg.has_part(part):
        return []
    # Reference order in the body, for position info.
    ref_order = [
        ref.get(qn("w:id"))
        for ref in pkg.root().iter(qn(ref_tag))
        if ref.get(qn("w:id")) is not None
    ]
    out = []
    for note in pkg.root(part).findall(qn(tag)):
        ntype = note.get(qn("w:type"))
        if ntype in _SEPARATOR_TYPES:
            continue
        nid = note.get(qn("w:id"))
        text = "\n".join(
            paragraph_text(p) for p in note.iter(qn("w:p"))
        ).strip()
        pos = ref_order.index(nid) + 1 if nid in ref_order else None
        out.append({"id": nid, "position": pos, "text": text})
    out.sort(key=lambda n: (n["position"] is None, n["position"]))
    return out


def list_footnotes(pkg: DocxPackage) -> list[dict]:
    return _list_notes(
        pkg, "word/footnotes.xml", "w:footnote", "w:footnoteReference"
    )


def list_endnotes(pkg: DocxPackage) -> list[dict]:
    return _list_notes(pkg, "word/endnotes.xml", "w:endnote", "w:endnoteReference")


# ------------------------------------------------------------ comments (read)


def get_comments(pkg: DocxPackage, *, author: str | None = None) -> list[dict]:
    if not pkg.has_part("word/comments.xml"):
        return []
    # Anchored text: map comment id -> text between commentRangeStart/End.
    anchors, deleted_anchors = _comment_anchors(pkg)
    # Threading + resolved state from commentsExtended (matched via paraId of
    # the comment's last paragraph).
    para_meta = _comments_extended(pkg)

    out = []
    for c in pkg.root("word/comments.xml").findall(qn("w:comment")):
        cid = c.get(qn("w:id"))
        entry = {
            "id": cid,
            "author": c.get(qn("w:author"), ""),
            "initials": c.get(qn("w:initials"), ""),
            "date": c.get(qn("w:date"), ""),
            "text": "\n".join(paragraph_text(p) for p in c.findall(qn("w:p"))).strip(),
            "anchored_text": anchors.get(cid, ""),
        }
        # Anchor range fully inside a pending tracked deletion: report the
        # deleted text plus an honest flag instead of a bare empty string
        # (field test, 2026-09-03).
        if not entry["anchored_text"] and deleted_anchors.get(cid):
            entry["anchored_text"] = deleted_anchors[cid]
            entry["anchor_deleted"] = True
        paras = c.findall(qn("w:p"))
        last_para_id = paras[-1].get(qn("w14:paraId")) if paras else None
        meta = para_meta.get(last_para_id, {})
        entry["resolved"] = meta.get("done", False)
        if meta.get("parent_para_id") is not None:
            entry["reply_to_para_id"] = meta["parent_para_id"]
        entry["para_id"] = last_para_id
        out.append(entry)

    # Convert reply_to_para_id -> reply_to comment id.
    by_para = {e["para_id"]: e["id"] for e in out if e["para_id"]}
    for e in out:
        parent = e.pop("reply_to_para_id", None)
        e["reply_to"] = by_para.get(parent)
        e.pop("para_id", None)
    if author is not None:
        out = [e for e in out if e["author"] == author]
    return out


def _comment_anchors(
    pkg: DocxPackage,
) -> tuple[dict[str, str], dict[str, str]]:
    """(comment id -> visible text spanned by its commentRangeStart/End,
    comment id -> tracked-DELETED text in that span). The second map lets
    get_comments report an anchor that survives only inside a pending
    tracked deletion instead of returning an empty string."""
    anchors: dict[str, list[str]] = {}
    deleted: dict[str, list[str]] = {}
    open_ids: set[str] = set()
    root = pkg.root()
    for el in root.iter():
        tag = etree.QName(el).localname
        if tag == "commentRangeStart":
            open_ids.add(el.get(qn("w:id")))
            anchors.setdefault(el.get(qn("w:id")), [])
        elif tag == "commentRangeEnd":
            open_ids.discard(el.get(qn("w:id")))
        elif tag == "r" and open_ids:
            text = run_text(el)
            if text:
                for cid in open_ids:
                    anchors[cid].append(text)
            else:
                del_text = run_text(el, include_deleted=True)
                if del_text:
                    for cid in open_ids:
                        deleted.setdefault(cid, []).append(del_text)
    return (
        {cid: "".join(parts) for cid, parts in anchors.items()},
        {cid: "".join(parts) for cid, parts in deleted.items()},
    )


def _comments_extended(pkg: DocxPackage) -> dict[str, dict]:
    """paraId -> {done, parent_para_id} from commentsExtended.xml."""
    part = "word/commentsExtended.xml"
    if not pkg.has_part(part):
        return {}
    w15 = NSMAP["w15"]
    out: dict[str, dict] = {}
    for ex in pkg.root(part).findall(f"{{{w15}}}commentEx"):
        pid = ex.get(f"{{{w15}}}paraId")
        out[pid] = {
            "done": ex.get(f"{{{w15}}}done") == "1",
            "parent_para_id": ex.get(f"{{{w15}}}paraIdParent"),
        }
    return out


# ------------------------------------------------------------ revisions (read)

_REVISION_TAGS = {
    "ins": "insertion",
    "del": "deletion",
    "rPrChange": "format_change_run",
    "pPrChange": "format_change_paragraph",
    "tblPrChange": "format_change_table",
    "sectPrChange": "format_change_section",
    "moveFrom": "move_from",
    "moveTo": "move_to",
    "cellIns": "cell_insertion",
    "cellDel": "cell_deletion",
}


def get_tracked_changes(pkg: DocxPackage, *, author: str | None = None) -> list[dict]:
    out = []
    parts = ["word/document.xml"]
    for extra in ("word/footnotes.xml", "word/endnotes.xml"):
        if pkg.has_part(extra):
            parts.append(extra)
    for part in parts:
        for el in pkg.root(part).iter():
            local = etree.QName(el).localname
            if local not in _REVISION_TAGS:
                continue
            rev_author = el.get(qn("w:author"), "")
            if author is not None and rev_author != author:
                continue
            if local == "ins":
                text = "".join(run_text(r) for r in el.iter(qn("w:r")))
            elif local in ("del", "moveFrom"):
                text = "".join(
                    run_text(r, include_deleted=True) for r in el.iter(qn("w:r"))
                )
            else:
                text = ""
            out.append(
                {
                    "type": _REVISION_TAGS[local],
                    "author": rev_author,
                    "date": el.get(qn("w:date"), ""),
                    "text": text,
                    "part": part,
                }
            )
    return out


def revision_summary(pkg: DocxPackage) -> dict:
    revs = get_tracked_changes(pkg)
    by_author: dict[str, int] = {}
    by_type: dict[str, int] = {}
    for r in revs:
        by_author[r["author"]] = by_author.get(r["author"], 0) + 1
        by_type[r["type"]] = by_type.get(r["type"], 0) + 1
    return {"total": len(revs), "by_author": by_author, "by_type": by_type}


# --------------------------------------------------------------- styles (read)


def _pt(twips: str | None, per_pt: int = 20) -> float | int | None:
    """Twips (or half-points etc.) -> points, integers kept integral."""
    if twips is None:
        return None
    v = int(twips) / per_pt
    return int(v) if v == int(v) else v


def _on_off(el: etree._Element | None) -> bool | None:
    """OOXML on/off element: absent -> None, val absent -> True, else by val."""
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val is None:
        return True
    return val not in ("0", "false", "none")


_JC_TO_ALIGNMENT = {
    "left": "left", "start": "left", "center": "center",
    "right": "right", "end": "right", "both": "justify",
}


def _style_paragraph_formatting(ppr: etree._Element | None) -> dict:
    """A style's EXPLICIT paragraph formatting, in the exact shape
    define_style's paragraph_formatting parameter accepts (inherited values
    are not synthesized — based_on carries the chain). Line spacing with an
    exact/atLeast rule has no define_style representation and is omitted."""
    if ppr is None:
        return {}
    fmt: dict = {}
    jc = ppr.find(qn("w:jc"))
    if jc is not None:
        alignment = _JC_TO_ALIGNMENT.get(jc.get(qn("w:val")))
        if alignment:
            fmt["alignment"] = alignment
    sp = ppr.find(qn("w:spacing"))
    if sp is not None:
        before = _pt(sp.get(qn("w:before")))
        if before is not None:
            fmt["space_before_pt"] = before
        after = _pt(sp.get(qn("w:after")))
        if after is not None:
            fmt["space_after_pt"] = after
        line = sp.get(qn("w:line"))
        if line is not None and sp.get(qn("w:lineRule"), "auto") == "auto":
            fmt["line_spacing"] = _pt(line, 240)
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        left = _pt(ind.get(qn("w:left")) or ind.get(qn("w:start")))
        if left is not None:
            fmt["indent_left_pt"] = left
        right = _pt(ind.get(qn("w:right")) or ind.get(qn("w:end")))
        if right is not None:
            fmt["indent_right_pt"] = right
        first = _pt(ind.get(qn("w:firstLine")))
        hanging = _pt(ind.get(qn("w:hanging")))
        if hanging is not None:
            fmt["first_line_indent_pt"] = -hanging
        elif first is not None:
            fmt["first_line_indent_pt"] = first
    for key, tag in (
        ("keep_with_next", "w:keepNext"),
        ("keep_lines_together", "w:keepLines"),
        ("page_break_before", "w:pageBreakBefore"),
        ("widow_control", "w:widowControl"),
    ):
        v = _on_off(ppr.find(qn(tag)))
        if v is not None:
            fmt[key] = v
    return fmt


def _style_character_formatting(rpr: etree._Element | None) -> dict:
    """A style's EXPLICIT run formatting, in the exact shape define_style's
    character_formatting parameter accepts. Theme font/color references have
    no concrete file-local value and are omitted rather than guessed."""
    if rpr is None:
        return {}
    fmt: dict = {}
    for key, tag in (
        ("bold", "w:b"), ("italic", "w:i"), ("strike", "w:strike"),
        ("small_caps", "w:smallCaps"), ("all_caps", "w:caps"),
        ("hidden", "w:vanish"), ("double_strike", "w:dstrike"),
    ):
        v = _on_off(rpr.find(qn(tag)))
        if v is not None:
            fmt[key] = v
    u = rpr.find(qn("w:u"))
    if u is not None:
        fmt["underline"] = u.get(qn("w:val")) != "none"
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        name = rfonts.get(qn("w:ascii"))
        if name:
            fmt["font"] = name
    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        fmt["size_pt"] = _pt(sz.get(qn("w:val")), 2)
    color = rpr.find(qn("w:color"))
    if color is not None:
        val = color.get(qn("w:val"))
        if val and val.lower() != "auto":
            fmt["color"] = val.upper()
    highlight = rpr.find(qn("w:highlight"))
    if highlight is not None and highlight.get(qn("w:val")):
        fmt["highlight"] = highlight.get(qn("w:val"))
    va = rpr.find(qn("w:vertAlign"))
    if va is not None:
        if va.get(qn("w:val")) == "superscript":
            fmt["superscript"] = True
        elif va.get(qn("w:val")) == "subscript":
            fmt["subscript"] = True
    csp = rpr.find(qn("w:spacing"))
    if csp is not None and csp.get(qn("w:val")):
        fmt["char_spacing_pt"] = _pt(csp.get(qn("w:val")))
    kern = rpr.find(qn("w:kern"))
    if kern is not None and kern.get(qn("w:val")):
        fmt["kerning_pt"] = _pt(kern.get(qn("w:val")), 2)
    pos = rpr.find(qn("w:position"))
    if pos is not None and pos.get(qn("w:val")):
        fmt["position_pt"] = _pt(pos.get(qn("w:val")), 2)
    lang = rpr.find(qn("w:lang"))
    if lang is not None:
        if lang.get(qn("w:val")):
            fmt["language"] = lang.get(qn("w:val"))
        if lang.get(qn("w:eastAsia")):
            fmt["east_asian_language"] = lang.get(qn("w:eastAsia"))
    return fmt


def list_styles(pkg: DocxPackage) -> list[dict]:
    if not pkg.has_part("word/styles.xml"):
        return []
    out = []
    for style in pkg.root("word/styles.xml").findall(qn("w:style")):
        name_el = style.find(qn("w:name"))
        based_el = style.find(qn("w:basedOn"))
        entry = {
            "id": style.get(qn("w:styleId")),
            "type": style.get(qn("w:type")),
            "name": name_el.get(qn("w:val")) if name_el is not None else None,
            "based_on": based_el.get(qn("w:val")) if based_el is not None else None,
        }
        pf = _style_paragraph_formatting(style.find(qn("w:pPr")))
        if pf:
            entry["paragraph_formatting"] = pf
        cf = _style_character_formatting(style.find(qn("w:rPr")))
        if cf:
            entry["character_formatting"] = cf
        out.append(entry)
    return out


def find_text(
    pkg: DocxPackage,
    query: str,
    *,
    regex: bool = False,
    context_chars: int = 60,
    include_textboxes: bool = False,
) -> list[dict]:
    from . import _regex
    from ..core.errors import WordMcpError

    if not query:
        raise WordMcpError("query must be non-empty")
    if regex and _regex.finditer(query, ""):
        raise WordMcpError(
            f"regex {query!r} can match the empty string and would match at "
            "every position; anchor the pattern"
        )
    matches = []
    pattern = _regex.compile_user_pattern(query) if regex else None
    for kind, idx, el in body_items(pkg):
        if kind == "paragraph":
            text = paragraph_text(el)
            spans = (
                [(m.start(), m.end()) for m in _regex.finditer(query, text)]
                if pattern
                else _literal_spans(text, query)
            )
            for start, end in spans:
                matches.append(
                    {
                        "paragraph_index": idx,
                        "match": text[start:end],
                        "context": text[
                            max(0, start - context_chars) : end + context_chars
                        ],
                    }
                )
        else:
            for r_i, tr in enumerate(el.findall(qn("w:tr"))):
                for c_i, tc in enumerate(tr.findall(qn("w:tc"))):
                    cell_text = "\n".join(
                        paragraph_text(p) for p in tc.findall(qn("w:p"))
                    )
                    spans = (
                        [
                            (m.start(), m.end())
                            for m in _regex.finditer(query, cell_text)
                        ]
                        if pattern
                        else _literal_spans(cell_text, query)
                    )
                    for start, end in spans:
                        matches.append(
                            {
                                "table_index": idx,
                                "row": r_i,
                                "cell": c_i,
                                "match": cell_text[start:end],
                                "context": cell_text[
                                    max(0, start - context_chars) : end + context_chars
                                ],
                            }
                        )
    # Text-box content: clearly-labeled ADDITIONAL matches (source "textbox")
    # read via ops/textboxes.py — never by re-including w:txbxContent in the
    # body walk (that doubled box text in v1.5). Body/table match entries and
    # their paragraph indices are unaffected by the flag.
    if include_textboxes:
        for box in _textbox_entries(pkg):
            box_text = box["text"]
            spans = (
                [(m.start(), m.end()) for m in _regex.finditer(query, box_text)]
                if pattern
                else _literal_spans(box_text, query)
            )
            for start, end in spans:
                matches.append(
                    {
                        "source": "textbox",
                        "box_index": box["box_index"],
                        "part": box["part"],
                        "shape_name": box["shape_name"],
                        "anchor_paragraph_index": box["anchor_paragraph_index"],
                        "match": box_text[start:end],
                        "context": box_text[
                            max(0, start - context_chars) : end + context_chars
                        ],
                    }
                )
    return matches


def _literal_spans(text: str, query: str) -> list[tuple[int, int]]:
    spans = []
    start = 0
    while True:
        pos = text.find(query, start)
        if pos < 0:
            break
        spans.append((pos, pos + len(query)))
        start = pos + 1
    return spans
