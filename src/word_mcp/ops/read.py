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
    """Full visible text of a w:p, descending into w:ins/hyperlinks/smartTags."""
    parts: list[str] = []
    for r in p.iter(qn("w:r")):
        # Runs inside w:del hold w:delText only; run_text already skips those
        # unless include_deleted is set.
        parts.append(run_text(r, include_deleted=include_deleted))
    return "".join(parts)


def _style_id(p: etree._Element) -> str | None:
    pstyle = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    return pstyle.get(qn("w:val")) if pstyle is not None else None


def _outline_level(p: etree._Element, style_outline: dict[str, int]) -> int | None:
    """Heading level 1-9, from direct outlineLvl, style name, or style's outlineLvl."""
    lvl = p.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
    if lvl is not None:
        return int(lvl.get(qn("w:val"))) + 1
    sid = _style_id(p)
    if sid:
        if sid in _HEADING_STYLES:
            return _HEADING_STYLES[sid]
        if sid in style_outline:
            return style_outline[sid]
    return None


def _style_outline_map(pkg: DocxPackage) -> dict[str, int]:
    """styleId -> heading level, from styles.xml outlineLvl definitions."""
    out: dict[str, int] = {}
    if not pkg.has_part("word/styles.xml"):
        return out
    for style in pkg.root("word/styles.xml").findall(qn("w:style")):
        sid = style.get(qn("w:styleId"))
        lvl = style.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
        if sid and lvl is not None:
            out[sid] = int(lvl.get(qn("w:val"))) + 1
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


def get_outline(pkg: DocxPackage) -> list[dict]:
    style_outline = _style_outline_map(pkg)
    out = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph":
            continue
        level = _outline_level(el, style_outline)
        if level is not None:
            text = paragraph_text(el).strip()
            if text:
                out.append({"paragraph_index": idx, "level": level, "text": text})
    return out


def get_paragraphs(
    pkg: DocxPackage,
    start: int = 0,
    end: int | None = None,
    *,
    contains: str | None = None,
    include_sdt: bool = True,
) -> list[dict]:
    if start < 0:
        from ..core.errors import WordMcpError

        raise WordMcpError("start must be >= 0")
    style_outline = _style_outline_map(pkg)
    out = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph" or idx < start or (end is not None and idx >= end):
            continue
        text = paragraph_text(el)
        if contains is not None and contains not in text:
            continue
        entry = {"index": idx, "text": text}
        sid = _style_id(el)
        if sid:
            entry["style"] = sid
        lvl = _outline_level(el, style_outline)
        if lvl is not None:
            entry["heading_level"] = lvl
        out.append(entry)
    # Block-level SDT content (TOC, caption lists, gallery bibliographies)
    # is not body-addressable (index None) but must be READABLE.
    if include_sdt:
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
                sid = _style_id(p)
                if sid:
                    entry["style"] = sid
                out.append(entry)
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
    anchors = _comment_anchors(pkg)
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


def _comment_anchors(pkg: DocxPackage) -> dict[str, str]:
    """comment id -> plain text spanned by its commentRangeStart/End."""
    anchors: dict[str, list[str]] = {}
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
    return {cid: "".join(parts) for cid, parts in anchors.items()}


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


def list_styles(pkg: DocxPackage) -> list[dict]:
    if not pkg.has_part("word/styles.xml"):
        return []
    out = []
    for style in pkg.root("word/styles.xml").findall(qn("w:style")):
        name_el = style.find(qn("w:name"))
        based_el = style.find(qn("w:basedOn"))
        out.append(
            {
                "id": style.get(qn("w:styleId")),
                "type": style.get(qn("w:type")),
                "name": name_el.get(qn("w:val")) if name_el is not None else None,
                "based_on": based_el.get(qn("w:val")) if based_el is not None else None,
            }
        )
    return out


def find_text(
    pkg: DocxPackage, query: str, *, regex: bool = False, context_chars: int = 60
) -> list[dict]:
    from . import _regex

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
