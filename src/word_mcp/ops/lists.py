"""Bulleted and numbered lists.

Real bullets/numbers are NOT a style: each list paragraph carries
w:pPr/w:numPr/(w:ilvl, w:numId), where numId points into word/numbering.xml
(w:num -> w:abstractNum with per-level formats). ListParagraph style alone
renders indented plain text — the bug this module exists to prevent.

Each add_list call creates its own w:num instance, so numbered lists restart
at 1 per call (Word's behavior for distinct lists).
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# (numFmt, lvlText, font) cycling across levels 0-8.
_BULLET_LEVELS = [
    ("bullet", "", "Symbol"),
    ("bullet", "o", "Courier New"),
    ("bullet", "", "Wingdings"),
]
_NUMBER_LEVELS = [
    ("decimal", "%{n}.", None),
    ("lowerLetter", "%{n}.", None),
    ("lowerRoman", "%{n}.", None),
]


def _ensure_numbering_part(pkg: DocxPackage) -> None:
    part = "word/numbering.xml"
    if pkg.has_part(part):
        return
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = etree.Element(qn("w:numbering"), nsmap={"w": w_ns})
    pkg.set_raw_part(
        part,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )
    ct_root = pkg.root("[Content_Types].xml")
    if not any(
        o.get("PartName") == "/" + part
        for o in ct_root.findall(f"{{{_CT_NS}}}Override")
    ):
        override = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        override.set("PartName", "/" + part)
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.numbering+xml",
        )
        pkg.mark_dirty("[Content_Types].xml")
    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/"
        "relationships/numbering"
    )
    if not any(r.get("Type") == rel_type for r in rels_root):
        existing = {r.get("Id") for r in rels_root}
        i = 1
        while f"rId{i}" in existing:
            i += 1
        rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
        rel.set("Id", f"rId{i}")
        rel.set("Type", rel_type)
        rel.set("Target", "numbering.xml")
        pkg.mark_dirty(rels_part)


def _new_numbering(pkg: DocxPackage, kind: str) -> int:
    """Create an abstractNum + num pair; return the numId to reference."""
    root = pkg.root("word/numbering.xml")
    abstract_ids = [
        int(a.get(qn("w:abstractNumId"), "0"))
        for a in root.findall(qn("w:abstractNum"))
    ]
    num_ids = [
        int(n.get(qn("w:numId"), "0")) for n in root.findall(qn("w:num"))
    ]
    abs_id = max(abstract_ids, default=-1) + 1
    num_id = max(num_ids, default=0) + 1

    levels = _BULLET_LEVELS if kind == "bullet" else _NUMBER_LEVELS
    abstract = etree.Element(qn("w:abstractNum"))
    abstract.set(qn("w:abstractNumId"), str(abs_id))
    ml = etree.SubElement(abstract, qn("w:multiLevelType"))
    ml.set(qn("w:val"), "hybridMultilevel")
    for ilvl in range(9):
        num_fmt, lvl_text, font = levels[ilvl % len(levels)]
        lvl = etree.SubElement(abstract, qn("w:lvl"))
        lvl.set(qn("w:ilvl"), str(ilvl))
        start = etree.SubElement(lvl, qn("w:start"))
        start.set(qn("w:val"), "1")
        fmt = etree.SubElement(lvl, qn("w:numFmt"))
        fmt.set(qn("w:val"), num_fmt)
        text_el = etree.SubElement(lvl, qn("w:lvlText"))
        text_el.set(qn("w:val"), lvl_text.replace("{n}", str(ilvl + 1)))
        jc = etree.SubElement(lvl, qn("w:lvlJc"))
        jc.set(qn("w:val"), "left")
        ppr = etree.SubElement(lvl, qn("w:pPr"))
        ind = etree.SubElement(ppr, qn("w:ind"))
        ind.set(qn("w:left"), str(720 * (ilvl + 1)))
        ind.set(qn("w:hanging"), "360")
        if font:
            rpr = etree.SubElement(lvl, qn("w:rPr"))
            rfonts = etree.SubElement(rpr, qn("w:rFonts"))
            for attr in ("w:ascii", "w:hAnsi", "w:hint"):
                rfonts.set(
                    qn(attr), font if attr != "w:hint" else "default"
                )

    # abstractNum elements must precede num elements in numbering.xml.
    nums = root.findall(qn("w:num"))
    if nums:
        nums[0].addprevious(abstract)
    else:
        root.append(abstract)
    num = etree.SubElement(root, qn("w:num"))
    num.set(qn("w:numId"), str(num_id))
    ref = etree.SubElement(num, qn("w:abstractNumId"))
    ref.set(qn("w:val"), str(abs_id))
    pkg.mark_dirty("word/numbering.xml")
    return num_id


def _ensure_list_style(pkg: DocxPackage) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    if "ListParagraph" in have:
        return
    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), "paragraph")
    s.set(qn("w:styleId"), "ListParagraph")
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), "List Paragraph")
    etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
    ppr = etree.SubElement(s, qn("w:pPr"))
    ind = etree.SubElement(ppr, qn("w:ind"))
    ind.set(qn("w:left"), "720")
    cs = etree.SubElement(ppr, qn("w:contextualSpacing"))
    pkg.mark_dirty("word/styles.xml")


def add_list(
    pkg: DocxPackage,
    items: list,
    *,
    kind: str = "bullet",
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
) -> dict:
    """Insert a bulleted or numbered list. Items: strings, or dicts
    {text, level} with level 0-8 for nesting. Each call is an independent
    list (numbering restarts at 1)."""
    if kind not in ("bullet", "number"):
        raise WordMcpError("kind must be 'bullet' or 'number'")
    if not items:
        raise WordMcpError("items must be a non-empty list")

    norm: list[tuple[str, int]] = []
    for item in items:
        if isinstance(item, dict):
            level = int(item.get("level", 0))
            if not 0 <= level <= 8:
                raise WordMcpError("level must be 0-8")
            norm.append((str(item["text"]), level))
        else:
            norm.append((str(item), 0))

    paragraphs, num_id = build_list_paragraphs(pkg, norm, kind)

    from .text import _body_paragraph, _resolve_anchor

    body = pkg.body()
    if at_end or (after_index is None and after_anchor is None):
        sectpr = body.find(qn("w:sectPr"))
        for p in paragraphs:
            if sectpr is not None:
                sectpr.addprevious(p)
            else:
                body.append(p)
    elif after_anchor is not None:
        ref = _resolve_anchor(pkg, after_anchor)
        for p in reversed(paragraphs):
            ref.addnext(p)
    else:
        ref = _body_paragraph(pkg, after_index)
        for p in reversed(paragraphs):
            ref.addnext(p)
    pkg.mark_dirty()
    return {"list_added": kind, "items": len(norm), "num_id": num_id}


def build_list_paragraphs(
    pkg: DocxPackage, norm: list[tuple[str, int]], kind: str
) -> tuple[list[etree._Element], int]:
    """Build (without inserting) the list's w:p elements, registering a
    fresh numbering instance in the package. norm: [(text, level 0-8)].
    Extracted from add_list so the batch layer's markdown list inserts
    share one construction path. Mutates numbering/styles parts only; the
    caller splices the paragraphs and calls mark_dirty()."""
    _ensure_numbering_part(pkg)
    _ensure_list_style(pkg)
    num_id = _new_numbering(pkg, kind)

    from . import _runmap

    paragraphs = []
    for text, level in norm:
        p = etree.Element(qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        pstyle = etree.SubElement(ppr, qn("w:pStyle"))
        pstyle.set(qn("w:val"), "ListParagraph")
        numpr = etree.SubElement(ppr, qn("w:numPr"))
        ilvl = etree.SubElement(numpr, qn("w:ilvl"))
        ilvl.set(qn("w:val"), str(level))
        nid = etree.SubElement(numpr, qn("w:numId"))
        nid.set(qn("w:val"), str(num_id))
        r = etree.SubElement(p, qn("w:r"))
        if text:
            t = etree.SubElement(r, qn("w:t"))
            t.text = text
            _runmap._preserve_space(t)
        paragraphs.append(p)
    return paragraphs, num_id


def get_lists(pkg: DocxPackage) -> list[dict]:
    """List paragraphs grouped by numbering instance (numId)."""
    from .read import body_items, paragraph_text

    out: dict[int, list] = {}
    for kind_, idx, el in body_items(pkg):
        if kind_ != "paragraph":
            continue
        numpr = el.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
        if numpr is None:
            continue
        nid_el = numpr.find(qn("w:numId"))
        ilvl_el = numpr.find(qn("w:ilvl"))
        if nid_el is None:
            continue
        nid = int(nid_el.get(qn("w:val"), "0"))
        out.setdefault(nid, []).append(
            {
                "paragraph_index": idx,
                "level": int(ilvl_el.get(qn("w:val"), "0")) if ilvl_el is not None else 0,
                "text": paragraph_text(el),
            }
        )
    return [{"num_id": nid, "items": items} for nid, items in sorted(out.items())]
