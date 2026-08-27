"""Structural operations: moving whole heading sections, document properties,
style definition, character styles, image alt text."""

from __future__ import annotations

import datetime

from lxml import etree

from ..core.errors import AmbiguousTarget, TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from .read import _outline_level, _style_outline_map, body_items, paragraph_text


# ------------------------------------------------------------- section moving


def _find_heading(pkg: DocxPackage, heading_text: str) -> tuple[int, int]:
    """(body child position, level) of the heading paragraph matching text."""
    style_outline = _style_outline_map(pkg)
    matches = []
    body = pkg.body()
    for pos, child in enumerate(body):
        if etree.QName(child).localname != "p":
            continue
        level = _outline_level(child, style_outline)
        if level is None:
            continue
        if paragraph_text(child).strip() == heading_text.strip():
            matches.append((pos, level))
    if not matches:
        raise TargetNotFound(f"no heading with text {heading_text!r}")
    if len(matches) > 1:
        raise AmbiguousTarget(
            f"{len(matches)} headings match {heading_text!r}; make it unique"
        )
    return matches[0]


def _section_block(pkg: DocxPackage, heading_text: str) -> list[etree._Element]:
    """The heading element plus every following body child until the next
    heading of the same or higher level."""
    style_outline = _style_outline_map(pkg)
    body = pkg.body()
    children = list(body)
    start, level = _find_heading(pkg, heading_text)
    block = [children[start]]
    for child in children[start + 1 :]:
        local = etree.QName(child).localname
        if local == "sectPr":
            break
        if local == "p":
            child_level = _outline_level(child, style_outline)
            if child_level is not None and child_level <= level:
                break
        block.append(child)
    return block


def _check_block_movable(block: list[etree._Element]) -> None:
    depth = 0
    for el in block:
        if etree.QName(el).localname == "p" and el.find(
            f"{qn('w:pPr')}/{qn('w:sectPr')}"
        ) is not None:
            raise WordMcpError(
                "the section contains a section break; moving it would "
                "restructure page layout — move content without the break"
            )
        for fc in el.iter(qn("w:fldChar")):
            t = fc.get(qn("w:fldCharType"))
            if t == "begin":
                depth += 1
            elif t == "end":
                depth -= 1
    if depth != 0:
        raise WordMcpError(
            "a field (TOC/PAGEREF/...) crosses the section boundary; "
            "moving would corrupt it"
        )


def move_section(
    pkg: DocxPackage,
    heading_text: str,
    *,
    before_heading: str | None = None,
    after_heading: str | None = None,
    at_end: bool = False,
) -> dict:
    """Move a heading and its ENTIRE section (everything until the next
    heading of the same or higher level, tables included) to a new location."""
    specified = sum(
        [before_heading is not None, after_heading is not None, bool(at_end)]
    )
    if specified != 1:
        raise WordMcpError(
            "specify exactly one of before_heading, after_heading, at_end"
        )
    block = _section_block(pkg, heading_text)
    _check_block_movable(block)
    block_set = {id(el) for el in block}
    body = pkg.body()

    if at_end:
        sectpr = body.find(qn("w:sectPr"))
        for el in block:
            if sectpr is not None:
                sectpr.addprevious(el)
            else:
                body.append(el)
    else:
        if before_heading is not None:
            anchor_el = body[ _find_heading(pkg, before_heading)[0] ]
            if id(anchor_el) in block_set:
                raise WordMcpError("destination lies inside the moved section")
            for el in block:
                anchor_el.addprevious(el)
        else:
            dest_block = _section_block(pkg, after_heading)
            if any(id(el) in block_set for el in dest_block):
                raise WordMcpError("destination lies inside the moved section")
            anchor_el = dest_block[-1]
            for el in reversed(block):
                anchor_el.addnext(el)
    pkg.mark_dirty()
    return {
        "moved": heading_text,
        "elements": len(block),
        "to": before_heading or after_heading or "end",
    }


def list_section_blocks(pkg: DocxPackage) -> list[dict]:
    """Headings with the size of each one's section block (for move planning)."""
    from .read import get_outline

    outline = get_outline(pkg)
    out = []
    for h in outline:
        try:
            block = _section_block(pkg, h["text"])
            out.append(
                {
                    "heading": h["text"],
                    "level": h["level"],
                    "elements": len(block),
                }
            )
        except AmbiguousTarget:
            out.append(
                {
                    "heading": h["text"],
                    "level": h["level"],
                    "elements": None,
                    "note": "duplicate heading text; not addressable by name",
                }
            )
    return out


# -------------------------------------------------------- document properties

_DC = "http://purl.org/dc/elements/1.1/"
_CP = (
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
)
_DCTERMS = "http://purl.org/dc/terms/"
_XSI = "http://www.w3.org/2001/XMLSchema-instance"


def set_document_properties(
    pkg: DocxPackage,
    *,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    comments: str | None = None,
) -> dict:
    """Set core document properties (File > Info metadata)."""
    part = "docProps/core.xml"
    if not pkg.has_part(part):
        root = etree.Element(
            f"{{{_CP}}}coreProperties",
            nsmap={
                "cp": _CP, "dc": _DC, "dcterms": _DCTERMS, "xsi": _XSI,
            },
        )
        pkg.set_raw_part(
            part,
            etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            ),
        )
    root = pkg.root(part)

    def set_el(ns: str, local: str, value: str | None, w3cdtf: bool = False):
        if value is None:
            return None
        el = root.find(f"{{{ns}}}{local}")
        if el is None:
            el = etree.SubElement(root, f"{{{ns}}}{local}")
        el.text = value
        if w3cdtf:
            el.set(f"{{{_XSI}}}type", "dcterms:W3CDTF")
        return local

    changed = [
        x
        for x in (
            set_el(_DC, "title", title),
            set_el(_DC, "creator", author),
            set_el(_DC, "subject", subject),
            set_el(_CP, "keywords", keywords),
            set_el(_CP, "category", category),
            set_el(_DC, "description", comments),
            set_el(
                _DCTERMS,
                "modified",
                datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"),
                w3cdtf=True,
            ),
        )
        if x
    ]
    pkg.mark_dirty(part)
    return {"properties_set": changed}


# ------------------------------------------------------------- style creation


def define_style(
    pkg: DocxPackage,
    *,
    style_id: str,
    name: str,
    style_type: str = "paragraph",
    based_on: str | None = "Normal",
    next_style: str | None = None,
    character_formatting: dict | None = None,
    paragraph_formatting: dict | None = None,
) -> dict:
    """Create (or replace) a custom style. character_formatting takes the
    format_text keys; paragraph_formatting takes alignment / spacing / indent
    keys from set_paragraph_format."""
    import re

    if style_type not in ("paragraph", "character"):
        raise WordMcpError("style_type must be paragraph or character")
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_\-]{0,49}", style_id):
        raise WordMcpError("style_id: letters/digits/_/-, starting with a letter")
    if style_type == "character" and paragraph_formatting:
        raise WordMcpError("character styles cannot carry paragraph formatting")

    root = pkg.root("word/styles.xml")
    for s in root.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == style_id:
            root.remove(s)

    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), style_type)
    s.set(qn("w:styleId"), style_id)
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), name)
    if based_on:
        etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), based_on)
    if next_style and style_type == "paragraph":
        etree.SubElement(s, qn("w:next")).set(qn("w:val"), next_style)
    etree.SubElement(s, qn("w:qFormat"))

    if paragraph_formatting:
        from .text import _PARA_FMT_KEYS, _check_keys

        allowed = _PARA_FMT_KEYS - {"borders", "shading", "tab_stops"}
        _check_keys(paragraph_formatting, allowed, "paragraph-formatting")
        ppr = etree.SubElement(s, qn("w:pPr"))
        pf = paragraph_formatting
        if "alignment" in pf:
            from .text import _ALIGN

            etree.SubElement(ppr, qn("w:jc")).set(qn("w:val"), _ALIGN[pf["alignment"]])
        if any(k in pf for k in ("space_before_pt", "space_after_pt", "line_spacing")):
            sp = etree.SubElement(ppr, qn("w:spacing"))
            if "space_before_pt" in pf:
                sp.set(qn("w:before"), str(int(pf["space_before_pt"] * 20)))
            if "space_after_pt" in pf:
                sp.set(qn("w:after"), str(int(pf["space_after_pt"] * 20)))
            if "line_spacing" in pf:
                sp.set(qn("w:line"), str(int(pf["line_spacing"] * 240)))
                sp.set(qn("w:lineRule"), "auto")
        if any(k in pf for k in ("indent_left_pt", "indent_right_pt", "first_line_indent_pt")):
            ind = etree.SubElement(ppr, qn("w:ind"))
            if "indent_left_pt" in pf:
                ind.set(qn("w:left"), str(int(pf["indent_left_pt"] * 20)))
            if "indent_right_pt" in pf:
                ind.set(qn("w:right"), str(int(pf["indent_right_pt"] * 20)))
            if "first_line_indent_pt" in pf:
                v = pf["first_line_indent_pt"]
                if v >= 0:
                    ind.set(qn("w:firstLine"), str(int(v * 20)))
                else:
                    ind.set(qn("w:hanging"), str(int(-v * 20)))

    if character_formatting:
        from .text import _make_rpr

        s.append(_make_rpr(character_formatting))

    pkg.mark_dirty("word/styles.xml")
    return {"style_defined": style_id, "type": style_type}


def apply_character_style(
    pkg: DocxPackage,
    *,
    find: str,
    style: str,
    occurrence: int = 1,
) -> dict:
    """Apply a character style to a text range (style must exist; create with
    define_style)."""
    from . import _runmap
    from .read import list_styles

    char_styles = {
        s["id"] for s in list_styles(pkg) if s["type"] == "character"
    }
    if style not in char_styles:
        raise TargetNotFound(
            f"no character style {style!r}; define_style first "
            f"(existing: {sorted(char_styles)[:15]})"
        )
    import re

    seen = 0
    for p in pkg.root().iter(qn("w:p")):
        text, _ = _runmap.build_map(p)
        for m in re.finditer(re.escape(find), text):
            seen += 1
            if seen == occurrence:
                runs = _runmap.split_for_range(p, m.start(), m.end())
                for run in runs:
                    rpr = run.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = etree.Element(qn("w:rPr"))
                        run.insert(0, rpr)
                    rs = rpr.find(qn("w:rStyle"))
                    if rs is None:
                        rs = etree.Element(qn("w:rStyle"))
                        rpr.insert(0, rs)
                    rs.set(qn("w:val"), style)
                pkg.mark_dirty()
                return {"styled": find, "style": style}
    raise TargetNotFound(f"occurrence {occurrence} of {find!r} not found")


# --------------------------------------------------------------- image alt text

_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def set_image_alt_text(
    pkg: DocxPackage,
    image_index: int,
    *,
    description: str,
    title: str | None = None,
) -> dict:
    """Accessibility alt text for an image (by list_images index)."""
    _A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    blips = list(pkg.root().iter(f"{{{_A}}}blip"))
    if not 0 <= image_index < len(blips):
        raise TargetNotFound(
            f"image index {image_index} out of range ({len(blips)} images)"
        )
    container = blips[image_index].getparent()
    while container is not None and not container.tag.endswith("}inline") and not container.tag.endswith("}anchor"):
        container = container.getparent()
    if container is None:
        raise WordMcpError("image has no drawing container")
    docpr = container.find(f"{{{_WP}}}docPr")
    if docpr is None:
        raise WordMcpError("image has no docPr element")
    docpr.set("descr", description)
    if title:
        docpr.set("title", title)
    pkg.mark_dirty()
    return {"alt_text_set": image_index}
