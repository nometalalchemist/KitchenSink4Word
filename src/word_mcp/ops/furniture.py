"""Headers, footers, page numbers, sections, page layout."""

from __future__ import annotations

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from .read import paragraph_text

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_HDR_FTR = {
    "header": {
        "root": "w:hdr",
        "ref": "w:headerReference",
        "prefix": "header",
        "content_type": (
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.header+xml"
        ),
        "rel_type": (
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/header"
        ),
        "style": "Header",
    },
    "footer": {
        "root": "w:ftr",
        "ref": "w:footerReference",
        "prefix": "footer",
        "content_type": (
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.footer+xml"
        ),
        "rel_type": (
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/footer"
        ),
        "style": "Footer",
    },
}

_REF_TYPES = ("default", "first", "even")

# CT_SectPr child sequence (ECMA-376). Elements must be inserted in this order
# or strict consumers flag the document.
_SECTPR_ORDER = [
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]


def _sectpr_get_or_add(sp: etree._Element, local: str) -> etree._Element:
    """Get the sectPr child `local`, creating it at its schema position."""
    existing = sp.find(qn(f"w:{local}"))
    if existing is not None:
        return existing
    el = etree.Element(qn(f"w:{local}"))
    my_rank = _SECTPR_ORDER.index(local)
    for child in sp:
        name = etree.QName(child).localname
        if name in _SECTPR_ORDER and _SECTPR_ORDER.index(name) > my_rank:
            child.addprevious(el)
            return el
    sp.append(el)
    return el


def _sect_prs(pkg: DocxPackage) -> list[etree._Element]:
    """All sectPr elements in document order (paragraph-level ones first, the
    body-final one last)."""
    body = pkg.body()
    out = [
        sp
        for p in body.findall(qn("w:p"))
        for sp in [p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")]
        if sp is not None
    ]
    final = body.find(qn("w:sectPr"))
    if final is not None:
        out.append(final)
    return out


def list_sections(pkg: DocxPackage) -> list[dict]:
    out = []
    for i, sp in enumerate(_sect_prs(pkg)):
        entry = {"index": i}
        pgsz = sp.find(qn("w:pgSz"))
        if pgsz is not None:
            entry["page_width_pt"] = int(pgsz.get(qn("w:w"), "0")) / 20
            entry["page_height_pt"] = int(pgsz.get(qn("w:h"), "0")) / 20
            entry["orientation"] = pgsz.get(qn("w:orient"), "portrait")
        pgmar = sp.find(qn("w:pgMar"))
        if pgmar is not None:
            entry["margins_pt"] = {
                side: int(pgmar.get(qn(f"w:{side}"), "0")) / 20
                for side in ("top", "bottom", "left", "right")
            }
        entry["headers"] = [
            {"type": r.get(qn("w:type"), "default")}
            for r in sp.findall(qn("w:headerReference"))
        ]
        entry["footers"] = [
            {"type": r.get(qn("w:type"), "default")}
            for r in sp.findall(qn("w:footerReference"))
        ]
        out.append(entry)
    return out


def get_headers_footers(pkg: DocxPackage) -> dict:
    """Text of every header/footer part."""
    import re as _re

    out = {"headers": [], "footers": []}
    for name in pkg.part_names():
        m = _re.fullmatch(r"word/(header|footer)(\d+)\.xml", name)
        if not m:
            continue
        text = "\n".join(
            paragraph_text(p) for p in pkg.root(name).iter(qn("w:p"))
        ).strip()
        has_page_field = any(
            (it.text or "").strip().startswith("PAGE")
            for it in pkg.root(name).iter(qn("w:instrText"))
        ) or any(
            "PAGE" in (fs.get(qn("w:instr")) or "")
            for fs in pkg.root(name).iter(qn("w:fldSimple"))
        )
        out[m.group(1) + "s"].append(
            {"part": name, "text": text, "has_page_number_field": has_page_field}
        )
    return out


def _next_part_number(pkg: DocxPackage, prefix: str) -> int:
    import re as _re

    nums = [
        int(m.group(1))
        for name in pkg.part_names()
        for m in [_re.fullmatch(rf"word/{prefix}(\d+)\.xml", name)]
        if m
    ]
    return max(nums, default=0) + 1


def _create_part(pkg: DocxPackage, kind: str, paragraphs: list[etree._Element]) -> str:
    """Create a header/footer part; returns its relationship id."""
    cfg = _HDR_FTR[kind]
    n = _next_part_number(pkg, cfg["prefix"])
    part = f"word/{cfg['prefix']}{n}.xml"
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = etree.Element(qn(cfg["root"]), nsmap={"w": w_ns, "r": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )})
    for p in paragraphs:
        root.append(p)
    pkg.set_raw_part(
        part,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )

    ct_root = pkg.root("[Content_Types].xml")
    override = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
    override.set("PartName", "/" + part)
    override.set("ContentType", cfg["content_type"])
    pkg.mark_dirty("[Content_Types].xml")

    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    existing = {r.get("Id") for r in rels_root}
    i = 1
    while f"rId{i}" in existing:
        i += 1
    rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", f"rId{i}")
    rel.set("Type", cfg["rel_type"])
    rel.set("Target", part.split("/", 1)[1])
    pkg.mark_dirty(rels_part)
    return f"rId{i}"


def set_columns(
    pkg: DocxPackage,
    *,
    section: int = 0,
    count: int = 1,
    space_pt: float = 36,
    separator: bool = False,
    widths_pt: list[float] | None = None,
) -> dict:
    """Multi-column text layout for a section. Equal widths by default;
    widths_pt gives explicit per-column widths (len must equal count)."""
    if not 1 <= count <= 45:
        raise WordMcpError("count must be 1-45")
    sects = _sect_prs(pkg)
    if not 0 <= section < len(sects):
        raise TargetNotFound(f"section {section} out of range")
    sp = sects[section]
    cols = _sectpr_get_or_add(sp, "cols")
    for child in list(cols):
        cols.remove(child)
    cols.attrib.clear()
    cols.set(qn("w:space"), str(int(space_pt * 20)))
    if count > 1:
        cols.set(qn("w:num"), str(count))
    if separator:
        cols.set(qn("w:sep"), "1")
    if widths_pt is not None:
        if len(widths_pt) != count:
            raise WordMcpError(f"widths_pt must have {count} entries")
        if any(w <= 0 for w in widths_pt):
            raise WordMcpError("column widths must be positive")
        cols.set(qn("w:equalWidth"), "0")
        for i, w in enumerate(widths_pt):
            col = etree.SubElement(cols, qn("w:col"))
            col.set(qn("w:w"), str(int(w * 20)))
            if i < count - 1:
                col.set(qn("w:space"), str(int(space_pt * 20)))
    pkg.mark_dirty()
    return {"section": section, "columns": count, "separator": separator}


def _make_hf_paragraph(
    kind: str, text: str, *, alignment: str = "center", page_field: bool = False,
    x_of_y: bool = False,
) -> etree._Element:
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), _HDR_FTR[kind]["style"])
    jc = etree.SubElement(ppr, qn("w:jc"))
    jc.set(
        qn("w:val"),
        {"left": "left", "center": "center", "right": "right"}[alignment],
    )
    if text:
        r = etree.SubElement(p, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    if page_field:
        if x_of_y and not text:
            lead = etree.SubElement(p, qn("w:r"))
            lt = etree.SubElement(lead, qn("w:t"))
            lt.text = "Page "
            lt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        fs = etree.SubElement(p, qn("w:fldSimple"))
        fs.set(qn("w:instr"), " PAGE \\* MERGEFORMAT ")
        r = etree.SubElement(fs, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = "1"
        if x_of_y:
            mid = etree.SubElement(p, qn("w:r"))
            mt = etree.SubElement(mid, qn("w:t"))
            mt.text = " of "
            mt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            fs2 = etree.SubElement(p, qn("w:fldSimple"))
            fs2.set(qn("w:instr"), " NUMPAGES \\* MERGEFORMAT ")
            r2 = etree.SubElement(fs2, qn("w:r"))
            t2 = etree.SubElement(r2, qn("w:t"))
            t2.text = "1"
    return p


def set_header_footer(
    pkg: DocxPackage,
    kind: str,
    text: str,
    *,
    section: int = 0,
    ref_type: str = "default",
    alignment: str = "center",
    include_page_number: bool = False,
    x_of_y: bool = False,
) -> dict:
    """Set a section's header or footer text. An existing referenced part is
    rewritten; otherwise a new part is created and referenced. ref_type
    'first'/'even' additionally flip the matching section/settings switches."""
    if kind not in _HDR_FTR:
        raise WordMcpError("kind must be 'header' or 'footer'")
    if ref_type not in _REF_TYPES:
        raise WordMcpError(f"ref_type must be one of {_REF_TYPES}")
    sects = _sect_prs(pkg)
    if not 0 <= section < len(sects):
        raise TargetNotFound(
            f"section {section} out of range (document has {len(sects)})"
        )
    sp = sects[section]
    cfg = _HDR_FTR[kind]

    existing_ref = next(
        (
            r
            for r in sp.findall(qn(cfg["ref"]))
            if r.get(qn("w:type"), "default") == ref_type
        ),
        None,
    )
    paragraphs = [
        _make_hf_paragraph(
            kind, text, alignment=alignment, page_field=include_page_number,
            x_of_y=x_of_y,
        )
    ]
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    if existing_ref is not None:
        rid = existing_ref.get(f"{{{r_ns}}}id")
        rels_root = pkg.root("word/_rels/document.xml.rels")
        target = next(
            (r.get("Target") for r in rels_root if r.get("Id") == rid), None
        )
        if target:
            part = "word/" + target.lstrip("/")
            root = pkg.root(part)
            # Watermark paragraphs must survive a header/footer rewrite
            # (bug found in Phase D: setting header text destroyed watermarks).
            preserved = [
                child
                for child in list(root)
                if any(
                    (s.get("id") or "").startswith("PowerPlusWaterMarkObject")
                    for s in child.iter(
                        "{urn:schemas-microsoft-com:vml}shape"
                    )
                )
            ]
            for child in list(root):
                root.remove(child)
            for wm in preserved:
                root.append(wm)
            for p in paragraphs:
                root.append(p)
            pkg.mark_dirty(part)
            _flip_type_switches(pkg, sp, ref_type)
            return {"set": kind, "section": section, "type": ref_type, "part": part}

    rid = _create_part(pkg, kind, paragraphs)
    ref = etree.SubElement(sp, qn(cfg["ref"]))
    ref.set(qn("w:type"), ref_type)
    ref.set(f"{{{r_ns}}}id", rid)
    # References must precede other sectPr children per schema; move to front.
    sp.remove(ref)
    sp.insert(0, ref)
    pkg.mark_dirty()
    _flip_type_switches(pkg, sp, ref_type)
    return {"set": kind, "section": section, "type": ref_type, "created": True}


def _flip_type_switches(pkg: DocxPackage, sp: etree._Element, ref_type: str) -> None:
    if ref_type == "first":
        if sp.find(qn("w:titlePg")) is None:
            # titlePg position: after headerReference/footerReference/footnotePr/endnotePr/type
            sp.insert(_after_refs_index(sp), etree.Element(qn("w:titlePg")))
            pkg.mark_dirty()
    elif ref_type == "even":
        if pkg.has_part("word/settings.xml"):
            root = pkg.root("word/settings.xml")
            if root.find(qn("w:evenAndOddHeaders")) is None:
                root.insert(0, etree.Element(qn("w:evenAndOddHeaders")))
                pkg.mark_dirty("word/settings.xml")


def _after_refs_index(sp: etree._Element) -> int:
    i = 0
    for child in sp:
        if etree.QName(child).localname in (
            "headerReference",
            "footerReference",
            "footnotePr",
            "endnotePr",
            "type",
            "pgSz",
            "pgMar",
        ):
            i += 1
        else:
            break
    return i


def add_page_numbers(
    pkg: DocxPackage,
    *,
    section: int = 0,
    position: str = "footer",
    alignment: str = "center",
    prefix: str = "",
    start_at: int | None = None,
    x_of_y: bool = False,
) -> dict:
    """Page numbers via a PAGE field in the header or footer. x_of_y renders
    'Page N of M' (NUMPAGES)."""
    result = set_header_footer(
        pkg,
        "footer" if position == "footer" else "header",
        prefix,
        section=section,
        alignment=alignment,
        include_page_number=True,
        x_of_y=x_of_y,
    )
    if start_at is not None:
        sp = _sect_prs(pkg)[section]
        pgnum = _sectpr_get_or_add(sp, "pgNumType")
        pgnum.set(qn("w:start"), str(start_at))
        pkg.mark_dirty()
    result["page_numbers"] = True
    return result


_PAGE_NUM_FORMATS = {
    "decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter",
    "ordinal", "cardinalText", "ordinalText", "decimalZero",
}


def set_page_number_format(
    pkg: DocxPackage,
    *,
    section: int = 0,
    number_format: str = "decimal",
    start_at: int | None = None,
) -> dict:
    """Page-number FORMAT for one section: lowerRoman for dissertation front
    matter, decimal restarting at 1 for the body, etc. Applies to PAGE fields
    in that section's headers/footers."""
    if number_format not in _PAGE_NUM_FORMATS:
        raise WordMcpError(
            f"number_format must be one of {sorted(_PAGE_NUM_FORMATS)}"
        )
    sects = _sect_prs(pkg)
    if not 0 <= section < len(sects):
        raise TargetNotFound(
            f"section {section} out of range (document has {len(sects)})"
        )
    pgnum = _sectpr_get_or_add(sects[section], "pgNumType")
    pgnum.set(qn("w:fmt"), number_format)
    if start_at is not None:
        pgnum.set(qn("w:start"), str(start_at))
    pkg.mark_dirty()
    return {
        "section": section,
        "number_format": number_format,
        "start_at": start_at,
    }


def set_line_numbering(
    pkg: DocxPackage,
    *,
    section: int = 0,
    count_by: int = 1,
    start: int = 1,
    restart: str = "continuous",
    distance_pt: float | None = None,
    remove: bool = False,
) -> dict:
    """Manuscript line numbering for a section (journals often require it).
    restart: continuous | newPage | newSection."""
    sects = _sect_prs(pkg)
    if not 0 <= section < len(sects):
        raise TargetNotFound(f"section {section} out of range")
    sp = sects[section]
    if remove:
        el = sp.find(qn("w:lnNumType"))
        if el is not None:
            sp.remove(el)
            pkg.mark_dirty()
        return {"section": section, "line_numbering": "removed"}
    if restart not in ("continuous", "newPage", "newSection"):
        raise WordMcpError("restart must be continuous | newPage | newSection")
    ln = _sectpr_get_or_add(sp, "lnNumType")
    ln.set(qn("w:countBy"), str(count_by))
    ln.set(qn("w:start"), str(start))
    ln.set(qn("w:restart"), restart)
    if distance_pt is not None:
        ln.set(qn("w:distance"), str(int(distance_pt * 20)))
    pkg.mark_dirty()
    return {"section": section, "line_numbering": restart, "count_by": count_by}


# Word's text-effect preset shapetype (t136) — required for robustness; the
# PowerPlusWaterMarkObject id prefix keeps Word's own Remove Watermark working.
_WATERMARK_SHAPETYPE = (
    '<v:shapetype xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" '
    'path="m@7,l@8,m@5,21600l@6,21600e">'
    '<v:formulas>'
    '<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/>'
    '<v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
    '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>'
    '<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/>'
    '<v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/>'
    '<v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>'
    '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>'
    "</v:formulas>"
    '<v:path textpathok="t" o:connecttype="custom" '
    'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" '
    'o:connectangles="270,180,90,0"/>'
    '<v:textpath on="t" fitshape="t"/>'
    '<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
    '<o:lock v:ext="edit" text="t" shapetype="t"/>'
    "</v:shapetype>"
)


def _watermark_paragraph(
    text: str, *, color: str, opacity: float, diagonal: bool, font: str,
    shape_num: int,
) -> etree._Element:
    rotation = "rotation:315;" if diagonal else ""
    shape = (
        '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        f'id="PowerPlusWaterMarkObject{357831000 + shape_num}" '
        f'o:spid="_x0000_s{2048 + shape_num}" type="#_x0000_t136" '
        f'style="position:absolute;margin-left:0;margin-top:0;'
        f"width:415.2pt;height:207.6pt;{rotation}z-index:-251654144;"
        "mso-position-horizontal:center;"
        "mso-position-horizontal-relative:margin;"
        "mso-position-vertical:center;"
        'mso-position-vertical-relative:margin" '
        f'o:allowincell="f" fillcolor="{color}" stroked="f">'
        f'<v:fill opacity="{opacity}"/>'
        f'<v:textpath style="font-family:&quot;{font}&quot;;font-size:1pt" '
        f'string="{text}"/>'
        '<w10:wrap anchorx="margin" anchory="margin"/>'
        "</v:shape>"
    )
    p = etree.Element(qn("w:p"))
    r = etree.SubElement(p, qn("w:r"))
    rpr = etree.SubElement(r, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:noProof"))
    pict = etree.SubElement(r, qn("w:pict"))
    pict.append(etree.fromstring(_WATERMARK_SHAPETYPE))
    pict.append(etree.fromstring(shape))
    return p


def add_watermark(
    pkg: DocxPackage,
    text: str = "DRAFT",
    *,
    color: str = "silver",
    opacity: float = 0.5,
    diagonal: bool = True,
    font: str = "Calibri",
) -> dict:
    """Text watermark (DRAFT / CONFIDENTIAL / ...) behind the text on every
    page: the shape is placed in EVERY header part the document references,
    and a default header is created for sections that have none."""
    if not text.strip():
        raise WordMcpError("watermark text must be non-empty")
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rels_root = pkg.root("word/_rels/document.xml.rels")
    rid_target = {r.get("Id"): r.get("Target") for r in rels_root}

    header_parts: set[str] = set()
    shape_num = 1
    for sp in _sect_prs(pkg):
        refs = sp.findall(qn("w:headerReference"))
        if not refs:
            wm = _watermark_paragraph(
                text, color=color, opacity=opacity, diagonal=diagonal,
                font=font, shape_num=shape_num,
            )
            shape_num += 1
            rid = _create_part(pkg, "header", [wm])
            ref = etree.SubElement(sp, qn("w:headerReference"))
            ref.set(qn("w:type"), "default")
            ref.set(f"{{{r_ns}}}id", rid)
            sp.remove(ref)
            sp.insert(0, ref)
            pkg.mark_dirty()
            rels_root = pkg.root("word/_rels/document.xml.rels")
            rid_target = {r.get("Id"): r.get("Target") for r in rels_root}
            target = rid_target.get(rid)
            if target:
                header_parts.add("word/" + target.lstrip("/"))
            continue
        for ref in refs:
            rid = ref.get(f"{{{r_ns}}}id")
            target = rid_target.get(rid)
            if target:
                header_parts.add("word/" + target.lstrip("/"))

    added = 0
    for part in sorted(header_parts):
        root = pkg.root(part)
        already = any(
            (el.get("id") or "").startswith("PowerPlusWaterMarkObject")
            for el in root.iter("{urn:schemas-microsoft-com:vml}shape")
        )
        if already:
            continue
        wm = _watermark_paragraph(
            text, color=color, opacity=opacity, diagonal=diagonal, font=font,
            shape_num=shape_num,
        )
        shape_num += 1
        root.insert(0, wm)
        pkg.mark_dirty(part)
        added += 1
    return {"watermark": text, "header_parts": added or len(header_parts)}


def remove_watermark(pkg: DocxPackage) -> dict:
    """Remove PowerPlusWaterMarkObject shapes from every header part."""
    import re as _re

    removed = 0
    for name in pkg.part_names():
        if not _re.fullmatch(r"word/header\d+\.xml", name):
            continue
        root = pkg.root(name)
        for shape in list(
            root.iter("{urn:schemas-microsoft-com:vml}shape")
        ):
            if (shape.get("id") or "").startswith("PowerPlusWaterMarkObject"):
                pict = shape.getparent()
                run = pict.getparent()
                para = run.getparent()
                para.remove(run)
                if not para.findall(qn("w:r")):
                    parent = para.getparent()
                    if parent is not None and len(parent.findall(qn("w:p"))) > 1:
                        parent.remove(para)
                pkg.mark_dirty(name)
                removed += 1
    if not removed:
        raise TargetNotFound("no watermark found in any header")
    return {"watermarks_removed": removed}


def _section_state(sp: etree._Element) -> dict:
    """Everything set_section_properties can set, as currently stored:
    page size, orientation, and all margin sides. Keys are omitted when the
    underlying element is absent (Word then applies its own defaults)."""
    state: dict = {}
    pgsz = sp.find(qn("w:pgSz"))
    if pgsz is not None:
        state["page_width_pt"] = int(pgsz.get(qn("w:w"), "0")) / 20
        state["page_height_pt"] = int(pgsz.get(qn("w:h"), "0")) / 20
        state["orientation"] = pgsz.get(qn("w:orient"), "portrait")
    pgmar = sp.find(qn("w:pgMar"))
    if pgmar is not None:
        state["margins_pt"] = {
            side: int(pgmar.get(qn(f"w:{side}"), "0")) / 20
            for side in (
                "top", "bottom", "left", "right", "header", "footer", "gutter",
            )
        }
    return state


def set_section_properties(
    pkg: DocxPackage,
    *,
    section: int = 0,
    orientation: str | None = None,
    page_width_pt: float | None = None,
    page_height_pt: float | None = None,
    margins_pt: dict | None = None,
) -> dict:
    sects = _sect_prs(pkg)
    if not 0 <= section < len(sects):
        raise TargetNotFound(
            f"section {section} out of range "
            f"({len(sects)} section(s), valid indices 0-{len(sects) - 1})"
        )
    sp = sects[section]
    changed = []
    if orientation or page_width_pt or page_height_pt:
        pgsz = sp.find(qn("w:pgSz"))
        if pgsz is None:
            pgsz = etree.SubElement(sp, qn("w:pgSz"))
            pgsz.set(qn("w:w"), "12240")
            pgsz.set(qn("w:h"), "15840")
        if orientation:
            if orientation not in ("portrait", "landscape"):
                raise WordMcpError("orientation must be portrait or landscape")
            cur_w = int(pgsz.get(qn("w:w")))
            cur_h = int(pgsz.get(qn("w:h")))
            needs_swap = (orientation == "landscape") != (cur_w > cur_h)
            if needs_swap:
                pgsz.set(qn("w:w"), str(cur_h))
                pgsz.set(qn("w:h"), str(cur_w))
            if orientation == "landscape":
                pgsz.set(qn("w:orient"), "landscape")
            else:
                pgsz.attrib.pop(qn("w:orient"), None)
            changed.append("orientation")
        if page_width_pt:
            pgsz.set(qn("w:w"), str(int(page_width_pt * 20)))
            changed.append("page_width")
        if page_height_pt:
            pgsz.set(qn("w:h"), str(int(page_height_pt * 20)))
            changed.append("page_height")
    if margins_pt:
        pgmar = sp.find(qn("w:pgMar"))
        if pgmar is None:
            pgmar = etree.SubElement(sp, qn("w:pgMar"))
            for side, default in (
                ("top", "1440"), ("right", "1440"), ("bottom", "1440"),
                ("left", "1440"), ("header", "720"), ("footer", "720"),
                ("gutter", "0"),
            ):
                pgmar.set(qn(f"w:{side}"), default)
        for side, val in margins_pt.items():
            if side not in ("top", "bottom", "left", "right", "header", "footer"):
                raise WordMcpError(f"unknown margin side: {side}")
            pgmar.set(qn(f"w:{side}"), str(int(val * 20)))
        changed.append("margins")
    if changed:
        pkg.mark_dirty()
    # Full state in every response: current state on a no-change call (safe
    # way to READ a section), post-change state otherwise.
    return {"section": section, "changed": changed, "state": _section_state(sp)}


def add_section_break(pkg: DocxPackage, *, after_index: int, break_type: str = "nextPage") -> dict:
    """Insert a section break after a body paragraph. The new section inherits
    the following section's properties (Word semantics: a paragraph-level
    sectPr closes the section that precedes it)."""
    import copy as _copy

    if break_type not in ("nextPage", "continuous", "evenPage", "oddPage"):
        raise WordMcpError("break_type must be nextPage|continuous|evenPage|oddPage")
    from .text import _body_paragraph

    p = _body_paragraph(pkg, after_index)
    # Find the sectPr governing this point (next one in document order).
    governing = None
    seen = False
    body = pkg.body()
    for child in body:
        if child is p:
            seen = True
        if seen:
            spr = (
                child.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
                if etree.QName(child).localname == "p"
                else None
            )
            if spr is not None:
                governing = spr
                break
    if governing is None:
        governing = body.find(qn("w:sectPr"))
    if governing is None:
        raise WordMcpError("document has no section properties to inherit")

    new_p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(new_p, qn("w:pPr"))
    new_sp = _copy.deepcopy(governing)
    type_el = new_sp.find(qn("w:type"))
    if type_el is None:
        type_el = etree.Element(qn("w:type"))
        new_sp.insert(_after_refs_index(new_sp), type_el)
    type_el.set(qn("w:val"), break_type)
    ppr.append(new_sp)
    p.addnext(new_p)
    pkg.mark_dirty()
    return {"section_break_after": after_index, "type": break_type}
