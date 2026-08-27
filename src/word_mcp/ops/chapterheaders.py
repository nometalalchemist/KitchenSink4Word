"""Chapter-aware running headers via STYLEREF fields.

A STYLEREF field in a header shows the text of the nearest preceding
paragraph in the referenced Heading style. Word evaluates it per page, so
one header serves every chapter — no per-chapter section breaks needed, and
it stays correct when chapters are renamed or reordered.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from .fields import _field_run
from .furniture import _after_refs_index, _create_part, _sect_prs
from .read import _outline_level, _style_outline_map, paragraph_text

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_VML_SHAPE = "{urn:schemas-microsoft-com:vml}shape"
_ALIGNMENTS = ("left", "center", "right")


def _heading_style_name(pkg: DocxPackage, level: int) -> str:
    """The style NAME (not id) STYLEREF must reference. Prefer the document's
    own name for Heading{level}; fall back to Word's built-in name."""
    if pkg.has_part("word/styles.xml"):
        for style in pkg.root("word/styles.xml").findall(qn("w:style")):
            if style.get(qn("w:styleId")) == f"Heading{level}":
                name_el = style.find(qn("w:name"))
                if name_el is not None and name_el.get(qn("w:val")):
                    return name_el.get(qn("w:val"))
    return f"heading {level}"


def _section_heading_levels(pkg: DocxPackage) -> list[set[int]]:
    """Heading levels present in each section, by section index. Body-level
    paragraphs only (headings inside tables are not scanned)."""
    style_map = _style_outline_map(pkg)
    n_sections = len(_sect_prs(pkg))
    out: list[set[int]] = [set() for _ in range(n_sections)]
    sec = 0
    for child in pkg.body():
        if etree.QName(child).localname != "p":
            continue
        lvl = _outline_level(child, style_map)
        if lvl is not None and paragraph_text(child).strip() and sec < n_sections:
            out[sec].add(lvl)
        if child.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            sec += 1
    return out


def _section_heading_texts(pkg: DocxPackage, section: int, level: int) -> list:
    """Texts of level-N headings inside one section (body-level only)."""
    style_map = _style_outline_map(pkg)
    out: list[str] = []
    sec = 0
    for child in pkg.body():
        if etree.QName(child).localname != "p":
            continue
        if sec == section:
            lvl = _outline_level(child, style_map)
            text = paragraph_text(child).strip()
            if lvl == level and text:
                out.append(text)
        if child.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            sec += 1
            if sec > section:
                break
    return out


def _default_header_ref(sp: etree._Element) -> etree._Element | None:
    for r in sp.findall(qn("w:headerReference")):
        if r.get(qn("w:type"), "default") == "default":
            return r
    return None


def _rid_to_part(pkg: DocxPackage, rid: str | None) -> str | None:
    if rid is None:
        return None
    rels_root = pkg.root("word/_rels/document.xml.rels")
    target = next((r.get("Target") for r in rels_root if r.get("Id") == rid), None)
    return "word/" + target.lstrip("/") if target else None


def _effective_default_header(
    pkg: DocxPackage, sects: list[etree._Element], index: int
) -> tuple[int, str] | None:
    """(owning section index, part name) of the default header Word shows for
    section `index` — sections without their own reference inherit backward."""
    for j in range(index, -1, -1):
        ref = _default_header_ref(sects[j])
        if ref is not None:
            part = _rid_to_part(pkg, ref.get(f"{{{_R_NS}}}id"))
            if part is not None and pkg.has_part(part):
                return j, part
            return None  # reference exists but is broken; treat as no header
    return None


def _chapter_header_paragraph(
    style_name: str, *, alignment: str, include_number: bool
) -> tuple[etree._Element, list[str]]:
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "Header")
    etree.SubElement(ppr, qn("w:jc")).set(qn("w:val"), alignment)
    codes = []
    if include_number:
        num_instr = f'STYLEREF "{style_name}" \\n'
        codes.append(num_instr)
        for el in _field_run(num_instr):
            p.append(el)
        r = etree.SubElement(p, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = "  "
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_instr = f'STYLEREF "{style_name}"'
    codes.append(text_instr)
    for el in _field_run(text_instr):
        p.append(el)
    return p, codes


def setup_chapter_headers(
    pkg: DocxPackage,
    *,
    level: int = 1,
    include_number: bool = False,
    alignment: str = "right",
    first_page_blank: bool = True,
    scope: str | list[int] = "auto",
) -> dict:
    """Put the current chapter title (Heading style of `level`) in the primary
    header via a STYLEREF field, optionally preceded by the heading number
    (STYLEREF \\n). Word evaluates header fields per page, so no field-update
    prompt is needed.

    scope 'auto' targets every section that contains body headings of that
    level; an explicit list of section indices overrides. first_page_blank
    sets titlePg so chapter-opening pages (section starts) show no header.
    Existing watermarks in rewritten headers are preserved; other existing
    header content is replaced (reported)."""
    if not 1 <= level <= 9:
        raise WordMcpError("level must be 1-9")
    if alignment not in _ALIGNMENTS:
        raise WordMcpError(f"alignment must be one of {_ALIGNMENTS}")
    sects = _sect_prs(pkg)
    levels_by_section = _section_heading_levels(pkg)

    if scope == "auto":
        targets = sorted(
            i for i, lvls in enumerate(levels_by_section) if level in lvls
        )
        # front-matter guard: when the doc has 2+ sections and every
        # heading in section 0 is a standard front-matter title, a running
        # header there would read "Abstract" — skip it (explicit scope
        # lists still target it deliberately)
        if len(levels_by_section) > 1 and 0 in targets:
            fm_titles = {
                "abstract", "acknowledgments", "acknowledgements",
                "dedication", "preface", "table of contents",
                "list of figures", "list of tables", "list of equations",
                "copyright", "declaration",
            }
            sec0_headings = _section_heading_texts(pkg, 0, level)
            if sec0_headings and all(
                h.strip().lower() in fm_titles for h in sec0_headings
            ):
                targets = [t for t in targets if t != 0]
        if not targets:
            raise TargetNotFound(
                f"no body headings of level {level}; nothing for the header "
                "to reference (add headings first, or pass explicit section "
                "indices as scope)"
            )
    elif isinstance(scope, list):
        targets = sorted(set(scope))
        for i in targets:
            if not isinstance(i, int) or not 0 <= i < len(sects):
                raise TargetNotFound(
                    f"scope section {i} out of range (document has {len(sects)})"
                )
    else:
        raise WordMcpError("scope must be 'auto' or a list of section indices")

    style_name = _heading_style_name(pkg, level)
    written_parts: dict[str, dict] = {}
    created_parts: list[str] = []
    replaced_content: list[str] = []
    sections_report: list[dict] = []
    field_codes: list[str] = []

    for i in targets:
        sp = sects[i]
        effective = _effective_default_header(pkg, sects, i)
        entry: dict = {"section": i}
        if effective is not None:
            owner, part = effective
            entry["header_part"] = part
            entry["inherited_from_section"] = None if owner == i else owner
            if part not in written_parts:
                root = pkg.root(part)
                # Watermark paragraphs must survive the rewrite (the v1.2
                # Phase D bug: rewriting a header destroyed watermarks).
                preserved = [
                    child
                    for child in list(root)
                    if any(
                        (s.get("id") or "").startswith("PowerPlusWaterMarkObject")
                        for s in child.iter(_VML_SHAPE)
                    )
                ]
                # Keep descendant proxies alive so their ids stay stable.
                preserved_desc = [q for wm in preserved for q in wm.iter()]
                preserved_ids = {id(q) for q in preserved_desc}
                had_other_content = any(
                    paragraph_text(p).strip()
                    for p in root.iter(qn("w:p"))
                    if id(p) not in preserved_ids
                )
                for child in list(root):
                    root.remove(child)
                for wm in preserved:
                    root.append(wm)
                para, codes = _chapter_header_paragraph(
                    style_name, alignment=alignment, include_number=include_number
                )
                field_codes = codes
                root.append(para)
                pkg.mark_dirty(part)
                written_parts[part] = {"watermark_preserved": bool(preserved)}
                if had_other_content:
                    replaced_content.append(part)
        else:
            para, codes = _chapter_header_paragraph(
                style_name, alignment=alignment, include_number=include_number
            )
            field_codes = codes
            rid = _create_part(pkg, "header", [para])
            ref = etree.SubElement(sp, qn("w:headerReference"))
            ref.set(qn("w:type"), "default")
            ref.set(f"{{{_R_NS}}}id", rid)
            # References must precede other sectPr children per schema.
            sp.remove(ref)
            sp.insert(0, ref)
            pkg.mark_dirty()
            part = _rid_to_part(pkg, rid)
            entry["header_part"] = part
            created_parts.append(part)
            written_parts[part] = {"watermark_preserved": False}
        if first_page_blank and sp.find(qn("w:titlePg")) is None:
            sp.insert(_after_refs_index(sp), etree.Element(qn("w:titlePg")))
            pkg.mark_dirty()
        sections_report.append(entry)

    # Sections OUTSIDE the scope whose effective header we rewrote now show
    # the chapter header too; report it rather than pretend otherwise.
    shared_out_of_scope = sorted(
        i
        for i in range(len(sects))
        if i not in targets
        for eff in [_effective_default_header(pkg, sects, i)]
        if eff is not None and eff[1] in written_parts
    )

    result = {
        "level": level,
        "style_referenced": style_name,
        "field_codes": field_codes,
        "sections": sections_report,
        "parts_written": sorted(written_parts),
        "parts_created": created_parts,
        "watermark_preserved_in": sorted(
            p for p, meta in written_parts.items() if meta["watermark_preserved"]
        ),
        "first_page_blank": first_page_blank,
    }
    if replaced_content:
        result["replaced_existing_header_content_in"] = replaced_content
    if shared_out_of_scope:
        result["also_affects_sections"] = shared_out_of_scope
    return result


def validate_chapter_headers(pkg: DocxPackage) -> dict:
    """Read back the chapter-header state: per section, the effective default
    header part, the STYLEREF field codes it carries, and the heading levels
    the section contains. A section is flagged missing when it has headings
    at the document's top heading level but its effective header carries no
    STYLEREF field."""
    sects = _sect_prs(pkg)
    levels_by_section = _section_heading_levels(pkg)
    all_levels = sorted({l for s in levels_by_section for l in s})
    chapter_level = all_levels[0] if all_levels else None

    sections_report = []
    missing = []
    for i, sp in enumerate(sects):
        effective = _effective_default_header(pkg, sects, i)
        stylerefs: list[str] = []
        part = None
        if effective is not None:
            _, part = effective
            for it in pkg.root(part).iter(qn("w:instrText")):
                text = (it.text or "").strip()
                if text.startswith("STYLEREF"):
                    stylerefs.append(text)
            for fs in pkg.root(part).iter(qn("w:fldSimple")):
                instr = (fs.get(qn("w:instr")) or "").strip()
                if instr.startswith("STYLEREF"):
                    stylerefs.append(instr)
        entry = {
            "section": i,
            "header_part": part,
            "styleref_fields": stylerefs,
            "heading_levels": sorted(levels_by_section[i]),
            "has_first_page_blank": sp.find(qn("w:titlePg")) is not None,
        }
        sections_report.append(entry)
        if (
            chapter_level is not None
            and chapter_level in levels_by_section[i]
            and not stylerefs
        ):
            missing.append(i)
    return {
        "chapter_level": chapter_level,
        "sections": sections_report,
        "sections_missing_chapter_header": missing,
        "ok": not missing,
    }
