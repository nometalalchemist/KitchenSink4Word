"""Word-native bibliography: sources, CITATION fields, BIBLIOGRAPHY field.

Ground truth (research/20260827_v12, topic 1, verified against Word 365):
- Sources live in customXml/itemN.xml (b:Sources) + itemPropsN.xml declaring
  the bibliography schemaRef; wired by a customXml relationship. NOT a
  dedicated part.
- The citation style is stored ON b:Sources (@SelectedStyle/@StyleName),
  nowhere in settings.xml.
- CITATION fields are sdt-wrapped complex fields; BIBLIOGRAPHY is a plain
  field whose result paragraphs use the Bibliography style.
- Fields render on update: com_refresh_fields for immediate, or the
  update-on-open flag.
"""

from __future__ import annotations

import re
import uuid

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap

_B = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
_DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CUSTOMXML_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
)
_PROPS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
)

# Verified name -> (StyleName attr, SelectedStyle xsl) map (COM-tested).
STYLES = {
    "APA": ("APA", "\\APASixthEditionOfficeOnline.xsl"),
    "Chicago": ("Chicago", "\\CHICAGO.XSL"),
    "MLA": ("MLA", "\\MLASeventhEditionOfficeOnline.xsl"),
    "IEEE": ("IEEE", "\\IEEE2006OfficeOnline.xsl"),
    "Turabian": ("Turabian", "\\TURABIAN.XSL"),
    "Harvard - Anglia": ("Harvard - Anglia", "\\HarvardAnglia2008OfficeOnline.xsl"),
    "GB7714": ("GB7714", "\\GB.XSL"),
    "GOST - Name Sort": ("GOST - Name Sort", "\\GostName.XSL"),
    "GOST - Title Sort": ("GOST - Title Sort", "\\GostTitle.XSL"),
    "ISO 690 - First Element and Date": (
        "ISO 690 - First Element and Date", "\\ISO690.XSL",
    ),
    "ISO 690 - Numerical Reference": (
        "ISO 690 - Numerical Reference", "\\ISO690Nmerical.XSL",
    ),
    "SIST02": ("SIST02", "\\SIST02.XSL"),
}

SOURCE_TYPES = {
    "ArticleInAPeriodical", "Book", "BookSection", "JournalArticle",
    "ConferenceProceedings", "Report", "SoundRecording", "Performance",
    "Art", "DocumentFromInternetSite", "InternetSite", "Film", "Interview",
    "Patent", "ElectronicSource", "Case", "Misc",
}

# Canonical element order Word writes (cleanliness only; schema is a choice).
_FIELD_ORDER = [
    "Title", "JournalName", "BookTitle", "InternetSiteTitle",
    "PublicationTitle", "ConferenceName", "Year", "Month", "Day", "City",
    "StateProvince", "CountryRegion", "Publisher", "Institution",
    "Department", "ThesisType", "Edition", "Pages", "Volume", "Issue",
    "NumberVolumes", "StandardNumber", "URL", "YearAccessed",
    "MonthAccessed", "DayAccessed", "Comments", "ShortTitle",
]


def _bq(local: str) -> str:
    return f"{{{_B}}}{local}"


# ------------------------------------------------------------- store plumbing


def _find_store(pkg: DocxPackage) -> str | None:
    """Part name of the bibliography customXml item, or None."""
    rels_part = "word/_rels/document.xml.rels"
    if not pkg.has_part(rels_part):
        return None
    for rel in pkg.root(rels_part):
        if rel.get("Type") != _CUSTOMXML_REL:
            continue
        target = rel.get("Target", "")
        item_part = target.replace("../", "")
        item_rels = (
            f"customXml/_rels/{item_part.rsplit('/', 1)[-1]}.rels"
        )
        if not pkg.has_part(item_rels):
            continue
        for r2 in pkg.root(item_rels):
            if r2.get("Type") == _PROPS_REL:
                props_part = "customXml/" + r2.get("Target")
                if pkg.has_part(props_part):
                    props = pkg.root(props_part)
                    for ref in props.iter(f"{{{_DS}}}schemaRef"):
                        if ref.get(f"{{{_DS}}}uri") == _B:
                            return item_part
    return None


def _ensure_store(pkg: DocxPackage, *, style: str = "APA") -> str:
    existing = _find_store(pkg)
    if existing:
        # Templates in the wild (python-docx's default among them) carry
        # legacy style paths like "/APA.XSL" that modern Word cannot resolve —
        # citations then render in Word's fallback style instead of the named
        # one. Normalize to the modern XSL for the store's declared style.
        root = pkg.root(existing)
        selected = root.get("SelectedStyle") or ""
        style_name = root.get("StyleName") or ""
        modern_files = {xsl for _, xsl in STYLES.values()}
        if selected not in modern_files:
            match = STYLES.get(style_name) or STYLES.get(style)
            if match:
                root.set("SelectedStyle", match[1])
                root.set("StyleName", match[0])
                pkg.mark_dirty(existing)
        return existing

    n = 1
    while pkg.has_part(f"customXml/item{n}.xml"):
        n += 1
    item_part = f"customXml/item{n}.xml"
    props_part = f"customXml/itemProps{n}.xml"
    item_rels_part = f"customXml/_rels/item{n}.xml.rels"

    style_name, style_xsl = STYLES[style]
    sources = etree.Element(
        _bq("Sources"), nsmap={"b": _B, None: _B}
    )
    sources.set("SelectedStyle", style_xsl)
    sources.set("StyleName", style_name)
    pkg.set_raw_part(
        item_part,
        etree.tostring(
            sources, xml_declaration=True, encoding="UTF-8", standalone=False
        ),
    )

    props = etree.Element(f"{{{_DS}}}datastoreItem", nsmap={"ds": _DS})
    props.set(f"{{{_DS}}}itemID", "{" + str(uuid.uuid4()).upper() + "}")
    refs = etree.SubElement(props, f"{{{_DS}}}schemaRefs")
    ref = etree.SubElement(refs, f"{{{_DS}}}schemaRef")
    ref.set(f"{{{_DS}}}uri", _B)
    pkg.set_raw_part(
        props_part,
        etree.tostring(
            props, xml_declaration=True, encoding="UTF-8", standalone=False
        ),
    )

    rels_root = etree.Element(
        f"{{{_REL_NS}}}Relationships", nsmap={None: _REL_NS}
    )
    rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", "rId1")
    rel.set("Type", _PROPS_REL)
    rel.set("Target", f"itemProps{n}.xml")
    pkg.set_raw_part(
        item_rels_part,
        etree.tostring(
            rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
    )

    # [Content_Types].xml: Override for props ONLY (item covered by xml Default).
    ct_root = pkg.root("[Content_Types].xml")
    if not any(
        d.get("Extension") == "xml"
        for d in ct_root.findall(f"{{{_CT_NS}}}Default")
    ):
        d = etree.SubElement(ct_root, f"{{{_CT_NS}}}Default")
        d.set("Extension", "xml")
        d.set("ContentType", "application/xml")
    override = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
    override.set("PartName", "/" + props_part)
    override.set(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
    )
    pkg.mark_dirty("[Content_Types].xml")

    doc_rels = pkg.root("word/_rels/document.xml.rels")
    existing_ids = {r.get("Id") for r in doc_rels}
    i = 1
    while f"rId{i}" in existing_ids:
        i += 1
    rel = etree.SubElement(doc_rels, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", f"rId{i}")
    rel.set("Type", _CUSTOMXML_REL)
    rel.set("Target", f"../{item_part}")
    pkg.mark_dirty("word/_rels/document.xml.rels")
    return item_part


# ------------------------------------------------------------------- sources


def add_source(
    pkg: DocxPackage,
    *,
    tag: str,
    source_type: str,
    title: str,
    year: str | None = None,
    authors: list | None = None,
    editors: list | None = None,
    journal_name: str | None = None,
    book_title: str | None = None,
    publisher: str | None = None,
    city: str | None = None,
    pages: str | None = None,
    volume: str | None = None,
    issue: str | None = None,
    edition: str | None = None,
    institution: str | None = None,
    url: str | None = None,
    internet_site_title: str | None = None,
    style: str = "APA",
    extra_fields: dict | None = None,
) -> dict:
    """Add a bibliography source. tag: unique citation key (letters/digits).
    authors/editors: [{last, first, middle?}] or [{corporate}]. Cite it with
    insert_citation(tag=...)."""
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_\-]{0,39}", tag):
        raise WordMcpError(
            "tag must start with a letter; letters/digits/_/- only, max 40"
        )
    if source_type not in SOURCE_TYPES:
        raise WordMcpError(
            f"source_type must be one of {sorted(SOURCE_TYPES)}"
        )
    part = _ensure_store(pkg, style=style)
    root = pkg.root(part)
    if any(
        (s.findtext(_bq("Tag")) or "") == tag
        for s in root.findall(_bq("Source"))
    ):
        raise WordMcpError(f"source tag {tag!r} already exists")

    src = etree.SubElement(root, _bq("Source"))
    etree.SubElement(src, _bq("Tag")).text = tag
    etree.SubElement(src, _bq("SourceType")).text = source_type

    def add_role(role: str, people: list) -> None:
        outer = src.find(_bq("Author"))
        if outer is None:
            outer = etree.SubElement(src, _bq("Author"))
        role_el = etree.SubElement(outer, _bq(role))
        corporate = [p for p in people if isinstance(p, dict) and "corporate" in p]
        persons = [p for p in people if isinstance(p, dict) and "last" in p]
        if corporate:
            etree.SubElement(role_el, _bq("Corporate")).text = corporate[0][
                "corporate"
            ]
            return
        if not persons:
            raise WordMcpError(
                f"{role} entries need {{last, first?}} or {{corporate}}"
            )
        nl = etree.SubElement(role_el, _bq("NameList"))
        for p in persons:
            person = etree.SubElement(nl, _bq("Person"))
            etree.SubElement(person, _bq("Last")).text = p["last"]
            if p.get("first"):
                etree.SubElement(person, _bq("First")).text = p["first"]
            if p.get("middle"):
                etree.SubElement(person, _bq("Middle")).text = p["middle"]

    if authors:
        add_role("Author", authors)
    if editors:
        add_role("Editor", editors)

    values = {
        "Title": title,
        "JournalName": journal_name,
        "BookTitle": book_title,
        "InternetSiteTitle": internet_site_title,
        "Year": year,
        "City": city,
        "Publisher": publisher,
        "Institution": institution,
        "Edition": edition,
        "Pages": pages,
        "Volume": volume,
        "Issue": issue,
        "URL": url,
    }
    if extra_fields:
        unknown = set(extra_fields) - set(_FIELD_ORDER)
        if unknown:
            raise WordMcpError(
                f"unknown extra_fields {sorted(unknown)}; "
                f"allowed: {_FIELD_ORDER}"
            )
        values.update(extra_fields)
    for name in _FIELD_ORDER:
        val = values.get(name)
        if val is not None and str(val) != "":
            etree.SubElement(src, _bq(name)).text = str(val)

    pkg.mark_dirty(part)
    return {"source_added": tag, "type": source_type, "store": part}


def list_sources(pkg: DocxPackage) -> list[dict]:
    part = _find_store(pkg)
    if not part:
        return []
    out = []
    for s in pkg.root(part).findall(_bq("Source")):
        entry = {
            "tag": s.findtext(_bq("Tag")),
            "type": s.findtext(_bq("SourceType")),
            "title": s.findtext(_bq("Title")),
            "year": s.findtext(_bq("Year")),
        }
        persons = [
            (p.findtext(_bq("Last")) or "")
            for p in s.iter(_bq("Person"))
        ]
        corporate = [c.text for c in s.iter(_bq("Corporate")) if c.text]
        entry["authors"] = persons or corporate
        out.append(entry)
    return out


def _cited_tags(pkg: DocxPackage) -> set[str]:
    tags = set()
    for it in pkg.root().iter(qn("w:instrText")):
        text = it.text or ""
        m = re.search(r"\bCITATION\s+([A-Za-z][A-Za-z0-9_\-]*)", text)
        if m:
            tags.add(m.group(1))
        for m2 in re.finditer(r"\\m\s+([A-Za-z][A-Za-z0-9_\-]*)", text):
            tags.add(m2.group(1))
    return tags


def delete_source(pkg: DocxPackage, tag: str, *, force: bool = False) -> dict:
    part = _find_store(pkg)
    if not part:
        raise TargetNotFound("document has no bibliography sources")
    root = pkg.root(part)
    src = next(
        (
            s
            for s in root.findall(_bq("Source"))
            if (s.findtext(_bq("Tag")) or "") == tag
        ),
        None,
    )
    if src is None:
        raise TargetNotFound(f"no source with tag {tag!r}")
    if tag in _cited_tags(pkg) and not force:
        raise WordMcpError(
            f"source {tag!r} is cited in the document; delete the citations "
            "first or pass force=True (citations will render as errors)"
        )
    root.remove(src)
    pkg.mark_dirty(part)
    return {"source_deleted": tag}


def set_bibliography_style(pkg: DocxPackage, style: str) -> dict:
    """Citation/bibliography style for the document. One of the 12 verified
    Word styles (APA, Chicago, MLA, IEEE, Turabian, ...). Re-render via
    com_refresh_fields or on next open."""
    if style not in STYLES:
        raise WordMcpError(f"style must be one of {sorted(STYLES)}")
    part = _ensure_store(pkg, style=style)
    root = pkg.root(part)
    style_name, style_xsl = STYLES[style]
    root.set("SelectedStyle", style_xsl)
    root.set("StyleName", style_name)
    pkg.mark_dirty(part)
    return {"bibliography_style": style}


# -------------------------------------------------------------------- fields


def _source_placeholder(pkg: DocxPackage, tag: str) -> str:
    part = _find_store(pkg)
    if part:
        for s in pkg.root(part).findall(_bq("Source")):
            if (s.findtext(_bq("Tag")) or "") == tag:
                last = next(
                    (
                        p.findtext(_bq("Last"))
                        for p in s.iter(_bq("Person"))
                        if p.findtext(_bq("Last"))
                    ),
                    None,
                )
                corp = next(
                    (c.text for c in s.iter(_bq("Corporate")) if c.text), None
                )
                year = s.findtext(_bq("Year")) or "n.d."
                name = last or corp or tag
                return f"({name}, {year})"
    return f"({tag})"


def insert_citation(
    pkg: DocxPackage,
    *,
    tag: str,
    anchor_text: str,
    occurrence: int = 1,
    pages: str | None = None,
    suppress_author: bool = False,
    suppress_year: bool = False,
    prefix: str | None = None,
    suffix: str | None = None,
) -> dict:
    """Insert a CITATION field for a source tag right after `anchor_text`.
    Renders per the document's bibliography style on field update; a
    pre-rendered (Author, Year) placeholder shows until then."""
    part = _find_store(pkg)
    if not part or tag not in {
        s.findtext(_bq("Tag")) for s in pkg.root(part).findall(_bq("Source"))
    }:
        raise TargetNotFound(
            f"no source with tag {tag!r}; add_source first"
        )

    seen = 0
    target = None
    for p in pkg.root().iter(qn("w:p")):
        text, _ = _runmap.build_map(p)
        for m in re.finditer(re.escape(anchor_text), text):
            seen += 1
            if seen == occurrence:
                target = (p, m.end())
                break
        if target:
            break
    if target is None:
        raise TargetNotFound(
            f"anchor text not found: {anchor_text!r}"
            + (f" (occurrence {occurrence}, saw {seen})" if seen else "")
        )
    p, end = target
    covered = _runmap.split_for_range(p, end - 1, end)

    instr = f" CITATION {tag} \\l 1033"
    if pages:
        instr += f" \\p {pages}"
    if suppress_author:
        instr += " \\n"
    if suppress_year:
        instr += " \\y"
    if prefix:
        instr += f' \\f "{prefix}"'
    if suffix:
        instr += f' \\s "{suffix}"'
    instr += " "

    sdt = etree.Element(qn("w:sdt"))
    sdtpr = etree.SubElement(sdt, qn("w:sdtPr"))
    id_el = etree.SubElement(sdtpr, qn("w:id"))
    id_el.set(qn("w:val"), str(-(abs(hash(tag + str(seen))) % 2000000000) - 1))
    etree.SubElement(sdtpr, qn("w:citation"))
    content = etree.SubElement(sdt, qn("w:sdtContent"))
    r1 = etree.SubElement(content, qn("w:r"))
    begin = etree.SubElement(r1, qn("w:fldChar"))
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    r2 = etree.SubElement(content, qn("w:r"))
    it = etree.SubElement(r2, qn("w:instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = instr
    r3 = etree.SubElement(content, qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    r4 = etree.SubElement(content, qn("w:r"))
    rpr = etree.SubElement(r4, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:noProof"))
    t = etree.SubElement(r4, qn("w:t"))
    t.text = _source_placeholder(pkg, tag)
    r5 = etree.SubElement(content, qn("w:r"))
    etree.SubElement(r5, qn("w:fldChar")).set(qn("w:fldCharType"), "end")

    covered[-1].addnext(sdt)
    pkg.mark_dirty()
    return {"citation_inserted": tag, "placeholder": t.text}


def insert_bibliography(
    pkg: DocxPackage,
    *,
    title: str | None = "Bibliography",
    after_index: int | None = None,
    at_end: bool = True,
    update_on_open: bool = True,
) -> dict:
    """Insert a BIBLIOGRAPHY field (renders the full reference list from the
    source store in the selected style on field update)."""
    if _find_store(pkg) is None:
        raise WordMcpError("no bibliography sources; add_source first")
    _ensure_bibliography_style_def(pkg)

    els: list[etree._Element] = []
    if title:
        from .text import ensure_heading_style

        style_id = ensure_heading_style(pkg, 1)
        hp = etree.Element(qn("w:p"))
        ppr = etree.SubElement(hp, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), style_id)
        r = etree.SubElement(hp, qn("w:r"))
        etree.SubElement(r, qn("w:t")).text = title
        els.append(hp)

    fp = etree.Element(qn("w:p"))
    ppr = etree.SubElement(fp, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "Bibliography")
    r1 = etree.SubElement(fp, qn("w:r"))
    begin = etree.SubElement(r1, qn("w:fldChar"))
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    r2 = etree.SubElement(fp, qn("w:r"))
    it = etree.SubElement(r2, qn("w:instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = " BIBLIOGRAPHY "
    r3 = etree.SubElement(fp, qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    r4 = etree.SubElement(fp, qn("w:r"))
    rpr = etree.SubElement(r4, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:noProof"))
    t4 = etree.SubElement(r4, qn("w:t"))
    t4.text = "Update fields to generate the bibliography."
    ep = etree.Element(qn("w:p"))
    r5 = etree.SubElement(ep, qn("w:r"))
    etree.SubElement(r5, qn("w:fldChar")).set(qn("w:fldCharType"), "end")
    els += [fp, ep]

    body = pkg.body()
    if at_end and after_index is None:
        sectpr = body.find(qn("w:sectPr"))
        for el in els:
            if sectpr is not None:
                sectpr.addprevious(el)
            else:
                body.append(el)
    else:
        from .text import _body_paragraph

        ref = _body_paragraph(pkg, after_index)
        for el in reversed(els):
            ref.addnext(el)
    pkg.mark_dirty()
    if update_on_open:
        from .toc import set_update_fields_flag

        set_update_fields_flag(pkg, True)
    return {"bibliography_inserted": True, "title": title}


def _ensure_bibliography_style_def(pkg: DocxPackage) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    if "Bibliography" in have:
        return
    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), "paragraph")
    s.set(qn("w:styleId"), "Bibliography")
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), "Bibliography")
    etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
    ppr = etree.SubElement(s, qn("w:pPr"))
    ind = etree.SubElement(ppr, qn("w:ind"))
    ind.set(qn("w:left"), "720")
    ind.set(qn("w:hanging"), "720")
    pkg.mark_dirty("word/styles.xml")
