"""Table of contents: insert, read, delete, refresh flag.

A TOC is a complex field (begin / instrText / separate / cached-result / end),
optionally wrapped in Word's SDT chrome. Page numbers exist only after Word
computes them: dirty="true" on the field asks politely; the settings.xml
w:updateFields flag forces recalculation on next open (with a prompt) and Word
removes the flag afterward; the COM module (Phase 6) refreshes immediately.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn


def insert_toc(
    pkg: DocxPackage,
    *,
    levels: str = "1-3",
    title: str | None = "Table of Contents",
    after_index: int | None = None,
    at_start: bool = False,
    update_on_open: bool = True,
) -> dict:
    """Insert a TOC. Placement: after body paragraph `after_index`, or at the
    document start. The field is marked dirty; with update_on_open Word also
    recalculates every field on next open (one confirmation prompt, one time)."""
    if _find_toc_field(pkg) is not None:
        raise WordMcpError(
            "document already contains a TOC; delete_toc first or edit it in Word"
        )

    _ensure_toc_styles(pkg, with_title=title is not None)

    sdt = etree.Element(qn("w:sdt"))
    sdtpr = etree.SubElement(sdt, qn("w:sdtPr"))
    dpo = etree.SubElement(sdtpr, qn("w:docPartObj"))
    etree.SubElement(dpo, qn("w:docPartGallery")).set(
        qn("w:val"), "Table of Contents"
    )
    etree.SubElement(dpo, qn("w:docPartUnique"))
    content = etree.SubElement(sdt, qn("w:sdtContent"))

    if title is not None:
        tp = etree.SubElement(content, qn("w:p"))
        ppr = etree.SubElement(tp, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "TOCHeading")
        r = etree.SubElement(tp, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = title

    fp = etree.SubElement(content, qn("w:p"))
    r1 = etree.SubElement(fp, qn("w:r"))
    begin = etree.SubElement(r1, qn("w:fldChar"))
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    r2 = etree.SubElement(fp, qn("w:r"))
    instr = etree.SubElement(r2, qn("w:instrText"))
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = f' TOC \\o "{levels}" \\h \\z \\u '
    r3 = etree.SubElement(fp, qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    r4 = etree.SubElement(fp, qn("w:r"))
    t4 = etree.SubElement(r4, qn("w:t"))
    t4.text = "Right-click and choose Update Field to populate this table."
    ep = etree.SubElement(content, qn("w:p"))
    r5 = etree.SubElement(ep, qn("w:r"))
    etree.SubElement(r5, qn("w:fldChar")).set(qn("w:fldCharType"), "end")

    body = pkg.body()
    if at_start or after_index is None:
        first = body.find(qn("w:p"))
        if first is not None:
            first.addprevious(sdt)
        else:
            sectpr = body.find(qn("w:sectPr"))
            if sectpr is not None:
                sectpr.addprevious(sdt)
            else:
                body.append(sdt)
    else:
        from .text import _body_paragraph

        _body_paragraph(pkg, after_index).addnext(sdt)
    pkg.mark_dirty()

    if update_on_open:
        set_update_fields_flag(pkg, True)
    return {"toc_inserted": True, "levels": levels, "update_on_open": update_on_open}


def insert_caption_list(
    pkg: DocxPackage,
    *,
    label: str = "Table",
    title: str | None = None,
    after_index: int | None = None,
    at_start: bool = False,
    update_on_open: bool = True,
) -> dict:
    """Insert a List of Tables/Figures/Equations: a TOC field over SEQ
    captions (TOC \\h \\z \\c "label"). Entries come from add_caption; page
    numbers appear on field update like the main TOC."""
    if label not in ("Table", "Figure", "Equation"):
        raise WordMcpError("label must be Table, Figure, or Equation")
    # One list per label.
    for p in _find_toc_fields(pkg):
        instr = " ".join(it.text or "" for it in p.iter(qn("w:instrText")))
        if f'\\c "{label}"' in instr:
            raise WordMcpError(
                f"document already has a caption list for {label}; delete it "
                "first (see read_toc for its index)"
            )
    if title is None:
        title = f"List of {label}s"
    _ensure_toc_styles(pkg, with_title=True)

    sdt = etree.Element(qn("w:sdt"))
    sdtpr = etree.SubElement(sdt, qn("w:sdtPr"))
    dpo = etree.SubElement(sdtpr, qn("w:docPartObj"))
    etree.SubElement(dpo, qn("w:docPartGallery")).set(
        qn("w:val"), "Table of Contents"
    )
    etree.SubElement(dpo, qn("w:docPartUnique"))
    content = etree.SubElement(sdt, qn("w:sdtContent"))
    tp = etree.SubElement(content, qn("w:p"))
    ppr = etree.SubElement(tp, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "TOCHeading")
    r = etree.SubElement(tp, qn("w:r"))
    etree.SubElement(r, qn("w:t")).text = title

    fp = etree.SubElement(content, qn("w:p"))
    r1 = etree.SubElement(fp, qn("w:r"))
    begin = etree.SubElement(r1, qn("w:fldChar"))
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    r2 = etree.SubElement(fp, qn("w:r"))
    instr_el = etree.SubElement(r2, qn("w:instrText"))
    instr_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_el.text = f' TOC \\h \\z \\c "{label}" '
    r3 = etree.SubElement(fp, qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    r4 = etree.SubElement(fp, qn("w:r"))
    etree.SubElement(r4, qn("w:t")).text = (
        "Right-click and choose Update Field to populate this list."
    )
    ep = etree.SubElement(content, qn("w:p"))
    r5 = etree.SubElement(ep, qn("w:r"))
    etree.SubElement(r5, qn("w:fldChar")).set(qn("w:fldCharType"), "end")

    body = pkg.body()
    if at_start or after_index is None:
        first = body.find(qn("w:p"))
        if first is not None:
            first.addprevious(sdt)
        else:
            sectpr = body.find(qn("w:sectPr"))
            if sectpr is not None:
                sectpr.addprevious(sdt)
            else:
                body.append(sdt)
    else:
        from .text import _body_paragraph

        _body_paragraph(pkg, after_index).addnext(sdt)
    pkg.mark_dirty()
    if update_on_open:
        set_update_fields_flag(pkg, True)
    return {"caption_list_inserted": label, "title": title}


def _find_toc_fields(pkg: DocxPackage) -> list[etree._Element]:
    """Paragraphs containing a TOC-family field's instruction, in document
    order. Includes List of Tables/Figures variants (TOC \\c)."""
    out = []
    for p in pkg.root().iter(qn("w:p")):
        for it in p.iter(qn("w:instrText")):
            if it.text and it.text.strip().startswith("TOC"):
                out.append(p)
                break
    return out


def _find_toc_field(pkg: DocxPackage):
    """First MAIN TOC (one without the \\c caption switch), or None. Used by
    insert_toc's duplicate check so a List of Tables does not block insertion."""
    for p in _find_toc_fields(pkg):
        instr = " ".join(it.text or "" for it in p.iter(qn("w:instrText")))
        if "\\c" not in instr:
            return p
    return None


def read_toc(pkg: DocxPackage) -> dict:
    """Report every TOC-family field and its cached entries. A document can
    hold several (main TOC, List of Tables, List of Figures)."""
    fields = _find_toc_fields(pkg)
    if not fields:
        return {"present": False, "tocs": []}
    from .read import paragraph_text

    tocs = []
    for i, field_p in enumerate(fields):
        # The field's OWN instruction: instrText runs before the separate
        # marker only (updated results embed nested PAGEREF instructions).
        instr_parts = []
        for el in field_p.iter():
            local = etree.QName(el).localname
            if local == "fldChar" and el.get(qn("w:fldCharType")) == "separate":
                break
            if local == "instrText":
                instr_parts.append(el.text or "")
        instr = " ".join(instr_parts).strip()
        entries = []
        scope = field_p.getparent()
        for p in scope.iter(qn("w:p")):
            ps = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
            style = ps.get(qn("w:val")) if ps is not None else None
            if style and (
                (style.startswith("TOC") and style != "TOCHeading")
                or style.startswith("TableofFigures")
            ):
                entries.append({"style": style, "text": paragraph_text(p)})
        kind = "caption_list" if "\\c" in instr else "main"
        tocs.append(
            {
                "index": i,
                "kind": kind,
                "instruction": instr,
                "cached_entries": entries if i == 0 or kind != "main" else entries,
            }
        )
    first = tocs[0]
    return {
        "present": True,
        "instruction": first["instruction"],
        "cached_entries": first["cached_entries"],
        "tocs": tocs,
    }


def delete_toc(pkg: DocxPackage, *, which: int = 0) -> dict:
    """Remove one TOC-family field by index (see read_toc's 'tocs' list):
    its SDT wrapper if present, else every paragraph from begin through end."""
    fields = _find_toc_fields(pkg)
    if not fields:
        raise TargetNotFound("no TOC field in this document")
    if not 0 <= which < len(fields):
        raise TargetNotFound(
            f"TOC index {which} out of range (document has {len(fields)})"
        )
    field_p = fields[which]
    # SDT-wrapped case.
    node = field_p
    while node is not None:
        if etree.QName(node).localname == "sdt":
            node.getparent().remove(node)
            pkg.mark_dirty()
            return {"toc_deleted": True, "form": "sdt"}
        node = node.getparent()
    # Bare-field case: walk siblings until the end marker balances.
    body = field_p.getparent()
    doomed = []
    depth = 0
    started = False
    for p in list(body):
        if not started and p is not field_p:
            continue
        started = True
        doomed.append(p)
        for fc in p.iter(qn("w:fldChar")):
            t = fc.get(qn("w:fldCharType"))
            if t == "begin":
                depth += 1
            elif t == "end":
                depth -= 1
        if depth == 0:
            break
    for p in doomed:
        body.remove(p)
    pkg.mark_dirty()
    return {"toc_deleted": True, "form": "bare", "paragraphs_removed": len(doomed)}


def set_update_fields_flag(pkg: DocxPackage, on: bool = True) -> dict:
    """settings.xml w:updateFields: Word recalculates all fields on next open
    (shows one confirmation prompt), then removes the flag itself."""
    if not pkg.has_part("word/settings.xml"):
        raise WordMcpError("document has no settings.xml")
    root = pkg.root("word/settings.xml")
    el = root.find(qn("w:updateFields"))
    if on:
        if el is None:
            el = etree.Element(qn("w:updateFields"))
            el.set(qn("w:val"), "true")
            root.insert(0, el)
    else:
        if el is not None:
            root.remove(el)
    pkg.mark_dirty("word/settings.xml")
    return {"update_fields_on_open": on}


def _ensure_toc_styles(pkg: DocxPackage, *, with_title: bool) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    dirty = False
    if with_title and "TOCHeading" not in have:
        s = etree.SubElement(root, qn("w:style"))
        s.set(qn("w:type"), "paragraph")
        s.set(qn("w:styleId"), "TOCHeading")
        etree.SubElement(s, qn("w:name")).set(qn("w:val"), "TOC Heading")
        base = "Heading1" if "Heading1" in have else "Normal"
        etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), base)
        ppr = etree.SubElement(s, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:outlineLvl")).set(qn("w:val"), "9")
        rpr = etree.SubElement(s, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:b"))
        etree.SubElement(rpr, qn("w:sz")).set(qn("w:val"), "32")
        dirty = True
    for lvl in (1, 2, 3):
        sid = f"TOC{lvl}"
        if sid not in have:
            s = etree.SubElement(root, qn("w:style"))
            s.set(qn("w:type"), "paragraph")
            s.set(qn("w:styleId"), sid)
            etree.SubElement(s, qn("w:name")).set(qn("w:val"), f"toc {lvl}")
            etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
            ppr = etree.SubElement(s, qn("w:pPr"))
            spacing = etree.SubElement(ppr, qn("w:spacing"))
            spacing.set(qn("w:after"), "100")
            if lvl > 1:
                ind = etree.SubElement(ppr, qn("w:ind"))
                ind.set(qn("w:left"), str(220 * (lvl - 1)))
            dirty = True
    if dirty:
        pkg.mark_dirty("word/styles.xml")
