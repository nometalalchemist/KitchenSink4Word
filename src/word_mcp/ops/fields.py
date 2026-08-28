"""Field-based features: bookmarks, cross-references, captions, hyperlinks.

TOC lives in ops/toc.py (Phase 5); this module holds the smaller field types.
Field results are left empty/dirty where Word must compute them — Word fills
them on open when settings.xml carries w:updateFields (see toc.set_update_flag).
"""

from __future__ import annotations

import re

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap
from .read import body_items

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _find_anchor_span(pkg: DocxPackage, anchor_text: str, occurrence: int = 1):
    seen = 0
    for p in pkg.root().iter(qn("w:p")):
        text, _ = _runmap.build_map(p)
        for m in re.finditer(re.escape(anchor_text), text):
            seen += 1
            if seen == occurrence:
                return p, m.start(), m.end()
    raise TargetNotFound(
        f"anchor text not found: {anchor_text!r}"
        + (f" (occurrence {occurrence}, saw {seen})" if seen else "")
    )


def _max_bookmark_id(pkg: DocxPackage) -> int:
    ids = [
        int(b.get(qn("w:id"), "0"))
        for b in pkg.root().iter(qn("w:bookmarkStart"))
    ]
    return max(ids, default=0)


def add_bookmark(
    pkg: DocxPackage, name: str, *, anchor_text: str, occurrence: int = 1
) -> dict:
    """Wrap `anchor_text` in a named bookmark (target for cross-references)."""
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,39}", name):
        raise WordMcpError(
            "bookmark names: letters/digits/underscore, start with a letter, "
            "max 40 chars"
        )
    existing = {
        b.get(qn("w:name")) for b in pkg.root().iter(qn("w:bookmarkStart"))
    }
    if name in existing:
        raise WordMcpError(f"bookmark {name!r} already exists")
    p, start, end = _find_anchor_span(pkg, anchor_text, occurrence)
    runs = _runmap.split_for_range(p, start, end)
    bid = str(_max_bookmark_id(pkg) + 1)
    bs = etree.Element(qn("w:bookmarkStart"))
    bs.set(qn("w:id"), bid)
    bs.set(qn("w:name"), name)
    be = etree.Element(qn("w:bookmarkEnd"))
    be.set(qn("w:id"), bid)
    runs[0].addprevious(bs)
    runs[-1].addnext(be)
    pkg.mark_dirty()
    return {"bookmark": name, "id": bid}


def list_bookmarks(pkg: DocxPackage) -> list[dict]:
    out = []
    for b in pkg.root().iter(qn("w:bookmarkStart")):
        name = b.get(qn("w:name"))
        if name and not name.startswith("_"):  # skip Word-internal bookmarks
            out.append({"name": name, "id": b.get(qn("w:id"))})
    return out


def _field_run(instr: str, placeholder: str = "") -> list[etree._Element]:
    """Complex field: begin(dirty) / instrText / separate / placeholder / end.
    Marked dirty so Word recomputes the result on open."""
    els = []
    r1 = etree.Element(qn("w:r"))
    fc = etree.SubElement(r1, qn("w:fldChar"))
    fc.set(qn("w:fldCharType"), "begin")
    fc.set(qn("w:dirty"), "true")
    els.append(r1)
    r2 = etree.Element(qn("w:r"))
    it = etree.SubElement(r2, qn("w:instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = f" {instr} "
    els.append(r2)
    r3 = etree.Element(qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    els.append(r3)
    if placeholder:
        r4 = etree.Element(qn("w:r"))
        t = etree.SubElement(r4, qn("w:t"))
        t.text = placeholder
        els.append(r4)
    r5 = etree.Element(qn("w:r"))
    etree.SubElement(r5, qn("w:fldChar")).set(qn("w:fldCharType"), "end")
    els.append(r5)
    return els


def add_cross_reference(
    pkg: DocxPackage,
    *,
    after_anchor: str,
    to_bookmark: str,
    kind: str = "page",
    occurrence: int = 1,
) -> dict:
    """Insert a cross-reference after `after_anchor`. kind: 'page' (PAGEREF —
    page number) or 'text' (REF — bookmarked text)."""
    existing = {
        b.get(qn("w:name")) for b in pkg.root().iter(qn("w:bookmarkStart"))
    }
    if to_bookmark not in existing:
        raise TargetNotFound(
            f"bookmark {to_bookmark!r} does not exist; create it first"
        )
    if kind == "page":
        instr = f"PAGEREF {to_bookmark} \\h"
    elif kind == "text":
        instr = f"REF {to_bookmark} \\h"
    else:
        raise WordMcpError("kind must be 'page' or 'text'")
    p, _, end = _find_anchor_span(pkg, after_anchor, occurrence)
    covered = _runmap.split_for_range(p, end - 1, end)
    ref = covered[-1]
    for el in reversed(_field_run(instr, placeholder="#")):
        ref.addnext(el)
    pkg.mark_dirty()
    return {"cross_reference": to_bookmark, "kind": kind}


def add_caption(
    pkg: DocxPackage,
    *,
    table_index: int | None = None,
    after_anchor: str | None = None,
    label: str = "Table",
    text: str,
    above: bool = True,
) -> dict:
    """Insert a numbered caption ('Table SEQ: text') above/below a table or at
    an anchor. Numbering uses a SEQ field, recomputed by Word."""
    if label not in ("Table", "Figure", "Equation"):
        raise WordMcpError("label must be Table, Figure, or Equation")
    cap = etree.Element(qn("w:p"))
    ppr = etree.SubElement(cap, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "Caption")
    r = etree.SubElement(cap, qn("w:r"))
    t = etree.SubElement(r, qn("w:t"))
    t.text = f"{label} "
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for el in _field_run(f"SEQ {label} \\* ARABIC", placeholder="#"):
        cap.append(el)
    r2 = etree.SubElement(cap, qn("w:r"))
    t2 = etree.SubElement(r2, qn("w:t"))
    t2.text = f": {text}"
    t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    _ensure_caption_style(pkg)

    if table_index is not None:
        tbl = None
        for k, idx, el in body_items(pkg):
            if k == "table" and idx == table_index:
                tbl = el
                break
        if tbl is None:
            raise TargetNotFound(f"no table with index {table_index}")
        if above:
            tbl.addprevious(cap)
        else:
            tbl.addnext(cap)
    elif after_anchor is not None:
        p, _, _ = _find_anchor_span(pkg, after_anchor)
        p.addnext(cap)
    else:
        raise WordMcpError("give table_index or after_anchor")
    pkg.mark_dirty()
    return {"caption": f"{label}: {text}"}


def _ensure_caption_style(pkg: DocxPackage) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    if "Caption" in have:
        return
    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), "paragraph")
    s.set(qn("w:styleId"), "Caption")
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), "caption")
    etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
    ppr = etree.SubElement(s, qn("w:pPr"))
    spacing = etree.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:after"), "200")
    rpr = etree.SubElement(s, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:b"))
    etree.SubElement(rpr, qn("w:sz")).set(qn("w:val"), "18")
    etree.SubElement(rpr, qn("w:szCs")).set(qn("w:val"), "18")
    pkg.mark_dirty("word/styles.xml")


def mark_index_entry(
    pkg: DocxPackage,
    *,
    anchor_text: str,
    entry: str,
    subentry: str | None = None,
    occurrence: int = 1,
    bold_page: bool = False,
    italic_page: bool = False,
    see: str | None = None,
) -> dict:
    """Mark a location for the index: inserts an invisible XE field after
    `anchor_text`. subentry nests under entry; see='other entry' makes a
    cross-reference instead of a page number."""
    p, _, end = _find_anchor_span(pkg, anchor_text, occurrence)
    covered = _runmap.split_for_range(p, end - 1, end)

    key = entry.replace('"', '\\"').replace(":", "\\:")
    if subentry:
        key += ":" + subentry.replace('"', '\\"').replace(":", "\\:")
    instr = f' XE "{key}"'
    if bold_page:
        instr += " \\b"
    if italic_page:
        instr += " \\i"
    if see:
        instr += f' \\t "See {see}"'
    instr += " "

    # Ground truth: XE = bare begin/instrText/end runs, NO separate, NO rPr.
    els = []
    r1 = etree.Element(qn("w:r"))
    etree.SubElement(r1, qn("w:fldChar")).set(qn("w:fldCharType"), "begin")
    els.append(r1)
    r2 = etree.Element(qn("w:r"))
    it = etree.SubElement(r2, qn("w:instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = instr
    els.append(r2)
    r3 = etree.Element(qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "end")
    els.append(r3)
    ref = covered[-1]
    for el in reversed(els):
        ref.addnext(el)
    pkg.mark_dirty()
    return {"index_entry": entry, "subentry": subentry, "see": see}


def list_index_entries(pkg: DocxPackage) -> list[dict]:
    """All XE index entries in the document."""
    import re as _re

    out = []
    for it in pkg.root().iter(qn("w:instrText")):
        text = (it.text or "").strip()
        if text.startswith("XE"):
            m = _re.search(r'XE\s+"((?:[^"\\]|\\.)*)"', text)
            if m:
                raw = m.group(1)
                parts = _re.split(r"(?<!\\):", raw)
                out.append(
                    {
                        "entry": parts[0].replace('\\"', '"').replace("\\:", ":"),
                        "subentry": (
                            parts[1].replace('\\"', '"').replace("\\:", ":")
                            if len(parts) > 1
                            else None
                        ),
                        "raw": text,
                    }
                )
    return out


def insert_index(
    pkg: DocxPackage,
    *,
    title: str | None = "Index",
    columns: int = 2,
    letter_headings: bool = True,
    after_index: int | None = None,
    at_end: bool = True,
    update_on_open: bool = True,
) -> dict:
    """Insert an INDEX field that compiles all XE entries into an index (with
    page numbers) when Word updates fields."""
    if not list_index_entries(pkg):
        raise WordMcpError(
            "no XE index entries in the document; mark_index_entry first"
        )
    if not 1 <= columns <= 4:
        raise WordMcpError("columns must be 1-4")
    _ensure_index_styles(pkg)

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

    instr = f' INDEX \\c "{columns}"'
    if letter_headings:
        instr += ' \\h "A"'
    instr += ' \\z "1033" '
    fp = etree.Element(qn("w:p"))
    r1 = etree.SubElement(fp, qn("w:r"))
    begin = etree.SubElement(r1, qn("w:fldChar"))
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    r2 = etree.SubElement(fp, qn("w:r"))
    it = etree.SubElement(r2, qn("w:instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = instr
    r3 = etree.SubElement(fp, qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    r4 = etree.SubElement(fp, qn("w:r"))
    t4 = etree.SubElement(r4, qn("w:t"))
    t4.text = "Update fields to generate the index."
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
    return {
        "index_inserted": True,
        "entries": len(list_index_entries(pkg)),
        "columns": columns,
    }


def _ensure_index_styles(pkg: DocxPackage) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    dirty = False
    for i in (1, 2):
        sid = f"Index{i}"
        if sid not in have:
            s = etree.SubElement(root, qn("w:style"))
            s.set(qn("w:type"), "paragraph")
            s.set(qn("w:styleId"), sid)
            etree.SubElement(s, qn("w:name")).set(qn("w:val"), f"index {i}")
            etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
            ppr = etree.SubElement(s, qn("w:pPr"))
            ind = etree.SubElement(ppr, qn("w:ind"))
            ind.set(qn("w:left"), str(220 * (i - 1)))
            ind.set(qn("w:hanging"), "220")
            dirty = True
    if "IndexHeading" not in have:
        s = etree.SubElement(root, qn("w:style"))
        s.set(qn("w:type"), "paragraph")
        s.set(qn("w:styleId"), "IndexHeading"),
        etree.SubElement(s, qn("w:name")).set(qn("w:val"), "index heading")
        etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
        rpr = etree.SubElement(s, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:b"))
        dirty = True
    if dirty:
        pkg.mark_dirty("word/styles.xml")


def add_hyperlink(
    pkg: DocxPackage, *, anchor_text: str, url: str, occurrence: int = 1
) -> dict:
    """Turn `anchor_text` into an external hyperlink."""
    p, start, end = _find_anchor_span(pkg, anchor_text, occurrence)
    runs = _runmap.split_for_range(p, start, end)
    for run in runs:
        if etree.QName(run.getparent()).localname == "hyperlink":
            raise WordMcpError("text is already inside a hyperlink")

    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    existing_ids = {r.get("Id") for r in rels_root}
    n = 1
    while f"rId{n}" in existing_ids:
        n += 1
    rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", f"rId{n}")
    rel.set(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    )
    rel.set("Target", url)
    rel.set("TargetMode", "External")
    pkg.mark_dirty(rels_part)

    link = etree.Element(qn("w:hyperlink"))
    link.set(qn("r:id"), f"rId{n}")
    runs[0].addprevious(link)
    for run in runs:
        link.append(run)  # moves the runs inside the wrapper
        rpr = run.find(qn("w:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("w:rPr"))
            run.insert(0, rpr)
        if rpr.find(qn("w:rStyle")) is None:
            st = etree.Element(qn("w:rStyle"))
            st.set(qn("w:val"), "Hyperlink")
            rpr.insert(0, st)
    _ensure_hyperlink_style(pkg)
    pkg.mark_dirty()
    return {"hyperlink": url, "text": anchor_text}


# --------------------------------------------------------- generic field codes

# Known-safe field codes for insert_field. Deliberately excludes anything that
# executes, links out, or rewrites content on update (INCLUDETEXT, LINK, DDE,
# AUTOTEXT, MACROBUTTON, IMPORT, GOTOBUTTON...): those are how field codes
# become an attack surface, and the reference/TOC/index field families already
# have dedicated tools with proper target validation.
FIELD_ALLOWLIST = frozenset(
    {
        "DATE", "TIME", "CREATEDATE", "SAVEDATE", "PRINTDATE", "EDITTIME",
        "FILENAME", "FILESIZE",
        "PAGE", "NUMPAGES", "NUMWORDS", "NUMCHARS", "SECTION", "SECTIONPAGES",
        "AUTHOR", "TITLE", "SUBJECT", "KEYWORDS", "COMMENTS", "LASTSAVEDBY",
        "USERNAME", "USERINITIALS",
        "SEQ",
    }
)

# Instruction text after validation may only contain these characters: word
# chars, whitespace, and the switch/format punctuation Word field syntax uses.
_FIELD_SAFE_RE = re.compile(r"""[^\w\s\\*@#"'.,:;/()\-%]""")


def _validate_field_code(field_code: str) -> str:
    code = " ".join(field_code.split())
    if not code:
        raise WordMcpError("field_code must be non-empty")
    keyword = code.split()[0].upper()
    if keyword not in FIELD_ALLOWLIST:
        raise WordMcpError(
            f"field code {keyword!r} is not on the allowlist of known-safe "
            f"codes: {sorted(FIELD_ALLOWLIST)}. Reference fields (REF/"
            "PAGEREF/SEQ captions), TOC, and INDEX have dedicated tools; "
            "codes that link out or execute are refused by design"
        )
    bad = _FIELD_SAFE_RE.search(code)
    if bad:
        raise WordMcpError(
            f"field code contains unsupported character {bad.group()!r}; "
            "only field keywords, arguments, and standard switches "
            '(\\* \\@ \\# formats, quoted strings) are accepted'
        )
    if code.count('"') % 2:
        raise WordMcpError("field code has an unbalanced quote")
    if keyword == "SEQ":
        if not re.match(r"(?i)SEQ\s+[A-Za-z_][A-Za-z0-9_]*(\s|$)", code):
            raise WordMcpError(
                "SEQ needs an identifier name, e.g. 'SEQ Exhibit \\* Arabic'"
            )
    return code


def insert_field(
    pkg: DocxPackage,
    *,
    field_code: str,
    after_anchor: str,
    occurrence: int = 1,
    placeholder: str = "",
) -> dict:
    """Insert a generic field (DATE, TIME, FILENAME, NUMPAGES, PAGE,
    SEQ <name>..., etc.) immediately after `after_anchor` text. The code is
    validated against FIELD_ALLOWLIST; anything else is refused naming the
    allowlist. The field is written dirty, so Word recomputes the result the
    next time fields refresh; `placeholder` is what shows until then."""
    code = _validate_field_code(field_code)
    p, _, end = _find_anchor_span(pkg, after_anchor, occurrence)
    covered = _runmap.split_for_range(p, end - 1, end)
    ref = covered[-1]
    for el in reversed(_field_run(code, placeholder=placeholder)):
        ref.addnext(el)
    pkg.mark_dirty()
    return {"field_inserted": code, "after": after_anchor}


def _story_parts_for_fields(pkg: DocxPackage) -> list[str]:
    parts = ["word/document.xml"]
    for name in sorted(pkg.part_names()):
        if re.fullmatch(r"word/(header|footer)\d+\.xml", name):
            parts.append(name)
    for name in ("word/footnotes.xml", "word/endnotes.xml"):
        if pkg.has_part(name):
            parts.append(name)
    return parts


def list_fields(pkg: DocxPackage) -> dict:
    """Every field in the body, headers, footers, footnotes, and endnotes:
    instruction code, field type (first keyword), current cached result, and
    location (part + body paragraph index where applicable). Covers complex
    (fldChar) and simple (fldSimple) fields; unclosed complex fields are
    flagged per entry."""
    from .reffields import (
        _body_paragraph_index_map,
        _containing_paragraph,
        _locate,
        scan_complex_fields,
    )

    index_map, _keepalive = _body_paragraph_index_map(pkg)
    out: list[dict] = []
    parts_scanned: list[str] = []
    for part in _story_parts_for_fields(pkg):
        if not pkg.has_part(part):
            continue
        parts_scanned.append(part)
        root = pkg.root(part)
        fields, _orphans = scan_complex_fields(root)
        for rec in fields:
            instr = rec["instr"]
            entry = {
                "code": instr,
                "type": instr.split()[0].upper() if instr.split() else "",
                "kind": "complex",
                "cached_result": rec["cached"],
                **_locate(rec["begin_el"], part, index_map),
            }
            if not rec["closed"]:
                entry["unclosed"] = True
            out.append(entry)
        for fs in root.iter(qn("w:fldSimple")):
            instr = " ".join((fs.get(qn("w:instr")) or "").split())
            cached = "".join(t.text or "" for t in fs.iter(qn("w:t")))
            p = _containing_paragraph(fs)
            loc = _locate(
                p if p is not None else fs, part, index_map
            )
            out.append(
                {
                    "code": instr,
                    "type": instr.split()[0].upper() if instr.split() else "",
                    "kind": "simple",
                    "cached_result": cached,
                    **loc,
                }
            )
    return {"fields": out, "total": len(out), "parts_scanned": parts_scanned}


def _ensure_hyperlink_style(pkg: DocxPackage) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    if "Hyperlink" in have:
        return
    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), "character")
    s.set(qn("w:styleId"), "Hyperlink")
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), "Hyperlink")
    etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "DefaultParagraphFont")
    rpr = etree.SubElement(s, qn("w:rPr"))
    color = etree.SubElement(rpr, qn("w:color"))
    color.set(qn("w:val"), "0563C1")
    color.set(qn("w:themeColor"), "hyperlink")
    etree.SubElement(rpr, qn("w:u")).set(qn("w:val"), "single")
    pkg.mark_dirty("word/styles.xml")
