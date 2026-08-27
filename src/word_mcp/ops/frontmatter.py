"""Front-matter assembly: title page, copyright page, abstract, TOC and
caption lists in one call, with the section split that lets the front matter
carry roman page numbers while the body restarts at 1.

The whole request is validated before anything is touched; a refusal changes
nothing. Page breaks separate front-matter pages; ONE section break separates
front matter from body so the two can carry different pgNumType formats.
"""

from __future__ import annotations

import copy
import re

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from . import toc as _tc
from .furniture import (
    _PAGE_NUM_FORMATS,
    _sect_prs,
    _sectpr_get_or_add,
    get_headers_footers,
)
from .read import paragraph_text
from .text import _make_paragraph, ensure_heading_style

_CAPTION_KINDS = {
    "list_of_tables": "Table",
    "list_of_figures": "Figure",
    "list_of_equations": "Equation",
}
_KNOWN_KINDS = {"title_page", "blank_or_copyright", "abstract", "toc"} | set(
    _CAPTION_KINDS
)


def _validate_spec(spec: dict) -> tuple[list[dict], str, str, int]:
    """Full up-front validation so a refusal leaves zero changes. Returns
    (sections, front_format, body_format, body_restart_at)."""
    if not isinstance(spec, dict):
        raise WordMcpError("spec must be a dict")
    sections = spec.get("sections")
    if not isinstance(sections, list) or not sections:
        raise WordMcpError('spec["sections"] must be a non-empty list')
    for i, sec in enumerate(sections):
        if not isinstance(sec, dict) or "kind" not in sec:
            raise WordMcpError(f'sections[{i}] must be a dict with a "kind"')
        kind = sec["kind"]
        if kind not in _KNOWN_KINDS:
            raise WordMcpError(
                f"sections[{i}]: unknown kind {kind!r}; "
                f"known kinds: {sorted(_KNOWN_KINDS)}"
            )
        if kind in ("title_page", "blank_or_copyright"):
            lines = sec.get("lines", [])
            if not isinstance(lines, list) or any(
                not isinstance(x, str) for x in lines
            ):
                raise WordMcpError(f'sections[{i}]: "lines" must be a list of strings')
            if kind == "title_page" and not lines:
                raise WordMcpError(f"sections[{i}]: title_page needs at least one line")
        elif kind == "abstract":
            if not isinstance(sec.get("text"), str) or not sec["text"].strip():
                raise WordMcpError(f'sections[{i}]: abstract needs non-empty "text"')
        # any section may carry a title; TOC-family sections carry levels —
        # both go straight into field codes, so type/shape-check them here
        # (validate-everything-first: a refusal must change nothing)
        if "title" in sec and not isinstance(sec["title"], str):
            raise WordMcpError(
                f'sections[{i}]: "title" must be a string, got '
                f"{type(sec['title']).__name__}"
            )
        if "levels" in sec:
            lv = sec["levels"]
            m = (
                re.fullmatch(r"([1-9])(?:-([1-9]))?", lv)
                if isinstance(lv, str)
                else None
            )
            if m is None or (
                m.group(2) and int(m.group(1)) > int(m.group(2))
            ):
                raise WordMcpError(
                    f'sections[{i}]: "levels" must look like "1-3" or "2" '
                    f"(low-high, 1-9), got {lv!r}"
                )

    numbering = spec.get("page_numbering", {})
    if not isinstance(numbering, dict):
        raise WordMcpError('spec["page_numbering"] must be a dict')
    front_fmt = numbering.get("front", "lowerRoman")
    body_fmt = numbering.get("body", "decimal")
    for label, fmt in (("front", front_fmt), ("body", body_fmt)):
        if fmt not in _PAGE_NUM_FORMATS:
            raise WordMcpError(
                f"page_numbering[{label!r}] must be one of "
                f"{sorted(_PAGE_NUM_FORMATS)}"
            )
    restart = numbering.get("body_restart_at", 1)
    if not isinstance(restart, int) or restart < 1:
        raise WordMcpError("page_numbering['body_restart_at'] must be an int >= 1")
    return sections, front_fmt, body_fmt, restart


def _existing_front_matter_reason(pkg: DocxPackage) -> str | None:
    for i, sp in enumerate(_sect_prs(pkg)):
        pgnum = sp.find(qn("w:pgNumType"))
        if pgnum is not None and pgnum.get(qn("w:fmt")) == "lowerRoman":
            return f"section {i} already carries lowerRoman page numbering"
    # A TOC within the first meaningful body blocks also signals front matter.
    checked = 0
    for child in pkg.body():
        local = etree.QName(child).localname
        if local == "sectPr":
            break
        for it in child.iter(qn("w:instrText")):
            if (it.text or "").strip().startswith("TOC"):
                return "document already starts with a TOC"
        if local == "p" and not paragraph_text(child).strip():
            continue
        checked += 1
        if checked >= 3:
            break
    return None


def _page_break_paragraph() -> etree._Element:
    p = etree.Element(qn("w:p"))
    r = etree.SubElement(p, qn("w:r"))
    etree.SubElement(r, qn("w:br")).set(qn("w:type"), "page")
    return p


def _centered_paragraph(text: str) -> etree._Element:
    p = _make_paragraph(text)
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    etree.SubElement(ppr, qn("w:jc")).set(qn("w:val"), "center")
    return p


def assemble_front_matter(pkg: DocxPackage, spec: dict) -> dict:
    """Insert standard long-document front matter at the START of the document;
    the existing content becomes the body.

    spec = {
        "sections": [
            {"kind": "title_page", "lines": [...]},          # centered lines
            {"kind": "blank_or_copyright", "lines": [...]},  # plain lines
            {"kind": "abstract", "title": "Abstract", "text": "..."},
            {"kind": "toc", "levels": "1-3", "title": ...},
            {"kind": "list_of_figures" | "list_of_tables" | "list_of_equations",
             "title": ...},
        ],
        "page_numbering": {"front": "lowerRoman", "body": "decimal",
                           "body_restart_at": 1},
        "force": bool,
    }

    Refuses when the document appears to have front matter already (starts
    with a TOC, or a section already uses lowerRoman numbering) unless
    force=true. With force, a TOC/caption-list request that duplicates an
    existing one is skipped and reported, not silently doubled.
    """
    sections, front_fmt, body_fmt, restart = _validate_spec(spec)
    if not spec.get("force"):
        reason = _existing_front_matter_reason(pkg)
        if reason is not None:
            raise WordMcpError(
                f"front matter appears to exist already ({reason}); "
                'pass "force": true to insert anyway'
            )

    body = pkg.body()
    children = list(body)
    # cursor = the first pre-existing child; everything front-matter goes
    # before it, so document order matches spec order.
    cursor = children[0] if children else None

    def insert(el: etree._Element) -> None:
        if cursor is not None:
            cursor.addprevious(el)
        else:
            body.append(el)

    inserted: list[dict] = []
    field_blocks = 0
    first = True
    for sec in sections:
        kind = sec["kind"]
        if not first:
            insert(_page_break_paragraph())
        first = False
        if kind == "title_page":
            for line in sec["lines"]:
                insert(_centered_paragraph(line))
            inserted.append({"kind": kind, "paragraphs": len(sec["lines"])})
        elif kind == "blank_or_copyright":
            lines = sec.get("lines", []) or [""]
            for line in lines:
                insert(_make_paragraph(line))
            inserted.append({"kind": kind, "paragraphs": len(lines)})
        elif kind == "abstract":
            title = sec.get("title", "Abstract")
            style_id = ensure_heading_style(pkg, 1)
            insert(_make_paragraph(title, style=style_id))
            chunks = [c for c in sec["text"].split("\n\n") if c.strip()]
            for chunk in chunks:
                insert(_make_paragraph(chunk))
            inserted.append({"kind": kind, "title": title, "paragraphs": len(chunks)})
        else:  # toc / caption lists: reuse the toc ops, then move into place
            # Hold the proxies in a list: lxml element ids are only stable
            # while a Python reference keeps the proxy alive.
            before = list(body)
            before_ids = {id(c) for c in before}
            try:
                if kind == "toc":
                    _tc.insert_toc(
                        pkg,
                        levels=sec.get("levels", "1-3"),
                        title=sec.get("title", "Table of Contents"),
                        at_start=True,
                        update_on_open=False,
                    )
                else:
                    _tc.insert_caption_list(
                        pkg,
                        label=_CAPTION_KINDS[kind],
                        title=sec.get("title"),
                        at_start=True,
                        update_on_open=False,
                    )
            except WordMcpError as exc:
                # Only reachable with force=true (duplicate TOC/list); the
                # insert refused before mutating, so skipping is clean.
                inserted.append({"kind": kind, "skipped": str(exc)})
                continue
            for el in [c for c in body if id(c) not in before_ids]:
                insert(el)  # addprevious MOVES the element to the cursor
            del before, before_ids
            field_blocks += 1
            inserted.append({"kind": kind})

    # Close the front matter with a SECTION break so numbering can switch.
    # The new sectPr inherits the properties governing the current start.
    governing = _sect_prs(pkg)[0]
    front_sp = copy.deepcopy(governing)
    _sectpr_get_or_add(front_sp, "type").set(qn("w:val"), "nextPage")
    pgnum = _sectpr_get_or_add(front_sp, "pgNumType")
    pgnum.set(qn("w:fmt"), front_fmt)
    pgnum.set(qn("w:start"), "1")
    break_p = etree.Element(qn("w:p"))
    etree.SubElement(break_p, qn("w:pPr")).append(front_sp)
    insert(break_p)

    # Body numbering: the first section AFTER the front matter restarts.
    sects = _sect_prs(pkg)
    body_pgnum = _sectpr_get_or_add(sects[1], "pgNumType")
    body_pgnum.set(qn("w:fmt"), body_fmt)
    body_pgnum.set(qn("w:start"), str(restart))
    pkg.mark_dirty()

    if field_blocks and pkg.has_part("word/settings.xml"):
        _tc.set_update_fields_flag(pkg, True)

    result = {
        "inserted": inserted,
        "front_matter": {
            "section_index": 0,
            "number_format": front_fmt,
            "start_at": 1,
        },
        "body": {
            "section_indices": list(range(1, len(sects))),
            "restart_applied_to_section": 1,
            "number_format": body_fmt,
            "start_at": restart,
        },
        "sections_total": len(sects),
        "update_fields_on_open": bool(field_blocks),
    }
    hf = get_headers_footers(pkg)
    has_page_field = any(
        e["has_page_number_field"] for e in hf["headers"] + hf["footers"]
    )
    if not has_page_field:
        result["note"] = (
            "no PAGE field found in any header/footer; the numbering formats "
            "are set but numbers will not display until page numbers are "
            "added (add_page_numbers)"
        )
    return result
