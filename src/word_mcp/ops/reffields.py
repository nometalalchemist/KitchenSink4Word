"""Reference-manager field inventory: Zotero, EndNote, Mendeley.

Citation managers store their live data inside complex fields (ADDIN fields):
the instruction text carries the manager's payload (CSL JSON, EN.CITE XML) and
the runs between the separate and end markers carry the rendered citation.
Editing tools that split, delete, or reorder runs can orphan one half of a
fldChar begin/separate/end triple, which silently disconnects the citation
from the manager. This module finds every such field and reports whether its
markers are still properly paired — the post-edit "did my edit break Zotero"
check.

Also home to scan_complex_fields(), the document-order field walker shared
with ops/integrity.py.
"""

from __future__ import annotations

from lxml import etree

from ..core.package import DocxPackage, qn

_STORY_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")


def scan_complex_fields(root: etree._Element) -> tuple[list[dict], list[dict]]:
    """Walk a story part in document order and pair up complex fields.

    Returns (fields, orphans). Each field dict: instr (joined instruction
    text), cached (joined rendered-result text), separated, closed, begin_el.
    Nested fields (a begin inside another field's result) are tracked on a
    stack, so their instruction/result text is attributed to the inner field.
    Orphans are separate/end markers that appear with no field open — each is
    {type, el} and always indicates damage.
    """
    fields: list[dict] = []
    stack: list[dict] = []
    orphans: list[dict] = []
    for el in root.iter():
        local = etree.QName(el).localname
        if local == "fldChar":
            ftype = el.get(qn("w:fldCharType"))
            if ftype == "begin":
                rec = {
                    "instr_parts": [],
                    "cached_parts": [],
                    "separated": False,
                    "closed": False,
                    "begin_el": el,
                }
                stack.append(rec)
                fields.append(rec)
            elif ftype == "separate":
                if stack:
                    stack[-1]["separated"] = True
                else:
                    orphans.append({"type": "separate", "el": el})
            elif ftype == "end":
                if stack:
                    stack.pop()["closed"] = True
                else:
                    orphans.append({"type": "end", "el": el})
        elif local == "instrText":
            if stack and not stack[-1]["separated"]:
                stack[-1]["instr_parts"].append(el.text or "")
        elif local == "t":
            if stack and stack[-1]["separated"]:
                stack[-1]["cached_parts"].append(el.text or "")
    for rec in fields:
        rec["instr"] = "".join(rec["instr_parts"]).strip()
        rec["cached"] = "".join(rec["cached_parts"])
    return fields, orphans


def _classify_manager(instr: str) -> tuple[str, str] | None:
    """(manager, kind) for a reference-manager ADDIN instruction, else None.

    Order matters: Zotero instructions contain the generic CSL_CITATION token
    that also identifies Mendeley, so Zotero is tested first.
    """
    s = " ".join(instr.split())
    if not s.startswith("ADDIN"):
        return None
    if "ZOTERO_ITEM" in s:
        return ("zotero", "citation")
    if "ZOTERO_BIBL" in s:
        return ("zotero", "bibliography")
    if "EN.CITE" in s:
        return ("endnote", "citation")
    if "EN.REFLIST" in s:  # EndNote's bibliography instruction
        return ("endnote", "bibliography")
    if "Mendeley Bibliography" in s or "CSL_BIBLIOGRAPHY" in s:
        return ("mendeley", "bibliography")
    if "CSL_CITATION" in s:
        return ("mendeley", "citation")
    return None


def _body_paragraph_index_map(
    pkg: DocxPackage,
) -> tuple[dict[int, int], list]:
    """(id(paragraph element) -> body index, keepalive list).

    The caller must hold the keepalive list while using the map: lxml element
    proxies are recreated per access unless a reference stays alive, and only
    a live proxy has a stable id()."""
    from .read import body_items

    els = [
        (el, idx) for kind, idx, el in body_items(pkg) if kind == "paragraph"
    ]
    return {id(el): idx for el, idx in els}, [el for el, _ in els]


def _containing_paragraph(el: etree._Element) -> etree._Element | None:
    node = el
    while node is not None and etree.QName(node).localname != "p":
        node = node.getparent()
    return node


def _locate(
    el: etree._Element, part: str, index_map: dict[int, int]
) -> dict:
    """Location info for a field's begin marker. paragraph_index is the
    body-level index; None when the paragraph is nested (table cell, SDT)
    or lives in a notes part."""
    loc: dict = {"part": part, "paragraph_index": None}
    p = _containing_paragraph(el)
    if p is not None and part == "word/document.xml":
        loc["paragraph_index"] = index_map.get(id(p))
    return loc


def list_reference_fields(pkg: DocxPackage) -> dict:
    """Inventory every Zotero/EndNote/Mendeley field in the document body,
    footnotes, and endnotes. Per field: manager, kind (citation or
    bibliography), location, the cached rendered text, and whether its field
    markers are intact. Broken pairs — a begin without an end, or stray
    separate/end markers — are reported separately and loudly, because they
    disconnect the citation from the reference manager."""
    index_map, _keepalive = _body_paragraph_index_map(pkg)
    out_fields: list[dict] = []
    broken: list[dict] = []
    orphan_chars: list[dict] = []
    unclosed_any = 0
    parts_scanned: list[str] = []

    for part in _STORY_PARTS:
        if not pkg.has_part(part):
            continue
        parts_scanned.append(part)
        fields, orphans = scan_complex_fields(pkg.root(part))
        for o in orphans:
            loc = _locate(o["el"], part, index_map)
            orphan_chars.append({"marker": o["type"], **loc})
        for rec in fields:
            if not rec["closed"]:
                unclosed_any += 1
            managed = _classify_manager(rec["instr"])
            if managed is None:
                continue
            manager, kind = managed
            loc = _locate(rec["begin_el"], part, index_map)
            entry = {
                "manager": manager,
                "kind": kind,
                **loc,
                "cached_text": rec["cached"],
                "intact": rec["closed"],
                "has_cached_result": rec["separated"],
            }
            out_fields.append(entry)
            if not rec["closed"]:
                broken.append(
                    {
                        "problem": "field begin without matching end",
                        "manager": manager,
                        "kind": kind,
                        **loc,
                    }
                )

    by_manager: dict[str, int] = {}
    for f in out_fields:
        by_manager[f["manager"]] = by_manager.get(f["manager"], 0) + 1
    return {
        "fields": out_fields,
        "total": len(out_fields),
        "by_manager": by_manager,
        "broken": broken,
        "orphan_field_chars": orphan_chars,
        "unclosed_fields_any_type": unclosed_any,
        "parts_scanned": parts_scanned,
    }


def check_reference_field_integrity(pkg: DocxPackage) -> dict:
    """Quick post-edit health check for reference-manager fields: counts by
    manager and kind, plus an ok flag. ok is False when any manager field has
    broken markers, when stray field markers exist anywhere, or when ANY
    complex field is left unclosed (unclosed fields corrupt everything after
    them, citations included). Run this after editing a document that
    contains Zotero/EndNote/Mendeley citations."""
    inv = list_reference_fields(pkg)
    citations = sum(1 for f in inv["fields"] if f["kind"] == "citation")
    bibliographies = sum(
        1 for f in inv["fields"] if f["kind"] == "bibliography"
    )
    ok = (
        not inv["broken"]
        and not inv["orphan_field_chars"]
        and inv["unclosed_fields_any_type"] == 0
    )
    result = {
        "ok": ok,
        "total_reference_fields": inv["total"],
        "citations": citations,
        "bibliographies": bibliographies,
        "by_manager": inv["by_manager"],
        "broken": inv["broken"],
        "orphan_field_chars": inv["orphan_field_chars"],
        "unclosed_fields_any_type": inv["unclosed_fields_any_type"],
    }
    if not ok:
        result["warning"] = (
            "field structure is damaged; citations may have lost their "
            "reference-manager link — restore from backup or repair in Word"
        )
    return result
