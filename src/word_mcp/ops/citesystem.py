"""Citation-system detection (API review, item 8).

Word documents can carry citations from several mutually incompatible
backends: Word's own citation feature (CITATION fields + a b:Sources store in
customXml), Zotero (ADDIN ZOTERO_ITEM fields), Mendeley (ADDIN CSL_CITATION
fields), EndNote (ADDIN EN.CITE fields), or none of them (plain typed text).
Mixing managed systems in one document creates a split-brain bibliography
where each manager only sees its own half. This module reports which
system(s) a document actually uses, with counts and a split_brain flag, so
citation work starts by matching the document's existing commitment.

Detection reuses the field walker and manager classifier from ops/reffields
and the sources-store locator from ops/bibliography; nothing is re-derived.
"""

from __future__ import annotations

from lxml import etree

from ..core.package import DocxPackage, qn
from .read import body_items, paragraph_text
from .reffields import _STORY_PARTS, _classify_manager, scan_complex_fields

_MANAGERS = ("zotero", "mendeley", "endnote")


def _classify_instr(instr: str) -> tuple[str, str] | None:
    """(system, kind) for a field instruction, covering the reference
    managers (via reffields) plus Word's native CITATION/BIBLIOGRAPHY."""
    managed = _classify_manager(instr)
    if managed is not None:
        return managed
    s = " ".join(instr.split())
    if s.startswith("CITATION"):
        return ("word_native", "citation")
    if s.startswith("BIBLIOGRAPHY"):
        return ("word_native", "bibliography")
    return None


def _count_store_sources(pkg: DocxPackage) -> int:
    """Sources in the Word-native bibliography store, 0 when there is none."""
    from .bibliography import _find_store

    part = _find_store(pkg)
    if part is None:
        return 0
    _B = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
    return len(pkg.root(part).findall(f"{{{_B}}}Source"))


def _plain_text_counts(pkg: DocxPackage) -> dict:
    """Heuristic counts of typed author-year citations in body paragraphs
    (reuses the APA patterns from ops/citecheck)."""
    from .citecheck import _NARRATIVE, _PAREN_CHUNK

    narrative = parenthetical = 0
    for kind, _, el in body_items(pkg):
        if kind != "paragraph":
            continue
        text = paragraph_text(el)
        narrative += sum(1 for _ in _NARRATIVE.finditer(text))
        parenthetical += sum(1 for _ in _PAREN_CHUNK.finditer(text))
    return {"narrative": narrative, "parenthetical": parenthetical,
            "total": narrative + parenthetical}


def detect_citation_system(pkg: DocxPackage) -> dict:
    """Which citation system(s) the document uses, with counts.

    Scans complex fields AND simple (fldSimple) fields across the body,
    footnotes, and endnotes; checks for a Word-native sources store; and
    counts plain-text author-year citations heuristically. More than one
    MANAGED system present sets split_brain=true.
    """
    counts: dict[str, dict[str, int]] = {}

    def bump(system: str, kind: str) -> None:
        entry = counts.setdefault(system, {"citations": 0, "bibliographies": 0})
        entry["citations" if kind == "citation" else "bibliographies"] += 1

    for part in _STORY_PARTS:
        if not pkg.has_part(part):
            continue
        root = pkg.root(part)
        fields, _ = scan_complex_fields(root)
        for rec in fields:
            classified = _classify_instr(rec["instr"])
            if classified is not None:
                bump(*classified)
        for fs in root.iter(qn("w:fldSimple")):
            classified = _classify_instr(fs.get(qn("w:instr"), ""))
            if classified is not None:
                bump(*classified)

    store_sources = _count_store_sources(pkg)
    plain = _plain_text_counts(pkg)

    systems: dict[str, dict] = {}
    if "word_native" in counts or store_sources:
        wn = counts.get("word_native", {"citations": 0, "bibliographies": 0})
        systems["word_native"] = {**wn, "sources_in_store": store_sources}
    for manager in _MANAGERS:
        if manager in counts:
            systems[manager] = counts[manager]

    managed_present = sorted(systems)
    split_brain = len(managed_present) > 1
    managed_citations = sum(s.get("citations", 0) for s in systems.values())

    if split_brain:
        summary = "mixed: " + " + ".join(managed_present)
    elif managed_present:
        summary = managed_present[0]
    elif plain["total"]:
        summary = "plain_text_only"
    else:
        summary = "none"

    result: dict = {
        "system": summary,
        "systems": systems,
        "split_brain": split_brain,
        "plain_text_citations": plain,
        "notes": [
            "plain_text_citations is a heuristic author-year pattern count; "
            "with a managed system present it largely reflects the cached "
            "renderings of the managed fields, not separate typed citations",
        ],
    }
    if split_brain:
        result["warning"] = (
            "more than one managed citation system is present "
            f"({', '.join(managed_present)}); each manager only maintains its "
            "own fields, so the bibliography is split-brain. Pick one system "
            "with the user before adding or converting citations."
        )
    elif managed_present and plain["total"] > managed_citations:
        result["notes"].append(
            "plain-text citation count exceeds the managed field count; the "
            "document may mix typed citations with managed ones (review)"
        )
    return result
