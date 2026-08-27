"""Style/template transfer: make document A look like reference document B.

Follows the merge-don't-replace algorithm from research topic 12 (modeled on
pandoc's reference-doc mechanism, reimplemented — no GPL code vendored):
styles/theme/fontTable copy wholesale, settings merge by allowlist, numbering
stays, body style references reconciled BY NAME (Heading1 vs '1' vs localized
ids), unmatched styles cloned so nothing silently falls back to Normal.
"""

from __future__ import annotations

import copy

from lxml import etree

from ..core.errors import DocumentNotFound, WordMcpError
from ..core.package import DocxPackage, qn

_SETTINGS_ALLOWLIST = (
    "w:compat",
    "w:defaultTabStop",
    "w:evenAndOddHeaders",
    "w:themeFontLang",
    "w:characterSpacingControl",
    "w:autoHyphenation",
    "w:clrSchemeMapping",
)

_STYLE_REF_TAGS = ("w:pStyle", "w:rStyle", "w:tblStyle")

_STORY_PARTS_PREFIXES = ("word/document.xml", "word/footnotes.xml",
                         "word/endnotes.xml")


def _style_maps(pkg: DocxPackage) -> tuple[dict, dict]:
    """(id -> name, name -> id) from styles.xml."""
    id_to_name: dict[str, str] = {}
    if pkg.has_part("word/styles.xml"):
        for s in pkg.root("word/styles.xml").findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            name_el = s.find(qn("w:name"))
            if sid and name_el is not None:
                id_to_name[sid] = name_el.get(qn("w:val")) or sid
    return id_to_name, {v: k for k, v in id_to_name.items()}


def _story_parts(pkg: DocxPackage) -> list[str]:
    import re

    parts = [p for p in _STORY_PARTS_PREFIXES if pkg.has_part(p)]
    parts += [
        p
        for p in pkg.part_names()
        if re.fullmatch(r"word/(header|footer)\d+\.xml", p)
    ]
    return parts


def _referenced_style_ids(pkg: DocxPackage) -> set[str]:
    refs: set[str] = set()
    for part in _story_parts(pkg):
        root = pkg.root(part)
        for tag in _STYLE_REF_TAGS:
            for el in root.iter(qn(tag)):
                val = el.get(qn("w:val"))
                if val:
                    refs.add(val)
    return refs


def apply_template(
    pkg: DocxPackage,
    reference_path: str,
    *,
    include_page_geometry: bool = True,
) -> dict:
    """Restyle this document to match `reference_path`: styles (with document
    defaults), theme, fonts, key layout settings, and optionally page
    geometry. Content is untouched; style references are remapped by NAME and
    unmatched styles are preserved by cloning (headings never silently lose
    their outline level)."""
    try:
        ref = DocxPackage(reference_path)
    except DocumentNotFound:
        raise
    if not ref.has_part("word/styles.xml"):
        raise WordMcpError("reference document has no styles.xml")

    # 1. Snapshot target's referenced ids and old styles for cloning.
    referenced = _referenced_style_ids(pkg)
    old_id_to_name, _ = _style_maps(pkg)
    old_styles_root = (
        copy.deepcopy(pkg.root("word/styles.xml"))
        if pkg.has_part("word/styles.xml")
        else None
    )

    # 2. Wholesale copies.
    copied = []
    for part in ("word/styles.xml", "word/theme/theme1.xml", "word/fontTable.xml"):
        if ref.has_part(part):
            pkg.set_raw_part(part, ref.raw_part(part))
            copied.append(part)

    # 3. Settings: merge allowlist only.
    merged_settings = []
    if pkg.has_part("word/settings.xml") and ref.has_part("word/settings.xml"):
        tgt = pkg.root("word/settings.xml")
        src = ref.root("word/settings.xml")
        for tag in _SETTINGS_ALLOWLIST:
            src_el = src.find(qn(tag))
            if src_el is None:
                continue
            tgt_el = tgt.find(qn(tag))
            new_el = copy.deepcopy(src_el)
            if tgt_el is not None:
                tgt_el.addprevious(new_el)
                tgt.remove(tgt_el)
            else:
                tgt.append(new_el)
            merged_settings.append(tag)
        pkg.mark_dirty("word/settings.xml")

    # 5. Reconcile references by name; clone anything unmatched.
    new_id_to_name, new_name_to_id = _style_maps(pkg)
    remapped: dict[str, str] = {}
    cloned: list[str] = []
    new_root = pkg.root("word/styles.xml")
    for sid in sorted(referenced):
        if sid in new_id_to_name:
            continue  # id exists in the new set
        name = old_id_to_name.get(sid)
        target_id = new_name_to_id.get(name) if name else None
        if target_id:
            remapped[sid] = target_id
        elif old_styles_root is not None:
            old_def = next(
                (
                    s
                    for s in old_styles_root.findall(qn("w:style"))
                    if s.get(qn("w:styleId")) == sid
                ),
                None,
            )
            if old_def is not None:
                new_root.append(copy.deepcopy(old_def))
                cloned.append(sid)
    if remapped:
        for part in _story_parts(pkg):
            root = pkg.root(part)
            dirty = False
            for tag in _STYLE_REF_TAGS:
                for el in root.iter(qn(tag)):
                    val = el.get(qn("w:val"))
                    if val in remapped:
                        el.set(qn("w:val"), remapped[val])
                        dirty = True
            if dirty:
                pkg.mark_dirty(part)
    pkg.mark_dirty("word/styles.xml")

    # 6. Page geometry from the reference's final section.
    geometry = False
    if include_page_geometry:
        ref_sect = ref.root().find(f"{qn('w:body')}/{qn('w:sectPr')}")
        if ref_sect is not None:
            from .furniture import _sect_prs, _sectpr_get_or_add

            for sp in _sect_prs(pkg):
                for tag in ("w:pgSz", "w:pgMar", "w:cols", "w:docGrid"):
                    src_el = ref_sect.find(qn(tag))
                    if src_el is None:
                        continue
                    dst_el = _sectpr_get_or_add(sp, tag.split(":")[1])
                    for k, v in src_el.attrib.items():
                        dst_el.set(k, v)
            geometry = True
            pkg.mark_dirty()

    # 7. Validate: no dangling references remain.
    final_ids = set(_style_maps(pkg)[0])
    dangling = sorted(_referenced_style_ids(pkg) - final_ids)
    return {
        "template_applied": reference_path,
        "parts_copied": copied,
        "settings_merged": merged_settings,
        "styles_remapped_by_name": remapped,
        "styles_cloned": cloned,
        "page_geometry_copied": geometry,
        "dangling_style_refs": dangling,  # should be empty
    }
