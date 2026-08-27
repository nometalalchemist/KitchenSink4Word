"""Footnote and endnote CRUD.

Conventions (per ECMA-376 and Word's own output; see research/20260827):
- footnotes.xml must carry the two special notes: separator id=-1 and
  continuationSeparator id=0. Real notes use positive ids.
- Note ids do NOT determine displayed numbers; Word numbers by body reference
  order, so inserting anywhere needs no renumbering.
- Superscript marks come entirely from the FootnoteReference/EndnoteReference
  character styles; both are injected if the document lacks them.
- Deleting must remove BOTH the definition and every body reference; a body
  reference without a definition is a repair prompt.
"""

from __future__ import annotations

import re

from lxml import etree

from ..core.errors import AmbiguousTarget, TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap

_KINDS = {
    "footnote": {
        "part": "word/footnotes.xml",
        "root": "w:footnotes",
        "note": "w:footnote",
        "body_ref": "w:footnoteReference",
        "self_ref": "w:footnoteRef",
        "text_style": ("FootnoteText", "footnote text"),
        "ref_style": ("FootnoteReference", "footnote reference"),
        "content_type": (
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.footnotes+xml"
        ),
        "rel_type": (
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/footnotes"
        ),
        "pr": "w:footnotePr",
        "pr_child": "w:footnote",
        "sep_el": "w:separator",
        "cont_el": "w:continuationSeparator",
    },
    "endnote": {
        "part": "word/endnotes.xml",
        "root": "w:endnotes",
        "note": "w:endnote",
        "body_ref": "w:endnoteReference",
        "self_ref": "w:endnoteRef",
        "text_style": ("EndnoteText", "endnote text"),
        "ref_style": ("EndnoteReference", "endnote reference"),
        "content_type": (
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.endnotes+xml"
        ),
        "rel_type": (
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/endnotes"
        ),
        "pr": "w:endnotePr",
        "pr_child": "w:endnote",
        "sep_el": "w:separator",
        "cont_el": "w:continuationSeparator",
    },
}

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


# ------------------------------------------------------------ infrastructure


def _ensure_part(pkg: DocxPackage, kind: str) -> None:
    """Create the notes part with only the two specials, plus content-type
    override, document relationship, and settings declaration."""
    cfg = _KINDS[kind]
    if pkg.has_part(cfg["part"]):
        _ensure_settings_pr(pkg, kind)
        return

    root = etree.Element(qn(cfg["root"]), nsmap={"w": qn("w:x").split("}")[0][1:]})
    for note_type, note_id, inner in (
        ("continuationSeparator", "0", cfg["cont_el"]),
        ("separator", "-1", cfg["sep_el"]),
    ):
        note = etree.SubElement(root, qn(cfg["note"]))
        note.set(qn("w:type"), note_type)
        note.set(qn("w:id"), note_id)
        p = etree.SubElement(note, qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        spacing = etree.SubElement(ppr, qn("w:spacing"))
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        r = etree.SubElement(p, qn("w:r"))
        etree.SubElement(r, qn(inner))
    pkg.set_raw_part(
        cfg["part"],
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
    )

    # [Content_Types].xml override
    ct_root = pkg.root("[Content_Types].xml")
    part_name = "/" + cfg["part"]
    if not any(
        o.get("PartName") == part_name
        for o in ct_root.findall(f"{{{_CT_NS}}}Override")
    ):
        override = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", cfg["content_type"])
        pkg.mark_dirty("[Content_Types].xml")

    # document.xml.rels relationship
    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    existing_ids = {r.get("Id") for r in rels_root}
    if not any(r.get("Type") == cfg["rel_type"] for r in rels_root):
        n = 1
        while f"rId{n}" in existing_ids:
            n += 1
        rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
        rel.set("Id", f"rId{n}")
        rel.set("Type", cfg["rel_type"])
        rel.set("Target", cfg["part"].split("/", 1)[1])
        pkg.mark_dirty(rels_part)

    _ensure_settings_pr(pkg, kind)


def _ensure_settings_pr(pkg: DocxPackage, kind: str) -> None:
    """settings.xml <w:footnotePr>/<w:endnotePr> with the special-note refs."""
    cfg = _KINDS[kind]
    if not pkg.has_part("word/settings.xml"):
        return
    root = pkg.root("word/settings.xml")
    if root.find(qn(cfg["pr"])) is not None:
        return
    pr = etree.Element(qn(cfg["pr"]))
    for special_id in ("-1", "0"):
        child = etree.SubElement(pr, qn(cfg["pr_child"]))
        child.set(qn("w:id"), special_id)
    # Schema position: footnotePr precedes endnotePr, both precede compat/rsids.
    anchor = None
    if kind == "footnote":
        anchor = root.find(qn("w:endnotePr"))
    if anchor is None:
        for tag in ("w:compat", "w:rsids", "w:clrSchemeMapping"):
            anchor = root.find(qn(tag))
            if anchor is not None:
                break
    if anchor is not None:
        anchor.addprevious(pr)
    else:
        root.append(pr)
    pkg.mark_dirty("word/settings.xml")


def _ensure_styles(pkg: DocxPackage, kind: str) -> None:
    cfg = _KINDS[kind]
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    dirty = False

    text_id, text_name = cfg["text_style"]
    if text_id not in have:
        s = etree.SubElement(root, qn("w:style"))
        s.set(qn("w:type"), "paragraph")
        s.set(qn("w:styleId"), text_id)
        etree.SubElement(s, qn("w:name")).set(qn("w:val"), text_name)
        etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
        etree.SubElement(s, qn("w:semiHidden"))
        etree.SubElement(s, qn("w:unhideWhenUsed"))
        ppr = etree.SubElement(s, qn("w:pPr"))
        spacing = etree.SubElement(ppr, qn("w:spacing"))
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), "240")
        spacing.set(qn("w:lineRule"), "auto")
        rpr = etree.SubElement(s, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:sz")).set(qn("w:val"), "20")
        etree.SubElement(rpr, qn("w:szCs")).set(qn("w:val"), "20")
        dirty = True

    ref_id, ref_name = cfg["ref_style"]
    if ref_id not in have:
        s = etree.SubElement(root, qn("w:style"))
        s.set(qn("w:type"), "character")
        s.set(qn("w:styleId"), ref_id)
        etree.SubElement(s, qn("w:name")).set(qn("w:val"), ref_name)
        etree.SubElement(s, qn("w:basedOn")).set(
            qn("w:val"), "DefaultParagraphFont"
        )
        etree.SubElement(s, qn("w:semiHidden"))
        etree.SubElement(s, qn("w:unhideWhenUsed"))
        rpr = etree.SubElement(s, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:vertAlign")).set(qn("w:val"), "superscript")
        dirty = True

    if dirty:
        pkg.mark_dirty("word/styles.xml")


def _next_id(pkg: DocxPackage, kind: str) -> int:
    cfg = _KINDS[kind]
    existing = {
        int(n.get(qn("w:id"), "0"))
        for n in pkg.root(cfg["part"]).findall(qn(cfg["note"]))
    }
    return max(existing | {0}) + 1


def _resolve_note(pkg: DocxPackage, kind: str, *, note_id=None, position=None):
    """Resolve a note by id or by 1-based displayed position."""
    from .read import list_endnotes, list_footnotes

    notes = list_footnotes(pkg) if kind == "footnote" else list_endnotes(pkg)
    if note_id is not None:
        match = [n for n in notes if n["id"] == str(note_id)]
        if not match:
            raise TargetNotFound(f"no {kind} with id {note_id}")
        return match[0]
    if position is not None:
        match = [n for n in notes if n["position"] == position]
        if not match:
            raise TargetNotFound(
                f"no {kind} at position {position} (document has {len(notes)})"
            )
        return match[0]
    raise WordMcpError("give note_id or position")


# ---------------------------------------------------------------- operations


def add_note(
    pkg: DocxPackage,
    kind: str,
    *,
    anchor_text: str,
    note_text: str,
    occurrence: int = 1,
) -> dict:
    """Add a footnote/endnote anchored immediately after `anchor_text` in the
    body. `occurrence` picks which match if the anchor appears more than once."""
    if kind not in _KINDS:
        raise WordMcpError("kind must be footnote or endnote")
    cfg = _KINDS[kind]
    _ensure_part(pkg, kind)
    _ensure_styles(pkg, kind)

    # Locate the anchor.
    target_p = None
    target_end = None
    seen = 0
    for p in pkg.root().iter(qn("w:p")):
        text, _ = _runmap.build_map(p)
        for m in re.finditer(re.escape(anchor_text), text):
            seen += 1
            if seen == occurrence:
                target_p, target_end = p, m.end()
                break
        if target_p is not None:
            break
    if target_p is None:
        if seen:
            raise TargetNotFound(
                f"anchor occurs {seen}x; occurrence {occurrence} not found"
            )
        raise TargetNotFound(f"anchor text not found: {anchor_text!r}")

    note_id = _next_id(pkg, kind)

    # Definition.
    notes_root = pkg.root(cfg["part"])
    note = etree.SubElement(notes_root, qn(cfg["note"]))
    note.set(qn("w:id"), str(note_id))
    _fill_note_content(note, cfg, note_text)
    pkg.mark_dirty(cfg["part"])

    # Body reference run, inserted after the run covering the anchor's end.
    covered = _runmap.split_for_range(target_p, target_end - 1, target_end)
    ref_run = etree.Element(qn("w:r"))
    rpr = etree.SubElement(ref_run, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:rStyle")).set(qn("w:val"), cfg["ref_style"][0])
    ref = etree.SubElement(ref_run, qn(cfg["body_ref"]))
    ref.set(qn("w:id"), str(note_id))
    covered[-1].addnext(ref_run)
    pkg.mark_dirty()

    from .read import list_endnotes, list_footnotes

    notes = list_footnotes(pkg) if kind == "footnote" else list_endnotes(pkg)
    pos = next((n["position"] for n in notes if n["id"] == str(note_id)), None)
    return {"added": kind, "id": note_id, "position": pos}


def _fill_note_content(note: etree._Element, cfg: dict, note_text: str) -> None:
    """Build the note body: FootnoteText paragraph(s), self-ref mark + space in
    the first."""
    for i, para_text in enumerate(note_text.split("\n")):
        p = etree.SubElement(note, qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:pStyle")).set(
            qn("w:val"), cfg["text_style"][0]
        )
        if i == 0:
            ref_run = etree.SubElement(p, qn("w:r"))
            rpr = etree.SubElement(ref_run, qn("w:rPr"))
            etree.SubElement(rpr, qn("w:rStyle")).set(
                qn("w:val"), cfg["ref_style"][0]
            )
            etree.SubElement(ref_run, qn(cfg["self_ref"]))
            sp = etree.SubElement(p, qn("w:r"))
            t = etree.SubElement(sp, qn("w:t"))
            t.text = " "
            _runmap._preserve_space(t)
        if para_text:
            run = etree.SubElement(p, qn("w:r"))
            t = etree.SubElement(run, qn("w:t"))
            t.text = para_text
            _runmap._preserve_space(t)


def edit_note(
    pkg: DocxPackage,
    kind: str,
    *,
    note_id=None,
    position=None,
    new_text: str,
) -> dict:
    """Replace a note's text, keeping its mark and styles."""
    cfg = _KINDS[kind]
    info = _resolve_note(pkg, kind, note_id=note_id, position=position)
    root = pkg.root(cfg["part"])
    note = next(
        n
        for n in root.findall(qn(cfg["note"]))
        if n.get(qn("w:id")) == info["id"]
    )
    for p in note.findall(qn("w:p")):
        note.remove(p)
    _fill_note_content(note, cfg, new_text)
    pkg.mark_dirty(cfg["part"])
    return {"edited": kind, "id": info["id"], "position": info["position"]}


def delete_note(
    pkg: DocxPackage, kind: str, *, note_id=None, position=None
) -> dict:
    """Delete a note: definition AND every body reference (both mandatory —
    an orphaned body reference corrupts the document)."""
    cfg = _KINDS[kind]
    info = _resolve_note(pkg, kind, note_id=note_id, position=position)
    if int(info["id"]) < 1:
        raise WordMcpError("refusing to touch separator notes")

    root = pkg.root(cfg["part"])
    note = next(
        n
        for n in root.findall(qn(cfg["note"]))
        if n.get(qn("w:id")) == info["id"]
    )
    root.remove(note)
    pkg.mark_dirty(cfg["part"])

    removed_refs = 0
    body_root = pkg.root()
    for ref in body_root.iter(qn(cfg["body_ref"])):
        if ref.get(qn("w:id")) == info["id"]:
            run = ref.getparent()
            container = run.getparent()
            container.remove(run)
            if (
                etree.QName(container).localname == "hyperlink"
                and len(container) == 0
            ):
                container.getparent().remove(container)
            removed_refs += 1
            break  # iterator invalidated; ids are unique anyway
    pkg.mark_dirty()
    return {
        "deleted": kind,
        "id": info["id"],
        "body_references_removed": removed_refs,
    }


def convert_notes(
    pkg: DocxPackage,
    direction: str,
    *,
    note_id=None,
    position=None,
) -> dict:
    """Convert footnotes to endnotes or vice versa. Give note_id/position for
    one note, neither for ALL notes. Content, order, and formatting carry
    over; Word renumbers automatically."""
    if direction == "footnotes_to_endnotes":
        src_kind, dst_kind = "footnote", "endnote"
    elif direction == "endnotes_to_footnotes":
        src_kind, dst_kind = "endnote", "footnote"
    else:
        raise WordMcpError(
            "direction must be footnotes_to_endnotes or endnotes_to_footnotes"
        )
    src_cfg, dst_cfg = _KINDS[src_kind], _KINDS[dst_kind]
    if not pkg.has_part(src_cfg["part"]):
        raise TargetNotFound(f"document has no {src_kind}s")

    if note_id is not None or position is not None:
        targets = [_resolve_note(pkg, src_kind, note_id=note_id, position=position)]
    else:
        from .read import list_endnotes, list_footnotes

        targets = (
            list_footnotes(pkg) if src_kind == "footnote" else list_endnotes(pkg)
        )
        if not targets:
            raise TargetNotFound(f"document has no {src_kind}s to convert")

    _ensure_part(pkg, dst_kind)
    _ensure_styles(pkg, dst_kind)

    src_root = pkg.root(src_cfg["part"])
    dst_root = pkg.root(dst_cfg["part"])
    body_root = pkg.root()
    converted = []
    for info in targets:
        old_id = info["id"]
        new_id = _next_id(pkg, dst_kind)
        note = next(
            n
            for n in src_root.findall(qn(src_cfg["note"]))
            if n.get(qn("w:id")) == old_id
        )
        # Move the definition, converting tag, styles, and self-reference.
        src_root.remove(note)
        note.tag = qn(dst_cfg["note"])
        note.set(qn("w:id"), str(new_id))
        for ps in note.iter(qn("w:pStyle")):
            if ps.get(qn("w:val")) == src_cfg["text_style"][0]:
                ps.set(qn("w:val"), dst_cfg["text_style"][0])
        for rs in note.iter(qn("w:rStyle")):
            if rs.get(qn("w:val")) == src_cfg["ref_style"][0]:
                rs.set(qn("w:val"), dst_cfg["ref_style"][0])
        for ref in note.iter(qn(src_cfg["self_ref"])):
            ref.tag = qn(dst_cfg["self_ref"])
        dst_root.append(note)

        # Retag the body reference.
        for body_ref in body_root.iter(qn(src_cfg["body_ref"])):
            if body_ref.get(qn("w:id")) == old_id:
                body_ref.tag = qn(dst_cfg["body_ref"])
                body_ref.set(qn("w:id"), str(new_id))
                run = body_ref.getparent()
                rpr = run.find(qn("w:rPr"))
                if rpr is not None:
                    rs = rpr.find(qn("w:rStyle"))
                    if rs is not None and rs.get(qn("w:val")) == src_cfg["ref_style"][0]:
                        rs.set(qn("w:val"), dst_cfg["ref_style"][0])
                break
        converted.append({"from_id": old_id, "to_id": str(new_id)})

    pkg.mark_dirty(src_cfg["part"])
    pkg.mark_dirty(dst_cfg["part"])
    pkg.mark_dirty()
    return {
        "converted": len(converted),
        "direction": direction,
        "mappings": converted,
    }


def validate_notes(pkg: DocxPackage) -> dict:
    """Two-direction integrity check for both note kinds.

    `ok` = no corruption (missing definitions / duplicate references).
    `needs_cleanup` = orphan definitions exist (harmless to Word but garbage;
    run cleanup via purge_orphans / the cleanup_orphan_notes tool)."""
    out = {}
    for kind, cfg in _KINDS.items():
        refs = [
            r.get(qn("w:id"))
            for r in pkg.root().iter(qn(cfg["body_ref"]))
        ]
        defs = set()
        if pkg.has_part(cfg["part"]):
            defs = {
                n.get(qn("w:id"))
                for n in pkg.root(cfg["part"]).findall(qn(cfg["note"]))
                if n.get(qn("w:type"))
                not in ("separator", "continuationSeparator", "continuationNotice")
            }
        missing = [r for r in refs if r not in defs]  # corrupting
        orphans = sorted(defs - set(refs))  # dead weight only
        dup_refs = sorted({r for r in refs if refs.count(r) > 1})
        out[kind + "s"] = {
            "references": len(refs),
            "definitions": len(defs),
            "missing_definitions": missing,
            "orphan_definitions": orphans,
            "duplicate_references": dup_refs,
            "ok": not missing and not dup_refs,
            "needs_cleanup": bool(orphans),
        }
    return out


def purge_orphans(pkg: DocxPackage) -> dict:
    """Remove note definitions that no body reference points to. Called
    automatically by content-deleting operations; also exposed as a tool."""
    removed: dict[str, list[str]] = {}
    for kind, cfg in _KINDS.items():
        if not pkg.has_part(cfg["part"]):
            continue
        refs = {
            r.get(qn("w:id")) for r in pkg.root().iter(qn(cfg["body_ref"]))
        }
        root = pkg.root(cfg["part"])
        doomed = [
            n
            for n in root.findall(qn(cfg["note"]))
            if n.get(qn("w:type"))
            not in ("separator", "continuationSeparator", "continuationNotice")
            and n.get(qn("w:id")) not in refs
        ]
        if doomed:
            for n in doomed:
                root.remove(n)
            pkg.mark_dirty(cfg["part"])
            removed[kind + "s"] = [n.get(qn("w:id")) for n in doomed]
    return {"purged": removed}
