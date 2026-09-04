"""Whole-document positional insertion: insert one .docx's entire body into
another at an exact position, with full resource reconciliation.

Built for the chapter-merge workflow (2026-08-28 merge test): merging a
100+-paragraph chapter with tables into a dissertation at a specific position
was previously only possible with a hand-written lxml script, because
com_merge_documents concatenates end-to-end and insert_paragraphs cannot carry
tables or large payloads. This module is that script, made safe.

What "safe" means here — a naive element transplant corrupts, so every
resource class the transplanted content references is reconciled:

- STYLES: matched to target styles BY NAME (the apply_template model, ops/
  template.py). Matching names remap styleIds to the target's id; the target's
  formatting governs (merge-don't-replace). Unmatched styles are cloned into
  the target with their basedOn/link/next dependency chains, under fresh ids
  when the id is taken.
- NUMBERING: every source list instance (numId) gets its own freshly-cloned
  abstractNum + num pair in the target with non-colliding ids, so numbering
  restart semantics are preserved exactly and no source list ever attaches to
  a target list.
- FOOTNOTES/ENDNOTES: definitions are copied with new non-colliding ids and
  the transplanted references retagged; note content goes through the same
  style/numbering/relationship reconciliation as body content.
- IMAGES: media parts copied under fresh names, new relationship ids, content
  types ensured, docPr ids renumbered for uniqueness.
- CHARTS: the chart part and its private subtree (embedded workbook, colors/
  style parts) are copied part-for-part with rewritten relationship targets
  and content-type overrides (the ops/charts.py plumbing pattern).
- HYPERLINKS and other external relationships: re-registered with fresh rIds.
- BOOKMARKS: ids remapped to fresh values; name collisions renamed with a
  suffix (reported), and w:anchor hyperlinks / REF-family fields inside the
  inserted content retargeted to the renamed bookmarks.
- COMMENTS: comment transplant is OUT OF SCOPE. Comment references and range
  markers in the source content are stripped cleanly and the count reported.
- TRACKED CHANGES: body-level revision markup (w:ins/w:del/...) is carried
  as-is.
- SECTIONS: the source's trailing sectPr is NEVER carried — body content
  only; headers, footers, and page setup stay the target's. Mid-content
  section breaks (paragraph-embedded sectPr) are stripped and reported, since
  carrying them would import the source's page furniture.

Anything that cannot be carried safely (OLE objects, ActiveX controls,
subdocuments, altChunks, SmartArt/unknown embedded parts) REFUSES the whole
insertion, naming the blocking content — nothing is ever half-applied.

insert_document's `formatting` parameter mirrors Word's paste options:
"source" (default — direct formatting preserved, Word InsertFile behavior),
"merge" (semantic emphasis kept, font/size/color/spacing/indent direct
overrides stripped so the target's styles size the carried text), and
"destination" (all direct rPr/pPr formatting stripped except structural
properties — numPr, outlineLvl, tab stops — so the target's styles govern
rendering entirely). Stripping happens on the COPIED elements only; the
source file is never modified.

copy_table() transplants a single top-level table through the exact same
reconciliation pipeline (styles by name, numbering, rels/images, notes,
bookmarks) scoped to that one element, with the same positioning contract
and refusal classes as insert_document.
"""

from __future__ import annotations

import copy
import posixpath
import re

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    TargetNotFound,
    UnsupportedStructure,
    ValidationFailed,
    WordMcpError,
)
from ..core.package import DocxPackage, qn
from . import lists as _lists
from . import media as _media
from . import notes as _notes
from .read import paragraph_text
from .template import _style_maps

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

_REL_IMAGE = _R_NS + "/image"
_REL_HYPERLINK = _R_NS + "/hyperlink"
_REL_CHART = _R_NS + "/chart"
_REL_CHARTEX = "http://schemas.microsoft.com/office/2014/relationships/chartEx"

# Internal relationship types we know how to transplant. Everything else that
# the inserted content references via r:id is a refusal (conservative mode).
_SUBTREE_REL_TYPES = {_REL_CHART, _REL_CHARTEX}

# Elements whose presence in the source content blocks the whole insertion.
_BLOCKED_ELEMENTS = {
    qn("w:object"): "embedded OLE object (w:object)",
    qn("w:control"): "ActiveX control (w:control)",
    qn("w:subDoc"): "subdocument reference (w:subDoc)",
    qn("w:altChunk"): "altChunk (embedded alternative-format content)",
    qn("w:movie"): "embedded movie (w:movie)",
    "{urn:schemas-microsoft-com:office:office}OLEObject": (
        "embedded OLE object (o:OLEObject)"
    ),
}

# Tags carrying style-id references, in content AND in style/numbering defs.
_STYLE_REF_TAGS = (
    "w:pStyle",
    "w:rStyle",
    "w:tblStyle",
    "w:basedOn",
    "w:link",
    "w:next",
    "w:styleLink",
    "w:numStyleLink",
)

# story -> (destination part, its rels part) — origin and destination match.
_STORY_PARTS = {
    "document": ("word/document.xml", "word/_rels/document.xml.rels"),
    "footnote": ("word/footnotes.xml", "word/_rels/footnotes.xml.rels"),
    "endnote": ("word/endnotes.xml", "word/_rels/endnotes.xml.rels"),
}


# ------------------------------------------------------------- small helpers


def _localname(el: etree._Element) -> str:
    return etree.QName(el).localname


def _rels_part_for(part: str) -> str:
    folder, name = part.rsplit("/", 1)
    return f"{folder}/_rels/{name}.rels"


def _resolve_rel_target(rels_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = rels_part.rsplit("_rels/", 1)[0].rstrip("/")
    return posixpath.normpath(posixpath.join(base, target))


def _ensure_rels_part(pkg: DocxPackage, name: str) -> None:
    if pkg.has_part(name):
        return
    root = etree.Element(f"{{{_REL_NS}}}Relationships", nsmap={None: _REL_NS})
    pkg.set_raw_part(
        name,
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
    )


def _iter_rid_attrs(el: etree._Element):
    """Yield (node, attr_key, rid) for every r:-namespace attribute below
    (and on) el."""
    prefix = "{" + _R_NS + "}"
    for node in el.iter():
        for key, val in node.attrib.items():
            if key.startswith(prefix):
                yield node, key, val


def _free_part_name(pkg: DocxPackage, like: str) -> str:
    """A part name in the same folder / same shape as `like` that does not
    exist in the target: trailing digits before the extension are replaced by
    the first free number."""
    folder, base = like.rsplit("/", 1)
    m = re.fullmatch(r"(.*?)(\d*)(\.[^./]+)?", base)
    stem, _, ext = m.group(1), m.group(2), m.group(3) or ""
    n = 1
    while True:
        cand = f"{folder}/{stem}{n}{ext}"
        if not pkg.has_part(cand):
            return cand
        n += 1


def _content_type_of(src: DocxPackage, part: str) -> tuple[str, str] | None:
    """('override'|'default', content_type) for a source part, from its
    [Content_Types].xml."""
    ct_root = src.root("[Content_Types].xml")
    part_name = "/" + part
    for o in ct_root.findall(f"{{{_CT_NS}}}Override"):
        if o.get("PartName") == part_name:
            return "override", o.get("ContentType")
    ext = part.rsplit(".", 1)[1].lower() if "." in part.rsplit("/", 1)[1] else ""
    for d in ct_root.findall(f"{{{_CT_NS}}}Default"):
        if (d.get("Extension") or "").lower() == ext:
            return "default", d.get("ContentType")
    return None


def _ensure_content_type(
    pkg: DocxPackage, src: DocxPackage, src_part: str, new_part: str
) -> None:
    """Replicate the source part's content-type declaration for its copy."""
    found = _content_type_of(src, src_part)
    ct_root = pkg.root("[Content_Types].xml")
    if found is None:
        # No declaration in the source either; fall back to a media guess.
        ext = "." + new_part.rsplit(".", 1)[1].lower() if "." in new_part else ""
        ctype = _media._EXT_TYPES.get(ext, "application/octet-stream")
        kind = "default"
    else:
        kind, ctype = found
    if kind == "override":
        part_name = "/" + new_part
        if not any(
            o.get("PartName") == part_name
            for o in ct_root.findall(f"{{{_CT_NS}}}Override")
        ):
            o = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
            o.set("PartName", part_name)
            o.set("ContentType", ctype)
            pkg.mark_dirty("[Content_Types].xml")
    else:
        ext_name = new_part.rsplit(".", 1)[1].lower()
        if not any(
            (d.get("Extension") or "").lower() == ext_name
            for d in ct_root.findall(f"{{{_CT_NS}}}Default")
        ):
            d = etree.SubElement(ct_root, f"{{{_CT_NS}}}Default")
            d.set("Extension", ext_name)
            d.set("ContentType", ctype)
            pkg.mark_dirty("[Content_Types].xml")


def _ensure_styles_part(pkg: DocxPackage) -> None:
    """Create a minimal word/styles.xml (plus content type and relationship)
    for the rare target that lacks one, so style cloning has a home."""
    part = "word/styles.xml"
    if pkg.has_part(part):
        return
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = etree.Element(qn("w:styles"), nsmap={"w": w_ns})
    pkg.set_raw_part(
        part,
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
    )
    ct_root = pkg.root("[Content_Types].xml")
    if not any(
        o.get("PartName") == "/" + part
        for o in ct_root.findall(f"{{{_CT_NS}}}Override")
    ):
        o = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        o.set("PartName", "/" + part)
        o.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.styles+xml",
        )
        pkg.mark_dirty("[Content_Types].xml")
    rels_part = "word/_rels/document.xml.rels"
    if pkg.has_part(rels_part):
        rels_root = pkg.root(rels_part)
        rel_type = _R_NS + "/styles"
        if not any(r.get("Type") == rel_type for r in rels_root):
            existing = {r.get("Id") for r in rels_root}
            n = 1
            while f"rId{n}" in existing:
                n += 1
            rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
            rel.set("Id", f"rId{n}")
            rel.set("Type", rel_type)
            rel.set("Target", "styles.xml")
            pkg.mark_dirty(rels_part)


# ------------------------------------------------------------- position logic


def _body_blocks(pkg: DocxPackage) -> list[etree._Element]:
    """Body items — paragraphs and tables as ONE document-order sequence."""
    return [
        c
        for c in pkg.body()
        if _localname(c) in ("p", "tbl")
    ]


def _resolve_position(
    pkg: DocxPackage,
    after_index: int | None,
    after_anchor: str | None,
    at_end: bool,
    before_first: bool = False,
) -> tuple[str, etree._Element | None, int]:
    """(mode, reference element, body-item index where inserted content
    starts). mode: 'append' (before the trailing sectPr), 'after' (after
    the reference element), or 'before' (before the reference element,
    used for insertion ahead of the first body item)."""
    blocks = _body_blocks(pkg)
    if before_first:
        if not blocks:
            return "append", None, 0
        return "before", blocks[0], 0
    if at_end:
        return "append", None, len(blocks)
    if after_index is not None:
        if not 0 <= after_index < len(blocks):
            raise TargetNotFound(
                f"after_index {after_index} out of range: the target has "
                f"{len(blocks)} body items (paragraphs + tables in document "
                f"order, valid indices 0-{len(blocks) - 1})"
            )
        return "after", blocks[after_index], after_index + 1
    # after_anchor: exact paragraph-text match, refused loudly on ambiguity —
    # heading text recurring in body prose must never resolve first-match.
    anchor = after_anchor.strip()
    matches = [
        (i, el)
        for i, el in enumerate(blocks)
        if _localname(el) == "p" and paragraph_text(el).strip() == anchor
    ]
    if not matches:
        raise TargetNotFound(
            f"no body paragraph's full text exactly matches {anchor!r} "
            "(after_anchor compares whole-paragraph text; use after_index "
            "for structural positions)"
        )
    if len(matches) > 1:
        locations = [
            {
                "body_item_index": i,
                "text": paragraph_text(el).strip()[:120],
            }
            for i, el in matches
        ]
        raise AmbiguousTarget(
            f"anchor text matches {len(matches)} paragraphs — refusing to "
            f"guess. Matches (body item indices): {locations}. Use "
            "after_index with the intended index instead."
        )
    i, el = matches[0]
    return "after", el, i + 1


# --------------------------------------------------------- style / numbering


class _StyleResolver:
    """By-name style reconciliation (the apply_template model): matching
    names remap ids to the target's; unmatched styles are cloned with their
    dependency chains under fresh ids."""

    def __init__(self, src: DocxPackage, pkg: DocxPackage):
        self.pkg = pkg
        self.src_id2name, _ = _style_maps(src)
        self.tgt_id2name, self.tgt_name2id = _style_maps(pkg)
        self.src_defs: dict[str, etree._Element] = {}
        if src.has_part("word/styles.xml"):
            for s in src.root("word/styles.xml").findall(qn("w:style")):
                sid = s.get(qn("w:styleId"))
                if sid:
                    self.src_defs[sid] = s
        self.remap: dict[str, str] = {}  # src id -> different target id
        self.matched: list[dict] = []
        self.cloned: list[dict] = []
        self.cloned_defs: list[etree._Element] = []
        self.unresolved: list[str] = []
        self._done: set[str] = set()

    def resolve(self, sid: str) -> None:
        if not sid or sid in self._done:
            return
        self._done.add(sid)
        name = self.src_id2name.get(sid)
        if name is None:
            # No definition in the source. If the target defines the id the
            # reference lands on that style; otherwise it dangles exactly as
            # it dangled in the source (Word falls back to Normal). Reported.
            if sid not in self.tgt_id2name:
                self.unresolved.append(sid)
            return
        tgt_id = self.tgt_name2id.get(name)
        if tgt_id is not None:
            # Name match: the target's definition (formatting) governs.
            if tgt_id != sid:
                self.remap[sid] = tgt_id
            self.matched.append(
                {"name": name, "source_id": sid, "target_id": tgt_id}
            )
            return
        # Clone, keeping the source id when free.
        new_id = sid
        n = 1
        while new_id in self.tgt_id2name:
            new_id = f"{sid}Ins{n}"
            n += 1
        d = copy.deepcopy(self.src_defs[sid])
        d.set(qn("w:styleId"), new_id)
        _ensure_styles_part(self.pkg)
        root = self.pkg.root("word/styles.xml")
        root.append(d)
        self.pkg.mark_dirty("word/styles.xml")
        self.tgt_id2name[new_id] = name
        self.tgt_name2id[name] = new_id
        if new_id != sid:
            self.remap[sid] = new_id
        self.cloned.append({"id": new_id, "name": name, "source_id": sid})
        self.cloned_defs.append(d)
        for dep_tag in ("w:basedOn", "w:link", "w:next"):
            dep = d.find(qn(dep_tag))
            if dep is not None:
                self.resolve(dep.get(qn("w:val")))


class _NumberingResolver:
    """Clone source abstractNum/num pairs with fresh non-colliding ids.
    Every source numId gets its OWN target instance (never merged with a
    target list), so restart semantics and lvlOverride/startOverride are
    preserved exactly."""

    def __init__(self, src: DocxPackage, pkg: DocxPackage):
        self.src = src
        self.pkg = pkg
        self.map: dict[str, str] = {}  # src numId -> tgt numId
        self.abs_map: dict[str, str] = {}
        self.cloned_abstracts: list[etree._Element] = []
        self.unresolved: list[str] = []
        self._done: set[str] = set()

    def _tgt_root(self) -> etree._Element:
        _lists._ensure_numbering_part(self.pkg)
        return self.pkg.root("word/numbering.xml")

    def resolve(self, num_id: str) -> None:
        if not num_id or num_id in self._done:
            return
        self._done.add(num_id)
        if num_id == "0":  # numId 0 = "numbering removed" marker; keep as-is
            self.map[num_id] = num_id
            return
        if not self.src.has_part("word/numbering.xml"):
            self.unresolved.append(num_id)
            return
        src_root = self.src.root("word/numbering.xml")
        num = next(
            (
                n
                for n in src_root.findall(qn("w:num"))
                if n.get(qn("w:numId")) == num_id
            ),
            None,
        )
        if num is None:
            self.unresolved.append(num_id)
            return
        tgt_root = self._tgt_root()
        abs_ref = num.find(qn("w:abstractNumId"))
        abs_id = abs_ref.get(qn("w:val")) if abs_ref is not None else None
        new_abs_id = self.abs_map.get(abs_id)
        if abs_id is not None and new_abs_id is None:
            abstract = next(
                (
                    a
                    for a in src_root.findall(qn("w:abstractNum"))
                    if a.get(qn("w:abstractNumId")) == abs_id
                ),
                None,
            )
            if abstract is not None:
                existing_abs = [
                    int(a.get(qn("w:abstractNumId"), "0") or 0)
                    for a in tgt_root.findall(qn("w:abstractNum"))
                ]
                new_abs_id = str(max(existing_abs, default=-1) + 1)
                clone = copy.deepcopy(abstract)
                clone.set(qn("w:abstractNumId"), new_abs_id)
                nums = tgt_root.findall(qn("w:num"))
                if nums:  # schema order: abstractNum before num
                    nums[0].addprevious(clone)
                else:
                    tgt_root.append(clone)
                self.abs_map[abs_id] = new_abs_id
                self.cloned_abstracts.append(clone)
        existing_nums = [
            int(n.get(qn("w:numId"), "0") or 0)
            for n in tgt_root.findall(qn("w:num"))
        ]
        new_num_id = str(max(existing_nums, default=0) + 1)
        num_clone = copy.deepcopy(num)
        num_clone.set(qn("w:numId"), new_num_id)
        clone_abs_ref = num_clone.find(qn("w:abstractNumId"))
        if clone_abs_ref is not None and new_abs_id is not None:
            clone_abs_ref.set(qn("w:val"), new_abs_id)
        tgt_root.append(num_clone)
        self.pkg.mark_dirty("word/numbering.xml")
        self.map[num_id] = new_num_id


def _collect_style_refs(elements) -> set[str]:
    refs: set[str] = set()
    for el in elements:
        for tag in _STYLE_REF_TAGS:
            for node in el.iter(qn(tag)):
                val = node.get(qn("w:val"))
                if val:
                    refs.add(val)
    return refs


def _collect_num_refs(elements) -> set[str]:
    refs: set[str] = set()
    for el in elements:
        for node in el.iter(qn("w:numId")):
            val = node.get(qn("w:val"))
            if val:
                refs.add(val)
    return refs


def _apply_style_remap(elements, remap: dict[str, str]) -> None:
    if not remap:
        return
    for el in elements:
        for tag in _STYLE_REF_TAGS:
            for node in el.iter(qn(tag)):
                val = node.get(qn("w:val"))
                if val in remap:
                    node.set(qn("w:val"), remap[val])


def _apply_num_remap(elements, num_map: dict[str, str], unresolved: set[str]):
    """Remap numIds; numPr pointing at a numbering instance the SOURCE never
    defined is stripped (it rendered plain in the source too) and counted."""
    stripped = 0
    for el in elements:
        for node in list(el.iter(qn("w:numId"))):
            val = node.get(qn("w:val"))
            if val in num_map:
                node.set(qn("w:val"), num_map[val])
            elif val in unresolved:
                numpr = node.getparent()  # w:numPr
                if numpr is not None and _localname(numpr) == "numPr":
                    numpr.getparent().remove(numpr)
                    stripped += 1
    return stripped


# ------------------------------- document-defaults reconciliation (source)
#
# The half-single-spaced-dissertation bug (field test, 2026-09-03): with
# formatting="source", paragraphs that carried NO explicit line_spacing or
# space_after relied on their SOURCE file's docDefaults for those values.
# After transplant they resolved against the TARGET's docDefaults instead,
# silently changing the rendered spacing. formatting="source" promises the
# source look, so for every tracked property where the two files'
# docDefaults differ, the source-effective value is baked as an explicit
# property onto carried paragraphs/runs that inherited it (direct values
# and style-chain values are left alone: direct already wins, and style
# values follow the documented by-name reconciliation contract).

# Attribute-level tracked properties (OOXML merges these attribute-wise).
_DD_PPR_ATTRS: dict[str, tuple[str, ...]] = {
    "spacing": (
        "after", "before", "line", "lineRule",
        "afterAutospacing", "beforeAutospacing",
    ),
    "ind": ("left", "start", "right", "end", "firstLine", "hanging"),
    "jc": ("val",),
}
_DD_RPR_ATTRS: dict[str, tuple[str, ...]] = {
    "rFonts": ("ascii", "hAnsi", "eastAsia", "cs"),
    "sz": ("val",),
    "szCs": ("val",),
}

# Word's built-in value when NEITHER file's docDefaults define an attribute
# the other file does define. Only pPr attributes with well-defined
# built-ins are bakeable from absence; run attributes (fonts, size) have
# theme-dependent built-ins and are baked only when the source defines them.
_DD_BUILTIN: dict[tuple[str, str], str] = {
    ("spacing", "after"): "0",
    ("spacing", "before"): "0",
    ("spacing", "line"): "240",
    ("spacing", "lineRule"): "auto",
    ("ind", "left"): "0",
    ("ind", "start"): "0",
    ("ind", "right"): "0",
    ("ind", "end"): "0",
    ("ind", "firstLine"): "0",
    ("ind", "hanging"): "0",
    ("jc", "val"): "left",
}

_RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
    "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
    "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath",
]


def _ordered_get_or_add(parent: etree._Element, local: str, order: list[str]):
    existing = parent.find(qn(f"w:{local}"))
    if existing is not None:
        return existing
    el = etree.Element(qn(f"w:{local}"))
    my_rank = order.index(local)
    for child in parent:
        name = _localname(child)
        if name in order and order.index(name) > my_rank:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def _docdefaults_props(
    pkg: DocxPackage, which: str
) -> dict[tuple[str, str], str]:
    """Tracked (element, attribute) -> value pairs from styles.xml
    docDefaults. which: 'pPr' | 'rPr'."""
    out: dict[tuple[str, str], str] = {}
    if not pkg.has_part("word/styles.xml"):
        return out
    dd = pkg.root("word/styles.xml").find(qn("w:docDefaults"))
    if dd is None:
        return out
    if which == "pPr":
        holder = dd.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}")
        table = _DD_PPR_ATTRS
    else:
        holder = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
        table = _DD_RPR_ATTRS
    if holder is None:
        return out
    for elem_name, attrs in table.items():
        el = holder.find(qn(f"w:{elem_name}"))
        if el is None:
            continue
        for a in attrs:
            v = el.get(qn(f"w:{a}"))
            if v is not None:
                out[(elem_name, a)] = v
    return out


def _direct_keys(
    holder: etree._Element | None, table: dict[str, tuple[str, ...]]
) -> set[tuple[str, str]]:
    """Tracked (element, attribute) keys explicitly present on a pPr/rPr."""
    keys: set[tuple[str, str]] = set()
    if holder is None:
        return keys
    for elem_name, attrs in table.items():
        el = holder.find(qn(f"w:{elem_name}"))
        if el is None:
            continue
        for a in attrs:
            if el.get(qn(f"w:{a}")) is not None:
                keys.add((elem_name, a))
    return keys


class _DefaultsBaker:
    """Bakes source-docDefaults-inherited properties onto carried copies
    (formatting='source' only). See the section comment above."""

    def __init__(self, src: DocxPackage, pkg: DocxPackage):
        self.src_ppr = _docdefaults_props(src, "pPr")
        self.tgt_ppr = _docdefaults_props(pkg, "pPr")
        self.src_rpr = _docdefaults_props(src, "rPr")
        self.tgt_rpr = _docdefaults_props(pkg, "rPr")
        self.diff_ppr = {
            k
            for k in set(self.src_ppr) | set(self.tgt_ppr)
            if self.src_ppr.get(k) != self.tgt_ppr.get(k)
        }
        self.diff_rpr = {
            k
            for k in set(self.src_rpr) | set(self.tgt_rpr)
            if self.src_rpr.get(k) != self.tgt_rpr.get(k)
        }
        # Source style definitions for inheritance-chain checks.
        self.style_el: dict[str, etree._Element] = {}
        self.based_on: dict[str, str] = {}
        self.default_para_style: str | None = None
        if src.has_part("word/styles.xml"):
            for s in src.root("word/styles.xml").findall(qn("w:style")):
                sid = s.get(qn("w:styleId"))
                if not sid:
                    continue
                self.style_el[sid] = s
                base = s.find(qn("w:basedOn"))
                if base is not None and base.get(qn("w:val")):
                    self.based_on[sid] = base.get(qn("w:val"))
                if (
                    s.get(qn("w:type")) == "paragraph"
                    and s.get(qn("w:default")) in ("1", "true", "on")
                ):
                    self.default_para_style = sid
        self._chain_cache: dict[tuple[str, str], set[tuple[str, str]]] = {}
        self.paragraphs_baked = 0
        self.runs_baked = 0

    @property
    def active(self) -> bool:
        return bool(self.diff_ppr or self.diff_rpr)

    def _chain_keys(self, sid: str | None, which: str) -> set[tuple[str, str]]:
        """Tracked keys DEFINED anywhere along a source style's basedOn
        chain (attribute-level; a defined key stops docDefaults
        inheritance for that attribute)."""
        if sid is None:
            return set()
        memo = self._chain_cache.get((sid, which))
        if memo is not None:
            return memo
        keys: set[tuple[str, str]] = set()
        table = _DD_PPR_ATTRS if which == "pPr" else _DD_RPR_ATTRS
        cur: str | None = sid
        seen: set[str] = set()
        while cur is not None and cur not in seen:
            seen.add(cur)
            s = self.style_el.get(cur)
            if s is None:
                break
            keys |= _direct_keys(s.find(qn(f"w:{which}")), table)
            cur = self.based_on.get(cur)
        self._chain_cache[(sid, which)] = keys
        return keys

    def _bake_value(self, key: tuple[str, str], from_rpr: bool) -> str | None:
        """The source-effective value to write, or None when it cannot be
        determined safely (run built-ins are theme-dependent)."""
        src = self.src_rpr if from_rpr else self.src_ppr
        if key in src:
            return src[key]
        if from_rpr:
            return None
        return _DD_BUILTIN.get(key)

    def bake_paragraph(self, p: etree._Element) -> None:
        ppr = p.find(qn("w:pPr"))
        pstyle = None
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                pstyle = ps.get(qn("w:val"))
        if pstyle is None:
            pstyle = self.default_para_style
        # ---- paragraph properties
        if self.diff_ppr:
            skip = _direct_keys(ppr, _DD_PPR_ATTRS)
            skip |= self._chain_keys(pstyle, "pPr")
            changed = False
            for key in sorted(self.diff_ppr):
                if key in skip:
                    continue
                value = self._bake_value(key, from_rpr=False)
                if value is None:
                    continue
                if ppr is None:
                    ppr = etree.Element(qn("w:pPr"))
                    p.insert(0, ppr)
                el = _ordered_get_or_add(ppr, key[0], _PPR_BAKE_ORDER)
                el.set(qn(f"w:{key[1]}"), value)
                changed = True
            if changed:
                self.paragraphs_baked += 1
        # ---- run properties
        if not self.diff_rpr:
            return
        para_chain = self._chain_keys(pstyle, "rPr")
        for r in p.iter(qn("w:r")):
            if r.getparent() is ppr:
                continue  # the paragraph-mark rPr holder is not a run
            rpr = r.find(qn("w:rPr"))
            rstyle = None
            if rpr is not None:
                rs = rpr.find(qn("w:rStyle"))
                if rs is not None:
                    rstyle = rs.get(qn("w:val"))
            skip = _direct_keys(rpr, _DD_RPR_ATTRS)
            skip |= para_chain
            skip |= self._chain_keys(rstyle, "rPr")
            changed = False
            for key in sorted(self.diff_rpr):
                if key in skip:
                    continue
                value = self._bake_value(key, from_rpr=True)
                if value is None:
                    continue
                if rpr is None:
                    rpr = etree.Element(qn("w:rPr"))
                    r.insert(0, rpr)
                el = _ordered_get_or_add(rpr, key[0], _RPR_ORDER)
                el.set(qn(f"w:{key[1]}"), value)
                changed = True
            if changed:
                self.runs_baked += 1

    def report(self) -> dict:
        differing = sorted(
            [f"pPr.{e}.{a}" for (e, a) in self.diff_ppr]
            + [f"rPr.{e}.{a}" for (e, a) in self.diff_rpr]
        )
        return {
            "differ": True,
            "differing_properties": differing,
            "paragraphs_baked": self.paragraphs_baked,
            "runs_baked": self.runs_baked,
            "note": (
                "the source and target files have different document "
                "defaults; source paragraphs that inherited these "
                "properties were given explicit values so they keep the "
                "source appearance. The target's own content is untouched."
            ),
        }


# pPr insertion order for baked elements (CT_PPr child sequence, abridged).
_PPR_BAKE_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
]


# ------------------------------------------------- formatting strip (modes)

_FORMATTING_MODES = ("source", "merge", "destination")

# `formatting="merge"` (Word's Merge Formatting paste mode): semantic
# emphasis stays direct (bold, italic, underline, strike, sub/superscript,
# highlight, and anything else not listed below), while direct font-family/
# size/color and character-spacing run overrides plus paragraph spacing/
# line-spacing/indent overrides are stripped so the TARGET's styles govern
# the look of the carried text.
_MERGE_STRIP_RPR = frozenset({"rFonts", "sz", "szCs", "color", "spacing"})
_MERGE_STRIP_PPR = frozenset({"spacing", "ind"})

# `formatting="destination"`: strip ALL direct rPr/pPr formatting except
# what is structurally required or non-visual. Kept on runs: style refs,
# proofing language / noProof, RTL and complex-script structure, hidden-text
# flags (stripping those would REVEAL content, not restyle it), math, and
# revision marks. Kept on paragraphs: style ref, numbering, outline level,
# tab stops, text direction, table-conditional/div plumbing, the (recursed)
# paragraph-mark rPr, and revision records.
_DEST_KEEP_RPR = frozenset(
    {
        "rStyle", "lang", "noProof", "rtl", "cs",
        "vanish", "webHidden", "specVanish", "oMath",
        "ins", "del", "rPrChange",
    }
)
_DEST_KEEP_PPR = frozenset(
    {
        "pStyle", "numPr", "outlineLvl", "tabs",
        "bidi", "textDirection", "divId", "cnfStyle",
        "rPr", "sectPr", "pPrChange",
    }
)


def _strip_rpr_direct(rpr: etree._Element, mode: str) -> bool:
    changed = False
    for child in list(rpr):
        name = _localname(child)
        drop = (
            name in _MERGE_STRIP_RPR
            if mode == "merge"
            else name not in _DEST_KEEP_RPR
        )
        if drop:
            rpr.remove(child)
            changed = True
    return changed


def _strip_direct_formatting(elements, mode: str) -> tuple[int, int]:
    """Strip direct formatting from COPIED content per `formatting` mode
    ("merge" | "destination"). Operates only on the deep copies headed into
    the target — the source file is never touched. Returns counts of runs
    and paragraphs that lost at least one direct property."""
    runs_stripped = 0
    paras_stripped = 0
    for el in elements:
        for p in el.iter(qn("w:p")):
            changed = False
            ppr = p.find(qn("w:pPr"))
            if ppr is not None:
                for child in list(ppr):
                    name = _localname(child)
                    if name == "rPr":  # paragraph-mark run properties
                        if _strip_rpr_direct(child, mode):
                            changed = True
                        if len(child) == 0:
                            ppr.remove(child)
                        continue
                    drop = (
                        name in _MERGE_STRIP_PPR
                        if mode == "merge"
                        else name not in _DEST_KEEP_PPR
                    )
                    if drop:
                        ppr.remove(child)
                        changed = True
                if len(ppr) == 0:
                    p.remove(ppr)
            if changed:
                paras_stripped += 1
        for r in el.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            if _strip_rpr_direct(rpr, mode):
                runs_stripped += 1
                if len(rpr) == 0:
                    r.remove(rpr)
    return runs_stripped, paras_stripped


# ---------------------------------------------------------------- main tools


def _check_positioners(
    after_index: int | None,
    after_anchor: str | None,
    at_end: bool,
    before_first: bool = False,
) -> None:
    positioners = (
        (after_index is not None) + (after_anchor is not None)
        + bool(at_end) + bool(before_first)
    )
    if positioners != 1:
        raise WordMcpError(
            "give exactly one positioner: after_index, after_anchor, "
            "at_end=True, or before_first=True"
        )


def _open_source(pkg: DocxPackage, source_path: str) -> DocxPackage:
    src = DocxPackage(source_path)
    try:
        same = src.path.resolve() == pkg.path.resolve()
    except OSError:  # pragma: no cover
        same = str(src.path) == str(pkg.path)
    if same:
        raise WordMcpError(
            "source and target are the same file — refusing to insert a "
            "document into itself"
        )
    return src


def insert_document(
    pkg: DocxPackage,
    source_path: str,
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    before_first: bool = False,
    formatting: str = "source",
) -> dict:
    """Insert the ENTIRE body content of the document at `source_path` into
    this document at one position.

    Positioning (exactly one required):
    - after_index: body ITEM index — paragraphs and tables counted together
      in document order (item 0 is the first block, whether paragraph or
      table). Insertion happens after that item. NOTE: this differs from the
      paragraph-only indices of get_text/insert_paragraphs when the target
      contains tables.
    - after_anchor: a paragraph whose FULL text exactly matches (whitespace-
      trimmed). More than one match refuses with every match location listed.
    - at_end: after the last body item, before the trailing sectPr.

    Carried: paragraphs, tables (incl. merged cells and nested tables),
    images, hyperlinks, numbered/bulleted lists, footnote/endnote references
    with their definitions, bookmarks, charts, tracked changes, equations.
    Styles reconcile by name (target formatting wins on a name match;
    unmatched styles are cloned). The source's trailing sectPr is never
    carried; mid-content section breaks are stripped and reported. Comment
    references are stripped and counted — comment transplant is out of scope.
    OLE objects, ActiveX controls, subdocuments, altChunks, and unknown
    embedded parts refuse the whole insertion (strip them from the source
    first); nothing is ever half-applied.

    formatting (Word paste modes, applied to the carried COPIES only):
    - "source" (default): direct formatting preserved — Word's InsertFile.
    - "merge": semantic emphasis direct formatting kept (bold, italic,
      underline, strike, sub/superscript, highlight) but direct font-family/
      size/color/character-spacing run overrides and paragraph spacing/
      line-spacing/indent overrides stripped — Word's Merge Formatting.
    - "destination": ALL direct rPr/pPr formatting stripped except what is
      structurally required (numPr, outlineLvl, tab stops stay); the
      target's styles govern rendering entirely.
    With "merge"/"destination" the result reports runs/paragraphs stripped.
    With "source", properties the source paragraphs inherited from their
    file's document defaults (docDefaults) are resolved to explicit values
    when the two files' defaults differ, so the carried content keeps the
    source's rendered spacing/indent/font; the result reports what was
    reconciled under "document_defaults".

    The source file is never modified."""
    _check_positioners(after_index, after_anchor, at_end, before_first)
    if formatting not in _FORMATTING_MODES:
        raise WordMcpError(
            f"formatting must be one of {list(_FORMATTING_MODES)}, "
            f"got {formatting!r}"
        )
    src = _open_source(pkg, source_path)
    copied: list[etree._Element] = []
    for c in src.body():
        if _localname(c) == "sectPr":
            continue
        if _localname(c) in ("commentRangeStart", "commentRangeEnd"):
            continue  # body-level comment markers; counted via references
        copied.append(copy.deepcopy(c))
    if not copied:
        raise TargetNotFound(
            f"source document has no body content to insert: {source_path}"
        )
    return _transplant(
        pkg,
        src,
        copied,
        after_index=after_index,
        after_anchor=after_anchor,
        at_end=at_end,
        before_first=before_first,
        formatting=formatting,
    )


def copy_table(
    pkg: DocxPackage,
    source_path: str,
    table_index: int,
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
) -> dict:
    """Transplant ONE top-level table from the document at `source_path` into
    this document, through the same reconciliation pipeline as
    insert_document (styles matched by name with the target's formatting
    governing, fresh numbering instances, images/hyperlinks re-registered,
    footnote/endnote definitions carried under new ids, bookmark ids
    remapped) scoped to the single table element.

    table_index counts the source's top-level body tables in document order
    (0-based; tables nested inside cells travel with their parent and are
    not separately addressable). Positioning follows the insert_document
    contract exactly: exactly one of after_index (body ITEM index of the
    TARGET, paragraphs and tables counted together), after_anchor (exact
    whole-paragraph text match, ambiguity refuses with every match listed),
    or at_end. Uncarryable content inside the table (OLE objects, ActiveX,
    subdocuments, altChunks, unknown embedded parts) refuses the whole copy;
    nothing is ever half-applied. The source file is never modified."""
    _check_positioners(after_index, after_anchor, at_end)
    if not isinstance(table_index, int) or isinstance(table_index, bool):
        raise WordMcpError(f"table_index must be an integer, got {table_index!r}")
    src = _open_source(pkg, source_path)
    tables = [el for el in _body_blocks(src) if _localname(el) == "tbl"]
    if not tables:
        raise TargetNotFound(
            f"source document has no top-level body tables: {source_path}"
        )
    if not 0 <= table_index < len(tables):
        raise TargetNotFound(
            f"table_index {table_index} out of range: the source has "
            f"{len(tables)} top-level body table(s) (valid indices "
            f"0-{len(tables) - 1})"
        )
    tbl = tables[table_index]
    rows = len(tbl.findall(qn("w:tr")))
    grid = tbl.find(qn("w:tblGrid"))
    cols = len(grid.findall(qn("w:gridCol"))) if grid is not None else None
    result = _transplant(
        pkg,
        src,
        [copy.deepcopy(tbl)],
        after_index=after_index,
        after_anchor=after_anchor,
        at_end=at_end,
        formatting="source",
    )
    result["source_table_index"] = table_index
    result["rows"] = rows
    if cols is not None:
        result["columns"] = cols
    return result


def _transplant(
    pkg: DocxPackage,
    src: DocxPackage,
    copied: list[etree._Element],
    *,
    after_index: int | None,
    after_anchor: str | None,
    at_end: bool,
    formatting: str,
    before_first: bool = False,
) -> dict:
    """Shared reconciliation-and-insert pipeline over already-deep-copied
    source elements. Phase A only scans (any refusal leaves the target
    untouched); phase B mutates the in-memory target, which is saved only by
    the caller."""
    pos_mode, ref_el, start_item = _resolve_position(
        pkg, after_index, after_anchor, at_end, before_first
    )

    # ---- phase A: prepare and SCAN (no target mutation on refusal)

    # Mid-content section breaks: strip BEFORE the relationship scan so their
    # header/footer references never enter the transplant set.
    section_breaks_stripped = 0
    for el in copied:
        for sp in list(el.iter(qn("w:sectPr"))):
            parent = sp.getparent()
            if parent is not None:
                parent.remove(sp)
                section_breaks_stripped += 1

    # Referenced note definitions, copied now so they join every later pass.
    note_plan: dict[str, list[tuple[str, etree._Element]]] = {}
    for kind, cfg in _notes._KINDS.items():
        ref_ids: list[str] = []
        for el in copied:
            for ref in el.iter(qn(cfg["body_ref"])):
                rid = ref.get(qn("w:id"))
                if rid is not None and rid not in ref_ids:
                    ref_ids.append(rid)
        if not ref_ids:
            continue
        if not src.has_part(cfg["part"]):
            raise UnsupportedStructure(
                f"source references {kind} ids {ref_ids} but has no "
                f"{cfg['part']} — the source document is corrupt"
            )
        defs = {
            n.get(qn("w:id")): n
            for n in src.root(cfg["part"]).findall(qn(cfg["note"]))
        }
        missing = [i for i in ref_ids if i not in defs]
        if missing:
            raise UnsupportedStructure(
                f"source references {kind} ids {missing} that have no "
                f"definition in {cfg['part']} — the source document is corrupt"
            )
        note_plan[kind] = [(i, copy.deepcopy(defs[i])) for i in ref_ids]

    units: list[tuple[etree._Element, str]] = [(el, "document") for el in copied]
    for kind, pairs in note_plan.items():
        units += [(d, kind) for _, d in pairs]

    # Comment references: stripped cleanly, counted, reported.
    comments_stripped = 0
    for el, _story in units:
        for node in list(el.iter(qn("w:commentReference"))):
            run = node.getparent()
            if run is not None and _localname(run) == "r":
                container = run.getparent()
                if container is not None:
                    container.remove(run)
            else:  # pragma: no cover - malformed but strip anyway
                node.getparent().remove(node)
            comments_stripped += 1
        for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
            for node in list(el.iter(qn(tag))):
                if node.getparent() is not None:
                    node.getparent().remove(node)

    # Blocking content scan.
    blocked: dict[str, int] = {}
    for el, _story in units:
        for tag, label in _BLOCKED_ELEMENTS.items():
            hits = len(list(el.iter(tag)))
            if hits:
                blocked[label] = blocked.get(label, 0) + hits
    if blocked:
        raise UnsupportedStructure(
            "source contains content that cannot be transplanted safely: "
            + "; ".join(f"{v}x {k}" for k, v in sorted(blocked.items()))
            + ". Nothing was inserted. Strip these from the source first "
            "(e.g., delete the embedded objects in Word) and retry."
        )

    # Relationship scan and classification.
    src_rels: dict[str, dict[str, etree._Element]] = {}
    for story, (_dst, rels_name) in _STORY_PARTS.items():
        rels: dict[str, etree._Element] = {}
        if src.has_part(rels_name):
            for rel in src.root(rels_name):
                rels[rel.get("Id")] = rel
        src_rels[story] = rels

    rel_uses: list[tuple[str, etree._Element, str, str]] = []
    rel_refusals: list[str] = []
    for el, story in units:
        for node, key, rid in _iter_rid_attrs(el):
            rel = src_rels[story].get(rid)
            if rel is None:
                rel_refusals.append(
                    f"relationship {rid} (referenced from {story} content) "
                    "has no definition in the source"
                )
                continue
            rtype = rel.get("Type")
            external = rel.get("TargetMode") == "External"
            if not external:
                target_part = _resolve_rel_target(
                    _STORY_PARTS[story][1], rel.get("Target", "")
                )
                if rtype == _REL_IMAGE or rtype in _SUBTREE_REL_TYPES:
                    if not src.has_part(target_part):
                        rel_refusals.append(
                            f"relationship {rid} targets missing source part "
                            f"{target_part}"
                        )
                else:
                    rel_refusals.append(
                        f"unsupported embedded content: relationship {rid} of "
                        f"type {rtype} -> {target_part}"
                    )
            rel_uses.append((story, node, key, rid))
    if rel_refusals:
        raise UnsupportedStructure(
            "source content references parts this tool cannot transplant "
            "safely: " + "; ".join(sorted(set(rel_refusals)))
            + ". Nothing was inserted. Strip that content from the source "
            "first and retry."
        )

    # Pre-existing target note problems (never blamed on this insertion).
    pre_notes = _notes.validate_notes(pkg)

    # ---- phase B: mutate the in-memory target (saved only by the caller,
    # so any exception below still leaves the file untouched)

    # B0. Formatting mode: strip direct formatting from the carried COPIES
    # (body content and note definitions alike; the source is untouched).
    runs_stripped = paras_stripped = 0
    if formatting != "source":
        runs_stripped, paras_stripped = _strip_direct_formatting(
            [el for el, _ in units], formatting
        )

    # B0b. formatting="source": reconcile document defaults. Properties the
    # source paragraphs inherited from their file's docDefaults are baked
    # as explicit values wherever the two files' defaults differ, so the
    # carried content keeps the source's rendered look instead of silently
    # resolving against the target's defaults (field test, 2026-09-03).
    dd_baker = None
    if formatting == "source":
        dd_baker = _DefaultsBaker(src, pkg)
        if dd_baker.active:
            for el, _story in units:
                if el.tag == qn("w:p"):
                    dd_baker.bake_paragraph(el)
                for p in el.iterdescendants(qn("w:p")):
                    dd_baker.bake_paragraph(p)
        else:
            dd_baker = None

    # B1+B2. Styles and numbering, to a fixpoint (cloned styles can reference
    # numbering; cloned numbering can reference styles via styleLink/pStyle).
    styles = _StyleResolver(src, pkg)
    numbering = _NumberingResolver(src, pkg)
    unit_els = [el for el, _ in units]
    pending_sids = _collect_style_refs(unit_els)
    pending_nids = _collect_num_refs(unit_els)
    for _round in range(10):
        if not pending_sids and not pending_nids:
            break
        for sid in sorted(pending_sids):
            styles.resolve(sid)
        for nid in sorted(pending_nids):
            numbering.resolve(nid)
        pending_sids = {
            s
            for s in _collect_style_refs(numbering.cloned_abstracts)
            | _collect_style_refs(styles.cloned_defs)
            if s not in styles._done
        }
        pending_nids = {
            n
            for n in _collect_num_refs(styles.cloned_defs)
            if n not in numbering._done
        }
    remap_targets = unit_els + styles.cloned_defs + numbering.cloned_abstracts
    _apply_style_remap(remap_targets, styles.remap)
    lists_stripped = _apply_num_remap(
        remap_targets, numbering.map, set(numbering.unresolved)
    )

    # B3. Notes: ensure parts/styles exist, assign fresh ids, retag, append.
    notes_carried = {"footnote": 0, "endnote": 0}
    for kind, pairs in note_plan.items():
        cfg = _notes._KINDS[kind]
        _notes._ensure_part(pkg, kind)
        _notes._ensure_styles(pkg, kind)
        notes_root = pkg.root(cfg["part"])
        id_map: dict[str, str] = {}
        for old_id, def_copy in pairs:
            new_id = str(_notes._next_id(pkg, kind))
            def_copy.set(qn("w:id"), new_id)
            notes_root.append(def_copy)
            id_map[old_id] = new_id
        pkg.mark_dirty(cfg["part"])
        for el in copied:
            for ref in el.iter(qn(cfg["body_ref"])):
                old = ref.get(qn("w:id"))
                if old in id_map:
                    ref.set(qn("w:id"), id_map[old])
        notes_carried[kind] = len(pairs)

    # B4. Relationships and parts.
    rid_memo: dict[tuple[str, str], str] = {}
    media_memo: dict[str, str] = {}
    subtree_memo: dict[str, str] = {}
    chart_parts: set[str] = set()
    rels_ids: dict[str, set[str]] = {}
    hyperlinks_carried = 0
    external_rels_carried = 0

    def _dst_rels_root(story: str) -> tuple[str, etree._Element]:
        name = _STORY_PARTS[story][1]
        _ensure_rels_part(pkg, name)
        root = pkg.root(name)
        if name not in rels_ids:
            rels_ids[name] = {r.get("Id") for r in root}
        return name, root

    def _add_rel(
        story: str, rtype: str, target: str, *, external: bool
    ) -> str:
        name, root = _dst_rels_root(story)
        n = 1
        while f"rId{n}" in rels_ids[name]:
            n += 1
        rid = f"rId{n}"
        rels_ids[name].add(rid)
        rel = etree.SubElement(root, f"{{{_REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rtype)
        rel.set("Target", target)
        if external:
            rel.set("TargetMode", "External")
        pkg.mark_dirty(name)
        return rid

    def _copy_media(src_part: str) -> str:
        if src_part in media_memo:
            return media_memo[src_part]
        ext = (
            "." + src_part.rsplit(".", 1)[1].lower()
            if "." in src_part.rsplit("/", 1)[1]
            else ""
        )
        n = 1
        while pkg.has_part(f"word/media/image{n}{ext}") or any(
            name.startswith(f"word/media/image{n}.")
            for name in pkg.part_names()
        ):
            n += 1
        new_part = f"word/media/image{n}{ext}"
        pkg.set_raw_part(new_part, src.raw_part(src_part))
        _ensure_content_type(pkg, src, src_part, new_part)
        media_memo[src_part] = new_part
        return new_part

    def _copy_subtree(src_part: str) -> str:
        """Copy a part plus its private dependency subtree (chart ->
        embedded workbook / colors / style parts), keeping the part's
        internal rIds valid by rewriting only relationship Targets."""
        if src_part in subtree_memo:
            return subtree_memo[src_part]
        new_part = _free_part_name(pkg, src_part)
        pkg.set_raw_part(new_part, b"")  # reserve the name before recursing
        subtree_memo[src_part] = new_part
        src_rels_name = _rels_part_for(src_part)
        if src.has_part(src_rels_name):
            rels_root = copy.deepcopy(src.root(src_rels_name))
            for rel in rels_root:
                if rel.get("TargetMode") == "External":
                    continue
                child_src = _resolve_rel_target(
                    src_rels_name, rel.get("Target", "")
                )
                if not src.has_part(child_src):
                    raise UnsupportedStructure(
                        f"source part {src_part} references missing part "
                        f"{child_src} — the source document is corrupt; "
                        "nothing was inserted"
                    )
                child_new = _copy_subtree(child_src)
                rel.set(
                    "Target",
                    posixpath.relpath(
                        child_new, start=new_part.rsplit("/", 1)[0]
                    ),
                )
            pkg.set_raw_part(
                _rels_part_for(new_part),
                etree.tostring(
                    rels_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                ),
            )
        pkg.set_raw_part(new_part, src.raw_part(src_part))
        _ensure_content_type(pkg, src, src_part, new_part)
        return new_part

    for story, node, key, rid in rel_uses:
        memo_key = (story, rid)
        new_rid = rid_memo.get(memo_key)
        if new_rid is None:
            rel = src_rels[story][rid]
            rtype = rel.get("Type")
            if rel.get("TargetMode") == "External":
                new_rid = _add_rel(
                    story, rtype, rel.get("Target", ""), external=True
                )
                external_rels_carried += 1
                if rtype == _REL_HYPERLINK:
                    hyperlinks_carried += 1
            else:
                src_part = _resolve_rel_target(
                    _STORY_PARTS[story][1], rel.get("Target", "")
                )
                if rtype == _REL_IMAGE:
                    new_part = _copy_media(src_part)
                else:  # chart / chartEx (phase A allowed nothing else)
                    new_part = _copy_subtree(src_part)
                    chart_parts.add(src_part)
                # Both document.xml and the notes parts live in word/, so the
                # relative target is the part name minus the word/ prefix.
                new_rid = _add_rel(
                    story, rtype, new_part.split("word/", 1)[1], external=False
                )
            rid_memo[memo_key] = new_rid
        node.set(key, new_rid)

    # B5. Bookmarks: fresh ids; name collisions renamed and retargeted.
    existing_names: set[str] = set()
    max_bm_id = 0
    for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if not pkg.has_part(part):
            continue
        for bs in pkg.root(part).iter(qn("w:bookmarkStart")):
            name = bs.get(qn("w:name"))
            if name:
                existing_names.add(name)
            try:
                max_bm_id = max(max_bm_id, int(bs.get(qn("w:id"), "0") or 0))
            except ValueError:  # pragma: no cover
                pass
    bm_id_map: dict[str, str] = {}
    bm_renames: list[dict] = []
    name_map: dict[str, str] = {}
    bookmarks_carried = 0
    for el, _story in units:
        for bs in el.iter(qn("w:bookmarkStart")):
            old_id = bs.get(qn("w:id"))
            name = bs.get(qn("w:name")) or ""
            max_bm_id += 1
            new_id = str(max_bm_id)
            if old_id is not None:
                bm_id_map[old_id] = new_id
            bs.set(qn("w:id"), new_id)
            if name:
                if name in existing_names:
                    n = 1
                    new_name = f"{name[:34]}_ins{n}"
                    while new_name in existing_names:
                        n += 1
                        new_name = f"{name[:34]}_ins{n}"
                    bs.set(qn("w:name"), new_name)
                    name_map[name] = new_name
                    bm_renames.append({"from": name, "to": new_name})
                    existing_names.add(new_name)
                else:
                    existing_names.add(name)
            bookmarks_carried += 1
    for el, _story in units:
        for be in el.iter(qn("w:bookmarkEnd")):
            old_id = be.get(qn("w:id"))
            if old_id in bm_id_map:
                be.set(qn("w:id"), bm_id_map[old_id])
    # Retarget internal links and REF-family fields to renamed bookmarks —
    # only within the inserted content (target content is never rewritten).
    refs_retargeted = 0
    if name_map:
        for el, _story in units:
            for link in el.iter(qn("w:hyperlink")):
                anchor = link.get(qn("w:anchor"))
                if anchor in name_map:
                    link.set(qn("w:anchor"), name_map[anchor])
                    refs_retargeted += 1
            for instr in el.iter(qn("w:instrText")):
                text = instr.text or ""
                for old, new in name_map.items():
                    pattern = (
                        r"\b(REF|PAGEREF|NOTEREF|HYPERLINK\s+\\l)(\s+\"?)"
                        + re.escape(old)
                        + r"\b"
                    )
                    new_text = re.sub(pattern, r"\g<1>\g<2>" + new, text)
                    if new_text != text:
                        text = new_text
                        refs_retargeted += 1
                instr.text = text

    # B6. docPr ids unique across the whole document (body + notes).
    max_docpr = 0
    for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if not pkg.has_part(part):
            continue
        for d in pkg.root(part).iter(f"{{{_WP}}}docPr"):
            try:
                max_docpr = max(max_docpr, int(d.get("id", "0") or 0))
            except ValueError:  # pragma: no cover
                pass
    for el, _story in units:
        for d in el.iter(f"{{{_WP}}}docPr"):
            max_docpr += 1
            d.set("id", str(max_docpr))

    # B7. Insert at the resolved position.
    body = pkg.body()
    if ref_el is None:
        sectpr = body.find(qn("w:sectPr"))
        for el in copied:
            if sectpr is not None:
                sectpr.addprevious(el)
            else:
                body.append(el)
    elif pos_mode == "before":
        for el in copied:
            ref_el.addprevious(el)
    else:
        for el in reversed(copied):
            ref_el.addnext(el)
    pkg.mark_dirty()

    # B8. Closure validation on everything just inserted; a failure here
    # aborts before the caller saves, leaving the file untouched.
    for el, story in units:
        rels_name, rels_root = _dst_rels_root(story)
        rel_map = {r.get("Id"): r for r in rels_root}
        for _node, _key, rid in _iter_rid_attrs(el):
            rel = rel_map.get(rid)
            if rel is None:
                raise ValidationFailed(
                    f"internal error: inserted content references undefined "
                    f"relationship {rid} in {rels_name}; document not saved"
                )
            if rel.get("TargetMode") != "External":
                resolved = _resolve_rel_target(
                    rels_name, rel.get("Target", "")
                )
                if not pkg.has_part(resolved):
                    raise ValidationFailed(
                        f"internal error: relationship {rid} targets missing "
                        f"part {resolved}; document not saved"
                    )
    post_notes = _notes.validate_notes(pkg)
    for kind_key, report in post_notes.items():
        pre = pre_notes.get(kind_key, {})
        new_missing = set(report["missing_definitions"]) - set(
            pre.get("missing_definitions", [])
        )
        new_dups = set(report["duplicate_references"]) - set(
            pre.get("duplicate_references", [])
        )
        if new_missing or new_dups:
            raise ValidationFailed(
                f"internal error: note integrity broke during insertion "
                f"({kind_key}: missing {sorted(new_missing)}, duplicates "
                f"{sorted(new_dups)}); document not saved"
            )

    n_paras = sum(1 for el in copied if _localname(el) == "p")
    n_tables = sum(1 for el in copied if _localname(el) == "tbl")
    n_items = n_paras + n_tables
    result = {
        "inserted_from": str(src.path),
        "position": {
            "mode": (
                "at_end"
                if at_end
                else "before_first"
                if before_first
                else ("after_index" if after_index is not None else "after_anchor")
            ),
            "starts_at_body_item": start_item,
        },
        "body_item_range": (
            [start_item, start_item + n_items - 1] if n_items else None
        ),
        "paragraphs": n_paras,
        "tables": n_tables,
        "images_carried": len(media_memo),
        "charts_carried": len(chart_parts),
        "hyperlinks_carried": hyperlinks_carried,
        "external_relationships_carried": external_rels_carried,
        "footnotes_carried": notes_carried["footnote"],
        "endnotes_carried": notes_carried["endnote"],
        "lists_carried": len(
            [k for k, v in numbering.map.items() if k != "0"]
        ),
        "styles": {
            "matched_by_name": len(styles.matched),
            "remapped_ids": dict(sorted(styles.remap.items())),
            "cloned": styles.cloned,
        },
        "bookmarks_carried": bookmarks_carried,
        "bookmarks_renamed": bm_renames,
        "bookmark_refs_retargeted": refs_retargeted,
        "comments_stripped": comments_stripped,
        "section_breaks_stripped": section_breaks_stripped,
        "formatting_mode": formatting,
    }
    if formatting != "source":
        result["runs_stripped"] = runs_stripped
        result["paragraphs_stripped"] = paras_stripped
    if dd_baker is not None:
        result["document_defaults"] = dd_baker.report()
    if styles.unresolved:
        result["style_refs_unresolved_in_source"] = sorted(styles.unresolved)
    if numbering.unresolved:
        result["numbering_refs_unresolved_in_source"] = sorted(
            numbering.unresolved
        )
        result["numbering_refs_stripped"] = lists_stripped
    return result
