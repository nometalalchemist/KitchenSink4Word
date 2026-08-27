"""Academic-integrity validators: cross-references and captions.

Read-only checks for the failure modes that survive proofreading: REF/PAGEREF
fields pointing at bookmarks that were deleted with their text, plain-text
references ("see Figure 3") that no longer match any actual figure, tables
and images that lost their captions in a restructure, and numbering that
mixes sequential ("Figure 3") with chapter-relative ("Figure 4.2") styles.

Field-number caveat: SEQ and REF results are computed by Word, not by this
library. A document whose fields were never updated carries placeholder
results ("#"), so number-matching checks against such a document report
unverified rather than verified — that is honest, not a bug. Update fields in
Word first for full verification.
"""

from __future__ import annotations

import re

from lxml import etree

from ..core.package import DocxPackage, qn
from .read import body_items, get_outline, paragraph_text
from .reffields import scan_complex_fields

_STORY_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")

# The textual-reference shapes worth flagging ("see Figure 3", "Chapter 4.2").
_TEXT_REF_RE = re.compile(
    r"\b(Figure|Table|Chapter|Section|Appendix)\s+([\dA-Z][\d.]*)"
)
_REF_FIELD_RE = re.compile(r'^\s*(REF|PAGEREF)\s+(?:"([^"]+)"|(\S+))')
_SEQ_LABEL_RE = re.compile(r"\bSEQ\s+(\w+)")


def _containing_paragraph(el: etree._Element) -> etree._Element | None:
    node = el
    while node is not None and etree.QName(node).localname != "p":
        node = node.getparent()
    return node


def _body_index_map(pkg: DocxPackage) -> tuple[dict[int, int], list]:
    """(id(paragraph element) -> body index, keepalive list). The caller must
    hold the keepalive list while using the map: lxml recreates element
    proxies unless a reference stays alive, and only a live proxy has a
    stable id()."""
    els = [
        (el, idx) for kind, idx, el in body_items(pkg) if kind == "paragraph"
    ]
    return {id(el): idx for el, idx in els}, [el for el, _ in els]


def _para_index_of(
    el: etree._Element, part: str, index_map: dict[int, int]
) -> int | None:
    p = _containing_paragraph(el)
    if p is not None and part == "word/document.xml":
        return index_map.get(id(p))
    return None


def _norm_number(num: str) -> str:
    # "4.2." at a sentence end -> "4.2"; bare trailing dots never carry meaning.
    return num.rstrip(".")


# ------------------------------------------------------------ cross-references


def validate_cross_references(pkg: DocxPackage) -> dict:
    """Check every REF/PAGEREF cross-reference field against the bookmarks
    that actually exist, and flag plain-text references ("see Figure 3") that
    match no caption or heading number.

    Findings, most to least serious:
    - broken: field references a bookmark that does not exist (the reference
      renders as an error in Word).
    - unverified_text_reference: text like "Figure 3" with no matching
      caption/heading number found. HEURISTIC — review candidates only; a
      document whose fields were never updated in Word still shows
      placeholder numbers and will land here.
    - unreferenced_bookmarks: user bookmarks nothing points at (informational;
      Word-internal underscore bookmarks are excluded).
    """
    index_map, _keepalive = _body_index_map(pkg)

    # Bookmark targets from every story part (REF can target notes bookmarks).
    bookmarks: dict[str, int | None] = {}
    for part in _STORY_PARTS:
        if not pkg.has_part(part):
            continue
        for bs in pkg.root(part).iter(qn("w:bookmarkStart")):
            name = bs.get(qn("w:name"))
            if name and name not in bookmarks:
                bookmarks[name] = _para_index_of(bs, part, index_map)

    broken: list[dict] = []
    referenced: set[str] = set()
    ref_fields = 0
    for part in _STORY_PARTS:
        if not pkg.has_part(part):
            continue
        root = pkg.root(part)
        fields, _ = scan_complex_fields(root)
        for rec in fields:
            m = _REF_FIELD_RE.match(rec["instr"])
            if not m:
                continue
            ref_fields += 1
            target = m.group(2) or m.group(3)
            referenced.add(target)
            if target not in bookmarks:
                broken.append(
                    {
                        "field": m.group(1),
                        "bookmark": target,
                        "part": part,
                        "paragraph_index": _para_index_of(
                            rec["begin_el"], part, index_map
                        ),
                    }
                )
        # Internal hyperlinks (w:anchor) also consume bookmarks; count them
        # as references so linked-to bookmarks are not reported unreferenced.
        for link in root.iter(qn("w:hyperlink")):
            anchor = link.get(qn("w:anchor"))
            if anchor:
                referenced.add(anchor)

    unreferenced = [
        {"name": name, "paragraph_index": bookmarks[name]}
        for name in sorted(bookmarks)
        if name not in referenced and not name.startswith("_")
    ]

    # --- textual references vs actual caption/heading numbers (heuristic)
    caption_numbers: dict[str, set[str]] = {}
    for p in pkg.root().iter(qn("w:p")):
        if not _paragraph_has_seq(p):
            continue
        m = re.match(
            r"^\s*(Figure|Table|Equation)\s+([\dA-Z][\d.]*)", paragraph_text(p)
        )
        if m:
            caption_numbers.setdefault(m.group(1), set()).add(
                _norm_number(m.group(2))
            )

    heading_numbers: dict[str, set[str]] = {}
    generic_heading_numbers: set[str] = set()
    for h in get_outline(pkg):
        text = h["text"]
        m = re.match(r"^(Chapter|Section|Appendix)\s+([\dA-Z][\d.]*)", text)
        if m:
            heading_numbers.setdefault(m.group(1), set()).add(
                _norm_number(m.group(2))
            )
        m = re.match(r"^(\d[\d.]*)\b", text)
        if m:
            generic_heading_numbers.add(_norm_number(m.group(1)))

    unverified: list[dict] = []
    checked = 0
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph":
            continue
        for m in _TEXT_REF_RE.finditer(paragraph_text(el)):
            label, num = m.group(1), _norm_number(m.group(2))
            if not num:
                continue
            checked += 1
            if label in ("Figure", "Table"):
                verified = num in caption_numbers.get(label, set())
            else:  # Chapter | Section | Appendix
                verified = (
                    num in heading_numbers.get(label, set())
                    or num in generic_heading_numbers
                )
            if not verified:
                unverified.append(
                    {"text": f"{label} {m.group(2)}", "paragraph_index": idx}
                )

    return {
        "bookmarks": len(bookmarks),
        "ref_fields": ref_fields,
        "broken": broken,
        "unreferenced_bookmarks": unreferenced,
        "text_references": {
            "checked": checked,
            "unverified": unverified,
            "note": (
                "heuristic text scan — unverified items are review "
                "candidates, not confirmed errors; numbers computed by Word "
                "(un-updated fields show placeholders) cannot be verified "
                "here"
            ),
        },
        "ok": not broken,
    }


# ------------------------------------------------------------------- captions


def _paragraph_has_seq(p: etree._Element) -> bool:
    return any(
        (it.text or "").strip().startswith("SEQ")
        for it in p.iter(qn("w:instrText"))
    )


def _caption_info(p: etree._Element) -> dict:
    """label / rendered number / numbering convention for a caption
    paragraph (one that carries a SEQ field)."""
    instrs = [(it.text or "") for it in p.iter(qn("w:instrText"))]
    joined = " ".join(instrs)
    label_m = _SEQ_LABEL_RE.search(joined)
    label = label_m.group(1) if label_m else None
    text = paragraph_text(p)
    num_m = re.match(r"^\s*\w+\s+([\dA-Z][\d.]*)", text)
    number = _norm_number(num_m.group(1)) if num_m else None
    # Chapter-relative markers: the SEQ \s switch (restart per heading level),
    # a STYLEREF companion field (the "4" in "4.2"), or a dotted cached number.
    chapter_relative = (
        "\\s" in joined
        or "STYLEREF" in joined
        or (number is not None and "." in number)
    )
    return {
        "label": label,
        "number": number,
        "text": text.strip(),
        "convention": "chapter_relative" if chapter_relative else "sequential",
    }


def _is_block(node: etree._Element) -> bool:
    return isinstance(node.tag, str) and etree.QName(node).localname in (
        "p",
        "tbl",
    )


def _adjacent_blocks(el: etree._Element) -> list[etree._Element]:
    """Nearest preceding and following body-level p/tbl siblings."""
    out = []
    node = el.getprevious()
    while node is not None and not _is_block(node):
        node = node.getprevious()
    if node is not None:
        out.append(node)
    node = el.getnext()
    while node is not None and not _is_block(node):
        node = node.getnext()
    if node is not None:
        out.append(node)
    return out


def validate_captions(pkg: DocxPackage) -> dict:
    """Check that every body-level table and image has an adjacent caption
    paragraph carrying a SEQ number field, and that caption numbering follows
    ONE convention per label (sequential "Figure 3" vs chapter-relative
    "Figure 4.2") rather than a mix.

    Only body-level tables and images are checked; drawings nested inside
    table cells or text boxes are out of scope (noted in the result)."""

    captions: list[dict] = []
    missing: list[dict] = []
    tables_checked = images_checked = 0

    for kind, idx, el in body_items(pkg):
        if kind == "paragraph" and _paragraph_has_seq(el):
            info = _caption_info(el)
            info["paragraph_index"] = idx
            captions.append(info)

    for kind, idx, el in body_items(pkg):
        if kind == "table":
            tables_checked += 1
            has_caption = any(
                etree.QName(n).localname == "p" and _paragraph_has_seq(n)
                for n in _adjacent_blocks(el)
            )
            if not has_caption:
                missing.append({"kind": "table", "table_index": idx})
        elif kind == "paragraph":
            drawing = el.find(f".//{qn('w:drawing')}")
            if drawing is None:
                continue
            images_checked += 1
            anchored = (
                drawing.find(
                    "{http://schemas.openxmlformats.org/drawingml/2006/"
                    "wordprocessingDrawing}anchor"
                )
                is not None
            )
            has_caption = any(
                etree.QName(n).localname == "p" and _paragraph_has_seq(n)
                for n in _adjacent_blocks(el)
            )
            if not has_caption:
                missing.append(
                    {
                        "kind": "image",
                        "paragraph_index": idx,
                        "anchored": anchored,
                    }
                )

    conventions: dict[str, str] = {}
    mixed: list[str] = []
    for cap in captions:
        label = cap["label"]
        if label is None:
            continue
        seen = conventions.get(label)
        if seen is None:
            conventions[label] = cap["convention"]
        elif seen != cap["convention"] and seen != "mixed":
            conventions[label] = "mixed"
            mixed.append(label)

    return {
        "tables_checked": tables_checked,
        "images_checked": images_checked,
        "captions": captions,
        "missing": missing,
        "conventions": conventions,
        "mixed_conventions": sorted(mixed),
        "ok": not missing and not mixed,
        "note": (
            "body-level tables and images only; drawings inside table cells "
            "or text boxes are not checked. Convention detection reads the "
            "SEQ switches and cached numbers — captions never updated in "
            "Word may lack a readable number."
        ),
    }
