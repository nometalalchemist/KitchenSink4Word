"""Text boxes and shape text frames: structured read and write access.

Word stores text-box content in ``w:txbxContent`` elements nested inside a
drawing (modern: ``mc:AlternateContent`` > ``mc:Choice`` > ``w:drawing`` >
``wps:txbx``; legacy: ``w:pict`` > ``v:textbox``), all buried inside a run of
some ordinary host paragraph.

Why a dedicated tool is needed — measured, not assumed:

- The generic read tools (get_text, find_text, the runmap that powers
  search_and_replace) iterate every descendant ``w:r`` of a paragraph, so box
  text is NOT invisible to them; it is SMEARED into the host paragraph's text
  with no boundary marker, and DUPLICATED — Word writes the same content twice
  (once in ``mc:Choice``, once in the ``mc:Fallback`` compatibility copy), so
  the generic tools report it doubled. Text that looks like
  ``"Host text.Box text.Box text."`` is a text box, not a corrupt paragraph.
- Because of that smearing, a replacement that matches across the host/box
  boundary (or across the Choice/Fallback copies) can splice text OUT of the
  box; and nothing in the generic output says which part of a paragraph lives
  inside a box. This module is the only per-box, structurally-addressed view.

Wire recommendation (for the integrating session; nothing here edits other
modules): find_text/get_text could gain an ``include_textboxes`` parameter,
and the runmap could skip ``w:txbxContent`` subtrees so box text stops
double-counting and stops being splice-able across the box boundary.
"""

from __future__ import annotations

import copy
import re

from lxml import etree

from ..core.errors import TargetNotFound, UnsupportedStructure
from ..core.package import DocxPackage, qn
from . import _runmap
from .read import body_items, paragraph_text

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_VML_NS = "urn:schemas-microsoft-com:vml"

_HEADER_FOOTER_RE = re.compile(r"word/(header|footer)\d+\.xml")


def _story_parts(pkg: DocxPackage) -> list[str]:
    parts = ["word/document.xml"]
    parts += [p for p in pkg.part_names() if _HEADER_FOOTER_RE.fullmatch(p)]
    return parts


def _in_fallback(el: etree._Element) -> bool:
    parent = el.getparent()
    while parent is not None:
        if parent.tag == f"{{{_MC_NS}}}Fallback":
            return True
        parent = parent.getparent()
    return False


def _shape_name(txbx_content: etree._Element) -> str | None:
    """Shape name from the nearest enclosing drawing's wp:docPr, or the
    legacy v:shape's alt/id."""
    parent = txbx_content.getparent()
    while parent is not None:
        tag = etree.QName(parent).localname
        if tag == "drawing":
            docpr = parent.find(f".//{qn('wp:docPr')}")
            if docpr is not None:
                return docpr.get("name")
            return None
        if parent.tag == f"{{{_VML_NS}}}shape":
            return parent.get("alt") or parent.get("id")
        parent = parent.getparent()
    return None


def _box_kind(txbx_content: etree._Element) -> str:
    parent = txbx_content.getparent()
    while parent is not None:
        tag = etree.QName(parent).localname
        if tag == "drawing":
            return "drawing"
        if tag in ("pict", "textbox"):
            return "vml"
        parent = parent.getparent()
    return "unknown"


def _host_paragraph(txbx_content: etree._Element) -> etree._Element | None:
    """The ordinary paragraph the box's run sits in (NOT a paragraph inside
    the box itself): first w:p ancestor strictly above the txbxContent."""
    parent = txbx_content.getparent()
    w_p = qn("w:p")
    while parent is not None:
        if parent.tag == w_p:
            return parent
        parent = parent.getparent()
    return None


def _enumerate_boxes(pkg: DocxPackage) -> list[dict]:
    """All primary (non-Fallback) txbxContent elements across body, headers,
    and footers, in document order per part. Fallback copies are the SAME
    content duplicated for old Word versions; they are tracked as twins, not
    listed as separate boxes."""
    boxes: list[dict] = []
    tag = qn("w:txbxContent")
    for part in _story_parts(pkg):
        if not pkg.has_part(part):
            continue
        if part == "word/document.xml":
            _keepalive = body_items(pkg)
            body_idx = {
                id(el): idx
                for kind, idx, el in _keepalive
                if kind == "paragraph"
            }
        else:
            body_idx = {}
        for el in pkg.root(part).iter(tag):
            if _in_fallback(el):
                continue
            host = _host_paragraph(el)
            anchor = body_idx.get(id(host)) if host is not None else None
            boxes.append(
                {
                    "element": el,
                    "part": part,
                    "kind": _box_kind(el),
                    "shape_name": _shape_name(el),
                    "anchor_paragraph_index": anchor,
                    "host": host,
                }
            )
    return boxes


def _fallback_twin(txbx_content: etree._Element) -> etree._Element | None:
    """The mc:Fallback copy of the same box, when one exists."""
    parent = txbx_content.getparent()
    while parent is not None:
        if parent.tag == f"{{{_MC_NS}}}AlternateContent":
            fb = parent.find(f"{{{_MC_NS}}}Fallback")
            if fb is not None:
                return fb.find(f".//{qn('w:txbxContent')}")
            return None
        parent = parent.getparent()
    return None


# ------------------------------------------------------------------- public


def get_textbox_text(pkg: DocxPackage) -> dict:
    """Text inside every text box / shape text frame, per box.

    Covers modern drawings (mc:AlternateContent > wps:txbx) and legacy VML
    (v:textbox) across the body, headers, and footers. Each box reports its
    text, its individual paragraphs, which part it lives in, the body index
    of the paragraph anchoring it (where determinable), and the shape name.

    Use THIS tool to read box content reliably: the generic read tools smear
    box text into the host paragraph's text without any boundary — and
    doubled, because Word stores each modern box twice (mc:Choice plus an
    mc:Fallback compatibility copy). A box anchored inside a table cell or a
    header has no body paragraph index (anchor_paragraph_index null).
    Read-only — the file is not modified."""
    boxes = _enumerate_boxes(pkg)
    out = []
    for i, b in enumerate(boxes):
        paras = [
            paragraph_text(p) for p in b["element"].findall(qn("w:p"))
        ]
        out.append(
            {
                "box_index": i,
                "part": b["part"],
                "kind": b["kind"],
                "shape_name": b["shape_name"],
                "anchor_paragraph_index": b["anchor_paragraph_index"],
                "paragraphs": paras,
                "text": "\n".join(paras),
            }
        )
    return {
        "count": len(out),
        "boxes": out,
        "note": (
            "generic read tools (get_text/find_text) exclude text-box "
            "content entirely — this tool is the one sanctioned reader; "
            "box_index above is the stable address for set_textbox_text"
        ),
    }


_NON_TEXT_CONTENT = ("w:tbl", "w:drawing", "w:pict", "a:blip")


def set_textbox_text(pkg: DocxPackage, box_index: int, text: str) -> dict:
    """Replace the text of one text box (box_index from get_textbox_text).

    The first paragraph's style and the first run's character formatting are
    kept and applied to the new content; '\\n' in `text` splits it into
    multiple paragraphs. When the box has an mc:Fallback compatibility copy,
    the copy is rewritten identically so old Word versions and the fallback
    path show the same text.

    Boxes holding non-text content (a nested table, an image, or another
    drawing) are refused rather than silently flattened — take the content
    out in Word first if that is really what you want."""
    boxes = _enumerate_boxes(pkg)
    if not 0 <= box_index < len(boxes):
        raise TargetNotFound(
            f"no text box with index {box_index}; document has "
            f"{len(boxes)} box(es) — run get_textbox_text for the list"
        )
    box = boxes[box_index]
    content = box["element"]

    for spec in _NON_TEXT_CONTENT:
        found = content.find(f".//{qn(spec)}")
        if found is not None:
            raise UnsupportedStructure(
                f"text box {box_index} ({box['shape_name'] or 'unnamed'}) "
                f"holds non-text content (<{spec}>); refusing to flatten it "
                "— edit that box in Word"
            )

    # Formatting to carry over: first paragraph's pPr, first run's rPr.
    first_p = content.find(qn("w:p"))
    ppr = rpr = None
    if first_p is not None:
        src_ppr = first_p.find(qn("w:pPr"))
        if src_ppr is not None:
            ppr = copy.deepcopy(src_ppr)
        first_r = first_p.find(qn("w:r"))
        if first_r is not None:
            src_rpr = first_r.find(qn("w:rPr"))
            if src_rpr is not None:
                rpr = copy.deepcopy(src_rpr)

    def build_content(target: etree._Element) -> int:
        for child in list(target):
            target.remove(child)
        lines = text.split("\n")
        for line in lines:
            p = etree.SubElement(target, qn("w:p"))
            if ppr is not None:
                p.append(copy.deepcopy(ppr))
            r = etree.SubElement(p, qn("w:r"))
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            t = etree.SubElement(r, qn("w:t"))
            t.text = line
            _runmap._preserve_space(t)
        return len(lines)

    n_paras = build_content(content)
    twin = _fallback_twin(content)
    if twin is not None:
        build_content(twin)
    pkg.mark_dirty(box["part"])
    return {
        "box_index": box_index,
        "part": box["part"],
        "shape_name": box["shape_name"],
        "paragraphs_written": n_paras,
        "fallback_copy_updated": twin is not None,
    }
