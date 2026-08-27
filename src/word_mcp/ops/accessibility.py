"""Accessibility audit and image-resolution checking (file-based OOXML only).

audit_accessibility covers the checks that can be decided honestly from the
package alone: heading hierarchy, image alt text (wp:docPr descr), table
header rows (w:tblHeader), explicit-color contrast, document title, and
generic hyperlink text. Anything requiring rendering or theme resolution is
out of scope and is skipped rather than guessed.

check_image_resolution derives effective DPI from the image part's native
pixel dimensions (PNG IHDR / JPEG SOF / GIF header, parsed by ops.media)
versus the displayed size in EMU.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from .media import _A, _R_NS, _WP, EMU_PER_INCH, _image_size_px
from .read import (
    _outline_level,
    _style_outline_map,
    body_items,
    paragraph_text,
    run_text,
)

# Word's fixed highlight palette (ST_HighlightColor -> sRGB hex).
_HIGHLIGHT_HEX = {
    "yellow": "FFFF00", "green": "00FF00", "cyan": "00FFFF",
    "magenta": "FF00FF", "blue": "0000FF", "red": "FF0000",
    "darkBlue": "000080", "darkCyan": "008080", "darkGreen": "008000",
    "darkMagenta": "800080", "darkRed": "800000", "darkYellow": "808000",
    "darkGray": "808080", "lightGray": "C0C0C0", "black": "000000",
    "white": "FFFFFF",
}

_GENERIC_LINK_TEXT = {"click here", "here", "link"}

# WCAG AA threshold for normal text.
_MIN_CONTRAST = 4.5


def _is_hex6(s: str | None) -> bool:
    if not s or len(s) != 6:
        return False
    try:
        int(s, 16)
        return True
    except ValueError:
        return False


def _luminance(hex6: str) -> float:
    def lin(c: int) -> float:
        v = c / 255
        return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hex6[i : i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def _contrast_ratio(a: str, b: str) -> float:
    la, lb = _luminance(a), _luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def _image_containers(pkg: DocxPackage):
    """(index, drawing container, media target) for every a:blip in
    document.xml, in document order — indices match list_images."""
    rid_target: dict = {}
    if pkg.has_part("word/_rels/document.xml.rels"):
        rid_target = {
            r.get("Id"): r.get("Target")
            for r in pkg.root("word/_rels/document.xml.rels")
        }
    for i, blip in enumerate(pkg.root().iter(f"{{{_A}}}blip")):
        container = blip.getparent()
        while container is not None and not (
            container.tag.endswith("}inline") or container.tag.endswith("}anchor")
        ):
            container = container.getparent()
        target = rid_target.get(blip.get(f"{{{_R_NS}}}embed"))
        yield i, container, target


# ------------------------------------------------------------------- the audit


def audit_accessibility(pkg: DocxPackage) -> dict:
    """Accessibility findings by category, each with a location and a fix hint.

    Categories:
    - heading_hierarchy: skipped levels, empty headings, no Heading 1 (or no
      headings at all).
    - images_missing_alt_text: inline/anchored images whose wp:docPr has no
      (or an empty) descr attribute.
    - tables_without_header_row: body tables whose first row lacks
      w:tblHeader.
    - low_contrast_text: runs with BOTH an explicit color and an explicit
      run-level background (w:shd fill or w:highlight) whose WCAG contrast
      ratio is below 4.5. Runs where either side is absent, "auto", or
      theme-indirected are skipped, not guessed.
    - document_title_missing: docProps/core.xml dc:title absent or empty.
    - link_text_generic: hyperlink display text in {"click here", "here",
      "link"}.

    Result: {findings: {category: [...]}, summary: {counts, total, pass}}.
    """
    findings: dict[str, list] = {
        "heading_hierarchy": [],
        "images_missing_alt_text": [],
        "tables_without_header_row": [],
        "low_contrast_text": [],
        "document_title_missing": [],
        "link_text_generic": [],
    }

    # ---- heading hierarchy
    style_outline = _style_outline_map(pkg)
    headings: list[tuple[int, int, str]] = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph":
            continue
        lvl = _outline_level(el, style_outline)
        if lvl is not None:
            headings.append((idx, lvl, paragraph_text(el).strip()))
    if not headings:
        findings["heading_hierarchy"].append(
            {
                "issue": "no_headings",
                "location": None,
                "fix": "add Heading 1-9 styles so assistive technology can "
                       "navigate the document",
            }
        )
    else:
        if not any(lvl == 1 for _, lvl, _ in headings):
            findings["heading_hierarchy"].append(
                {
                    "issue": "no_heading_1",
                    "location": {"paragraph_index": headings[0][0]},
                    "fix": "the top-level heading should be Heading 1; screen "
                           "readers use it as the document's entry point",
                }
            )
        prev = headings[0][1]
        for idx, lvl, text in headings:
            if not text:
                findings["heading_hierarchy"].append(
                    {
                        "issue": "empty_heading",
                        "location": {"paragraph_index": idx},
                        "fix": "give the heading text or change its style to "
                               "a body style; empty headings confuse "
                               "screen-reader navigation",
                    }
                )
            if lvl > prev + 1:
                findings["heading_hierarchy"].append(
                    {
                        "issue": "skipped_level",
                        "location": {"paragraph_index": idx},
                        "found": f"Heading {lvl} follows Heading {prev}",
                        "fix": f"use Heading {prev + 1} (or restructure) so "
                               "levels increase one at a time",
                    }
                )
            prev = lvl

    # ---- images missing alt text
    for i, container, target in _image_containers(pkg):
        if container is None:
            continue
        docpr = container.find(f"{{{_WP}}}docPr")
        descr = docpr.get("descr") if docpr is not None else None
        if not (descr or "").strip():
            findings["images_missing_alt_text"].append(
                {
                    "location": {"image_index": i},
                    "target": target,
                    "fix": "add a description with set_image_alt_text "
                           "(decorative images should say so explicitly)",
                }
            )

    # ---- tables without a header row
    for kind, idx, el in body_items(pkg):
        if kind != "table":
            continue
        first_tr = el.find(qn("w:tr"))
        if first_tr is None:
            continue
        trpr = first_tr.find(qn("w:trPr"))
        has_header = trpr is not None and trpr.find(qn("w:tblHeader")) is not None
        if not has_header:
            findings["tables_without_header_row"].append(
                {
                    "location": {"table_index": idx},
                    "rows": len(el.findall(qn("w:tr"))),
                    "fix": "mark the first row as a repeating header with "
                           "set_header_row_repeat so screen readers announce "
                           "column context",
                }
            )

    # ---- low-contrast text (explicit color vs explicit run background only)
    body_idx = {
        id(el): idx for kind, idx, el in body_items(pkg) if kind == "paragraph"
    }
    for p in pkg.root().iter(qn("w:p")):
        loc: dict = (
            {"paragraph_index": body_idx[id(p)]}
            if id(p) in body_idx
            else {"paragraph_index": None, "container": "nested (table cell "
                  "or text box)"}
        )
        for r in p.iter(qn("w:r")):
            if not run_text(r).strip():
                continue
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            c = rpr.find(qn("w:color"))
            if c is None or c.get(qn("w:themeColor")):
                continue
            fg = (c.get(qn("w:val")) or "").upper()
            if not _is_hex6(fg):
                continue  # "auto" or malformed: skip, do not guess
            bg = None
            shd = rpr.find(qn("w:shd"))
            if shd is not None:
                fill = (shd.get(qn("w:fill")) or "").upper()
                if _is_hex6(fill):
                    bg = fill
            if bg is None:
                hl = rpr.find(qn("w:highlight"))
                if hl is not None:
                    bg = _HIGHLIGHT_HEX.get(hl.get(qn("w:val")))
            if bg is None:
                continue  # no explicit background: skip, do not guess
            ratio = _contrast_ratio(fg, bg)
            if ratio < _MIN_CONTRAST:
                snippet = run_text(r).strip()
                findings["low_contrast_text"].append(
                    {
                        "location": loc,
                        "text": snippet[:60],
                        "color": fg,
                        "background": bg,
                        "contrast_ratio": round(ratio, 2),
                        "fix": f"contrast {round(ratio, 2)}:1 is below the "
                               f"WCAG AA minimum of {_MIN_CONTRAST}:1; darken "
                               "the text or lighten the background",
                    }
                )
                break  # one finding per paragraph keeps the report readable

    # ---- document title
    title = None
    if pkg.has_part("docProps/core.xml"):
        el = pkg.root("docProps/core.xml").find(
            "{http://purl.org/dc/elements/1.1/}title"
        )
        if el is not None:
            title = el.text
    if not (title or "").strip():
        findings["document_title_missing"].append(
            {
                "location": {"part": "docProps/core.xml"},
                "fix": "set a title with set_document_properties; screen "
                       "readers announce it instead of the filename",
            }
        )

    # ---- generic hyperlink text
    for link in pkg.root().iter(qn("w:hyperlink")):
        text = "".join(run_text(r) for r in link.iter(qn("w:r"))).strip()
        if text.lower() not in _GENERIC_LINK_TEXT:
            continue
        p = link.getparent()
        while p is not None and etree.QName(p).localname != "p":
            p = p.getparent()
        loc = (
            {"paragraph_index": body_idx.get(id(p))}
            if p is not None
            else {"paragraph_index": None}
        )
        findings["link_text_generic"].append(
            {
                "location": loc,
                "text": text,
                "fix": "rewrite the link text to describe the destination "
                       "(e.g. 'the 2026 annual report'), not the action",
            }
        )

    counts = {cat: len(items) for cat, items in findings.items()}
    total = sum(counts.values())
    return {
        "findings": findings,
        "summary": {"counts": counts, "total": total, "pass": total == 0},
    }


# --------------------------------------------------------------- image DPI


def check_image_resolution(pkg: DocxPackage, min_dpi: int = 300) -> dict:
    """Effective print resolution of every image in the document body.

    For each image: native pixel dimensions (parsed from the PNG/JPEG/GIF part
    bytes) divided by the displayed size (wp:extent, 914400 EMU per inch)
    gives horizontal and vertical DPI. Images below min_dpi are flagged with
    the actual numbers. EMF/WMF are vector formats and are reported as
    "vector (not applicable)"; other formats (BMP, TIFF, ...) are reported as
    "unchecked (<format>)" rather than guessed.

    Result: {min_dpi, images: [{image_index, target, status, pixels?,
    display_in?, dpi?}], low_resolution: [indices], pass}.
    """
    if not isinstance(min_dpi, int) or min_dpi <= 0:
        raise WordMcpError("min_dpi must be a positive integer")
    images: list[dict] = []
    low: list[int] = []
    for i, container, target in _image_containers(pkg):
        entry: dict = {"image_index": i, "target": target}
        images.append(entry)
        if target is None:
            entry["status"] = "unchecked (no relationship target)"
            continue
        ext = (
            "." + target.rsplit(".", 1)[1].lower() if "." in target else ""
        )
        if ext in (".emf", ".wmf"):
            entry["status"] = "vector (not applicable)"
            continue
        if container is None:
            entry["status"] = "unchecked (no inline/anchor container)"
            continue
        extent = container.find(f"{{{_WP}}}extent")
        if extent is None:
            entry["status"] = "unchecked (no display extent)"
            continue
        cx, cy = int(extent.get("cx", "0")), int(extent.get("cy", "0"))
        if cx <= 0 or cy <= 0:
            entry["status"] = "unchecked (zero display extent)"
            continue
        entry["display_in"] = [
            round(cx / EMU_PER_INCH, 2),
            round(cy / EMU_PER_INCH, 2),
        ]
        if ext not in (".png", ".jpg", ".jpeg", ".gif"):
            entry["status"] = f"unchecked ({ext.lstrip('.') or 'unknown format'})"
            continue
        part = "word/" + target.lstrip("/")
        if not pkg.has_part(part):
            entry["status"] = "unchecked (media part missing)"
            continue
        px = _image_size_px(pkg.raw_part(part), ext)
        if px is None:
            entry["status"] = (
                f"unchecked (could not parse {ext.lstrip('.')} dimensions)"
            )
            continue
        dpi_h = px[0] * EMU_PER_INCH / cx
        dpi_v = px[1] * EMU_PER_INCH / cy
        entry["pixels"] = [px[0], px[1]]
        entry["dpi"] = [round(dpi_h, 1), round(dpi_v, 1)]
        if min(dpi_h, dpi_v) < min_dpi:
            entry["status"] = "low"
            entry["fix"] = (
                f"effective resolution {round(min(dpi_h, dpi_v), 1)} DPI is "
                f"below {min_dpi}; use a larger source image or shrink the "
                "displayed size"
            )
            low.append(i)
        else:
            entry["status"] = "ok"
    return {
        "min_dpi": min_dpi,
        "images": images,
        "low_resolution": low,
        "pass": not low,
    }
