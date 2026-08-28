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
    _HEADING_STYLES,
    _outline_level,
    _style_id,
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
    # _keepalive holds the body elements while body_idx maps their id()s:
    # lxml proxies are ephemeral, and a recycled proxy address would report
    # the wrong paragraph_index. The map is used again in the hyperlink
    # section below, so the list must outlive both.
    _keepalive = [
        (el, idx) for kind, idx, el in body_items(pkg) if kind == "paragraph"
    ]
    body_idx = {id(el): idx for el, idx in _keepalive}
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


# ------------------------------------------------------------------- the fixes
#
# fix_accessibility is the mutation twin of audit_accessibility. Detection is
# deliberately IDENTICAL to the audit's (same helpers, same conditions), so an
# audit -> fix -> audit round trip shows the opted-in findings resolved.
# Nothing runs unless its category is explicitly opted in, and every category
# reports fixed / skipped-with-reason / needs-human-review. Ambiguity is
# refused, never guessed.

_FIX_CATEGORIES = (
    "alt_text_placeholders",
    "heading_skips",
    "table_headers",
    "doc_title",
)

_HEADING_STRATEGIES = ("promote", "demote_following")


def _defined_style_ids(pkg: DocxPackage) -> set[str]:
    ids: set[str] = set()
    if pkg.has_part("word/styles.xml"):
        for style in pkg.root("word/styles.xml").findall(qn("w:style")):
            sid = style.get(qn("w:styleId"))
            if sid:
                ids.add(sid)
    return ids


def _looks_numeric(s: str) -> bool:
    """True when a cell's text reads as data, not a label: at least one digit
    and nothing outside digits/number punctuation/currency."""
    s = s.strip()
    if not s or not any(ch.isdigit() for ch in s):
        return False
    allowed = " \t.,%+-()/:$€£¥"
    return all(ch.isdigit() or ch in allowed for ch in s)


# ---- category: alt text placeholders


def _fix_alt_text(pkg: DocxPackage, dry_run: bool) -> dict:
    fixed: list[dict] = []
    skipped: list[dict] = []
    review: list[dict] = []
    for i, container, target in _image_containers(pkg):
        if container is None:
            skipped.append(
                {
                    "image_index": i,
                    "target": target,
                    "reason": "no inline/anchor drawing container; cannot "
                              "reach docPr to set alt text",
                }
            )
            continue
        docpr = container.find(f"{{{_WP}}}docPr")
        if docpr is None:
            skipped.append(
                {
                    "image_index": i,
                    "target": target,
                    "reason": "drawing has no docPr element",
                }
            )
            continue
        if (docpr.get("descr") or "").strip():
            continue  # already has alt text; the audit does not flag it
        hint = target.rsplit("/", 1)[-1] if target else f"image {i}"
        placeholder = f"IMAGE: needs description: {hint}"
        if not dry_run:
            docpr.set("descr", placeholder)
        fixed.append(
            {"image_index": i, "target": target, "placeholder": placeholder}
        )
        review.append(
            {
                "image_index": i,
                "target": target,
                "reason": "placeholder alt text only; no description was "
                          "generated. Replace it with a real description via "
                          "set_image_alt_text (or mark the image decorative).",
            }
        )
    if fixed and not dry_run:
        pkg.mark_dirty()
    return {"fixed": fixed, "skipped": skipped, "needs_human_review": review}


# ---- category: heading skips


def _builtin_target_style(tgt: int, defined: set[str]) -> str | None:
    for cand in (f"Heading{tgt}", f"heading {tgt}"):
        if cand in defined:
            return cand
    return None


def _heading_change_blocker(
    p, tgt: int, defined: set[str], style_outline: dict[str, int]
) -> str | None:
    """Reason this paragraph's heading level cannot be changed safely, or None."""
    sid = _style_id(p)
    direct = p.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
    if sid and sid in _HEADING_STYLES:
        if _builtin_target_style(tgt, defined) is None:
            return (
                f"target style Heading{tgt} is not defined in this "
                "document's styles.xml"
            )
        return None
    if direct is not None:
        return None  # paragraph-level outlineLvl can be adjusted in place
    if sid and sid in style_outline:
        return (
            f"heading level comes from custom style '{sid}'; changing the "
            "style would change every paragraph using it, and a "
            "paragraph-level override would diverge from the style's visual "
            "formatting"
        )
    return "could not determine how this paragraph gets its heading level"


def _apply_heading_level(p, tgt: int, defined: set[str]) -> None:
    sid = _style_id(p)
    direct = p.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
    if sid and sid in _HEADING_STYLES:
        pstyle = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        pstyle.set(qn("w:val"), _builtin_target_style(tgt, defined))
    if direct is not None:
        direct.set(qn("w:val"), str(tgt - 1))


def _fix_heading_skips(pkg: DocxPackage, strategy: str, dry_run: bool) -> dict:
    if strategy not in _HEADING_STRATEGIES:
        raise WordMcpError(
            "heading_strategy must be 'promote' or 'demote_following', "
            f"got {strategy!r}"
        )
    style_outline = _style_outline_map(pkg)
    defined = _defined_style_ids(pkg)
    # _items keeps the body elements alive: lxml proxies are ephemeral, and
    # headings[] below holds references into this list for later mutation.
    _items = body_items(pkg)
    headings: list[tuple[int, int, object]] = []
    for kind, idx, el in _items:
        if kind != "paragraph":
            continue
        lvl = _outline_level(el, style_outline)
        if lvl is not None:
            headings.append((idx, lvl, el))
    empty = {"fixed": [], "skipped": [], "needs_human_review": []}
    if not headings:
        return empty
    orig = [lvl for _, lvl, _ in headings]
    n = len(orig)
    plans: dict[int, int] = {}  # heading position -> new level
    ambiguous: list[str] = []

    if strategy == "promote":
        # Pull each skipping heading (and its contiguous deeper block) up by
        # the gap, exactly closing to prev+1. A skip detected INSIDE an
        # already-shifted block is a nested/mixed pattern: refuse, don't guess.
        sim = list(orig)
        shifted: set[int] = set()
        prev = sim[0]
        for j in range(n):
            cur = sim[j]
            if cur > prev + 1:
                if j in shifted:
                    ambiguous.append(
                        f"nested skip pattern at paragraph "
                        f"{headings[j][0]}: promoting the enclosing block "
                        f"still leaves Heading {cur} after Heading {prev}"
                    )
                    break
                gap = cur - (prev + 1)
                k = j
                while k < n and sim[k] >= cur:
                    sim[k] -= gap
                    plans[k] = sim[k]
                    shifted.add(k)
                    k += 1
            prev = sim[j]
    else:  # demote_following
        # Read the skip as the PRECEDING heading being too shallow: demote it
        # so the following deep heading fits under it. Refuse on conflicting
        # demotions; the simulation below refuses anything that still skips.
        prev = orig[0]
        for j in range(n):
            cur = orig[j]
            if cur > prev + 1:
                pos, tgt = j - 1, cur - 1
                if pos in plans and plans[pos] != tgt:
                    ambiguous.append(
                        f"conflicting demotions for paragraph "
                        f"{headings[pos][0]} (would need Heading {plans[pos]} "
                        f"and Heading {tgt} at once)"
                    )
                    break
                plans[pos] = tgt
            prev = cur

    if not ambiguous and plans:
        # Verify with the audit's own walk: the repaired sequence must have no
        # skips left, and must not lose the document's only Heading 1.
        sim = [plans.get(j, orig[j]) for j in range(n)]
        prev = sim[0]
        for j in range(n):
            if sim[j] > prev + 1:
                ambiguous.append(
                    f"repair would leave Heading {sim[j]} after Heading "
                    f"{prev} at paragraph {headings[j][0]}"
                )
            prev = sim[j]
        if any(l == 1 for l in orig) and not any(l == 1 for l in sim):
            ambiguous.append("repair would remove every Heading 1")

    if ambiguous:
        return {
            "fixed": [],
            "skipped": [],
            "needs_human_review": [
                {
                    "reason": "heading structure is ambiguous; nothing was "
                              "changed. Restructure manually (apply_style or "
                              "add_heading) or try the other strategy.",
                    "strategy": strategy,
                    "details": ambiguous,
                }
            ],
        }
    if not plans:
        return empty  # no skipped levels to repair

    problems = [
        {"paragraph_index": headings[j][0], "reason": blocker}
        for j, tgt in sorted(plans.items())
        if (blocker := _heading_change_blocker(
            headings[j][2], tgt, defined, style_outline
        ))
    ]
    if problems:
        return {
            "fixed": [],
            "skipped": [],
            "needs_human_review": [
                {
                    "reason": "some headings cannot be changed safely; "
                              "nothing was changed (a partial repair would "
                              "corrupt the hierarchy)",
                    "strategy": strategy,
                    "items": problems,
                }
            ],
        }

    fixed: list[dict] = []
    for j in sorted(plans):
        idx, lvl, el = headings[j]
        if not dry_run:
            _apply_heading_level(el, plans[j], defined)
        fixed.append(
            {
                "paragraph_index": idx,
                "text": paragraph_text(el).strip()[:60],
                "from_level": lvl,
                "to_level": plans[j],
            }
        )
    if fixed and not dry_run:
        pkg.mark_dirty()
    return {"fixed": fixed, "skipped": [], "needs_human_review": []}


# ---- category: table header rows


def _fix_table_headers(pkg: DocxPackage, dry_run: bool) -> dict:
    fixed: list[dict] = []
    skipped: list[dict] = []
    # local list keeps the table elements alive while we mutate them
    for kind, idx, el in body_items(pkg):
        if kind != "table":
            continue
        trs = el.findall(qn("w:tr"))
        if not trs:
            continue
        first_tr = trs[0]
        trpr = first_tr.find(qn("w:trPr"))
        if trpr is not None and trpr.find(qn("w:tblHeader")) is not None:
            continue  # already a header row; the audit does not flag it
        cells = [
            "".join(paragraph_text(p) for p in tc.findall(qn("w:p"))).strip()
            for tc in first_tr.findall(qn("w:tc"))
        ]
        nonempty = [c for c in cells if c]
        if len(trs) == 1:
            skipped.append(
                {
                    "table_index": idx,
                    "first_row": cells,
                    "reason": "single-row table; marking its only row as a "
                              "header needs human judgment",
                }
            )
            continue
        if not nonempty:
            skipped.append(
                {
                    "table_index": idx,
                    "first_row": cells,
                    "reason": "first row is empty; an empty header row would "
                              "not help assistive technology. Add real column "
                              "labels first.",
                }
            )
            continue
        if all(_looks_numeric(c) for c in nonempty):
            skipped.append(
                {
                    "table_index": idx,
                    "first_row": cells,
                    "reason": "first row is all-numeric and looks like data, "
                              "not column labels; not marked. Add a real "
                              "header row, then set_header_row_repeat.",
                }
            )
            continue
        if not dry_run:
            if trpr is None:
                trpr = etree.Element(qn("w:trPr"))
                first_tr.insert(0, trpr)
            etree.SubElement(trpr, qn("w:tblHeader"))
        fixed.append({"table_index": idx, "first_row": cells})
    if fixed and not dry_run:
        pkg.mark_dirty()
    return {"fixed": fixed, "skipped": skipped, "needs_human_review": []}


# ---- category: document title


def _fix_doc_title(pkg: DocxPackage, dry_run: bool) -> dict:
    empty = {"fixed": [], "skipped": [], "needs_human_review": []}
    title = None
    if pkg.has_part("docProps/core.xml"):
        el = pkg.root("docProps/core.xml").find(
            "{http://purl.org/dc/elements/1.1/}title"
        )
        if el is not None:
            title = el.text
    if (title or "").strip():
        return {
            **empty,
            "skipped": [
                {
                    "reason": "document already has a title; existing titles "
                              "are never overwritten",
                    "title": title.strip(),
                }
            ],
        }
    style_outline = _style_outline_map(pkg)
    first_h1: tuple[int, str] | None = None
    for kind, idx, el in body_items(pkg):
        if kind == "paragraph" and _outline_level(el, style_outline) == 1:
            first_h1 = (idx, paragraph_text(el).strip())
            break
    if first_h1 is None:
        return {
            **empty,
            "needs_human_review": [
                {
                    "reason": "no Heading 1 to take the title from; set a "
                              "title with set_document_properties",
                }
            ],
        }
    idx, text = first_h1
    if not text:
        return {
            **empty,
            "needs_human_review": [
                {
                    "reason": "the first Heading 1 is empty; set a title "
                              "with set_document_properties",
                    "source": {"paragraph_index": idx},
                }
            ],
        }
    if not dry_run:
        from .structure import set_document_properties

        set_document_properties(pkg, title=text)
    return {
        **empty,
        "fixed": [{"title": text, "source": {"paragraph_index": idx}}],
    }


# ---- the tool entry point


def fix_accessibility(
    pkg: DocxPackage,
    *,
    alt_text_placeholders: bool = False,
    heading_skips: bool = False,
    heading_strategy: str = "promote",
    table_headers: bool = False,
    doc_title: bool = False,
    dry_run: bool = False,
) -> dict:
    """Repair audit_accessibility findings, one opted-in category at a time.

    Detection mirrors the audit exactly, so audit -> fix -> audit shows the
    opted-in categories resolved. Categories not opted in are never touched
    and do not appear in the result. With no category selected this raises
    rather than guessing an intent.

    - alt_text_placeholders: images with no alt text get a clearly-marked
      placeholder ("IMAGE: needs description: <filename>"); every touched
      image is also listed under needs_human_review. No descriptions are
      invented.
    - heading_skips: skipped levels repaired via heading_strategy 'promote'
      (pull the deep heading and its contiguous deeper block up to close the
      gap) or 'demote_following' (demote the preceding shallow heading so the
      following deep one fits). Nested/mixed patterns, custom-style headings,
      and any repair that would leave a skip or remove the only Heading 1 are
      refused with details, nothing half-applied.
    - table_headers: sets w:tblHeader on first rows the audit flags, skipping
      single-row tables and first rows that are empty or all-numeric (visibly
      data), with reasons.
    - doc_title: sets the core title from the first Heading 1 only when the
      title is empty; an existing title is never overwritten.

    dry_run=True computes the identical report without changing anything.
    Result: {dry_run, categories: {name: {fixed, skipped,
    needs_human_review}}, summary}.
    """
    if not (alt_text_placeholders or heading_skips or table_headers or doc_title):
        raise WordMcpError(
            "no fix categories selected; pass at least one of "
            + ", ".join(_FIX_CATEGORIES)
            + " (each category runs only when explicitly opted in)"
        )
    categories: dict[str, dict] = {}
    if alt_text_placeholders:
        categories["alt_text_placeholders"] = _fix_alt_text(pkg, dry_run)
    if heading_skips:
        categories["heading_skips"] = _fix_heading_skips(
            pkg, heading_strategy, dry_run
        )
    if table_headers:
        categories["table_headers"] = _fix_table_headers(pkg, dry_run)
    if doc_title:
        categories["doc_title"] = _fix_doc_title(pkg, dry_run)
    summary = {
        key: sum(len(cat[key]) for cat in categories.values())
        for key in ("fixed", "skipped", "needs_human_review")
    }
    summary["applied"] = not dry_run
    return {"dry_run": dry_run, "categories": categories, "summary": summary}
