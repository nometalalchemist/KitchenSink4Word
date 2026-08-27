"""Template and brand compliance checking against a user-supplied ruleset.

The ruleset is a plain dict a user can write directly from a university
formatting guide or a style manual (see check_template_compliance for the
schema). Every rule key is optional; unknown keys are rejected so a typo in
the ruleset fails loudly instead of silently checking nothing.

Everything here is read-only OOXML inspection: sectPr for page geometry,
styles.xml (docDefaults + the basedOn chain) for effective fonts, sizes and
spacing, pgNumType for page-number formats, and the heading outline.

Honesty policy (refuse-rather-than-guess): values this module cannot resolve
from the package alone — theme font indirection (w:asciiTheme, "+mn-lt"-style
references into theme1.xml), runs with no font defined anywhere — are reported
in the result's "unverified" list, never silently passed or failed.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from .furniture import _PAGE_NUM_FORMATS, _sect_prs, list_sections
from .read import (
    _outline_level,
    _style_outline_map,
    body_items,
    paragraph_text,
    run_text,
)

# Portrait-normalized page sizes in points (Word's letter = 12240x15840 twips,
# A4 = 11906x16838 twips).
_PAGE_SIZES_PT = {"letter": (612.0, 792.0), "a4": (595.3, 841.9)}

# Font/spacing/color checks sample the first N non-empty body paragraphs; the
# result notes when a longer document was truncated.
_SAMPLE_LIMIT = 500

# Per-rule violation cap so a systemic problem (every paragraph in Calibri)
# yields a readable report instead of thousands of identical entries.
_MAX_PER_RULE = 20

_TEMPLATE_RULE_KEYS = {
    "page", "fonts", "line_spacing", "headings", "page_numbering",
    "required_headings_in_order",
}
_BRAND_RULE_KEYS = _TEMPLATE_RULE_KEYS | {"colors"}

# Fixed evaluation order so reports are stable regardless of dict order.
_CHECK_ORDER = [
    "page", "fonts", "line_spacing", "headings", "page_numbering",
    "required_headings_in_order", "colors",
]


def _check_subkeys(d: dict, allowed: set, what: str) -> None:
    if not isinstance(d, dict):
        raise WordMcpError(f"{what} must be a dict; allowed keys: {sorted(allowed)}")
    unknown = set(d) - allowed
    if unknown:
        raise WordMcpError(
            f"unknown {what} key(s) {sorted(unknown)}; allowed: {sorted(allowed)}"
        )


def _add(
    report: dict,
    rule: str,
    expected,
    found,
    location,
    severity: str = "error",
) -> None:
    counts = report["_counts"]
    counts[rule] = counts.get(rule, 0) + 1
    if counts[rule] > _MAX_PER_RULE:
        return
    report["violations"].append(
        {
            "rule": rule,
            "expected": expected,
            "found": found,
            "location": location,
            "severity": severity,
        }
    )


# ------------------------------------------------- effective-format resolution


class _FormatResolver:
    """Effective run/paragraph formatting: explicit properties first, then the
    character-style chain, then the paragraph-style chain (following basedOn),
    then styles.xml docDefaults. Theme font indirection is surfaced as
    ('theme', name) and never resolved into a concrete font name."""

    def __init__(self, pkg: DocxPackage):
        self._styles: dict[str, etree._Element] = {}
        self._default_para: str | None = None
        self._doc_rpr = None
        self._doc_ppr = None
        if pkg.has_part("word/styles.xml"):
            root = pkg.root("word/styles.xml")
            for st in root.findall(qn("w:style")):
                sid = st.get(qn("w:styleId"))
                if sid:
                    self._styles[sid] = st
                if (
                    st.get(qn("w:default")) in ("1", "true")
                    and st.get(qn("w:type")) == "paragraph"
                ):
                    self._default_para = sid
            dd = root.find(qn("w:docDefaults"))
            if dd is not None:
                self._doc_rpr = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
                self._doc_ppr = dd.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}")

    def _chain(self, style_id: str | None):
        seen: set[str] = set()
        while style_id and style_id in self._styles and style_id not in seen:
            seen.add(style_id)
            st = self._styles[style_id]
            yield st
            based = st.find(qn("w:basedOn"))
            style_id = based.get(qn("w:val")) if based is not None else None

    def para_style_id(self, p: etree._Element) -> str | None:
        ps = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        return ps.get(qn("w:val")) if ps is not None else self._default_para

    @staticmethod
    def _font_of(rpr) -> tuple[str, str] | None:
        if rpr is None:
            return None
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            return None
        name = rfonts.get(qn("w:ascii"))
        if name:
            return ("name", name)
        theme = rfonts.get(qn("w:asciiTheme"))
        if theme:
            return ("theme", theme)
        return None

    @staticmethod
    def _size_of(rpr) -> float | None:
        if rpr is None:
            return None
        sz = rpr.find(qn("w:sz"))
        if sz is None:
            return None
        val = sz.get(qn("w:val"))
        return int(val) / 2 if val else None

    def _run_chain(self, p: etree._Element, r: etree._Element, getter):
        rpr = r.find(qn("w:rPr"))
        v = getter(rpr)
        if v is not None:
            return v
        if rpr is not None:
            rstyle = rpr.find(qn("w:rStyle"))
            if rstyle is not None:
                for st in self._chain(rstyle.get(qn("w:val"))):
                    v = getter(st.find(qn("w:rPr")))
                    if v is not None:
                        return v
        for st in self._chain(self.para_style_id(p)):
            v = getter(st.find(qn("w:rPr")))
            if v is not None:
                return v
        return getter(self._doc_rpr)

    def font(self, p, r) -> tuple[str, str] | None:
        """('name', font) | ('theme', themeRef) | None (nothing defined)."""
        return self._run_chain(p, r, self._font_of)

    def size_pt(self, p, r) -> float | None:
        return self._run_chain(p, r, self._size_of)

    @staticmethod
    def _spacing_of(ppr) -> tuple[int, str] | None:
        if ppr is None:
            return None
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            return None
        line = sp.get(qn("w:line"))
        if line is None:
            return None
        return (int(line), sp.get(qn("w:lineRule"), "auto"))

    def line_spacing(self, p) -> tuple[float, str]:
        """(value, rule). rule 'auto' means value is a multiple (1.0, 2.0);
        'exact'/'atLeast' mean value is in points. No spacing defined anywhere
        resolves to Word's baseline single spacing (1.0, 'auto')."""
        v = self._spacing_of(p.find(qn("w:pPr")))
        if v is None:
            for st in self._chain(self.para_style_id(p)):
                v = self._spacing_of(st.find(qn("w:pPr")))
                if v is not None:
                    break
        if v is None:
            v = self._spacing_of(self._doc_ppr)
        if v is None:
            return (1.0, "auto")
        line, rule = v
        if rule == "auto":
            return (round(line / 240, 2), "auto")
        return (round(line / 20, 1), rule)

    @staticmethod
    def run_color(r) -> tuple[str, str] | None:
        """EXPLICIT run color only: ('hex', 'RRGGBB') | ('theme', name) | None.
        Style-inherited colors are deliberately not resolved here."""
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            return None
        c = rpr.find(qn("w:color"))
        if c is None:
            return None
        theme = c.get(qn("w:themeColor"))
        if theme:
            return ("theme", theme)
        val = c.get(qn("w:val"))
        if not val or val.lower() == "auto":
            return None
        return ("hex", val.upper())


class _Ctx:
    """Lazily-built shared state so checkers never re-parse the same parts."""

    def __init__(self, pkg: DocxPackage):
        self.pkg = pkg
        self._resolver = None
        self._style_outline = None
        self._outline = None
        self._sample = None
        self.sample_truncated = False
        self.sample_used = False

    @property
    def resolver(self) -> _FormatResolver:
        if self._resolver is None:
            self._resolver = _FormatResolver(self.pkg)
        return self._resolver

    @property
    def style_outline(self) -> dict:
        if self._style_outline is None:
            self._style_outline = _style_outline_map(self.pkg)
        return self._style_outline

    @property
    def outline(self) -> list[dict]:
        if self._outline is None:
            from .read import get_outline

            self._outline = get_outline(self.pkg)
        return self._outline

    @property
    def sample_paragraphs(self) -> list[tuple[int, etree._Element]]:
        """(body index, element) for non-empty body paragraphs, capped at
        _SAMPLE_LIMIT."""
        self.sample_used = True
        if self._sample is None:
            out = []
            for kind, idx, el in body_items(self.pkg):
                if kind != "paragraph" or not paragraph_text(el).strip():
                    continue
                if len(out) >= _SAMPLE_LIMIT:
                    self.sample_truncated = True
                    break
                out.append((idx, el))
            self._sample = out
        return self._sample


# ----------------------------------------------------------------- rule checks


def _check_page(pkg, ctx, rule, report) -> None:
    _check_subkeys(
        rule, {"margins_pt", "tolerance_pt", "size", "orientation"}, "page rule"
    )
    tol = float(rule.get("tolerance_pt", 1))
    margins = rule.get("margins_pt")
    if margins is not None:
        _check_subkeys(
            margins, {"top", "bottom", "left", "right"}, "page.margins_pt"
        )
    size = rule.get("size")
    if size is not None and size not in _PAGE_SIZES_PT:
        raise WordMcpError(
            f"page.size must be one of {sorted(_PAGE_SIZES_PT)}, got {size!r}"
        )
    orient = rule.get("orientation")
    if orient is not None and orient not in ("portrait", "landscape"):
        raise WordMcpError("page.orientation must be portrait or landscape")

    for s in list_sections(pkg):
        loc = {"section": s["index"]}
        if margins is not None:
            found_m = s.get("margins_pt")
            if found_m is None:
                _add(report, "page.margins", margins,
                     "section has no explicit margins (w:pgMar missing)", loc)
            else:
                for side, exp in margins.items():
                    got = found_m.get(side, 0)
                    if abs(got - exp) > tol:
                        _add(report, f"page.margins.{side}", exp, got, loc)
        w, h = s.get("page_width_pt"), s.get("page_height_pt")
        if size is not None:
            if w is None or h is None:
                _add(report, "page.size", size,
                     "section has no explicit page size (w:pgSz missing)", loc)
            else:
                # Orientation-agnostic: a landscape letter page is still
                # "letter"; orientation is its own rule below.
                exp_dims = sorted(_PAGE_SIZES_PT[size])
                got_dims = sorted((w, h))
                if any(
                    abs(a - b) > max(tol, 1.0)
                    for a, b in zip(exp_dims, got_dims)
                ):
                    _add(report, "page.size", size,
                         {"width_pt": w, "height_pt": h}, loc)
        if orient is not None:
            if w is not None and h is not None:
                # Dimensions are the ground truth; the w:orient attribute can
                # be absent on hand-built landscape sections.
                found_o = "landscape" if w > h else "portrait"
            else:
                found_o = s.get("orientation", "portrait")
            if found_o != orient:
                _add(report, "page.orientation", orient, found_o, loc)


def _check_fonts(pkg, ctx, rule, report) -> None:
    _check_subkeys(rule, {"allowed", "body_size_pt"}, "fonts rule")
    allowed = rule.get("allowed")
    if allowed is not None and (
        not isinstance(allowed, list) or not all(isinstance(a, str) for a in allowed)
    ):
        raise WordMcpError("fonts.allowed must be a list of font names")
    allowed_cf = {a.casefold() for a in allowed} if allowed is not None else None
    body_size = rule.get("body_size_pt")

    theme_seen: dict[str, int] = {}
    unsized_at: int | None = None
    for idx, p in ctx.sample_paragraphs:
        is_heading = _outline_level(p, ctx.style_outline) is not None
        font_flagged = size_flagged = False
        for r in p.iter(qn("w:r")):
            if not run_text(r).strip():
                continue
            if allowed_cf is not None and not font_flagged:
                got = ctx.resolver.font(p, r)
                if got is None:
                    theme_seen.setdefault(
                        "(no font defined; application default)", idx
                    )
                elif got[0] == "theme":
                    theme_seen.setdefault(got[1], idx)
                elif got[1].casefold() not in allowed_cf:
                    _add(report, "fonts.allowed", allowed, got[1],
                         {"paragraph_index": idx})
                    font_flagged = True
            # Heading sizes legitimately differ; body_size_pt is body-only.
            if body_size is not None and not is_heading and not size_flagged:
                sz = ctx.resolver.size_pt(p, r)
                if sz is None:
                    unsized_at = idx if unsized_at is None else unsized_at
                elif abs(sz - body_size) > 0.01:
                    _add(report, "fonts.body_size_pt", body_size, sz,
                         {"paragraph_index": idx})
                    size_flagged = True
            done_font = allowed_cf is None or font_flagged
            done_size = body_size is None or is_heading or size_flagged
            if done_font and done_size:
                break
    for theme, idx in theme_seen.items():
        report["unverified"].append(
            {
                "rule": "fonts.allowed",
                "location": {"paragraph_index": idx},
                "reason": (
                    f"font resolves to theme indirection {theme!r}; theme "
                    "fonts are reported, not resolved — verify the theme "
                    "part (word/theme/theme1.xml) manually"
                ),
            }
        )
    if unsized_at is not None:
        report["unverified"].append(
            {
                "rule": "fonts.body_size_pt",
                "location": {"paragraph_index": unsized_at},
                "reason": (
                    "no font size defined in the run, its styles, or "
                    "docDefaults; the application default applies and cannot "
                    "be verified from the file"
                ),
            }
        )


def _check_line_spacing(pkg, ctx, rule, report) -> None:
    _check_subkeys(rule, {"body"}, "line_spacing rule")
    exp = rule.get("body")
    if exp is None:
        return
    for idx, p in ctx.sample_paragraphs:
        if _outline_level(p, ctx.style_outline) is not None:
            continue
        val, lr = ctx.resolver.line_spacing(p)
        if lr == "auto":
            if abs(val - exp) > 0.05:
                _add(report, "line_spacing.body", exp, val,
                     {"paragraph_index": idx})
        else:
            # exact/atLeast rules are point values, not multiples; a spacing
            # requirement stated as a multiple cannot be met by them.
            _add(report, "line_spacing.body", exp,
                 f"{val}pt ({lr} line rule, not a spacing multiple)",
                 {"paragraph_index": idx}, severity="warning")


def _check_headings(pkg, ctx, rule, report) -> None:
    _check_subkeys(rule, {"max_skip", "required_first_level"}, "headings rule")
    outline = ctx.outline
    if not outline:
        _add(report, "headings", rule, "document has no headings",
             {"paragraph_index": None})
        return
    rfl = rule.get("required_first_level")
    if rfl is not None and outline[0]["level"] != rfl:
        _add(report, "headings.required_first_level", rfl,
             outline[0]["level"],
             {"paragraph_index": outline[0]["paragraph_index"]})
    max_skip = rule.get("max_skip")
    if max_skip is not None:
        prev = outline[0]["level"]
        for h in outline[1:]:
            if h["level"] > prev + 1 + max_skip:
                _add(report, "headings.max_skip",
                     f"level <= {prev + 1 + max_skip} after level {prev}",
                     h["level"], {"paragraph_index": h["paragraph_index"]})
            prev = h["level"]


def _check_page_numbering(pkg, ctx, rule, report) -> None:
    if not isinstance(rule, list):
        raise WordMcpError(
            "page_numbering must be a list of "
            "{section, format, restart_at} entries"
        )
    sects = _sect_prs(pkg)
    for entry in rule:
        _check_subkeys(
            entry, {"section", "format", "restart_at"}, "page_numbering entry"
        )
        i = entry.get("section")
        if not isinstance(i, int):
            raise WordMcpError(
                "each page_numbering entry needs an integer 'section'"
            )
        fmt_exp = entry.get("format")
        if fmt_exp is not None and fmt_exp not in _PAGE_NUM_FORMATS:
            raise WordMcpError(
                f"page_numbering format must be one of "
                f"{sorted(_PAGE_NUM_FORMATS)}, got {fmt_exp!r}"
            )
        loc = {"section": i}
        if not 0 <= i < len(sects):
            _add(report, "page_numbering.section", f"section {i} exists",
                 f"document has {len(sects)} section(s)", loc)
            continue
        pgnum = sects[i].find(qn("w:pgNumType"))
        # OOXML default when w:fmt is absent is decimal.
        found_fmt = (
            pgnum.get(qn("w:fmt"), "decimal") if pgnum is not None else "decimal"
        )
        if fmt_exp is not None and found_fmt != fmt_exp:
            _add(report, "page_numbering.format", fmt_exp, found_fmt, loc)
        if "restart_at" in entry:
            start = pgnum.get(qn("w:start")) if pgnum is not None else None
            found = int(start) if start is not None else None
            if found != entry["restart_at"]:
                _add(report, "page_numbering.restart_at", entry["restart_at"],
                     found if found is not None
                     else "no restart (numbering continues)", loc)


def _check_required_headings(pkg, ctx, rule, report) -> None:
    if not isinstance(rule, list) or not all(isinstance(x, str) for x in rule):
        raise WordMcpError(
            "required_headings_in_order must be a list of heading strings"
        )
    texts = [(h["text"].strip().casefold(), h) for h in ctx.outline]
    pos = 0
    for want in rule:
        w = want.strip().casefold()
        hit = next((j for j in range(pos, len(texts)) if texts[j][0] == w), None)
        if hit is not None:
            pos = hit + 1
            continue
        anywhere = next((j for j in range(len(texts)) if texts[j][0] == w), None)
        if anywhere is None:
            _add(report, "required_headings_in_order", want,
                 "heading not found", {"paragraph_index": None})
        else:
            h = texts[anywhere][1]
            _add(report, "required_headings_in_order",
                 f"{want!r} after the previous required heading",
                 f"found earlier in the document (paragraph "
                 f"{h['paragraph_index']}), out of order",
                 {"paragraph_index": h["paragraph_index"]},
                 severity="warning")


def _check_colors(pkg, ctx, rule, report) -> None:
    _check_subkeys(rule, {"allowed_hex"}, "colors rule")
    allowed = rule.get("allowed_hex")
    if allowed is None:
        return
    if not isinstance(allowed, list) or not all(
        isinstance(c, str) for c in allowed
    ):
        raise WordMcpError("colors.allowed_hex must be a list of hex strings")
    allowed_set = {c.lstrip("#").upper() for c in allowed}
    theme_seen: dict[str, int] = {}
    for idx, p in ctx.sample_paragraphs:
        for r in p.iter(qn("w:r")):
            if not run_text(r).strip():
                continue
            got = _FormatResolver.run_color(r)
            if got is None:
                continue
            if got[0] == "theme":
                theme_seen.setdefault(got[1], idx)
                continue
            if got[1] not in allowed_set:
                _add(report, "colors.allowed_hex", sorted(allowed_set),
                     got[1], {"paragraph_index": idx})
                break  # one violation per paragraph keeps reports readable
    for theme, idx in theme_seen.items():
        report["unverified"].append(
            {
                "rule": "colors.allowed_hex",
                "location": {"paragraph_index": idx},
                "reason": (
                    f"run color is theme-indirected ({theme!r}); theme colors "
                    "are reported, not resolved — verify manually"
                ),
            }
        )


_CHECKERS = {
    "page": _check_page,
    "fonts": _check_fonts,
    "line_spacing": _check_line_spacing,
    "headings": _check_headings,
    "page_numbering": _check_page_numbering,
    "required_headings_in_order": _check_required_headings,
    "colors": _check_colors,
}


# ----------------------------------------------------------------- entry points


def _evaluate(pkg: DocxPackage, rules: dict, allowed: set, what: str) -> dict:
    if not isinstance(rules, dict):
        raise WordMcpError(
            f"{what} rules must be a dict; allowed keys: {sorted(allowed)}"
        )
    unknown = set(rules) - allowed
    if unknown:
        raise WordMcpError(
            f"unknown {what} rule key(s) {sorted(unknown)}; "
            f"allowed: {sorted(allowed)}"
        )
    ctx = _Ctx(pkg)
    report: dict = {
        "rules_checked": [k for k in _CHECK_ORDER if k in rules],
        "violations": [],
        "unverified": [],
        "_counts": {},
    }
    for key in report["rules_checked"]:
        _CHECKERS[key](pkg, ctx, rules[key], report)
    counts = report.pop("_counts")
    suppressed = {
        rule: n - _MAX_PER_RULE for rule, n in counts.items() if n > _MAX_PER_RULE
    }
    if suppressed:
        report["suppressed"] = suppressed
    if ctx.sample_truncated:
        report["unverified"].append(
            {
                "rule": "sampling",
                "location": None,
                "reason": (
                    f"font/spacing/color checks sampled only the first "
                    f"{_SAMPLE_LIMIT} non-empty body paragraphs; later "
                    "paragraphs were not examined"
                ),
            }
        )
    report["compliant"] = not report["violations"]
    report["violation_count"] = sum(counts.values())
    return report


def check_template_compliance(pkg: DocxPackage, rules: dict) -> dict:
    """Validate a document against a formatting ruleset (e.g. a university
    dissertation guide). Every key is optional; unknown keys are rejected.

    Ruleset schema::

        {"page": {"margins_pt": {"top": 72, "bottom": 72, "left": 90,
                                 "right": 72},
                  "tolerance_pt": 1,            # default 1
                  "size": "letter",             # "letter" | "a4"
                  "orientation": "portrait"},   # "portrait" | "landscape"
         "fonts": {"allowed": ["Times New Roman"],   # all text incl. headings
                   "body_size_pt": 12},              # non-heading text only
         "line_spacing": {"body": 2.0},              # multiple; non-heading
         "headings": {"max_skip": 0,            # extra levels a heading may
                                                # jump past prev + 1
                      "required_first_level": 1},
         "page_numbering": [{"section": 0, "format": "lowerRoman"},
                            {"section": 1, "format": "decimal",
                             "restart_at": 1}],
         "required_headings_in_order": ["Abstract", "Acknowledgments"]}

    Result: {compliant, violations: [{rule, expected, found, location,
    severity}], unverified: [...], rules_checked, violation_count}.

    Notes on scope and honesty:
    - Fonts/sizes/spacing are resolved from explicit run/paragraph properties
      plus the named-style basedOn chain and styles.xml docDefaults, on a
      sample of up to 500 non-empty body paragraphs.
    - Theme font indirection (w:asciiTheme / "+mn-lt") is reported in
      "unverified" as a theme reference, never resolved to a concrete name.
    - page.size is orientation-agnostic (a landscape letter page still counts
      as "letter"); use the orientation rule for orientation.
    - At most 20 violations per rule are listed; the overflow count appears
      under "suppressed".
    """
    return _evaluate(pkg, rules, _TEMPLATE_RULE_KEYS, "template-compliance")


def check_brand_compliance(pkg: DocxPackage, rules: dict) -> dict:
    """Brand-guide variant of check_template_compliance: the same engine and
    ruleset schema, plus one extra rule::

        {"colors": {"allowed_hex": ["1F4E79", "FF0000"]}}

    which checks EXPLICIT run color values (w:color) used in body text against
    the allowed palette. Hex values are compared case-insensitively, with or
    without a leading '#'. Theme-indirected colors are reported in
    "unverified"; colors inherited from styles are not resolved. See
    check_template_compliance for the shared schema and result format.
    """
    return _evaluate(pkg, rules, _BRAND_RULE_KEYS, "brand-compliance")
