"""Style-aware find: locate text by its EFFECTIVE formatting.

find_formatted() answers "where is everything bold?", "find 'delta' but only
where it is italic", "which runs are 12.5pt Garamond?" — questions plain text
search cannot answer. Formatting is resolved the way Word resolves it:
explicit run properties first, then the character-style chain, then the
paragraph-style chain (following basedOn), then styles.xml docDefaults. Each
match reports WHICH level satisfied each criterion, so explicit bold and
Heading-inherited bold are distinguishable.

Honesty policy: theme font/color indirection (w:asciiTheme, w:themeColor)
cannot be resolved to a concrete value from the package alone. Runs whose
effective value is a theme reference are never silently matched or rejected —
they are counted in the result's `unresolved_theme_runs` note.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap
from .compliance import _FormatResolver
from .read import body_items

_ALLOWED_KEYS = {
    "bold", "italic", "underline", "strike",
    "font", "size_pt", "color", "highlight", "style",
}
_TOGGLE_TAGS = {
    "bold": "w:b",
    "italic": "w:i",
    "underline": "w:u",
    "strike": "w:strike",
}
_SCOPES = {
    "body": ("word/document.xml",),
    "all": ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"),
}
_MAX_MATCHES = 500


# ------------------------------------------------------------------ getters


def _toggle_getter(tag: str):
    def get(rpr):
        if rpr is None:
            return None
        el = rpr.find(qn(tag))
        if el is None:
            return None
        val = el.get(qn("w:val"))
        if val is None:
            return True
        if tag == "w:u":
            return val != "none"
        return val not in ("0", "false", "none")

    return get


def _color_of(rpr):
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


def _highlight_of(rpr):
    if rpr is None:
        return None
    h = rpr.find(qn("w:highlight"))
    if h is None:
        return None
    return h.get(qn("w:val")) or None


class _Resolver(_FormatResolver):
    """_FormatResolver plus source attribution: every lookup reports whether
    the value came from the run itself, a character style, a paragraph style,
    or the document defaults."""

    def resolve(self, p, r, getter):
        rpr = r.find(qn("w:rPr"))
        v = getter(rpr)
        if v is not None:
            return v, "explicit"
        if rpr is not None:
            rstyle = rpr.find(qn("w:rStyle"))
            if rstyle is not None:
                for st in self._chain(rstyle.get(qn("w:val"))):
                    v = getter(st.find(qn("w:rPr")))
                    if v is not None:
                        return v, "character_style"
        for st in self._chain(self.para_style_id(p)):
            v = getter(st.find(qn("w:rPr")))
            if v is not None:
                return v, "paragraph_style"
        v = getter(self._doc_rpr)
        if v is not None:
            return v, "document_defaults"
        return None, None

    def style_names(self) -> dict[str, str]:
        out = {}
        for sid, st in self._styles.items():
            name = st.find(qn("w:name"))
            out[sid] = name.get(qn("w:val")) if name is not None else sid
        return out

    def run_char_style(self, r) -> str | None:
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            return None
        rstyle = rpr.find(qn("w:rStyle"))
        return rstyle.get(qn("w:val")) if rstyle is not None else None

    def resolve_para(self, p, getter):
        """Paragraph-level lookup with source attribution: the paragraph's own
        pPr first, then the paragraph-style chain (following basedOn), then
        styles.xml docDefaults. Returns (value, source) — (None, None) when
        the property is defined nowhere."""
        v = getter(p.find(qn("w:pPr")))
        if v is not None:
            return v, "explicit"
        for st in self._chain(self.para_style_id(p)):
            v = getter(st.find(qn("w:pPr")))
            if v is not None:
                return v, "paragraph_style"
        v = getter(self._doc_ppr)
        if v is not None:
            return v, "document_defaults"
        return None, None


# ----------------------------------------------------------------- matching


def _normalize(formatting: dict) -> dict:
    if not isinstance(formatting, dict) or not formatting:
        raise WordMcpError(
            "formatting must be a non-empty dict; "
            f"allowed keys: {sorted(_ALLOWED_KEYS)}"
        )
    unknown = set(formatting) - _ALLOWED_KEYS
    if unknown:
        raise WordMcpError(
            f"unknown formatting key(s) {sorted(unknown)}; "
            f"allowed: {sorted(_ALLOWED_KEYS)}"
        )
    crit = dict(formatting)
    for k in _TOGGLE_TAGS:
        if k in crit and not isinstance(crit[k], bool):
            raise WordMcpError(f"{k} must be true or false")
    if "size_pt" in crit and not isinstance(crit["size_pt"], (int, float)):
        raise WordMcpError("size_pt must be a number (points)")
    for k in ("font", "color", "highlight", "style"):
        if k in crit and not isinstance(crit[k], str):
            raise WordMcpError(f"{k} must be a string")
    if "color" in crit:
        crit["color"] = crit["color"].lstrip("#").upper()
    if "highlight" in crit:
        crit["highlight"] = crit["highlight"].lower()
    return crit


def _run_matches(
    resolver: _Resolver,
    style_names: dict[str, str],
    p,
    r,
    crit: dict,
    theme_counter: list[int],
) -> dict | None:
    """matched_via dict when the run satisfies every criterion, else None."""
    via: dict[str, str] = {}
    for key in ("bold", "italic", "underline", "strike"):
        if key not in crit:
            continue
        v, src = resolver.resolve(p, r, _toggle_getter(_TOGGLE_TAGS[key]))
        effective = bool(v)  # not defined anywhere = off, Word's baseline
        if effective != crit[key]:
            return None
        via[key] = src if src else "absent_default_off"
    if "font" in crit:
        v, src = resolver.resolve(p, r, _FormatResolver._font_of)
        if v is None:
            return None
        kind, name = v
        if kind == "theme":
            theme_counter[0] += 1
            return None
        if name.lower() != crit["font"].lower():
            return None
        via["font"] = src
    if "size_pt" in crit:
        v, src = resolver.resolve(p, r, _FormatResolver._size_of)
        if v is None or abs(v - float(crit["size_pt"])) > 0.01:
            return None
        via["size_pt"] = src
    if "color" in crit:
        v, src = resolver.resolve(p, r, _color_of)
        if v is None:
            return None
        kind, val = v
        if kind == "theme":
            theme_counter[0] += 1
            return None
        if val != crit["color"]:
            return None
        via["color"] = src
    if "highlight" in crit:
        v, src = resolver.resolve(p, r, _highlight_of)
        if v is None or v.lower() != crit["highlight"]:
            return None
        via["highlight"] = src
    if "style" in crit:
        want = crit["style"].lower()
        char_sid = resolver.run_char_style(r)
        para_sid = resolver.para_style_id(p)
        if char_sid and want in (
            char_sid.lower(),
            style_names.get(char_sid, "").lower(),
        ):
            via["style"] = "character_style"
        elif para_sid and want in (
            para_sid.lower(),
            style_names.get(para_sid, "").lower(),
        ):
            via["style"] = "paragraph_style"
        else:
            return None
    return via


def _matching_stretches(
    resolver: _Resolver,
    style_names: dict[str, str],
    p,
    crit: dict,
    theme_counter: list[int],
) -> tuple[str, list, list[tuple[int, int, dict]]]:
    """(text, segments, stretches) for one paragraph. A stretch is a
    (start, end, matched_via) span of contiguous runs whose EFFECTIVE
    formatting satisfies every criterion, merged only while the per-criterion
    sources stay identical. Shared by find_formatted and replace_formatted so
    the two always agree on what matches."""
    text, segments = _runmap.build_map(p)
    if not segments:
        return text, segments, []
    # Per-run spans in paragraph offsets, document order.
    run_spans: list[tuple] = []  # (run, start, end)
    for seg in segments:
        if run_spans and run_spans[-1][0] is seg.run:
            run_spans[-1] = (seg.run, run_spans[-1][1], seg.end)
        else:
            run_spans.append((seg.run, seg.start, seg.end))
    # Contiguous stretches of matching runs with identical sources.
    stretches: list[tuple[int, int, dict]] = []
    for run, start, end in run_spans:
        via = _run_matches(resolver, style_names, p, run, crit, theme_counter)
        if via is None:
            continue
        if (
            stretches
            and stretches[-1][1] == start
            and stretches[-1][2] == via
        ):
            stretches[-1] = (stretches[-1][0], end, via)
        else:
            stretches.append((start, end, via))
    return text, segments, stretches


def find_formatted(
    pkg: DocxPackage,
    query: str | None = None,
    *,
    formatting: dict,
    scope: str = "body",
) -> dict:
    """Find text by its effective formatting (read-only).

    formatting: any of bold/italic/underline/strike (true/false), font,
    size_pt, color (hex), highlight (Word highlight name), style (style id or
    name — matches the style actually applied to the run or its paragraph,
    including the document's default paragraph style). All given criteria
    must hold (AND). Unknown keys are rejected with the allowed list.

    query=None returns every contiguous stretch of text carrying the
    formatting; with a query (case-sensitive), only occurrences of that text
    lying entirely inside such a stretch are returned. scope: 'body'
    (document.xml, table cells included) or 'all' (adds footnotes/endnotes).

    Each match reports paragraph_index (None inside table cells and notes
    parts), character offsets within the paragraph, the matched text, and
    matched_via — per criterion, whether it held explicitly on the run or
    arrived through a character style, paragraph style, or document defaults.
    Runs whose relevant value is a theme reference (theme fonts/colors) are
    not matchable from the file alone and are counted in
    unresolved_theme_runs instead of being guessed at.

    Text-box content is excluded from the walk entirely (Word stores every
    box TWICE — mc:Choice + mc:Fallback — so walking boxes reports doubled
    matches), consistent with get_text/search_and_replace and with
    replace_formatted. Boxes are readable via get_textbox_text or get_text
    include_textboxes."""
    crit = _normalize(formatting)
    if query is not None and query == "":
        raise WordMcpError("query must be non-empty (or None for all runs)")
    if scope not in _SCOPES:
        raise WordMcpError(
            f"scope must be one of {sorted(_SCOPES)}, got {scope!r}"
        )

    resolver = _Resolver(pkg)
    style_names = resolver.style_names()
    body_paras = [
        (el, idx) for kind, idx, el in body_items(pkg) if kind == "paragraph"
    ]
    body_idx = {id(el): idx for el, idx in body_paras}
    keepalive = [el for el, _ in body_paras]  # live proxies keep id() stable

    matches: list[dict] = []
    theme_counter = [0]
    truncated = False

    for part in _SCOPES[scope]:
        if not pkg.has_part(part):
            continue
        root = pkg.root(part)
        for p in root.iter(qn("w:p")):
            # Text-box paragraphs are a separate story (stored doubled in
            # mc:Choice AND mc:Fallback); walking them here returned box
            # text twice. Skip boxes entirely, exactly as replace_formatted
            # and search_and_replace do.
            if _runmap._in_textbox(p, root):
                continue
            text, segments, stretches = _matching_stretches(
                resolver, style_names, p, crit, theme_counter
            )
            if not stretches:
                continue
            para_index = (
                body_idx.get(id(p)) if part == "word/document.xml" else None
            )
            for start, end, via in stretches:
                if query is None:
                    spans = [(start, end)]
                else:
                    spans = []
                    pos = text.find(query, start)
                    while pos != -1 and pos + len(query) <= end:
                        spans.append((pos, pos + len(query)))
                        pos = text.find(query, pos + 1)
                for s, e in spans:
                    if len(matches) >= _MAX_MATCHES:
                        truncated = True
                        break
                    matches.append(
                        {
                            "part": part,
                            "paragraph_index": para_index,
                            "start": s,
                            "end": e,
                            "text": text[s:e],
                            "matched_via": dict(via),
                        }
                    )

    del keepalive
    result = {
        "criteria": crit,
        "query": query,
        "scope": scope,
        "total": len(matches),
        "matches": matches,
        "unresolved_theme_runs": theme_counter[0],
    }
    if truncated:
        result["truncated"] = (
            f"stopped at {_MAX_MATCHES} matches; narrow the criteria"
        )
    if theme_counter[0]:
        result["note"] = (
            "some runs use theme font/color references that cannot be "
            "resolved from the file alone; they were skipped, not matched"
        )
    return result


# ------------------------------------------------------ style-aware replace


def replace_formatted(
    pkg: DocxPackage,
    *,
    formatting: dict,
    replace: str,
    find: str | None = None,
    scope: str = "body",
    regex: bool = False,
    max_replacements: int | None = None,
) -> dict:
    """Replace text ONLY inside stretches carrying the given EFFECTIVE
    formatting — the mutation twin of find_formatted, sharing its resolution
    machinery and criteria keys (bold/italic/underline/strike, font, size_pt,
    color, highlight, style), so what find_formatted reports matched is
    exactly what this rewrites.

    find=None replaces each entire matching stretch with `replace`; with
    `find` (literal, case-sensitive), only occurrences lying entirely inside
    a matching stretch are replaced. Matching is per paragraph; the
    replacement text inherits the formatting of the first character it
    replaces (which is inside the matched stretch, so the formatting is
    kept). Fragmented runs are handled through the same offset-map mechanism
    search_and_replace uses, so a match spanning several runs is safe.

    regex combined with formatting criteria is refused (offsets from a regex
    interacting with stretch boundaries are not guaranteed solid); use
    search_and_replace for regex, or a literal find here. Text-box content is
    never touched (set_textbox_text is the box editor). Runs whose relevant
    value is a theme reference are skipped, not guessed, and counted in
    unresolved_theme_runs. max_replacements is a blast-radius guard: over the
    limit, NOTHING is replaced and the error reports the count."""
    crit = _normalize(formatting)
    if regex:
        raise WordMcpError(
            "regex combined with formatting criteria is not supported "
            "(conservative refusal); pass a literal find string, or use "
            "search_and_replace for pure-text regex work"
        )
    if find is not None and find == "":
        raise WordMcpError(
            "find must be non-empty (or None to replace whole stretches)"
        )
    if not isinstance(replace, str):
        raise WordMcpError("replace must be a string (may be empty)")
    if scope not in _SCOPES:
        raise WordMcpError(
            f"scope must be one of {sorted(_SCOPES)}, got {scope!r}"
        )

    resolver = _Resolver(pkg)
    style_names = resolver.style_names()
    body_paras = [
        (el, idx) for kind, idx, el in body_items(pkg) if kind == "paragraph"
    ]
    body_idx = {id(el): idx for el, idx in body_paras}
    keepalive = [el for el, _ in body_paras]  # live proxies keep id() stable

    theme_counter = [0]
    # Plan first (so the guard can refuse with nothing changed), then apply.
    plan: list[tuple] = []  # (part, p, text, segments, [(start, end, via)])
    for part in _SCOPES[scope]:
        if not pkg.has_part(part):
            continue
        root = pkg.root(part)
        for p in root.iter(qn("w:p")):
            # Box paragraphs are a separate story; search_and_replace skips
            # them for the same reason (Choice/Fallback duplication).
            if _runmap._in_textbox(p, root):
                continue
            text, segments, stretches = _matching_stretches(
                resolver, style_names, p, crit, theme_counter
            )
            if not stretches:
                continue
            spans: list[tuple[int, int, dict]] = []
            for start, end, via in stretches:
                if find is None:
                    spans.append((start, end, via))
                else:
                    pos = text.find(find, start)
                    while pos != -1 and pos + len(find) <= end:
                        spans.append((pos, pos + len(find), via))
                        pos = text.find(find, pos + len(find))
            if spans:
                plan.append((part, p, text, segments, spans))

    total = sum(len(spans) for *_, spans in plan)
    if max_replacements is not None and total > max_replacements:
        raise WordMcpError(
            f"would make {total} replacements, over the max_replacements "
            f"guard of {max_replacements}; nothing was changed — narrow the "
            "criteria/find or raise the limit"
        )

    details: list[dict] = []
    truncated = False
    dirty_parts: set[str] = set()
    for part, p, text, segments, spans in plan:
        para_index = (
            body_idx.get(id(p)) if part == "word/document.xml" else None
        )
        # Right-to-left so earlier offsets stay valid within one snapshot.
        for start, end, via in reversed(spans):
            _runmap.replace_range(p, segments, start, end, replace)
        dirty_parts.add(part)
        for start, end, via in spans:
            if len(details) >= _MAX_MATCHES:
                truncated = True
                break
            details.append(
                {
                    "part": part,
                    "paragraph_index": para_index,
                    "start": start,
                    "end": end,
                    "replaced_text": text[start:end],
                    "matched_via": dict(via),
                }
            )
    for part in dirty_parts:
        pkg.mark_dirty(part)

    del keepalive
    key = find if find is not None else "(formatted stretch)"
    result = {
        "replaced": {key: total},
        "total": total,
        "criteria": crit,
        "find": find,
        "scope": scope,
        "unresolved_theme_runs": theme_counter[0],
        "replacements": details,
    }
    if truncated:
        result["replacements_truncated"] = (
            f"detail list stopped at {_MAX_MATCHES} entries; counts above "
            "cover everything"
        )
    if theme_counter[0]:
        result["note"] = (
            "some runs use theme font/color references that cannot be "
            "resolved from the file alone; they were skipped, not replaced"
        )
    return result


# ------------------------------------------------- paragraph-format reading

_WORD_BASELINES = {
    "alignment": "left",
    "space_before_pt": 0,
    "space_after_pt": 0,
    "indent_left_pt": 0,
    "indent_right_pt": 0,
    "first_line_indent_pt": 0,
    "keep_with_next": False,
    "widow_control": True,
    "page_break_before": False,
}

_JC_MAP = {"both": "justify", "start": "left", "end": "right"}


def _num(raw: str | None, per_pt: int = 20) -> float | int | None:
    if raw is None:
        return None
    v = int(raw) / per_pt
    return int(v) if v == int(v) else v


def _jc_of(ppr):
    if ppr is None:
        return None
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        return None
    val = jc.get(qn("w:val"))
    return _JC_MAP.get(val, val)


def _spacing_attr(attr: str):
    def get(ppr):
        if ppr is None:
            return None
        sp = ppr.find(qn("w:spacing"))
        return _num(sp.get(qn(attr))) if sp is not None else None

    return get


def _line_of(ppr):
    """(value, rule): rule 'auto' -> value is a multiple (1.0, 2.0);
    'exact'/'atLeast' -> value is in points."""
    if ppr is None:
        return None
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        return None
    line = sp.get(qn("w:line"))
    if line is None:
        return None
    rule = sp.get(qn("w:lineRule"), "auto")
    if rule == "auto":
        return (round(int(line) / 240, 2), "auto")
    return (round(int(line) / 20, 1), rule)


def _ind_attr(*attrs: str):
    def get(ppr):
        if ppr is None:
            return None
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            return None
        for attr in attrs:
            v = ind.get(qn(attr))
            if v is not None:
                return _num(v)
        return None

    return get


def _first_line_of(ppr):
    if ppr is None:
        return None
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        return None
    hanging = ind.get(qn("w:hanging"))
    if hanging is not None:
        v = _num(hanging)
        return -v if v is not None else None
    return _num(ind.get(qn("w:firstLine")))


def _outline_of(ppr):
    if ppr is None:
        return None
    el = ppr.find(qn("w:outlineLvl"))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    return int(val) if val is not None else None


def _ppr_toggle(tag: str):
    def get(ppr):
        if ppr is None:
            return None
        el = ppr.find(qn(tag))
        if el is None:
            return None
        val = el.get(qn("w:val"))
        if val is None:
            return True
        return val not in ("0", "false", "none")

    return get


_PARA_PROPS: list[tuple[str, object]] = [
    ("alignment", _jc_of),
    ("space_before_pt", _spacing_attr("w:before")),
    ("space_after_pt", _spacing_attr("w:after")),
    ("line_spacing", _line_of),
    ("indent_left_pt", _ind_attr("w:left", "w:start")),
    ("indent_right_pt", _ind_attr("w:right", "w:end")),
    ("first_line_indent_pt", _first_line_of),
    ("keep_with_next", _ppr_toggle("w:keepNext")),
    ("widow_control", _ppr_toggle("w:widowControl")),
    ("page_break_before", _ppr_toggle("w:pageBreakBefore")),
    ("outline_level", _outline_of),
]


def get_paragraph_format(
    pkg: DocxPackage, start: int, end: int | None = None
) -> dict:
    """EFFECTIVE paragraph-level formatting for body paragraphs [start, end]
    (inclusive; end defaults to start), resolved the way Word resolves it:
    the paragraph's own pPr, then the paragraph-style chain (basedOn), then
    styles.xml docDefaults, then Word's built-in baseline. Every property
    reports its value AND its source — 'explicit' | 'paragraph_style' |
    'document_defaults' | 'word_default' — mirroring find_formatted's
    matched_via convention. line_spacing also carries its rule: 'auto' means
    the value is a multiple (1.0, 2.0); 'exact'/'atLeast' mean points.
    Indents contributed by list numbering (w:numPr level definitions) are
    NOT resolved here; paragraphs carrying numbering are flagged with
    has_numbering so the caller knows the rendered indent may differ.
    outline_level (w:outlineLvl, 0 = top) reports source 'explicit' |
    'paragraph_style' | 'none' — it is what Word's navigation pane and TOC
    field harvesting honor, so 'none' on a visually-formatted heading means
    the paragraph will not appear in the outline or TOC (set it with
    set_paragraph_format's outline_level key)."""
    if start < 0:
        raise WordMcpError("start must be >= 0")
    end = start if end is None else end
    if end < start:
        raise WordMcpError("end must be >= start")

    resolver = _Resolver(pkg)
    paras = [
        (idx, el) for kind, idx, el in body_items(pkg) if kind == "paragraph"
    ]
    total = len(paras)
    if end >= total:
        raise WordMcpError(
            f"paragraph range {start}-{end} exceeds the document body "
            f"({total} paragraph(s)"
            + (f", valid indices 0-{total - 1})" if total else ")")
        )

    out = []
    for idx, p in paras[start : end + 1]:
        fmt: dict = {}
        for key, getter in _PARA_PROPS:
            v, src = resolver.resolve_para(p, getter)
            if key == "line_spacing":
                if v is None:
                    value, rule, src = 1.0, "auto", "word_default"
                else:
                    value, rule = v
                fmt[key] = {"value": value, "rule": rule, "source": src}
                continue
            if key == "outline_level":
                # No baseline exists: body text simply has no outlineLvl.
                if v is None:
                    fmt[key] = {"value": None, "source": "none"}
                else:
                    fmt[key] = {"value": v, "source": src}
                continue
            if v is None:
                v, src = _WORD_BASELINES[key], "word_default"
            fmt[key] = {"value": v, "source": src}
        text, _ = _runmap.build_map(p)
        entry = {
            "index": idx,
            "style": resolver.para_style_id(p) or "Normal",
            "text_preview": text[:60],
            "format": fmt,
        }
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:numPr")) is not None:
            entry["has_numbering"] = True
        out.append(entry)
    return {"start": start, "end": end, "paragraphs": out}
