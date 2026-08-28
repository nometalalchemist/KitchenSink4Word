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
    unresolved_theme_runs instead of being guessed at."""
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
        for p in pkg.root(part).iter(qn("w:p")):
            text, segments = _runmap.build_map(p)
            if not segments:
                continue
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
                via = _run_matches(
                    resolver, style_names, p, run, crit, theme_counter
                )
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
