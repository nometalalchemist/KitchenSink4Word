"""Journal-ready word counts with named exclusion zones.

Journals count words differently from Word's status bar: most limits exclude
the reference list, figure/table captions, and often footnotes, tables, the
abstract, or block quotations. This module computes one consistent total and
subtracts named zones so the number reported to a journal is defensible.

Counting rules are the SAME as ops/stats.py (whitespace-delimited tokens,
``\\S+``), so numbers here reconcile with get_word_count — with two
deliberate differences, both documented in the result: heading text is part
of the total here (get_word_count skips it), and footnote/endnote text is
part of the total here (get_word_count never reads the note parts). The
arithmetic invariant always holds: total = included + sum(excluded zones).

Zone detection is heuristic where Word gives no structure (the reference
list is found by its heading, the same way check_citation_parity finds it);
every detected zone's location is reported so the caller can eyeball it.
"""

from __future__ import annotations

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from .citecheck import _REF_HEADINGS
from .localization import canonical_for_name, cjk_aware_word_count
from .read import (
    _outline_level,
    _style_id,
    _style_outline_map,
    body_items,
    get_table,
    list_endnotes,
    list_footnotes,
    list_styles,
    paragraph_text,
)

ALLOWED_EXCLUSIONS = (
    "references",
    "captions",
    "footnotes",
    "endnotes",
    "block_quotes",
    "tables",
    "headings",
    "front_matter",
    "abstract",
)

# Zone precedence for body paragraphs: a paragraph lands in the FIRST active
# zone that claims it, so nothing is double-subtracted (a heading inside the
# reference list counts as "references", not "headings").
_PRECEDENCE = (
    "references",
    "abstract",
    "front_matter",
    "headings",
    "captions",
    "block_quotes",
)

_QUOTE_STYLE_KEYS = {"quote", "blockquote", "blocktext", "intensequote"}
_CAPTION_STYLE_KEYS = {"caption", "tablecaption", "figurecaption"}


def _norm(s: str | None) -> str:
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())


def _style_key_sets(pkg: DocxPackage) -> tuple[set[str], set[str]]:
    """(quote style ids, caption style ids) by id or display name. Style IDs
    stay English in every language version, but DISPLAY names localize
    (Caption is 캡션 on a Korean install), so names are additionally resolved
    through the localization aliases."""
    quote_ids: set[str] = set()
    caption_ids: set[str] = set()
    for s in list_styles(pkg):
        keys = {_norm(s["id"]), _norm(s["name"])}
        canonical = canonical_for_name(s["name"])
        if keys & _QUOTE_STYLE_KEYS or canonical in ("quote", "block_text"):
            quote_ids.add(s["id"])
        if keys & _CAPTION_STYLE_KEYS or canonical == "caption":
            caption_ids.add(s["id"])
    return quote_ids, caption_ids


def _indented_both_sides(p) -> bool:
    """Direct formatting: left AND right indent > 0 — the classic manually
    formatted block quotation."""
    ind = p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    if ind is None:
        return False

    def _val(*attrs: str) -> int:
        for a in attrs:
            v = ind.get(qn(a))
            if v is not None:
                try:
                    return int(v)
                except ValueError:
                    return 0
        return 0

    return _val("w:left", "w:start") > 0 and _val("w:right", "w:end") > 0


def word_count_with_exclusions(
    pkg: DocxPackage,
    *,
    exclude: tuple[str, ...] | list[str] = (
        "references",
        "captions",
        "footnotes",
    ),
) -> dict:
    """Word count minus named zones — the number a journal actually wants.

    exclude: any of references (the reference-list section, located by its
    References/Bibliography heading), captions (Caption-style paragraphs),
    footnotes / endnotes (the note parts), block_quotes (Quote/Block Text
    styles or paragraphs indented on both sides), tables (all table text),
    headings (all heading paragraphs), front_matter (everything before the
    first level-1 heading), abstract (an Abstract-headed section). Unknown
    names are rejected with the allowed list.

    Returns total, the excluded breakdown by zone, and the included count;
    total = included + sum(excluded) always. Counting uses the same
    whitespace tokenization as get_word_count for space-delimited text;
    Japanese/Chinese (CJK) characters count one word per character, reported
    via cjk_chars and counting ("spaces"|"cjk"|"mixed"). Unlike
    get_word_count, the total here includes heading text and footnote/endnote
    text (so they can be excluded explicitly and the arithmetic stays
    honest). Zone locations are reported for review. Read-only — the file is
    not modified."""
    exclude = tuple(exclude)
    unknown = set(exclude) - set(ALLOWED_EXCLUSIONS)
    if unknown:
        raise WordMcpError(
            f"unknown exclusion name(s) {sorted(unknown)}; "
            f"allowed: {list(ALLOWED_EXCLUSIONS)}"
        )
    active = set(exclude)

    style_outline = _style_outline_map(pkg)
    quote_ids, caption_ids = _style_key_sets(pkg)
    items = body_items(pkg)

    # CJK-aware counting: identical to stats' \S+ tokenization for pure
    # space-delimited text (English, Korean, ...); Japanese/Chinese text adds
    # one word per CJK character (the academic convention). cjk_total is
    # surfaced in the result so the caller can see when it mattered.
    cjk_total = 0

    def _count_words(text: str) -> int:
        nonlocal cjk_total
        r = cjk_aware_word_count(text)
        cjk_total += r["cjk_chars"]
        return r["words"]

    # ---- locate heading-delimited zones on the body paragraph index axis
    headings: list[tuple[int, int, str]] = []  # (paragraph_index, level, text)
    for kind, idx, el in items:
        if kind != "paragraph":
            continue
        level = _outline_level(el, style_outline)
        text = paragraph_text(el).strip()
        if level is not None and text:
            headings.append((idx, level, text))

    def _span_after(start_idx: int) -> int | None:
        nxt = [i for i, _, _ in headings if i > start_idx]
        return min(nxt) if nxt else None

    ref_span = None
    ref_heading = next(
        (h for h in headings if _REF_HEADINGS.match(h[2])), None
    )
    if ref_heading is not None:
        ref_span = (ref_heading[0], _span_after(ref_heading[0]))

    abstract_span = None
    # Text-matched like the reference heading: English plus the Korean
    # equivalents (초록 / 요약); other languages' abstract headings are not
    # yet covered.
    _abstract_titles = {"abstract", "초록", "요약"}
    abstract_heading = next(
        (h for h in headings if h[2].strip().lower() in _abstract_titles),
        None,
    )
    if abstract_heading is not None:
        abstract_span = (abstract_heading[0], _span_after(abstract_heading[0]))

    first_h1 = next((i for i, lvl, _ in headings if lvl == 1), None)

    def _in_span(idx: int | None, span) -> bool:
        if idx is None or span is None:
            return False
        start, end = span
        return idx >= start and (end is None or idx < end)

    # ---- classify and count
    zone_words = {z: 0 for z in exclude}
    included = 0
    total = 0

    def _zone_for(idx: int | None, level, style, indented: bool) -> str | None:
        for z in _PRECEDENCE:
            if z not in active:
                continue
            if z == "references" and _in_span(idx, ref_span):
                return z
            if z == "abstract" and _in_span(idx, abstract_span):
                return z
            if (
                z == "front_matter"
                and first_h1 is not None
                and idx is not None
                and idx < first_h1
            ):
                return z
            if z == "headings" and level is not None:
                return z
            if z == "captions" and (
                style in caption_ids or _norm(style) in _CAPTION_STYLE_KEYS
            ):
                return z
            if z == "block_quotes" and (
                style in quote_ids
                or _norm(style) in _QUOTE_STYLE_KEYS
                or indented
            ):
                return z
        return None

    last_p_idx: int | None = None
    for kind, idx, el in items:
        if kind == "paragraph":
            last_p_idx = idx
            text = paragraph_text(el)
            words = _count_words(text)
            if not words:
                continue
            total += words
            zone = _zone_for(
                idx,
                _outline_level(el, style_outline),
                _style_id(el),
                _indented_both_sides(el),
            )
            if zone:
                zone_words[zone] += words
            else:
                included += words
        else:  # table — position on the paragraph axis = last paragraph seen
            table = get_table(pkg, idx)
            table_text = " ".join(
                c["text"] for row in table["cells"] for c in row
            )
            words = _count_words(table_text)
            if not words:
                continue
            total += words
            zone = None
            for z in ("references", "abstract", "front_matter"):
                if z not in active:
                    continue
                if z == "references" and _in_span(last_p_idx, ref_span):
                    zone = z
                elif z == "abstract" and _in_span(last_p_idx, abstract_span):
                    zone = z
                elif (
                    z == "front_matter"
                    and first_h1 is not None
                    and last_p_idx is not None
                    and last_p_idx < first_h1
                ):
                    zone = z
                if zone:
                    break
            if zone is None and "tables" in active:
                zone = "tables"
            if zone:
                zone_words[zone] += words
            else:
                included += words

    # ---- note parts
    for zone_name, lister in (
        ("footnotes", list_footnotes),
        ("endnotes", list_endnotes),
    ):
        words = sum(_count_words(n["text"]) for n in lister(pkg))
        if not words:
            continue
        total += words
        if zone_name in active:
            zone_words[zone_name] += words
        else:
            included += words

    detected = {
        "references_section": (
            {
                "heading": ref_heading[2],
                "start_paragraph": ref_span[0],
                "end_paragraph": ref_span[1],
            }
            if ref_span
            else None
        ),
        "abstract_section": (
            {
                "heading": abstract_heading[2],
                "start_paragraph": abstract_span[0],
                "end_paragraph": abstract_span[1],
            }
            if abstract_span
            else None
        ),
        "first_level1_heading_paragraph": first_h1,
    }

    notes = [
        "tokenization: whitespace-delimited tokens (\\S+), same rule as "
        "get_word_count",
        "total here includes heading and footnote/endnote text so those "
        "zones can be excluded explicitly; get_word_count's total differs "
        "by exactly those words",
    ]
    if "references" in active and ref_span is None:
        notes.append(
            "no References/Bibliography heading found — the references "
            "exclusion matched nothing"
        )
    if "abstract" in active and abstract_span is None:
        notes.append("no Abstract heading found — the abstract exclusion matched nothing")
    if "front_matter" in active and first_h1 is None:
        notes.append(
            "no level-1 heading found — front_matter is undefined and "
            "matched nothing (refusing to treat the whole document as "
            "front matter)"
        )

    if cjk_total:
        notes.append(
            "CJK text detected: Japanese/Chinese characters are counted one "
            "word per character (space-delimited text is unaffected)"
        )

    result = {
        "total": total,
        "included": included,
        "excluded": zone_words,
        "excluded_total": sum(zone_words.values()),
        "exclusions_applied": list(exclude),
        "cjk_chars": cjk_total,
        "counting": (
            "spaces"
            if not cjk_total
            else ("cjk" if total == cjk_total else "mixed")
        ),
        "zones_detected": detected,
        "notes": notes,
    }
    # Hard invariant — if this ever fails it is a bug, not a document quirk.
    assert result["total"] == result["included"] + result["excluded_total"]
    return result
