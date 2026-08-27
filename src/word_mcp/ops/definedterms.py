"""Defined-terms audit for legal documents.

Contracts introduce capitalized defined terms («"Confidential Information"
means...», «(the "Agreement")») and then rely on them; the classic drafting
failures are terms defined but never used, terms used but never defined,
terms defined twice, and terms used before their definition. This module
finds all four, with paragraph indices for every finding.

Detection is HEURISTIC where it has to be: definitions are found by pattern
(overridable), but "used but never defined" can only guess at what LOOKS
like a defined term (capitalized multi-word phrases, CamelCase words), so
those results are flagged as review candidates, never as certainties.

Scope: body-level paragraphs only (the stable paragraph-index addressing
scheme). Definitions placed inside table cells or footnotes are not scanned;
that limitation is stated in the result notes rather than hidden.
"""

from __future__ import annotations

import re

from ..core.errors import WordMcpError
from ..core.package import DocxPackage
from . import _regex
from .read import _outline_level, _style_outline_map, body_items, paragraph_text

# Each pattern must expose the term as capture group 1. Straight quotes only:
# input text is quote-normalized before matching, so Word's curly quotes work.
DEFAULT_DEFINITION_PATTERNS = [
    # "Term" means ... / "Term" shall mean ...
    r'"([A-Z][^"\n]{0,80}?)"\s+(?:shall\s+mean|means)\b',
    # (the "Term")
    r'\(\s*the\s+"([A-Z][^"\n]{0,80}?)"\s*\)',
    # (each, a "Term") / (each an "Term")
    r'\(\s*each,?\s+an?\s+"([A-Z][^"\n]{0,80}?)"\s*\)',
]

# Candidate shapes for the used-but-never-defined heuristic.
_MULTIWORD = re.compile(r"\b[A-Z][a-z0-9]+(?:\s+[A-Z][a-z0-9]+)+\b")
_CAMELCASE = re.compile(r"\b[A-Z][a-z0-9]+(?:[A-Z][a-z0-9]+)+\b")

_QUOTE_MAP = str.maketrans({"“": '"', "”": '"',
                            "‘": "'", "’": "'"})


def _normalize_quotes(text: str) -> str:
    return text.translate(_QUOTE_MAP)


def _sentence_start(text: str, pos: int) -> bool:
    """Is the character position the start of a sentence (or paragraph)?
    Walks back over whitespace, quotes, and opening brackets."""
    i = pos - 1
    while i >= 0 and text[i] in " \t\"'([“‘":
        i -= 1
    return i < 0 or text[i] in ".!?:;"


def _term_occurrences(term: str, text: str) -> list[int]:
    """Word-boundary, case-sensitive occurrences of a term."""
    return [
        m.start()
        for m in re.finditer(
            r"(?<!\w)" + re.escape(term) + r"(?!\w)", text
        )
    ]


def check_defined_terms(
    pkg: DocxPackage, *, definition_patterns: list | None = None
) -> dict:
    """Audit the document's defined terms.

    definition_patterns: optional list of regex strings replacing the
    defaults; each must contain a capture group for the term (refused
    otherwise). Defaults recognize «"Term" means», «"Term" shall mean»,
    «(the "Term")», and «(each, a "Term")», with straight or curly quotes.

    Returns, all with paragraph indices:
    - defined_terms: every definition found, with definition sites and use
      counts (a "use" is a word-boundary occurrence of the exact term
      outside its own quoted definition span)
    - defined_never_used: defined but never referenced again
    - defined_multiple_times: two or more definition sites
    - first_use_before_definition: earliest use paragraph precedes the
      earliest definition paragraph (paragraph granularity; a use and its
      definition in the SAME paragraph is not flagged)
    - used_never_defined: HEURISTIC review candidates — capitalized
      multi-word or CamelCase terms appearing 2+ times with at least one
      mid-sentence occurrence and no definition. A phrase capitalized only
      at sentence starts is excluded (ordinary sentence capitalization, not
      a term). Verify these by eye; proper nouns and headings can slip
      through despite the filters.
    """
    raw_patterns = (
        definition_patterns
        if definition_patterns is not None
        else DEFAULT_DEFINITION_PATTERNS
    )
    if not raw_patterns:
        raise WordMcpError("definition_patterns must not be empty")
    for pat in raw_patterns:
        compiled = _regex.compile_user_pattern(pat)
        if compiled.groups < 1:
            raise WordMcpError(
                f"definition pattern {pat!r} has no capture group; group 1 "
                "must capture the term"
            )

    # ---- collect paragraphs (body-level only; heading flag for candidates)
    style_outline = _style_outline_map(pkg)
    paras: list[tuple[int, str, bool]] = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph":
            continue
        text = _normalize_quotes(paragraph_text(el))
        is_heading = _outline_level(el, style_outline) is not None
        paras.append((idx, text, is_heading))

    # ---- definitions: term -> {para_idx: [term-group spans]}
    defs: dict[str, dict[int, list[tuple[int, int]]]] = {}
    for idx, text, _ in paras:
        for pat in raw_patterns:
            for m in _regex.finditer(pat, text):
                term = m.group(1)
                if not term or not term.strip():
                    continue
                term = term.strip()
                defs.setdefault(term, {}).setdefault(idx, []).append(
                    m.span(1)
                )

    # ---- uses of each defined term (outside its quoted definition spans)
    defined_terms: list[dict] = []
    never_used: list[dict] = []
    multiple: list[dict] = []
    use_before_def: list[dict] = []
    for term in sorted(defs):
        def_sites = sorted(defs[term])
        # terms that shadow this one ("Master Agreement" shadows
        # "Agreement"): a hit inside the longer term's own occurrence is a
        # use of the LONGER term, not this one
        shadows = [
            other for other in defs
            if other != term and len(other) > len(term) and term in other
        ]
        use_paras: list[int] = []
        for idx, text, _ in paras:
            spans = defs[term].get(idx, [])
            shadow_spans = [
                (p, p + len(other))
                for other in shadows
                for p in _term_occurrences(other, text)
            ]
            for pos in _term_occurrences(term, text):
                end = pos + len(term)
                inside_def = any(
                    pos < d_end and end > d_start for d_start, d_end in spans
                )
                inside_shadow = any(
                    pos >= s_start and end <= s_end
                    for s_start, s_end in shadow_spans
                )
                if not inside_def and not inside_shadow:
                    use_paras.append(idx)
        entry = {
            "term": term,
            "defined_at": def_sites,
            "use_count": len(use_paras),
            "used_at": sorted(set(use_paras)),
        }
        defined_terms.append(entry)
        if not use_paras:
            never_used.append({"term": term, "defined_at": def_sites})
        if len(def_sites) > 1:
            multiple.append({"term": term, "defined_at": def_sites})
        if use_paras and min(use_paras) < def_sites[0]:
            use_before_def.append(
                {
                    "term": term,
                    "first_use_paragraph": min(use_paras),
                    "first_definition_paragraph": def_sites[0],
                }
            )

    # ---- used-but-never-defined heuristic
    candidates: dict[str, dict] = {}
    defined_set = set(defs)
    for idx, text, is_heading in paras:
        if is_heading:
            continue  # Title Case headings are not term uses
        for m in list(_MULTIWORD.finditer(text)) + list(
            _CAMELCASE.finditer(text)
        ):
            phrase = m.group(0)
            if phrase in defined_set:
                continue
            entry = candidates.setdefault(
                phrase,
                {"count": 0, "paragraphs": set(), "mid_sentence": False},
            )
            entry["count"] += 1
            entry["paragraphs"].add(idx)
            if not _sentence_start(text, m.start()):
                entry["mid_sentence"] = True
    used_never_defined = [
        {
            "term": phrase,
            "count": data["count"],
            "paragraphs": sorted(data["paragraphs"]),
            "note": (
                "heuristic review candidate — looks like a defined term "
                "(capitalized, recurring, used mid-sentence) but no "
                "definition matched; may be a proper noun"
            ),
        }
        for phrase, data in sorted(candidates.items())
        if data["count"] >= 2 and data["mid_sentence"]
    ]

    return {
        "defined_terms": defined_terms,
        "defined_never_used": never_used,
        "used_never_defined": used_never_defined,
        "defined_multiple_times": multiple,
        "first_use_before_definition": use_before_def,
        "definition_patterns": list(raw_patterns),
        "paragraphs_scanned": len(paras),
        "notes": [
            "body-level paragraphs only — definitions inside table cells, "
            "footnotes, or headers are not scanned",
            "used_never_defined is a shape heuristic (2+ occurrences, at "
            "least one mid-sentence); it flags review candidates, not "
            "certainties",
            "use detection is case-sensitive and exact — inflected or "
            "abbreviated references to a term are not counted",
        ],
    }
