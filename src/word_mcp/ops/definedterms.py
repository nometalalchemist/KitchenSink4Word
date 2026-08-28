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


# ------------------------------------------------------------------ glossary

DEFINITION_NEEDED = "[DEFINITION NEEDED]"

# «"Term" means X.» / «"Term" shall mean X.» — group 1 is the definition
# tail starting right after the verb. Built per term at harvest time.
_SENTENCE_END = re.compile(r"^(.{3,600}?[.;])(?:\s|$)", re.S)


def _harvest_definition(term: str, text: str) -> str | None:
    """The definition sentence fragment for a term from its defining
    paragraph (quote-normalized text), or None when it cannot be extracted
    cleanly. Only the «"Term" means/shall mean ...» shape yields a
    harvestable definition; parenthetical definitions ((the "Term")) define
    by surrounding prose and cannot be cut out mechanically."""
    m = re.search(
        r'"' + re.escape(term) + r'"\s+(?:shall\s+mean|means)\s+(.+)',
        text,
    )
    if not m:
        return None
    tail = m.group(1).strip()
    m2 = _SENTENCE_END.match(tail)
    if not m2:
        return None
    definition = m2.group(1).strip()
    # A fragment that is all punctuation or suspiciously short is a mangled
    # extraction, not a definition.
    if len(definition.strip(".;,: ")) < 3:
        return None
    return definition


def insert_glossary(
    pkg: DocxPackage,
    *,
    heading: str = "Glossary",
    heading_level: int = 1,
    after_index: int | None = None,
    at_end: bool = True,
    definition_patterns: list | None = None,
) -> dict:
    """Build a glossary section from the document's defined terms (the same
    detection check_defined_terms uses): a heading followed by one paragraph
    per term, alphabetized, with the term in bold and the definition
    harvested from the defining sentence. Terms whose definition cannot be
    extracted cleanly get a [DEFINITION NEEDED] marker instead of a mangled
    fragment (fill those in by hand). Placed at the end of the body by
    default, or after body paragraph `after_index`."""
    from lxml import etree

    from ..core.package import qn

    if not 1 <= heading_level <= 9:
        raise WordMcpError("heading_level must be 1-9")
    audit = check_defined_terms(pkg, definition_patterns=definition_patterns)
    terms = [e["term"] for e in audit["defined_terms"]]
    if not terms:
        raise WordMcpError(
            "no defined terms found in the document (patterns: "
            f"{audit['definition_patterns']}); nothing to build"
        )

    # First defining paragraph per term, for harvesting.
    first_def_para = {
        e["term"]: e["defined_at"][0] for e in audit["defined_terms"]
    }
    para_text: dict[int, str] = {}
    for kind, idx, el in body_items(pkg):
        if kind == "paragraph" and idx in set(first_def_para.values()):
            para_text[idx] = _normalize_quotes(paragraph_text(el))

    entries: list[tuple[str, str]] = []
    needing: list[str] = []
    for term in sorted(terms, key=str.casefold):
        definition = _harvest_definition(
            term, para_text.get(first_def_para[term], "")
        )
        if definition is None:
            definition = DEFINITION_NEEDED
            needing.append(term)
        entries.append((term, definition))

    from .text import ensure_heading_style

    els: list = []
    hp = etree.Element(qn("w:p"))
    ppr = etree.SubElement(hp, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(
        qn("w:val"), ensure_heading_style(pkg, heading_level)
    )
    hr = etree.SubElement(hp, qn("w:r"))
    etree.SubElement(hr, qn("w:t")).text = heading
    els.append(hp)
    for term, definition in entries:
        p = etree.Element(qn("w:p"))
        r1 = etree.SubElement(p, qn("w:r"))
        rpr = etree.SubElement(r1, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:b"))
        t1 = etree.SubElement(r1, qn("w:t"))
        t1.text = term
        t1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r2 = etree.SubElement(p, qn("w:r"))
        t2 = etree.SubElement(r2, qn("w:t"))
        t2.text = f": {definition}"
        t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        els.append(p)

    body = pkg.body()
    if after_index is not None:
        from .text import _body_paragraph

        ref = _body_paragraph(pkg, after_index)
        for el in reversed(els):
            ref.addnext(el)
        position = f"after paragraph {after_index}"
    elif at_end:
        sectpr = body.find(qn("w:sectPr"))
        for el in els:
            if sectpr is not None:
                sectpr.addprevious(el)
            else:
                body.append(el)
        position = "end of body"
    else:
        raise WordMcpError("give after_index or leave at_end=True")

    pkg.mark_dirty()
    result = {
        "glossary_inserted": True,
        "heading": heading,
        "terms": len(entries),
        "position": position,
    }
    if needing:
        result["needing_definition"] = needing
        result["note"] = (
            f"{len(needing)} term(s) had no cleanly extractable definition "
            f"and carry the {DEFINITION_NEEDED} marker; fill those in "
            "manually"
        )
    return result
