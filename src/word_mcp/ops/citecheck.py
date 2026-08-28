"""Citation-reference parity checking (APA-style, heuristic).

Cross-checks in-text citations against the reference list in both directions:
- citations whose (surname, year) has no reference entry  -> missing_references
- reference entries never cited in the body              -> uncited_references

This is a FLAGGING tool, not a fixer: APA parsing is heuristic (organizational
authors, legal citations, and unusual formats can evade the patterns), so
results are review candidates, not verdicts. Parenthetical (Author, 2020;
Other, 2021), narrative Author (2020), et al., &, and year-letters (2020a)
are handled.
"""

from __future__ import annotations

import re

from ..core.errors import TargetNotFound
from ..core.package import DocxPackage
from .read import get_outline, get_paragraphs

# Reference-list heading, matched by TEXT (localized): English plus Korean
# (참고문헌 / 참고 문헌), German (Literaturverzeichnis), French
# (Bibliographie), Spanish (Bibliografía), Japanese+Chinese (参考文献),
# Portuguese (Referências), Italian/Portuguese (Bibliografia). Shared by
# journalcount, styleconvert, and anonymize — extending this extends them all.
_REF_HEADINGS = re.compile(
    r"^\s*(references|bibliography|works cited|reference list"
    r"|참고\s*문헌"          # ko
    r"|literaturverzeichnis"  # de
    r"|bibliographie"         # fr (also de alternative)
    r"|bibliograf[íi]a"       # es / it+pt
    r"|参考文献"              # ja + zh
    r"|refer[êe]ncias"        # pt (accented and plain)
    r")\s*$",
    re.I,
)

# "n.d." (no date) is a legal APA year; hangul surnames are legal authors —
# both were invisible to the Latin-only patterns (v1.5 adversarial F4).
_YEAR = r"(?:(?:1[89]\d\d|20\d\d)[a-z]?|n\.d\.)"
# one author token: a capitalized Latin word (accents included — Müller,
# García, Łukasz; \w continues with any Unicode letter) OR a hangul word
_AUT = r"(?:[A-ZÀ-ÞĀ-Žƀ-Ƀ][\w'’\-]+|[가-힣]+)"
# Narrative: Smith (2026) / Smith and Jones (2026) / Smith et al. (2026)
_NARRATIVE = re.compile(
    rf"(?<![\w가-힣])({_AUT})"
    rf"(?:\s+(?:and|&|와|과)\s+{_AUT})*"
    rf"(?:\s+et al\.?)?"
    rf"\s*\(({_YEAR})(?:,\s*(?:p{{1,2}}\.\s*[\d\-–, ]+))?\)"
)
# Parenthetical content chunk: Smith, 2026 / Smith & Jones, 2026, p. 4 /
# Smith et al., 2026
_PAREN_CHUNK = re.compile(
    rf"({_AUT})"
    rf"(?:,?\s+(?:and|&|와|과)\s+{_AUT})*"
    rf"(?:,?\s+et al\.?)?"
    rf",\s*({_YEAR})"
)
# Reference entries: Surname, I. (2026). | Carter, J. (1977a, July 21). |
# Congressional Record, 122(Pt. 24), 30367 (1976, September 15).
# Key = first capitalized word + the year from the first paren containing one
# (full APA date forms and preceding non-year parens are tolerated).
_REF_LEAD = re.compile(rf"^\s*({_AUT})")
_REF_YEAR = re.compile(rf"\(({_YEAR})\b[^)]*\)")


def check_citation_parity(pkg: DocxPackage) -> dict:
    paras = get_paragraphs(pkg)
    outline = get_outline(pkg)

    ref_heading = next(
        (h for h in outline if _REF_HEADINGS.match(h["text"])), None
    )
    if ref_heading is None:
        # Fall back: any paragraph that IS exactly a reference heading.
        candidates = [
            p for p in paras if _REF_HEADINGS.match(p["text"] or "")
        ]
        if not candidates:
            raise TargetNotFound(
                "no References/Bibliography heading found; cannot locate the "
                "reference list"
            )
        ref_start = candidates[-1]["index"]
    else:
        ref_start = ref_heading["paragraph_index"]

    # End of reference list: next heading after ref_start, or document end.
    next_headings = [
        h["paragraph_index"]
        for h in outline
        if h["paragraph_index"] > ref_start
    ]
    ref_end = min(next_headings) if next_headings else None

    body_text = "\n".join(
        p["text"] for p in paras if p["index"] < ref_start
    )
    ref_paras = [
        p["text"]
        for p in paras
        if p["index"] > ref_start
        and (ref_end is None or p["index"] < ref_end)
        and p["text"].strip()
    ]

    # ---- collect in-text citations
    cited: dict[tuple[str, str], int] = {}
    for m in _NARRATIVE.finditer(body_text):
        key = (m.group(1).lower(), m.group(2).lower())
        cited[key] = cited.get(key, 0) + 1
    for paren in re.finditer(r"\(([^()]{4,300}?)\)", body_text):
        inner = paren.group(1)
        if not re.search(_YEAR, inner):
            continue
        for m in _PAREN_CHUNK.finditer(inner):
            key = (m.group(1).lower(), m.group(2).lower())
            cited[key] = cited.get(key, 0) + 1

    # ---- collect reference entries
    listed: dict[tuple[str, str], str] = {}
    unparsed: list[str] = []
    for entry in ref_paras:
        lead = _REF_LEAD.match(entry)
        year = _REF_YEAR.search(entry)
        if lead and year:
            listed[(lead.group(1).lower(), year.group(1).lower())] = entry[:120]
        else:
            unparsed.append(entry[:120])

    missing = sorted(
        {
            f"{surname.title()} ({year})"
            for (surname, year) in cited
            if (surname, year) not in listed
        }
    )
    uncited = sorted(
        listed[key] for key in listed if key not in cited
    )

    return {
        "in_text_citations": sum(cited.values()),
        "unique_cited_works": len(cited),
        "reference_entries": len(listed),
        "missing_references": missing,  # cited but not in the list — serious
        "uncited_references": uncited,  # listed but never cited — review
        "unparsed_reference_entries": unparsed,  # could not extract key
        "parity_ok": not missing and not uncited,
        "note": (
            "Heuristic APA matching on (first-author surname, year). "
            "Organizational authors and unusual formats may need manual "
            "review; unparsed entries were not checked."
        ),
    }
