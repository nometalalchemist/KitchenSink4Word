"""Document statistics: word counts overall and per heading section."""

from __future__ import annotations

import re

from ..core.package import DocxPackage
from .localization import cjk_aware_word_count
from .read import (
    _outline_level,
    _style_outline_map,
    body_items,
    paragraph_text,
)


def _count_words(text: str) -> int:
    return len(re.findall(r"\S+", text))


def word_count(pkg: DocxPackage, *, by_section: bool = True) -> dict:
    """Total and per-heading-section word/character/paragraph counts. A
    'section' runs from one heading to the next heading of the same or higher
    level; table text counts toward the section containing the table.

    'words' is the space-delimited (\\S+) token count, unchanged for every
    document. For Japanese/Chinese text (which has no word spaces) that
    under-counts, so totals additionally carry 'cjk_chars' (CJK characters,
    the academic zh/ja count unit) and the result carries 'counting'
    ("spaces"|"cjk"|"mixed") flagging when cjk_chars is the number to use."""
    style_outline = _style_outline_map(pkg)
    totals = {
        "words": 0, "characters": 0, "paragraphs": 0, "tables": 0,
        "cjk_chars": 0,
    }
    non_cjk_words = 0
    sections: list[dict] = []
    stack: list[dict] = []  # open sections by level

    def close_to_level(level: int) -> None:
        while stack and stack[-1]["level"] >= level:
            stack.pop()

    def add_counts(words: int, chars: int, text: str) -> None:
        nonlocal non_cjk_words
        totals["words"] += words
        totals["characters"] += chars
        cjk = cjk_aware_word_count(text)
        totals["cjk_chars"] += cjk["cjk_chars"]
        non_cjk_words += cjk["words"] - cjk["cjk_chars"]
        for s in stack:
            s["words"] += words

    for kind, idx, el in body_items(pkg):
        if kind == "paragraph":
            text = paragraph_text(el)
            level = _outline_level(el, style_outline)
            if level is not None and text.strip():
                close_to_level(level)
                entry = {
                    "heading": text.strip(),
                    "level": level,
                    "paragraph_index": idx,
                    "words": 0,
                }
                sections.append(entry)
                stack.append(entry)
                continue  # heading text itself not counted as body words
            totals["paragraphs"] += 1
            add_counts(_count_words(text), len(text), text)
        else:
            totals["tables"] += 1
            from .read import get_table

            table = get_table(pkg, idx)
            table_text = " ".join(
                c["text"] for row in table["cells"] for c in row
            )
            add_counts(_count_words(table_text), len(table_text), table_text)

    result = {
        "totals": totals,
        "counting": (
            "spaces"
            if not totals["cjk_chars"]
            else ("cjk" if not non_cjk_words else "mixed")
        ),
    }
    if by_section:
        result["sections"] = sections
    return result
