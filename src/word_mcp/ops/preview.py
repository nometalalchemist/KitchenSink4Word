"""Dry-run preview for search_and_replace. The file is NEVER touched.

Uses the same matching engine as the real tool — the same runmap over
fragmented runs, the same guarded regex execution (ReDoS timeout), the same
per-paragraph / per-item sequencing in which later items see the results of
earlier items — so the previewed matches are exactly the matches a real run
would make. Nothing is written and no part is ever marked dirty; the module
asserts that on every call.

Intended flow: run preview_replace, review the matches, then run
search_and_replace with confidence — and with max_replacements set to the
previewed total, so any drift between preview and execution aborts the real
run instead of over-replacing.
"""

from __future__ import annotations

import re as _re

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from . import _regex, _runmap
from .read import body_items
from .text import _replace_parts

# XML 1.0-forbidden control characters (includes \x07, Word's internal
# table-cell separator) plus DEL 0x7F, which Word strips silently on open
# — same class the mail-merge guard refuses.
_BAD_CHARS_RE = _re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

_MATCH_DETAIL_CAP = 300
_CONTEXT = 60


def _validate_items(replacements: list[dict]) -> list[tuple[str, str, bool]]:
    if not replacements:
        raise WordMcpError(
            "give at least one replacement: {find, replace, regex?}"
        )
    items: list[tuple[str, str, bool]] = []
    for i, item in enumerate(replacements):
        if not isinstance(item, dict) or "find" not in item or "replace" not in item:
            raise WordMcpError(
                f"replacement {i} must be a dict with 'find' and 'replace' "
                "keys (optional 'regex': true)"
            )
        unknown = set(item) - {"find", "replace", "regex"}
        if unknown:
            raise WordMcpError(
                f"replacement {i} has unknown key(s) {sorted(unknown)}; "
                "allowed: find, replace, regex"
            )
        if not isinstance(item["find"], str) or not isinstance(
            item["replace"], str
        ):
            raise WordMcpError(
                f"replacement {i}: 'find' and 'replace' must be strings"
            )
        items.append((item["find"], item["replace"], bool(item.get("regex"))))
    return items


def _pre_flight(items: list[tuple[str, str, bool]]) -> list[dict]:
    """Problems the REAL run would hit, reported instead of hit.

    - invalid regex: compiled here so the error surfaces now (raised — the
      real run raises the identical error before changing anything);
    - zero-width-matchable regex: search_and_replace silently skips
      zero-length matches, which usually means the pattern is not doing what
      the author thinks — reported;
    - empty literal find: silently skipped by the real run — reported;
    - control characters in the replacement (\\x07 etc.): XML cannot store
      them, the real run would fail mid-flight — reported."""
    refusals: list[dict] = []
    for i, (find, repl, use_regex) in enumerate(items):
        if use_regex:
            _regex.compile_user_pattern(find)  # raises on invalid pattern
            if _regex.finditer(find, ""):
                refusals.append(
                    {
                        "item": i,
                        "find": find,
                        "problem": (
                            "regex can match the empty string; "
                            "search_and_replace silently skips zero-length "
                            "matches, so this pattern likely does not do "
                            "what you intend — anchor it"
                        ),
                    }
                )
        elif not find:
            refusals.append(
                {
                    "item": i,
                    "find": find,
                    "problem": (
                        "empty literal find string; search_and_replace "
                        "silently skips it (zero matches)"
                    ),
                }
            )
        if _BAD_CHARS_RE.search(repl):
            refusals.append(
                {
                    "item": i,
                    "find": find,
                    "problem": (
                        "replacement contains control characters (e.g. "
                        "\\x07) that XML cannot store; the real run would "
                        "fail — remove them before replacing"
                    ),
                }
            )
    return refusals


def _has_paragraph_ancestor(p) -> bool:
    """True for paragraphs nested inside another paragraph (text-box content
    inside w:txbxContent). The real engine's host-paragraph pass already
    covers their text (the runmap descends into them), so counting them
    again here would double-book every text-box match."""
    w_p = qn("w:p")
    parent = p.getparent()
    while parent is not None:
        if parent.tag == w_p:
            return True
        parent = parent.getparent()
    return False


def _item_matches(
    text: str, find: str, repl: str, use_regex: bool
) -> list[tuple[int, int, str]]:
    """Non-overlapping (start, end, replacement_text) — the exact match set
    search_and_replace computes for one (paragraph, item) pass."""
    if use_regex:
        return [
            (m.start(), m.end(), m.expand(repl))
            for m in _regex.finditer(find, text)
            if m.start() != m.end()
        ]
    if not find:
        return []
    out = []
    pos = 0
    while True:
        pos = text.find(find, pos)
        if pos < 0:
            break
        out.append((pos, pos + len(find), repl))
        pos += len(find)
    return out


def preview_replace(
    pkg: DocxPackage, replacements: list[dict], *, scope: str = "body"
) -> dict:
    """DRY RUN for search_and_replace: show exactly what a real run with the
    same replacements and scope would change, without touching the file.

    Each item: {find, replace, regex?: bool}; scope: body | footnotes |
    headers | all — identical to search_and_replace. Returns per-item match
    counts, a detailed match list (paragraph index, matched text, ~60 chars
    of context before and after the change), the grand total, and a
    `refusals` list of problems the real run would hit or silently absorb
    (invalid patterns raise, exactly as the real run would).

    Sequencing matches the real engine: within each paragraph, item N's
    matches are computed on the text AFTER items 1..N-1 have been applied,
    so chained replacements preview correctly.

    Recommended flow: run this, review the matches, then run
    search_and_replace with confidence (and max_replacements set to the
    previewed total, so any drift aborts the real run instead of
    over-replacing). Read-only — the file is never modified."""
    items = _validate_items(replacements)
    refusals = _pre_flight(items)
    parts = _replace_parts(pkg, scope)  # validates scope, same as real run

    dirty_before = set(pkg._dirty)

    per_item = [0] * len(items)
    matches: list[dict] = []
    truncated = False
    nested_seen = False

    for part in parts:
        if part == "word/document.xml":
            _keepalive = body_items(pkg)
            body_idx = {
                id(el): idx
                for kind, idx, el in _keepalive
                if kind == "paragraph"
            }
        else:
            body_idx = {}
        part_ordinal = 0
        for p in pkg.root(part).iter(qn("w:p")):
            if _has_paragraph_ancestor(p):
                nested_seen = True
                continue  # text-box paragraphs: covered by their host's pass
            text, _segments = _runmap.build_map(p)
            for i, (find, repl, use_regex) in enumerate(items):
                found = _item_matches(text, find, repl, use_regex)
                for start, end, actual in found:
                    per_item[i] += 1
                    if len(matches) >= _MATCH_DETAIL_CAP:
                        truncated = True
                        continue
                    entry: dict = {
                        "item": i,
                        "part": part,
                        "match": text[start:end],
                        "before": text[
                            max(0, start - _CONTEXT) : end + _CONTEXT
                        ],
                        "after": (
                            text[max(0, start - _CONTEXT) : start]
                            + actual
                            + text[end : end + _CONTEXT]
                        ),
                    }
                    if part == "word/document.xml":
                        idx = body_idx.get(id(p))
                        entry["paragraph_index"] = idx
                        if idx is None:
                            entry["location"] = (
                                "nested paragraph (table cell or text box)"
                            )
                    else:
                        entry["paragraph_in_part"] = part_ordinal
                    matches.append(entry)
                # Apply this item's matches to the STRING so the next item
                # sees post-replacement text — the real engine's semantics.
                for start, end, actual in reversed(found):
                    text = text[:start] + actual + text[end:]
            part_ordinal += 1

    # Read-only contract: nothing may have been marked dirty by this call.
    assert pkg._dirty == dirty_before, (
        "preview_replace must never dirty the package"
    )

    total = sum(per_item)
    if nested_seen:
        refusals.append(
            {
                "item": None,
                "find": None,
                "problem": (
                    "document contains text boxes; their text is previewed "
                    "through the host paragraph (matching the real engine's "
                    "primary pass). If a replacement string itself contains "
                    "a find pattern, the real run can re-apply it inside "
                    "boxes — avoid self-recreating replacements here."
                ),
            }
        )
    return {
        "scope": scope,
        "parts_searched": parts,
        "items": [
            {
                "find": find,
                "replace": repl,
                "regex": use_regex,
                "matches": per_item[i],
            }
            for i, (find, repl, use_regex) in enumerate(items)
        ],
        "total": total,
        "matches": matches,
        "matches_truncated": truncated,
        "refusals": refusals,
        "file_untouched": True,
        "next_step": (
            "review the matches, then run search_and_replace with the same "
            f"replacements and scope, and max_replacements={total} so any "
            "drift aborts instead of over-replacing"
        ),
    }
