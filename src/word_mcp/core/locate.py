"""The v2 location-object resolver (V2_DESIGN Section 6).

Every positional v2 tool takes a ``location`` dict carrying EXACTLY ONE
selector key plus an optional position modifier, and resolves it here
before acting:

    {"paragraph": 15}
    {"after_heading": {"text": "Chapter 3", "occurrence": 1, "match": "exact"}}
    {"outline": "3.2"}
    {"bookmark": "methodology_section"}
    {"search": {"text": "specific text", "occurrence": 1, "match_case": false}}
    {"anchor": "a3f9"}          (stub until the Phase 3 view layer)
    {"cursor": true}            (live mode only, via an injected reader)

    "position": "before" | "after" | "replace" | "start" | "end"
                (default "after")

Range-taking tools pass ``{"start": <location>, "end": <location>}`` to
resolve_range; both endpoints go through the same resolver and an inverted
range refuses with RANGE_OUT_OF_BOUNDS.

Design rules implemented here:

- Ambiguity is a loud refusal (Section 6.2). A text selector matching more
  than one place without an ``occurrence`` raises AmbiguousTarget carrying a
  ``matches`` list (paragraph index, outline path where derivable, context
  snippet); envelope.refusal surfaces it as error["matches"]. No caller in
  v2 acts on first-match.
- after_heading honors BOTH heading systems: built-in Heading styles AND
  w:outlineLvl (direct or style-inherited), the L8/NSU lesson, by reusing
  the ops/read.py detection helpers.
- The merge-test lesson: heading text recurs in body prose on real
  documents, so after_heading also scans prose for the text (always as a
  substring, whatever the heading match mode) and refuses ambiguous when
  prose recurrence exists and no occurrence was given. ``occurrence``
  counts MATCHING HEADINGS in document order; prose never resolves.
- search matches PLAIN text: XML entities are matched as their literal
  characters (Bug 11), and the not-found message hints at entity confusion,
  curly/straight quote mismatches (the L6 lesson), and cheap near-misses.
- paragraph index 0 addresses the default empty paragraph of a fresh
  document (the WS0 fix). Resolution NEVER mutates the package; ops-side
  consumers materialize the implicit paragraph when they act on it.
- cursor resolves through an injected ``cursor_reader`` callable so this
  module never imports COM; without one it refuses with WordNotRunning.

This module only READS the package. It imports read helpers from
ops/read.py (which itself depends only on core.package), so no import
cycle exists.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from typing import Any, Callable

from .errors import (
    AmbiguousTarget,
    RangeOutOfBounds,
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
    WordNotRunning,
)
from .package import DocxPackage, qn

SELECTORS = (
    "paragraph",
    "after_heading",
    "outline",
    "bookmark",
    "search",
    "anchor",
    "cursor",
)
POSITIONS = ("before", "after", "replace", "start", "end")
DEFAULT_POSITION = "after"

_MATCH_MODES = ("exact", "contains")
_ENTITY_RE = re.compile(r"&amp;|&lt;|&gt;|&quot;|&apos;")
_QUOTE_MAP = str.maketrans(
    {"‘": "'", "’": "'", "“": '"', "”": '"'}
)
_MAX_LISTED = 25  # cap on matches/outline entries echoed into messages


# ------------------------------------------------------------------ results


@dataclass(frozen=True)
class ResolvedLocation:
    """Where a location object landed, ready for an ops function to consume.

    paragraph_index: 0-based body paragraph index (index 0 is valid on a
        fresh document whose body holds no w:p yet; consumers materialize
        the implicit paragraph, per the WS0 fix; matched carries
        ``implicit: True`` in that case).
    position: before | after | replace | start | end.
    selector: the selector key that resolved it.
    matched: selector-specific info about the matched element (text or
        context, heading level / outline path / detected_via for headings,
        bookmark name, search span and matched text).
    char_start / char_end: search selector only, the matched span within
        the paragraph's plain text; None for whole-paragraph selectors.
    """

    paragraph_index: int
    position: str
    selector: str
    matched: dict[str, Any] = field(default_factory=dict)
    char_start: int | None = None
    char_end: int | None = None


@dataclass(frozen=True)
class ResolvedRange:
    """A resolved {start, end} pair, ordered (never inverted)."""

    start: ResolvedLocation
    end: ResolvedLocation

    @property
    def start_index(self) -> int:
        return self.start.paragraph_index

    @property
    def end_index(self) -> int:
        return self.end.paragraph_index


# ------------------------------------------------------------------ helpers


def _read():
    """ops/read.py, imported lazily to keep module import order flexible."""
    from ..ops import read

    return read


def _body_paragraphs(pkg: DocxPackage) -> list[tuple[int, Any]]:
    rd = _read()
    return [(idx, el) for kind, idx, el in rd.body_items(pkg) if kind == "paragraph"]


def _outline_entries(pkg: DocxPackage) -> list[dict]:
    """Headings in document order with their numbered outline paths.

    Detection reuses ops/read.py and therefore honors BOTH Heading styles
    and outlineLvl (direct or inherited through basedOn). Path "3.2" means
    the second level-2 heading under the third level-1. A heading whose
    parent level never appeared gets a 0 segment for that level (a level-2
    heading before any level-1 is "0.1"); paths stay unique either way.
    """
    rd = _read()
    style_outline = rd._style_outline_map(pkg)
    counters = [0] * 9
    entries: list[dict] = []
    for idx, el in _body_paragraphs(pkg):
        level, via = rd._outline_level_detected(el, style_outline)
        if level is None:
            continue
        text = rd.paragraph_text(el).strip()
        if not text:
            continue
        counters[level - 1] += 1
        for i in range(level, 9):
            counters[i] = 0
        path = ".".join(str(counters[i]) for i in range(level))
        entries.append(
            {
                "paragraph_index": idx,
                "level": level,
                "text": text,
                "path": path,
                "detected_via": via,
            }
        )
    return entries


def _snippet(text: str, start: int, end: int, radius: int = 30) -> str:
    lo = max(0, start - radius)
    hi = min(len(text), end + radius)
    out = text[lo:hi]
    if lo > 0:
        out = "..." + out
    if hi < len(text):
        out = out + "..."
    return out


def _clip(text: str, limit: int = 80) -> str:
    return text if len(text) <= limit else text[: limit - 3] + "..."


def _normalize_quotes(text: str) -> str:
    return text.translate(_QUOTE_MAP)


def _entity_hint(query: str) -> str | None:
    if _ENTITY_RE.search(query):
        return (
            "note: matching is against the PLAIN text, never the XML; "
            "write '&' not '&amp;', '<' not '&lt;', '>' not '&gt;'"
        )
    return None


def _ambiguous(message: str, matches: list[dict]) -> AmbiguousTarget:
    exc = AmbiguousTarget(message)
    exc.matches = matches  # envelope.refusal surfaces this as error["matches"]
    return exc


def _spec_dict(
    value: Any,
    name: str,
    example: str,
    required: dict[str, type],
    optional: dict[str, type],
) -> dict:
    if not isinstance(value, dict):
        raise WordMcpError(f"{name} selector takes an object like {example}")
    unknown = sorted(set(value) - set(required) - set(optional))
    if unknown:
        raise WordMcpError(
            f"{name} selector got unknown key(s) {unknown}; "
            f"it takes {sorted(required)} plus optional {sorted(optional)}"
        )
    for key, typ in required.items():
        if key not in value:
            raise WordMcpError(f"{name} selector requires {key!r}, like {example}")
        if not isinstance(value[key], typ) or isinstance(value[key], bool):
            raise WordMcpError(f"{name} selector: {key!r} must be a {typ.__name__}")
    for key, typ in optional.items():
        if key not in value:
            continue
        got = value[key]
        if typ is bool:
            if not isinstance(got, bool):
                raise WordMcpError(f"{name} selector: {key!r} must be a bool")
        elif not isinstance(got, typ) or isinstance(got, bool):
            raise WordMcpError(f"{name} selector: {key!r} must be a {typ.__name__}")
    return value


def _check_occurrence(occurrence: int, name: str) -> None:
    if occurrence < 1:
        raise WordMcpError(
            f"{name} selector: occurrence is 1-based, got {occurrence}"
        )


def _heading_resolved(entry: dict, position: str, selector: str) -> ResolvedLocation:
    return ResolvedLocation(
        paragraph_index=entry["paragraph_index"],
        position=position,
        selector=selector,
        matched={
            "paragraph": entry["paragraph_index"],
            "text": entry["text"],
            "level": entry["level"],
            "outline": entry["path"],
            "detected_via": entry["detected_via"],
        },
    )


# ------------------------------------------------------------------ selectors


def _resolve_paragraph(pkg: DocxPackage, value: Any, position: str) -> ResolvedLocation:
    if isinstance(value, bool) or not isinstance(value, int):
        raise WordMcpError(
            "paragraph selector takes a 0-based integer body paragraph index"
        )
    if value < 0:
        raise WordMcpError(f"paragraph index must be >= 0, got {value}")
    paras = _body_paragraphs(pkg)
    if not paras:
        if value == 0:
            # Fresh document: the body holds no w:p yet, but Word displays
            # one empty paragraph and index 0 addresses it (the WS0 fix).
            # Resolution never mutates; the consumer materializes it.
            return ResolvedLocation(
                0, position, "paragraph",
                {"paragraph": 0, "text": "", "implicit": True},
            )
        raise TargetNotFound(
            f"paragraph index {value} out of range, document has 1 paragraph "
            "(the default empty paragraph of a fresh document; only index 0 "
            "is valid)"
        )
    if value >= len(paras):
        raise TargetNotFound(
            f"paragraph index {value} out of range, document has {len(paras)} "
            f"paragraph(s) (valid indices 0-{len(paras) - 1})"
        )
    text = _read().paragraph_text(paras[value][1])
    return ResolvedLocation(
        value, position, "paragraph", {"paragraph": value, "text": _clip(text)}
    )


def _match_heading(entry_text: str, query: str, mode: str) -> bool:
    if mode == "exact":
        return entry_text == query
    return query in entry_text


def _heading_not_found(
    query: str, mode: str, entries: list[dict], prose: list[dict]
) -> TargetNotFound:
    parts = [f"no heading matched {query!r}"]
    if prose:
        where = ", ".join(str(m["paragraph"]) for m in prose[:5])
        parts.append(
            f"the text does occur in body prose (paragraph {where}); "
            "after_heading addresses headings, use search or paragraph "
            "for prose"
        )
    ent = _entity_hint(query)
    if ent:
        parts.append(ent)
    nq = _normalize_quotes(query)
    quote_hits = [
        e for e in entries
        if _match_heading(_normalize_quotes(e["text"]), nq, mode)
    ]
    if quote_hits:
        e = quote_hits[0]
        parts.append(
            f"a heading matches once curly and straight quotes are treated "
            f"alike: {e['text']!r} (paragraph {e['paragraph_index']}, outline "
            f"{e['path']}); Word autocorrects straight quotes to curly (use "
            "the curly characters)"
        )
    else:
        close = difflib.get_close_matches(
            query, [e["text"] for e in entries], n=2, cutoff=0.7
        )
        if close:
            parts.append(
                "close heading matches: " + ", ".join(repr(c) for c in close)
            )
    return TargetNotFound(". ".join(parts) + ".")


def _resolve_after_heading(
    pkg: DocxPackage, value: Any, position: str
) -> ResolvedLocation:
    spec = _spec_dict(
        value,
        "after_heading",
        '{"text": "Chapter 3", "occurrence": 1, "match": "exact"}',
        required={"text": str},
        optional={"occurrence": int, "match": str},
    )
    query = spec["text"]
    if not query.strip():
        raise WordMcpError("after_heading selector: 'text' must be non-empty")
    mode = spec.get("match", "exact")
    if mode not in _MATCH_MODES:
        raise WordMcpError(
            f"after_heading selector: match must be one of {_MATCH_MODES}, "
            f"got {mode!r}"
        )
    occurrence = spec.get("occurrence")
    if occurrence is not None:
        _check_occurrence(occurrence, "after_heading")

    entries = _outline_entries(pkg)
    heads = [e for e in entries if _match_heading(e["text"], query, mode)]

    # The merge-test lesson: heading text recurs in body prose on real
    # documents, so scan prose too (always substring, whatever the heading
    # match mode) and never resolve first-match past a recurrence.
    heading_indices = {e["paragraph_index"] for e in entries}
    prose: list[dict] = []
    rd = _read()
    for idx, el in _body_paragraphs(pkg):
        if idx in heading_indices:
            continue
        ptext = rd.paragraph_text(el)
        pos = ptext.find(query)
        if pos >= 0:
            prose.append(
                {
                    "paragraph": idx,
                    "context": _snippet(ptext, pos, pos + len(query)),
                }
            )

    if not heads:
        raise _heading_not_found(query, mode, entries, prose)
    if occurrence is not None:
        if occurrence > len(heads):
            raise TargetNotFound(
                f"after_heading {query!r}: occurrence {occurrence} out of "
                f"range, only {len(heads)} heading(s) match"
            )
        return _heading_resolved(heads[occurrence - 1], position, "after_heading")
    if len(heads) > 1 or prose:
        matches = [
            {
                "paragraph": e["paragraph_index"],
                "outline": e["path"],
                "context": _clip(e["text"]),
            }
            for e in heads
        ] + prose
        matches.sort(key=lambda m: m["paragraph"])
        raise _ambiguous(
            f"after_heading {query!r} matched {len(matches)} paragraphs "
            f"({len(heads)} heading(s), {len(prose)} in body prose). Pass "
            "occurrence (1-based, counting matching headings in document "
            "order), or address by outline/paragraph/anchor.",
            matches[:_MAX_LISTED],
        )
    return _heading_resolved(heads[0], position, "after_heading")


def _resolve_outline(pkg: DocxPackage, value: Any, position: str) -> ResolvedLocation:
    if not isinstance(value, str) or not re.fullmatch(r"\d+(\.\d+)*", value):
        raise WordMcpError(
            'outline selector takes a numbered path string like "3.2" '
            "(second level-2 heading under the third level-1)"
        )
    entries = _outline_entries(pkg)
    for e in entries:
        if e["path"] == value:
            return _heading_resolved(e, position, "outline")
    listing = "; ".join(
        f"{e['path']} {_clip(e['text'], 40)!r}" for e in entries[:_MAX_LISTED]
    )
    more = (
        f" (and {len(entries) - _MAX_LISTED} more)"
        if len(entries) > _MAX_LISTED
        else ""
    )
    raise TargetNotFound(
        f"no heading at outline path {value!r}; document outline: "
        + (listing + more if entries else "no headings detected")
    )


def _resolve_bookmark(pkg: DocxPackage, value: Any, position: str) -> ResolvedLocation:
    if not isinstance(value, str) or not value:
        raise WordMcpError("bookmark selector takes a bookmark name (string)")
    root = pkg.root()
    target = None
    names: list[str] = []
    for bs in root.iter(qn("w:bookmarkStart")):
        name = bs.get(qn("w:name"))
        if name:
            names.append(name)
        if name == value and target is None:
            target = bs
    if target is None:
        msg = (
            f"no bookmark named {value!r}; document has {len(names)} "
            "bookmark(s)"
        )
        ci = [n for n in names if n.lower() == value.lower()]
        if ci:
            msg += f". Bookmark names are case-sensitive: did you mean {ci[0]!r}?"
        else:
            close = difflib.get_close_matches(value, names, n=3, cutoff=0.6)
            if close:
                msg += ". Close matches: " + ", ".join(repr(c) for c in close)
        raise TargetNotFound(msg)

    body = pkg.body()
    # Keep the paragraph list ALIVE and compare proxies by identity: lxml
    # keeps at most one live proxy per node, but id() of a dead proxy gets
    # reused, so an id-keyed map over a discarded list mis-resolves.
    paras = _body_paragraphs(pkg)

    def _index_of(el) -> int | None:
        for idx, cand in paras:
            if cand is el:
                return idx
        return None

    node = target
    while node.getparent() is not None and node.getparent() is not body:
        node = node.getparent()
    rd = _read()
    if node.tag == qn("w:p") and _index_of(node) is not None:
        idx = _index_of(node)
        return ResolvedLocation(
            idx, position, "bookmark",
            {
                "paragraph": idx,
                "bookmark": value,
                "text": _clip(rd.paragraph_text(node)),
            },
        )
    if node.tag == qn("w:tbl"):
        raise UnsupportedStructure(
            f"bookmark {value!r} is inside a table; positional tools address "
            "body paragraphs. Use the table tools (set_cells, get_table) for "
            "table content."
        )
    if node is target:
        # bookmarkStart sitting directly at body level: attach to the
        # nearest following body paragraph, else the nearest preceding one.
        for step, key in (("getnext", "following"), ("getprevious", "preceding")):
            cand = getattr(node, step)()
            while cand is not None:
                idx = _index_of(cand) if cand.tag == qn("w:p") else None
                if idx is not None:
                    return ResolvedLocation(
                        idx, position, "bookmark",
                        {
                            "paragraph": idx,
                            "bookmark": value,
                            "text": _clip(rd.paragraph_text(cand)),
                            "adjacency": key,
                        },
                    )
                cand = getattr(cand, step)()
        return ResolvedLocation(
            0, position, "bookmark",
            {"paragraph": 0, "bookmark": value, "text": "", "implicit": True},
        )
    raise UnsupportedStructure(
        f"bookmark {value!r} is not inside a body paragraph (it sits inside "
        "a nested structure this resolver does not address); use a "
        "paragraph index or search instead."
    )


def _resolve_search(pkg: DocxPackage, value: Any, position: str) -> ResolvedLocation:
    spec = _spec_dict(
        value,
        "search",
        '{"text": "specific text", "occurrence": 1, "match_case": false}',
        required={"text": str},
        optional={"occurrence": int, "match_case": bool},
    )
    query = spec["text"]
    if not query:
        raise WordMcpError("search selector: 'text' must be non-empty")
    match_case = spec.get("match_case", False)
    occurrence = spec.get("occurrence")
    if occurrence is not None:
        _check_occurrence(occurrence, "search")

    rd = _read()
    paras = _body_paragraphs(pkg)
    outline_by_idx = {
        e["paragraph_index"]: e["path"] for e in _outline_entries(pkg)
    }

    def _spans(text: str, needle: str) -> list[tuple[int, int]]:
        if match_case:
            return rd._literal_spans(text, needle)
        return rd._literal_spans(text.lower(), needle.lower())

    hits: list[dict] = []
    for idx, el in paras:
        ptext = rd.paragraph_text(el)
        for s, e in _spans(ptext, query):
            entry = {
                "paragraph": idx,
                "match": ptext[s:e],
                "char_start": s,
                "char_end": e,
                "context": _snippet(ptext, s, e),
            }
            if idx in outline_by_idx:
                entry["outline"] = outline_by_idx[idx]
            hits.append(entry)

    if not hits:
        parts = [f"search text not found: {query!r}"]
        ent = _entity_hint(query)
        if ent:
            parts.append(ent)
        nq = _normalize_quotes(query)
        quote_hits: list[int] = []
        for idx, el in paras:
            if _spans(_normalize_quotes(rd.paragraph_text(el)), nq):
                quote_hits.append(idx)
        if quote_hits:
            where = ", ".join(str(i) for i in quote_hits[:5])
            parts.append(
                f"the text matches once curly and straight quotes are "
                f"treated alike (paragraph {where}); Word autocorrects "
                "straight quotes to curly, so search with the curly "
                "characters"
            )
        if match_case:
            ci_hits = [
                idx for idx, el in paras
                if rd._literal_spans(
                    rd.paragraph_text(el).lower(), query.lower()
                )
            ]
            if ci_hits:
                where = ", ".join(str(i) for i in ci_hits[:5])
                parts.append(
                    f"matches exist ignoring case (paragraph {where}); pass "
                    "match_case: false or fix the casing"
                )
        raise TargetNotFound(". ".join(parts) + ".")

    if occurrence is not None:
        if occurrence > len(hits):
            raise TargetNotFound(
                f"search {query!r}: occurrence {occurrence} out of range, "
                f"only {len(hits)} match(es)"
            )
        chosen = hits[occurrence - 1]
    elif len(hits) > 1:
        raise _ambiguous(
            f"search {query!r} matched {len(hits)} places. Pass occurrence "
            "(1-based, document order), or address by "
            "outline/paragraph/anchor.",
            [
                {
                    k: h[k]
                    for k in ("paragraph", "outline", "context")
                    if k in h
                }
                for h in hits[:_MAX_LISTED]
            ],
        )
    else:
        chosen = hits[0]
    return ResolvedLocation(
        paragraph_index=chosen["paragraph"],
        position=position,
        selector="search",
        matched=chosen,
        char_start=chosen["char_start"],
        char_end=chosen["char_end"],
    )


def _resolve_anchor_stub(value: Any) -> ResolvedLocation:
    if not isinstance(value, str) or not value:
        raise WordMcpError(
            "anchor selector takes a get_document_view anchor id (string)"
        )
    raise WordMcpError(
        "anchor addressing is not available yet: anchors are issued by "
        "get_document_view and arrive with the view/batch layer (v2 "
        "Phase 3). Address by outline, paragraph, bookmark, or search for "
        "now."
    )


def _resolve_cursor(
    pkg: DocxPackage,
    value: Any,
    position: str,
    cursor_reader: Callable[[], int] | None,
) -> ResolvedLocation:
    if value is not True:
        raise WordMcpError('cursor selector takes true, as in {"cursor": true}')
    if cursor_reader is None:
        raise WordNotRunning(
            "cursor addressing needs a live Word session with this document "
            "open and active; no cursor reader is available in file mode. "
            "Address by paragraph, outline, bookmark, or search instead."
        )
    idx = cursor_reader()
    if isinstance(idx, bool) or not isinstance(idx, int) or idx < 0:
        raise WordMcpError(
            f"cursor reader returned {idx!r}; expected a 0-based body "
            "paragraph index"
        )
    paras = _body_paragraphs(pkg)
    count = len(paras) if paras else 1  # fresh doc displays one paragraph
    if idx >= count:
        raise TargetNotFound(
            f"cursor reports paragraph {idx} but the document has {count} "
            "paragraph(s); the file and the open document may be out of sync"
        )
    text = _read().paragraph_text(paras[idx][1]) if paras else ""
    return ResolvedLocation(
        idx, position, "cursor",
        {"paragraph": idx, "via": "cursor", "text": _clip(text)},
    )


# ------------------------------------------------------------------ public API


def is_range_spec(obj: Any) -> bool:
    """True when obj looks like a {start, end} range rather than a single
    location object (used by Phase 2 wrappers on dual-shape parameters)."""
    return (
        isinstance(obj, dict)
        and ("start" in obj or "end" in obj)
        and not any(k in obj for k in SELECTORS)
    )


def resolve_location(
    pkg: DocxPackage,
    location: Any,
    *,
    cursor_reader: Callable[[], int] | None = None,
) -> ResolvedLocation:
    """Resolve one location object against an open package.

    location: dict with EXACTLY ONE selector key from SELECTORS plus an
        optional "position" (default "after"). Zero or multiple selector
        keys refuse (BAD_PARAMS).
    cursor_reader: live-mode injection for the cursor selector; a callable
        returning the 0-based body paragraph index of the Word selection
        start. None (file mode) makes {"cursor": true} refuse with
        WordNotRunning. This module never imports COM.

    Raises WordMcpError (BAD_PARAMS), TargetNotFound (NOT_FOUND),
    AmbiguousTarget with a .matches list (AMBIGUOUS_LOCATION),
    UnsupportedStructure (UNSUPPORTED_CONTENT), or WordNotRunning.
    """
    if not isinstance(location, dict):
        raise WordMcpError(
            "location must be an object with exactly one selector key from "
            f"{list(SELECTORS)} plus optional 'position'"
        )
    unknown = sorted(set(location) - set(SELECTORS) - {"position"})
    if unknown:
        if "start" in location or "end" in location:
            raise WordMcpError(
                "this looks like a {start, end} range, not a single "
                "location; pass it to a range-taking parameter, or pass one "
                "location object with a single selector key"
            )
        raise WordMcpError(
            f"unknown location key(s) {unknown}; selectors are "
            f"{list(SELECTORS)}, plus optional 'position'"
        )
    present = [k for k in SELECTORS if k in location]
    if len(present) != 1:
        raise WordMcpError(
            "location needs exactly one selector key, got "
            f"{len(present)} ({present if present else 'none'}); selectors "
            f"are {list(SELECTORS)}"
        )
    position = location.get("position", DEFAULT_POSITION)
    if position not in POSITIONS:
        raise WordMcpError(
            f"position must be one of {list(POSITIONS)}, got {position!r}"
        )
    sel = present[0]
    value = location[sel]
    if sel == "paragraph":
        return _resolve_paragraph(pkg, value, position)
    if sel == "after_heading":
        return _resolve_after_heading(pkg, value, position)
    if sel == "outline":
        return _resolve_outline(pkg, value, position)
    if sel == "bookmark":
        return _resolve_bookmark(pkg, value, position)
    if sel == "search":
        return _resolve_search(pkg, value, position)
    if sel == "anchor":
        return _resolve_anchor_stub(value)
    return _resolve_cursor(pkg, value, position, cursor_reader)


def resolve_range(
    pkg: DocxPackage,
    range_spec: Any,
    *,
    cursor_reader: Callable[[], int] | None = None,
) -> ResolvedRange:
    """Resolve a {"start": <location>, "end": <location>} range.

    Both endpoints go through resolve_location (same refusal behavior).
    An inverted range (end before start, by paragraph index, or by char
    span when both endpoints are search hits in the same paragraph)
    refuses with RangeOutOfBounds (RANGE_OUT_OF_BOUNDS).
    """
    if not isinstance(range_spec, dict):
        raise WordMcpError(
            'range must be an object like {"start": <location>, '
            '"end": <location>}'
        )
    unknown = sorted(set(range_spec) - {"start", "end"})
    if unknown:
        raise WordMcpError(
            f"unknown range key(s) {unknown}; a range takes exactly "
            "'start' and 'end', each a location object"
        )
    for key in ("start", "end"):
        if key not in range_spec:
            raise WordMcpError(
                f"range is missing {key!r}; a range takes 'start' and "
                "'end', each a location object"
            )
    start = resolve_location(
        pkg, range_spec["start"], cursor_reader=cursor_reader
    )
    end = resolve_location(pkg, range_spec["end"], cursor_reader=cursor_reader)
    if end.paragraph_index < start.paragraph_index:
        raise RangeOutOfBounds(
            f"inverted range: end (paragraph {end.paragraph_index}) precedes "
            f"start (paragraph {start.paragraph_index})"
        )
    if (
        end.paragraph_index == start.paragraph_index
        and start.char_start is not None
        and end.char_start is not None
        and end.char_start < start.char_start
    ):
        raise RangeOutOfBounds(
            f"inverted range: within paragraph {start.paragraph_index} the "
            f"end span (char {end.char_start}) precedes the start span "
            f"(char {start.char_start})"
        )
    return ResolvedRange(start=start, end=end)
