"""Redaction: TRUE removal of matched text from a .docx package.

Unlike a black highlight or hidden-text formatting (which leave the text in
the XML for anyone to extract), redact_text REPLACES the matched characters
in the underlying XML with the replacement string. Matching is runmap-safe:
a secret Word has fragmented across several runs (formatting boundaries,
revision-save ids) is still found and removed as one match.

What is scrubbed (location classes):
- body text including tables (word/document.xml)
- headers and footers
- footnotes and endnotes
- comment text (word/comments.xml)
- document properties (docProps/core.xml, app.xml, custom.xml text values)
- hyperlink display text (regular runs, covered by the story pass),
  hyperlink tooltips (w:tooltip), and hyperlink URL targets in .rels parts
- complex-field instruction text (w:instrText) and cached field results
  (counted separately, since a stale cached result can leak after the
  instruction is cleaned)
- tracked-change deleted text (w:delText, node-level; see caveat below)

What is NOT examined — reported under "not_examined" so nobody mistakes this
for image redaction:
- embedded images and media (pixels are never OCRed; text drawn in a scanned
  page or screenshot is NOT removed)
- charts, OLE/embedded objects, VBA projects, thumbnails, embedded fonts,
  and any other binary part

Caveats reported honestly rather than papered over:
- Deleted tracked-change text is scrubbed node by node; a match split across
  two deleted fragments can survive. The post-redaction verification pass
  concatenates deleted text per paragraph, so such a survivor flips
  verified_clean to False with a location. Accept/reject revisions first for
  a guaranteed-clean result.
- The verification scan is conservative: it also checks concatenated
  visible+hidden text per paragraph and every XML attribute, so it can flag
  a residual that is not actually contiguous on screen. A false alarm is
  preferred over a false all-clear.

Atomicity: all targets are validated before anything is touched, and nothing
in this module writes to disk — the caller saves via DocxPackage.save, which
is atomic and validated. Any error leaves the original file untouched.
"""

from __future__ import annotations

import re

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from . import _regex, _runmap

# ---------------------------------------------------------------- validation


def _validate_targets(targets: list[dict]) -> list[tuple[str, bool]]:
    """Normalize and validate every target BEFORE any mutation (atomicity).

    Refuses empty literals and regexes that can match the empty string
    (mirrors the find_text guard: a zero-width match would 'redact' at every
    position and rewrite the whole document into replacement strings).
    """
    if not targets:
        raise WordMcpError("give at least one target: {find, regex?}")
    norm: list[tuple[str, bool]] = []
    for i, t in enumerate(targets):
        if not isinstance(t, dict) or "find" not in t:
            raise WordMcpError(
                f"target {i} must be a dict with a 'find' key "
                "(optional 'regex': true)"
            )
        find = t["find"]
        unknown = set(t) - {"find", "regex"}
        if unknown:
            raise WordMcpError(
                f"target {i} has unknown key(s) {sorted(unknown)}; "
                "allowed: find, regex"
            )
        if not isinstance(find, str):
            raise WordMcpError(f"target {i}: 'find' must be a string")
        use_regex = bool(t.get("regex"))
        if use_regex:
            if _regex.finditer(find, ""):
                raise WordMcpError(
                    f"regex {find!r} can match the empty string and would "
                    "match at every position; anchor the pattern — nothing "
                    "was changed"
                )
        elif not find:
            raise WordMcpError(
                f"target {i}: empty literal find string refused — nothing "
                "was changed"
            )
        norm.append((find, use_regex))
    return norm


def _spans(text: str, find: str, use_regex: bool) -> list[tuple[int, int]]:
    """Non-overlapping match spans of one pattern in `text`."""
    if not text:
        return []
    if use_regex:
        return [
            (m.start(), m.end())
            for m in _regex.finditer(find, text)
            if m.start() != m.end()
        ]
    spans: list[tuple[int, int]] = []
    pos = 0
    while True:
        pos = text.find(find, pos)
        if pos < 0:
            break
        spans.append((pos, pos + len(find)))
        pos += len(find)
    return spans


def _apply_to_string(
    value: str, patterns: list[tuple[str, bool]], replacement: str
) -> tuple[str, int]:
    """Replace every match of every pattern in a plain string. Spans are
    computed per pattern before splicing and applied right-to-left, so a
    replacement containing its own find string cannot loop."""
    total = 0
    for find, use_regex in patterns:
        spans = _spans(value, find, use_regex)
        for start, end in reversed(spans):
            value = value[:start] + replacement + value[end:]
        total += len(spans)
    return value, total


# ------------------------------------------------------------- part routing

_HEADER_FOOTER_RE = re.compile(r"word/(header|footer)\d+\.xml")


def _all_story_parts(pkg: DocxPackage) -> dict[str, list[str]]:
    """Every story part present, grouped by scope keyword."""
    return {
        "body": ["word/document.xml"],
        "headers": [
            p for p in pkg.part_names() if _HEADER_FOOTER_RE.fullmatch(p)
        ],
        "footnotes": [
            p
            for p in ("word/footnotes.xml", "word/endnotes.xml")
            if pkg.has_part(p)
        ],
    }


def _scoped_parts(pkg: DocxPackage, scope: str) -> tuple[list[str], list[str]]:
    """(parts in scope, story parts EXCLUDED by scope). Unknown scope refused."""
    groups = _all_story_parts(pkg)
    if scope == "all":
        keys = ["body", "headers", "footnotes"]
    elif scope in groups:
        keys = [scope]
    else:
        raise WordMcpError(
            f"unknown scope: {scope!r}; use body | headers | footnotes | all"
        )
    in_scope = [p for k in keys for p in groups[k]]
    excluded = [p for k in groups for p in groups[k] if p not in in_scope]
    return in_scope, excluded


def _story_class(part: str) -> str:
    if part == "word/document.xml":
        return "body"
    if part == "word/footnotes.xml":
        return "footnotes"
    if part == "word/endnotes.xml":
        return "endnotes"
    if part == "word/comments.xml":
        return "comments"
    return "headers_footers"


# ------------------------------------------------------------ field mapping


def _field_result_runs(root: etree._Element) -> set[int]:
    """ids of w:r elements sitting inside a complex field's CACHED RESULT
    (between a fldChar 'separate' and its 'end'). iter() is document order:
    a run is yielded before its own fldChar child, so the runs holding the
    begin/separate markers are correctly excluded and the runs holding the
    cached text are included.

    Returns (id_set, keepalive): lxml proxies are EPHEMERAL — without the
    keepalive list their ids get recycled by later proxies and body counts
    misattribute to field_results (found by the Wave B adversarial round)."""
    result: set[int] = set()
    keepalive: list = []
    state: list[str] = []  # one entry per open (possibly nested) field
    for el in root.iter():
        tag = etree.QName(el).localname
        if tag == "fldChar":
            t = el.get(qn("w:fldCharType"))
            if t == "begin":
                state.append("instr")
            elif t == "separate":
                if state:
                    state[-1] = "result"
            elif t == "end":
                if state:
                    state.pop()
        elif tag == "r" and state and state[-1] == "result":
            result.add(id(el))
            keepalive.append(el)
    return result, keepalive


# ----------------------------------------------------------- scrubbing passes


def _redact_story_part(
    pkg: DocxPackage,
    part: str,
    patterns: list[tuple[str, bool]],
    replacement: str,
    counts: dict[str, int],
) -> None:
    """Runmap-safe redaction of one story part: visible run text (fragmented
    runs handled), instruction text, deleted text, hyperlink tooltips."""
    root = pkg.root(part)
    part_class = _story_class(part)
    field_result_runs, _fr_keepalive = _field_result_runs(root)
    dirty = False

    def bump(cls: str, n: int = 1) -> None:
        counts[cls] = counts.get(cls, 0) + n

    # Visible text, per paragraph. One runmap snapshot per (paragraph,
    # pattern); matches applied right-to-left so earlier offsets stay valid
    # and replacement text is never re-matched (mirrors search_and_replace).
    for p in root.iter(qn("w:p")):
        for find, use_regex in patterns:
            text, segments = _runmap.build_map(p)
            for start, end in reversed(_spans(text, find, use_regex)):
                affected = [
                    s for s in segments if s.start < end and s.end > start
                ]
                in_result = any(
                    id(s.run) in field_result_runs for s in affected
                )
                _runmap.replace_range(p, segments, start, end, replacement)
                bump("field_results" if in_result else part_class)
                dirty = True

    # Field instruction text (HYPERLINK URLs, mail-merge field names...).
    # instrText is invisible to the runmap; scrubbed node by node.
    for it in root.iter(qn("w:instrText")):
        new, n = _apply_to_string(it.text or "", patterns, replacement)
        if n:
            it.text = new
            bump("field_instructions", n)
            dirty = True

    # Tracked-change deleted text: still present in the XML, invisible on
    # screen. Node-level scrub; cross-fragment matches are caught by the
    # verification pass instead of guessed at here.
    for dt in root.iter(qn("w:delText")):
        new, n = _apply_to_string(dt.text or "", patterns, replacement)
        if n:
            dt.text = new
            bump("tracked_deletions", n)
            dirty = True

    # Hyperlink tooltips (attribute, not run text).
    for link in root.iter(qn("w:hyperlink")):
        tip = link.get(qn("w:tooltip"))
        if tip:
            new, n = _apply_to_string(tip, patterns, replacement)
            if n:
                link.set(qn("w:tooltip"), new)
                bump("hyperlink_tooltips", n)
                dirty = True

    if dirty:
        pkg.mark_dirty(part)


_DOCPROPS_PARTS = ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml")


def _redact_doc_properties(
    pkg: DocxPackage,
    patterns: list[tuple[str, bool]],
    replacement: str,
    counts: dict[str, int],
) -> None:
    """Scrub every text value in the docProps parts: title, subject,
    keywords, description (core.xml) plus the app.xml fields Word derives
    from content (TitlesOfParts caches heading text) and custom properties."""
    for part in _DOCPROPS_PARTS:
        if not pkg.has_part(part):
            continue
        dirty = False
        for el in pkg.root(part).iter():
            if el.text:
                new, n = _apply_to_string(el.text, patterns, replacement)
                if n:
                    el.text = new
                    counts["doc_properties"] = (
                        counts.get("doc_properties", 0) + n
                    )
                    dirty = True
        if dirty:
            pkg.mark_dirty(part)


_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _redact_hyperlink_rels(
    pkg: DocxPackage,
    patterns: list[tuple[str, bool]],
    replacement: str,
    counts: dict[str, int],
) -> None:
    """Scrub hyperlink URL targets in every .rels part. A URL whose matched
    portion is removed stops resolving — in a redaction that is the point.
    A target left empty becomes about:blank so the XML stays valid."""
    for part in pkg.part_names():
        if not part.endswith(".rels"):
            continue
        dirty = False
        for rel in pkg.root(part).findall(f"{{{_REL_NS}}}Relationship"):
            if not (rel.get("Type") or "").endswith("/hyperlink"):
                continue
            target = rel.get("Target") or ""
            new, n = _apply_to_string(target, patterns, replacement)
            if n:
                rel.set("Target", new if new else "about:blank")
                counts["hyperlink_urls"] = counts.get("hyperlink_urls", 0) + n
                dirty = True
        if dirty:
            pkg.mark_dirty(part)


# ------------------------------------------------------------- verification


def _under_paragraph(el: etree._Element) -> bool:
    w_p = qn("w:p")
    parent = el.getparent()
    while parent is not None:
        if parent.tag == w_p:
            return True
        parent = parent.getparent()
    return False


def _paragraph_all_text(p: etree._Element) -> str:
    """Every descendant text node of a paragraph concatenated in document
    order — visible text, deleted text, and instruction text together. Used
    only for verification: catches matches fragmented across ANY node kind,
    at the cost of possible false positives across node boundaries."""
    return "".join(t for t in p.itertext())


def _scan_package(
    pkg: DocxPackage, patterns: list[tuple[str, bool]]
) -> tuple[list[dict], int]:
    """Re-scan every XML part for the patterns: paragraph visible text
    (runmap, so fragmented runs are caught), paragraph concatenated hidden
    text, every free-standing text node, and every attribute value.
    Returns (deduped findings, parts scanned)."""
    findings: list[dict] = []
    seen: set[tuple[str, str, str]] = set()
    scanned = 0

    def hit(part: str, where: str, find: str) -> None:
        key = (part, where, find)
        if key not in seen:
            seen.add(key)
            findings.append({"part": part, "where": where, "pattern": find})

    for name in pkg.part_names():
        if not name.endswith((".xml", ".rels")):
            continue
        try:
            root = pkg.root(name)
        except Exception:
            hit(name, "part could not be parsed for scanning", "*")
            continue
        scanned += 1
        for p in root.iter(qn("w:p")):
            visible, _ = _runmap.build_map(p)
            hidden = _paragraph_all_text(p)
            for find, use_regex in patterns:
                if _spans(visible, find, use_regex):
                    hit(name, "paragraph visible text", find)
                elif _spans(hidden, find, use_regex):
                    hit(
                        name,
                        "paragraph hidden text (deleted/instruction/"
                        "concatenated fragments)",
                        find,
                    )
        for el in root.iter():
            tag = etree.QName(el).localname
            for attr, value in el.attrib.items():
                for find, use_regex in patterns:
                    if _spans(value, find, use_regex):
                        hit(
                            name,
                            f"attribute {etree.QName(attr).localname} "
                            f"on <{tag}>",
                            find,
                        )
            if el.text and not _under_paragraph(el) and tag != "p":
                for find, use_regex in patterns:
                    if _spans(el.text, find, use_regex):
                        hit(name, f"text of <{tag}>", find)
    return findings, scanned


# ------------------------------------------------------- honesty accounting

_BINARY_CLASSES = (
    ("word/media/", "embedded images/media"),
    ("word/charts/", "charts"),
    ("word/embeddings/", "OLE/embedded objects"),
    ("word/vbaProject", "VBA macro project"),
    ("docProps/thumbnail", "document thumbnail"),
    ("word/fonts/", "embedded fonts"),
)


def _not_examined(pkg: DocxPackage, excluded_parts: list[str]) -> list[dict]:
    """Location classes this run did NOT scrub, listed so the caller cannot
    mistake text redaction for full-content redaction."""
    out: list[dict] = []
    covered: set[str] = set()
    for prefix, label in _BINARY_CLASSES:
        parts = [n for n in pkg.part_names() if n.startswith(prefix)]
        if parts:
            covered.update(parts)
            out.append(
                {
                    "class": label,
                    "parts": len(parts),
                    "reason": (
                        "binary content — pixels/objects are never examined "
                        "or OCRed; text rendered inside is NOT redacted"
                    ),
                }
            )
    other_binary = [
        n
        for n in pkg.part_names()
        if not n.endswith((".xml", ".rels")) and n not in covered
    ]
    if other_binary:
        out.append(
            {
                "class": "other binary parts",
                "parts": len(other_binary),
                "names": other_binary[:20],
                "reason": "binary content — not examined",
            }
        )
    if excluded_parts:
        out.append(
            {
                "class": "story parts excluded by scope",
                "names": excluded_parts,
                "reason": (
                    "outside the requested scope — verification still scans "
                    "them, so a secret there flips verified_clean to False"
                ),
            }
        )
    return out


# ------------------------------------------------------------------- public


def redact_text(
    pkg: DocxPackage,
    targets: list[dict],
    *,
    replacement: str = "[REDACTED]",
    scope: str = "all",
) -> dict:
    """Permanently remove every match of every target from the document.

    Each target: {find, regex?: bool}. Literal matching by default; regex
    targets run through the guarded engine (timeout, zero-width refusal).
    scope limits the STORY parts (body | headers | footnotes | all);
    comments, document properties, and hyperlink URL targets are always
    scrubbed regardless of scope, because leaving a secret in metadata while
    'redacting' the body would be a false promise.

    The replacement inherits the formatting of the first character it
    replaces. Matches cannot span paragraphs (same contract as
    search_and_replace).

    Returns per-location-class counts, the classes scrubbed, the classes NOT
    examined (images, charts, OLE objects...), and verified_clean — the
    result of a full post-redaction re-scan of every XML part. This is TEXT
    redaction only; it does not touch pixels.
    """
    patterns = _validate_targets(targets)
    if not isinstance(replacement, str):
        raise WordMcpError("replacement must be a string (may be empty)")
    story_parts, excluded = _scoped_parts(pkg, scope)

    counts: dict[str, int] = {}
    scrub_parts = list(story_parts)
    if pkg.has_part("word/comments.xml"):
        scrub_parts.append("word/comments.xml")
    for part in scrub_parts:
        _redact_story_part(pkg, part, patterns, replacement, counts)
    _redact_doc_properties(pkg, patterns, replacement, counts)
    _redact_hyperlink_rels(pkg, patterns, replacement, counts)

    residual, parts_scanned = _scan_package(pkg, patterns)
    scrubbed_classes = [
        "body text incl. tables" if "word/document.xml" in story_parts else None,
        "headers/footers" if any(_HEADER_FOOTER_RE.fullmatch(p) for p in story_parts) else None,
        "footnotes/endnotes" if any(p.startswith("word/footnotes") or p.startswith("word/endnotes") for p in story_parts) else None,
        "comments" if "word/comments.xml" in scrub_parts else None,
        "document properties (docProps core/app/custom)",
        "hyperlink display text, tooltips, and URL targets",
        "field instruction text and cached field results",
        "tracked-change deleted text (node-level; verification catches "
        "cross-fragment survivors)",
    ]
    result = {
        "redacted": counts,
        "total": sum(counts.values()),
        "replacement": replacement,
        "scope": scope,
        "scrubbed_location_classes": [c for c in scrubbed_classes if c],
        "not_examined": _not_examined(pkg, excluded),
        "verified_clean": not residual,
        "parts_scanned": parts_scanned,
        "note": (
            "text redaction only — content stored as pixels (scanned pages, "
            "screenshots, charts rendered as images) is never examined"
        ),
    }
    if residual:
        result["residual"] = residual
    return result


def verify_redaction(pkg: DocxPackage, targets: list[dict]) -> dict:
    """Standalone re-scan: do any of the target patterns still appear
    anywhere in the package's XML? Use it to check a third-party file, or a
    file redacted by another tool. Scans paragraph visible text (fragmented
    runs included), hidden text (deleted/instruction), free-standing text
    nodes, and every attribute in every XML part. Read-only.

    clean=True means no XML text or attribute matched; binary parts (images,
    OLE objects) are listed under not_examined, never silently trusted."""
    patterns = _validate_targets(targets)
    residual, parts_scanned = _scan_package(pkg, patterns)
    return {
        "clean": not residual,
        "residual": residual,
        "patterns_checked": len(patterns),
        "parts_scanned": parts_scanned,
        "not_examined": _not_examined(pkg, []),
    }
