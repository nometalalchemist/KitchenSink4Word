"""Peer-review anonymization: reversible masking of self-citations.

Distinct from redaction (ops/redaction.py destroys text permanently): this
module masks the AUTHOR'S OWN citations and reference entries the way journal
submission guidelines ask ("Author, 1999" / "Author (1999). [Details removed
for peer review.]"), scrubs identifying metadata, and writes a reversal
mapping so the accepted manuscript can be restored exactly.

What it edits automatically:
- self-citations in body/footnote/endnote text: "Hurd (1999)" -> "Author
  (1999)", "(Hurd, 1999)" -> "(Author, 1999)", "Hurd & Lake (2005)" ->
  "Author (2005)" — the year (and any page numbers) are kept;
- matching entries in the reference list -> "Author (1999). [Details removed
  for peer review.]";
- document metadata (creator, lastModifiedBy, Company/Manager) via the same
  scrub the submission-prep pass uses (ops/cleanup.py).

What it deliberately does NOT edit — flagged with locations instead:
self-identifying prose ("my previous work", "our earlier study", "this
author"), an Acknowledgments section, and any leftover mention of the
author's surname outside citation syntax. Rewriting prose changes meaning;
that is the human's job, and the flags say exactly where to look.

The mapping file re-identifies the author by construction. Keep it PRIVATE —
never upload it with the submission.
"""

from __future__ import annotations

import datetime as _dt
import json
import re
from pathlib import Path

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from ..core.sandbox import check_path
from . import _runmap
from .citecheck import _REF_HEADINGS, _REF_YEAR, _YEAR
from .cleanup import _CP, _DC, _EP, _empty_text
from .read import (
    _outline_level,
    _style_outline_map,
    body_items,
    paragraph_text,
)

MAPPING_FORMAT = "word-mcp.anonymization/1"

# A citation author group: "Hurd", "Hurd and Lake", "Hurd, Lake, & Smith",
# "Hurd et al."
_AUTHOR_GROUP = (
    r"[A-Z][A-Za-z'’\-]+"
    r"(?:(?:,\s*|,?\s+(?:and|&)\s+)[A-Z][A-Za-z'’\-]+)*"
    r"(?:,?\s+et al\.?)?"
)
# Narrative: author group directly followed by "(1999" — replace the group.
_NARRATIVE_RE = re.compile(rf"(?P<a>{_AUTHOR_GROUP})(?=\s*\((?:{_YEAR}))")
# Parenthetical: author group followed by ", 1999" — replace the group.
_PAREN_RE = re.compile(rf"(?P<a>{_AUTHOR_GROUP})(?=,\s*(?:{_YEAR}))")

_SELF_PHRASES = (
    "my previous work",
    "my earlier work",
    "my previous study",
    "my earlier study",
    "our previous work",
    "our earlier work",
    "our previous study",
    "our earlier study",
    "my dissertation",
    "my thesis",
    "this author",
    "the present author",
    "as i argued",
    "as i have argued",
    "as we argued",
    "as we have argued",
)


def _surnames(author_names: list[str]) -> list[str]:
    if not author_names or not all(
        isinstance(n, str) and n.strip() for n in author_names
    ):
        raise WordMcpError(
            "author_names must be a non-empty list of names "
            "(e.g. ['Ian Hurd'] or ['Hurd'])"
        )
    out = []
    for name in author_names:
        surname = name.strip().split()[-1].strip(".,;:")
        if surname:
            out.append(surname)
    return out


def _surname_re(surnames: list[str]) -> re.Pattern:
    return re.compile(
        r"\b(?:" + "|".join(re.escape(s) for s in surnames) + r")\b"
    )


def _iter_ordinals(root: etree._Element) -> list[etree._Element]:
    """All w:p in document order — the stable per-part paragraph addressing
    used in the mapping file (nested text-box paragraphs included, so
    ordinals survive as long as the structure does; content drift is caught
    by the verification pass)."""
    return list(root.iter(qn("w:p")))


def _has_paragraph_ancestor(p: etree._Element) -> bool:
    w_p = qn("w:p")
    parent = p.getparent()
    while parent is not None:
        if parent.tag == w_p:
            return True
        parent = parent.getparent()
    return False


def _citation_spans(text: str, name_re: re.Pattern) -> list[tuple[int, int]]:
    """Spans of citation author-groups that include one of the target
    surnames. Only the author-group is replaced; the year, page numbers, and
    parentheses stay."""
    spans: list[tuple[int, int]] = []
    for pattern in (_NARRATIVE_RE, _PAREN_RE):
        for m in pattern.finditer(text):
            if name_re.search(m.group("a")):
                spans.append(m.span("a"))
    spans.sort()
    # Drop overlaps (defensive; the two patterns should not overlap).
    merged: list[tuple[int, int]] = []
    for s in spans:
        if merged and s[0] < merged[-1][1]:
            continue
        merged.append(s)
    return merged


def _reference_section_span(pkg: DocxPackage) -> tuple[int, int | None] | None:
    """(start_paragraph_index, end_paragraph_index) of the reference list on
    the body axis, located by its heading the same way check_citation_parity
    does. None when no References/Bibliography heading exists."""
    style_outline = _style_outline_map(pkg)
    headings = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph":
            continue
        level = _outline_level(el, style_outline)
        text = paragraph_text(el).strip()
        if level is not None and text:
            headings.append((idx, text))
    ref = next((h for h in headings if _REF_HEADINGS.match(h[1])), None)
    if ref is None:
        return None
    nxt = [i for i, _ in headings if i > ref[0]]
    return (ref[0], min(nxt) if nxt else None)


def _apply_paragraph_edits(
    p: etree._Element, edits: list[tuple[int, int, str]]
) -> list[dict]:
    """Apply (start, end, new_text) edits to one paragraph via the runmap
    (right-to-left, same discipline as search_and_replace) and return
    mapping records with FINAL positions — the positions the edited text
    occupies after all edits, which is what deanonymize verifies against."""
    text, segments = _runmap.build_map(p)
    edits = sorted(edits)
    for start, end, new in reversed(edits):
        _runmap.replace_range(p, segments, start, end, new)
    records = []
    delta = 0
    for start, end, new in edits:
        records.append(
            {
                "start": start + delta,
                "original": text[start:end],
                "replaced_with": new,
            }
        )
        delta += len(new) - (end - start)
    return records


def _default_mapping_path(pkg: DocxPackage) -> Path:
    return pkg.path.with_name(pkg.path.stem + ".anonymization.json")


# ------------------------------------------------------------------- public


def anonymize_for_review(
    pkg: DocxPackage,
    author_names: list[str],
    *,
    replacement: str = "Author",
    mapping_path: str | None = None,
) -> dict:
    """Anonymize a manuscript for double-blind peer review, reversibly.

    author_names: the manuscript author(s) — surnames are taken from the
    last word of each name. Self-citations by those surnames are masked with
    `replacement` (years and page numbers kept); reference-list entries led
    by or co-authored by those surnames become
    "Author (Year). [Details removed for peer review.]"; and identifying
    metadata (creator, lastModifiedBy, Company, Manager) is scrubbed.

    Self-identifying PROSE is flagged, not edited: phrases like "my previous
    work" or "this author", an Acknowledgments section, and any remaining
    surname mention outside citation syntax are listed with paragraph
    locations. Rewriting prose changes meaning — that part is yours.

    A reversal mapping (every change, its location, and the original text)
    is written as JSON to mapping_path (default: <name>.anonymization.json
    beside the document; an existing file is never overwritten). KEEP THE
    MAPPING PRIVATE — it re-identifies the author. Restore later with
    deanonymize; the mapping is written before the document is saved, so if
    the save fails, delete the stale mapping."""
    surnames = _surnames(author_names)
    name_re = _surname_re(surnames)
    if not isinstance(replacement, str) or not replacement:
        raise WordMcpError("replacement must be a non-empty string")

    map_path = (
        Path(mapping_path) if mapping_path else _default_mapping_path(pkg)
    )
    check_path(map_path, "write anonymization mapping")
    if map_path.exists():
        raise WordMcpError(
            f"mapping file already exists: {map_path} — refusing to "
            "overwrite it (it may hold the reversal record of a previous "
            "anonymization). Move or delete it first."
        )

    ref_span = _reference_section_span(pkg)
    _keepalive = body_items(pkg)
    body_idx = {
        id(el): idx for kind, idx, el in _keepalive if kind == "paragraph"
    }

    changes: list[dict] = []
    counts = {"self_citations": 0, "reference_entries": 0}

    parts = ["word/document.xml"] + [
        p
        for p in ("word/footnotes.xml", "word/endnotes.xml")
        if pkg.has_part(p)
    ]
    for part in parts:
        root = pkg.root(part)
        ordinals = _iter_ordinals(root)
        ordinal_of = {id(p): i for i, p in enumerate(ordinals)}
        dirty = False
        for p in ordinals:
            if _has_paragraph_ancestor(p):
                continue  # text-box content is edited through its host
            text, _ = _runmap.build_map(p)
            if not text.strip():
                continue
            idx = body_idx.get(id(p)) if part == "word/document.xml" else None
            in_refs = (
                part == "word/document.xml"
                and ref_span is not None
                and idx is not None
                and idx > ref_span[0]
                and (ref_span[1] is None or idx < ref_span[1])
            )
            edits: list[tuple[int, int, str]] = []
            kind = None
            if in_refs:
                year_m = _REF_YEAR.search(text)
                author_segment = text[: year_m.start()] if year_m else text
                if name_re.search(author_segment):
                    year = year_m.group(1) if year_m else "n.d."
                    masked = (
                        f"{replacement} ({year}). "
                        "[Details removed for peer review.]"
                    )
                    edits = [(0, len(text), masked)]
                    kind = "reference_entry"
            else:
                spans = _citation_spans(text, name_re)
                if spans:
                    edits = [(s, e, replacement) for s, e in spans]
                    kind = "self_citation"
            if not edits:
                continue
            records = _apply_paragraph_edits(p, edits)
            for r in records:
                r.update(
                    {
                        "part": part,
                        "paragraph": ordinal_of[id(p)],
                        "kind": kind,
                    }
                )
                changes.append(r)
            counts[
                "reference_entries"
                if kind == "reference_entry"
                else "self_citations"
            ] += len(records)
            dirty = True
        if dirty:
            pkg.mark_dirty(part)

    # ---- metadata scrub (same fields the submission-prep pass empties),
    # with originals recorded for reversal.
    metadata: dict[str, str] = {}
    if pkg.has_part("docProps/core.xml"):
        core = pkg.root("docProps/core.xml")
        for ns, local, key in (
            (_DC, "creator", "creator"),
            (_CP, "lastModifiedBy", "lastModifiedBy"),
        ):
            el = core.find(f"{{{ns}}}{local}")
            original = el.text if el is not None else None
            if _empty_text(core, ns, local):
                metadata[key] = original or ""
                pkg.mark_dirty("docProps/core.xml")
    if pkg.has_part("docProps/app.xml"):
        app = pkg.root("docProps/app.xml")
        for local in ("Company", "Manager"):
            el = app.find(f"{{{_EP}}}{local}")
            original = el.text if el is not None else None
            if _empty_text(app, _EP, local):
                metadata[local] = original or ""
                pkg.mark_dirty("docProps/app.xml")

    # ---- flag-only pass: self-identifying prose this tool will NOT edit.
    flags: list[dict] = []
    style_outline = _style_outline_map(pkg)
    for kind_, idx, el in _keepalive:
        if kind_ != "paragraph":
            continue
        text = paragraph_text(el)
        low = text.lower()
        for phrase in _SELF_PHRASES:
            pos = low.find(phrase)
            if pos >= 0:
                flags.append(
                    {
                        "kind": "self_identifying_phrase",
                        "paragraph_index": idx,
                        "phrase": phrase,
                        "snippet": text[max(0, pos - 40) : pos + len(phrase) + 40],
                    }
                )
        level = _outline_level(el, style_outline)
        if level is not None and "acknowledg" in low:
            flags.append(
                {
                    "kind": "acknowledgments_section",
                    "paragraph_index": idx,
                    "heading": text.strip(),
                }
            )
        # Surname still visible after masking = a mention outside citation
        # syntax ("As Hurd argues", a book title, a running head...).
        current, _ = _runmap.build_map(el)
        m = name_re.search(current)
        if m:
            flags.append(
                {
                    "kind": "surname_outside_citation",
                    "paragraph_index": idx,
                    "surname": m.group(0),
                    "snippet": current[
                        max(0, m.start() - 40) : m.end() + 40
                    ],
                }
            )

    mapping = {
        "format": MAPPING_FORMAT,
        "document": str(pkg.path),
        "created": _dt.datetime.now().isoformat(timespec="seconds"),
        "author_names": author_names,
        "replacement": replacement,
        "changes": changes,
        "metadata": metadata,
    }
    map_path.write_text(
        json.dumps(mapping, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    return {
        "changed": {
            "self_citations": counts["self_citations"],
            "reference_entries": counts["reference_entries"],
            "metadata_fields": sorted(metadata),
        },
        "flagged_not_changed": flags,
        "mapping_path": str(map_path),
        "reference_section_found": ref_span is not None,
        "warnings": [
            "KEEP THE MAPPING FILE PRIVATE — it re-identifies the author; "
            "never upload it with the submission",
            "flagged prose was NOT edited: rewriting self-identifying "
            "sentences changes meaning and is the author's job",
        ]
        + (
            []
            if ref_span is not None
            else [
                "no References/Bibliography heading found — reference-list "
                "entries were not masked"
            ]
        ),
    }


def deanonymize(pkg: DocxPackage, mapping_path: str | None = None) -> dict:
    """Restore a manuscript anonymized by anonymize_for_review.

    Reads the mapping file (default: <name>.anonymization.json beside the
    document), VERIFIES that every recorded change is still where the
    mapping says it is, and only then restores the original text and
    metadata. If the document drifted since anonymization (edits moved or
    changed the masked text), NOTHING is restored and the refusal lists
    every mismatch — restore those spots by hand or from the mapping's
    recorded originals.

    The mapping file is left on disk (delete it yourself once the restore
    is confirmed)."""
    map_path = (
        Path(mapping_path) if mapping_path else _default_mapping_path(pkg)
    )
    check_path(map_path, "read anonymization mapping")
    if not map_path.exists():
        raise WordMcpError(f"no mapping file at {map_path}")
    try:
        mapping = json.loads(map_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise WordMcpError(f"cannot read mapping file {map_path}: {exc}") from exc
    if mapping.get("format") != MAPPING_FORMAT:
        raise WordMcpError(
            f"{map_path} is not a word-mcp anonymization mapping "
            f"(format={mapping.get('format')!r})"
        )

    changes = mapping.get("changes", [])
    by_part: dict[str, list[dict]] = {}
    for c in changes:
        by_part.setdefault(c["part"], []).append(c)

    # ---- verify EVERYTHING before touching anything (atomicity).
    drift: list[dict] = []
    resolved: dict[str, list[tuple[etree._Element, dict]]] = {}
    for part, part_changes in by_part.items():
        if not pkg.has_part(part):
            drift.append({"part": part, "problem": "part no longer exists"})
            continue
        ordinals = _iter_ordinals(pkg.root(part))
        resolved[part] = []
        for c in part_changes:
            n = c["paragraph"]
            if n >= len(ordinals):
                drift.append(
                    {
                        "part": part,
                        "paragraph": n,
                        "problem": "paragraph no longer exists",
                    }
                )
                continue
            p = ordinals[n]
            text, _ = _runmap.build_map(p)
            start = c["start"]
            expected = c["replaced_with"]
            if text[start : start + len(expected)] != expected:
                drift.append(
                    {
                        "part": part,
                        "paragraph": n,
                        "start": start,
                        "expected": expected,
                        "found": text[start : start + len(expected)],
                        "problem": "text at recorded position has changed",
                    }
                )
                continue
            resolved[part].append((p, c))

    meta_restore: list[tuple[str, str, str, str]] = []  # part, ns, local, val
    metadata = mapping.get("metadata", {})
    _META_FIELDS = {
        "creator": ("docProps/core.xml", _DC, "creator"),
        "lastModifiedBy": ("docProps/core.xml", _CP, "lastModifiedBy"),
        "Company": ("docProps/app.xml", _EP, "Company"),
        "Manager": ("docProps/app.xml", _EP, "Manager"),
    }
    for key, original in metadata.items():
        if key not in _META_FIELDS:
            drift.append(
                {"metadata": key, "problem": "unknown metadata key in mapping"}
            )
            continue
        part, ns, local = _META_FIELDS[key]
        if not pkg.has_part(part):
            drift.append({"metadata": key, "problem": f"{part} missing"})
            continue
        el = pkg.root(part).find(f"{{{ns}}}{local}")
        current = el.text if el is not None else None
        if current not in (None, "", original):
            drift.append(
                {
                    "metadata": key,
                    "expected_empty_or": original,
                    "found": current,
                    "problem": "metadata field was changed since anonymization",
                }
            )
            continue
        if el is not None and current != original:
            meta_restore.append((part, ns, local, original))

    if drift:
        raise WordMcpError(
            "document has drifted since anonymization; NOTHING was restored. "
            "Mismatches: "
            + json.dumps(drift, ensure_ascii=False)[:2000]
            + " — fix these spots manually using the mapping's recorded "
            "originals, or restore from backup."
        )

    # ---- apply, right-to-left within each paragraph so recorded positions
    # of earlier changes stay valid.
    restored = 0
    for part, pairs in resolved.items():
        if not pairs:
            continue
        by_para: dict[int, list[tuple[etree._Element, dict]]] = {}
        for p, c in pairs:
            by_para.setdefault(id(p), []).append((p, c))
        for group in by_para.values():
            group.sort(key=lambda pc: pc[1]["start"], reverse=True)
            for p, c in group:
                text, segments = _runmap.build_map(p)
                _runmap.replace_range(
                    p,
                    segments,
                    c["start"],
                    c["start"] + len(c["replaced_with"]),
                    c["original"],
                )
                restored += 1
        pkg.mark_dirty(part)

    meta_count = 0
    for part, ns, local, value in meta_restore:
        el = pkg.root(part).find(f"{{{ns}}}{local}")
        if el is not None:
            el.text = value
            pkg.mark_dirty(part)
            meta_count += 1

    return {
        "restored_text_changes": restored,
        "restored_metadata_fields": meta_count,
        "mapping_path": str(map_path),
        "note": (
            "mapping file left on disk — delete it once the restore is "
            "confirmed; it re-identifies the author"
        ),
    }
