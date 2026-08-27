"""Live implementations of the high-value tools, mirroring the file-based
parameter shapes and result schemas (plus the standard live fields).

Addressing conventions match the file-based layer:
- paragraph indices are 0-based over BODY-LEVEL paragraphs (paragraphs inside
  tables are excluded from the index, same as ops/read.py's body_items);
- table_index is 0-based over body-level tables; row/cell are 0-based.

Replacement strategy: literal finds go through Word's own Find.Execute
one-at-a-time forward-only (no wrap — immune to offset drift from fields and
to self-referencing replacement loops); regex finds run our ReDoS-guarded
regex over story text with right-to-left application. Regex offsets can drift
in stories containing complex fields, so regex replacements re-verify the
matched text before touching it and refuse on mismatch.
"""

from __future__ import annotations

import contextlib
import re

from ..core.errors import (
    AmbiguousTarget,
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from ..ops import _regex as _rx
from .live import check_text_safe, insert_text_chunked, run_live

# story types
_MAIN_STORY = 1
_FOOTNOTES_STORY = 2
_ENDNOTES_STORY = 3
_HEADER_FOOTER_STORIES = {6, 7, 8, 9, 10, 11}

_WD_REPLACE_ONE = 1
_WD_FIND_STOP = 0
_WD_WITH_IN_TABLE = 12  # wdWithInTable
_WD_UNDERLINE_SINGLE = 1
_WD_UNDERLINE_NONE = 0

_SCOPE_STORIES = {
    "body": {_MAIN_STORY},
    "footnotes": {_FOOTNOTES_STORY},
    "headers": _HEADER_FOOTER_STORIES,
    "all": {_MAIN_STORY, _FOOTNOTES_STORY, _ENDNOTES_STORY}
    | _HEADER_FOOTER_STORIES,
}

_CHAR_KEYS = {
    "bold",
    "italic",
    "underline",
    "strike",
    "font",
    "size_pt",
    "color",
    "highlight",
    "superscript",
    "subscript",
    "small_caps",
    "all_caps",
    "hidden",
    "double_strike",
    "char_spacing_pt",
    "kerning_pt",
    "position_pt",
}
# accepted by the file layer but not implementable through Font live
_CHAR_KEYS_FILE_ONLY = {"language", "east_asian"}

# OOXML highlight name -> wdColorIndex
_HIGHLIGHT_INDEX = {
    "black": 1,
    "blue": 2,
    "cyan": 3,
    "green": 11,
    "magenta": 5,
    "red": 6,
    "yellow": 7,
    "white": 8,
    "darkBlue": 9,
    "darkCyan": 10,
    "darkGreen": 11,
    "darkMagenta": 12,
    "darkRed": 13,
    "darkYellow": 14,
    "darkGray": 15,
    "lightGray": 16,
    "none": 0,
}


# ------------------------------------------------------------------ helpers


def _stories(doc, scope: str):
    """All story ranges for a scope, following NextStoryRange chains."""
    wanted = _SCOPE_STORIES.get(scope)
    if wanted is None:
        raise WordMcpError(
            f"unknown scope {scope!r}; allowed: {sorted(_SCOPE_STORIES)}"
        )
    out = []
    for story in doc.StoryRanges:
        s = story
        while s is not None:
            if s.StoryType in wanted:
                out.append(s)
            try:
                s = s.NextStoryRange
            except Exception:
                break
    return out


def _in_content_control(p) -> bool:
    try:
        return p.Range.ParentContentControl is not None
    except Exception:
        return False


_TOC_HEADING_STYLES = {"toc heading", "tocheading"}


def _sdt_regions(doc) -> list:
    """(start, end) spans of gallery SDT blocks (TOC / List of Figures).

    Word hides gallery SDTs completely from COM — they are absent from
    doc.ContentControls and ParentContentControl returns None inside them
    (verified empirically 2026-08-28) — so the block is reconstructed from
    the TOC-family field range, extended over an adjacent preceding
    TOC-Heading title paragraph (the gallery's standard shape). A custom
    SDT that is neither a content control nor a TOC gallery stays
    invisible; that residual case is documented in the tool docstrings."""
    regions = []
    for collection in ("TablesOfContents", "TablesOfFigures"):
        with contextlib.suppress(Exception):
            coll = getattr(doc, collection)
            for i in range(1, coll.Count + 1):
                rng = coll(i).Range
                start = rng.Start
                with contextlib.suppress(Exception):
                    first = doc.Range(start, start).Paragraphs(1)
                    prev = first.Previous()
                    if (
                        prev is not None
                        and prev.Range.End == first.Range.Start
                        and prev.Style.NameLocal.strip().lower()
                        in _TOC_HEADING_STYLES
                    ):
                        start = prev.Range.Start
                regions.append((start, rng.End))
    return regions


def _body_paragraphs(doc) -> list:
    """Body-level paragraphs, 0-based, matching the file layer's indexing:
    paragraphs inside tables AND inside content-control/gallery SDT blocks
    (TOCs, controls) are excluded, exactly like ops/read.py's body_items.
    Getting this wrong shifts every index and turns mixed-layer workflows
    into wrong-target edits."""
    regions = _sdt_regions(doc)
    out = []
    for p in doc.Paragraphs:
        if p.Range.Information(_WD_WITH_IN_TABLE):
            continue
        if _in_content_control(p):
            continue
        p_start = p.Range.Start
        if any(s <= p_start < e for s, e in regions):
            continue
        out.append(p)
    return out


def _deleted_spans(doc) -> list:
    """(start, end) ranges of tracked DELETIONS in the main story.

    Range.Text includes deleted text ONLY when the user's view shows all
    markup inline (wdRevisionsMarkupAll = 2); in Simple/No-Markup views the
    deleted text is absent and subtracting these spans would eat VISIBLE
    text (verified empirically 2026-08-28). The view state is read, never
    changed. When the view is unreadable, no subtraction happens — wrong
    extra text is safer than silently dropped text."""
    spans = []
    try:
        if doc.ActiveWindow.View.RevisionsFilter.Markup != 2:
            return spans
    except Exception:
        return spans
    with contextlib.suppress(Exception):
        for rev in doc.Revisions:
            if rev.Type == 2:  # wdRevisionDelete
                r = rev.Range
                if r.StoryType == _MAIN_STORY:
                    spans.append((r.Start, r.End))
    return spans


def _para_text(p, deleted_spans: list | None = None) -> str:
    """Paragraph text matching file-layer semantics: no trailing mark, no
    cell marker, no section-break character, tracked-deletion text removed."""
    rng = p.Range
    text = rng.Text.rstrip("\r\x07")
    if deleted_spans:
        base = rng.Start
        keep = []
        for i, ch in enumerate(text):
            pos = base + i
            if any(s <= pos < e for s, e in deleted_spans):
                continue
            keep.append(ch)
        text = "".join(keep)
    return text.replace("\x0c", "")


def _hex_to_wdcolor(color: str) -> int:
    h = color.lstrip("#")
    if len(h) != 6:
        raise WordMcpError(f"color must be 6-digit hex, got {color!r}")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


def _track_guard(session, track: bool, author: str):
    doc, app, guard = session.doc, session.app, session.guard
    if track:
        # per-author revision counts BEFORE the edit: Revisions is ordered
        # by document position, so "the last revision" is NOT the newest —
        # the effective author is found later by diffing these counts
        counts = {}
        with contextlib.suppress(Exception):
            for rev in doc.Revisions:
                counts[rev.Author] = counts.get(rev.Author, 0) + 1
        session._rev_authors_before = counts
        guard.set(doc, "TrackRevisions", True)
        guard.set(app, "UserName", author)
        guard.set(app, "UserInitials", "".join(
            w[0] for w in author.split()[:3] if w
        ) or "A")
    else:
        # forced off: visible tracked deletions re-match Find patterns
        guard.set(doc, "TrackRevisions", False)


def _tracked_result(session, track: bool, author: str, payload: dict) -> dict:
    """Live tracked-change attribution is BEST-EFFORT: Word 365 signed into
    an Office account overrides Application.UserName with the account
    identity. Report the effective author instead of silently lying."""
    if not track:
        return payload
    payload["tracked"] = True
    payload["author_requested"] = author
    try:
        doc = session.doc
        before = getattr(session, "_rev_authors_before", {})
        after = {}
        for rev in doc.Revisions:
            after[rev.Author] = after.get(rev.Author, 0) + 1
        grown = [a for a, n in after.items() if n > before.get(a, 0)]
        if len(grown) == 1:
            effective = grown[0]
            payload["author_effective"] = effective
            if effective != author:
                payload["author_note"] = (
                    "Word is signed into an Office account and attributed "
                    "the revisions to it; the requested author name applies "
                    "only when Word's 'always use these values' privacy "
                    "option is on"
                )
        elif grown:
            payload["author_effective"] = grown
    except Exception:
        pass
    return payload


# ------------------------------------------------------------ search/replace


def _count_matches(story, find: str) -> int:
    """Case-sensitive literal count on the story's display text. NEVER
    count via Find.Execute loops: Word's Find repositions unpredictably
    around field results and the loop spins (observed with citation
    fields, 2026-08-28)."""
    return story.Text.count(find)


def _containing_field_end(story, start: int, end: int):
    """If [start, end) lies inside a field's code-or-result span, return the
    position just past that field; else None. Fields are re-collected per
    call because edits shift their positions."""
    for fld in story.Fields:
        try:
            f_start = fld.Code.Start - 1      # include the begin marker
            f_end = fld.Result.End + 1        # include the end marker
        except Exception:
            continue
        if start >= f_start and end <= f_end:
            return f_end
    return None


def _protected_span_end(story, rng):
    """If the found range lies inside machinery Word regenerates — a field's
    code-or-result span, or a content control (citation/bibliography SDTs
    wrap their fields so story.Fields does NOT list them) — return the
    position just past it; else None. Replacing inside either loops forever
    (Word re-renders, the text re-matches) and achieves nothing anyway."""
    past = _containing_field_end(story, rng.Start, rng.End)
    if past is not None:
        return past
    try:
        cc = rng.ParentContentControl
        if cc is not None:
            return cc.Range.End + 1
    except Exception:
        pass
    return None


def _replace_literal(story, find: str, replace: str) -> tuple:
    """Find-only then manual text assignment, one match at a time, forward
    only. Matches inside fields/content controls are SKIPPED (see
    _protected_span_end). Self-referencing replacements can't loop because
    the search resumes after each replacement, and the iteration budget is
    tied to the PRE-EDIT match count so no regeneration pathology can spin.
    Returns (replaced, skipped_in_fields, skipped_in_deletions)."""
    budget = _count_matches(story, find) + 50
    rng = story.Duplicate
    done = 0
    skipped = 0
    skipped_deleted = 0
    for _ in range(budget):
        f = rng.Find
        f.ClearFormatting()
        f.Text = find
        f.Forward = True
        f.Wrap = _WD_FIND_STOP
        f.MatchWildcards = False
        f.MatchCase = True
        if not f.Execute():          # find only — rng now covers the match
            break
        past_protected = _protected_span_end(story, rng)
        deleted_rev_end = None
        if past_protected is None:
            # a match inside a tracked DELETION re-matches after every
            # tracked replacement (the deleted copy stays findable) — skip
            # it; replacing deleted text is meaningless anyway
            with contextlib.suppress(Exception):
                for rev in rng.Revisions:
                    if rev.Type == 2:  # wdRevisionDelete
                        deleted_rev_end = max(rev.Range.End, rng.End)
                        break
        if past_protected is not None:
            skipped += 1
            resume = past_protected
        elif deleted_rev_end is not None:
            skipped_deleted += 1
            resume = deleted_rev_end
        else:
            check_text_safe(replace)
            rng.Text = replace       # rng covers the replacement afterwards
            done += 1
            resume = rng.End
        rng.SetRange(resume, max(story.End, resume))
        if rng.Start >= rng.End:
            break
    return done, skipped, skipped_deleted


def _replace_regex(story, find: str, replace: str) -> tuple:
    """Python-regex over story text, applied right-to-left. Each span is
    verified to still hold the matched text before replacing; a mismatch
    means field machinery shifted COM offsets past that point, so that
    match is SKIPPED and reported rather than corrupting or refusing the
    whole story (matches before the first field always line up).
    Returns (replaced, skipped_offset_drift)."""
    if _rx.finditer(find, ""):
        # z*-style patterns match the empty string at every position and
        # would shred the document (same guard as the file layer)
        raise WordMcpError(
            f"regex {find!r} can match the empty string, which would insert "
            "the replacement between every pair of characters; anchor the "
            "pattern so every match is non-empty"
        )
    text = story.Text
    matches = _rx.finditer(find, text)
    done = 0
    skipped = 0
    for m in reversed(matches):
        sub = story.Duplicate
        sub.SetRange(m.start(), m.end())
        if sub.Text != m.group(0):
            skipped += 1
            continue
        replacement = m.expand(replace)
        check_text_safe(replacement)
        sub.Text = replacement
        done += 1
    return done, skipped


def search_and_replace(
    path: str,
    replacements: list[dict],
    scope: str = "body",
    max_replacements: int | None = None,
    track: bool = False,
    author: str = "Claude",
) -> dict:
    def body(session):
        doc = session.doc
        stories = _stories(doc, scope)
        _track_guard(session, track, author)

        # pre-count for the blast-radius guard, before ANY change
        if max_replacements is not None:
            total = 0
            for item in replacements:
                find = item["find"]
                if item.get("regex"):
                    for story in stories:
                        total += len(_rx.finditer(find, story.Text))
                else:
                    for story in stories:
                        total += _count_matches(story, find)
            if total > max_replacements:
                raise WordMcpError(
                    f"{total} matches exceed max_replacements="
                    f"{max_replacements}; nothing was changed"
                )

        per_item = []
        total = 0
        for item in replacements:
            find, replace = item["find"], item.get("replace", "")
            check_text_safe(replace)
            n = in_fields = drift = in_deletions = 0
            if item.get("regex"):
                for story in stories:
                    done, skipped = _replace_regex(story, find, replace)
                    n += done
                    drift += skipped
            else:
                for story in stories:
                    done, skipped, skipped_del = _replace_literal(
                        story, find, replace
                    )
                    n += done
                    in_fields += skipped
                    in_deletions += skipped_del
            entry = {"find": find, "replacements": n}
            if in_fields:
                entry["skipped_inside_fields"] = in_fields
                entry["note"] = (
                    "matches inside field results are skipped — Word "
                    "regenerates field results, so edits there do not stick"
                )
            if in_deletions:
                entry["skipped_inside_tracked_deletions"] = in_deletions
            if drift:
                entry["skipped_offset_drift"] = drift
                entry["note"] = (
                    "matches positioned after complex fields were skipped "
                    "(COM offsets drift there); use a literal find for "
                    "those, or edit the closed file"
                )
            per_item.append(entry)
            total += n
        return _tracked_result(
            session, track, author,
            {"total_replacements": total, "items": per_item, "scope": scope},
        )

    return run_live(path, "search and replace", body)


# ------------------------------------------------------------- paragraphs


def get_text(
    path: str,
    start: int = 0,
    end: int | None = None,
    contains: str | None = None,
) -> dict:
    def body(session):
        doc = session.doc
        deleted = _deleted_spans(doc)
        paras = _body_paragraphs(doc)
        out = []
        # end is EXCLUSIVE, matching the file layer's slice semantics
        stop = len(paras) if end is None else min(end, len(paras))
        for i in range(max(0, start), stop):
            p = paras[i]
            text = _para_text(p, deleted)
            if contains is not None and contains not in text:
                continue
            entry = {"index": i, "text": text}
            try:
                entry["style"] = p.Style.NameLocal
            except Exception:
                entry["style"] = None
            try:
                lvl = p.OutlineLevel
                if 1 <= lvl <= 9:
                    entry["heading_level"] = lvl
            except Exception:
                pass
            out.append(entry)
        # SDT (content-control/gallery) paragraphs match the file layer:
        # listed with index None so they are readable but never shift body
        # indices
        if contains is None and start == 0 and end is None:
            regions = _sdt_regions(doc)
            for p in doc.Paragraphs:
                if p.Range.Information(_WD_WITH_IN_TABLE):
                    continue
                in_sdt = _in_content_control(p) or any(
                    s <= p.Range.Start < e for s, e in regions
                )
                if in_sdt:
                    text = _para_text(p, deleted)
                    if text:
                        out.append(
                            {"index": None, "in_sdt": True, "text": text}
                        )
        return {"paragraphs": out, "total_body_paragraphs": len(paras)}

    return run_live(path, "read text", body, mutating=False)


def get_outline(path: str) -> dict:
    def body(session):
        paras = _body_paragraphs(session.doc)
        out = []
        for i, p in enumerate(paras):
            try:
                lvl = p.OutlineLevel
            except Exception:
                continue
            text = _para_text(p).strip()
            if 1 <= lvl <= 9 and text:
                out.append(
                    {"paragraph_index": i, "level": lvl, "text": text}
                )
        return {"outline": out}

    return run_live(path, "read outline", body, mutating=False)


def get_document_info(path: str) -> dict:
    def body(session):
        doc = session.doc
        info = {
            "paragraphs": len(_body_paragraphs(doc)),
            "tables": doc.Tables.Count,
            "footnotes": doc.Footnotes.Count,
            "endnotes": doc.Endnotes.Count,
            "comments": doc.Comments.Count,
            "revisions": doc.Revisions.Count,
            "sections": doc.Sections.Count,
            "words": doc.Words.Count,
            "track_revisions": bool(doc.TrackRevisions),
        }
        return info

    return run_live(path, "document info", body, mutating=False)


def find_text(
    path: str, query: str, regex: bool = False, context_chars: int = 60
) -> dict:
    if not query:
        raise WordMcpError("query must be non-empty")
    if regex and _rx.finditer(query, ""):
        raise WordMcpError(
            f"regex {query!r} can match the empty string and would match at "
            "every position; anchor the pattern"
        )

    def body(session):
        doc = session.doc
        deleted = _deleted_spans(doc)
        regions = _sdt_regions(doc)
        matches = []
        body_idx = -1
        for p in doc.Paragraphs:
            in_table = bool(p.Range.Information(_WD_WITH_IN_TABLE))
            in_sdt = _in_content_control(p) or any(
                s <= p.Range.Start < e for s, e in regions
            )
            if not in_table and not in_sdt:
                body_idx += 1
            text = _para_text(p, deleted)
            found = (
                _rx.finditer(query, text)
                if regex
                else list(re.finditer(re.escape(query), text))
            )
            for m in found:
                lo = max(0, m.start() - context_chars)
                hi = min(len(text), m.end() + context_chars)
                entry = {
                    "match": m.group(0),
                    "context": text[lo:hi],
                    "location": (
                        "table cell" if in_table
                        else "content control" if in_sdt
                        else "body"
                    ),
                }
                if not in_table and not in_sdt:
                    entry["paragraph"] = body_idx
                matches.append(entry)
                if len(matches) >= 500:
                    return {"matches": matches, "truncated": True}
        return {"matches": matches, "truncated": False}

    return run_live(path, "find text", body, mutating=False)


def insert_paragraphs(
    path: str,
    paragraphs: list[dict],
    after_index: int | None = None,
    before_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    track: bool = False,
    author: str = "Claude",
) -> dict:
    targets = [after_index is not None, before_index is not None,
               after_anchor is not None, at_end]
    if sum(targets) != 1:
        raise WordMcpError(
            "give exactly one of after_index, before_index, after_anchor, at_end"
        )

    def body(session):
        doc = session.doc
        _track_guard(session, track, author)
        paras = _body_paragraphs(doc)

        # Insertion point + mode. "prefix" inserts "\r"+text just BEFORE the
        # anchor's paragraph mark: collapsing to the anchor's END would land
        # INSIDE a following table's first cell when the anchor abuts a
        # table, so we never cross the paragraph mark at all.
        if at_end:
            point = doc.Content.End - 1
            prefix = True
        elif after_anchor is not None:
            hits = [
                i for i, p in enumerate(paras) if after_anchor in _para_text(p)
            ]
            if not hits:
                raise TargetNotFound(f"anchor text {after_anchor!r} not found")
            if len(hits) > 1:
                raise AmbiguousTarget(
                    f"anchor text {after_anchor!r} matches paragraphs {hits}; "
                    "use after_index instead"
                )
            point = paras[hits[0]].Range.End - 1
            prefix = True
        elif after_index is not None:
            if not 0 <= after_index < len(paras):
                raise TargetNotFound(
                    f"after_index {after_index} out of range "
                    f"({len(paras)} body paragraphs)"
                )
            point = paras[after_index].Range.End - 1
            prefix = True
        else:
            if not 0 <= before_index < len(paras):
                raise TargetNotFound(
                    f"before_index {before_index} out of range "
                    f"({len(paras)} body paragraphs)"
                )
            point = paras[before_index].Range.Start
            prefix = False

        insertion = doc.Range(point, point)
        inserted = 0
        for spec in paragraphs:
            text = spec.get("text", "")
            check_text_safe(text)
            payload = ("\r" + text) if prefix else (text + "\r")
            rng = insertion.Duplicate
            insert_text_chunked(rng, payload)
            # rng covers the payload; the new paragraph's own text span:
            if prefix:
                p_start, p_end = rng.Start + 1, rng.End
            else:
                p_start, p_end = rng.Start, max(rng.Start, rng.End - 1)
            styled = doc.Range(p_start, max(p_start, p_end))
            if spec.get("style"):
                try:
                    styled.Style = spec["style"]
                except Exception as exc:
                    raise TargetNotFound(
                        f"style {spec['style']!r} does not exist in this "
                        "document"
                    ) from exc
            else:
                # match the file layer: unstyled inserts are Normal, not the
                # anchor's style that the split mark inherited
                styled.Style = doc.Styles(-1)  # wdStyleNormal
            insertion.SetRange(rng.End, rng.End)
            inserted += 1
        return _tracked_result(session, track, author, {"inserted": inserted})

    return run_live(path, "insert paragraphs", body)


def delete_paragraphs(
    path: str,
    start: int,
    end: int | None = None,
    track: bool = False,
    author: str = "Claude",
) -> dict:
    def body(session):
        doc = session.doc
        last = start if end is None else end
        paras = _body_paragraphs(doc)
        if not (0 <= start <= last < len(paras)):
            raise TargetNotFound(
                f"paragraph range [{start}..{last}] out of range "
                f"({len(paras)} body paragraphs)"
            )
        rng = doc.Range(paras[start].Range.Start, paras[last].Range.End)
        if rng.Tables.Count:
            raise UnsupportedStructure(
                "paragraph range spans a table (body-level paragraph indices "
                "skip tables, so the range crosses one); split the deletion "
                "around the table"
            )
        if rng.Sections.Count > 1 or "\x0c" in rng.Text:
            raise UnsupportedStructure(
                "range carries a section break; delete around it or use "
                "the file-based tool on the closed file"
            )
        for f in rng.Fields:
            code = f.Code
            result = f.Result
            if code.Start < rng.Start or result.End > rng.End:
                raise UnsupportedStructure(
                    "range would cut through a field; adjust the range"
                )
        _track_guard(session, track, author)
        deleted = last - start + 1
        # Content controls (citation/bibliography SDTs) inside the range:
        # Text-assignment silently leaves them (and their field) behind —
        # delete each control WITH its contents first, then the remainder.
        ccs = []
        with contextlib.suppress(Exception):
            ccs = list(rng.ContentControls)
        for cc in ccs:
            if cc.Range.Start < rng.Start or cc.Range.End > rng.End:
                raise UnsupportedStructure(
                    "a content control extends beyond the paragraph range; "
                    "deleting would remove content outside the range"
                )
        for cc in ccs:
            cc.Delete(True)
        # Bare fields survive text-assignment too (Word preserves field
        # structures through Range.Text writes) — delete them explicitly.
        removed_fields = False
        with contextlib.suppress(Exception):
            for fld in list(rng.Fields):
                fld.Delete()
                removed_fields = True
        if ccs or removed_fields:  # boundaries shifted; rebuild the range
            rng = doc.Range(paras[start].Range.Start, paras[last].Range.End)
        # Text-assignment, NOT Range.Delete: Delete refuses to remove a
        # paragraph mark that would land the deletion flush against a
        # following table and leaves a stray empty paragraph; assigning ""
        # removes the same span cleanly (verified empirically, 2026-08-28).
        rng.Text = ""
        return _tracked_result(session, track, author, {"deleted": deleted})

    return run_live(path, "delete paragraphs", body)


# ------------------------------------------------------------------ tables


def set_cells(
    path: str,
    table_index: int,
    edits: list[dict],
    track: bool = False,
    author: str = "Claude",
) -> dict:
    def body(session):
        doc = session.doc
        if not 0 <= table_index < doc.Tables.Count:
            raise TargetNotFound(
                f"no table with index {table_index} "
                f"({doc.Tables.Count} tables)"
            )
        table = doc.Tables(table_index + 1)
        # Pre-flight EVERY target before writing ANY cell: Word forbids
        # Rows(r).Cells on vertically merged tables (raw error mid-batch
        # would leave a partial write). Refusal is atomic.
        targets = []
        for edit in edits:
            r, c, text = edit["row"], edit["cell"], edit["text"]
            check_text_safe(text)
            if not 0 <= r < table.Rows.Count:
                raise TargetNotFound(
                    f"row {r} out of range (table has {table.Rows.Count})"
                )
            try:
                row = table.Rows(r + 1)
                cell_count = row.Cells.Count
            except Exception as exc:
                raise UnsupportedStructure(
                    "this table has vertically merged cells, which Word's "
                    "live row addressing cannot handle — use the file-based "
                    "set_cells on the closed file (it is merge-aware)"
                ) from exc
            if not 0 <= c < cell_count:
                raise TargetNotFound(
                    f"cell {c} out of range (row {r} has {cell_count})"
                )
            targets.append((row.Cells(c + 1), text))
        _track_guard(session, track, author)
        applied = 0
        for cell, text in targets:
            cell_rng = cell.Range
            cell_rng.SetRange(cell_rng.Start, cell_rng.End - 1)  # drop \x07
            cell_rng.Text = text
            applied += 1
        return _tracked_result(
            session, track, author,
            {"cells_set": applied, "table": table_index},
        )

    return run_live(path, "set cells", body)


# -------------------------------------------------------------- formatting


def format_text(
    path: str,
    formatting: dict,
    paragraph_index: int | None = None,
    find: str | None = None,
    occurrence: int = 1,
) -> dict:
    file_only = set(formatting) & _CHAR_KEYS_FILE_ONLY
    if file_only:
        raise UnsupportedStructure(
            f"formatting key(s) {sorted(file_only)} are not supported on "
            "open documents; close the file and use the file-based tool"
        )
    unknown = set(formatting) - _CHAR_KEYS
    if unknown:
        raise WordMcpError(
            f"unknown character-formatting key(s) {sorted(unknown)}; "
            f"allowed: {sorted(_CHAR_KEYS)}"
        )
    if formatting.get("superscript") and formatting.get("subscript"):
        raise WordMcpError("superscript and subscript are mutually exclusive")
    if find is None and paragraph_index is None:
        raise WordMcpError("need paragraph_index, find, or both")

    def body(session):
        doc = session.doc
        if paragraph_index is not None:
            paras = _body_paragraphs(doc)
            if not 0 <= paragraph_index < len(paras):
                raise TargetNotFound(
                    f"paragraph_index {paragraph_index} out of range "
                    f"({len(paras)} body paragraphs)"
                )
            candidates = [(paras[paragraph_index], paragraph_index)]
        else:
            candidates = []
            body_idx = -1
            for p in doc.Paragraphs:
                in_table = bool(p.Range.Information(_WD_WITH_IN_TABLE))
                if not in_table:
                    body_idx += 1
                candidates.append((p, None if in_table else body_idx))

        seen = 0
        for p, idx in candidates:
            text = _para_text(p)
            if find is None:
                spans = [(0, len(text))] if text else []
            else:
                spans = []
                pos = text.find(find)
                while pos != -1:
                    spans.append((pos, pos + len(find)))
                    pos = text.find(find, pos + 1)
            for lo, hi in spans:
                seen += 1
                if seen != occurrence:
                    continue
                base = p.Range.Start
                rng = doc.Range(base + lo, base + hi)
                if find is not None and rng.Text != find:
                    raise UnsupportedStructure(
                        "character offsets do not line up with Word's "
                        "positions in this paragraph (complex fields "
                        "present); use the file-based tool on the closed file"
                    )
                _apply_char_formatting(rng, formatting)
                loc = {"start": lo, "end": hi}
                if idx is not None:
                    loc["paragraph"] = idx
                else:
                    loc["location"] = "table cell"
                return {"formatted": loc}
        raise TargetNotFound(
            f"occurrence {occurrence} of {find!r} not found"
            if find
            else "paragraph has no text to format"
        )

    return run_live(path, "format text", body)


# -------------------------------------------------------- live-only tools


def insert_at_cursor(path: str, text: str, *, newline: bool = False) -> dict:
    """Insert text at the user's cursor. Selection is READ once (its start
    position); it is never written, so the user's selection/cursor state is
    left exactly as Word adjusts it around the insertion."""

    def body(session):
        doc, app = session.doc, session.app
        check_text_safe(text)
        try:
            sel = app.Selection
            sel_doc = sel.Document
            if sel_doc.FullName.lower() != doc.FullName.lower():
                raise WordMcpError(
                    "the cursor is not in this document (active document is "
                    f"{sel_doc.Name}); click into the target document first"
                )
            point = sel.Range.Start
            story = sel.StoryType
        except WordMcpError:
            raise
        except Exception as exc:
            raise WordMcpError(
                "could not read the cursor position"
            ) from exc
        if story != _MAIN_STORY:
            raise UnsupportedStructure(
                "the cursor is in a footnote/header/other pane; live cursor "
                "insert currently supports the main text only"
            )
        payload = text + ("\r" if newline else "")
        rng = doc.Range(point, point)
        insert_text_chunked(rng, payload)
        return {"inserted_at": point, "chars": len(payload)}

    return run_live(path, "insert at cursor", body)


def scroll_to(
    path: str,
    find: str | None = None,
    paragraph_index: int | None = None,
) -> dict:
    """THE one sanctioned scroll: bring a location into view in the user's
    window WITHOUT selecting it or moving the cursor."""
    if (find is None) == (paragraph_index is None):
        raise WordMcpError("give exactly one of find, paragraph_index")

    def body(session):
        doc = session.doc
        if paragraph_index is not None:
            paras = _body_paragraphs(doc)
            if not 0 <= paragraph_index < len(paras):
                raise TargetNotFound(
                    f"paragraph_index {paragraph_index} out of range "
                    f"({len(paras)} body paragraphs)"
                )
            rng = paras[paragraph_index].Range
            label = {"paragraph": paragraph_index}
        else:
            rng = doc.Content.Duplicate
            f = rng.Find
            f.ClearFormatting()
            f.Text = find
            f.Forward = True
            f.Wrap = _WD_FIND_STOP
            if not f.Execute():
                raise TargetNotFound(f"{find!r} not found in the body")
            label = {"found": find}
        window = doc.ActiveWindow
        window.ScrollIntoView(rng, True)
        label["page"] = rng.Information(1)  # wdActiveEndAdjustedPageNumber
        return label

    return run_live(path, "scroll to", body, mutating=False)


def set_track_changes(path: str, enabled: bool) -> dict:
    """Deliberate PERSISTENT toggle of track changes on the open document —
    the one live tool whose purpose IS a state change, so no guard/restore."""

    def body(session):
        doc = session.doc
        previous = bool(doc.TrackRevisions)
        doc.TrackRevisions = enabled
        return {"track_changes": enabled, "was": previous}

    return run_live(path, "toggle track changes", body)


def _apply_char_formatting(rng, fmt: dict):
    font = rng.Font
    if "bold" in fmt:
        font.Bold = bool(fmt["bold"])
    if "italic" in fmt:
        font.Italic = bool(fmt["italic"])
    if "underline" in fmt:
        font.Underline = (
            _WD_UNDERLINE_SINGLE if fmt["underline"] else _WD_UNDERLINE_NONE
        )
    if "strike" in fmt:
        font.StrikeThrough = bool(fmt["strike"])
    if "font" in fmt:
        font.Name = fmt["font"]
    if "size_pt" in fmt:
        font.Size = float(fmt["size_pt"])
    if "color" in fmt:
        font.Color = _hex_to_wdcolor(fmt["color"])
    if "highlight" in fmt:
        name = fmt["highlight"]
        if name not in _HIGHLIGHT_INDEX:
            raise WordMcpError(
                f"unknown highlight {name!r}; "
                f"allowed: {sorted(_HIGHLIGHT_INDEX)}"
            )
        rng.HighlightColorIndex = _HIGHLIGHT_INDEX[name]
    if "superscript" in fmt:
        font.Superscript = bool(fmt["superscript"])
    if "subscript" in fmt:
        font.Subscript = bool(fmt["subscript"])
    if "small_caps" in fmt:
        font.SmallCaps = bool(fmt["small_caps"])
    if "all_caps" in fmt:
        font.AllCaps = bool(fmt["all_caps"])
    if "hidden" in fmt:
        font.Hidden = bool(fmt["hidden"])
    if "double_strike" in fmt:
        font.DoubleStrikeThrough = bool(fmt["double_strike"])
    if "char_spacing_pt" in fmt:
        font.Spacing = float(fmt["char_spacing_pt"])
    if "kerning_pt" in fmt:
        font.Kerning = float(fmt["kerning_pt"])
    if "position_pt" in fmt:
        font.Position = int(fmt["position_pt"])
