"""Live implementations of the high-value tools, mirroring the file-based
parameter shapes and result schemas (plus the standard live fields).

SCHEMA PARITY RULING (L7+L9, 2026-08-28): the canonical result shape for
every dual-mode tool is the FILE-mode shape; dict results additionally carry
"live": true (+ undo/dirty metadata), list results stay flat lists (a list
cannot carry the live key). Live-only information is ADDITIVE keys only —
never a different top-level shape. One parsing path for callers.

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
    StaleAnchor,
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from ..ops import _regex as _rx
from ..ops.localization import canonical_for_name, style_name_matches
from .live import check_text_safe, insert_text_chunked, live_session, run_live

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

# ComputeStatistics codes. NEVER count words via doc.Words.Count — it counts
# punctuation runs and paragraph marks as "words" (~18% high on real prose;
# L1, 2026-08-28). ComputeStatistics(wdStatisticWords) matches Word's own
# status-bar number.
_WD_STAT_WORDS = 0
_WD_STAT_CHARS_WITH_SPACES = 5

# Word's Find.Text COM property rejects strings beyond ~255 characters
# ('String parameter too long'). Longer finds are located via a prefix
# search + range extension + full verification (L5).
_FIND_TEXT_LIMIT = 255
_FIND_PREFIX_CHARS = 250

# Range.Text single-call assignment limit mirror (live.TEXT_CHUNK); longer
# replacements go through delete + chunked insert.
_ASSIGN_CHUNK = 30000

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
    "language",
    "east_asian_language",
}
# BCP-47 language tag -> Word LCID, for Font.LanguageID /
# Font.LanguageIDFarEast (the live twin of the file layer's w:lang keys).
# Values are the documented WdLanguageID constants.
_LANGUAGE_LCIDS = {
    "en-us": 1033, "en-gb": 2057, "en-au": 3081, "en-ca": 4105,
    "ko-kr": 1042,
    "ja-jp": 1041,
    "zh-cn": 2052, "zh-tw": 1028, "zh-hk": 3076,
    "de-de": 1031, "de-at": 3079, "de-ch": 2055,
    "fr-fr": 1036, "fr-ca": 3084, "fr-ch": 4108,
    "es-es": 3082, "es-mx": 2058,
    "it-it": 1040,
    "pt-br": 1046, "pt-pt": 2070,
    "ru-ru": 1049,
    "nl-nl": 1043,
    "pl-pl": 1045,
    "tr-tr": 1055,
    "ar-sa": 1025,
    "he-il": 1037,
    "hi-in": 1081,
    "th-th": 1054,
    "vi-vn": 1066,
    "id-id": 1057,
    "sv-se": 1053, "da-dk": 1030, "nb-no": 1044, "fi-fi": 1035,
    "cs-cz": 1029, "el-gr": 1032, "hu-hu": 1038, "ro-ro": 1048,
    "uk-ua": 1058,
}


# languages Word stores ONLY in the East-Asian slot: writing their LCID to
# Range.LanguageID is silently IGNORED by Word (verified 2026-08-28 — the
# Latin slot keeps its old value), so honesty demands routing them through
# east_asian_language instead of reporting a success that did nothing.
_EAST_ASIAN_PRIMARY = {"ko", "ja", "zh"}


def _lcid(tag: str) -> int:
    lcid = _LANGUAGE_LCIDS.get(str(tag).lower())
    if lcid is None:
        raise WordMcpError(
            f"language tag {tag!r} has no live LCID mapping; supported live: "
            f"{sorted(_LANGUAGE_LCIDS)} — for other tags close the document "
            "and use the file-based tool (it writes the tag verbatim)"
        )
    return lcid

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


def _run_live_list(path: str, tool_name: str, body) -> list:
    """Run a READ-ONLY body whose result is a LIST, returned unchanged.

    Live/file schema parity (L7+L9 ruling, 2026-08-28): the canonical result
    shape is the FILE-mode shape. File-mode list results therefore stay flat
    lists in live mode too — a list cannot carry the "live": true key, so
    list-shaped reads return without live metadata (nothing was mutated, so
    no undo/restore report is lost)."""
    with live_session(path, tool_name, mutating=False) as session:
        return body(session)


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
                    # NameLocal is the LOCALIZED display name ("TOC Heading"
                    # only on English installs — 목차 제목 on Korean, etc.),
                    # so the match routes through the localization aliases.
                    if (
                        prev is not None
                        and prev.Range.End == first.Range.Start
                        and style_name_matches(
                            prev.Style.NameLocal, "toc_heading"
                        )
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


# --------------------------------------------- snapshot staleness guard
#
# Live-branch location resolution reads a DISK SNAPSHOT of the open
# document (server._resolve_for_live), so a text selector (search,
# after_heading, outline, anchor) resolved while Word holds unsaved
# changes can point at the wrong live paragraph. Bodies that accept a
# verify_text compare the live paragraph's text against the snapshot
# paragraph the selector matched, and refuse (STALE_ANCHOR) on mismatch.
# The check is skipped when doc.Saved is true (live == disk, nothing can
# be stale), so the refusal's remedy — save the document in Word — always
# clears it.

_NOTE_REF_RE = re.compile(r"\[(?:fn|en):\d+\]")


def _texts_equivalent(file_text: str, live_text: str) -> bool:
    """File-layer paragraph text vs Range.Text, normalized: note-reference
    markers ([fn:3] vs \\x02), soft breaks (\\n vs \\x0b), non-breaking
    hyphens, and object placeholders/control chars the two extractors
    render differently."""
    f = _NOTE_REF_RE.sub("\x02", file_text)
    lv = (
        live_text.replace("\x0b", "\n")
        .replace("\x1e", "-")
        .replace("\x1f", "")
    )
    keep = "\t\n\x02"
    f = "".join(ch for ch in f if ch >= " " or ch in keep)
    lv = "".join(ch for ch in lv if ch >= " " or ch in keep)
    return f == lv


def _stale_guard(session, paras, index: int, expected: str | None,
                 what: str) -> None:
    """Refuse when a snapshot-resolved index cannot be trusted: the open
    document has unsaved changes AND the live paragraph no longer carries
    the text the selector matched on disk. expected=None disables the
    check (index/cursor addressing, where no snapshot text was involved)."""
    if expected is None:
        return
    with contextlib.suppress(Exception):
        if session.doc.Saved:
            return
    if not 0 <= index < len(paras):
        return  # the op's own bounds check raises the precise error
    if not _texts_equivalent(expected, _para_text(paras[index])):
        raise StaleAnchor(
            f"{what}: the open document has UNSAVED changes and paragraph "
            f"{index} no longer matches the saved file this location was "
            "resolved against (locations resolve against the last saved "
            "state). Save the document in Word (com_save_document) and "
            "retry."
        )


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
    payload["tracked_as"] = author       # file-mode parity key
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


def _assign_text(rng, text: str):
    """Range text assignment respecting COM's single-string size limit.
    Short strings assign directly; long ones clear the range then insert in
    chunks (InsertAfter extends the range to cover everything inserted, so
    rng covers the full replacement afterwards either way)."""
    if len(text) <= _ASSIGN_CHUNK:
        rng.Text = text
        return
    rng.Text = ""                    # collapse the match away first
    insert_text_chunked(rng, text)   # rng grows to cover the inserted text


def _replace_literal(story, find: str, replace: str) -> tuple:
    """Find-only then manual text assignment, one match at a time, forward
    only. Matches inside fields/content controls are SKIPPED (see
    _protected_span_end). Self-referencing replacements can't loop because
    the search resumes after each replacement, and the iteration budget is
    tied to the PRE-EDIT match count so no regeneration pathology can spin.

    Finds longer than Word's ~255-char Find.Text limit (L5) are located via
    the first ~250 chars, the range is extended to the full find length, and
    the extended range's text is VERIFIED against the whole find before
    anything is touched; a prefix hit that is not a full match is skipped.
    The replacement side never goes through Find.Replacement (same 255
    limit) — text assignment has no such limit.
    Returns (replaced, skipped_in_fields, skipped_in_deletions)."""
    long_find = len(find) > _FIND_TEXT_LIMIT
    probe = find[:_FIND_PREFIX_CHARS] if long_find else find
    # budget on PROBE occurrences: >= full-string occurrences, so prefix
    # hits that fail full verification can never starve real matches
    budget = _count_matches(story, probe) + 50
    rng = story.Duplicate
    done = 0
    skipped = 0
    skipped_deleted = 0
    for _ in range(budget):
        f = rng.Find
        f.ClearFormatting()
        f.Text = probe
        f.Forward = True
        f.Wrap = _WD_FIND_STOP
        f.MatchWildcards = False
        f.MatchCase = True
        if not f.Execute():          # find only — rng now covers the match
            break
        if long_find:
            # extend over the full find length and verify the whole string;
            # offsets are per-character within one story, and the explicit
            # Text comparison catches any field-machinery drift
            full_end = rng.Start + len(find)
            if full_end > story.End:
                break                # too close to the story end to match
            candidate = story.Duplicate
            candidate.SetRange(rng.Start, full_end)
            if candidate.Text != find:
                resume = rng.End     # prefix-only hit: not a real match
                rng.SetRange(resume, max(story.End, resume))
                if rng.Start >= rng.End:
                    break
                continue
            rng.SetRange(candidate.Start, candidate.End)
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
            _assign_text(rng, replace)  # rng covers the replacement after
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
        _assign_text(sub, replacement)
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

        # RESULT SHAPE (L7 parity ruling): canonical file-mode shape —
        # {"replaced": {find: n}, "total": n} — plus live-only additive keys
        # (skip counters, notes). The old items[]/total_replacements live
        # shape is retired.
        replaced: dict = {}
        skipped_fields: dict = {}
        skipped_deletions: dict = {}
        skipped_drift: dict = {}
        notes: list = []
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
            replaced[find] = replaced.get(find, 0) + n
            if in_fields:
                skipped_fields[find] = in_fields
                note = (
                    "matches inside field results are skipped — Word "
                    "regenerates field results, so edits there do not stick"
                )
                if note not in notes:
                    notes.append(note)
            if in_deletions:
                skipped_deletions[find] = in_deletions
            if drift:
                skipped_drift[find] = drift
                note = (
                    "matches positioned after complex fields were skipped "
                    "(COM offsets drift there); use a literal find for "
                    "those, or edit the closed file"
                )
                if note not in notes:
                    notes.append(note)
            total += n
        result = {"replaced": replaced, "total": total, "scope": scope}
        if skipped_fields:
            result["skipped_inside_fields"] = skipped_fields
        if skipped_deletions:
            result["skipped_inside_tracked_deletions"] = skipped_deletions
        if skipped_drift:
            result["skipped_offset_drift"] = skipped_drift
        if notes:
            result["notes"] = notes
        return _tracked_result(session, track, author, result)

    return run_live(path, "search and replace", body)


# ------------------------------------------------------------- paragraphs


def get_text(
    path: str,
    start: int = 0,
    end: int | None = None,
    contains: str | None = None,
) -> list:
    """FILE-MODE SHAPE (L9 parity): a flat list of paragraph entries
    {index, text, style, heading_level?}, SDT paragraphs appended with
    index None + in_sdt true. Live styles are Word's LOCALIZED display
    names (Style.NameLocal); file mode reports style IDs."""

    def body(session):
        doc = session.doc
        deleted = _deleted_spans(doc)
        paras = _body_paragraphs(doc)
        out = []

        def _entry(p, index):
            text = _para_text(p, deleted)
            entry = {"index": index, "text": text}
            try:
                entry["style"] = p.Style.NameLocal
            except Exception:
                entry["style"] = None
            return entry

        # end is EXCLUSIVE, matching the file layer's slice semantics
        stop = len(paras) if end is None else min(end, len(paras))
        for i in range(max(0, start), stop):
            entry = _entry(paras[i], i)
            if contains is not None and contains not in entry["text"]:
                continue
            try:
                lvl = paras[i].OutlineLevel
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
                    entry = _entry(p, None)
                    if entry["text"]:
                        entry["in_sdt"] = True
                        out.append(entry)
        return out

    return _run_live_list(path, "read text", body)


def get_outline(path: str) -> list:
    """FILE-MODE SHAPE (L9 parity): flat list of
    {paragraph_index, level, text, detected_via}. Word's OutlineLevel is the
    EFFECTIVE value, so outlineLvl-based template headings (L8) are seen
    live for free; detected_via distinguishes built-in Heading styles from
    outline-level overrides via the paragraph's style name."""

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
                via = "outline_level"
                with contextlib.suppress(Exception):
                    canonical = canonical_for_name(p.Style.NameLocal)
                    if canonical and canonical.startswith("heading"):
                        via = "heading_style"
                out.append(
                    {
                        "paragraph_index": i,
                        "level": lvl,
                        "text": text,
                        "detected_via": via,
                    }
                )
        return out

    return _run_live_list(path, "read outline", body)


def get_document_info(path: str) -> dict:
    def body(session):
        doc = session.doc
        info = {
            "path": doc.FullName,
            "paragraphs": len(_body_paragraphs(doc)),
            "tables": doc.Tables.Count,
            "sections": doc.Sections.Count,
            "footnotes": doc.Footnotes.Count,
            "endnotes": doc.Endnotes.Count,
            "comments": doc.Comments.Count,
            "revisions": doc.Revisions.Count,
            # L1 fix: ComputeStatistics matches Word's status bar; Words.Count
            # counted punctuation and paragraph marks (~18% high).
            "words": int(doc.ComputeStatistics(_WD_STAT_WORDS)),
            "track_revisions": bool(doc.TrackRevisions),
        }
        with contextlib.suppress(Exception):
            info["images"] = int(doc.InlineShapes.Count) + int(
                doc.Shapes.Count
            )
        with contextlib.suppress(Exception):
            props = doc.BuiltInDocumentProperties
            for com_name, key in (("Title", "title"), ("Author", "author")):
                val = str(props(com_name).Value or "")
                if val:
                    info[key] = val
        # file-mode's "parts" (package part list) has no live equivalent;
        # "words"/"track_revisions" are live extras (file word counts live in
        # the word_count tool)
        return info

    return run_live(path, "document info", body, mutating=False)


def _table_spans(doc) -> list:
    """[(start, end, table_index)] for body tables, for match attribution."""
    spans = []
    with contextlib.suppress(Exception):
        for i in range(1, doc.Tables.Count + 1):
            rng = doc.Tables(i).Range
            spans.append((rng.Start, rng.End, i - 1))
    return spans


def find_text(
    path: str, query: str, regex: bool = False, context_chars: int = 60
) -> list:
    """FILE-MODE SHAPE (L9 parity): flat list of match entries —
    {paragraph_index, match, context} for body paragraphs,
    {table_index, row, cell, match, context} for table cells, plus
    live-only labeled extras {in_sdt: true, match, context} for content
    controls/TOC galleries (file mode does not search SDT blocks). Capped
    at 500 matches; a trailing {truncated: true} sentinel entry marks the
    cut."""
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
        tables = _table_spans(doc)
        matches = []
        truncated_note = {
            "truncated": True,
            "note": "500-match cap reached; narrow the query",
        }
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
            if not found:
                continue
            # cell attribution once per paragraph, not per match
            location: dict = {}
            if in_sdt:
                location["in_sdt"] = True
            elif in_table:
                p_start = p.Range.Start
                t_idx = next(
                    (t for s, e, t in tables if s <= p_start < e), None
                )
                if t_idx is not None:
                    location["table_index"] = t_idx
                with contextlib.suppress(Exception):
                    cell = p.Range.Cells(1)
                    location["row"] = cell.RowIndex - 1
                    location["cell"] = cell.ColumnIndex - 1
            else:
                location["paragraph_index"] = body_idx
            for m in found:
                lo = max(0, m.start() - context_chars)
                hi = min(len(text), m.end() + context_chars)
                entry = dict(location)
                entry["match"] = m.group(0)
                entry["context"] = text[lo:hi]
                matches.append(entry)
                if len(matches) >= 500:
                    matches.append(truncated_note)
                    return matches
        return matches

    return _run_live_list(path, "find text", body)


def word_count(path: str, by_section: bool = True) -> dict:
    """Live word count via Word's own statistics engine (L2).

    FILE-MODE SHAPE: {totals: {words, characters, paragraphs, tables,
    cjk_chars}, counting, sections?} plus live keys. totals.words is
    ComputeStatistics(wdStatisticWords) — the number Word's status bar
    shows — which counts slightly differently from file mode's
    whitespace-token count, so the two modes can differ by a few words on
    the same content. Per-section counts are best-effort mirrors of the
    file logic (a section runs from a heading to the next heading of the
    same or higher level; nested headings' own text is excluded), computed
    with Range.ComputeStatistics."""

    def body(session):
        doc = session.doc
        paras = _body_paragraphs(doc)

        # headings: (list position, body index, level, text, paragraph)
        headings = []
        non_heading_paragraphs = 0
        for i, p in enumerate(paras):
            lvl = None
            with contextlib.suppress(Exception):
                v = p.OutlineLevel
                if 1 <= v <= 9:
                    lvl = v
            text = _para_text(p).strip()
            if lvl is not None and text:
                headings.append((i, lvl, text, p))
            else:
                non_heading_paragraphs += 1

        def _range_words(start, end) -> int:
            if start >= end:
                return 0
            try:
                return int(
                    doc.Range(start, end).ComputeStatistics(_WD_STAT_WORDS)
                )
            except Exception:
                return 0

        totals = {
            "words": int(doc.ComputeStatistics(_WD_STAT_WORDS)),
            "characters": int(
                doc.ComputeStatistics(_WD_STAT_CHARS_WITH_SPACES)
            ),
            "paragraphs": non_heading_paragraphs,
            "tables": int(doc.Tables.Count),
        }
        # CJK character count (the zh/ja academic unit) from the body text —
        # best-effort live mirror of the file layer's cjk counting
        from ..ops.localization import cjk_aware_word_count

        body_text = ""
        with contextlib.suppress(Exception):
            body_text = doc.Content.Text
        cjk = cjk_aware_word_count(body_text)
        totals["cjk_chars"] = cjk["cjk_chars"]
        non_cjk = cjk["words"] - cjk["cjk_chars"]
        result = {
            "totals": totals,
            "counting": (
                "spaces"
                if not totals["cjk_chars"]
                else ("cjk" if not non_cjk else "mixed")
            ),
            "note": (
                "counted live by Word (ComputeStatistics — matches Word's "
                "status bar); file mode counts whitespace tokens, so the "
                "modes can differ by a few words on identical content"
            ),
        }
        if by_section:
            sections = []
            doc_end = doc.Content.End
            for pos, (i, lvl, text, p) in enumerate(headings):
                # section body: from the end of this heading's paragraph to
                # the start of the next heading of the same-or-higher level
                sec_start = p.Range.End
                sec_end = doc_end
                for j, jlvl, _, jp in headings[pos + 1:]:
                    if jlvl <= lvl:
                        sec_end = jp.Range.Start
                        break
                words = _range_words(sec_start, sec_end)
                # exclude nested sub-headings' own text, like the file layer
                for j, jlvl, _, jp in headings[pos + 1:]:
                    js, je = jp.Range.Start, jp.Range.End
                    if js >= sec_end:
                        break
                    if js >= sec_start:
                        words -= _range_words(js, je)
                sections.append(
                    {
                        "heading": text,
                        "level": lvl,
                        "paragraph_index": i,
                        "words": max(0, words),
                    }
                )
            result["sections"] = sections
        return result

    return run_live(path, "word count", body, mutating=False)


def get_comments(path: str, author: str | None = None) -> list:
    """Live comments read (L3). FILE-MODE SHAPE: flat list of {id, author,
    initials, date, text, anchored_text, resolved, reply_to}. Live ids are
    the comment's POSITION (Word's 1-based Comment.Index as a string) —
    COM does not expose the XML w:id — so ids are stable within one read
    but are not the file layer's XML ids."""

    def body(session):
        doc = session.doc
        out = []
        index_ids = {}
        for i in range(1, doc.Comments.Count + 1):
            c = doc.Comments(i)
            cid = str(c.Index)
            index_ids[c.Index] = cid
            entry = {"id": cid}
            with contextlib.suppress(Exception):
                entry["author"] = c.Author or ""
            with contextlib.suppress(Exception):
                entry["initials"] = c.Initial or ""
            entry.setdefault("author", "")
            entry.setdefault("initials", "")
            try:
                entry["date"] = c.Date.isoformat()
            except Exception:
                try:
                    entry["date"] = str(c.Date)
                except Exception:
                    entry["date"] = ""
            try:
                entry["text"] = (
                    c.Range.Text.replace("\r", "\n").replace("\x07", "")
                ).strip()
            except Exception:
                entry["text"] = ""
            try:
                entry["anchored_text"] = (
                    c.Scope.Text.replace("\r", "").replace("\x07", "")
                )
            except Exception:
                entry["anchored_text"] = ""
            try:
                entry["resolved"] = bool(c.Done)
            except Exception:
                entry["resolved"] = False
            reply_to = None
            with contextlib.suppress(Exception):
                anc = c.Ancestor
                if anc is not None:
                    reply_to = str(anc.Index)
            entry["reply_to"] = reply_to
            out.append(entry)
        if author is not None:
            out = [e for e in out if e["author"] == author]
        return out

    return _run_live_list(path, "read comments", body)


def replace_paragraph_text(
    path: str, index: int, new_text: str, expect: str | None = None,
    verify_text: str | None = None,
) -> dict:
    """Live full-paragraph text replacement (L4). FILE-MODE SHAPE:
    {replaced_paragraph: index, replaced_text: old}. expect guards stale
    indices: when given, the paragraph's current text must contain it or
    the call refuses with nothing changed (Bug 11: index-addressed
    replacement after insert/delete shifts silently hits the wrong
    paragraph). The paragraph MARK is never touched, so
    the style and any section break riding the mark survive; the first
    character's formatting carries into the replacement (Word's assignment
    semantics, matching the file layer's keep-first-run-format).
    Paragraphs carrying tracked revisions are refused (accept/reject first,
    or edit the closed file) — replacing part-deleted/part-inserted text
    live has no faithful semantics. Bare fields and content controls inside
    the paragraph are deleted explicitly first (Range.Text assignment
    silently leaves both behind)."""
    return run_live(
        path, "replace paragraph text",
        replace_paragraph_text_body(index, new_text, expect,
                                    verify_text=verify_text),
    )


def replace_paragraph_text_body(
    index: int, new_text: str, expect: str | None = None,
    verify_text: str | None = None,
):
    """Session-level body factory; the apply_edits live route composes
    these inside ONE undo group. Same code the public tool runs."""
    check_text_safe(new_text)

    def body(session):
        doc = session.doc
        paras = _body_paragraphs(doc)
        if not 0 <= index < len(paras):
            raise TargetNotFound(
                f"no body paragraph with index {index}; the document has "
                f"{len(paras)} body paragraphs"
            )
        _stale_guard(session, paras, index, verify_text, "set_text")
        p = paras[index]
        rng = doc.Range(p.Range.Start, max(p.Range.Start, p.Range.End - 1))
        current = rng.Text or ""
        if expect is not None and expect not in current:
            raise TargetNotFound(
                f"paragraph {index} does not contain the expected text "
                f"{expect[:80]!r}; its current text begins: "
                f"{current[:120]!r}. Paragraph indices shift after "
                "insert/delete operations - re-read with get_text, or use "
                "search_and_replace for text-anchored replacement. "
                "Nothing was changed."
            )
        try:
            has_revisions = bool(rng.Revisions.Count)
        except Exception:
            has_revisions = False
        if has_revisions:
            raise UnsupportedStructure(
                f"paragraph {index} carries tracked revisions; accept or "
                "reject them first (accept_revisions/reject_revisions on "
                "the closed file, or in Word), then retry"
            )
        # content controls: refuse any that extend beyond the paragraph,
        # then delete contained ones WITH contents (text assignment leaves
        # them behind otherwise)
        ccs = []
        with contextlib.suppress(Exception):
            ccs = list(rng.ContentControls)
        for cc in ccs:
            if cc.Range.Start < rng.Start or cc.Range.End > rng.End:
                raise UnsupportedStructure(
                    "a content control extends beyond this paragraph; "
                    "replacing would destroy content outside it"
                )
        removed = False
        for cc in ccs:
            cc.Delete(True)
            removed = True
        # bare fields survive Range.Text writes — delete explicitly
        with contextlib.suppress(Exception):
            for fld in list(rng.Fields):
                fld.Delete()
                removed = True
        if removed:  # boundaries shifted; re-resolve the paragraph
            p = _body_paragraphs(doc)[index]
            rng = doc.Range(
                p.Range.Start, max(p.Range.Start, p.Range.End - 1)
            )
        _assign_text(rng, new_text)
        result = {"replaced_paragraph": index, "replaced_text": current}
        with contextlib.suppress(Exception):
            if doc.TrackRevisions:
                result["note"] = (
                    "the document's track-changes is ON, so the replacement "
                    "was recorded as tracked changes"
                )
        return result

    return body


def insert_paragraphs(
    path: str,
    paragraphs: list[dict],
    after_index: int | None = None,
    before_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    track: bool = False,
    author: str = "Claude",
    verify_text: str | None = None,
) -> dict:
    return run_live(
        path, "insert paragraphs",
        insert_paragraphs_body(
            paragraphs, after_index=after_index, before_index=before_index,
            after_anchor=after_anchor, at_end=at_end, track=track,
            author=author, verify_text=verify_text,
        ),
    )


def insert_paragraphs_body(
    paragraphs: list[dict],
    *,
    after_index: int | None = None,
    before_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    track: bool = False,
    author: str = "Claude",
    verify_text: str | None = None,
):
    """Session-level body factory; see replace_paragraph_text_body."""
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
            _stale_guard(session, paras, after_index, verify_text, "insert")
            point = paras[after_index].Range.End - 1
            prefix = True
        else:
            if not 0 <= before_index < len(paras):
                raise TargetNotFound(
                    f"before_index {before_index} out of range "
                    f"({len(paras)} body paragraphs)"
                )
            _stale_guard(session, paras, before_index, verify_text, "insert")
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
            if spec.get("heading_level"):
                # built-in heading via the wdStyleHeadingN constant
                # (-2 .. -10): locale-independent, unlike the name form
                level = int(spec["heading_level"])
                styled.Style = doc.Styles(-(level + 1))
            elif spec.get("outline_heading"):
                # outline-based heading (academic-template pattern): keep
                # the Normal look, set the outline level directly
                styled.Style = doc.Styles(-1)  # wdStyleNormal
                first = styled.Paragraphs(1)
                first.OutlineLevel = int(spec["outline_heading"])
            elif spec.get("style"):
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

    return body


def delete_paragraphs(
    path: str,
    start: int,
    end: int | None = None,
    track: bool = False,
    author: str = "Claude",
    verify_start_text: str | None = None,
    verify_end_text: str | None = None,
) -> dict:
    return run_live(
        path, "delete paragraphs",
        delete_paragraphs_body(start, end, track=track, author=author,
                               verify_start_text=verify_start_text,
                               verify_end_text=verify_end_text),
    )


def delete_paragraphs_body(
    start: int,
    end: int | None = None,
    *,
    track: bool = False,
    author: str = "Claude",
    verify_start_text: str | None = None,
    verify_end_text: str | None = None,
):
    """Session-level body factory; see replace_paragraph_text_body."""
    def body(session):
        doc = session.doc
        last = start if end is None else end
        paras = _body_paragraphs(doc)
        if not (0 <= start <= last < len(paras)):
            raise TargetNotFound(
                f"paragraph range [{start}..{last}] out of range "
                f"({len(paras)} body paragraphs)"
            )
        _stale_guard(session, paras, start, verify_start_text, "delete")
        if last != start:
            _stale_guard(session, paras, last, verify_end_text, "delete")
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
        # file-mode parity: tracked deletes report deleted_tracked (nothing
        # is removed until the revision is accepted), untracked report deleted
        payload = (
            {"deleted_tracked": deleted} if track else {"deleted": deleted}
        )
        return _tracked_result(session, track, author, payload)

    return body


# ------------------------------------------------------------------ tables


def set_cells(
    path: str,
    table_index: int,
    edits: list[dict],
    track: bool = False,
    author: str = "Claude",
) -> dict:
    return run_live(
        path, "set cells",
        set_cells_body(table_index, edits, track=track, author=author),
    )


def set_cells_body(
    table_index: int,
    edits: list[dict],
    *,
    track: bool = False,
    author: str = "Claude",
):
    """Session-level body factory; see replace_paragraph_text_body."""
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
            # "cells_written" is the file-mode key (L7 parity); "table" is a
            # live-only additive echo of the target
            {"cells_written": applied, "table": table_index},
        )

    return body


# -------------------------------------------------------------- formatting


def format_text(
    path: str,
    formatting: dict,
    paragraph_index: int | None = None,
    find: str | None = None,
    occurrence: int = 1,
    verify_text: str | None = None,
) -> dict:
    return run_live(
        path, "format text",
        format_text_body(
            formatting, paragraph_index=paragraph_index, find=find,
            occurrence=occurrence, verify_text=verify_text,
        ),
    )


def format_text_body(
    formatting: dict,
    *,
    paragraph_index: int | None = None,
    find: str | None = None,
    occurrence: int = 1,
    verify_text: str | None = None,
):
    """Session-level body factory; see replace_paragraph_text_body."""
    unknown = set(formatting) - _CHAR_KEYS
    if unknown:
        raise WordMcpError(
            f"unknown character-formatting key(s) {sorted(unknown)}; "
            f"allowed: {sorted(_CHAR_KEYS)}"
        )
    if formatting.get("superscript") and formatting.get("subscript"):
        raise WordMcpError("superscript and subscript are mutually exclusive")
    for key in ("language", "east_asian_language"):
        if key in formatting:
            _lcid(formatting[key])  # validate the tag BEFORE touching Word
    if "language" in formatting:
        primary = str(formatting["language"]).lower().split("-")[0]
        if primary in _EAST_ASIAN_PRIMARY:
            raise WordMcpError(
                f"language {formatting['language']!r} is an East-Asian "
                "proofing language, which Word stores in the east-asian "
                "slot and silently IGNORES in the latin slot — pass it as "
                "east_asian_language instead (matches Word's own behavior)"
            )
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
            _stale_guard(session, paras, paragraph_index, verify_text,
                         "format")
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

    return body


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
    if find is not None and len(find) > _FIND_TEXT_LIMIT:
        raise WordMcpError(
            f"find string is {len(find)} characters; Word's Find accepts at "
            f"most ~{_FIND_TEXT_LIMIT} — scroll with a shorter unique prefix"
        )

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
    # proofing language (the live twin of the file layer's w:lang keys);
    # ko-KR and friends map through the LCID table. NOTE: the language
    # properties live on the RANGE, not the Font — Word's Font object has
    # no LanguageID, and assigning one onto a gen_py Font proxy silently
    # sets a phantom Python attribute (verified 2026-08-28).
    if "language" in fmt:
        rng.LanguageID = _lcid(fmt["language"])
    if "east_asian_language" in fmt:
        rng.LanguageIDFarEast = _lcid(fmt["east_asian_language"])


# ---------------------------------------------------------------------------
# LIVE-PARITY AUDIT (WS-L, 2026-08-28) — source material for the docstring
# pass. Status of every file-mode tool family against the live pathway.
#
# HAS A LIVE ROUTE (dual-mode via server._route_live; result = file-mode
# shape + "live": true, list-shaped reads stay flat lists):
#   get_document_info, get_text, get_outline, find_text, word_count,
#   get_comments, search_and_replace, insert_paragraphs, delete_paragraphs,
#   replace_paragraph_text, format_text, set_cells
# LIVE-ONLY: live_insert_at_cursor, live_scroll_to, live_set_track_changes,
#   word_live_repair (+ the com_* bridge tools, which run their own
#   invisible instance or message the user's).
#
# COULD GAIN A LIVE ROUTE (COM exposes the data; not yet built — candidates
# for a later slice, roughly in order of value):
#   read-only: get_styles (doc.Styles), list_tables/get_table (doc.Tables),
#     list_footnotes/list_endnotes (doc.Footnotes/Endnotes),
#     get_tracked_changes/revision_summary/revision_analytics
#     (doc.Revisions), list_bookmarks (doc.Bookmarks), get_headers_footers
#     (section.Headers/Footers), list_images (InlineShapes/Shapes),
#     get_protection (doc.ProtectionType), list_sections
#     (doc.Sections), read_toc (TablesOfContents), get_textbox_text
#     (Shapes.TextFrame), list_form_fields (FormFields/ContentControls),
#     word_count_with_exclusions (Range arithmetic, laborious but possible)
#   mutating: apply_style/set_paragraph_format (Paragraph.Style/Format),
#     add_comment/reply/resolve/delete (Comments), accept/reject_revisions
#     (Revisions), insert/delete_rows/columns (Table.Rows/Columns),
#     add_heading/add_page_break (insert + style), change_case
#     (Range.Case), format_cells (Cell.Range.Font)
#
# MUST REFUSE LIVE, AND WHY (structural XML surgery or whole-package
# operations with no faithful COM equivalent on an OPEN document):
#   diagnose_document (reads the saved package's XML — stale + locked while
#     Word holds unsaved changes; com_validate_opens_clean / live
#     get_document_info are the open-document checks),
#   split_document / copy_document / apply_template / fill_template /
#     mail_merge (whole-file operations on the package),
#   redact_text / verify_redaction / anonymize_for_review / prepare_for_
#     submission (guarantee depends on rewriting the saved XML incl. rsids
#     and metadata),
#   convert_notes / cleanup_orphan_notes / notes CRUD (note-part XML
#     rewrites; COM note objects exist but the file layer's id-integrity
#     guarantees do not translate),
#   insert_toc/delete_toc/insert_index & friends, insert_citation/
#     bibliography tools (field+SDT machinery Word regenerates; gallery
#     SDTs are invisible to COM — see _sdt_regions),
#   structured_diff / compare tools (com_compare_documents is the live-ish
#     equivalent via a dedicated invisible instance),
#   backup/manage_backups (the open file's bytes are stale by definition;
#     Word AutoRecover owns the open document),
#   protection set/remove (Word ignores or dialogs on programmatic
#     protection changes to the active document; password flows especially).
# ---------------------------------------------------------------------------


# ---- paragraph formatting (Bug 12: completes the live citation workflow) ----

_WD_ALIGN_MAP = {
    "left": 0, "center": 1, "right": 2, "justify": 3,
}

_PARA_FMT_KEYS_LIVE = {
    "alignment", "space_before_pt", "space_after_pt", "line_spacing",
    "indent_left_pt", "indent_right_pt", "first_line_indent_pt",
    "keep_with_next", "keep_lines_together", "page_break_before",
    "widow_control", "outline_level",
}

_PARA_FMT_UNSUPPORTED_LIVE = {"shading", "borders", "tab_stops"}


def set_paragraph_format(
    path: str, indices: list[int], formatting: dict
) -> dict:
    """Live paragraph formatting via COM Paragraph.Format properties."""
    return run_live(
        path, "set paragraph format",
        set_paragraph_format_body(indices, formatting),
    )


def set_paragraph_format_body(indices: list[int], formatting: dict,
                              verify_texts: list[str | None] | None = None):
    """Session-level body factory; see replace_paragraph_text_body.
    verify_texts, when given, is parallel to indices (None entries skip)."""
    unknown = set(formatting) - _PARA_FMT_KEYS_LIVE - _PARA_FMT_UNSUPPORTED_LIVE
    if unknown:
        raise WordMcpError(
            f"unknown paragraph-formatting key(s) {sorted(unknown)}; "
            f"allowed (live): {sorted(_PARA_FMT_KEYS_LIVE)}"
        )
    unsupported = set(formatting) & _PARA_FMT_UNSUPPORTED_LIVE
    if unsupported:
        raise WordMcpError(
            f"paragraph-formatting key(s) {sorted(unsupported)} are not "
            "supported in live mode (they require XML-level operations). "
            "Close the document and use file-mode set_paragraph_format, or "
            "apply the formatting manually in Word."
        )
    if "alignment" in formatting:
        val = formatting["alignment"]
        if val not in _WD_ALIGN_MAP:
            raise WordMcpError(
                f"alignment must be one of {list(_WD_ALIGN_MAP)}, got {val!r}"
            )
    if "outline_level" in formatting:
        lvl = formatting["outline_level"]
        if lvl is not None and (
            isinstance(lvl, bool) or not isinstance(lvl, int)
            or not 0 <= lvl <= 8
        ):
            raise WordMcpError(
                "outline_level must be an integer 0-8 or null to remove"
            )

    def body(session):
        doc = session.doc
        paras = _body_paragraphs(doc)
        applied = []
        for pos, idx in enumerate(indices):
            if not 0 <= idx < len(paras):
                raise TargetNotFound(
                    f"paragraph index {idx} out of range "
                    f"({len(paras)} body paragraphs)"
                )
            if verify_texts is not None and pos < len(verify_texts):
                _stale_guard(session, paras, idx, verify_texts[pos],
                             "set_paragraph_format")
            p = paras[idx]
            fmt = p.Format
            keys_set = []
            if "alignment" in formatting:
                fmt.Alignment = _WD_ALIGN_MAP[formatting["alignment"]]
                keys_set.append("alignment")
            if "space_before_pt" in formatting:
                fmt.SpaceBefore = float(formatting["space_before_pt"])
                keys_set.append("space_before_pt")
            if "space_after_pt" in formatting:
                fmt.SpaceAfter = float(formatting["space_after_pt"])
                keys_set.append("space_after_pt")
            if "line_spacing" in formatting:
                val = float(formatting["line_spacing"])
                if val <= 6:
                    fmt.LineSpacingRule = 5  # wdLineSpaceMultiple
                    fmt.LineSpacing = val * 12
                else:
                    fmt.LineSpacingRule = 4  # wdLineSpaceExactly
                    fmt.LineSpacing = val
                keys_set.append("line_spacing")
            if "indent_left_pt" in formatting:
                fmt.LeftIndent = float(formatting["indent_left_pt"])
                keys_set.append("indent_left_pt")
            if "indent_right_pt" in formatting:
                fmt.RightIndent = float(formatting["indent_right_pt"])
                keys_set.append("indent_right_pt")
            if "first_line_indent_pt" in formatting:
                fmt.FirstLineIndent = float(formatting["first_line_indent_pt"])
                keys_set.append("first_line_indent_pt")
            if "keep_with_next" in formatting:
                fmt.KeepWithNext = bool(formatting["keep_with_next"])
                keys_set.append("keep_with_next")
            if "keep_lines_together" in formatting:
                fmt.KeepTogether = bool(formatting["keep_lines_together"])
                keys_set.append("keep_lines_together")
            if "page_break_before" in formatting:
                fmt.PageBreakBefore = bool(formatting["page_break_before"])
                keys_set.append("page_break_before")
            if "widow_control" in formatting:
                fmt.WidowControl = bool(formatting["widow_control"])
                keys_set.append("widow_control")
            if "outline_level" in formatting:
                lvl = formatting["outline_level"]
                if lvl is None:
                    p.OutlineLevel = 10  # wdOutlineLevelBodyText
                else:
                    p.OutlineLevel = lvl + 1  # COM: 1-9 = Heading 1-9
                keys_set.append("outline_level")
            applied.append({"paragraph": idx, "keys_set": keys_set})
        return {"paragraphs_formatted": len(applied), "applied": applied}

    return body
