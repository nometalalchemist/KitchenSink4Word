"""word-mcp: FastMCP server exposing the full Word editing suite.

Contract for every mutating tool:
- `file_path` is the target .docx (absolute path).
- Before each mutation the current content is rotated into stable backup
  slots (prev.docx / anchor.docx) inside a hidden .ks4w-backups/ folder next
  to the file, unless backup=False (see core.safesave; manage_backups
  lists/restores/purges them).
- Mutations of one file are serialized (in-process mutex + cross-process
  advisory lockfile), so parallel calls see each other's results.
- Saves are atomic and validated; on any failure the original is untouched
  and the error message says exactly what was wrong.
- A file open in Word is edited live by dual-mode tools (live='auto');
  tools with no live route refuse with a clear message until it is closed.
"""

from __future__ import annotations

from typing import Any

from fastmcp import FastMCP

from .core.package import DocxPackage
from .ops import (
    bibliography as _bib,
    citecheck as _cc,
    comments as _cm,
    fields as _fl,
    furniture as _fu,
    lists as _ls,
    media as _md,
    notes as _nt,
    protection as _pr,
    read as _rd,
    structure as _sx,
    template as _tp,
    revisions as _rv,
    stats as _st,
    tables as _tb,
    text as _tx,
    toc as _tc,
    accessibility as _ac,
    batch as _bt,
    chapterheaders as _ch,
    cleanup as _cu,
    compliance as _cp,
    definedterms as _dt,
    diagnostics as _dg,
    forms as _fm,
    frontmatter as _fma,
    integrity as _ig,
    mailmerge as _mm,
    redaction as _rx,
    reffields as _rf,
    anonymize as _an,
    assembly as _asm,
    backups as _bk,
    charts as _charts,
    citesystem as _cs,
    dataio as _dio,
    workflows as _wf,
    equations as _eq,
    journalcount as _jc,
    preview as _pv,
    reviewcycle as _rc,
    styleconvert as _sc2,
    stylefind as _sf,
    textboxes as _tbx,
    zoterolib as _zl,
)

mcp = FastMCP(
    "kitchensink4word",
    instructions=(
        "Full-featured Word (.docx) editor: text, tables (including column "
        "insert/delete and bulk cell edits), footnotes/endnotes, TOC, "
        "headers/footers, images, charts, citations, comments, tracked "
        "changes, document assembly. File-based with auto-backup before "
        "every mutation; dual-mode tools edit documents open in Word live "
        "(live='auto'), tools without a live route refuse until the file "
        "is closed."
    ),
)


def _edit(file_path: str, fn, *, backup: bool = True) -> dict:
    from .core.safesave import write_lock

    # Serialize the full read-modify-validate-save cycle per file: concurrent
    # mutations of one document otherwise race (stale response metadata, lost
    # updates). Covers threads in this process AND other server processes.
    with write_lock(file_path):
        pkg = DocxPackage(file_path)
        result = fn(pkg)
        saved = pkg.save(do_backup=backup)
    result["saved"] = str(saved)
    return result


def _route_live(live: str, file_call, live_call) -> dict:
    """live='auto': file-based first, live layer when the doc is open in Word.
    'force': straight to the live layer. 'off': v1 behavior (locked = refuse)."""
    from .core.errors import DocumentLocked

    if live not in ("auto", "force", "off"):
        raise ValueError(f"live must be auto|force|off, got {live!r}")
    if live == "force":
        return live_call()
    try:
        return file_call()
    except DocumentLocked:
        if live == "off":
            raise
        return live_call()


# ------------------------------------------------------------------- document


@mcp.tool
def get_document_info(file_path: str, live: str = "auto") -> dict:
    """Read a one-call document overview:
    paragraph/table/footnote/comment/revision counts, sections, and package
    parts. Documents open in Word are read live (same key names; live adds
    'words' from Word's own ComputeStatistics counter plus track_revisions,
    and omits the part list). Read-only.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_document_info(DocxPackage(file_path)),
        lambda: _lo.get_document_info(file_path),
    )


@mcp.tool
def get_outline(file_path: str, live: str = "auto") -> list | dict:
    """List every heading with its body paragraph index and level. Detects
    both heading systems: built-in Heading styles AND w:outlineLvl
    overrides (direct or style-inherited, the academic-template pattern on
    Normal-styled paragraphs); detected_via on each entry names which.
    Documents open in Word are read live (same flat-list shape). Read-only.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_outline(DocxPackage(file_path)),
        lambda: _lo.get_outline(file_path),
    )


@mcp.tool
def get_text(
    file_path: str,
    start: int = 0,
    end: int | None = None,
    contains: str | None = None,
    include_textboxes: bool = False,
    live: str = "auto",
) -> list | dict:
    """Read body paragraphs as [{index, text, style, ...}]; every paragraph
    reports its effective style, including the default (usually Normal).
    start/end slice by paragraph index (0-based, end EXCLUSIVE); contains
    filters to matching paragraphs. include_textboxes=True appends text-box
    content as labeled extra entries (source 'textbox', box_index
    addressing set_textbox_text) without touching body indices; file-mode
    only. Documents open in Word are read live with the same shape (live
    styles are Word's localized display names; file styles are style ids).
    Equations are invisible here (list_equations reads those).
    """
    from .com import live_ops as _lo

    def _live_call():
        if include_textboxes:
            from .core.errors import WordMcpError

            raise WordMcpError(
                "include_textboxes is file-mode only; close the document in "
                "Word, use get_textbox_text, or pass include_textboxes=False"
            )
        return _lo.get_text(file_path, start, end, contains)

    return _route_live(
        live,
        lambda: _rd.get_paragraphs(
            DocxPackage(file_path), start, end, contains=contains,
            include_textboxes=include_textboxes,
        ),
        _live_call,
    )


@mcp.tool
def find_text(
    file_path: str,
    query: str,
    regex: bool = False,
    context_chars: int = 60,
    include_textboxes: bool = False,
    live: str = "auto",
) -> list | dict:
    """Find text in paragraphs and table cells; returns locations + context.
    include_textboxes=True also searches text-box/shape content, returned as
    clearly-labeled extra matches (source 'textbox' with box_index) without
    affecting body match entries or paragraph indices; file-mode only.
    Documents open in Word are searched live (current in-memory state, same
    flat-list shape; live additionally reports matches inside content
    controls as labeled in_sdt entries, and caps at 500 matches with a
    trailing truncated sentinel entry)."""
    from .com import live_ops as _lo

    def _live_call():
        if include_textboxes:
            from .core.errors import WordMcpError

            raise WordMcpError(
                "include_textboxes is file-mode only; close the document in "
                "Word, use get_textbox_text, or pass include_textboxes=False"
            )
        return _lo.find_text(file_path, query, regex, context_chars)

    return _route_live(
        live,
        lambda: _rd.find_text(
            DocxPackage(file_path), query, regex=regex,
            context_chars=context_chars, include_textboxes=include_textboxes,
        ),
        _live_call,
    )


@mcp.tool
def get_styles(file_path: str) -> list:
    """List the styles defined in the document: id, name, type, based_on, plus
    each style's EXPLICITLY defined paragraph_formatting and
    character_formatting in the exact shape define_style accepts (template
    cloning is one read + one define). Inherited values are not
    synthesized; follow the based_on chain. Exact/atLeast line spacing and
    theme font/color references have no define_style representation and are
    omitted. Read-only.
    """
    return _rd.list_styles(DocxPackage(file_path))


@mcp.tool
def word_count(file_path: str, by_section: bool = True, live: str = "auto") -> dict:
    """Count words/characters/paragraphs, total and per heading section.
    Documents open in Word are counted live via Word's own statistics
    engine (ComputeStatistics, the status-bar number); file mode counts
    whitespace tokens, so the two can differ by a few words on identical
    content. Live per-section counts are best-effort mirrors of the
    file-mode logic. For journal counts that exclude references, captions,
    or other zones, use word_count_with_exclusions. Read-only.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _st.word_count(DocxPackage(file_path), by_section=by_section),
        lambda: _lo.word_count(file_path, by_section=by_section),
    )


@mcp.tool
def check_citation_parity(file_path: str) -> dict:
    """Cross-check APA in-text citations against the reference list in both
    directions: cited-but-not-listed (serious) and listed-but-never-cited
    (review). Heuristic flagging: results are review candidates, not
    verdicts. Read-only.
    """
    return _cc.check_citation_parity(DocxPackage(file_path))


@mcp.tool
def validate_document(file_path: str) -> dict:
    """Run a quick structural check: the package opens, footnotes/endnotes are
    consistent, and field begin/end markers balance. Returns {package_ok,
    notes, fields_balanced, field_begins, field_ends}. For the deep
    multi-check health report (relationships, orphan parts, undefined
    styles, bookmarks, images), use diagnose_document; for Word's own
    verdict use com_validate_opens_clean. Read-only.
    """
    pkg = DocxPackage(file_path)
    notes_report = _nt.validate_notes(pkg)
    xml = pkg.raw_part("word/document.xml").decode("utf-8", errors="replace")
    begins = xml.count('w:fldCharType="begin"')
    ends = xml.count('w:fldCharType="end"')
    return {
        "package_ok": True,
        "notes": notes_report,
        "fields_balanced": begins == ends,
        "field_begins": begins,
        "field_ends": ends,
    }


@mcp.tool
def create_document(file_path: str, title: str | None = None) -> dict:
    """Create a new blank .docx file, optionally setting the title core
    property. Refuses to overwrite an existing file (use copy_document with
    overwrite=True for that). Parent directories are created automatically.
    Populate the document afterward with insert_paragraphs, create_table,
    define_style, and other tools."""
    from pathlib import Path

    from docx import Document

    from .core.sandbox import check_path

    p = Path(check_path(file_path, "create document"))
    if p.exists():
        raise FileExistsError(f"{file_path} already exists")
    doc = Document()
    if title:
        doc.core_properties.title = title
    p.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(p))
    return {"created": str(p)}


@mcp.tool
def copy_document(
    file_path: str, dest_path: str, overwrite: bool = False
) -> dict:
    """Copy a document byte-for-byte, e.g. to a new DTG-stamped filename
    before editing (create_snapshot names such a copy for you; this tool
    takes an explicit dest_path). Refuses an existing dest_path unless
    overwrite=True; an overwritten destination's previous content rotates
    into its .ks4w-backups prev slot first, so the overwrite is undoable
    via manage_backups restore.
    """
    import shutil
    from pathlib import Path

    from .core import safesave
    from .core.sandbox import check_path

    file_path = check_path(file_path, "copy source (read)")
    dest_path = check_path(dest_path, "copy destination")
    dest = Path(dest_path)
    overwrote = False
    if dest.exists():
        if not overwrite:
            raise FileExistsError(
                f"{dest_path} already exists (pass overwrite=True to "
                "replace it; its current content will be kept in the "
                ".ks4w-backups prev slot)"
            )
        with safesave.write_lock(dest_path):
            safesave.rotate_slots(dest_path)
            # prev.docx may be a hardlink to dest: never write into dest
            # in place (it would write through the link and corrupt the
            # backup); copy aside, then atomically replace.
            tmp_copy = dest.with_name(dest.name + ".ks4w-copy-tmp")
            shutil.copy2(file_path, tmp_copy)
            safesave.replace_with_retry(tmp_copy, dest_path)
        overwrote = True
    else:
        shutil.copy2(file_path, dest_path)
    return {"copied_to": dest_path, "overwrote_existing": overwrote}


@mcp.tool
def manage_backups(
    action: str,
    file_path: str | None = None,
    directory: str | None = None,
    source: str | None = None,
    scope: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Manage the automatic backups under the hidden .ks4w-backups/ folder
    next to each mutated document: two stable slots per document, prev
    (state before the most recent mutation) and anchor (session start).
    Snapshots from create_snapshot are permanent keepers this tool never
    touches.

    action='list': slot files with sizes and mtimes, legacy *.bak-* files
    from the pre-v1.6 scheme, and orphaned slot folders whose source
    document is gone; give file_path for one document or directory for a
    folder.

    action='restore': overwrite file_path with a backup; source is 'prev',
    'anchor', or a legacy *.bak-* path. The current content rotates into
    prev FIRST, so a restore is itself undoable. Refuses documents open in
    Word; the payload is validated before the atomic replace.

    action='purge': delete backups; scope: 'legacy', 'orphans', or 'slots'.
    dry_run defaults to TRUE (report only); dry_run=False deletes. Exact
    paths and sizes are reported either way.
    """
    return _bk.manage_backups(
        action,
        file_path=file_path,
        directory=directory,
        source=source,
        scope=scope,
        dry_run=dry_run,
    )


# ----------------------------------------------------------------------- text


@mcp.tool
def search_and_replace(
    file_path: str,
    replacements: list[dict],
    scope: str = "body",
    max_replacements: int | None = None,
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Batch find/replace, safe across Word's fragmented runs. Each item:
    {find, replace, regex?}; scope: body | footnotes | headers | all.
    max_replacements aborts, changing nothing, when total matches would
    exceed it: set it for broad regexes (preview_replace dry-runs this tool
    and yields the count). track records each replacement as a tracked
    change by `author`. Siblings: replace_formatted replaces only text with
    specific formatting; replace_paragraph_text rewrites one whole
    paragraph by index. Live edits appear immediately as one Ctrl+Z step,
    unsaved until the user saves; the live result adds "live": true and
    skip counters, and literal finds beyond Word's ~255-char limit are
    handled automatically. Auto-backup in file mode: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Documents open in Word are edited live.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _search_and_replace_file(
            file_path, replacements, scope, max_replacements, track, author,
            backup,
        ),
        lambda: _lo.search_and_replace(
            file_path, replacements, scope, max_replacements, track, author
        ),
    )


def _search_and_replace_file(
    file_path, replacements, scope, max_replacements, track, author, backup
) -> dict:
    return _edit(
        file_path,
        lambda pkg: _tx.search_and_replace(
            pkg,
            replacements,
            scope=scope,
            max_replacements=max_replacements,
            track=track,
            author=author,
        ),
        backup=backup,
    )


@mcp.tool
def insert_paragraphs(
    file_path: str,
    paragraphs: list[dict],
    after_index: int | None = None,
    before_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    inherit_format: bool = False,
    copy_format_from: int | None = None,
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Insert paragraphs ({text, style?, formatting?}) at ONE position:
    after_index | before_index | after_anchor | at_end. Anchors match a
    paragraph's PLAIN text (write '&', never the XML entity '&amp;') and
    must be unique; heading text recurs in body prose, so prefer index
    addressing for structural work. On a fresh document, index 0 addresses
    the implicit empty paragraph. inherit_format=True clones the anchor
    paragraph's direct formatting onto the inserted paragraphs (a reference
    entry matches its neighbors' hanging indent and font in one call); with
    before_index/at_end use copy_format_from=<body paragraph index> instead
    (mutually exclusive). Explicit per-item style/formatting wins; track
    records tracked insertions by `author`. Format cloning is file-mode
    only. Auto-backup in file mode: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Documents
    open in Word are edited live.
    """
    from .com import live_ops as _lo

    def _live() -> dict:
        if inherit_format or copy_format_from is not None:
            from .core.errors import WordMcpError

            raise WordMcpError(
                "inherit_format/copy_format_from are file-mode features and "
                "this document is open in Word: close it in Word and retry, "
                "or insert without format cloning"
            )
        return _lo.insert_paragraphs(
            file_path,
            paragraphs,
            after_index=after_index,
            before_index=before_index,
            after_anchor=after_anchor,
            at_end=at_end,
            track=track,
            author=author,
        )

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.insert_paragraphs(
                pkg,
                paragraphs,
                after_index=after_index,
                before_index=before_index,
                after_anchor=after_anchor,
                at_end=at_end,
                inherit_format=inherit_format,
                copy_format_from=copy_format_from,
                track=track,
                author=author,
            ),
            backup=backup,
        ),
        _live,
    )


@mcp.tool
def delete_paragraphs(
    file_path: str,
    start: int,
    end: int | None = None,
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Delete body paragraphs start..end (0-based, INCLUSIVE; end defaults to
    start). Refuses ranges that cut through a field or carry a section
    break. Deleting every paragraph leaves one empty paragraph behind (a
    document always keeps one). track marks tracked deletions by `author`
    instead of removing (result key deleted_tracked). Auto-backup in file
    mode: prev/anchor slots in .ks4w-backups (backup=False skips rotation
    only); atomic validated save. Documents open in Word are edited live.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.delete_paragraphs(
                pkg, start, end, track=track, author=author
            ),
            backup=backup,
        ),
        lambda: _lo.delete_paragraphs(
            file_path, start, end, track=track, author=author
        ),
    )


@mcp.tool
def replace_paragraph_text(
    file_path: str, index: int, new_text: str, expect: str | None = None,
    backup: bool = True, live: str = "auto",
) -> dict:
    """Replace one paragraph's full text, keeping style and base formatting.
    IMPORTANT: paragraph indices SHIFT after insert/delete operations -
    pass expect (a substring the target paragraph currently contains) to
    refuse instead of silently hitting the wrong paragraph, and verify the
    returned replaced_text (the old text) matches what you meant to
    replace. For text-anchored replacement immune to index shifts, use
    search_and_replace; this tool is the natural fallback when a find
    string would exceed live search_and_replace's length limit (no limit
    here). Auto-backup in file mode: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Documents
    open in Word are edited live. Live edits leave the paragraph mark
    untouched (style and section breaks survive); paragraphs carrying
    tracked revisions are refused live - accept/reject first.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.replace_paragraph_text(
                pkg, index, new_text, expect=expect
            ),
            backup=backup,
        ),
        lambda: _lo.replace_paragraph_text(
            file_path, index, new_text, expect=expect
        ),
    )


@mcp.tool
def add_heading(
    file_path: str,
    text: str,
    level: int = 1,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    backup: bool = True,
) -> dict:
    """Insert a NEW heading paragraph (level 1-9; Heading styles auto-created
    if missing) positioned by after_index, after_anchor, or at_end. Anchors
    match plain paragraph text, and heading text recurs in body prose, so
    prefer after_index for structural work. To restyle an EXISTING
    paragraph use apply_style; to shift existing heading levels use
    change_heading_level. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tx.add_heading(
            pkg,
            text,
            level,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
        ),
        backup=backup,
    )


@mcp.tool
def add_page_break(file_path: str, after_index: int, backup: bool = True) -> dict:
    """Insert a page break after body paragraph after_index (0-based).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _tx.add_page_break(pkg, after_index=after_index),
        backup=backup,
    )


@mcp.tool
def format_text(
    file_path: str,
    formatting: dict,
    paragraph_index: int | None = None,
    find: str | None = None,
    occurrence: int = 1,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Apply character formatting to a text range. formatting keys: bold,
    italic, underline, strike, font, size_pt, color, highlight,
    superscript, subscript, small_caps, all_caps, hidden, double_strike,
    char_spacing_pt, kerning_pt, position_pt, language, east_asian_language
    (BCP-47; ko/ja/zh go in east_asian_language, Word's East-Asian proofing
    slot; live maps tags to LCIDs, unmapped tags are file-mode only).
    Target: paragraph_index + find (substring), find alone (occurrence
    picks the match), or the whole paragraph. For a named character style
    use apply_character_style. Auto-backup in file mode: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Documents open in Word are edited live.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.format_text(
                pkg,
                paragraph_index=paragraph_index,
                find=find,
                occurrence=occurrence,
                formatting=formatting,
            ),
            backup=backup,
        ),
        lambda: _lo.format_text(
            file_path,
            formatting,
            paragraph_index=paragraph_index,
            find=find,
            occurrence=occurrence,
        ),
    )


@mcp.tool
def set_paragraph_format(
    file_path: str, indices: list[int], formatting: dict, backup: bool = True,
    live: str = "auto",
) -> dict:
    """Set paragraph formatting on a batch of paragraphs (0-based indices).
    Keys: alignment, space_before_pt, space_after_pt, line_spacing,
    indent_left_pt, indent_right_pt, first_line_indent_pt, keep_with_next,
    outline_level. outline_level (0-8, 0 = top; null removes the override)
    sets w:outlineLvl without touching style or visual formatting: the way
    to give template headings (Normal-styled, direct-formatted) a place in
    Word's navigation pane and TOC harvesting, where apply_style would
    change their look. get_paragraph_format is the matching reader.
    Documents open in Word are edited live (shading, borders, and tab_stops
    are refused live; all other keys work). Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save."""
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.set_paragraph_format(pkg, indices, formatting),
            backup=backup,
        ),
        lambda: _lo.set_paragraph_format(file_path, indices, formatting),
    )


@mcp.tool
def apply_style(
    file_path: str, indices: list[int], style: str, backup: bool = True
) -> dict:
    """Apply a paragraph style (by id or name) to a batch of paragraphs,
    changing their full look; Heading1-9 are auto-created when missing. To
    give a paragraph an outline level WITHOUT changing its appearance use
    set_paragraph_format's outline_level; to promote/demote existing
    built-in headings use change_heading_level. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _tx.apply_style(pkg, indices, style), backup=backup
    )


@mcp.tool
def add_list(
    file_path: str,
    items: list,
    kind: str = "bullet",
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    backup: bool = True,
) -> dict:
    """Insert a bulleted or numbered list with real bullet/number glyphs
    (numbering.xml infrastructure created as needed). items: strings or
    {text, level} dicts (level 0-8 nests); kind: bullet | number. Each call
    is an independent list, so numbering restarts at 1. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _ls.add_list(
            pkg,
            items,
            kind=kind,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
        ),
        backup=backup,
    )


@mcp.tool
def get_lists(file_path: str) -> list:
    """Numbered and bulleted list paragraphs grouped by numbering instance,
    with each item's nesting level (0-based), numbering style (bullet or
    decimal/lowerLetter/lowerRoman/...), and text. Useful for verifying
    list restarts and nesting after insert_document or add_list. Read-only."""
    return _ls.get_lists(DocxPackage(file_path))


@mcp.tool
def change_case(
    file_path: str,
    transform: str,
    indices: list[int] | None = None,
    find: str | None = None,
    backup: bool = True,
) -> dict:
    """Change text case: transform is upper | lower | title | sentence.
    Target: paragraph indices, or every occurrence of `find`
    (case-insensitive). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tx.change_case(pkg, transform, indices=indices, find=find),
        backup=backup,
    )


# --------------------------------------------------------------------- tables


@mcp.tool
def list_tables(file_path: str) -> list:
    """List tables with dimensions, merge flags, and header previews; the
    position is the 0-based table_index the other table tools take.
    Read-only.
    """
    return _rd.list_tables(DocxPackage(file_path))


@mcp.tool
def get_table(file_path: str, table_index: int) -> dict:
    """Read one body-level table in full: every cell's text, the merge map
    (gridSpan for horizontal, vMerge for vertical), column widths, and
    the table style. table_index is 0-based in list_tables order. For
    tables nested inside a cell, use get_nested_table instead. Use
    set_cells/set_cells_block to write cell values. Read-only."""
    return _rd.get_table(DocxPackage(file_path), table_index)


@mcp.tool
def create_table(
    file_path: str,
    data: list[list[str]],
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    header_row: bool = True,
    width_pt: float | None = None,
    backup: bool = True,
) -> dict:
    """Create a table from 2D string data with single-line borders. To build a
    table from a CSV/JSON file, use import_table. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.create_table(
            pkg,
            data,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
            header_row=header_row,
            width_pt=width_pt,
        ),
        backup=backup,
    )


@mcp.tool
def delete_table(file_path: str, table_index: int, backup: bool = True) -> dict:
    """Delete a whole table (table_index from list_tables). Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _tb.delete_table(pkg, table_index), backup=backup
    )


@mcp.tool
def insert_rows(
    file_path: str,
    table_index: int,
    at: int,
    count: int = 1,
    copy_format_from: int | None = None,
    backup: bool = True,
) -> dict:
    """Insert rows before row `at` (at = row count appends), copying structure
    and formatting from an existing row (copy_format_from). Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.insert_rows(
            pkg, table_index, at=at, count=count, copy_format_from=copy_format_from
        ),
        backup=backup,
    )


@mcp.tool
def delete_rows(
    file_path: str,
    table_index: int,
    start: int,
    end: int | None = None,
    backup: bool = True,
) -> dict:
    """Delete rows start..end (inclusive). Vertical merges are re-rooted, not
    broken. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.delete_rows(pkg, table_index, start, end),
        backup=backup,
    )


@mcp.tool
def insert_columns(
    file_path: str,
    table_index: int,
    at: int,
    count: int = 1,
    width_pt: float | None = None,
    backup: bool = True,
) -> dict:
    """Insert grid columns before position `at` (at = column count appends).
    Merge-aware: inserting inside a merged cell widens it. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.insert_columns(
            pkg, table_index, at=at, count=count, width_pt=width_pt
        ),
        backup=backup,
    )


@mcp.tool
def delete_columns(
    file_path: str, table_index: int, columns: list[int], backup: bool = True
) -> dict:
    """Delete grid columns (0-based list). Merge-aware: merged cells shrink,
    single cells are removed, the grid stays consistent. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.delete_columns(pkg, table_index, columns),
        backup=backup,
    )


@mcp.tool
def set_cells(
    file_path: str,
    table_index: int,
    edits: list[dict],
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Write many table cells in ONE call: edits = [{row, cell, text}]. Use
    this for any multi-cell edit instead of per-cell calls; set_cells_block
    writes a contiguous 2D block instead. track records tracked changes by
    `author`; the result reports cells_written. Live mode refuses
    vertically merged tables (the file-based path is merge-aware).
    Auto-backup in file mode: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Documents
    open in Word are edited live.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tb.set_cells(
                pkg, table_index, edits, track=track, author=author
            ),
            backup=backup,
        ),
        lambda: _lo.set_cells(
            file_path, table_index, edits, track=track, author=author
        ),
    )


@mcp.tool
def set_cells_block(
    file_path: str,
    table_index: int,
    origin_row: int,
    origin_cell: int,
    data: list[list[str]],
    backup: bool = True,
) -> dict:
    """Write a 2D block of values starting at (origin_row, origin_cell); for
    scattered single-cell edits use set_cells. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.set_cells_grid(
            pkg, table_index, origin_row=origin_row, origin_cell=origin_cell, data=data
        ),
        backup=backup,
    )


@mcp.tool
def format_cells(
    file_path: str,
    table_index: int,
    targets: list[dict],
    formatting: dict,
    backup: bool = True,
) -> dict:
    """Format table cells in bulk. targets: [{row, cell?}] ({row} alone =
    whole row); formatting: shading, bold, italic, alignment, valign,
    padding_pt. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.format_cells(pkg, table_index, targets, formatting),
        backup=backup,
    )


@mcp.tool
def merge_cells(
    file_path: str,
    table_index: int,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    backup: bool = True,
) -> dict:
    """Merge a rectangle of cells (grid coordinates, inclusive). Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.merge_cells(
            pkg,
            table_index,
            start_row=start_row,
            end_row=end_row,
            start_col=start_col,
            end_col=end_col,
        ),
        backup=backup,
    )


@mcp.tool
def unmerge_cells(
    file_path: str, table_index: int, row: int, cell: int, backup: bool = True
) -> dict:
    """Split a merged cell back into single cells (horizontal and vertical).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.unmerge_cells(pkg, table_index, row=row, cell=cell),
        backup=backup,
    )


@mcp.tool
def set_column_widths(
    file_path: str, table_index: int, widths_pt: list[float], backup: bool = True
) -> dict:
    """Set every grid column width in points (one value per grid column).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.set_column_widths(pkg, table_index, widths_pt),
        backup=backup,
    )


@mcp.tool
def apply_table_style(
    file_path: str,
    table_index: int,
    style: str = "TableGrid",
    banded_rows: bool = True,
    first_row_header: bool = True,
    backup: bool = True,
) -> dict:
    """Apply a named table style: TableGrid (bordered) | PlainTable |
    BandedTable (alternating shading); definitions are injected when
    missing. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.apply_table_style(
            pkg, table_index, style, banded_rows=banded_rows,
            first_row_header=first_row_header,
        ),
        backup=backup,
    )


@mcp.tool
def sort_table(
    file_path: str,
    table_index: int,
    column: int,
    numeric: bool = False,
    descending: bool = False,
    has_header: bool = True,
    backup: bool = True,
) -> dict:
    """Sort data rows by a column (CELL index, 0-based). Refused when vertical
    merges make rows interdependent. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.sort_table(
            pkg, table_index, column=column, numeric=numeric,
            descending=descending, has_header=has_header,
        ),
        backup=backup,
    )


@mcp.tool
def split_table(
    file_path: str, table_index: int, at_row: int, backup: bool = True
) -> dict:
    """Split a table into two at a row (that row starts the new table).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.split_table(pkg, table_index, at_row=at_row),
        backup=backup,
    )


@mcp.tool
def set_header_row_repeat(
    file_path: str,
    table_index: int,
    rows: int = 1,
    on: bool = True,
    backup: bool = True,
) -> dict:
    """Repeat the first N rows as the table header on every page. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.set_header_row_repeat(pkg, table_index, rows=rows, on=on),
        backup=backup,
    )


@mcp.tool
def get_nested_table(
    file_path: str,
    table_index: int,
    row: int,
    cell: int,
    nested_index: int = 0,
) -> dict:
    """Read a table nested inside a cell of a body-level table: rows, cells,
    text, and merge info, exactly like get_table but addressed by the host
    table_index, host row/cell, and nested_index (0 when the cell holds one
    nested table; higher for multiple). Use set_nested_cells to write.
    Read-only."""
    return _tb.get_nested_table(
        DocxPackage(file_path), table_index, row=row, cell=cell,
        nested_index=nested_index,
    )


@mcp.tool
def set_nested_cells(
    file_path: str,
    table_index: int,
    row: int,
    cell: int,
    edits: list[dict],
    nested_index: int = 0,
    backup: bool = True,
) -> dict:
    """Write cells inside a NESTED table (edits = [{row, cell, text}]; the
    host cell is addressed by table_index/row/cell, nested_index picks
    among several). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.set_nested_cells(
            pkg, table_index, row=row, cell=cell, edits=edits,
            nested_index=nested_index,
        ),
        backup=backup,
    )


# ---------------------------------------------------------- footnotes/endnotes


@mcp.tool
def list_footnotes(file_path: str) -> list:
    """Every footnote in the document: id, the body paragraph where the
    reference appears (display position), and the note's full text. Use
    add_footnote/edit_footnote/delete_footnote to modify; validate_notes to
    check integrity. Read-only."""
    return _rd.list_footnotes(DocxPackage(file_path))


@mcp.tool
def list_endnotes(file_path: str) -> list:
    """Every endnote in the document: id, the body paragraph where the
    reference appears (display position), and the note's full text. Use
    add_endnote/edit_endnote/delete_endnote to modify; validate_notes to
    check integrity. Read-only."""
    return _rd.list_endnotes(DocxPackage(file_path))


@mcp.tool
def add_footnote(
    file_path: str,
    anchor_text: str,
    note_text: str,
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Add a footnote anchored after the occurrence-th match of anchor_text;
    creates all note infrastructure (parts, styles, superscript) if the
    document has none. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.add_note(
            pkg,
            "footnote",
            anchor_text=anchor_text,
            note_text=note_text,
            occurrence=occurrence,
        ),
        backup=backup,
    )


@mcp.tool
def add_endnote(
    file_path: str,
    anchor_text: str,
    note_text: str,
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Add an endnote anchored after anchor_text (same mechanics as
    add_footnote). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.add_note(
            pkg,
            "endnote",
            anchor_text=anchor_text,
            note_text=note_text,
            occurrence=occurrence,
        ),
        backup=backup,
    )


@mcp.tool
def edit_footnote(
    file_path: str,
    new_text: str,
    note_id: str | None = None,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Rewrite a footnote's text, addressed by note_id or 1-based display
    position (exactly one). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.edit_note(
            pkg, "footnote", note_id=note_id, position=position, new_text=new_text
        ),
        backup=backup,
    )


@mcp.tool
def edit_endnote(
    file_path: str,
    new_text: str,
    note_id: str | None = None,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Rewrite an endnote's text, addressed by note_id or 1-based display
    position. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.edit_note(
            pkg, "endnote", note_id=note_id, position=position, new_text=new_text
        ),
        backup=backup,
    )


@mcp.tool
def delete_footnote(
    file_path: str,
    note_id: str | None = None,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Delete a footnote: the definition AND the body reference mark, always
    both. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.delete_note(
            pkg, "footnote", note_id=note_id, position=position
        ),
        backup=backup,
    )


@mcp.tool
def delete_endnote(
    file_path: str,
    note_id: str | None = None,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Delete an endnote: the definition AND the body reference mark.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.delete_note(
            pkg, "endnote", note_id=note_id, position=position
        ),
        backup=backup,
    )


@mcp.tool
def validate_notes(file_path: str) -> dict:
    """Check footnote/endnote structural integrity: whether note
    definitions match their body references and whether orphan definitions
    exist. ok=true means no corruption; needs_cleanup=true means orphan
    note definitions were found (run cleanup_orphan_notes to purge them).
    Use list_footnotes/list_endnotes to see note contents. Read-only."""
    return _nt.validate_notes(DocxPackage(file_path))


@mcp.tool
def cleanup_orphan_notes(file_path: str, backup: bool = True) -> dict:
    """Remove footnote/endnote definitions no body reference points to.
    Content-deleting tools already do this automatically; use this for
    documents that arrived with orphans. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(file_path, lambda pkg: _nt.purge_orphans(pkg), backup=backup)


# --------------------------------------------------------- references & links


@mcp.tool
def add_bookmark(
    file_path: str, name: str, anchor_text: str, occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Bookmark a text range (the target add_cross_reference points at).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.add_bookmark(
            pkg, name, anchor_text=anchor_text, occurrence=occurrence
        ),
        backup=backup,
    )


@mcp.tool
def list_bookmarks(file_path: str) -> list:
    """Every user-visible bookmark: name, the paragraph index where it
    starts, and the bookmarked text. Bookmarks serve as targets for
    add_cross_reference and as stable anchors for navigation. Internal
    bookmarks (TOC, field-generated) are excluded. Read-only."""
    return _fl.list_bookmarks(DocxPackage(file_path))


@mcp.tool
def add_cross_reference(
    file_path: str,
    after_anchor: str,
    to_bookmark: str,
    kind: str = "page",
    backup: bool = True,
) -> dict:
    """Insert a cross-reference field after anchor text. kind: 'page' (page
    number) or 'text' (the bookmarked text); Word computes the value on
    field update. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.add_cross_reference(
            pkg, after_anchor=after_anchor, to_bookmark=to_bookmark, kind=kind
        ),
        backup=backup,
    )


@mcp.tool
def add_caption(
    file_path: str,
    text: str,
    table_index: int | None = None,
    after_anchor: str | None = None,
    label: str = "Table",
    above: bool = True,
    backup: bool = True,
) -> dict:
    """Insert a numbered caption (SEQ field), 'Table N: text', above/below a
    table or at an anchor. label: Table | Figure | Equation. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.add_caption(
            pkg,
            table_index=table_index,
            after_anchor=after_anchor,
            label=label,
            text=text,
            above=above,
        ),
        backup=backup,
    )


@mcp.tool
def add_hyperlink(
    file_path: str, anchor_text: str, url: str, occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Turn existing text (occurrence-th match of anchor_text) into an
    external hyperlink. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.add_hyperlink(
            pkg, anchor_text=anchor_text, url=url, occurrence=occurrence
        ),
        backup=backup,
    )


# ------------------------------------------------------------------------ TOC


@mcp.tool
def insert_toc(
    file_path: str,
    levels: str = "1-3",
    title: str | None = "Table of Contents",
    after_index: int | None = None,
    at_start: bool = False,
    update_on_open: bool = True,
    backup: bool = True,
) -> dict:
    """Insert a Table of Contents (SDT-wrapped TOC field); levels like '1-3'
    picks heading depth. Page numbers appear after Word updates fields:
    automatically on next open (update_on_open) or immediately via
    com_refresh_fields. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tc.insert_toc(
            pkg,
            levels=levels,
            title=title,
            after_index=after_index,
            at_start=at_start,
            update_on_open=update_on_open,
        ),
        backup=backup,
    )


@mcp.tool
def read_toc(file_path: str) -> dict:
    """Read every TOC-family field (main TOC, List of Tables, List of
    Figures) and its cached entries: title, page number, heading level.
    The cached text reflects the state at last field update; refresh with
    com_refresh_fields (requires Word) then re-read. Use insert_toc/
    insert_caption_list/delete_toc to manage TOCs. Read-only."""
    return _tc.read_toc(DocxPackage(file_path))


@mcp.tool
def delete_toc(file_path: str, which: int = 0, backup: bool = True) -> dict:
    """Delete one TOC-family field by index (read_toc's tocs list order).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _tc.delete_toc(pkg, which=which), backup=backup
    )


@mcp.tool
def set_update_fields_flag(
    file_path: str, on: bool = True, backup: bool = True
) -> dict:
    """Toggle 'update all fields on next open' (one-shot; Word clears it after
    updating). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _tc.set_update_fields_flag(pkg, on), backup=backup
    )


@mcp.tool
def insert_caption_list(
    file_path: str,
    label: str = "Table",
    title: str | None = None,
    after_index: int | None = None,
    at_start: bool = False,
    update_on_open: bool = True,
    backup: bool = True,
) -> dict:
    """Insert a List of Tables/Figures/Equations built from add_caption SEQ
    entries. label: Table | Figure | Equation; default title 'List of Xs'.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tc.insert_caption_list(
            pkg,
            label=label,
            title=title,
            after_index=after_index,
            at_start=at_start,
            update_on_open=update_on_open,
        ),
        backup=backup,
    )


@mcp.tool
def convert_notes(
    file_path: str,
    direction: str,
    note_id: str | None = None,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Convert footnotes to endnotes or back. direction: footnotes_to_endnotes
    | endnotes_to_footnotes; give note_id or position for ONE note, neither
    for ALL. Word renumbers automatically. Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.convert_notes(
            pkg, direction, note_id=note_id, position=position
        ),
        backup=backup,
    )


# ------------------------------------------------------- bibliography (native)


@mcp.tool
def add_source(
    file_path: str,
    tag: str,
    source_type: str,
    title: str,
    year: str | None = None,
    authors: list | None = None,
    editors: list | None = None,
    journal_name: str | None = None,
    book_title: str | None = None,
    publisher: str | None = None,
    city: str | None = None,
    pages: str | None = None,
    volume: str | None = None,
    issue: str | None = None,
    edition: str | None = None,
    institution: str | None = None,
    url: str | None = None,
    internet_site_title: str | None = None,
    style: str = "APA",
    extra_fields: dict | None = None,
    backup: bool = True,
) -> dict:
    """Add a Word-native bibliography source to the document store. tag =
    unique citation key; source_type: JournalArticle | Book | BookSection |
    Report | InternetSite | ...; authors/editors: [{last, first, middle?}]
    or [{corporate}]. Cite with insert_citation; render the list with
    insert_bibliography. Run detect_citation_system first on unfamiliar
    documents: mixing Word-native and Zotero/Mendeley/EndNote citations
    creates a bibliography no single manager maintains. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _bib.add_source(
            pkg, tag=tag, source_type=source_type, title=title, year=year,
            authors=authors, editors=editors, journal_name=journal_name,
            book_title=book_title, publisher=publisher, city=city,
            pages=pages, volume=volume, issue=issue, edition=edition,
            institution=institution, url=url,
            internet_site_title=internet_site_title, style=style,
            extra_fields=extra_fields,
        ),
        backup=backup,
    )


@mcp.tool
def list_sources(file_path: str) -> list:
    """Bibliography sources stored in the document's XML source store: tag,
    type (JournalArticle, Book, ...), author, title, year, and all other
    fields. These feed insert_citation and insert_bibliography when using
    Word's native citation system (not Zotero/Mendeley/EndNote). Run
    detect_citation_system first on unfamiliar documents. Read-only."""
    return _bib.list_sources(DocxPackage(file_path))


@mcp.tool
def delete_source(file_path: str, tag: str, force: bool = False, backup: bool = True) -> dict:
    """Delete a bibliography source by tag (refused while cited unless
    force=True). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _bib.delete_source(pkg, tag, force=force),
        backup=backup,
    )


@mcp.tool
def set_bibliography_style(file_path: str, style: str, backup: bool = True) -> dict:
    """Set the style Word-native CITATION and BIBLIOGRAPHY fields render in:
    APA | Chicago | MLA | IEEE | Turabian | Harvard - Anglia | GB7714 |
    GOST - Name Sort | GOST - Title Sort | ISO 690 - First Element and Date
    | ISO 690 - Numerical Reference | SIST02. Affects Word-native fields
    only: manager citations restyle in their own manager, plain-text
    citations take convert_citation_style. Run detect_citation_system first
    on unfamiliar documents: mixing Word-native and Zotero/Mendeley/EndNote
    citations creates a bibliography no single manager maintains.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _bib.set_bibliography_style(pkg, style),
        backup=backup,
    )


@mcp.tool
def insert_citation(
    file_path: str,
    tag: str,
    anchor_text: str,
    occurrence: int = 1,
    pages: str | None = None,
    suppress_author: bool = False,
    suppress_year: bool = False,
    prefix: str | None = None,
    suffix: str | None = None,
    backup: bool = True,
) -> dict:
    """Insert a CITATION field for a stored source (tag from add_source /
    list_sources) after the occurrence-th anchor_text match; renders in the
    document citation style on field update (placeholder until then).
    pages, prefix/suffix, and the suppress flags shape the rendered cite.
    Run detect_citation_system first on unfamiliar documents: mixing
    Word-native and Zotero/Mendeley/EndNote citations creates a
    bibliography no single manager maintains. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _bib.insert_citation(
            pkg, tag=tag, anchor_text=anchor_text, occurrence=occurrence,
            pages=pages, suppress_author=suppress_author,
            suppress_year=suppress_year, prefix=prefix, suffix=suffix,
        ),
        backup=backup,
    )


@mcp.tool
def insert_bibliography(
    file_path: str,
    title: str | None = "Bibliography",
    after_index: int | None = None,
    at_end: bool = True,
    update_on_open: bool = True,
    backup: bool = True,
) -> dict:
    """Insert the BIBLIOGRAPHY field: Word generates the full styled reference
    list from the source store on field update. Run detect_citation_system
    first on unfamiliar documents: mixing Word-native and
    Zotero/Mendeley/EndNote citations creates a bibliography no single
    manager maintains. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _bib.insert_bibliography(
            pkg, title=title, after_index=after_index, at_end=at_end,
            update_on_open=update_on_open,
        ),
        backup=backup,
    )


# -------------------------------------------------------------------- indexing


@mcp.tool
def mark_index_entry(
    file_path: str,
    anchor_text: str,
    entry: str,
    subentry: str | None = None,
    occurrence: int = 1,
    bold_page: bool = False,
    italic_page: bool = False,
    see: str | None = None,
    backup: bool = True,
) -> dict:
    """Mark a location for the index (invisible XE field at the anchor).
    see='other entry' writes a cross-reference instead of a page number.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.mark_index_entry(
            pkg, anchor_text=anchor_text, entry=entry, subentry=subentry,
            occurrence=occurrence, bold_page=bold_page,
            italic_page=italic_page, see=see,
        ),
        backup=backup,
    )


@mcp.tool
def list_index_entries(file_path: str) -> list:
    """Every XE index entry marked in the document: main text, sub-entry,
    see/see-also references, and the paragraph where each appears. Use
    mark_index_entry to add entries and insert_index to generate the
    compiled index. Read-only."""
    return _fl.list_index_entries(DocxPackage(file_path))


@mcp.tool
def insert_index(
    file_path: str,
    title: str | None = "Index",
    columns: int = 2,
    letter_headings: bool = True,
    after_index: int | None = None,
    at_end: bool = True,
    update_on_open: bool = True,
    backup: bool = True,
) -> dict:
    """Insert the INDEX field compiling all XE entries with page numbers
    (generated when Word updates fields). Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.insert_index(
            pkg, title=title, columns=columns,
            letter_headings=letter_headings, after_index=after_index,
            at_end=at_end, update_on_open=update_on_open,
        ),
        backup=backup,
    )


# ------------------------------------------------------ structure & templates


@mcp.tool
def move_section(
    file_path: str,
    heading_text: str,
    before_heading: str | None = None,
    after_heading: str | None = None,
    at_end: bool = False,
    backup: bool = True,
) -> dict:
    """Move a heading and its ENTIRE section (content and tables until the
    next same-or-higher heading) before/after another heading or to the
    end. Sections are addressed by exact heading text; only headings are
    matched, so body prose repeating the words does not collide. Refuses
    moves that would cut fields or section breaks. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.move_section(
            pkg, heading_text, before_heading=before_heading,
            after_heading=after_heading, at_end=at_end,
        ),
        backup=backup,
    )


@mcp.tool
def list_section_blocks(file_path: str) -> list:
    """Headings with each section's body-element count, paragraph count,
    and table count: the planning aid for move_section (shows what will
    travel with each heading) and for insert_document/copy_table (shows
    body-item indices at section boundaries). Read-only."""
    return _sx.list_section_blocks(DocxPackage(file_path))


@mcp.tool
def apply_template(
    file_path: str,
    reference_path: str,
    include_page_geometry: bool = True,
    backup: bool = True,
) -> dict:
    """Restyle this document to match a reference document: styles, theme,
    fonts, layout settings, optionally page geometry. This is also the
    import-styles-from-another-document tool: style definitions are
    remapped by name, unmatched styles preserved, content untouched.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tp.apply_template(
            pkg, reference_path, include_page_geometry=include_page_geometry
        ),
        backup=backup,
    )


@mcp.tool
def set_document_properties(
    file_path: str,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    comments: str | None = None,
    backup: bool = True,
) -> dict:
    """Set core document metadata (File > Info): title, author, subject,
    keywords, category, comments; only the parameters given change.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.set_document_properties(
            pkg, title=title, author=author, subject=subject,
            keywords=keywords, category=category, comments=comments,
        ),
        backup=backup,
    )


@mcp.tool
def define_style(
    file_path: str,
    style_id: str,
    name: str,
    style_type: str = "paragraph",
    based_on: str | None = "Normal",
    next_style: str | None = None,
    character_formatting: dict | None = None,
    paragraph_formatting: dict | None = None,
    backup: bool = True,
) -> dict:
    """Create or replace a custom style (paragraph or character) with full
    formatting control; get_styles returns existing definitions in this
    exact input shape, so cloning a style is one read + one define.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.define_style(
            pkg, style_id=style_id, name=name, style_type=style_type,
            based_on=based_on, next_style=next_style,
            character_formatting=character_formatting,
            paragraph_formatting=paragraph_formatting,
        ),
        backup=backup,
    )


@mcp.tool
def apply_character_style(
    file_path: str,
    find: str,
    style: str,
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Apply a named character style to a text range (occurrence-th match of
    find). For direct formatting without a style, use format_text.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.apply_character_style(
            pkg, find=find, style=style, occurrence=occurrence
        ),
        backup=backup,
    )


@mcp.tool
def set_image_alt_text(
    file_path: str,
    image_index: int,
    description: str,
    title: str | None = None,
    backup: bool = True,
) -> dict:
    """Set accessibility alt text for an image (image_index from list_images).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.set_image_alt_text(
            pkg, image_index, description=description, title=title
        ),
        backup=backup,
    )


# ---------------------------------------------------- watermark & protection


@mcp.tool
def add_watermark(
    file_path: str,
    text: str = "DRAFT",
    color: str = "silver",
    opacity: float = 0.5,
    diagonal: bool = True,
    backup: bool = True,
) -> dict:
    """Add a text watermark behind the text on every page (all header parts;
    compatible with Word's own Remove Watermark command). Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.add_watermark(
            pkg, text, color=color, opacity=opacity, diagonal=diagonal
        ),
        backup=backup,
    )


@mcp.tool
def remove_watermark(file_path: str, backup: bool = True) -> dict:
    """Remove watermark shapes from every header. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(file_path, lambda pkg: _fu.remove_watermark(pkg), backup=backup)


@mcp.tool
def set_document_protection(
    file_path: str,
    edit: str = "trackedChanges",
    password: str | None = None,
    restrict_formatting: bool = False,
    backup: bool = True,
) -> dict:
    """Restrict editing: edit = readOnly | comments | trackedChanges | forms.
    trackedChanges forces every edit by the recipient to be tracked (the
    send-to-committee mode). Password hashing is Word-compatible SHA-512;
    this is NOT encryption (com_save_with_password encrypts). Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _pr.set_document_protection(
            pkg, edit=edit, password=password,
            restrict_formatting=restrict_formatting,
        ),
        backup=backup,
    )


@mcp.tool
def remove_document_protection(file_path: str, backup: bool = True) -> dict:
    """Lift the editing restriction set by set_document_protection.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _pr.remove_document_protection(pkg),
        backup=backup,
    )


@mcp.tool
def get_protection(file_path: str) -> dict:
    """Current document protection state: mode (none, readOnly, comments,
    trackedChanges, forms), whether a password is set, and whether
    formatting restrictions are active. Use set_document_protection to
    enable and remove_document_protection to disable. Read-only."""
    return _pr.get_protection(DocxPackage(file_path))


# ---------------------------------------------------------- headers & layout


@mcp.tool
def get_headers_footers(file_path: str) -> dict:
    """Every header and footer part across all sections: text content,
    whether it contains a PAGE-number field, and the part type (default,
    first page, even page). Use set_header/set_footer to write them and
    setup_chapter_headers for per-chapter running headers. Read-only."""
    return _fu.get_headers_footers(DocxPackage(file_path))


@mcp.tool
def set_header(
    file_path: str,
    text: str,
    section: int = 0,
    ref_type: str = "default",
    alignment: str = "center",
    backup: bool = True,
) -> dict:
    """Set a section's header text. ref_type: default | first | even.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.set_header_footer(
            pkg, "header", text, section=section, ref_type=ref_type,
            alignment=alignment,
        ),
        backup=backup,
    )


@mcp.tool
def set_footer(
    file_path: str,
    text: str,
    section: int = 0,
    ref_type: str = "default",
    alignment: str = "center",
    include_page_number: bool = False,
    backup: bool = True,
) -> dict:
    """Set a section's footer text, optionally with a page-number field.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.set_header_footer(
            pkg, "footer", text, section=section, ref_type=ref_type,
            alignment=alignment, include_page_number=include_page_number,
        ),
        backup=backup,
    )


@mcp.tool
def add_page_numbers(
    file_path: str,
    section: int = 0,
    position: str = "footer",
    alignment: str = "center",
    prefix: str = "",
    start_at: int | None = None,
    x_of_y: bool = False,
    backup: bool = True,
) -> dict:
    """Add page numbers (PAGE field) to the header or footer; x_of_y renders
    'Page N of M'. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.add_page_numbers(
            pkg, section=section, position=position, alignment=alignment,
            prefix=prefix, start_at=start_at, x_of_y=x_of_y,
        ),
        backup=backup,
    )


@mcp.tool
def set_columns(
    file_path: str,
    section: int = 0,
    count: int = 1,
    space_pt: float = 36,
    separator: bool = False,
    widths_pt: list[float] | None = None,
    backup: bool = True,
) -> dict:
    """Set multi-column text layout for a section (equal widths, or explicit
    widths_pt per column; separator draws a line between columns).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.set_columns(
            pkg, section=section, count=count, space_pt=space_pt,
            separator=separator, widths_pt=widths_pt,
        ),
        backup=backup,
    )


@mcp.tool
def set_page_number_format(
    file_path: str,
    section: int = 0,
    number_format: str = "decimal",
    start_at: int | None = None,
    backup: bool = True,
) -> dict:
    """Set the page-number FORMAT per section: lowerRoman for front matter,
    decimal restarting at 1 for the body, upperRoman, letters, etc.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.set_page_number_format(
            pkg, section=section, number_format=number_format, start_at=start_at
        ),
        backup=backup,
    )


@mcp.tool
def set_line_numbering(
    file_path: str,
    section: int = 0,
    count_by: int = 1,
    start: int = 1,
    restart: str = "continuous",
    distance_pt: float | None = None,
    remove: bool = False,
    backup: bool = True,
) -> dict:
    """Set manuscript line numbering for a section (journal submissions).
    restart: continuous | newPage | newSection; remove=True clears it.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.set_line_numbering(
            pkg, section=section, count_by=count_by, start=start,
            restart=restart, distance_pt=distance_pt, remove=remove,
        ),
        backup=backup,
    )


@mcp.tool
def list_sections(file_path: str) -> list:
    """Every section: page width/height, orientation (portrait/landscape),
    all margins, and header/footer part references. Use
    set_section_properties to change page geometry and add_section_break to
    create new sections. Read-only."""
    return _fu.list_sections(DocxPackage(file_path))


@mcp.tool
def set_section_properties(
    file_path: str,
    section: int = 0,
    orientation: str | None = None,
    page_width_pt: float | None = None,
    page_height_pt: float | None = None,
    margins_pt: dict | None = None,
    backup: bool = True,
) -> dict:
    """Set page size, orientation, and margins for one section. Every response
    carries the section's full state; call with no change parameters to
    read the current values. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.set_section_properties(
            pkg, section=section, orientation=orientation,
            page_width_pt=page_width_pt, page_height_pt=page_height_pt,
            margins_pt=margins_pt,
        ),
        backup=backup,
    )


@mcp.tool
def add_section_break(
    file_path: str, after_index: int, break_type: str = "nextPage",
    backup: bool = True,
) -> dict:
    """Insert a section break after a body paragraph: nextPage | continuous |
    evenPage | oddPage. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fu.add_section_break(
            pkg, after_index=after_index, break_type=break_type
        ),
        backup=backup,
    )


# --------------------------------------------------------------------- images


@mcp.tool
def add_image(
    file_path: str,
    image_path: str,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    width_pt: float | None = None,
    alignment: str = "center",
    backup: bool = True,
) -> dict:
    """Insert an inline image (PNG/JPEG/GIF/BMP/TIFF), aspect ratio kept.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _md.add_image(
            pkg, image_path, after_index=after_index, after_anchor=after_anchor,
            at_end=at_end, width_pt=width_pt, alignment=alignment,
        ),
        backup=backup,
    )


@mcp.tool
def list_images(file_path: str) -> list:
    """Every inline image: index (use with resize_image, replace_image,
    set_image_alt_text), display size in points, native pixel dimensions,
    effective DPI, alt text, and the media part path. For print-quality
    checks, use check_image_resolution instead. Read-only."""
    return _md.list_images(DocxPackage(file_path))


@mcp.tool
def resize_image(
    file_path: str, image_index: int, width_pt: float, backup: bool = True
) -> dict:
    """Resize an inline image to width_pt, aspect ratio kept (image_index from
    list_images). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _md.resize_image(pkg, image_index, width_pt=width_pt),
        backup=backup,
    )


@mcp.tool
def replace_image(
    file_path: str, image_index: int, new_image_path: str, backup: bool = True
) -> dict:
    """Swap an image's file, keeping placement and display size. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _md.replace_image(pkg, image_index, new_image_path),
        backup=backup,
    )


@mcp.tool
def add_chart(
    file_path: str,
    chart_type: str,
    data: list | dict | str,
    title: str | None = None,
    width_pt: float | None = None,
    height_pt: float | None = None,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    alignment: str = "center",
    legend: bool = True,
    colors: list | None = None,
    backup: bool = True,
) -> dict:
    """Insert a native, theme-following Word chart built from data (no image
    rendering). chart_type: 'bar' (horizontal), 'column' (vertical),
    'line', 'pie', or 'scatter'. data shapes: rows like import_table ([["",
    "S1", "S2"], ["Alpha", 4.3, 1.2], ...]; first row series names, first
    column categories), a dict {"categories": [...], "series": [{"name",
    "values"}]} (scatter: {"series": [{"name", "x", "y"}]}), or a
    .csv/.json file path. The chart part carries literal data caches AND a
    matching embedded workbook, so it renders everywhere and right-click >
    Edit Data works in Word. Size defaults to 6.0 x 3.5 in
    (width_pt/height_pt override); position with at most one of
    after_index/after_anchor/at_end (default document end); colors takes
    one hex per series (default: theme accents). Refused: other chart types
    (3D/area/doughnut/radar/...), ragged or non-numeric data, multi-series
    pie. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _charts.add_chart(
            pkg,
            chart_type,
            data,
            title=title,
            width_pt=width_pt,
            height_pt=height_pt,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
            alignment=alignment,
            legend=legend,
            colors=colors,
        ),
        backup=backup,
    )


@mcp.tool
def update_chart_data(
    file_path: str,
    chart_index: int,
    data: list | dict | str,
    series_names: list | None = None,
    backup: bool = True,
) -> dict:
    """Replace the data of an existing bar/column/line/pie/scatter chart in
    place (chart_index from list_charts): literal caches, range formulas,
    and the embedded workbook are rewritten together, so rendering and Edit
    Data stay in sync; formatting and styles are preserved. data takes the
    same shapes as add_chart and must keep the existing series COUNT (point
    count may change); series_names renames. Refused with the reason named:
    chartex/modern, combo, 3D/area/doughnut/radar/surface/stock/bubble,
    multi-level categories, series-count changes, ragged data. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _charts.update_chart_data(
            pkg,
            chart_index,
            data,
            series_names=series_names,
        ),
        backup=backup,
    )


@mcp.tool
def list_charts(file_path: str) -> list:
    """Enumerate every chart in the document body, in document order. Each
    entry reports index (use it with update_chart_data), chart part name,
    plot type (barChart/lineChart/pieChart/scatterChart/... or 'chartex'
    for modern charts like treemap/sunburst/waterfall), title, series
    names, point count, size in points, whether an embedded data workbook
    exists, and supported_for_update with a reason when updating is
    refused (chartex, combo, 3D/area/radar/..., multi-level categories).
    Read-only against the document."""
    return _charts.list_charts(DocxPackage(file_path))


# ------------------------------------------------------------------- comments


@mcp.tool
def get_comments(
    file_path: str, author: str | None = None, live: str = "auto"
) -> list:
    """Comments with authors, anchored text, threading, resolved state.
    Documents open in Word are read live (same entry shape; live ids are
    the comment's position, not the XML id)."""
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_comments(DocxPackage(file_path), author=author),
        lambda: _lo.get_comments(file_path, author=author),
    )


@mcp.tool
def add_comment(
    file_path: str,
    anchor_text: str,
    text: str,
    author: str = "Claude",
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Add a comment on a text range (threaded-comment infrastructure created
    as needed). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _cm.add_comment(
            pkg, anchor_text=anchor_text, text=text, author=author,
            occurrence=occurrence,
        ),
        backup=backup,
    )


@mcp.tool
def reply_to_comment(
    file_path: str, comment_id: str, text: str, author: str = "Claude",
    backup: bool = True,
) -> dict:
    """Add a threaded reply to an existing comment (comment_id from
    get_comments). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _cm.reply_to_comment(
            pkg, comment_id=comment_id, text=text, author=author
        ),
        backup=backup,
    )


@mcp.tool
def resolve_comment(
    file_path: str, comment_id: str, done: bool = True, backup: bool = True
) -> dict:
    """Mark a comment thread resolved, or reopen it with done=False.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _cm.resolve_comment(pkg, comment_id=comment_id, done=done),
        backup=backup,
    )


@mcp.tool
def delete_comment(file_path: str, comment_id: str, backup: bool = True) -> dict:
    """Delete a comment and its replies, including all body markers.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _cm.delete_comment(pkg, comment_id=comment_id),
        backup=backup,
    )


# ------------------------------------------------------------------ revisions


@mcp.tool
def get_tracked_changes(file_path: str, author: str | None = None) -> list:
    """Every tracked change: type (insertion, deletion, move, format
    change), author, date, the affected text, and paragraph location.
    Filter by author to see one reviewer's edits. Use accept_revisions/
    reject_revisions to resolve, or revision_summary/revision_analytics
    for aggregate views. Read-only."""
    return _rd.get_tracked_changes(DocxPackage(file_path), author=author)


@mcp.tool
def revision_summary(file_path: str) -> dict:
    """Summarize tracked-change counts by author and type. For word-level
    analytics and per-section concentration use revision_analytics.
    Read-only.
    """
    return _rd.revision_summary(DocxPackage(file_path))


@mcp.tool
def accept_revisions(
    file_path: str, author: str | None = None, backup: bool = True
) -> dict:
    """Accept tracked changes: all of them, or one author's only. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _rv.accept_revisions(pkg, author=author),
        backup=backup,
    )


@mcp.tool
def reject_revisions(
    file_path: str, author: str | None = None, backup: bool = True
) -> dict:
    """Reject tracked changes: all of them, or one author's only. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _rv.reject_revisions(pkg, author=author),
        backup=backup,
    )


# ---------------------------------------------------------------- COM (Word)


@mcp.tool
def com_word_status() -> dict:
    """Check whether Word is running and what state it is in: interactive_state
    (ready, busy if a dialog is open, blocked if a long operation is running,
    not_running), and a list of open documents with per-document dirty flag
    and autosave state. Use before live editing to confirm Word is responsive,
    or to discover which files are locked. No file_path needed. Read-only."""
    from .com import bridge, live

    out = bridge.word_status()
    status = live.interactive_status()
    out["interactive_state"] = status["interactive_state"]
    if status["open_documents"]:
        out["open_documents"] = status["open_documents"]
    if status.get("protected_view_documents"):
        out["protected_view_documents"] = status["protected_view_documents"]
    return out


@mcp.tool
def live_insert_at_cursor(
    file_path: str, text: str, newline: bool = False
) -> dict:
    """Insert text at the user's cursor position in the open document (main
    body text only; headers/footers/footnotes not supported). The cursor
    position is read once and never moved; the user's selection is
    untouched. newline=True ends the insertion with a paragraph break.
    Use insert_paragraphs for index-addressed insertion instead. Requires
    the document to be open in Word."""
    from .com import live_ops as _lo

    return _lo.insert_at_cursor(file_path, text, newline=newline)


@mcp.tool
def live_scroll_to(
    file_path: str,
    find: str | None = None,
    paragraph_index: int | None = None,
) -> dict:
    """Scroll the user's Word window to show a location without selecting
    anything or moving their cursor (useful after a live edit to show the
    user what changed). Target by find text (first match) or body
    paragraph_index. The document must be open in Word. Read-only."""
    from .com import live_ops as _lo

    return _lo.scroll_to(file_path, find=find, paragraph_index=paragraph_index)


@mcp.tool
def live_set_track_changes(file_path: str, enabled: bool) -> dict:
    """Turn track changes on or off on the open document. This is a
    persistent state change (the document remembers the setting), unlike
    the track flag on edit tools like search_and_replace which auto-restores.
    Returns the previous state so you can restore it later. The document
    must be open in Word."""
    from .com import live_ops as _lo

    return _lo.set_track_changes(file_path, enabled)


@mcp.tool
def word_live_repair() -> dict:
    """Recovery tool: if a crashed live edit left the user's Word frozen
    (ScreenUpdating off), alerts suppressed, or an undo record open, a fresh
    attach fixes all three. Safe to run anytime; reports what it fixed."""
    from .com import live

    return live.live_repair()


@mcp.tool
def com_refresh_fields(file_path: str) -> dict:
    """Update every field in the document (TOC page numbers, PAGEREF,
    NUMPAGES, SEQ, cross-references) via an invisible Word instance, giving
    correct page numbers and computed values immediately. Use after
    insert_toc/insert_index/insert_caption_list for real page numbers, or
    after any edit that shifts pages. The source file is modified in place.
    Requires Word installed."""
    from .com import bridge

    return bridge.refresh_fields(file_path)


@mcp.tool
def com_export_pdf(file_path: str, pdf_path: str | None = None) -> dict:
    """Export the document to PDF via an invisible Word instance with full
    fidelity (fields resolved, footnotes, TOC, headers/footers, images).
    pdf_path defaults to the source filename with .pdf extension in the
    same directory. The source .docx is never modified. Requires Word
    installed."""
    from .com import bridge

    return bridge.export_pdf(file_path, pdf_path)


@mcp.tool
def com_compare_documents(
    original_path: str,
    revised_path: str,
    output_path: str | None = None,
    author: str = "word-mcp compare",
) -> dict:
    """Word-native compare of two documents (e.g. two DTG versions of a
    draft): produces a NEW file where every difference is a tracked change,
    plus a revision summary. Inputs are untouched. Default output:
    <revised>_COMPARE.docx."""
    from .com import bridge

    return bridge.compare_documents(
        original_path, revised_path, output_path, author=author
    )


@mcp.tool
def com_merge_documents(
    paths: list[str],
    output_path: str,
    section_break_between: bool = True,
) -> dict:
    """Concatenate whole documents in order into ONE new file via Word, full
    fidelity (styles, footnotes, numbering); section breaks between parts
    keep per-chapter headers/numbering possible. To insert a document INTO
    an existing one at a chosen position with style reconciliation, use
    insert_document; for a single table, copy_table.
    """
    from .com import bridge

    return bridge.merge_documents(
        paths, output_path, section_break_between=section_break_between
    )


@mcp.tool
def com_combine_documents(
    original_path: str,
    revised_path: str,
    output_path: str | None = None,
) -> dict:
    """Combine two reviewers' tracked changes into one document (both sets
    of revisions with their original attributions preserved). Use when two
    people edited copies of the same draft and you need a unified redline.
    For diffing two versions to discover what changed, use
    com_compare_documents instead. Output defaults beside the original.
    Requires Word installed."""
    from .com import bridge

    return bridge.combine_documents(original_path, revised_path, output_path)


@mcp.tool
def com_save_open_document(file_path: str) -> dict:
    """Tell the user's running Word to SAVE a document it has open, so
    file-based tools can read the current state. Use before diagnose_document
    or any file-mode tool that needs the latest edits flushed to disk.
    Does nothing if Word is not running or the file is not open. Requires
    Word installed."""
    from .com import bridge

    return bridge.save_open_document(file_path)


@mcp.tool
def com_close_open_document(file_path: str, save: bool = True) -> dict:
    """Tell the user's running Word to CLOSE an open document, releasing
    the file lock so file-based tools can edit it. Saves by default; pass
    save=False to discard unsaved changes. Use when switching from live
    editing to file-mode tools. Requires Word installed."""
    from .com import bridge

    return bridge.close_open_document(file_path, save=save)


@mcp.tool
def com_proofing_errors(file_path: str, limit: int = 100) -> dict:
    """Word's own spelling and grammar error lists with surrounding
    context: each error, its type (spelling/grammar), the sentence
    containing it, and suggested corrections. Useful as a review aid
    before submission (note: proper nouns and technical terms appear as
    spelling errors). Requires Word installed; opens an invisible instance.
    Read-only."""
    from .com import bridge

    return bridge.proofing_errors(file_path, limit=limit)


@mcp.tool
def com_readability_statistics(file_path: str) -> dict:
    """Word's own readability statistics via COM: Flesch Reading Ease,
    Flesch-Kincaid Grade Level, word/sentence/paragraph counts, and
    averages. Requires Word installed; opens an invisible instance to
    compute the statistics. Read-only."""
    from .com import bridge

    return bridge.readability_statistics(file_path)


@mcp.tool
def com_save_with_password(
    file_path: str, password: str, output_path: str | None = None
) -> dict:
    """Save an ENCRYPTED copy requiring a password to open (real AES
    encryption applied by Word, unlike document protection which is an
    editing restriction). output_path defaults to the same file; a
    different path saves a new encrypted copy. Requires Word installed;
    opens an invisible instance to apply the encryption."""
    from .com import bridge

    return bridge.save_with_password(
        file_path, output_path, password=password
    )


@mcp.tool
def com_validate_opens_clean(file_path: str) -> dict:
    """Run the definitive corruption check: open the file in an invisible Word
    instance and report clean/fail. The Word-verdict companion to
    validate_document / diagnose_document.
    """
    from .com import bridge

    return bridge.validate_opens_clean(file_path)



# ==================================================== workflow suites


@mcp.tool
def validate_cross_references(file_path: str) -> dict:
    """Check every REF/PAGEREF cross-reference against the bookmarks that
    actually exist: broken refs (target missing; renders as an error in
    Word), bookmarks nothing references (informational), and plain-text
    references like 'see Figure 3' that match no caption or heading number
    (heuristic review candidates). Paragraph indices included for every
    finding. Read-only.
    """
    return _ig.validate_cross_references(DocxPackage(file_path))


@mcp.tool
def validate_captions(file_path: str) -> dict:
    """Check that every body-level table and image has an adjacent caption
    paragraph with a SEQ number field, report the ones missing captions with
    locations, and flag mixed numbering conventions (sequential "Figure 3"
    vs chapter-relative "Figure 4.2") per label."""
    return _ig.validate_captions(DocxPackage(file_path))


@mcp.tool
def prepare_for_submission(
    file_path: str,
    accept_revisions: bool = True,
    remove_comments: bool = True,
    scrub_metadata: bool = True,
    keep_title: bool = True,
    backup: bool = True,
) -> dict:
    """Prepare a manuscript for submission in one call: accept all tracked
    changes (every author), delete all comments including the
    comments-family parts, and scrub identifying metadata (author,
    last-modified-by, company; title kept unless keep_title=False). Content
    is never removed: footnotes, fields, and citations all stay. Refuses
    protected documents rather than delivering a half-clean file. Reports
    what was done, what was deliberately left (rsids), and what remains.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _cu.prepare_for_submission(
            pkg,
            accept_revisions=accept_revisions,
            remove_comments=remove_comments,
            scrub_metadata=scrub_metadata,
            keep_title=keep_title,
        ),
        backup=backup,
    )


@mcp.tool
def list_reference_fields(file_path: str) -> dict:
    """Inventory every Zotero, EndNote, and Mendeley field in the body,
    footnotes, and endnotes: manager, kind (citation or bibliography),
    location, cached rendered text, and whether the field markers are still
    intact (broken pairs are reported loudly; they disconnect the citation
    from its manager). Mendeley/EndNote support is preservation-only: their
    fields survive edits, but this server cannot insert or generate them
    (Zotero insertion exists: insert_zotero_citation). Read-only.
    """
    return _rf.list_reference_fields(DocxPackage(file_path))


@mcp.tool
def check_reference_field_integrity(file_path: str) -> dict:
    """Check reference-manager citations after an edit: counts by manager and
    kind plus an ok flag that goes False on any broken or stray field
    marker. Run it after editing a document containing Zotero, EndNote, or
    Mendeley citations (the latter two are preservation-only: fields
    survive edits, no insertion). Read-only.
    """
    return _rf.check_reference_field_integrity(DocxPackage(file_path))


@mcp.tool
def check_template_compliance(file_path: str, rules: dict) -> dict:
    """Validate the document against a formatting ruleset (university
    dissertation guide, journal style sheet). Every rule key is optional;
    unknown keys are rejected with the allowed list. Example ruleset:

    {"page": {"margins_pt": {"top": 72, "bottom": 72, "left": 90,
                           "right": 72},
              "tolerance_pt": 1, "size": "letter",
              "orientation": "portrait"},
     "fonts": {"allowed": ["Times New Roman"], "body_size_pt": 12},
     "line_spacing": {"body": 2.0},
     "headings": {"max_skip": 0, "required_first_level": 1},
     "page_numbering": [{"section": 0, "format": "lowerRoman"},
                        {"section": 1, "format": "decimal",
                         "restart_at": 1}],
     "required_headings_in_order": ["Abstract", "Acknowledgments"]}

    Returns {compliant, violations: [{rule, expected, found, location,
    severity}], unverified, rules_checked}. Fonts/sizes/spacing resolve
    through explicit formatting, the basedOn chain, and docDefaults;
    theme-indirected fonts land in 'unverified', never guessed. Read-only.
    """
    return _cp.check_template_compliance(DocxPackage(file_path), rules)


@mcp.tool
def check_brand_compliance(file_path: str, rules: dict) -> dict:
    """Brand-guide compliance: the same engine and ruleset schema as
    check_template_compliance, plus a colors rule:

    {"fonts": {"allowed": ["Georgia", "Arial"]},
     "colors": {"allowed_hex": ["1F4E79", "C00000"]}}

    colors checks explicit run color values (w:color) in body text against
    the allowed palette; hex comparison ignores case and a leading '#'.
    Theme-indirected colors go to "unverified". Read-only."""
    return _cp.check_brand_compliance(DocxPackage(file_path), rules)


@mcp.tool
def audit_accessibility(file_path: str) -> dict:
    """Audit accessibility, read-only: heading hierarchy (skipped levels,
    empty headings, no Heading 1), images missing alt text, tables whose
    first row is not a repeating header, low-contrast text (explicit color
    vs explicit run background below WCAG 4.5:1; skipped, not guessed, when
    either side is absent), missing document title, and generic hyperlink
    text ('click here'). Each finding carries a location and a fix hint;
    the summary has per-category counts and a pass flag. fix_accessibility
    repairs opted-in categories.
    """
    return _ac.audit_accessibility(DocxPackage(file_path))


@mcp.tool
def fix_accessibility(
    file_path: str,
    alt_text_placeholders: bool = False,
    heading_skips: bool = False,
    heading_strategy: str = "promote",
    table_headers: bool = False,
    doc_title: bool = False,
    dry_run: bool = False,
    backup: bool = True,
) -> dict:
    """Repair audit_accessibility findings, one opted-in category at a time
    (detection mirrors the audit, so audit -> fix -> audit shows the
    opted-in findings resolved). Every category defaults to OFF; with none
    enabled the call refuses. Each category reports fixed items, skipped
    items with reasons, and items needing human review.

    alt_text_placeholders: images without alt text get a clearly marked
    placeholder (never an invented description), all listed for a human
    pass via set_image_alt_text. heading_skips: repairs skipped levels via
    heading_strategy 'promote' or 'demote_following'; nested/mixed
    patterns, custom-style headings, and any repair that would leave a skip
    or remove the only Heading 1 are refused with details. table_headers:
    sets w:tblHeader on flagged first rows, skipping single-row tables and
    empty or all-numeric first rows. doc_title: fills an EMPTY title from
    the first Heading 1, never overwriting.

    dry_run=True returns the identical report without touching the file.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    if dry_run:
        return _ac.fix_accessibility(
            DocxPackage(file_path),
            alt_text_placeholders=alt_text_placeholders,
            heading_skips=heading_skips,
            heading_strategy=heading_strategy,
            table_headers=table_headers,
            doc_title=doc_title,
            dry_run=True,
        )
    return _edit(
        file_path,
        lambda pkg: _ac.fix_accessibility(
            pkg,
            alt_text_placeholders=alt_text_placeholders,
            heading_skips=heading_skips,
            heading_strategy=heading_strategy,
            table_headers=table_headers,
            doc_title=doc_title,
            dry_run=False,
        ),
        backup=backup,
    )


@mcp.tool
def check_image_resolution(file_path: str, min_dpi: int = 300) -> dict:
    """Effective print resolution of every image: native pixel size (PNG/JPEG/
    GIF parsed from the part bytes) vs displayed size gives horizontal and
    vertical DPI; images below min_dpi are flagged with the actual numbers.
    EMF/WMF report "vector (not applicable)"; other formats report
    "unchecked (<format>)". Publishers typically require 300 DPI for print
    figures. Read-only."""
    return _ac.check_image_resolution(DocxPackage(file_path), min_dpi=min_dpi)


@mcp.tool
def list_template_placeholders(file_path: str) -> dict:
    """Find every {{name}} placeholder and legacy MERGEFIELD in a template
    (body, tables, headers/footers, footnotes), with counts and locations.
    The returned names are the keys fill_template and mail_merge expect."""
    return _mm.list_template_placeholders(DocxPackage(file_path))


@mcp.tool
def fill_template(
    file_path: str,
    data: dict,
    missing: str = "error",
    backup: bool = True,
) -> dict:
    """Fill a template IN PLACE: replace every {{name}} with data[name] (safe
    across fragmented runs, first run's formatting kept) and set
    MERGEFIELDs to their values as plain text. missing: 'error' refuses and
    changes nothing if the template needs a key data lacks; 'skip' leaves
    those markers; 'empty' fills them with empty strings. To produce filled
    COPIES instead of editing in place, use mail_merge. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _mm.fill_template(pkg, data, missing=missing),
        backup=backup,
    )


@mcp.tool
def mail_merge(
    template_path: str,
    data_rows: list[dict] | str,
    output_dir: str,
    filename_pattern: str = "{row_index}.docx",
    missing: str = "error",
) -> dict:
    """Mail merge: one filled .docx per data row, saved into output_dir.
    data_rows: a list of dicts, or a path to a .csv (header = placeholder
    names) or .json (array of objects). filename_pattern supports
    {row_index} (1-based) and any {column}. Refuses BEFORE writing anything
    on existing-file collisions, duplicate output names, or (missing='error')
    rows lacking values the template needs. The template is never modified."""
    return _mm.mail_merge(
        template_path,
        data_rows,
        output_dir,
        filename_pattern=filename_pattern,
        missing=missing,
    )


# --------------------------------------------------------- batch operations


@mcp.tool
def batch_apply(
    file_paths: list[str],
    operations: list[dict],
    stop_on_error: bool = True,
    backup: bool = True,
) -> dict:
    """Apply the same operations to MANY documents in one call (e.g. update
    a footer across 50 templates). operations: [{'tool': name, 'params':
    {...}}] with each tool's normal parameters minus file_path. Allowed
    tools: search_and_replace, insert_paragraphs, delete_paragraphs,
    replace_paragraph_text, format_text, set_paragraph_format, apply_style,
    set_header, set_footer, add_page_numbers, set_page_number_format,
    set_document_properties, set_cells, add_watermark, remove_watermark.
    Per file: all operations run, then one prev/anchor slot rotation and one
    atomic validated save; a failing operation leaves that file untouched,
    and stop_on_error=True skips remaining files (already-saved files keep
    their changes). Refuses files open in Word.
    """
    return _bt.batch_apply(
        file_paths, operations, stop_on_error=stop_on_error, backup=backup
    )


# ------------------------------------------------------------- form fields


@mcp.tool
def list_form_fields(file_path: str) -> dict:
    """Every fillable form field in the document: legacy fields (FORMTEXT /
    FORMCHECKBOX / FORMDROPDOWN with name, value, options) and modern
    content controls (text, rich text, checkbox, dropdown/combo, date, with
    tag/alias and placeholder state), across the body and tables."""
    return _fm.list_form_fields(DocxPackage(file_path))


@mcp.tool
def fill_form_fields(
    file_path: str,
    values: dict,
    missing: str = "error",
    backup: bool = True,
) -> dict:
    """Set form-field values by name (legacy fields) or tag/alias (content
    controls): text fields take strings, checkboxes booleans, dropdowns
    only values from their options (refused otherwise). Duplicate field
    names are refused with locations. missing: 'error' refuses, changing
    nothing, if a key matches no field; 'skip' ignores and reports them.
    For SDT types this tool skips, use set_content_control_value.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fm.fill_form_fields(pkg, values, missing=missing),
        backup=backup,
    )


@mcp.tool
def validate_form_completeness(
    file_path: str, required: list[str] | None = None
) -> dict:
    """Report unfilled form fields: empty text, placeholder text still
    showing, and (for names listed in `required`) unchecked checkboxes or
    fields missing from the document entirely. Without `required`, every
    field is checked."""
    return _fm.validate_form_completeness(
        DocxPackage(file_path), required=required
    )




@mcp.tool
def assemble_front_matter(file_path: str, spec: dict, backup: bool = True) -> dict:
    """Assemble long-document front matter in one call: the requested sequence
    is inserted at the START (existing content becomes the body), page
    breaks separate front-matter pages, and one section break lets page
    numbering switch (front lowerRoman, body decimal restarting at 1, by
    default).

    spec = {"sections": [{"kind": "title_page", "lines": [...]},
                         {"kind": "blank_or_copyright", "lines": [...]},
                         {"kind": "abstract", "title": "Abstract",
                          "text": "..."},
                         {"kind": "toc"}, {"kind": "list_of_figures"},
                         {"kind": "list_of_tables"}],
            "page_numbering": {"front": "lowerRoman", "body": "decimal",
                               "body_restart_at": 1}}

    Title-page lines are centered; the abstract gets a Heading-styled
    title. Refuses when front matter appears to exist (leading TOC, or a
    lowerRoman section) unless spec has "force": true. Reports what was
    inserted and which sections carry which numbering. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fma.assemble_front_matter(pkg, spec),
        backup=backup,
    )


@mcp.tool
def setup_chapter_headers(
    file_path: str,
    level: int = 1,
    include_number: bool = False,
    alignment: str = "right",
    first_page_blank: bool = True,
    scope: str | list = "auto",
    backup: bool = True,
) -> dict:
    """Put the current chapter title in the running header via a STYLEREF
    field referencing the Heading style of `level`: the standard Word
    mechanism, evaluated per page, no per-chapter section breaks needed.
    include_number adds the heading number (STYLEREF \\n). scope 'auto'
    targets every section containing body headings of that level; pass a
    list of section indices to override. first_page_blank sets titlePg (no
    header on section-opening pages). Watermarks are preserved; other
    header content is replaced and reported, with the exact field codes
    written. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _ch.setup_chapter_headers(
            pkg,
            level=level,
            include_number=include_number,
            alignment=alignment,
            first_page_blank=first_page_blank,
            scope=scope,
        ),
        backup=backup,
    )


@mcp.tool
def validate_chapter_headers(file_path: str) -> dict:
    """Read back the chapter-header state: which sections carry STYLEREF
    header fields (with the field codes), which heading levels each section
    contains, and which sections with chapter-level headings lack a STYLEREF
    header (the gaps setup_chapter_headers would fill)."""
    return _ch.validate_chapter_headers(DocxPackage(file_path))


@mcp.tool
def diagnose_document(file_path: str) -> dict:
    """Produce a one-call structural health report, read-only: content-type
    coverage, dangling relationships and orphan parts, field balance per
    story part, footnote/endnote integrity, references to undefined styles
    and numbering, content-control and bookmark sanity, duplicate revision
    ids, missing image targets, broken cross-references, and a per-part
    size profile. Never fails on a weird-but-openable document; every check
    degrades to a reported problem, and ok=false only for problems that
    render broken or lose content in Word. The deep companion to
    validate_document's quick check.

    No live mode BY DESIGN: this reads the saved package's XML, stale while
    Word holds unsaved changes. Close the document first
    (com_close_open_document saves and closes), or use
    com_validate_opens_clean / live get_document_info.
    """
    from .core.errors import DocumentLocked

    try:
        return _dg.diagnose_document(DocxPackage(file_path))
    except DocumentLocked as exc:
        raise DocumentLocked(
            f"{file_path} is open in Word (or locked). diagnose_document "
            "reads the saved file's XML, which is stale while Word holds "
            "unsaved changes, so there is no live route by design. Close "
            "the document (com_close_open_document saves and closes) and "
            "rerun; for open-document checks use com_validate_opens_clean "
            "or get_document_info (live)."
        ) from exc


@mcp.tool
def redact_text(
    file_path: str,
    targets: list[dict],
    replacement: str = "[REDACTED]",
    scope: str = "all",
    backup: bool = True,
) -> dict:
    """Permanently REMOVE matched text (true redaction: the characters are
    replaced in the XML, not highlighted). targets: [{find, regex?}].
    Runmap-safe: secrets fragmented across Word's split runs are removed as
    one match. Scrubs body incl. tables, headers/footers,
    footnotes/endnotes (per scope: body | headers | footnotes | all), and
    ALWAYS: comment text, document properties, hyperlink text/tooltips/URL
    targets, field instructions, cached field results, tracked-change
    deleted text. Reports per-class counts, what was NOT examined (images,
    charts, OLE; text drawn in an image is not redacted), and
    verified_clean from a full post-redaction re-scan. Zero-width regexes
    and empty finds are refused up front; any error leaves the file
    unchanged. Irreversible in the saved file: the prev backup slot is the
    undo. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _rx.redact_text(
            pkg, targets, replacement=replacement, scope=scope
        ),
        backup=backup,
    )


@mcp.tool
def verify_redaction(file_path: str, targets: list[dict]) -> dict:
    """Re-scan a document for the given patterns without changing anything:
    do they still appear ANYWHERE in the XML (visible text across fragmented
    runs, deleted tracked-change text, field instructions, attributes,
    metadata, hyperlink targets)? Use it to audit a third-party file or a
    file redacted elsewhere. clean=True covers every XML part; binary parts
    (images, OLE objects) are listed under not_examined, never silently
    trusted. Read-only."""
    return _rx.verify_redaction(DocxPackage(file_path), targets)


@mcp.tool
def check_defined_terms(
    file_path: str, definition_patterns: list | None = None
) -> dict:
    """BETA (heuristic): review the flagged-items list in the result rather
    than trusting silently. Audit a legal document's defined terms. Finds
    definitions («"Term" means», «"Term" shall mean», «(the "Term")»,
    «(each, a "Term")»; defaults overridable via definition_patterns, each
    regex capturing the term as group 1) and reports, with paragraph
    indices: defined_never_used, defined_multiple_times,
    first_use_before_definition, and used_never_defined (a HEURISTIC list
    of capitalized recurring terms, filtered so sentence-start capitals are
    not flagged; treat as review candidates). Body-level paragraphs only.
    Read-only.
    """
    return _dt.check_defined_terms(
        DocxPackage(file_path), definition_patterns=definition_patterns
    )


@mcp.tool
def com_import_pdf(pdf_path: str, output_path: str | None = None) -> dict:
    """Convert a PDF to .docx via Word's built-in PDF reflow, in a dedicated
    invisible Word instance. Output defaults next to the PDF with a .docx
    extension; an existing output file is refused; the produced .docx is
    validated by a full package round-trip. Text-based PDFs convert well,
    complex layouts may reflow imperfectly, and scanned image PDFs yield
    little or no text (Word does not OCR; a near-zero word count triggers
    an explicit warning).
    """
    from word_mcp.com import convert  # when pasted: from .com import convert

    return convert.import_pdf(pdf_path, output_path)




@mcp.tool
def add_equation(
    file_path: str,
    latex: str,
    display: bool = True,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    anchor_text: str | None = None,
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Insert a LaTeX equation as NATIVE Word math (editable in Word's
    equation editor, not an image). display=True: a block equation in its
    own paragraph, positioned by exactly one of after_index / after_anchor
    / at_end. display=False: inline immediately after anchor_text within
    that paragraph (occurrence picks the match), surrounding text
    untouched. Supports the standard academic repertoire: fractions, roots,
    sums/integrals with limits, matrices, cases, align* (aligned is
    rewritten automatically), Greek, operators, accents, sub/superscripts,
    \\text{}. Unconvertible LaTeX raises a clear error; the document is NOT
    modified. Equations are invisible to get_text/find_text
    (search-and-replace is equation-safe); read them with list_equations.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _eq.add_equation(
            pkg,
            latex,
            display=display,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
            anchor_text=anchor_text,
            occurrence=occurrence,
        ),
        backup=backup,
    )


@mcp.tool
def list_equations(file_path: str) -> dict:
    """List every equation in the document (body, tables, footnotes,
    endnotes): index, display vs inline, location, and a plain-text
    approximation of the math. THIS is how equation content is read;
    equations are invisible to get_text and find_text because math runs sit
    outside the plain-text layer. The index is the handle delete_equation
    takes. Read-only.
    """
    return _eq.list_equations(DocxPackage(file_path))


@mcp.tool
def delete_equation(file_path: str, index: int, backup: bool = True) -> dict:
    """Delete an equation by its list_equations index. A display equation's
    paragraph is removed with it when the paragraph holds nothing else
    (the usual case); an inline equation is removed from within its
    paragraph, surrounding text untouched. Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word."""
    return _edit(
        file_path,
        lambda pkg: _eq.delete_equation(pkg, index),
        backup=backup,
    )


@mcp.tool
def search_zotero_library(
    query: str, db_path: str | None = None, limit: int = 20
) -> dict:
    """Search the user's LOCAL Zotero library (read-only; the database is
    never modified). Every query word must match the title, a creator's
    last name, the year, or the publication name. Returns each match's
    Zotero item key (the handle insert_zotero_citation takes) plus type,
    title, creators, year, and publication. Attachments, notes, and trashed
    items are excluded. The database defaults to
    <home>/Zotero/zotero.sqlite; pass db_path for a nonstandard data
    directory. Reads a point-in-time snapshot, so last-moment edits in a
    running Zotero may not appear yet.
    """
    return _zl.search_zotero_library(query, db_path=db_path, limit=limit)


@mcp.tool
def insert_zotero_citation(
    file_path: str,
    item_keys: list[str],
    anchor_text: str,
    occurrence: int = 1,
    page: str | None = None,
    prefix: str | None = None,
    suffix: str | None = None,
    db_path: str | None = None,
    backup: bool = True,
) -> dict:
    """Insert a REAL Zotero citation field (ADDIN ZOTERO_ITEM CSL_CITATION)
    after anchor_text, built exactly as the Zotero Word plugin builds it,
    so Zotero recognizes, refreshes, and bibliographs it. item_keys come
    from search_zotero_library; page becomes the locator; prefix/suffix
    attach to the cited item (single-item citations only; several keys
    refuse rather than guess). The visible text is a plain (Author, Year)
    PLACEHOLDER until the user clicks Refresh in Zotero's Word plugin. The
    Zotero database is only ever read. Run detect_citation_system first on
    unfamiliar documents: mixing Word-native and Zotero/Mendeley/EndNote
    citations creates a bibliography no single manager maintains.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _zl.insert_zotero_citation(
            pkg,
            item_keys,
            anchor_text=anchor_text,
            occurrence=occurrence,
            page=page,
            prefix=prefix,
            suffix=suffix,
            db_path=db_path,
        ),
        backup=backup,
    )


@mcp.tool
def find_formatted(
    file_path: str,
    formatting: dict,
    query: str | None = None,
    scope: str = "body",
) -> dict:
    """Find text by its EFFECTIVE formatting: bold/italic/underline/strike
    (true/false), font, size_pt, color (hex), highlight, style (id or
    name); all criteria must hold together. Resolution follows Word (run
    properties, then character style, then paragraph-style chain, then
    document defaults); every match reports which level satisfied each
    criterion, so explicit and Heading-inherited bold are distinguishable.
    query=None returns every stretch with the formatting; a query
    (case-sensitive) returns only occurrences inside such stretches. scope:
    'body' (tables included) or 'all' (adds footnotes/endnotes). Text-box
    content is EXCLUDED: read boxes via get_textbox_text or get_text
    include_textboxes. Theme references are counted, never guessed.
    replace_formatted is the mutation twin. Read-only.
    """
    return _sf.find_formatted(
        DocxPackage(file_path), query, formatting=formatting, scope=scope
    )


@mcp.tool
def get_paragraph_format(
    file_path: str, start: int, end: int | None = None
) -> dict:
    """Read EFFECTIVE paragraph formatting for body paragraphs start..end
    (inclusive; end defaults to start): alignment, space_before_pt,
    space_after_pt, line_spacing, indent_left_pt, indent_right_pt,
    first_line_indent_pt (negative = hanging), keep_with_next,
    widow_control, page_break_before. Resolution follows Word (own
    properties, then the style basedOn chain, then document defaults);
    every property reports value AND source ('explicit' | 'paragraph_style'
    | 'document_defaults' | 'word_default'), mirroring find_formatted's
    matched_via, so direct overrides and style inheritance are
    distinguishable. line_spacing carries its rule ('auto' = multiple;
    'exact'/'atLeast' = points). Numbering-contributed indents are not
    resolved (flagged has_numbering). set_paragraph_format takes the same
    keys. Read-only.
    """
    return _sf.get_paragraph_format(DocxPackage(file_path), start, end)


@mcp.tool
def comment_report(file_path: str, include_resolved: bool = True) -> dict:
    """The full reviewer matrix in one call: every comment thread with author,
    initials, date, the anchored text it targets, the comment text, replies
    nested under their parents, resolved flag, and a locator (paragraph index
    plus the heading path of the section it falls under). Summary gives
    per-author, open/resolved, and per-section counts. Built for processing
    committee feedback without re-reading comments piecemeal. Set
    include_resolved=False to drop resolved threads (the excluded count is
    reported). Read-only."""
    return _rc.comment_report(
        DocxPackage(file_path), include_resolved=include_resolved
    )


@mcp.tool
def comment_report_multi(file_paths: list[str]) -> dict:
    """Merge the reviewer matrix across several documents (e.g. three
    committee members' copies of the same draft). Entries are keyed by
    (author, anchored text) with per-file provenance on every occurrence,
    and collisions (the same span commented on by two or more reviewers)
    are detected and listed. Copies should share the draft text; merging
    matches anchored spans verbatim. Read-only on every file.
    """
    return _rc.comment_report_multi(file_paths)


@mcp.tool
def revision_analytics(file_path: str) -> dict:
    """Tracked-change analytics per author: insertion/deletion counts, words
    added and removed, move and format-change counts, the date range of each
    author's edits, and a per-section (heading-path) breakdown of where their
    changes concentrate. Includes the 10 heaviest body paragraphs by revision
    churn. Footnote/endnote revisions count under "[footnotes]"/"[endnotes]".
    Read-only."""
    return _rc.revision_analytics(DocxPackage(file_path))


@mcp.tool
def structured_diff(old_path: str, new_path: str, detail_cap: int = 200) -> dict:
    """Diff two saved drafts agent-readably, computed without Word:
    unchanged/modified/inserted/deleted paragraphs with indices in both
    documents, moved paragraphs (identical text at a new position),
    intra-paragraph change opcodes, per-section change summary, table
    changes (dimensions plus changed cells, compared positionally),
    footnote/endnote count deltas, and heading structure changes.
    Deliberately NOT a redline: use com_compare_documents for a
    Word-rendered tracked-changes comparison. Per-change detail is capped
    at detail_cap entries on huge diffs (counts stay complete). Read-only
    on both files.
    """
    return _rc.structured_diff(old_path, new_path, detail_cap=detail_cap)


@mcp.tool
def get_textbox_text(file_path: str) -> dict:
    """Read the text inside every text box / shape text frame, per box, across
    the body, headers, and footers (modern drawings and legacy VML alike).
    Each box reports its text, paragraphs, part, anchoring body paragraph
    index (where determinable), and shape name. Use this instead of
    get_text for box content: the generic tools smear box text into the
    host paragraph (doubled by the mc:Fallback compatibility copy) with no
    boundary. box_index is the address set_textbox_text takes. Read-only.
    """
    return _tbx.get_textbox_text(DocxPackage(file_path))


@mcp.tool
def set_textbox_text(
    file_path: str, box_index: int, text: str, backup: bool = True
) -> dict:
    """Replace the text of one text box (box_index from get_textbox_text).
    Keeps the first paragraph's style and first run's formatting; '\\n'
    splits into multiple paragraphs; the mc:Fallback compatibility copy is
    rewritten to match. Boxes holding non-text content (nested tables,
    images) are refused rather than flattened. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tbx.set_textbox_text(pkg, box_index, text),
        backup=backup,
    )


@mcp.tool
def preview_replace(
    file_path: str, replacements: list[dict], scope: str = "body"
) -> dict:
    """Dry-run search_and_replace; the file is NEVER modified. Same items
    ({find, replace, regex?}), scope (body | footnotes | headers | all),
    and matching engine (fragmented-run map, guarded regex, chained items),
    so the preview shows exactly what a real run would change: per-item
    counts, each match with paragraph index and ~60 chars of before/after
    context, the grand total, and the refusals a real run would hit.
    Review, then run search_and_replace with max_replacements set to the
    previewed total so any drift aborts instead of over-replacing.
    """
    return _pv.preview_replace(
        DocxPackage(file_path), replacements, scope=scope
    )


@mcp.tool
def word_count_with_exclusions(
    file_path: str,
    exclude: list[str] = ["references", "captions", "footnotes"],
) -> dict:
    """Count words minus named zones: the number a journal actually wants.
    exclude any of: references, captions, footnotes, endnotes,
    block_quotes, tables, headings, front_matter, abstract (unknown names
    rejected with the allowed list). Returns total, per-zone excluded
    breakdown, and the included count; total = included + sum(excluded)
    always. Same whitespace tokenization as word_count's file mode;
    detected zone locations are reported for review. Read-only.
    """
    return _jc.word_count_with_exclusions(
        DocxPackage(file_path), exclude=tuple(exclude)
    )


@mcp.tool
def anonymize_for_review(
    file_path: str,
    author_names: list[str],
    replacement: str = "Author",
    mapping_path: str | None = None,
    backup: bool = True,
) -> dict:
    """BETA (heuristic): review the flagged-items list in the result rather
    than trusting silently. Anonymize a manuscript for double-blind peer
    review, reversibly. Masks self-citations by the named authors ('Hurd
    (1999)' -> 'Author (1999)') keeping years and pages, rewrites their
    reference-list entries to 'Author (Year). [Details removed for peer
    review.]', and scrubs identifying metadata. Prose that identifies the
    author (Acknowledgments, 'my previous work', surnames outside citation
    syntax) is FLAGGED with locations, never auto-edited. Writes a reversal
    mapping JSON (default <name>.anonymization.json, never overwritten) for
    deanonymize_document; KEEP IT PRIVATE. Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _an.anonymize_for_review(
            pkg,
            author_names,
            replacement=replacement,
            mapping_path=mapping_path,
        ),
        backup=backup,
    )


@mcp.tool
def deanonymize_document(
    file_path: str, mapping_path: str | None = None, backup: bool = True
) -> dict:
    """Reverse anonymize_for_review using its mapping file (default
    <name>.anonymization.json beside the document). Every recorded change
    is verified to still sit where the mapping says before anything is
    restored; if the document drifted since anonymization, NOTHING is
    restored and the refusal lists every mismatch. The mapping file is left
    on disk to delete once the restore is confirmed. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _an.deanonymize(pkg, mapping_path=mapping_path),
        backup=backup,
    )


@mcp.tool
def export_table(
    file_path: str,
    table_index: int,
    format: str = "csv",
    output_path: str | None = None,
    include_merges: bool = True,
) -> dict:
    """Export a body-level table (list_tables index) to CSV or JSON as a
    rows x grid-columns matrix. Merged cells keep their value in the anchor
    (top-left) position with empty strings in the covered positions;
    include_merges adds the merge topology as {row, col, rowspan, colspan}
    entries (inside the JSON document; as a report field alongside CSV).
    Nested tables are flattened into the host cell's text and flagged.
    output_path=None returns the data inline ('csv' string / 'data' rows)
    instead of writing a file; an existing output file is refused. The
    document itself is never modified."""
    return _dio.export_table(
        DocxPackage(file_path),
        table_index,
        format=format,
        output_path=output_path,
        include_merges=include_merges,
    )


@mcp.tool
def import_table(
    file_path: str,
    data: list | str,
    table_index: int | None = None,
    at_end: bool = False,
    after_anchor: str | None = None,
    has_header: bool = True,
    backup: bool = True,
) -> dict:
    """Fill a table from data: a .csv path, a .json path (export_table's form
    or a bare row list), or an inline list of lists. Without table_index a
    NEW table is created (at_end/after_anchor, default end; has_header
    bolds and repeats the first row). With table_index the existing table's
    cell texts are OVERWRITTEN in place: data must exactly match rows x
    grid-columns (refused otherwise, listing both shapes); merged cells
    take their value at the anchor position; values in merge-covered
    positions are refused. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _dio.import_table(
            pkg,
            data,
            table_index=table_index,
            at_end=at_end,
            after_anchor=after_anchor,
            has_header=has_header,
        ),
        backup=backup,
    )


@mcp.tool
def extract_images(
    file_path: str, output_dir: str, prefix: str | None = None
) -> dict:
    """Extract every image in the document to files in output_dir, named by
    list_images index plus the original extension (image0.png, ...; prefix
    replaces 'image'). Media referenced only from headers/footers/notes gets
    the next indices with the referencing part reported. Each entry reports
    the output file, native pixel dimensions, and where the image appears
    (body paragraph/table index). Collisions with existing files are refused
    before anything is written. Read-only against the document."""
    return _dio.extract_images(
        DocxPackage(file_path), output_dir, prefix=prefix
    )


@mcp.tool
def split_document(
    file_path: str,
    output_dir: str,
    level: int = 1,
    filename_from: str = "heading",
) -> dict:
    """Split a document into one standalone .docx per heading section, the
    inverse of com_merge_documents (no Word needed). Sections start at each
    heading of `level` (or higher) and carry everything to the next one;
    content before the first heading becomes 00_front_matter.docx when
    non-empty. Outputs carry the source's styles, numbering, settings,
    fonts, and themes; other sections' note definitions and image parts are
    stripped; every output is round-trip validated. filename_from:
    'heading' gives 01_Heading Text.docx, 'index' gives 01.docx. The source
    file is never modified; existing output files are refused before
    anything is written.
    """
    return _dio.split_document(
        file_path, output_dir, level=level, filename_from=filename_from
    )


@mcp.tool
def parse_references(file_path: str, style_hint: str | None = None) -> dict:
    """Parse the manuscript's reference list (References / Bibliography /
    Works Cited heading) and in-text citations into a structured model.
    Read-only stage 1 of publication-style conversion.

    Heuristic text parsing: every entry gets parse_confidence
    (full/partial/failed); failed entries are returned verbatim and will
    never be converted. Also reports in-text citations with positions
    (parenthetical, narrative, [n] bracket, superscript), whether the
    document uses Word-native or Zotero/Mendeley/EndNote citation FIELDS
    (which text conversion refuses to touch), and the detected citation
    system. style_hint (e.g. 'apa7', 'ieee', 'turabian') tightens parsing
    when the source style is known."""
    return _sc2.parse_references(DocxPackage(file_path), style_hint=style_hint)


@mcp.tool
def convert_citation_style(
    file_path: str,
    target_style: str,
    source_style: str = "auto",
    dry_run: bool = False,
    backup: bool = True,
) -> dict:
    """BETA (heuristic): review the flagged-items list in the result rather
    than trusting silently. Convert a manuscript's PLAIN-TEXT citations and
    reference list to a target publication style: apa7,
    chicago17-author-date, chicago17-notes (= Turabian), mla9, harvard,
    ieee, vancouver, asa.

    Heuristic text conversion, not a citation processor: only fully-parsed
    reference entries and unambiguously-resolved citations are converted;
    everything else is left verbatim and flagged for human review. Supports
    author-date <-> numbered ([n] or Vancouver superscripts, numbered by
    first appearance), author-date -> Chicago notes (each citation becomes
    a REAL footnote), and notes -> author-date (only recognizably pure
    citation footnotes are harvested; mixed ones are flagged and left
    alone). Narrative citations keep the author's name; reference-list
    italics, ordering, heading, and hanging indents are handled. Documents
    using Word-native or Zotero/Mendeley/EndNote citation FIELDS are routed
    away (set_bibliography_style, or restyle in the manager); fields are
    never rewritten as text. Run detect_citation_system first on unfamiliar
    documents: mixing Word-native and Zotero/Mendeley/EndNote citations
    creates a bibliography no single manager maintains.

    dry_run=True returns the complete change plan and leaves the file
    byte-identical; parse_references is the read-only stage 1. Any error
    during a real run leaves the original untouched. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    if dry_run:
        return _sc2.convert_citation_style(
            DocxPackage(file_path), target_style,
            source_style=source_style, dry_run=True,
        )
    return _edit(
        file_path,
        lambda pkg: _sc2.convert_citation_style(
            pkg, target_style, source_style=source_style, dry_run=False
        ),
        backup=backup,
    )


@mcp.tool
def apply_manuscript_format(
    file_path: str,
    style: str,
    running_head: str | None = None,
    author_last_name: str | None = None,
    backup: bool = True,
) -> dict:
    """Apply a style's page-level manuscript conventions where publicly
    well-defined: 'apa7' (= apa7-student: 1in margins, double spacing, page
    number top right, APA heading formats L1-L5), 'apa7-professional' (adds
    the running head; running_head or the title, flagged), 'mla9' (surname
    + page header; author_last_name or the author metadata, flagged),
    'chicago17' (= turabian: 12pt base, double-spaced body, single-spaced
    footnotes, page number bottom center). All three set a hanging indent
    on a found reference list. Anything contested (run-in headings, casing,
    indents near front matter) is NOT applied and is itemized in
    not_applied. IEEE/Vancouver/ASA/Harvard page formats are
    journal-template-specific and refused. Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sc2.apply_manuscript_format(
            pkg, style, running_head=running_head,
            author_last_name=author_last_name,
        ),
        backup=backup,
    )



# ------------------------------------------------- api-review absorptions


@mcp.tool
def get_workflows(task: str | None = None) -> dict:
    """Recommended tool sequences for common multi-step tasks, with a
    one-line why per step. Call with no task to list the available tasks
    ('process-feedback', 'prepare-submission', 'format-citations',
    'build-from-template', 'heavy-editing'); call with task='<name>' for
    that task's step-by-step sequence and notes. Pure guidance: reads
    nothing, changes nothing."""
    return _wf.get_workflows(task)


@mcp.tool
def detect_citation_system(file_path: str) -> dict:
    """Which citation system(s) the document uses: Word native (CITATION
    fields + sources store), Zotero (ADDIN ZOTERO_ITEM), Mendeley (ADDIN
    CSL_CITATION), EndNote (ADDIN EN.CITE), or plain typed text only. Counts
    per system across body, footnotes, and endnotes, plus a split_brain flag
    when more than one managed system is present (a split-brain bibliography:
    each manager only maintains its own fields). Run this BEFORE any citation
    work on an unfamiliar document."""
    return _cs.detect_citation_system(DocxPackage(file_path))


@mcp.tool
def change_heading_level(
    file_path: str,
    delta: int,
    heading_text: str | None = None,
    paragraph_index: int | None = None,
    subtree: bool = False,
    backup: bool = True,
) -> dict:
    """Promote (delta=-1) or demote (delta=1) a heading, addressed by exact
    heading_text or by paragraph_index. subtree=True also shifts every
    subordinate heading beneath it (until the next same-or-higher heading)
    by the same delta, keeping the branch's shape. Refuses, changing
    nothing, if any affected heading would leave levels 1-9 (the blocker is
    named) or if a heading's level comes from an outlineLvl override or
    custom style rather than a built-in Heading style (adjust those via
    set_paragraph_format's outline_level). Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.change_heading_level(
            pkg,
            delta=delta,
            heading_text=heading_text,
            paragraph_index=paragraph_index,
            subtree=subtree,
        ),
        backup=backup,
    )


@mcp.tool
def insert_field(
    file_path: str,
    field_code: str,
    after_anchor: str,
    occurrence: int = 1,
    placeholder: str = "",
    backup: bool = True,
) -> dict:
    """Insert a generic Word field right after `after_anchor` text (plain
    paragraph text, literal characters, not XML entities; occurrence picks
    which match when the anchor appears more than once). field_code
    examples: 'DATE', 'TIME \\@ "HH:mm"', 'FILENAME', 'NUMPAGES', 'PAGE',
    'SEQ Exhibit \\* Arabic'. Codes are validated against an allowlist of
    known-safe fields (document info, page/date/time numbers, SEQ);
    anything that links out or executes is refused, naming the allowlist.
    The field is written dirty so Word computes the result on the next
    refresh; `placeholder` shows until then. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word."""
    return _edit(
        file_path,
        lambda pkg: _fl.insert_field(
            pkg,
            field_code=field_code,
            after_anchor=after_anchor,
            occurrence=occurrence,
            placeholder=placeholder,
        ),
        backup=backup,
    )


@mcp.tool
def list_fields(file_path: str) -> dict:
    """Every field in the body, headers, footers, footnotes, and endnotes:
    instruction code, field type (DATE, PAGE, SEQ, TOC, ...), the current
    cached result text, and location (part + body paragraph index where
    applicable). Covers complex (fldChar) and simple (fldSimple) fields;
    unclosed complex fields are flagged."""
    return _fl.list_fields(DocxPackage(file_path))


@mcp.tool
def create_snapshot(
    file_path: str,
    label: str | None = None,
    dest_dir: str | None = None,
) -> dict:
    """Save a DTG-stamped permanent copy of the document:
    YYYYMMDD_HHMM_<name>.docx (an existing leading DTG in the name is
    replaced, not stacked), with an optional short label suffix. Snapshots
    are the PERMANENT keepers that complement the automatic prev/anchor
    backup slots: the slots rotate on every mutation, snapshots are never
    auto-pruned and manage_backups never touches them. Never overwrites
    (collisions get a numeric suffix); returns the created path. The source
    document is not modified."""
    return _bk.create_snapshot(file_path, label=label, dest_dir=dest_dir)


@mcp.tool
def list_content_controls(file_path: str) -> dict:
    """Every content control (SDT) in the document body, including the types
    fill_form_fields skips: tag, alias, type (text / richtext / checkbox /
    dropdown / combo / date / picture / group / citation / bibliography /
    equation / gallery / repeating_section), current value, lock state,
    placeholder flag, and block/inline. The index is the addressing handle
    for set_content_control_value when tags are missing or duplicated."""
    return _fm.list_content_controls(DocxPackage(file_path))


@mcp.tool
def set_content_control_value(
    file_path: str,
    value: str | bool,
    tag: str | None = None,
    index: int | None = None,
    backup: bool = True,
) -> dict:
    """Set one content control's value, addressed by tag or by
    list_content_controls index (exactly one). Text, rich-text, combo, and
    date controls take a string; checkboxes a boolean; dropdowns one of
    their options (refused otherwise, naming the options). Locked controls
    and unwritable types (gallery, repeating section, citation,
    bibliography, picture, group, equation) are refused with nothing
    changed. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fm.set_content_control_value(
            pkg, value, tag=tag, index=index
        ),
        backup=backup,
    )


@mcp.tool
def insert_content_control(
    file_path: str,
    tag: str,
    after_anchor: str,
    alias: str | None = None,
    text: str = "",
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Insert a new PLAIN-TEXT content control (inline SDT) right after
    `after_anchor` text (plain paragraph text, literal characters;
    occurrence picks which match when the anchor appears more than once),
    with a unique tag (refused if it exists) and optional alias and initial
    text. Plain text is the one control type this server can build safely;
    creating checkbox, dropdown, date, picture, gallery, or repeating
    controls is refused (list/fill still cover those when a template
    provides them). Fill it later via fill_form_fields or
    set_content_control_value by its tag. Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word."""
    return _edit(
        file_path,
        lambda pkg: _fm.insert_content_control(
            pkg,
            tag=tag,
            after_anchor=after_anchor,
            alias=alias,
            text=text,
            occurrence=occurrence,
        ),
        backup=backup,
    )


@mcp.tool
def insert_glossary(
    file_path: str,
    heading: str = "Glossary",
    heading_level: int = 1,
    after_index: int | None = None,
    at_end: bool = True,
    definition_patterns: list | None = None,
    backup: bool = True,
) -> dict:
    """Build a glossary section from the document's defined terms (same
    detection as check_defined_terms): a heading plus one alphabetized
    paragraph per term, term in bold, definition harvested from the
    defining sentence. Terms whose definition cannot be extracted cleanly
    get a [DEFINITION NEEDED] marker instead of a mangled fragment; the
    result lists them for manual completion. Placed at the body end by
    default, or after body paragraph after_index. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _dt.insert_glossary(
            pkg,
            heading=heading,
            heading_level=heading_level,
            after_index=after_index,
            at_end=at_end,
            definition_patterns=definition_patterns,
        ),
        backup=backup,
    )


# ---------------------------------------------------------- document assembly


@mcp.tool
def insert_document(
    target_path: str,
    source_path: str,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    formatting: str = "source",
    backup: bool = True,
) -> dict:
    """Insert the ENTIRE body of source_path into target_path at one position,
    with full resource reconciliation: the document-assembly tool for
    merging chapter files into one manuscript. com_merge_documents only
    concatenates whole files into a new one; copy_table transplants a
    single table.

    Position (exactly one): after_index is a body ITEM index, counting
    paragraphs AND tables together in document order (unlike the
    paragraph-only indices elsewhere); insertion lands after that item.
    after_anchor matches a paragraph whose FULL plain text equals the
    anchor; several matches refuse and list every location, so recurring
    heading text cannot land content at the wrong spot (prefer after_index
    for structural work). at_end appends after the last body item.

    Carried: tables, images, charts, hyperlinks, lists (fresh numbering),
    footnotes/endnotes (new ids), bookmarks (remapped, collisions renamed
    and reported), tracked changes, equations. Styles reconcile BY NAME
    (the target's formatting wins on a match; unmatched styles are cloned
    in with dependency chains). The source's section setup is never
    carried; mid-content section breaks and comment references are stripped
    and reported. OLE objects, ActiveX, subdocuments, and altChunks refuse
    the whole insertion, naming the blocker; nothing is half-applied.

    formatting mirrors Word's paste modes on the carried copies: 'source'
    (default) keeps direct formatting; 'merge' keeps bold/italic/emphasis
    but strips direct font/size/color/spacing/indent overrides;
    'destination' strips all direct formatting except structural
    properties. Returns per-resource counts, style remaps, bookmark
    renames, and the occupied body-item range. The source file is never
    modified. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        target_path,
        lambda pkg: _asm.insert_document(
            pkg,
            source_path,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
            formatting=formatting,
        ),
        backup=backup,
    )


# ------------------------------------------------------- style-aware replace


@mcp.tool
def replace_formatted(
    file_path: str,
    formatting: dict,
    replace: str,
    find: str | None = None,
    scope: str = "body",
    max_replacements: int | None = None,
    backup: bool = True,
) -> dict:
    """Replace text ONLY where it carries specific EFFECTIVE formatting: the
    mutation twin of find_formatted, same criteria keys
    (bold/italic/underline/strike, font, size_pt, color, highlight, style),
    all required together, resolved the way Word resolves them. find=None
    replaces each entire matching stretch; with find (literal,
    case-sensitive), only occurrences wholly inside a matching stretch.
    Safe across fragmented runs; the replacement keeps the matched
    formatting. Result mirrors search_and_replace ({replaced, total}) plus
    per-replacement matched_via. No regex here (search_and_replace covers
    that); text boxes are never touched (set_textbox_text edits those).
    max_replacements aborts, changing nothing, past the cap. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sf.replace_formatted(
            pkg,
            formatting=formatting,
            replace=replace,
            find=find,
            scope=scope,
            max_replacements=max_replacements,
        ),
        backup=backup,
    )


@mcp.tool
def copy_table(
    target_path: str,
    source_path: str,
    table_index: int,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    backup: bool = True,
) -> dict:
    """Transplant ONE table from source_path into target_path: the
    single-element sibling of insert_document, reusing its resource
    reconciliation scoped to the table (styles matched BY NAME, target
    wins, unmatched cloned; fresh numbering; images/hyperlinks
    re-registered; footnote/endnote definitions under new ids; bookmarks
    remapped, collisions renamed).

    table_index counts the SOURCE's top-level body tables (0-based,
    list_tables order; nested tables travel with their parent). Target
    position (exactly one): after_index is a body ITEM index counting
    paragraphs AND tables together (not the paragraph-only index other
    tools use); after_anchor matches a paragraph's FULL plain text,
    refusing with every location when it recurs; at_end appends.
    OLE/ActiveX/altChunks inside the table refuse the whole copy; nothing
    is half-applied. Returns row/column counts, the occupied body-item
    position, and per-resource counts. The source file is never modified.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        target_path,
        lambda pkg: _asm.copy_table(
            pkg,
            source_path,
            table_index,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
        ),
        backup=backup,
    )


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
