"""word-mcp: FastMCP server exposing the full Word editing suite.

Contract for every mutating tool:
- `file_path` is the target .docx (absolute path).
- A timestamped .bak copy is written next to the file before the first
  mutation unless backup=False.
- Saves are atomic and validated; on any failure the original is untouched
  and the error message says exactly what was wrong.
- A file open in Word is refused with a clear message (COM tools still work).
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
)

mcp = FastMCP(
    "word",
    instructions=(
        "Full-featured Word (.docx) editor: text, tables (including column "
        "insert/delete and bulk cell edits), footnotes/endnotes, TOC, "
        "headers/footers, images, comments, tracked changes. File-based; "
        "documents open in Word are refused (use the com_* tools for those "
        "cases). Every mutation auto-backs up the file first."
    ),
)


def _edit(file_path: str, fn, *, backup: bool = True) -> dict:
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
    """Overview: paragraph/table/footnote/comment/revision counts, sections, parts.
    Documents open in Word are read live."""
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_document_info(DocxPackage(file_path)),
        lambda: _lo.get_document_info(file_path),
    )


@mcp.tool
def get_outline(file_path: str, live: str = "auto") -> list | dict:
    """Headings with paragraph indices and levels. Documents open in Word
    are read live."""
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
    live: str = "auto",
) -> list | dict:
    """Body paragraphs with indices, styles, heading levels. Slice with
    start/end; filter with contains. If the document is open in Word the
    CURRENT in-memory state is read live (live='off' restores v1 refusal)."""
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_paragraphs(
            DocxPackage(file_path), start, end, contains=contains
        ),
        lambda: _lo.get_text(file_path, start, end, contains),
    )


@mcp.tool
def find_text(
    file_path: str,
    query: str,
    regex: bool = False,
    context_chars: int = 60,
    live: str = "auto",
) -> list | dict:
    """Find text in paragraphs and table cells; returns locations + context.
    Documents open in Word are searched live (current in-memory state)."""
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.find_text(
            DocxPackage(file_path), query, regex=regex,
            context_chars=context_chars,
        ),
        lambda: _lo.find_text(file_path, query, regex, context_chars),
    )


@mcp.tool
def get_styles(file_path: str) -> list:
    """Styles defined in the document (id, name, type, based_on)."""
    return _rd.list_styles(DocxPackage(file_path))


@mcp.tool
def word_count(file_path: str, by_section: bool = True) -> dict:
    """Word/character/paragraph counts, total and per heading section."""
    return _st.word_count(DocxPackage(file_path), by_section=by_section)


@mcp.tool
def check_citation_parity(file_path: str) -> dict:
    """Cross-check APA in-text citations against the reference list, both
    directions: cited-but-not-listed (serious) and listed-but-never-cited
    (review). Heuristic flagging tool — results are review candidates."""
    return _cc.check_citation_parity(DocxPackage(file_path))


@mcp.tool
def validate_document(file_path: str) -> dict:
    """Structural integrity: package opens, notes consistent, fields balanced."""
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
    """Create a new blank .docx (refuses to overwrite an existing file)."""
    from pathlib import Path

    from docx import Document

    p = Path(file_path)
    if p.exists():
        raise FileExistsError(f"{file_path} already exists")
    doc = Document()
    if title:
        doc.core_properties.title = title
    p.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(p))
    return {"created": str(p)}


@mcp.tool
def copy_document(file_path: str, dest_path: str) -> dict:
    """Copy a document (e.g. to a new DTG-stamped filename before editing)."""
    import shutil
    from pathlib import Path

    if Path(dest_path).exists():
        raise FileExistsError(f"{dest_path} already exists")
    shutil.copy2(file_path, dest_path)
    return {"copied_to": dest_path}


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
    {find, replace, regex?}. scope: body | footnotes | headers | all.
    max_replacements: abort (changing nothing) if the total match count would
    exceed it — set this when using broad regex patterns. track: record each
    replacement as a tracked change attributed to `author`. Documents open in
    Word are edited LIVE (visible immediately, one Ctrl+Z step, nothing saved
    to disk until the user saves); live='off' restores the v1 refusal."""
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
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Insert paragraphs ({text, style?, formatting?}) at one position:
    after_index | before_index | after_anchor (unique text) | at_end.
    track: record as tracked insertions by `author`. Documents open in Word
    are edited live."""
    from .com import live_ops as _lo

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
                track=track,
                author=author,
            ),
            backup=backup,
        ),
        lambda: _lo.insert_paragraphs(
            file_path,
            paragraphs,
            after_index=after_index,
            before_index=before_index,
            after_anchor=after_anchor,
            at_end=at_end,
            track=track,
            author=author,
        ),
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
    """Delete body paragraphs [start..end] inclusive. Refuses ranges that cut
    through a field or carry a section break. track: mark as tracked deletions
    by `author` instead of removing. Documents open in Word are edited live."""
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
    file_path: str, index: int, new_text: str, backup: bool = True
) -> dict:
    """Replace one paragraph's full text, keeping style and base formatting."""
    return _edit(
        file_path,
        lambda pkg: _tx.replace_paragraph_text(pkg, index, new_text),
        backup=backup,
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
    """Insert a heading (level 1-9). Heading styles are auto-created if the
    document lacks them."""
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
    """Insert a page break after a body paragraph."""
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
    """Character formatting on a text range. formatting keys: bold, italic,
    underline, strike, font, size_pt, color, highlight, superscript, subscript.
    Target: paragraph_index + find (substring), find alone, or whole paragraph.
    Documents open in Word are edited live."""
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
    file_path: str, indices: list[int], formatting: dict, backup: bool = True
) -> dict:
    """Paragraph formatting for a batch of paragraphs. Keys: alignment,
    space_before_pt, space_after_pt, line_spacing, indent_left_pt,
    indent_right_pt, first_line_indent_pt, keep_with_next."""
    return _edit(
        file_path,
        lambda pkg: _tx.set_paragraph_format(pkg, indices, formatting),
        backup=backup,
    )


@mcp.tool
def apply_style(
    file_path: str, indices: list[int], style: str, backup: bool = True
) -> dict:
    """Apply a paragraph style (by id or name) to a batch of paragraphs.
    Heading1-9 are auto-created when missing."""
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
    {text, level} dicts (level 0-8 nests). kind: bullet | number. Each call
    is an independent list, so numbering restarts at 1."""
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
    """List paragraphs grouped by numbering instance, with levels and text."""
    return _ls.get_lists(DocxPackage(file_path))


@mcp.tool
def change_case(
    file_path: str,
    transform: str,
    indices: list[int] | None = None,
    find: str | None = None,
    backup: bool = True,
) -> dict:
    """Change text case: upper | lower | title | sentence. Target: paragraph
    indices, or every occurrence of `find` (case-insensitive)."""
    return _edit(
        file_path,
        lambda pkg: _tx.change_case(pkg, transform, indices=indices, find=find),
        backup=backup,
    )


# --------------------------------------------------------------------- tables


@mcp.tool
def list_tables(file_path: str) -> list:
    """Tables with dimensions, merge flags, and header previews."""
    return _rd.list_tables(DocxPackage(file_path))


@mcp.tool
def get_table(file_path: str, table_index: int) -> dict:
    """Full table content + structure: cell texts, gridSpan/vMerge map, widths."""
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
    """Create a table from 2D data with single-line borders."""
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
    """Delete a whole table."""
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
    """Insert rows before row `at` (at=row count appends), copying structure
    and formatting from an existing row."""
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
    """Delete rows [start..end]. Vertical merges are re-rooted, not broken."""
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
    """Insert grid columns before position `at` (at=column count appends).
    Merge-aware: inserting inside a merged cell widens it."""
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
    single cells are removed, the grid stays consistent."""
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
    """Bulk cell writes in ONE call: [{row, cell, text}, ...]. Use for any
    multi-cell edit instead of per-cell calls. track: record as tracked
    changes by `author`. Documents open in Word are edited live."""
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
    """Write a 2D block of values starting at (origin_row, origin_cell)."""
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
    """Bulk cell formatting. targets: [{row, cell?}] ({row} alone = whole row).
    formatting: shading, bold, italic, alignment, valign, padding_pt."""
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
    """Merge a rectangle of cells (grid coordinates, inclusive)."""
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
    """Split a merged cell back into single cells (horizontal and vertical)."""
    return _edit(
        file_path,
        lambda pkg: _tb.unmerge_cells(pkg, table_index, row=row, cell=cell),
        backup=backup,
    )


@mcp.tool
def set_column_widths(
    file_path: str, table_index: int, widths_pt: list[float], backup: bool = True
) -> dict:
    """Set every grid column width in points."""
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
    BandedTable (alternating shading). Definitions injected when missing."""
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
    """Sort data rows by a column (CELL index). Refused when vertical merges
    make rows interdependent."""
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
    """Split a table into two at a row (that row starts the new table)."""
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
    """Repeat the first N rows as the table header on every page."""
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
    """Read a table nested inside a cell of a body-level table."""
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
    """Bulk cell writes inside a NESTED table ([{row, cell, text}])."""
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
    """Footnotes with ids, display positions, and text."""
    return _rd.list_footnotes(DocxPackage(file_path))


@mcp.tool
def list_endnotes(file_path: str) -> list:
    """Endnotes with ids, display positions, and text."""
    return _rd.list_endnotes(DocxPackage(file_path))


@mcp.tool
def add_footnote(
    file_path: str,
    anchor_text: str,
    note_text: str,
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Add a footnote anchored after `anchor_text`. Creates all note
    infrastructure (parts, styles, superscript) if the document has none."""
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
    """Add an endnote anchored after `anchor_text`."""
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
    """Rewrite a footnote's text, addressed by id or 1-based display position."""
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
    """Rewrite an endnote's text, addressed by id or display position."""
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
    """Delete a footnote: definition AND body reference (both, always)."""
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
    """Delete an endnote: definition AND body reference."""
    return _edit(
        file_path,
        lambda pkg: _nt.delete_note(
            pkg, "endnote", note_id=note_id, position=position
        ),
        backup=backup,
    )


@mcp.tool
def validate_notes(file_path: str) -> dict:
    """Footnote/endnote integrity. ok=no corruption; needs_cleanup=orphan
    definitions exist (run cleanup_orphan_notes)."""
    return _nt.validate_notes(DocxPackage(file_path))


@mcp.tool
def cleanup_orphan_notes(file_path: str, backup: bool = True) -> dict:
    """Remove footnote/endnote definitions no body reference points to.
    Content-deleting tools already do this automatically; use this for
    documents that arrived with orphans."""
    return _edit(file_path, lambda pkg: _nt.purge_orphans(pkg), backup=backup)


# --------------------------------------------------------- references & links


@mcp.tool
def add_bookmark(
    file_path: str, name: str, anchor_text: str, occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Bookmark a text range (target for cross-references)."""
    return _edit(
        file_path,
        lambda pkg: _fl.add_bookmark(
            pkg, name, anchor_text=anchor_text, occurrence=occurrence
        ),
        backup=backup,
    )


@mcp.tool
def list_bookmarks(file_path: str) -> list:
    """User-visible bookmarks in the document."""
    return _fl.list_bookmarks(DocxPackage(file_path))


@mcp.tool
def add_cross_reference(
    file_path: str,
    after_anchor: str,
    to_bookmark: str,
    kind: str = "page",
    backup: bool = True,
) -> dict:
    """Insert a cross-reference after anchor text. kind: 'page' (page number)
    or 'text' (bookmarked text). Word computes the value on field update."""
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
    """Numbered caption (SEQ field): 'Table N: text' above/below a table or at
    an anchor. label: Table | Figure | Equation."""
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
    """Turn existing text into an external hyperlink."""
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
    """Insert a TOC (SDT-wrapped field). Page numbers appear after Word
    updates fields: automatically on next open (update_on_open) or immediately
    via com_refresh_fields."""
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
    """All TOC-family fields (main TOC, List of Tables/Figures) and their
    cached entries."""
    return _tc.read_toc(DocxPackage(file_path))


@mcp.tool
def delete_toc(file_path: str, which: int = 0, backup: bool = True) -> dict:
    """Delete one TOC-family field by index (see read_toc's tocs list)."""
    return _edit(
        file_path, lambda pkg: _tc.delete_toc(pkg, which=which), backup=backup
    )


@mcp.tool
def set_update_fields_flag(
    file_path: str, on: bool = True, backup: bool = True
) -> dict:
    """Toggle 'update all fields on next open' (one-shot; Word clears it)."""
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
    entries. label: Table | Figure | Equation. Default title 'List of Xs'."""
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
    """Convert footnotes<->endnotes. direction: footnotes_to_endnotes |
    endnotes_to_footnotes. Give note_id or position for ONE note, neither
    for ALL. Word renumbers automatically."""
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
    """Add a Word-native bibliography source. tag = unique citation key;
    source_type: JournalArticle | Book | BookSection | Report | InternetSite
    | ...; authors/editors: [{last, first, middle?}] or [{corporate}].
    Cite with insert_citation; render the list with insert_bibliography."""
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
    """Bibliography sources in the document source store."""
    return _bib.list_sources(DocxPackage(file_path))


@mcp.tool
def delete_source(file_path: str, tag: str, force: bool = False, backup: bool = True) -> dict:
    """Delete a bibliography source (refused while cited unless force)."""
    return _edit(
        file_path, lambda pkg: _bib.delete_source(pkg, tag, force=force),
        backup=backup,
    )


@mcp.tool
def set_bibliography_style(file_path: str, style: str, backup: bool = True) -> dict:
    """Citation/bibliography style: APA | Chicago | MLA | IEEE | Turabian |
    Harvard - Anglia | GB7714 | GOST - Name Sort | GOST - Title Sort |
    ISO 690 - First Element and Date | ISO 690 - Numerical Reference | SIST02."""
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
    """Insert a CITATION field for a stored source after anchor_text. Renders
    in the document citation style on field update (placeholder until then)."""
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
    """Insert the BIBLIOGRAPHY field: Word generates the full styled
    reference list from the source store on field update."""
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
    """Mark a location for the index (invisible XE field). see='other entry'
    creates a cross-reference instead of a page number."""
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
    """All XE index entries marked in the document."""
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
    (generated when Word updates fields)."""
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
    """Move a heading and its ENTIRE section (content + tables until the next
    same-or-higher heading) to a new location. Refuses moves that would cut
    fields or section breaks."""
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
    """Headings with each section's element count (planning aid for
    move_section)."""
    return _sx.list_section_blocks(DocxPackage(file_path))


@mcp.tool
def apply_template(
    file_path: str,
    reference_path: str,
    include_page_geometry: bool = True,
    backup: bool = True,
) -> dict:
    """Restyle this document to match a reference document: styles, theme,
    fonts, layout settings, optionally page geometry. Content untouched;
    style references remapped by name, unmatched styles preserved."""
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
    """Core document metadata (File > Info): title, author, subject,
    keywords, category, comments."""
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
    formatting control."""
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
    """Apply a character style to a text range."""
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
    """Accessibility alt text for an image (list_images index)."""
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
    """Text watermark behind the text on every page (all header parts;
    compatible with Word's own Remove Watermark command)."""
    return _edit(
        file_path,
        lambda pkg: _fu.add_watermark(
            pkg, text, color=color, opacity=opacity, diagonal=diagonal
        ),
        backup=backup,
    )


@mcp.tool
def remove_watermark(file_path: str, backup: bool = True) -> dict:
    """Remove watermark shapes from every header."""
    return _edit(file_path, lambda pkg: _fu.remove_watermark(pkg), backup=backup)


@mcp.tool
def set_document_protection(
    file_path: str,
    edit: str = "trackedChanges",
    password: str | None = None,
    restrict_formatting: bool = False,
    backup: bool = True,
) -> dict:
    """Restrict editing: readOnly | comments | trackedChanges | forms.
    trackedChanges forces every edit by the recipient to be tracked — the
    send-to-committee mode. Password hashing is Word-compatible (SHA-512).
    NOT encryption."""
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
    """Lift the editing restriction."""
    return _edit(
        file_path, lambda pkg: _pr.remove_document_protection(pkg),
        backup=backup,
    )


@mcp.tool
def get_protection(file_path: str) -> dict:
    """Current protection state (mode, password, formatting restriction)."""
    return _pr.get_protection(DocxPackage(file_path))


# ---------------------------------------------------------- headers & layout


@mcp.tool
def get_headers_footers(file_path: str) -> dict:
    """Every header/footer part with its text and page-number-field flag."""
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
    """Set a section's header. ref_type: default | first | even."""
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
    """Set a section's footer, optionally with a page-number field."""
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
    """Add page numbers (PAGE field) to the header or footer. x_of_y renders
    'Page N of M'."""
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
    """Multi-column text layout for a section (equal widths, or explicit
    widths_pt per column; separator draws a line between columns)."""
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
    """Page-number FORMAT per section: lowerRoman for front
    matter, decimal restarting at 1 for the body, upperRoman, letters, etc."""
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
    """Manuscript line numbering for a section (journal submissions).
    restart: continuous | newPage | newSection. remove=True clears it."""
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
    """Sections with page size, orientation, margins, header/footer refs."""
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
    """Page size, orientation, margins for one section."""
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
    """Insert a section break: nextPage | continuous | evenPage | oddPage."""
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
    """Insert an inline image (PNG/JPEG/GIF/BMP/TIFF); aspect ratio kept."""
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
    """Images with sizes and media targets."""
    return _md.list_images(DocxPackage(file_path))


@mcp.tool
def resize_image(
    file_path: str, image_index: int, width_pt: float, backup: bool = True
) -> dict:
    """Resize an inline image (aspect ratio kept)."""
    return _edit(
        file_path,
        lambda pkg: _md.resize_image(pkg, image_index, width_pt=width_pt),
        backup=backup,
    )


@mcp.tool
def replace_image(
    file_path: str, image_index: int, new_image_path: str, backup: bool = True
) -> dict:
    """Swap an image's file, keeping placement and display size."""
    return _edit(
        file_path,
        lambda pkg: _md.replace_image(pkg, image_index, new_image_path),
        backup=backup,
    )


# ------------------------------------------------------------------- comments


@mcp.tool
def get_comments(file_path: str, author: str | None = None) -> list:
    """Comments with authors, anchored text, threading, resolved state."""
    return _rd.get_comments(DocxPackage(file_path), author=author)


@mcp.tool
def add_comment(
    file_path: str,
    anchor_text: str,
    text: str,
    author: str = "Claude",
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Comment on a text range (threaded-comment infrastructure included)."""
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
    """Threaded reply to an existing comment."""
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
    """Mark a comment thread resolved (or reopen with done=False)."""
    return _edit(
        file_path,
        lambda pkg: _cm.resolve_comment(pkg, comment_id=comment_id, done=done),
        backup=backup,
    )


@mcp.tool
def delete_comment(file_path: str, comment_id: str, backup: bool = True) -> dict:
    """Delete a comment and its replies, including all body markers."""
    return _edit(
        file_path,
        lambda pkg: _cm.delete_comment(pkg, comment_id=comment_id),
        backup=backup,
    )


# ------------------------------------------------------------------ revisions


@mcp.tool
def get_tracked_changes(file_path: str, author: str | None = None) -> list:
    """Tracked changes (insertions, deletions, moves, format changes)."""
    return _rd.get_tracked_changes(DocxPackage(file_path), author=author)


@mcp.tool
def revision_summary(file_path: str) -> dict:
    """Revision counts by author and type."""
    return _rd.revision_summary(DocxPackage(file_path))


@mcp.tool
def accept_revisions(
    file_path: str, author: str | None = None, backup: bool = True
) -> dict:
    """Accept tracked changes — all, or one author's only."""
    return _edit(
        file_path, lambda pkg: _rv.accept_revisions(pkg, author=author),
        backup=backup,
    )


@mcp.tool
def reject_revisions(
    file_path: str, author: str | None = None, backup: bool = True
) -> dict:
    """Reject tracked changes — all, or one author's only."""
    return _edit(
        file_path, lambda pkg: _rv.reject_revisions(pkg, author=author),
        backup=backup,
    )


# ---------------------------------------------------------------- COM (Word)


@mcp.tool
def com_word_status() -> dict:
    """Is Word running and responsive, and which documents does it have open
    (with per-document dirty/autosave state)? interactive_state:
    ready | busy (dialog open) | blocked (long operation) | not_running."""
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
    """Insert text at the USER'S CURSOR in the open document (main text
    only). The cursor position is read once; the selection itself is never
    touched. newline: end the insertion with a paragraph break."""
    from .com import live_ops as _lo

    return _lo.insert_at_cursor(file_path, text, newline=newline)


@mcp.tool
def live_scroll_to(
    file_path: str,
    find: str | None = None,
    paragraph_index: int | None = None,
) -> dict:
    """Scroll the user's Word window to show a location ("here is what I
    changed") WITHOUT selecting anything or moving their cursor. Target by
    find text or body paragraph index."""
    from .com import live_ops as _lo

    return _lo.scroll_to(file_path, find=find, paragraph_index=paragraph_index)


@mcp.tool
def live_set_track_changes(file_path: str, enabled: bool) -> dict:
    """Turn track changes on/off on the OPEN document, persistently (this is
    a deliberate state change, unlike the auto-restored track flags on edit
    tools). Returns the previous state."""
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
    """Update every field (TOC page numbers, PAGEREF, SEQ) immediately via an
    invisible Word instance. Use after insert_toc for instant page numbers."""
    from .com import bridge

    return bridge.refresh_fields(file_path)


@mcp.tool
def com_export_pdf(file_path: str, pdf_path: str | None = None) -> dict:
    """Export to PDF via Word (full fidelity: fields, footnotes, TOC)."""
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
    """Merge documents in order into ONE file with full fidelity (styles,
    footnotes, numbering) — e.g. chapters into one manuscript. Section
    breaks between parts keep per-chapter headers/numbering possible."""
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
    """Combine two documents' TRACKED CHANGES into one (two reviewers' edits
    of the same draft, both attributions preserved). Compare diffs content;
    combine merges revisions."""
    from .com import bridge

    return bridge.combine_documents(original_path, revised_path, output_path)


@mcp.tool
def com_save_open_document(file_path: str) -> dict:
    """Tell the user's running Word to SAVE a document it has open, so
    file-based tools can read the current state."""
    from .com import bridge

    return bridge.save_open_document(file_path)


@mcp.tool
def com_close_open_document(file_path: str, save: bool = True) -> dict:
    """Tell the user's running Word to CLOSE an open document (saving by
    default), releasing the lock for file-based editing."""
    from .com import bridge

    return bridge.close_open_document(file_path, save=save)


@mcp.tool
def com_proofing_errors(file_path: str, limit: int = 100) -> dict:
    """Word's own spelling/grammar error lists with context (review aid;
    includes proper nouns it does not recognize)."""
    from .com import bridge

    return bridge.proofing_errors(file_path, limit=limit)


@mcp.tool
def com_readability_statistics(file_path: str) -> dict:
    """Word's readability statistics (Flesch, grade level, counts)."""
    from .com import bridge

    return bridge.readability_statistics(file_path)


@mcp.tool
def com_save_with_password(
    file_path: str, password: str, output_path: str | None = None
) -> dict:
    """Save an ENCRYPTED copy requiring a password to open (real encryption,
    unlike document protection)."""
    from .com import bridge

    return bridge.save_with_password(
        file_path, output_path, password=password
    )


@mcp.tool
def com_validate_opens_clean(file_path: str) -> dict:
    """Definitive corruption check: open in invisible Word, report clean/fail."""
    from .com import bridge

    return bridge.validate_opens_clean(file_path)



# ==================================================== workflow suites


@mcp.tool
def validate_cross_references(file_path: str) -> dict:
    """Check every REF/PAGEREF cross-reference against the bookmarks that
    actually exist: broken refs (target bookmark missing — renders as an
    error in Word), bookmarks nothing references (informational), and
    plain-text references like "see Figure 3" that match no caption or
    heading number (heuristic review candidates). Paragraph indices included
    for every finding."""
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
    """One-call submission prep: accept all tracked changes (every author),
    delete all comments including the comments-family parts, and scrub
    identifying metadata (author, last-modified-by, company; title kept
    unless keep_title=False). Content is never removed — footnotes, fields,
    and citations all stay. Refuses protected documents rather than
    delivering a half-clean file. Returns exactly what was done, what was
    deliberately left (rsids), and what remains in the document."""
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
    intact. Broken pairs are reported loudly — they disconnect the citation
    from the reference manager."""
    return _rf.list_reference_fields(DocxPackage(file_path))


@mcp.tool
def check_reference_field_integrity(file_path: str) -> dict:
    """Post-edit health check for reference-manager citations: counts by
    manager and kind plus an ok flag that goes False on any broken or stray
    field marker. Run after editing a document that contains Zotero, EndNote,
    or Mendeley citations to confirm the edit broke nothing."""
    return _rf.check_reference_field_integrity(DocxPackage(file_path))


@mcp.tool
def check_template_compliance(file_path: str, rules: dict) -> dict:
    """Validate the document against a formatting ruleset (university
    dissertation guide, journal style sheet). Every rule key is optional;
    unknown keys are rejected with the allowed list. Example ruleset:

    {"page": {"margins_pt": {"top": 72, "bottom": 72, "left": 90, "right": 72},
              "tolerance_pt": 1, "size": "letter", "orientation": "portrait"},
     "fonts": {"allowed": ["Times New Roman"], "body_size_pt": 12},
     "line_spacing": {"body": 2.0},
     "headings": {"max_skip": 0, "required_first_level": 1},
     "page_numbering": [{"section": 0, "format": "lowerRoman"},
                        {"section": 1, "format": "decimal", "restart_at": 1}],
     "required_headings_in_order": ["Abstract", "Acknowledgments"]}

    Returns {compliant, violations: [{rule, expected, found, location,
    severity}], unverified, rules_checked}. Fonts/sizes/spacing are resolved
    through explicit formatting, the style basedOn chain, and docDefaults;
    theme-indirected fonts are listed under "unverified" (reported, never
    guessed). Read-only — the file is not modified."""
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
    """Accessibility audit: heading hierarchy (skipped levels, empty headings,
    no Heading 1), images missing alt text, tables whose first row is not a
    repeating header, low-contrast text (explicit color vs explicit run
    background below WCAG 4.5:1 — skipped, not guessed, when either side is
    absent or automatic), missing document title, and generic hyperlink text
    ("click here", "here", "link"). Each finding carries a location and a fix
    hint; the summary has per-category counts and a pass flag. Read-only."""
    return _ac.audit_accessibility(DocxPackage(file_path))


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
    """Fill a template IN PLACE: replace every {{name}} with data[name]
    (safe across fragmented runs, first run's formatting kept) and set
    MERGEFIELDs to their values as plain text. missing: 'error' refuses and
    changes nothing if the template needs a key data lacks; 'skip' leaves
    those markers; 'empty' fills them with empty strings. To produce copies
    instead of editing in place, use mail_merge."""
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
    """Apply the same operations to MANY documents. Each operation:
    {'tool': name, 'params': {...}} with the tool's normal parameters minus
    file_path (e.g. {'tool': 'search_and_replace', 'params': {'replacements':
    [{'find': 'a', 'replace': 'b'}]}}). Allowed tools: search_and_replace,
    insert_paragraphs, delete_paragraphs, replace_paragraph_text,
    format_text, set_paragraph_format, apply_style, set_header, set_footer,
    add_page_numbers, set_page_number_format, set_document_properties,
    set_cells, add_watermark, remove_watermark. Per file: ALL operations,
    then one backup and one atomic save; if any operation fails that file is
    left untouched, and stop_on_error=True also skips the remaining files
    (already-saved files keep their changes and are reported)."""
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
    """Set form-field values by name (legacy) or tag/alias (content
    control): text fields take strings, checkboxes booleans, dropdowns only
    values from their options (refused otherwise). Duplicate field names are
    refused with locations. missing: 'error' refuses (changing nothing) if a
    key matches no field; 'skip' ignores such keys and reports them."""
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
    """Set up standard long-document front matter in one call: the requested
    sequence is inserted at the START of the document (existing content
    becomes the body), page breaks separate front-matter pages, and one
    section break separates front matter from body so page numbering can
    switch (front lowerRoman, body decimal restarting at 1, by default).

    spec = {"sections": [{"kind": "title_page", "lines": [...]},
                         {"kind": "blank_or_copyright", "lines": [...]},
                         {"kind": "abstract", "title": "Abstract", "text": "..."},
                         {"kind": "toc"}, {"kind": "list_of_figures"},
                         {"kind": "list_of_tables"}],
            "page_numbering": {"front": "lowerRoman", "body": "decimal",
                               "body_restart_at": 1}}

    Title-page lines are centered; the abstract gets a Heading-styled title.
    Refuses when front matter appears to exist already (document starts with
    a TOC, or a section already uses lowerRoman numbering) unless spec has
    "force": true. Reports exactly what was inserted and which section
    indices carry which numbering."""
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
    field referencing the Heading style of `level` — the standard Word
    mechanism, evaluated per page, so it updates automatically and needs no
    per-chapter section breaks. include_number adds the heading number
    (STYLEREF \\n) before the title. scope 'auto' targets every section
    containing body headings of that level; pass a list of section indices
    to override. first_page_blank sets titlePg so section-opening pages show
    no header. Watermarks in rewritten headers are preserved; other existing
    header content is replaced and reported. Reports the sections touched
    and the exact field codes written."""
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
    """One-call structural health report, read-only: content-type coverage,
    dangling relationships and orphan parts, field begin/end balance per
    story part, footnote/endnote integrity, references to undefined styles
    and numbering, content-control and bookmark sanity, duplicate revision
    ids, images whose targets are missing, broken cross-references, and a
    per-part size profile. Never fails on a weird-but-openable document —
    every check degrades to a reported problem. ok=false only for problems
    that render broken or lose content in Word."""
    return _dg.diagnose_document(DocxPackage(file_path))


@mcp.tool
def redact_text(
    file_path: str,
    targets: list[dict],
    replacement: str = "[REDACTED]",
    scope: str = "all",
    backup: bool = True,
) -> dict:
    """Permanently REMOVE matched text from the document (true redaction,
    not black highlighting — the characters are replaced in the XML). Each
    target: {find, regex?}. Runmap-safe: secrets fragmented across Word's
    split runs are found and removed as one match. Scrubs body incl. tables,
    headers/footers, footnotes/endnotes (per scope: body | headers |
    footnotes | all), and ALWAYS: comment text, document properties,
    hyperlink display text/tooltips/URL targets, field instruction text,
    cached field results, and tracked-change deleted text.

    The result reports per-location-class counts, what was scrubbed, what
    was NOT examined (embedded images, charts, OLE objects — text drawn in
    an image is NOT redacted), and verified_clean from a full post-redaction
    re-scan of every XML part. Zero-width regexes and empty finds are
    refused before anything is touched; any error leaves the file unchanged.
    Irreversible in the saved file — the auto-backup is the undo."""
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
    """Legal-document defined-terms audit. Finds definitions («"Term"
    means», «"Term" shall mean», «(the "Term")», «(each, a "Term")» —
    defaults overridable via definition_patterns, each regex capturing the
    term as group 1) and reports, with paragraph indices: defined_never_used,
    defined_multiple_times, first_use_before_definition, and
    used_never_defined — a HEURISTIC list of capitalized recurring terms
    with no definition, filtered so words capitalized only at sentence
    starts are not flagged; treat those as review candidates. Body-level
    paragraphs only (stated in the result notes). Read-only."""
    return _dt.check_defined_terms(
        DocxPackage(file_path), definition_patterns=definition_patterns
    )


@mcp.tool
def com_import_pdf(pdf_path: str, output_path: str | None = None) -> dict:
    """Convert a PDF to .docx via Word's built-in PDF reflow, in a dedicated
    invisible Word instance. Output defaults next to the PDF with a .docx
    extension; an existing output file is refused. The produced .docx is
    validated by a full package round-trip. Reflow quality depends on the
    PDF: text-based PDFs convert well, complex layouts may reflow
    imperfectly, and scanned image PDFs yield little or no text (Word does
    not OCR) — a near-zero word count triggers an explicit warning."""
    from word_mcp.com import convert  # when pasted: from .com import convert

    return convert.import_pdf(pdf_path, output_path)



def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
