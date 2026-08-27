<!-- mcp-name: io.github.nometalalchemist/kitchensink4word -->

# 🚰 KitchenSink4Word

**Everything plus the kitchen sink for Microsoft Word.** The most complete
Word (.docx) MCP server available — 119 tools, engineered not to corrupt, stress-tested against long, heavily formatted real-world documents.

> **The origin story:** an AI agent once needed *fifteen minutes* to edit
> twenty table cells in a Word document, because no existing Word MCP could
> delete a table column, bulk-edit cells, manage footnotes, AND insert a TOC.
> So instead of installing three mediocre servers, this one got built in a
> day — and now your agents won't have that problem.

## Why this one (the honest comparison)

Every public Word MCP server was surveyed before building this (August 2026):

| Capability | KitchenSink4Word | GongRzhe Office-Word (2.1k★, archived) | word-mcp-live (195★) | SecurityRonin docx-mcp (43★) |
|---|---|---|---|---|
| Table column insert/delete | ✅ merge-aware | ❌ | ❌ | ❌ |
| Bulk cell edits (one call) | ✅ | ❌ | ❌ | ❌ |
| Cell merge/unmerge | ✅ | merge only | ❌ | ❌ |
| Footnotes AND endnotes CRUD | ✅ + conversion | add only | ❌ | ✅ |
| TOC insert + refresh | ✅ | ❌ | ❌ | ❌ |
| Native citations/bibliography | ✅ 12 styles | ❌ | ❌ | ❌ |
| Index generation | ✅ | ❌ | ❌ | ❌ |
| Tracked-change WRITING | ✅ | ❌ | ✅ | ✅ |
| Accept/reject by author | ✅ | ❌ | partial | ✅ |
| Document compare + combine | ✅ Word-native | ❌ | ❌ | buggy |
| Watermarks / protection / line numbers | ✅ | protect only | ❌ | ❌ |
| Section moving / template transfer | ✅ | ❌ | ❌ | ❌ |
| Atomic saves + auto-backup | ✅ | ❌ | ❌ | ❌ |

**119 tools** across: text and formatting, tables (including merge-aware column
insert/delete and one-call bulk cell edits), footnotes/endnotes (full CRUD +
footnote↔endnote conversion), TOC and caption lists, headers/footers/sections,
images, bulleted/numbered lists, threaded comments, tracked changes (read,
accept/reject by author, AND write edits as tracked changes), plus
Word-COM-backed document compare, field refresh, PDF export, and open-clean
validation on Windows.

## What's new in v1.2 (2026-08-28)

- **Word-native citations & bibliography**: structured source store, CITATION
  fields with page/suppress switches, BIBLIOGRAPHY field, 12 selectable styles
  (APA, Chicago, MLA, IEEE, ...) — verified end-to-end: Word renders the
  fields in the selected style.
- **Index** (XE entries with nesting and see-references + INDEX field) and
  **caption lists** (List of Tables/Figures).
- **Long-document layout kit**: roman-numeral page formats, line numbering,
  watermarks (compatible with Word's Remove Watermark), multi-column
  sections, Page X of Y.
- **Document protection** with Word-compatible SHA-512 password hashing
  (verified byte-for-byte against Word's own output); the trackedChanges mode
  forces recipients' edits to be tracked.
- **Structure ops**: move_section (a heading plus its entire section, tables
  included), template transfer (restyle to match a reference document with
  name-based style remapping), custom styles, character styles, document
  properties, image alt text.
- **Table completions**: gridBefore/gridAfter rows, named table styles, sort,
  split, header-row repeat, nested-table read/write, cell text direction.
- **Formatting long-tail**: small caps, hidden text, character spacing,
  kerning, CJK language tagging, tab stops with leaders, paragraph
  borders/shading, widow control, change-case.
- **Analysis**: per-section word counts, APA citation-parity checking, Word's
  proofing errors and readability statistics, chapter merge, reviewer
  combine, password-encrypted saving.

## What makes it different

- **Merge-aware table column operations.** `delete_columns` / `insert_columns`
  work correctly through horizontally and vertically merged cells (gridSpan
  shrinks, vMerge chains re-root). At the time of writing, no other public
  Word MCP has this.
- **Bulk-first API.** Editing 20 cells is ONE `set_cells` call with a payload,
  not 20 round-trips.
- **Tracked-change writing.** `track=True, author="Jane"` on replace/insert/
  delete/cell tools produces real Word revisions the recipient can
  accept/reject — proven round-trip against the server's own revision engine.
- **Document compare.** `com_compare_documents` produces a Word-native redline
  between two versions of a document.
- **Never corrupts.** Atomic saves (temp file → structural validation →
  replace), automatic timestamped backups before every mutation, byte-identical
  passthrough of anything not being edited (equations, textboxes, content
  controls survive untouched), and clean typed errors — a file open in Word is
  refused with a message, not a hang.
- **Fragmented-run safe.** Find/replace works across Word's arbitrarily split
  runs while preserving per-character formatting, with a ReDoS timeout guard on
  user regex and an optional `max_replacements` blast-radius limit.

## Requirements

- Windows (COM tools require Microsoft Word installed; all other tools are
  pure file manipulation and work without Word)
- Python 3.12+ (developed on 3.14)

## Install

```
git clone https://github.com/nometalalchemist/KitchenSink4Word
cd KitchenSink4Word
python -m venv .venv
.venv\Scripts\pip install -e .
```

Register with Claude Code:

```
claude mcp add word -s user -- <absolute-path-to>\word-mcp\.venv\Scripts\word-mcp.exe
```

## Safety model

- Every mutating tool takes `file_path` and writes a `<name>.bak-<timestamp>.docx`
  beside it before the first change (`backup=False` to skip).
- Saves are atomic and validated; a failed operation leaves the original
  byte-identical.
- Deleting content that carries footnote references automatically removes the
  now-orphaned definitions; `validate_document` / `validate_notes` report
  integrity in both directions.
- Paragraph deletion refuses ranges that would cut a field (TOC, PAGEREF) in
  half or silently swallow a section break.

## Testing

335 tests. The full suite runs against a corpus of real-world documents
(book-length chapters, a document with 171 footnotes, a manuscript with 126
tracked changes and reviewer comments) that is **private and not shipped** —
corpus-dependent tests skip cleanly with an explanation. See
`tests/conftest.py` for how to supply your own corpus. `tests/word_validator.py`
opens outputs in invisible Word and fails on any repair prompt — the definitive
corruption check.

Development history: built with Claude Code in a single day (2026-08-27),
tested through three rounds — unit gates per phase, an edge-case session
(89 calls), and two "insane mode" rounds through the raw MCP stdio transport (scale
torture, pathological merge topologies, Unicode/schema fuzzing, ReDoS,
Word-lock lifecycle, COM leak checks), and a dedicated cross-feature
interaction bug-hunt. All findings fixed with regression tests. `research/` documents the OOXML
algorithms and pitfalls the implementation is built on, with attribution to
the MIT-licensed reference implementations studied.

## Known limits

- No live editing of documents open in Word (clean refusal; COM tools work on
  saved files).
- TOC/caption-list page numbers require a field update: automatic on next Word
  open (`update_on_open`), or immediate via `com_refresh_fields`.

## License

**PolyForm Noncommercial 1.0.0** — free for personal, academic, research, and
any other noncommercial use. Using it to make money? That's fine too — just
ask first: open an issue to arrange a commercial license.
