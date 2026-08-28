<!-- mcp-name: io.github.nometalalchemist/kitchensink4word -->

# 🚰 KitchenSink4Word

[![Tests](https://github.com/nometalalchemist/KitchenSink4Word/actions/workflows/tests.yml/badge.svg)](https://github.com/nometalalchemist/KitchenSink4Word/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4word)](https://pypi.org/project/kitchensink4word/)
[![License: PolyForm NC](https://img.shields.io/badge/license-PolyForm%20Noncommercial-blue)](LICENSE)

**Everything plus the kitchen sink for Microsoft Word.** The most complete
Word (.docx) MCP server available — 170 tools, engineered not to corrupt, stress-tested against long, heavily formatted real-world documents. Now with live editing: documents open in Word are edited in place, visibly, with each tool call landing as a single Ctrl+Z step.

> **The origin story:** an AI agent once needed *fifteen minutes* to edit
> twenty table cells in a Word document, because no existing Word MCP could
> delete a table column, bulk-edit cells, manage footnotes, AND insert a TOC.
> So instead of installing three mediocre servers, this one got built in a
> day — and now your agents won't have that problem.

## Why this one (the honest comparison)

Every public Word MCP server was surveyed before building this (August 2026):

| Capability | KitchenSink4Word | GongRzhe Office-Word (2.1k★, archived) | word-mcp-live (195★) | SecurityRonin docx-mcp (43★) |
|---|---|---|---|---|
| Live editing while the doc is open in Word | ✅ cursor-safe, one Ctrl+Z per call | ❌ | ✅ | ❌ |
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

**170 tools** across: equations (LaTeX → Word math), Zotero library citations, publication style conversion (8 styles, beta), review-cycle analytics, workflow suites (mail merge, batch operations, redaction, compliance/accessibility audits, submission prep, front matter, diagnostics), text and formatting, tables (including merge-aware column
insert/delete and one-call bulk cell edits), footnotes/endnotes (full CRUD +
footnote↔endnote conversion), TOC and caption lists, headers/footers/sections,
images, bulleted/numbered lists, threaded comments, tracked changes (read,
accept/reject by author, AND write edits as tracked changes), plus
Word-COM-backed document compare, field refresh, PDF export, and open-clean
validation on Windows.

## What's new in v1.4 (2026-08-28)

**The workflow tier** — 24 new tools that turn document primitives into
complete jobs:

- **Document production**: `mail_merge` (template + CSV/JSON → one document
  per row, atomic collision refusal), `fill_template` ({{placeholders}} and
  MERGEFIELDs, safe across fragmented runs), `batch_apply` (the same edits
  across dozens of files, all-or-nothing per file), form-field tooling
  (legacy form fields AND content controls: list, fill, validate
  completeness).
- **Pre-flight checks**: `check_template_compliance` (validate margins,
  fonts, spacing, heading structure, page-numbering and required sections
  against a ruleset you write from any formatting guide),
  `check_brand_compliance`, `audit_accessibility` (heading hierarchy, alt
  text, table headers, contrast, link text), `check_image_resolution`
  (effective DPI vs a print threshold), `validate_cross_references` (broken
  REF/PAGEREF targets plus "see Figure 3" text references that match
  nothing), `validate_captions`, `check_defined_terms` (legal drafts:
  defined-but-unused, used-but-undefined, defined twice, used before
  defined).
- **Finishing moves**: `prepare_for_submission` (accept all revisions,
  strip comments and identifying metadata in one call — content untouched,
  idempotent), `redact_text` (TRUE removal from text, tables, headers,
  footnotes, comments, properties, hyperlinks, field codes and tracked
  deletions, with a verification re-scan and an explicit list of what was
  NOT examined), `assemble_front_matter` (title page through TOC with
  roman/arabic numbering switch in one call), `setup_chapter_headers`
  (chapter-title headers via STYLEREF), `diagnose_document` (13 structural
  health checks that never crash on a broken file), `com_import_pdf`
  (PDF → editable .docx via Word's own converter), reference-manager field
  inventory + integrity checking (Zotero, EndNote, Mendeley).

## What's new in v1.3 (2026-08-28)

**Live editing.** A document open in Word no longer has to be closed first —
the high-value tools route to a live COM layer automatically when the file is
locked (or explicitly with `live="force"`; `live="off"` restores the old
refusal):

- Edits appear in the Word window immediately; nothing touches the disk until
  the USER saves (or an explicit save tool is called). Unsaved work is never
  auto-saved or auto-closed.
- **One tool call = one Ctrl+Z step.** Every mutating call is wrapped in a
  custom undo record, so a bulk replace of 40 matches reverts with a single
  undo.
- **The cursor is sacred.** All addressing is Range-based; the user's
  selection, scroll position, and view are never touched. (One deliberate
  exception: `live_scroll_to` brings a location into view on request —
  without selecting it.)
- Routed tools: `search_and_replace`, `insert_paragraphs`,
  `delete_paragraphs`, `set_cells`, `format_text`, `get_text`, `find_text`,
  `get_outline`, `get_document_info`. Live-only tools:
  `live_insert_at_cursor`, `live_scroll_to`, `live_set_track_changes`,
  `word_live_repair` (recovery if a crashed client left Word in a bad state).
- Honest state reporting on every live result: document dirty, AutoSave
  state, read-only, enforced tracking under protection, and per-item skip
  counts when matches sit inside field results (Word regenerates those, so
  editing them is refused rather than faked).
- Protection-restricted documents get typed refusals up front; documents in
  Protected View are named as such.
- Reference-manager field codes (Zotero and friends) survive live and
  file-based edits around them — verified by dedicated tests.

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

Two minutes, start to editing:

```
pip install kitchensink4word
claude mcp add word -s user -- kitchensink4word
```

For Claude Desktop / Cursor / any MCP client, point the server command at the
installed `kitchensink4word` (or `word-mcp`) executable:

```json
{"mcpServers": {"word": {"command": "kitchensink4word"}}}
```

From source instead:

```
git clone https://github.com/nometalalchemist/KitchenSink4Word
cd KitchenSink4Word
python -m venv .venv
.venv\Scripts\pip install -e .
claude mcp add word -s user -- <absolute-path>\.venv\Scripts\word-mcp.exe
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

344 tests, running everywhere: the suite was developed against a private
corpus of real-world documents (book-length chapters, a document with 171
footnotes, a manuscript with 126 tracked changes and reviewer comments), and
CI auto-generates **structurally equivalent synthetic stand-ins**
(`tests/make_corpus.py`) so the full suite runs on any machine — including
yours and every pull request. Local real documents, when present, take
precedence. `tests/word_validator.py`
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

## Maturity — what the version number does and does not claim

This project moves fast and is honest about what backs it. What the test
record covers: every release passes the full suite (626 tests) plus
dedicated adversarial rounds through the raw MCP transport (~2,500 calls
to date), against a corpus of long, heavily formatted real-world documents
— with zero corruption across all of it. What it does not yet cover: other
machines, Word builds older than current Microsoft 365, non-English Word
installs (some tools reference styles by localized display name), RTL
scripts, and the diversity of documents only real users bring. The safety
net while the tool earns that mileage is structural: automatic timestamped
backups before every mutation and atomic validated saves, so a bad outcome
is a restore, not a loss. If something misbehaves on your documents,
[an issue](https://github.com/nometalalchemist/KitchenSink4Word/issues)
with the symptom (never the document itself, unless it contains nothing
private) is the most valuable thing you can send.

**Beta-labeled tools** — heuristic by nature; review their flagged-items
list rather than trusting silently: `convert_citation_style`,
`anonymize_for_review`, `check_defined_terms`, and the text-reference
scan inside `validate_cross_references`. Each returns an explicit list of
what it could not confidently handle.

## Known limits

- Live regex replacements skip matches positioned after complex fields in a
  story (COM character offsets drift there); the skip count is reported and a
  literal find still works. Live `set_cells` refuses vertically merged tables
  (the file-based tool is merge-aware — close the doc for those).
- Live tracked-change attribution is best-effort: Word signed into an Office
  account attributes revisions to that account; results report the effective
  author honestly.
- TOC/caption-list page numbers require a field update: automatic on next Word
  open (`update_on_open`), or immediate via `com_refresh_fields`.

## License

**PolyForm Noncommercial 1.0.0** — free for personal, academic, research, and
any other noncommercial use. Commercial use requires a separate license **before** use — open an issue
to arrange one.
