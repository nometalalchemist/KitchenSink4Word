<!-- mcp-name: io.github.nometalalchemist/kitchensink4word -->

# 🚰 KitchenSink4Word

[![Tests](https://github.com/nometalalchemist/KitchenSink4Word/actions/workflows/tests.yml/badge.svg)](https://github.com/nometalalchemist/KitchenSink4Word/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4word)](https://pypi.org/project/kitchensink4word/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](LICENSE)
[![Listed on mcpservers.org](https://mcpservers.org/badge.svg)](https://mcpservers.org/servers/nometalalchemist/kitchensink4word)

[Landing page](https://nometalalchemist.github.io/KitchenSink4Word/) · [llms.txt](https://nometalalchemist.github.io/KitchenSink4Word/llms.txt) (machine-readable capability manifest for agents and LLM crawlers)

**Everything plus the kitchen sink for Microsoft Word.** The most complete
Word (.docx) MCP server available: **182 document operations** across **108
tools**, one consistent grammar, engineered not to corrupt and stress-tested
against long, heavily formatted real-world documents. Live editing included:
documents open in Word are edited in place, visibly, with each tool call
landing as a single Ctrl+Z step.

> ### ⚠️ v2.0 is a breaking change
> Every v1.x tool name changed. The 189-tool v1.6 surface was rebuilt as a
> consolidated set of 108 tools that cover every prior capability under one
> grammar. If you are upgrading from v1.x, read the
> **[migration guide](docs/MIGRATION_V2.md)** first: it maps every old tool
> name to its v2 home, and `get_workflows("migrate-from-v1")` returns the
> same map in-session. New installs need nothing extra.

## Two numbers that matter

- **182 document operations, 108 tools.** The operation count went up and the
  tool count came down on purpose. v1 spread similar jobs across many
  competing names; v2 gives each concept exactly one name built from a small
  verb table (`insert_`, `set_`, `manage_`, `list_elements`, `validate`,
  `delete_element`), so an agent picks the right tool the first time and
  carries less schema to do it. Fewer tools, more reach.
- **Tiered loading: starts at about 7.5k tokens, scales to everything.** A
  fresh session loads the 28-tool lite core (about 7,500 tokens) and turns on
  capability packs only when a task needs them, with one `enable_tools` call.
  Load every pack and the full surface measures about 26,400 tokens, down from
  about 34,400 in v1.6: roughly 23% less for the whole sink, about 78% less at
  lite start. (All figures are script-measured; see
  [Context cost](#context-cost-measured) below.)

## The packs

Numbers below come straight from `scripts/measure_surface.py`, never
hand-counted. The lite core loads at startup; the seven packs load on demand.

| Pack | Tools | Approx tokens | What it carries |
|---|---:|---:|---|
| **lite** (startup) | 28 | ~7.5k | Everyday reading and editing: text, paragraphs, tables, cells, lists, find and replace, outline, document view, backups, workflow guide, pack toggles |
| references | 8 | ~2.4k | Word-native citations and bibliography, Zotero search and cite, parity checks, style conversion and detection |
| review | 9 | ~1.8k | Tracked changes (read, accept/reject, reports), threaded comments, structured diff, anonymize and deanonymize |
| academic | 23 | ~5.7k | Footnotes and endnotes, TOC, index, captions, cross-references, front matter, chapter headers, sections, styles, word counts, validation batteries, submission prep, accessibility |
| assembly | 7 | ~1.7k | Insert and split documents, move sections, copy tables across files, apply and fill templates, mail merge |
| media-forms | 16 | ~4.0k | Images, charts, equations, text boxes, hyperlinks, table structure and styling, form fields, content controls, field codes |
| com-live | 13 | ~2.1k | Drives a local Microsoft Word: PDF import/export, compare and combine, proofing, readability, field refresh, live editing of open documents |
| protection-io | 6 | ~1.2k | Document protection, watermarks, redaction with verification, table data import and export |
| **Full surface** | **110** | **~26.4k** | Everything (108 document tools plus `enable_tools` / `disable_tools`) |

## Quickstart: start lite, enable what you need

A session begins with the lite core. When a task needs more, the agent turns
on the pack by name:

```
enable_tools(["references"])          # citations, bibliography, Zotero
enable_tools(["academic", "review"])  # notes, TOC, tracked changes, comments
```

Lite-tool refusals name the pack and the exact `enable_tools` call to run, and
`get_workflows` recipes name the packs each workflow needs, so discovery is
built in. To skip tiering entirely, start the server with `KS4W_MODE=full` and
every tool is present from the first call.

## Why this one (the capability comparison)

Every public Word MCP server was surveyed before building this. The honest
comparison is about what each one can do, not how many names it has:

| Capability | KitchenSink4Word | GongRzhe Office-Word (2.1k★, archived) | word-mcp-live (195★) | SecurityRonin docx-mcp (43★) |
|---|---|---|---|---|
| Tiered context loading (lite core, packs on demand) | ✅ from ~7.5k tokens | ❌ | ❌ | ❌ |
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

Capability survey compiled from public repositories, documentation, and issue
trackers. Corrections welcome: [open an issue](https://github.com/nometalalchemist/KitchenSink4Word/issues).

## What the 182 operations cover

The full surface spans equations (LaTeX to Word math), native charts, document
assembly (chapter files into one manuscript), Zotero library citations,
publication style conversion (8 styles, beta), review-cycle analytics, and
workflow suites (mail merge, batch operations, redaction, compliance and
accessibility audits and fixes, submission prep, front matter, diagnostics),
plus text and formatting, tables (including merge-aware column insert/delete
and one-call bulk cell edits), footnotes and endnotes (full lifecycle plus
footnote and endnote conversion), TOC and caption lists, headers, footers, and
sections, images, bulleted and numbered lists, content controls and fields,
threaded comments, tracked changes (read, accept and reject by author, and
writing edits as tracked changes), and Word-COM-backed document compare, field
refresh, PDF export and import, and open-clean validation on Windows.

## What makes it different

- **Merge-aware table column operations.** `modify_table_structure` inserts
  and deletes columns correctly through horizontally and vertically merged
  cells (gridSpan shrinks, vMerge chains re-root). At the time of writing, no
  other public Word MCP has this.
- **Bulk-first API.** Editing 20 cells is ONE `set_cells` call with a payload,
  not 20 round-trips.
- **Tracked-change writing.** `track=True, author="Jane"` on edit tools
  produces real Word revisions the recipient can accept or reject, proven
  round-trip against the server's own revision engine.
- **Document compare.** `com_multi_document(action="compare")` produces a
  Word-native redline between two versions of a document.
- **Never corrupts.** Atomic saves (temp file, structural validation, replace),
  automatic slot backups before every mutation, byte-identical passthrough of
  anything not being edited (equations, textboxes, content controls survive
  untouched), and clean typed errors: a file open in Word is refused with a
  message, not a hang.
- **Fragmented-run safe.** Find and replace works across Word's arbitrarily
  split runs while preserving per-character formatting, with a ReDoS timeout
  guard on user regex and an optional blast-radius limit.

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

Claude Desktop one-click: download `kitchensink4word.mcpb` from the latest
release, then in Desktop use Settings > Extensions > Advanced settings >
Install extension and pick the file. Requires [uv](https://docs.astral.sh/uv/)
on your PATH (`pip install uv`), which the bundle uses to launch the server.

For other MCP clients, point the server command at the installed
`kitchensink4word` (or `word-mcp`) executable:

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

## Context cost (measured)

Almost no MCP server tells you what it costs to load. Here is the bill, from
`scripts/measure_surface.py`:

- **Lite start:** 28 tools, about 7,500 tokens, loaded when the session opens.
- **Full surface:** 110 tools (108 document tools plus the two pack toggles),
  about 26,400 tokens with every pack enabled.
- **Versus v1.6:** the old full surface was about 34,400 tokens. v2 is roughly
  23% smaller at full load and about 78% smaller at lite start.
- Clients that defer tool schemas until first use (for example Claude Code)
  pay close to zero until a tool is actually called.

## Safety model

- Every mutating tool takes `file_path` and rotates the current content into
  stable backup slots before the change (`backup=False` to skip the rotation;
  the atomic validated save always applies). Backups live in a hidden
  `.ks4w-backups/` folder next to the document, one subfolder per document,
  with exactly two slots: `prev.docx` (state before the most recent mutation)
  and `anchor.docx` (session start, rotating after 60+ minutes of idle).
  Storage stays bounded at roughly two copies per document no matter how many
  edits a session makes. `manage_backups` lists, restores (undoably: the
  pre-restore state rotates into `prev` first), and purges them, including
  leftover `*.bak-*` files from earlier schemes.
- Exclude `.ks4w-backups/` from cloud sync tools (OneDrive, Dropbox, Google
  Drive): the slots churn on every edit and sync clients can hold locks that
  slow saves down.
- Mutations of the same file are serialized (in-process and across server
  processes via an advisory lockfile), so parallel calls cannot clobber each
  other and response metadata reflects settled document state.
- COM/live mutations (documents open in Word) remain outside this backup
  system; Word's own AutoRecover covers the open document.
- Saves are atomic and validated; a failed operation leaves the original
  byte-identical.
- Deleting content that carries footnote references automatically removes the
  now-orphaned definitions; `validate` reports integrity in both directions.
- Paragraph deletion refuses ranges that would cut a field (TOC, PAGEREF) in
  half or silently swallow a section break.

### Sandboxing (opt-in)

Off by default: with nothing configured, the server behaves exactly as it
always has. Set the `KS4W_ALLOWED_ROOTS` environment variable to a list of
directories separated by the OS path separator (`;` on Windows, `:`
elsewhere), for example `C:\Users\me\Documents;D:\Work`, and every path the
server touches must resolve inside one of those directories. Reads are gated
as well as writes, since a read outside the sandbox exfiltrates content just
as surely as a write plants it. The containment check runs on canonicalized
paths, so `..\` traversal, symlink and junction escapes, 8.3 short names,
extended-length prefixes, case tricks, and lookalike sibling directories
(`Documents2` against an allowed `Documents`) are all caught, and UNC network
paths are refused unless an allowed root is itself a UNC path that contains
them. A blocked call refuses with a typed error naming the offending path and
the allowed roots before any file is opened. Recommended whenever the server
runs against untrusted or semi-trusted agent traffic.

## Testing

1,255 tests (1,193 run everywhere; 62 live-marked tests drive a real Word
instance on Windows): the suite was developed against a private corpus of
real-world documents (book-length chapters, a document with 171 footnotes, a
manuscript with 126 tracked changes and reviewer comments), and CI
auto-generates **structurally equivalent synthetic stand-ins**
(`tests/make_corpus.py`) so the full suite runs on any machine, including
yours and every pull request. Local real documents, when present, take
precedence. `tests/word_validator.py` opens outputs in invisible Word and
fails on any repair prompt, the definitive corruption check.

Development history: prototyped with Claude Code in a day (2026-08-27), then
hardened across release cycles through dedicated adversarial rounds (scale
torture, pathological merge topologies, Unicode and schema fuzzing, ReDoS,
Word-lock lifecycle, live-editing interaction hunts, COM leak checks) plus
per-phase unit gates. Every finding fixed with a regression test, same session
it was found. `research/` documents the OOXML algorithms and pitfalls the
implementation is built on, with attribution to the MIT-licensed reference
implementations studied.

## Maturity: what the version number does and does not claim

This project moves fast and is honest about what backs it. What the test
record covers: every release passes the full suite plus dedicated adversarial
rounds through the raw MCP transport, against a corpus of long, heavily
formatted real-world documents with zero corruption across all of it. What it
does not yet cover: other machines, Word builds older than current Microsoft
365, non-English Word installs (some tools reference styles by localized
display name), RTL scripts, and the diversity of documents only real users
bring. The safety net while the tool earns that mileage is structural:
automatic slot backups before every mutation and atomic validated saves, so a
bad outcome is a restore, not a loss. If something misbehaves on your
documents, [an issue](https://github.com/nometalalchemist/KitchenSink4Word/issues)
with the symptom (never the document itself, unless it contains nothing
private) is the most valuable thing you can send.

**Beta-labeled tools**, heuristic by nature: review their flagged-items list
rather than trusting silently: `convert_citation_style`,
`anonymize_for_review`, `validate(checks=["defined_terms"])`, and the
text-reference scan inside `validate(checks=["cross_references"])`. Each
returns an explicit list of what it could not confidently handle.

## Known limits

- Live regex replacements skip matches positioned after complex fields in a
  story (COM character offsets drift there); the skip count is reported and a
  literal find still works. Live `set_cells` refuses vertically merged tables
  (the file-based tool is merge-aware; close the doc for those).
- Live tracked-change attribution is best-effort: Word signed into an Office
  account attributes revisions to that account; results report the effective
  author honestly.
- TOC and caption-list page numbers require a field update: automatic on next
  Word open, or immediate via `com_refresh_fields`.

## License

**AGPL-3.0**: free for any use under AGPL terms (copyleft: modifications to
served copies must be open-sourced). Commercial license available for use
cases that cannot comply;
[open an issue](https://github.com/nometalalchemist/KitchenSink4Word/issues/new?template=commercial_license.yml)
to arrange one.

---

*Not affiliated with or endorsed by Microsoft Corporation. Microsoft and Word
are trademarks of Microsoft Corporation. "For Microsoft Word" describes
file-format compatibility only.*
