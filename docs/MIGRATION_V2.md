# Migrating from KitchenSink4Word v1.x to v2.0

v2.0 is a hard break: the 189-tool v1.6 surface was rebuilt as a
consolidated set of 108 document tools (plus the two pack toggles,
`enable_tools` and `disable_tools`). Every v1 capability survives; only
the grammar changed. `migration/v1_to_v2.json` is the machine-readable
map (one entry per v1.6 tool), and `get_workflows("migrate-from-v1")`
returns the short in-session recipe. This guide is the narrative version.

## 1. What changed and why

Two costs drove the redesign. First, the token bill: a client that loads
every tool schema paid for 189 schemas in every conversation. v2 cuts
the full surface by consolidating overlapping tools into multiplexes
(one `list_elements` instead of 18 `list_*` tools, one `validate`
instead of 14 checkers), and tiered loading cuts it further: a fresh
session starts with the 28-tool lite core and enables the rest by pack,
on demand. Second, selection accuracy: with 189 names, similar tools
competed for the same task and agents picked wrong. v2 gives each
concept exactly one name, built from a small verb table (`insert_`,
`set_`, `manage_`, `list_elements`, `validate`, `delete_element`), so
the right tool is the only candidate.

The consolidation also added two capabilities v1 never had:
`get_document_view` (an anchored markdown projection of the document,
cheap enough to read whole sections in one call) and `apply_edits` (a
batch of anchor-addressed edits applied with one lock, one backup, one
validated save). Section 6 covers when to switch to them.

## 2. Breaking changes at a glance

- Every v1 tool name maps to a v2 home; many single-purpose tools became
  parameterized calls on a multiplex (see the rename table below).
- Positioning parameters (`after_index`, `before_index`, `after_anchor`,
  `at_end`, `anchor_text`/`occurrence`) became one `location` object
  (section 4).
- Tools that used to act on the first text match now refuse loudly when
  a text selector is ambiguous (section 5.1).
- `remove_watermark` and `remove_document_protection` became
  set-to-none calls on their `set_` siblings (section 5.2).
- The `validate_*` / `check_*` battery became `validate(checks=[...])`
  with a fixed check-name vocabulary (section 5.3).
- Responses carry a uniform envelope; refusals are structured objects
  with stable error codes, never raw exception strings (section 7).
- No capability was removed: the map is total, and there are no one-way
  doors. Live editing of open documents kept every v1 route and gained
  new guards (section 8).

## 3. The rename table

The table below is generated from `migration/v1_to_v2.json` by
`scripts/generate_migration_table.py`. Regenerate it after any map
change; do not edit it by hand.

<!-- BEGIN GENERATED RENAME TABLE (scripts/generate_migration_table.py) -->

All 189 v1.6 tools, grouped by their v2 destination (107 v2 tools receive them). `inject` values are the literal v2 parameters that reproduce the v1 tool's fixed behavior; parameter moves use dot paths into nested objects (so `location.search.text` means `location={"search": {"text": ...}}`).

### -> `anonymize_for_review`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| anonymize_for_review | `anonymize_for_review` | - | - |

### -> `apply_edits`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| batch_apply | `apply_edits` | - | op-type absorption into the batch layer |

### -> `apply_manuscript_format`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| apply_manuscript_format | `apply_manuscript_format` | - | - |

### -> `apply_style`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| apply_character_style | `apply_style` | `find` -> `target.search.text`; `occurrence` -> `target.search.occurrence`; unchanged: `style` | absorbed; the target selector routes to the character-style path. |
| apply_style | `apply_style` | `indices` -> `range` | a contiguous indices list [a..b] becomes range={start:a,end:b}; non-contiguous lists need one call per contiguous run; style carries over by name. |

### -> `apply_template`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| apply_template | `apply_template` | - | - |

### -> `assemble_front_matter`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| assemble_front_matter | `assemble_front_matter` | - | - |

### -> `change_heading_level`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| change_heading_level | `change_heading_level` | - | - |

### -> `com_export_pdf`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_export_pdf | `com_export_pdf` | - | - |

### -> `com_import_pdf`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_import_pdf | `com_import_pdf` | - | - |

### -> `com_multi_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_combine_documents | `com_multi_document(action="combine")` | `original_path` -> `files[0]`; `revised_path` -> `files[1]`; unchanged: `output_path` | merged; output still defaults beside the original. |
| com_compare_documents | `com_multi_document(action="compare")` | `original_path` -> `files[0]`; `revised_path` -> `files[1]`; unchanged: `output_path`, `author` | merged; output still defaults to <revised>_COMPARE.docx. |
| com_merge_documents | `com_multi_document(action="merge")` | `paths` -> `files`; unchanged: `output_path`, `section_break_between` | merged; output_path stays required for merge. |

### -> `com_proofing_errors`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_proofing_errors | `com_proofing_errors` | - | - |

### -> `com_readability_statistics`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_readability_statistics | `com_readability_statistics` | - | - |

### -> `com_refresh_fields`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_refresh_fields | `com_refresh_fields` | - | - |

### -> `com_save_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_close_open_document | `com_save_document(close=true)` | unchanged: `save` | merged; save=false still discards unsaved changes on close. |
| com_save_open_document | `com_save_document` | - | merged; the default call (save=true, close=false) reproduces it exactly. |
| com_save_with_password | `com_save_document` | unchanged: `password`, `output_path` | merged; runs on its own invisible instance as in v1; not combinable with close. |

### -> `com_validate_opens_clean`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_validate_opens_clean | `com_validate_opens_clean` | - | - |

### -> `com_word_status`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| com_word_status | `com_word_status` | - | - |

### -> `comment_report`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| comment_report | `comment_report` | `file_path` -> `files[0]` | single file becomes a one-element files array; include_resolved keeps its name (single-file mode only) |
| comment_report_multi | `comment_report` | `file_paths` -> `files` | - |

### -> `convert_citation_style`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| convert_citation_style | `convert_citation_style` | - | - |

### -> `convert_notes`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| convert_notes | `convert_notes` | - | - |

### -> `copy_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| copy_document | `copy_document` | - | - |

### -> `copy_table`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| copy_table | `copy_table` | - | identity; added day three post-design, rides insert_document's engine; docstring rewritten to budget in v2 |

### -> `create_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| create_document | `create_document` | - | - |

### -> `create_table`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| create_table | `create_table` | `after_index` -> `location.paragraph`; `after_anchor` -> `location.search.text` | at_end:true (or no position param) -> omit location entirely; location's default position 'after' matches v1 after_index semantics. v1 after_anchor matched FULL paragraph plain text; v2 search matches substrings and refuses ambiguity loudly, so a recurring anchor text that silently hit first-match in v1 now refuses with the match list. |

### -> `deanonymize_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| deanonymize_document | `deanonymize_document` | - | - |

### -> `define_style`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| define_style | `define_style` | - | identity. |

### -> `delete_element`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| delete_equation | `delete_element(type="equation")` | `index` -> `id` | - |
| delete_toc | `delete_element(type="reference_list")` | `which` -> `id` | v1 which was the read_toc order index; v2 ids come from list_elements(type='toc') |

### -> `delete_paragraphs`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| delete_paragraphs | `delete_paragraphs` | - | identity; start/end integer shorthand unchanged, range={start,end} location objects added. |

### -> `delete_table`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| delete_table | `delete_table` | - | - |

### -> `detect_citation_system`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| detect_citation_system | `detect_citation_system` | - | - |

### -> `diagnose_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| diagnose_document | `diagnose_document` | - | counted under Validation and workflow in 2.16; still no live route BY DESIGN; v2: top-level ok re-keyed to healthy (the envelope ok now means the call succeeded) |

### -> `export_images`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| extract_images | `export_images` | - | - |

### -> `export_table`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| export_table | `export_table` | - | - |

### -> `fill_template`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| fill_template | `fill_template` | - | - |

### -> `find_text`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| find_formatted | `find_text` | unchanged: `formatting`, `query`, `scope` | formatting mode is file-mode only and refuses regex=True or include_textboxes=True |
| find_text | `find_text` | - | - |

### -> `fix_accessibility`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| fix_accessibility | `fix_accessibility` | - | - |

### -> `format_cells`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| format_cells | `format_cells` | - | - |

### -> `format_text`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| change_case | `format_text` | `transform` -> `case`; `indices` -> `range`; unchanged: `find` | absorbed; a contiguous indices list [a..b] becomes range={start:a,end:b}; NON-CONTIGUOUS lists need one call per contiguous run. 'toggle' was never a v1 transform and is not in v2.0. |
| format_text | `format_text` | `paragraph_index` -> `range.start` | set range.end to the same index (single-paragraph ranges only in formatting mode); find/occurrence/formatting carry over by name. |

### -> `get_comments`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_comments | `get_comments` | - | identity; keeps the live route and the live='auto' param |

### -> `get_document_info`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_document_info | `get_document_info` | - | - |

### -> `get_headers_footers`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_headers_footers | `get_headers_footers` | - | - |

### -> `get_outline`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_outline | `get_outline` | - | - |

### -> `get_paragraph_format`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_paragraph_format | `get_paragraph_format` | - | - |

### -> `get_protection`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_protection | `get_protection` | - | - |

### -> `get_revision_report`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| revision_analytics | `get_revision_report(mode="analytics")` | - | - |
| revision_summary | `get_revision_report(mode="summary")` | - | - |

### -> `get_styles`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_styles | `get_styles` | - | - |

### -> `get_table`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_nested_table | `get_table` | `row` -> `nested.row`; `cell` -> `nested.cell`; `nested_index` -> `nested.index` | absorbed via the nested address object |
| get_table | `get_table` | - | identity; gains the optional nested={row, cell, index} address |

### -> `get_text`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_text | `get_text` | - | - |
| get_textbox_text | `get_text(textbox=true)` | - | full per-box shape verbatim; {"index": n} narrows to one box; file-mode only |

### -> `get_tracked_changes`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_tracked_changes | `get_tracked_changes` | - | - |

### -> `get_workflows`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_workflows | `get_workflows` | - | gains migrate-from-v1 and bulk-edit task entries (ops/workflows.py addition, see wave_A.md entry 22) |

### -> `import_table`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| import_table | `import_table` | - | identity; keeps at_end/after_anchor (table import does not take the location object) |

### -> `insert_bookmark`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_bookmark | `insert_bookmark` | - | - |

### -> `insert_break`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_page_break | `insert_break(type="page")` | `after_index` -> `location.paragraph` | - |
| add_section_break | `insert_break` | `after_index` -> `location.paragraph`; `break_type` -> `type` | value map: nextPage->section_next, continuous->section_continuous, evenPage->section_even, oddPage->section_odd (v1 default nextPage -> inject type:'section_next' when break_type was omitted) |

### -> `insert_caption`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_caption | `insert_caption` | - | - |

### -> `insert_chart`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_chart | `insert_chart` | `after_index` -> `location.paragraph` | after_anchor -> location:{search:{text}}; at_end -> omit location |

### -> `insert_citation`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| insert_citation | `insert_citation` | `anchor_text` -> `location.search.text`; `occurrence` -> `location.search.occurrence` | name unchanged; anchoring moves to the location object; pages/suppress flags/prefix/suffix keep their names |

### -> `insert_content_control`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| insert_content_control | `insert_content_control` | - | - |

### -> `insert_cross_reference`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_cross_reference | `insert_cross_reference` | `after_anchor` -> `location.search.text`; unchanged: `to_bookmark`, `kind` | renamed to the insert_ verb family; v1 had no occurrence param (acted on first match), the v2 search selector defaults occurrence=1 |

### -> `insert_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| insert_document | `insert_document` | `after_index` -> `location.paragraph`; `after_anchor` -> `location.search.text` | at_end=True maps to omitting location (v2 default appends at end). CAVEAT: v1 after_index counted body ITEMS (paragraphs AND tables); v2 location.paragraph counts paragraphs only, so indices must be re-derived on table-bearing documents. v1 after_anchor matched FULL paragraph text; v2 location.search matches content substrings, so pass occurrence or tighten the string on recurring text. |

### -> `insert_equation`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_equation | `insert_equation` | `after_index` -> `location.paragraph` | display placement moves to the location object; inline placement (display=false) keeps anchor_text/occurrence unchanged |

### -> `insert_field`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| insert_field | `insert_field` | `after_anchor` -> `location.search.text`; `occurrence` -> `location.search.occurrence` | name unchanged; anchoring moves to the location object; field_code/placeholder keep their names |

### -> `insert_hyperlink`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_hyperlink | `insert_hyperlink` | - | - |

### -> `insert_image`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_image | `insert_image` | `after_index` -> `location.paragraph` | v1 at_end=true or no position -> omit location (document end); after_anchor -> location:{search:{text}} |

### -> `insert_list`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_list | `insert_list` | `after_index` -> `location.paragraph` | v1 at_end=true -> location:{position:'end'} or omit location |

### -> `insert_paragraphs`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_heading | `insert_paragraphs` | `text` -> `paragraphs[0].text`; `level` -> `paragraphs[0].heading_level`; `after_index` -> `location.paragraph`; `after_anchor` -> `location.search.text` | merged; heading_level maps to the Heading style, or to outlineLvl on outline-based documents (the NSU lesson); at_end=true -> omit location entirely (document-end default). |
| insert_paragraphs | `insert_paragraphs` | `after_index` -> `location.paragraph`; `before_index` -> `location.paragraph`; `after_anchor` -> `location.search.text` | after_index keeps the default position 'after'; before_index also sets location.position='before'; after_anchor becomes a search selector (v1 anchors matched paragraph plain text as a substring, same as the search selector); at_end=true -> omit location entirely (document-end default, the cross-wave convention). Items gain optional heading_level. |

### -> `insert_reference_list`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| insert_bibliography | `insert_reference_list(type="bibliography")` | `after_index` -> `location.paragraph`; unchanged: `title`, `update_on_open` | v1 at_end=True is the v2 default when location is omitted |
| insert_caption_list | `insert_reference_list` | `after_index` -> `location.paragraph`; unchanged: `title`, `update_on_open` | v1 label decides the v2 type: Table -> type='table_list', Figure -> 'figure_list', Equation -> 'equation_list'; default placement when location omitted matches v1 (document start) |
| insert_glossary | `insert_reference_list(type="glossary")` | `heading` -> `title`; `heading_level` -> `options.heading_level`; `definition_patterns` -> `options.definition_patterns`; `after_index` -> `location.paragraph` | v1 at_end=True is the v2 default when location is omitted; glossary writes literal paragraphs, so update_on_open does not apply |
| insert_index | `insert_reference_list(type="index")` | `columns` -> `options.columns`; `letter_headings` -> `options.letter_headings`; `after_index` -> `location.paragraph`; unchanged: `title`, `update_on_open` | v1 at_end=True is the v2 default when location is omitted |
| insert_toc | `insert_reference_list(type="toc")` | `levels` -> `options.levels`; `after_index` -> `location.paragraph`; unchanged: `title`, `update_on_open` | v1 at_start=True (or after_index=None) is the v2 default placement when location is omitted |

### -> `insert_zotero_citation`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| insert_zotero_citation | `insert_zotero_citation` | `anchor_text` -> `location.search.text`; `occurrence` -> `location.search.occurrence` | name unchanged; anchoring moves to the location object; item_keys/page/prefix/suffix/db_path keep their names |

### -> `list_elements`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| get_lists | `list_elements(type="lists")` | - | - |
| list_bookmarks | `list_elements(type="bookmarks")` | - | - |
| list_charts | `list_elements(type="charts")` | - | - |
| list_content_controls | `list_elements(type="content_controls")` | - | items = v1 controls array; the v1 count key is dropped in favor of the envelope count |
| list_endnotes | `list_elements(type="endnotes")` | - | - |
| list_equations | `list_elements(type="equations")` | - | items = v1 equations array; equation_count sibling carried |
| list_fields | `list_elements(type="fields")` | - | items = v1 fields array; total and parts_scanned siblings carried |
| list_footnotes | `list_elements(type="footnotes")` | - | - |
| list_form_fields | `list_elements(type="form_fields")` | - | items = v1 fields array; the v1 count key is dropped in favor of the envelope count |
| list_images | `list_elements(type="images")` | - | - |
| list_index_entries | `list_elements(type="index_entries")` | - | - |
| list_reference_fields | `list_elements(type="reference_fields")` | - | items = v1 fields array; by_manager, broken, and the other v1 keys carried as siblings |
| list_section_blocks | `list_elements(type="section_blocks")` | - | - |
| list_sections | `list_elements(type="sections")` | - | - |
| list_sources | `list_elements(type="sources")` | - | - |
| list_tables | `list_elements(type="tables")` | - | - |
| list_template_placeholders | `list_elements(type="template_placeholders")` | - | items = v1 placeholders array; names and mergefields carried as siblings |
| read_toc | `list_elements(type="toc")` | - | items = v1 tocs array; present/instruction/cached_entries carried as siblings |

### -> `live_insert_at_cursor`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| live_insert_at_cursor | `live_insert_at_cursor` | - | - |

### -> `live_repair`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| word_live_repair | `live_repair` | - | renamed into the live_ namespace |

### -> `live_scroll_to`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| live_scroll_to | `live_scroll_to` | - | - |

### -> `live_set_track_changes`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| live_set_track_changes | `live_set_track_changes` | - | - |

### -> `mail_merge`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| mail_merge | `mail_merge` | - | - |

### -> `manage_backups`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| create_snapshot | `manage_backups(action="snapshot")` | unchanged: `label`, `dest_dir` | snapshots stay permanent keepers; no purge scope touches them |
| manage_backups | `manage_backups` | - | - |

### -> `manage_comment`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_comment | `manage_comment(action="add")` | `anchor_text` -> `location.search.text`; `occurrence` -> `location.search.occurrence`; unchanged: `text`, `author` | - |
| delete_comment | `manage_comment(action="delete")` | unchanged: `comment_id` | - |
| reply_to_comment | `manage_comment(action="reply")` | unchanged: `comment_id`, `text`, `author` | - |
| resolve_comment | `manage_comment(action="resolve")` | unchanged: `comment_id`, `done` | - |

### -> `manage_note`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_endnote | `manage_note(action="insert", note_type="endnote")` | `anchor_text` -> `location.search.text`; `occurrence` -> `location.search.occurrence`; `note_text` -> `text` | - |
| add_footnote | `manage_note(action="insert", note_type="footnote")` | `anchor_text` -> `location.search.text`; `occurrence` -> `location.search.occurrence`; `note_text` -> `text` | - |
| cleanup_orphan_notes | `manage_note(action="cleanup_orphans")` | - | covers footnotes AND endnotes in one call, as in v1 |
| delete_endnote | `manage_note(action="delete", note_type="endnote")` | unchanged: `note_id`, `position` | - |
| delete_footnote | `manage_note(action="delete", note_type="footnote")` | unchanged: `note_id`, `position` | - |
| edit_endnote | `manage_note(action="edit", note_type="endnote")` | `new_text` -> `text`; unchanged: `note_id`, `position` | - |
| edit_footnote | `manage_note(action="edit", note_type="footnote")` | `new_text` -> `text`; unchanged: `note_id`, `position` | - |

### -> `manage_source`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_source | `manage_source(action="add")` | - | all bibliographic params keep their v1 names (tag, source_type, title, year, authors, editors, journal_name, book_title, publisher, city, pages, volume, issue, edition, institution, url, internet_site_title, style, extra_fields) |
| delete_source | `manage_source(action="delete")` | unchanged: `tag`, `force` | - |

### -> `mark_index_entry`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| mark_index_entry | `mark_index_entry` | `anchor_text` -> `location.search.text`; `occurrence` -> `location.search.occurrence` | name unchanged; anchoring moves to the location object; entry/subentry/bold_page/italic_page/see keep their names |

### -> `modify_table_structure`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| delete_columns | `modify_table_structure(action="delete", target="columns")` | - | columns keeps its name |
| delete_rows | `modify_table_structure(action="delete", target="rows")` | - | start, end keep their names (end None = start only) |
| insert_columns | `modify_table_structure(action="insert", target="columns")` | - | at, count, width_pt keep their names |
| insert_rows | `modify_table_structure(action="insert", target="rows")` | - | at, count, copy_format_from keep their names |
| merge_cells | `modify_table_structure(action="merge")` | `start_row` -> `range.start_row`; `end_row` -> `range.end_row`; `start_col` -> `range.start_col`; `end_col` -> `range.end_col` | - |
| split_table | `modify_table_structure(action="split")` | - | at_row keeps its name |
| unmerge_cells | `modify_table_structure(action="unmerge")` | - | row, cell keep their names |

### -> `move_section`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| move_section | `move_section` | - | - |

### -> `parse_references`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| parse_references | `parse_references` | - | - |

### -> `prepare_for_submission`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| prepare_for_submission | `prepare_for_submission` | - | - |

### -> `redact_text`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| redact_text | `redact_text` | - | - |

### -> `resolve_revisions`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| accept_revisions | `resolve_revisions(action="accept")` | unchanged: `author` | - |
| reject_revisions | `resolve_revisions(action="reject")` | unchanged: `author` | - |

### -> `search_and_replace`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| preview_replace | `search_and_replace(preview=true)` | - | absorbed; preview is the dry-run flag of the same operation, result shape unchanged. |
| replace_formatted | `search_and_replace` | `formatting` -> `find_formatting`; `replace` -> `replacements[0].replace`; `find` -> `replacements[0].find` | absorbed; find=None (omit replacements[0].find) keeps the whole-matching-stretch semantics; scope and max_replacements carry over by name; no regex or track in this mode, as in v1. |
| search_and_replace | `search_and_replace` | - | identity; keeps its name; gains preview and find_formatting |

### -> `search_zotero_library`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| search_zotero_library | `search_zotero_library` | - | - |

### -> `set_bibliography_style`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_bibliography_style | `set_bibliography_style` | - | - |

### -> `set_cells`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_cells | `set_cells` | - | identity in scatter mode; gains block and nested modes |
| set_cells_block | `set_cells` | `origin_row` -> `block.origin.row`; `origin_cell` -> `block.origin.cell`; `data` -> `block.values` | block mode is file-only, as the v1 tool was; no track, no live route |
| set_nested_cells | `set_cells` | `row` -> `nested.row`; `cell` -> `nested.cell`; `nested_index` -> `nested.index` | edits keeps its name; nested mode is file-only, as the v1 tool was |

### -> `set_chart_data`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| update_chart_data | `set_chart_data` | `chart_index` -> `chart_id` | - |

### -> `set_content_control`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_content_control_value | `set_content_control` | - | - |

### -> `set_document_properties`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_document_properties | `set_document_properties` | - | - |
| set_update_fields_flag | `set_document_properties` | `on` -> `update_fields_on_open` | it is a document setting; one edit pass covers metadata plus the flag |

### -> `set_document_protection`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| remove_document_protection | `set_document_protection(protection="none")` | - | - |
| set_document_protection | `set_document_protection` | `edit` -> `protection` | - |

### -> `set_form_fields`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| fill_form_fields | `set_form_fields` | - | - |

### -> `set_header_footer`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_footer | `set_header_footer(part="footer")` | - | - |
| set_header | `set_header_footer(part="header")` | - | - |

### -> `set_image`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| replace_image | `set_image` | `image_index` -> `image_id`; `new_image_path` -> `source` | - |
| resize_image | `set_image` | `image_index` -> `image_id` | - |
| set_image_alt_text | `set_image` | `image_index` -> `image_id`; `description` -> `alt_text`; `title` -> `alt_title` | - |

### -> `set_page_numbers`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_page_numbers | `set_page_numbers` | `start_at` -> `format.start_at` | - |
| set_page_number_format | `set_page_numbers` | `number_format` -> `format.number_format`; `start_at` -> `format.start_at` | format-only call: leave position unset |

### -> `set_paragraph_format`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_paragraph_format | `set_paragraph_format` | - | identity. |

### -> `set_paragraph_text`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| replace_paragraph_text | `set_paragraph_text` | `index` -> `location.paragraph` | renamed to the set_ verb family; expect guard and replaced_text return unchanged |

### -> `set_section_properties`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_columns | `set_section_properties` | `count` -> `columns.count`; `space_pt` -> `columns.space_pt`; `separator` -> `columns.separator`; `widths_pt` -> `columns.widths_pt` | - |
| set_line_numbering | `set_section_properties` | `count_by` -> `line_numbering.count_by`; `start` -> `line_numbering.start`; `restart` -> `line_numbering.restart`; `distance_pt` -> `line_numbering.distance_pt` | v1 remove=true -> line_numbering:'none' (value transform, not a path rename) |
| set_section_properties | `set_section_properties` | - | - |

### -> `set_table_properties`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| apply_table_style | `set_table_properties` | - | style, banded_rows, first_row_header keep their names; banded_rows/first_row_header only apply when style is passed |
| set_column_widths | `set_table_properties` | `widths_pt` -> `column_widths` | - |
| set_header_row_repeat | `set_table_properties` | `rows` -> `header_row_repeat` | on:false -> header_row_repeat:false regardless of rows; on:true with rows N -> header_row_repeat:N (true alone = 1 row) |

### -> `set_textbox_text`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| set_textbox_text | `set_textbox_text` | - | - |

### -> `set_watermark`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| add_watermark | `set_watermark` | `text` -> `watermark.text`; `color` -> `watermark.color`; `opacity` -> `watermark.opacity`; `diagonal` -> `watermark.diagonal` | - |
| remove_watermark | `set_watermark(watermark="none")` | - | set-to-none is the removal idiom |

### -> `setup_chapter_headers`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| setup_chapter_headers | `setup_chapter_headers` | - | - |

### -> `sort_table`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| sort_table | `sort_table` | - | - |

### -> `split_document`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| split_document | `split_document` | - | - |

### -> `structured_diff`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| structured_diff | `structured_diff` | - | - |

### -> `validate`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| audit_accessibility | `validate(checks=["accessibility"])` | - | v2: report top level is {passed, results} (ok belongs to the envelope) |
| check_brand_compliance | `validate(checks=["brand"])` | `rules` -> `options.brand` | options.brand takes the same rules dict the v1 tool took; v2: report top level is {passed, results} (ok belongs to the envelope) |
| check_citation_parity | `validate(checks=["citation_parity"])` | - | v2: report top level is {passed, results} (ok belongs to the envelope) |
| check_defined_terms | `validate(checks=["defined_terms"])` | `definition_patterns` -> `options.defined_terms.definition_patterns` | v2: report top level is {passed, results} (ok belongs to the envelope) |
| check_image_resolution | `validate(checks=["image_resolution"])` | `min_dpi` -> `options.image_resolution.min_dpi` | runs as one check inside the validate battery; v2: report top level is {passed, results} (ok belongs to the envelope) |
| check_reference_field_integrity | `validate(checks=["reference_fields"])` | - | v2: report top level is {passed, results} (ok belongs to the envelope) |
| check_template_compliance | `validate(checks=["template"])` | `rules` -> `options.template` | options.template takes the v1 rules DICT (same flag as brand); v2: report top level is {passed, results} (ok belongs to the envelope) |
| validate_captions | `validate(checks=["captions"])` | - | v2: report top level is {passed, results} (ok belongs to the envelope) |
| validate_chapter_headers | `validate(checks=["chapter_headers"])` | - | v2: report top level is {passed, results} (ok belongs to the envelope) |
| validate_cross_references | `validate(checks=["cross_references"])` | - | v2: report top level is {passed, results} (ok belongs to the envelope) |
| validate_document | `validate(checks=["core"])` | - | core is the default check set; v1 result dict lands verbatim under results.core.findings; v2: report top level is {passed, results} (ok belongs to the envelope) |
| validate_form_completeness | `validate(checks=["forms"])` | `required` -> `options.forms.required` | v2: report top level is {passed, results} (ok belongs to the envelope) |
| validate_notes | `validate(checks=["notes"])` | - | v2: report top level is {passed, results} (ok belongs to the envelope) |
| verify_redaction | `validate(checks=["redaction"])` | `targets` -> `options.redaction.targets` | the redaction check re-scans for the given patterns without changing anything; v2: report top level is {passed, results} (ok belongs to the envelope) |

### -> `word_count`

| v1 tool | v2 call | parameter moves | notes |
|---|---|---|---|
| word_count | `word_count` | - | - |
| word_count_with_exclusions | `word_count` | `exclude` -> `exclusions` | exclusions mode is file-mode only; by_section not consulted there |

<!-- END GENERATED RENAME TABLE -->

### Worked example: add_footnote to manage_note

v1:

```json
{"tool": "add_footnote",
 "params": {"file_path": "d.docx", "anchor_text": "as shown",
            "occurrence": 2, "note_text": "See appendix B."}}
```

v2 (one lifecycle tool per note family; `action` picks the operation,
`note_type` picks footnote or endnote):

```json
{"tool": "manage_note",
 "params": {"file_path": "d.docx", "action": "insert",
            "note_type": "footnote",
            "location": {"search": {"text": "as shown", "occurrence": 2}},
            "text": "See appendix B."}}
```

`edit_footnote` becomes `action: "edit"`, `delete_footnote` becomes
`action: "delete"`, `list_footnotes` becomes
`list_elements(type="footnotes")`, and the same pattern covers endnotes
via `note_type: "endnote"`.

### Worked example: insert_rows to modify_table_structure

v1:

```json
{"tool": "insert_rows",
 "params": {"file_path": "d.docx", "table_index": 0, "at": 3,
            "count": 2, "copy_format_from": 2}}
```

v2 (`action` + `target` pick the structural operation; `at`, `count`,
and `copy_format_from` keep their names):

```json
{"tool": "modify_table_structure",
 "params": {"file_path": "d.docx", "table_index": 0, "action": "insert",
            "target": "rows", "at": 3, "count": 2, "copy_format_from": 2}}
```

`delete_rows`, `insert_columns`, `delete_columns`, `merge_cells`,
`unmerge_cells`, `split_table`, and `set_column_widths` are the other
faces of the same tool; each map entry gives the exact `action`/`target`
pair.

### Worked example: list_* to list_elements

v1 had one enumerator per element family (`list_tables`,
`list_footnotes`, `list_images`, ...). v2 has one:

```json
{"tool": "list_elements",
 "params": {"file_path": "d.docx", "type": "tables"}}
```

The per-type result shapes are unchanged from the v1 enumerators, so
whatever parsed `list_tables` output parses
`list_elements(type="tables")` output.

## 4. The location object

Positional tools take one `location` parameter with exactly one selector
key plus an optional `position` modifier
(`before | after | replace | start | end`; inserters default to `after`,
omitting `location` entirely means document end).

| v1 addressing | v2 location |
|---|---|
| `after_index: 15` | `{"paragraph": 15, "position": "after"}` |
| `before_index: 15` | `{"paragraph": 15, "position": "before"}` |
| `at_end: true` | omit `location` (or `{"position": "end"}`) |
| `after_anchor: "Results"` / `anchor_text` + `occurrence` | `{"search": {"text": "Results", "occurrence": 1}}` |
| (no v1 equivalent) | `{"after_heading": {"text": "Chapter 3"}}`, `{"outline": "3.2"}`, `{"bookmark": "methods"}`, `{"anchor": "a3f9"}`, `{"cursor": true}` |

Notes:

- `paragraph` indices are 0-based body-paragraph indices, exactly as in
  v1 (paragraphs inside tables do not count).
- `outline` is the preferred structural address (`"3.2"` means the
  second level-2 heading under the third level-1), because heading text
  recurs in body prose on real documents.
- `anchor` takes an anchor id from `get_document_view`; anything the
  view shows, every positional tool can address.
- `cursor` works only on documents open in Word; it reads the user's
  selection start and never moves it.
- Range-taking tools (`delete_paragraphs`, `format_text`, `apply_style`)
  accept `range={"start": ..., "end": ...}` where each endpoint is a
  location object or a bare paragraph index.

## 5. Behavior changes that are not renames

### 5.1 Ambiguous text selectors refuse loudly

In v1, several anchor-text tools acted on the first match
(`add_cross_reference` had no occurrence parameter at all). In v2, a
text selector matching more than one place without an `occurrence`
refuses with the code `AMBIGUOUS_LOCATION` and returns every match with
its paragraph index and context, plus a hint to pass `occurrence` or
address by `outline`/`paragraph`/`anchor`. No v2 tool acts on first
match. Zero matches also refuse, with nearest-miss hints (including the
curly-versus-straight-quote and XML-entity cases). If your v1 call
relied on first-match behavior, add `occurrence: 1` to reproduce it
deliberately.

### 5.2 Set-to-none removals

Two v1 removal tools became the `"none"` value of their setter:

- `remove_watermark` is now `set_watermark(watermark="none")`
- `remove_document_protection` is now
  `set_document_protection(protection="none")`

### 5.3 validate check names

The v1 validation battery merged into one read-only tool:
`validate(file_path, checks=[...], options={...})`. The check names are
a fixed vocabulary; each check's findings keep the v1 result shape:

`core` (was `validate_document`), `captions`, `chapter_headers`,
`cross_references`, `notes`, `forms`, `citation_parity`,
`defined_terms`, `brand`, `template`, `reference_fields`,
`image_resolution`, `accessibility`, `redaction`.

Per-check options are namespaced (for example
`options.brand`, `options.template`, `options.redaction`). `validate`
never mutates; repair tools stay separate (`fix_accessibility`,
`redact_text`, orphan-note cleanup via
`manage_note(action="cleanup_orphans")`).

## 6. The view/batch layer

`get_document_view` returns the document as markdown with a stable
anchor id per paragraph and pipe-rendered tables (cells addressable as
`t:hex:rNcN`). `apply_edits` takes a list of ops (`replace`, `set_text`,
`insert`, `delete`, `set_style`, `format`, `set_paragraph_format`,
`set_cell`) addressed by those anchors and applies the whole batch with
one lock, one backup, one validated save; one stale anchor refuses the
entire batch before anything mutates.

Rule of thumb: three or more edits in one section, take a scoped view
and send one batch; for one or two surgical changes, the fine-grained
tools are cheaper. The fine-grained tools accept the view's anchors too
(`{"anchor": "a3f9"}`), so the two styles mix freely.

## 7. The response envelope

- Success responses carry `ok: true` and (when the call named a file)
  `file`. Mutation responses keep their operation-specific fields and
  the save/backup fields (`saved`, `backup`).
- Refusals are structured: `{"ok": false, "error": {"code", "message",
  "hint", "matches?", "detail?"}}` with `isError` set at the MCP level.
  The code vocabulary is closed and documented: `AMBIGUOUS_LOCATION`,
  `NOT_FOUND`, `DOCUMENT_LOCKED`, `WORD_NOT_RUNNING`, `WORD_BUSY`,
  `WORD_BLOCKED`, `PROTECTED_VIEW`, `VALIDATION_FAILED`, `STALE_ANCHOR`,
  `RANGE_OUT_OF_BOUNDS`, `UNSUPPORTED_CONTENT`, `CONFLICT`,
  `BAD_PARAMS`.
- File mode is the canonical shape. Live execution only ADDS fields
  (`live: true` plus undo/dirty metadata) and never changes the shape,
  so one parsing path covers both modes. List-shaped reader results stay
  flat lists.
- If your v1 parser read file-mode responses, it needs little or no
  change; v1 live-mode special-casing can be deleted.

## 8. Live mode (documents open in Word)

The dual-mode contract carried over: tools with a live route try the
file first and switch to the open document when Word holds the lock
(`live="auto"`, with `"force"` and `"off"` overrides). Every v1
live-routed capability kept its route under its v2 name, and the batch
layer added one: `apply_edits` applies live as a single undo step.

Live-routed v2 tools: `get_text`, `find_text`, `get_outline`,
`get_document_info`, `get_comments`, `word_count`, `insert_paragraphs`,
`delete_paragraphs`, `set_paragraph_text` (was
`replace_paragraph_text`), `search_and_replace`, `format_text`,
`set_paragraph_format`, `set_cells`, `apply_edits`.

What to know when a document is open in Word:

- **Locations resolve against the last saved state.** Word does not
  expose unsaved edits to other processes, so text selectors (`search`,
  `after_heading`, `outline`, `anchor`) resolve on the saved file. Every
  live write then re-verifies its target inside Word before touching it:
  if unsaved changes moved the target, the call refuses with
  `STALE_ANCHOR` and nothing is changed. Saving the document in Word
  (or `com_save_document`) always clears the refusal. Plain `paragraph`
  indices keep the v1 contract (trusted as given; use `expect` on
  `set_paragraph_text` to guard them).
- **Some sub-modes are file-only.** Markdown lists and pipe tables in
  `apply_edits` insert ops, format cloning
  (`inherit_format`/`copy_format_from`), `find_formatting` and `preview`
  on `search_and_replace`, `case` mode on `format_text`, `exclusions`
  on `word_count`, and text-box modes on `get_text`/`find_text` refuse
  live with a message naming the fix (usually: close or save the file,
  or use the file-mode sibling). `heading_level` items on
  `insert_paragraphs` DO work live: they apply Word's built-in heading
  styles (or direct outline levels on outline-numbered documents).
- **Readers without a live route refuse rather than lie.** Tools that
  parse the saved XML (for example `diagnose_document`,
  `list_elements`, `validate`) refuse an open document, because the
  saved XML is stale while Word holds unsaved changes; the refusal names
  the live alternatives. `get_document_view` is the exception: it reads
  the last saved state and says so explicitly in the result
  (`live: true` plus a note).
- **Saving is the user's.** Live edits appear in the open window as one
  undo step and stay unsaved until the user (or `com_save_document`)
  saves. File-mode tools keep writing through to disk with auto-backup
  and validated saves, unchanged from v1.
- The `com_` tools (PDF export/import, compare/combine, proofing,
  readability, field refresh) and `live_` tools (cursor insert, scroll,
  track-changes toggle, repair) kept their prefixes and semantics; see
  the rename table for the few that merged.

## 9. Tiered loading (packs)

A fresh v2 session exposes the 28-tool lite core (including the two
toggles). The other 82 tools are grouped into seven packs:
`references`, `review`, `academic`, `assembly`, `media-forms`,
`com-live`, `protection-io`. Call `enable_tools(pack)` to add a pack
mid-session; calling a disabled tool returns a signpost naming its pack
and the exact enabling call, so discovery is self-serve. Startup surface
is configurable with `KS4W_MODE` (`lite`, `full`, or a comma-separated
pack list), and `KS4W_PACK_POLICY=locked` pins the surface for hosts
that manage it themselves. The `.mcpb` bundle exposes the same two
choices as checkboxes writing `KS4W_ALL_TOOLS` and `KS4W_LOCK_TOOLS`;
the two power-user variables above beat them when set. Nothing persists
across sessions by design.

## 10. Known one-way doors

None. The migration map is total: all 189 v1.6 tools have a v2 home with
equal or wider capability, and the map file ships with the package so
the translation can be automated.
