"""v2 migration map: structure and completeness, both ENFORCED (Phase 2).

Phase 0 checked structure only and xfail'd completeness while the map was
a skeleton; the Phase 2 integration merged all five wave fragments plus
the batch_apply row, so every check is enforced. The v1-side authority is
the FROZEN list of the 189 tools that shipped in v1.6.0 (the live
registry now serves the v2 surface, so it can no longer vouch for v1
names); the v2-side check runs against the live registry, with the
declared Phase 3 arrivals (apply_edits, get_document_view) pending.
"""

from __future__ import annotations

import asyncio
import json
from pathlib import Path

from word_mcp import server

MAP_PATH = (
    Path(__file__).resolve().parents[2] / "migration" / "v1_to_v2.json"
)

# The 189 tools registered by the shipped v1.6.0 server.py, frozen at the
# Phase 2 rebuild (extracted from the last v1 revision's @mcp.tool defs).
V1_TOOLS = frozenset({
    'accept_revisions', 'add_bookmark', 'add_caption', 'add_chart',
    'add_comment', 'add_cross_reference', 'add_endnote', 'add_equation',
    'add_footnote', 'add_heading', 'add_hyperlink', 'add_image', 'add_list',
    'add_page_break', 'add_page_numbers', 'add_section_break', 'add_source',
    'add_watermark', 'anonymize_for_review', 'apply_character_style',
    'apply_manuscript_format', 'apply_style', 'apply_table_style',
    'apply_template', 'assemble_front_matter', 'audit_accessibility',
    'batch_apply', 'change_case', 'change_heading_level',
    'check_brand_compliance', 'check_citation_parity', 'check_defined_terms',
    'check_image_resolution', 'check_reference_field_integrity',
    'check_template_compliance', 'cleanup_orphan_notes',
    'com_close_open_document', 'com_combine_documents',
    'com_compare_documents', 'com_export_pdf', 'com_import_pdf',
    'com_merge_documents', 'com_proofing_errors',
    'com_readability_statistics', 'com_refresh_fields',
    'com_save_open_document', 'com_save_with_password',
    'com_validate_opens_clean', 'com_word_status', 'comment_report',
    'comment_report_multi', 'convert_citation_style', 'convert_notes',
    'copy_document', 'copy_table', 'create_document', 'create_snapshot',
    'create_table', 'deanonymize_document', 'define_style', 'delete_columns',
    'delete_comment', 'delete_endnote', 'delete_equation', 'delete_footnote',
    'delete_paragraphs', 'delete_rows', 'delete_source', 'delete_table',
    'delete_toc', 'detect_citation_system', 'diagnose_document',
    'edit_endnote', 'edit_footnote', 'export_table', 'extract_images',
    'fill_form_fields', 'fill_template', 'find_formatted', 'find_text',
    'fix_accessibility', 'format_cells', 'format_text', 'get_comments',
    'get_document_info', 'get_headers_footers', 'get_lists',
    'get_nested_table', 'get_outline', 'get_paragraph_format',
    'get_protection', 'get_styles', 'get_table', 'get_text',
    'get_textbox_text', 'get_tracked_changes', 'get_workflows',
    'import_table', 'insert_bibliography', 'insert_caption_list',
    'insert_citation', 'insert_columns', 'insert_content_control',
    'insert_document', 'insert_field', 'insert_glossary', 'insert_index',
    'insert_paragraphs', 'insert_rows', 'insert_toc',
    'insert_zotero_citation', 'list_bookmarks', 'list_charts',
    'list_content_controls', 'list_endnotes', 'list_equations', 'list_fields',
    'list_footnotes', 'list_form_fields', 'list_images', 'list_index_entries',
    'list_reference_fields', 'list_section_blocks', 'list_sections',
    'list_sources', 'list_tables', 'list_template_placeholders',
    'live_insert_at_cursor', 'live_scroll_to', 'live_set_track_changes',
    'mail_merge', 'manage_backups', 'mark_index_entry', 'merge_cells',
    'move_section', 'parse_references', 'prepare_for_submission',
    'preview_replace', 'read_toc', 'redact_text', 'reject_revisions',
    'remove_document_protection', 'remove_watermark', 'replace_formatted',
    'replace_image', 'replace_paragraph_text', 'reply_to_comment',
    'resize_image', 'resolve_comment', 'revision_analytics',
    'revision_summary', 'search_and_replace', 'search_zotero_library',
    'set_bibliography_style', 'set_cells', 'set_cells_block',
    'set_column_widths', 'set_columns', 'set_content_control_value',
    'set_document_properties', 'set_document_protection', 'set_footer',
    'set_header', 'set_header_row_repeat', 'set_image_alt_text',
    'set_line_numbering', 'set_nested_cells', 'set_page_number_format',
    'set_paragraph_format', 'set_section_properties', 'set_textbox_text',
    'set_update_fields_flag', 'setup_chapter_headers', 'sort_table',
    'split_document', 'split_table', 'structured_diff', 'unmerge_cells',
    'update_chart_data', 'validate_captions', 'validate_chapter_headers',
    'validate_cross_references', 'validate_document',
    'validate_form_completeness', 'validate_notes', 'verify_redaction',
    'word_count', 'word_count_with_exclusions', 'word_live_repair'
})

# v2 targets declared in the map but arriving with the Phase 3 view/batch
# layer; drop entries here as the tools land.
PENDING_V2_TARGETS = {"apply_edits", "get_document_view"}


def _load():
    return json.loads(MAP_PATH.read_text(encoding="utf-8"))


def test_map_structure():
    data = _load()
    assert data["schema"] == 1
    assert data["v1_version"].startswith("1.6")
    assert data["v2_version"].startswith("2.0")
    tools = data["tools"]
    assert isinstance(tools, dict) and tools
    for v1_name, entry in tools.items():
        assert isinstance(entry, dict), f"{v1_name} entry is not a dict"
        assert entry.get("to"), f"{v1_name} entry lacks 'to'"
        for field in entry:
            assert field in {"to", "inject", "params", "notes"}, (
                f"{v1_name} entry has unknown field {field!r}"
            )
        if "params" in entry:
            assert isinstance(entry["params"], dict)
        if "inject" in entry:
            assert isinstance(entry["inject"], dict)


def test_map_entries_name_real_v1_tools():
    """Every v1-side key must be a tool that actually shipped in v1.6.0
    (a typo here would silently strand migrators)."""
    for v1_name in _load()["tools"]:
        assert v1_name in V1_TOOLS, (
            f"map entry {v1_name!r} is not a shipped v1.6.0 tool"
        )


def test_map_targets_are_registered_v2_tools():
    """Every 'to' target must exist on the rebuilt v2 surface (or be a
    declared Phase 3 arrival), so no migrator is pointed at a ghost."""
    registered = {t.name for t in asyncio.run(server.mcp.list_tools())}
    for v1_name, entry in _load()["tools"].items():
        target = entry["to"]
        assert target in registered or target in PENDING_V2_TARGETS, (
            f"map entry {v1_name!r} points at {target!r}, which is neither "
            "registered nor a declared pending Phase 3 tool"
        )


def test_map_completeness():
    """All 189 shipped v1 tools appear in the map: the mechanical 'no
    functionality lost' gate, ENFORCED from Phase 2."""
    mapped = set(_load()["tools"])
    missing = sorted(V1_TOOLS - mapped)
    assert not missing, (
        f"{len(missing)} shipped v1 tools missing from the migration map: "
        f"{missing[:10]}..."
    )
    stray = sorted(mapped - V1_TOOLS)
    assert not stray, f"map entries that were never v1 tools: {stray}"
    assert len(mapped) == 189
