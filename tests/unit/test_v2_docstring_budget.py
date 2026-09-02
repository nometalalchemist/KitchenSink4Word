"""v2 docstring budgets, ported from the pptx budget test.

Phase 0 status: the no-em-dash check is ENFORCED now (v1 already passes
it). The token budget runs against the v1 surface but is expected to fail
(v1 was not written to the v2 budgets: 95 of 189 descriptions sit outside
[60, 130] and two exceed even the multiplex cap), so it carries xfail.
Phase 2 flip: delete the pytest.mark.xfail line and update MULTIPLEX to
the final v2 multiplex set; nothing else changes.
"""

from __future__ import annotations

import asyncio

import pytest

from word_mcp import server


def _tools():
    """The shipped surface only. Test files import gitignored integration/
    staging snippets that re-register a handful of tools on the shared
    FastMCP instance with STALE docstrings (paste-ready duplicates of code
    already merged into server.py), so under a full-suite run some
    registry entries carry pre-merge text. Filter to tools whose function
    lives in a word_mcp module; the budgets police shipped code, not
    staging artifacts."""
    tools = asyncio.run(server.mcp.list_tools())
    return [
        t for t in tools
        if getattr(getattr(t, "fn", None), "__module__", "").startswith(
            "word_mcp"
        )
    ]


# The planned v2 multiplex/high-traffic tools that earn the ~350-token cap
# (finalized in Phase 2; the v1 names below keep the test meaningful while
# it runs against the v1 surface).
MULTIPLEX = {
    "list_elements", "validate", "manage_note", "manage_comment",
    "manage_source", "modify_table_structure", "com_multi_document",
    "insert_reference_list", "apply_edits", "get_document_view",
    "search_and_replace", "resolve_revisions", "manage_backups",
    "enable_tools", "get_workflows", "diagnose_document",
    "insert_document", "convert_citation_style",
}


def test_no_em_dashes_in_descriptions():
    """Standing rule: no em dashes anywhere public. Enforced from Phase 0
    (v1 passes it today; keep it that way)."""
    for tool in _tools():
        desc = tool.description or ""
        assert "—" not in desc, f"{tool.name} description has an em dash"


@pytest.mark.xfail(
    reason="v2 budgets apply from Phase 2 (v1 docstrings predate the "
    "80-120/350 budget)",
    strict=False,
)
def test_docstring_budget():
    """Every tool description inside 60-130 tokens (chars/4; multiplex
    tools get headroom to ~350)."""
    for tool in _tools():
        desc = tool.description or ""
        assert desc, f"{tool.name} has no description"
        tokens = len(desc) / 4
        cap = 350 if tool.name in MULTIPLEX else 130
        assert 60 <= tokens <= cap, (
            f"{tool.name} description is ~{tokens:.0f} tokens, "
            f"outside [60, {cap}]"
        )
