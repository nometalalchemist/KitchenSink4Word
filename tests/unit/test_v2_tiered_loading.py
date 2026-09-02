"""v2 Phase 4: tiered loading wired for real.

Pack membership finalized against the 108-tool surface (+2 toggles),
enable_tools/disable_tools registered, KS4W_MODE/KS4W_PACK_POLICY
startup handling, the fastmcp 3.x visibility route (global transform at
startup, session-scoped toggles mid-session), and the discoverability
contract. The red-gate discoverability scenarios (fresh agents enabling
the right pack unprompted) run in Phase 6; the stubs here prove the
plumbing they depend on.

Wire tests ride an in-process fastmcp Client and replicate main()'s
startup wiring (global Visibility transform) in a try/finally that
restores process-global state, because the FastMCP instance and the
packs bookkeeping are shared with every other test in the suite.
"""

from __future__ import annotations

import asyncio
import json

import pytest

from fastmcp import Client
from fastmcp.exceptions import ToolError
from fastmcp.server.transforms.visibility import Visibility

from word_mcp import envelope, packs, server
from word_mcp.core.errors import WordMcpError


def _shipped_tools():
    """The shipped surface only (same filter as the budget test: staging
    snippets re-register stale duplicates under a full-suite run)."""
    tools = asyncio.run(server.mcp.list_tools())
    return {
        t.name: t for t in tools
        if getattr(getattr(t, "fn", None), "__module__", "").startswith(
            "word_mcp"
        )
    }


# The Phase 4 lite core, finalized (26 content tools + the 2 toggles).
LITE = {
    "create_document", "copy_document", "manage_backups",
    "get_document_info", "diagnose_document", "get_workflows",
    "get_text", "find_text", "get_outline", "get_document_view",
    "apply_edits", "search_and_replace", "insert_paragraphs",
    "delete_paragraphs", "set_paragraph_text", "insert_break",
    "insert_list", "format_text", "apply_style", "set_paragraph_format",
    "create_table", "delete_table", "get_table", "set_cells",
    "list_elements", "delete_element", "enable_tools", "disable_tools",
}


@pytest.fixture
def restore_enabled():
    """Snapshot and restore the process-global enabled bookkeeping."""
    saved = dict(packs._ENABLED)
    yield
    packs._ENABLED.clear()
    packs._ENABLED.update(saved)
    server._PENDING_VISIBILITY.clear()


# ------------------------------------------------- membership integrity


def test_every_tool_in_exactly_one_pack():
    """No orphans, no double-homing: the registry partitions the shipped
    surface."""
    shipped = set(_shipped_tools())
    seen: dict[str, str] = {}
    for pack, members in packs.tool_names().items():
        for name in members:
            assert name not in seen, (
                f"{name} is in both {seen[name]} and {pack}"
            )
            seen[name] = pack
    assert set(seen) == shipped, (
        f"registry/registration drift: only-registry="
        f"{sorted(set(seen) - shipped)} only-registered="
        f"{sorted(shipped - set(seen))}"
    )


def test_lite_membership_finalized():
    assert set(packs.pack_tools("lite")) == LITE


def test_surface_counts():
    """108 consolidated tools + enable_tools/disable_tools."""
    assert len(_shipped_tools()) == 110
    assert len(packs.pack_tools("lite")) == 28


def test_toggles_always_on():
    """enable_tools/disable_tools live in the lite core (visible in every
    mode) and lite itself can never be toggled off."""
    assert packs.pack_of("enable_tools") == "lite"
    assert packs.pack_of("disable_tools") == "lite"
    with pytest.raises(WordMcpError):
        packs.disable(["lite"])


def test_menu_matches_registry_and_costs():
    """enable_tools' docstring is the pack menu: script-generated from
    PACK_SUMMARIES + pack_cost, never hand-listed. Every pack name and
    its measured bill appear; the bills match a fresh recomputation."""
    desc = _shipped_tools()["enable_tools"].description
    for pack in packs.pack_names():
        cost = sum(
            packs.approx_tokens(t)
            for t in packs._REGISTRY[pack].values()
        )
        line = f"- {pack} (~{cost / 1000:.1f}k): "
        assert line in desc, f"menu line for {pack} missing or stale"
        assert packs.PACK_SUMMARIES[pack] in desc


def test_pack_bills_cost_aware():
    """The 2026-09-02 granularity ruling: lite within reach of the gate
    (honest ceiling asserted, exact number in the phase report) and no
    pack below ~1.5k except the two documented exceptions: com-live is
    environment-gated at any size, protection-io is kept as rare-use."""
    lite_cost = sum(
        packs.approx_tokens(t) for t in packs._REGISTRY["lite"].values()
    )
    assert lite_cost <= 8000, f"lite regressed to ~{lite_cost} tokens"
    for pack in packs.pack_names():
        if pack in ("com-live", "protection-io"):
            continue
        assert packs.pack_cost(pack) >= 1500, (
            f"{pack} fell below the 1.5k line; merge or justify"
        )


# ------------------------------------------------- mode startup matrix


def _reset_to_lite():
    for name in packs._ENABLED:
        packs._ENABLED[name] = packs.pack_of(name) == "lite"


def test_mode_default_lite(restore_enabled, monkeypatch):
    monkeypatch.delenv("KS4W_MODE", raising=False)
    _reset_to_lite()
    assert packs.apply_startup_mode() == "lite"
    assert server._startup_disabled_names() == (
        set(_shipped_tools()) - LITE
    )


def test_mode_full(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4W_MODE", "full")
    _reset_to_lite()
    packs.apply_startup_mode()
    assert server._startup_disabled_names() == set()


def test_mode_comma_list(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4W_MODE", "references, com-live")
    _reset_to_lite()
    packs.apply_startup_mode()
    hidden = server._startup_disabled_names()
    assert "insert_citation" not in hidden
    assert "com_export_pdf" not in hidden
    assert "get_comments" in hidden


def test_mode_tolerates_lite_full_tokens(restore_enabled, monkeypatch):
    """The pptx round 1 M5 / round 2 L4 regressions: mode tokens inside
    comma lists must not brick startup."""
    monkeypatch.setenv("KS4W_MODE", "lite, references")
    _reset_to_lite()
    packs.apply_startup_mode()
    assert packs.is_tool_enabled("insert_citation")
    assert not packs.is_tool_enabled("get_comments")


def test_mode_typo_fails_loudly(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4W_MODE", "references,citatons")
    _reset_to_lite()
    with pytest.raises(WordMcpError):
        packs.apply_startup_mode()


def test_locked_policy_refuses(restore_enabled, monkeypatch):
    monkeypatch.setenv("KS4W_PACK_POLICY", "locked")
    with pytest.raises(WordMcpError) as exc:
        asyncio.run(server.enable_tools(["references"]))
    assert getattr(exc.value, "code", None) == "CONFLICT"


# --------------------------------------------- wire: toggle round trip


def test_toggle_round_trip_and_signpost(restore_enabled):
    """Startup lite -> disabled tool signposts pack + exact call ->
    enable -> visible + list_changed -> disable -> hidden -> signpost.
    Replicates main()'s startup wiring against fastmcp 3.4.7 reality
    (verified: disabled tools raise NotFoundError('Unknown tool'), the
    DisabledToolSignpost converts it; session visibility rules override
    the global transform and send ToolListChangedNotification)."""

    async def run():
        out = {}
        _reset_to_lite()
        server._PENDING_VISIBILITY.clear()
        transform = Visibility(
            False, names=server._startup_disabled_names()
        )
        server.mcp.add_transform(transform)
        notes: list[str] = []

        async def handler(message):
            notes.append(getattr(
                getattr(message, "root", None), "method", "?"))

        try:
            async with Client(server.mcp, message_handler=handler) as c:
                names = {t.name for t in await c.list_tools()}
                out["startup"] = names

                with pytest.raises(ToolError) as exc:
                    await c.call_tool("insert_citation", {
                        "file_path": "x.docx", "source_tag": "t",
                    })
                out["signpost"] = str(exc.value)

                res = await c.call_tool(
                    "enable_tools", {"packs": ["references"]})
                out["enable"] = res.structured_content
                await asyncio.sleep(0.1)
                out["notes_enable"] = list(notes)
                out["after_enable"] = {
                    t.name for t in await c.list_tools()}

                notes.clear()
                await c.call_tool(
                    "disable_tools", {"packs": ["references"]})
                await asyncio.sleep(0.1)
                out["notes_disable"] = list(notes)
                out["after_disable"] = {
                    t.name for t in await c.list_tools()}

                with pytest.raises(ToolError) as exc2:
                    await c.call_tool("insert_citation", {
                        "file_path": "x.docx", "source_tag": "t",
                    })
                out["signpost2"] = str(exc2.value)
        finally:
            server.mcp._transforms.remove(transform)
        return out

    out = asyncio.run(run())
    assert "enable_tools" in out["startup"]
    assert "insert_citation" not in out["startup"]
    assert "'references' pack" in out["signpost"]
    assert "enable_tools(packs=['references'])" in out["signpost"]

    payload = out["enable"]
    assert payload["enabled"] == ["references"]
    assert payload["approx_tokens_added"] > 0
    assert payload["active_tools"] == len(out["after_enable"])
    assert "notifications/tools/list_changed" in out["notes_enable"]
    assert "insert_citation" in out["after_enable"]

    assert "notifications/tools/list_changed" in out["notes_disable"]
    assert "insert_citation" not in out["after_disable"]
    assert "enable_tools(packs=['references'])" in out["signpost2"]


# ------------------------------------- discoverability contract stubs


def test_refusal_pack_hint_uses_real_registry(restore_enabled):
    """Rule 2 plumbing: a refusal declaring hint_tools resolves the pack
    from the REAL registry and spells the exact enable_tools call. The
    message text is never scanned."""
    _reset_to_lite()
    exc = WordMcpError("citation work needs the references tools")
    exc.hint_tools = ["insert_citation"]
    payload = envelope.refusal(exc)
    hint = payload["error"]["hint"]
    assert "insert_citation (pack 'references')" in hint
    assert "enable_tools(packs=['references'])" in hint

    # enabled tools do not hint (nothing to enable)
    packs._ENABLED["insert_citation"] = True
    assert envelope.pack_hint(exc) is None

    # message text mentioning a tool name must NOT trigger (pptx M8)
    exc2 = WordMcpError("try insert_citation for this")
    assert envelope.pack_hint(exc2) is None


def test_apply_edits_op_signpost_lands_on_pack(restore_enabled):
    """A pack tool name passed as an apply_edits op signposts its pack
    now that membership is real (the batch.py raise site sets
    hint_tools=[op])."""
    _reset_to_lite()
    exc = WordMcpError("op 'insert_citation' is not an apply_edits op")
    exc.hint_tools = ["insert_citation"]
    hint = envelope.refusal(exc)["error"]["hint"]
    assert "enable_tools(packs=['references'])" in hint


def test_workflow_recipes_name_packs():
    """Rule 4: get_workflows stays in lite; recipes name the packs their
    steps need; migrate-from-v1 and bulk-edit entries exist and reference
    registered tools."""
    assert packs.pack_of("get_workflows") == "lite"
    shipped = set(_shipped_tools())

    listing = server.get_workflows()
    tasks = {t["task"] for t in listing["tasks"]}
    assert {"migrate-from-v1", "bulk-edit"} <= tasks

    heavy = server.get_workflows("heavy-editing")
    assert "academic" in heavy["packs_required"]
    for step in heavy["steps"]:
        assert step["tool"] in shipped
        pack = packs.pack_of(step["tool"])
        if pack != "lite":
            assert step["pack"] == pack
        else:
            assert "pack" not in step
    assert "enable_tools(packs=" in heavy["note"]

    bulk = server.get_workflows("bulk-edit")
    for step in bulk["steps"]:
        assert step["tool"] in shipped


def test_lite_docstrings_carry_pack_adverts():
    """Rule 3: the budgeted advert lines landed where the design requires
    them (table family, style definition, TOC, validate battery,
    assembly, the list-then-act multiplex)."""
    tools = _shipped_tools()
    expected = {
        "get_table": "media-forms pack",
        "apply_style": "academic pack",
        "get_outline": "academic pack",
        "diagnose_document": "academic pack",
        "create_document": "assembly pack",
        "copy_document": "assembly pack",
        "create_table": "protection-io pack",
        "list_elements": "enable_tools",
    }
    for name, needle in expected.items():
        assert needle in (tools[name].description or ""), (
            f"{name} lost its pack-advert line"
        )


def test_lite_has_no_stand_in_for_demoted_tools():
    """Rule 1 spot check: the demoted names are really out of lite, so
    the only route back is the pack (no degraded stand-ins)."""
    demoted_academic = {
        "set_section_properties", "define_style", "set_page_numbers",
        "set_header_footer", "change_heading_level", "word_count",
    }
    demoted_media = {
        "insert_hyperlink", "modify_table_structure",
        "set_table_properties", "format_cells", "sort_table",
    }
    for name in demoted_academic:
        assert packs.pack_of(name) == "academic", name
    for name in demoted_media:
        assert packs.pack_of(name) == "media-forms", name


def test_enable_result_json_serializable():
    """The informed-approval payload must survive the wire as JSON."""
    report = packs.surface_report()
    json.dumps(report)
    assert report["active_tools"] >= 28
