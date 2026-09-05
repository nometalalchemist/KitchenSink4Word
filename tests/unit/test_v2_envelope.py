"""v2 Phase 0: the ported refusal envelope and discoverability signpost.

envelope.py is standalone (never imports server.py), so these tests prove
the machinery ready for the Phase 2 server rebuild without touching v1.
"""

from __future__ import annotations

import asyncio

import pytest

from word_mcp import envelope, packs
from word_mcp.core import errors as err
from word_mcp.core.safesave import MutationLockTimeout
from word_mcp.core.sandbox import SandboxViolation


class DummyTool:
    def __init__(self):
        self.description = "d" * 400
        self.parameters = {"properties": {}}


@pytest.fixture(autouse=True)
def _clean_registry(monkeypatch):
    monkeypatch.delenv("KS4W_MODE", raising=False)
    monkeypatch.delenv("KS4W_PACK_POLICY", raising=False)
    monkeypatch.delenv("KS4W_ALL_TOOLS", raising=False)
    monkeypatch.delenv("KS4W_LOCK_TOOLS", raising=False)
    saved_reg = {p: dict(t) for p, t in packs._REGISTRY.items()}
    saved_en = dict(packs._ENABLED)
    yield
    packs._REGISTRY.clear()
    packs._REGISTRY.update({p: dict(t) for p, t in saved_reg.items()})
    packs._ENABLED.clear()
    packs._ENABLED.update(saved_en)


def test_all_mapped_codes_are_closed_vocabulary():
    for _etype, code in envelope.CODE_MAP:
        assert code in envelope.CLOSED_CODES, (
            f"{code} is outside the closed vocabulary"
        )


def test_every_closed_code_has_a_hint_or_is_deliberate():
    """Hints exist for every code an exception can map to today;
    STALE_ANCHOR and RANGE_OUT_OF_BOUNDS hints are pre-staged for the
    Phase 1/3 raisers; BAD_PARAMS and CONFLICT carry their message."""
    for code in envelope.CLOSED_CODES - {"BAD_PARAMS", "CONFLICT",
                                         "UNSUPPORTED_CONTENT"}:
        assert code in envelope.HINTS, f"{code} has no hint"


@pytest.mark.parametrize("exc,code", [
    (err.AmbiguousTarget("two matches"), "AMBIGUOUS_LOCATION"),
    (err.TargetNotFound("no such paragraph"), "NOT_FOUND"),
    (err.DocumentNotFound("missing.docx"), "NOT_FOUND"),
    (MutationLockTimeout("locked"), "DOCUMENT_LOCKED"),
    (err.DocumentLocked("open in Word"), "DOCUMENT_LOCKED"),
    (err.DocumentCorrupt("bad zip"), "UNSUPPORTED_CONTENT"),
    (err.ValidationFailed("package invalid"), "VALIDATION_FAILED"),
    (err.ProtectedViewRefused("protected"), "PROTECTED_VIEW"),
    (err.WordNotRunning("no instance"), "APP_NOT_RUNNING"),
    (err.WordBusy("dialog open"), "APP_BUSY"),
    (err.WordBlocked("hung"), "APP_BLOCKED"),
    (err.WordDisconnected("closed mid-call"), "CONFLICT"),
    (SandboxViolation("outside root"), "BAD_PARAMS"),
    (err.WordMcpError("generic"), "BAD_PARAMS"),
    (FileExistsError("exists"), "CONFLICT"),
    (FileNotFoundError("gone"), "NOT_FOUND"),
    (ValueError("bad"), "BAD_PARAMS"),
    (RecursionError("deep"), "UNSUPPORTED_CONTENT"),
])
def test_classification(exc, code):
    out = envelope.refusal(exc)
    assert out["ok"] is False
    assert out["error"]["code"] == code
    assert out["error"]["code"] in envelope.CLOSED_CODES


def test_explicit_code_attribute_wins():
    exc = err.WordMcpError("surface locked")
    exc.code = "CONFLICT"
    assert envelope.refusal(exc)["error"]["code"] == "CONFLICT"


def test_short_lookup_error_message_is_named():
    out = envelope.refusal(KeyError(0))
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "internal lookup failed" in out["error"]["message"]


def test_detail_rides_along():
    exc = err.AmbiguousTarget("3 matches for 'Introduction'")
    exc.detail = {"candidates": [4, 17, 92]}
    out = envelope.refusal(exc)
    assert out["error"]["detail"] == {"candidates": [4, 17, 92]}


def test_refusal_result_is_dict_like_and_iserror():
    """The fastmcp 3.x RefusalResult contract: mapping access for
    in-process callers, isError=true on the wire, payload intact in both
    content and structuredContent."""
    exc = err.TargetNotFound("paragraph 99 does not exist")
    res = envelope.refuse(exc)
    assert isinstance(res, envelope.RefusalResult)
    assert res["ok"] is False
    assert res["error"]["code"] == "NOT_FOUND"
    assert "error" in res
    assert res.get("missing", "dflt") == "dflt"

    mcp_result = res.to_mcp_result()
    assert mcp_result.isError is True
    assert mcp_result.structuredContent["error"]["code"] == "NOT_FOUND"
    assert "paragraph 99" in mcp_result.content[0].text


def test_pack_hint_names_enable_tools():
    """Discoverability rule 2: a refusal that DECLARES it directs the
    caller to a disabled tool (exc.hint_tools, set at the raise site) must
    carry the exact enable_tools call. Message text is never scanned."""
    packs.register("insert_citation", "references", DummyTool())
    exc = err.UnsupportedStructure(
        "citations go through insert_citation, not raw field edits"
    )
    exc.hint_tools = ["insert_citation"]
    out = envelope.refusal(exc)
    hint = out["error"]["hint"]
    assert "enable_tools" in hint
    assert "references" in hint


def test_pack_hint_absent_when_pack_enabled():
    packs.register("insert_citation", "references", DummyTool())
    packs.enable(["references"])
    exc = err.UnsupportedStructure("citations go through insert_citation")
    exc.hint_tools = ["insert_citation"]
    out = envelope.refusal(exc)
    assert "enable_tools" not in out["error"]["hint"]


def test_pack_hint_ignores_message_text():
    """pptx M8: tool names appearing in the MESSAGE alone never trigger
    the hint; only the structured hint_tools attribute does."""
    packs.register("insert_citation", "references", DummyTool())
    exc = err.UnsupportedStructure(
        "citations go through insert_citation, not raw field edits"
    )
    out = envelope.refusal(exc)
    assert "enable_tools" not in out["error"]["hint"]


def test_pack_hint_skips_lite_and_unknown_tools():
    packs.register("get_text_lite", None, DummyTool())
    exc = err.UnsupportedStructure("use get_text_lite or wholly_unknown")
    exc.hint_tools = ["get_text_lite", "wholly_unknown"]
    assert envelope.pack_hint(exc) is None


def test_disabled_tool_signpost_middleware():
    """A tools/call to a registered but disabled tool must name the pack,
    not dead-end as Unknown tool."""
    from fastmcp.exceptions import NotFoundError, ToolError

    packs.register("insert_citation", "references", DummyTool())
    mw = envelope.DisabledToolSignpost()

    class Ctx:
        class message:
            name = "insert_citation"

    async def call_next(_ctx):
        raise NotFoundError("Unknown tool: insert_citation")

    with pytest.raises(ToolError) as exc:
        asyncio.run(mw.on_call_tool(Ctx(), call_next))
    msg = str(exc.value)
    assert "references" in msg
    assert "enable_tools" in msg

    class CtxUnknown:
        class message:
            name = "never_registered"

    with pytest.raises(NotFoundError):
        asyncio.run(mw.on_call_tool(CtxUnknown(), call_next))


def test_no_em_dashes_in_envelope_strings():
    for code, hint in envelope.HINTS.items():
        assert "—" not in hint, f"hint for {code} has an em dash"
