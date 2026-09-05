"""The v2 refusal envelope: structured refusals with isError=true and the
pack-hint discoverability contract.

Ported from KitchenSink4PPT server.py (the _RefusalResult machinery and the
hint_tools signpost, both proven through the pptx production round) as a
standalone module so the Phase 2 server rebuild can wire it in without
touching the v1 server. Nothing here imports server.py.

Contract:
- Refusals are structured: {ok: false, error: {code, message, hint}} with a
  closed code vocabulary, never raw tracebacks.
- RefusalResult serializes over MCP with isError=true (production-test
  finding: refusals riding out as isError=false successes) while staying
  indexable like a dict for in-process callers and the test harness.
- Discoverability rule 2: a refusal that directs the caller to tools that
  exist but are disabled names the pack and the exact enable_tools call.
  The raise site declares the tools via exc.hint_tools; the message text is
  NEVER scanned, because a message echoing user input that happens to match
  a tool name must not trigger the hint (pptx Phase 8 finding M8).

fastmcp 3.x adaptation: the pptx original (fastmcp 2.14) made RefusalResult
inherit from BOTH dict and ToolResult; in 3.x ToolResult is a pydantic
model, so dual inheritance hits an instance lay-out conflict. RefusalResult
now subclasses ToolResult alone (is_error=True at construction) and exposes
the mapping protocol over structured_content, which preserves both halves
of the old contract.

Phase 4 verification (fastmcp 3.4.7, in-process client): a tools/call to
a visibility-disabled tool has get_tool return None, and the server
raises NotFoundError("Unknown tool: ...") exactly as 2.14 did, so
DisabledToolSignpost's except clause holds unchanged; the signpost was
observed on the wire converting it into the pack-naming ToolError.
"""

from __future__ import annotations

import json as _json
from typing import Any
from xml.etree.ElementTree import ParseError as _XmlParseError

from fastmcp.exceptions import NotFoundError as _FmcpNotFound
from fastmcp.exceptions import ToolError as _FmcpToolError
from fastmcp.server.middleware import Middleware as _FmcpMiddleware
from fastmcp.tools.tool import ToolResult as _FmcpToolResult
from lxml import etree as _lxml_etree

from . import packs as _packs
from .core import errors as _err
from .core.safesave import MutationLockTimeout
from .core.sandbox import SandboxViolation

# Order matters: specific classes before their WordMcpError base.
CODE_MAP: tuple[tuple[type[BaseException], str], ...] = (
    (_err.AmbiguousTarget, "AMBIGUOUS_LOCATION"),
    (_err.StaleAnchor, "STALE_ANCHOR"),
    (_err.RangeOutOfBounds, "RANGE_OUT_OF_BOUNDS"),
    (_err.TargetNotFound, "NOT_FOUND"),
    (_err.DocumentNotFound, "NOT_FOUND"),
    (MutationLockTimeout, "DOCUMENT_LOCKED"),
    (_err.DocumentLocked, "DOCUMENT_LOCKED"),
    (_err.DocumentCorrupt, "UNSUPPORTED_CONTENT"),
    (_err.DocumentProtected, "UNSUPPORTED_CONTENT"),
    (_err.UnsupportedStructure, "UNSUPPORTED_CONTENT"),
    (_err.ValidationFailed, "VALIDATION_FAILED"),
    (_err.ProtectedViewRefused, "PROTECTED_VIEW"),
    (_err.WordNotRunning, "APP_NOT_RUNNING"),
    (_err.DocumentNotOpenInWord, "APP_NOT_RUNNING"),
    (_err.LiveLockTimeout, "APP_BUSY"),
    (_err.WordBusy, "APP_BUSY"),
    (_err.WordBlocked, "APP_BLOCKED"),
    (_err.WordDisconnected, "CONFLICT"),
    (SandboxViolation, "BAD_PARAMS"),
    (_err.WordMcpError, "BAD_PARAMS"),
    (FileExistsError, "CONFLICT"),
    (FileNotFoundError, "NOT_FOUND"),
    (ValueError, "BAD_PARAMS"),
    (TypeError, "BAD_PARAMS"),
    # Deliberate widening carried from the pptx hardening rounds: parser
    # and recursion errors from hostile XML input must refuse in-envelope,
    # never surface as raw FastMCP tool errors. Ops-level guards refuse
    # first with better messages; these are the backstop.
    (_XmlParseError, "BAD_PARAMS"),
    (_lxml_etree.LxmlError, "BAD_PARAMS"),
    (AttributeError, "BAD_PARAMS"),
    (RecursionError, "UNSUPPORTED_CONTENT"),
    (OverflowError, "BAD_PARAMS"),
    # A dict indexed with the wrong key type otherwise dies as a raw
    # KeyError repr ("0") on the client. Ops-level guards refuse first
    # with real messages; this catches stragglers.
    (KeyError, "BAD_PARAMS"),
    (IndexError, "BAD_PARAMS"),
)
CATCHABLE = tuple(t for t, _ in CODE_MAP)

# The closed code vocabulary. STALE_ANCHOR and RANGE_OUT_OF_BOUNDS are
# raised by the Phase 1 locate resolver and the Phase 3 view/batch layer.
CLOSED_CODES = frozenset({
    "AMBIGUOUS_LOCATION", "NOT_FOUND", "DOCUMENT_LOCKED", "APP_NOT_RUNNING",
    "APP_BUSY", "APP_BLOCKED", "PROTECTED_VIEW", "VALIDATION_FAILED",
    "STALE_ANCHOR", "RANGE_OUT_OF_BOUNDS", "UNSUPPORTED_CONTENT",
    "CONFLICT", "BAD_PARAMS",
})

HINTS: dict[str, str] = {
    "AMBIGUOUS_LOCATION": (
        "several targets matched; use the unambiguous address from the "
        "candidates in the message"
    ),
    "NOT_FOUND": (
        "re-run get_document_view or get_outline to see current content, "
        "indices, and anchors"
    ),
    "STALE_ANCHOR": (
        "the document changed since the view was taken; re-run "
        "get_document_view and resend with fresh anchors. If you made "
        "live edits on an open document, com_save_document first "
        "(anchors resolve from the last SAVED state)"
    ),
    "RANGE_OUT_OF_BOUNDS": (
        "the range exceeds the document; the message names the valid "
        "bounds"
    ),
    "DOCUMENT_LOCKED": (
        "close the file in Word (or wait out the other process) and "
        "retry; dual-mode tools accept live='auto' to edit the open copy "
        "instead"
    ),
    "VALIDATION_FAILED": (
        "the original file was NOT modified; the message says what the "
        "produced package failed"
    ),
    "APP_NOT_RUNNING": "this operation needs Word installed and reachable",
    "APP_BUSY": (
        "Word is showing a dialog or running a command; clear it and "
        "retry"
    ),
    "APP_BLOCKED": "Word is not answering; wait or restart it",
    "PROTECTED_VIEW": "click Enable Editing in Word first",
}


def classify(exc: BaseException) -> str:
    for etype, code in CODE_MAP:
        if isinstance(exc, etype):
            return code
    return "BAD_PARAMS"


def pack_hint(exc: BaseException) -> str | None:
    """If a refusal explicitly directs the caller to tools that exist but
    are disabled, say exactly how to turn them on (discoverability rule 2:
    the refusal IS the signpost). The raise site declares the tools via
    exc.hint_tools; the message text is NEVER scanned (pptx M8)."""
    names = getattr(exc, "hint_tools", None)
    if not names:
        return None
    needed: dict[str, str] = {}
    for name in names:
        pack = _packs.pack_of(name)
        if pack in (None, "lite"):
            continue
        if not _packs.is_tool_enabled(name):
            needed[name] = pack
    if not needed:
        return None
    pack_list = sorted(set(needed.values()))
    named = ", ".join(f"{n} (pack {p!r})" for n, p in sorted(needed.items()))
    return (
        f"the tool(s) named here are registered but currently disabled: "
        f"{named}. Call enable_tools(packs={pack_list}) to turn them on."
    )


def refusal(exc: BaseException) -> dict:
    """Build the {ok: false, error: {code, message, hint}} payload."""
    code = getattr(exc, "code", None) or classify(exc)
    message = str(exc)
    if isinstance(exc, LookupError) and len(message) < 40:
        # A bare KeyError/IndexError repr ("0") is useless on its own;
        # name the failure mode. Ops-level guards should refuse first.
        message = (
            f"internal lookup failed on {message}: a nested parameter "
            "probably has the wrong shape (list where a dict belongs, or "
            "vice versa)"
        )
    hint = HINTS.get(code, "")
    ph = pack_hint(exc)
    if ph:
        hint = f"{hint} {ph}".strip()
    error: dict[str, Any] = {"code": code, "message": message, "hint": hint}
    # Ambiguity refusals from the locate resolver (Section 6.2) carry every
    # match on the exception; surface them per the declared refusal shape.
    matches = getattr(exc, "matches", None)
    if matches:
        error["matches"] = matches
    detail = getattr(exc, "detail", None)
    if detail:
        error["detail"] = detail
    return {"ok": False, "error": error}


class RefusalResult(_FmcpToolResult):
    """A structured refusal that is BOTH indexable like the
    {ok: false, error: ...} dict (in-process callers and the test harness)
    AND a FastMCP ToolResult whose MCP serialization sets isError=true, so
    spec-compliant clients see the failure flag. The JSON payload stays
    intact in the content AND in structuredContent; only the flag changes.
    See the module docstring for the fastmcp 3.x shape of this class."""

    def __init__(self, payload: dict):
        text = _json.dumps(payload, indent=2, ensure_ascii=False)
        super().__init__(
            content=text, structured_content=payload, is_error=True
        )

    # Mapping protocol over the payload, replacing the 2.14 dict base.
    def __getitem__(self, key):
        return self.structured_content[key]

    def __contains__(self, key) -> bool:
        return key in self.structured_content

    def get(self, key, default=None):
        return self.structured_content.get(key, default)

    def keys(self):
        return self.structured_content.keys()


def refuse(exc: BaseException) -> RefusalResult:
    """One-call convenience for the Phase 2 tool wrapper."""
    return RefusalResult(refusal(exc))


class DisabledToolSignpost(_FmcpMiddleware):
    """Discoverability rule 2 at the transport layer (pptx M4 fix): a
    tools/call to a registered but currently disabled tool must name the
    owning pack and the exact enable_tools call, not dead-end with a bare
    "Unknown tool"."""

    async def on_call_tool(self, context, call_next):
        try:
            return await call_next(context)
        except _FmcpNotFound as exc:
            name = getattr(context.message, "name", "")
            pack = _packs.pack_of(name)
            if pack and pack != "lite" and not _packs.is_tool_enabled(name):
                raise _FmcpToolError(
                    f"tool {name!r} exists but is currently disabled: it "
                    f"belongs to the {pack!r} pack. Call "
                    f"enable_tools(packs=['{pack}']) to turn it on, "
                    "then retry this call."
                ) from exc
            raise
