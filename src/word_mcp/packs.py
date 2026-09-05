"""Tiered loading: the pack registry and the enable/disable machinery.

Ported from KitchenSink4PPT packs.py (proven in production) and adapted to
fastmcp 3.x. Every tool is registered with FastMCP up front; non-lite tools
start disabled so a fresh session pays for the lite core, not the whole
surface. enable_tools flips packs on mid-session; clients re-fetch
tools/list on notifications/tools/list_changed.

fastmcp 3.x adaptation (the pptx original rode fastmcp 2.14, where
Tool.enable()/disable() flipped a per-tool flag and queued the
list_changed notification; 3.0 REMOVED Tool.enable/disable): this module
now keeps its own per-tool enabled bookkeeping (authoritative for
surface_report and the informed-approval token math) and mirrors every
change through an injectable visibility hook. server.py wires the hook to
the fastmcp 3.x visibility API (Phase 4, VERIFIED against an in-process
fastmcp 3.4.7 client session):
- startup surface: main() applies apply_startup_mode() to bookkeeping,
  then ONE global transform, mcp.add_transform(Visibility(False,
  names=disabled)) (3.4.7 has no server.disable method);
- mid-session toggles: session-scoped ctx.enable_components /
  ctx.disable_components(names={...}), whose rules override the global
  transform (mark-based, later marks win) and send
  ToolListChangedNotification to the session only (observed on the wire:
  tools/resources/prompts list_changed all fire per toggle).

Env contract:
- KS4W_MODE: startup surface for clients without reliable list_changed.
  "lite" (default), "full", or a comma-separated pack list
  ("references,com-live").
- KS4W_PACK_POLICY: "auto" (default; the CLIENT's permission prompt gates
  enable_tools, which is deliberately a plain tool call) or "locked"
  (enable_tools/disable_tools refuse; the surface is fixed at startup).

The BOOLEAN toggles behind the .mcpb checkboxes (see parse_toggle). Each
accepts the literal "true"/"false" Claude Desktop writes, treats empty and
absent as off, and refuses to start on anything else:
- KS4W_ALL_TOOLS: "true" loads every pack at startup, "false" or empty
  keeps the lite default. KS4W_MODE wins when it is set to a non-empty
  value, so a power user's pack list is never overridden by a checkbox.
- KS4W_LOCK_TOOLS: "true" fixes the surface at startup (enable_tools and
  disable_tools refuse), "false" or empty leaves it adjustable.
  KS4W_PACK_POLICY wins when it is set to a non-empty value.

There is deliberately no per-pack startup toggle: see toggle_env_names.
startup_note() names which setting decided the surface and
apply_startup_mode() writes that line to stderr.

No persistence, by design: every session starts at KS4W_MODE.

server.py populates the registry via register(); this module never imports
FastMCP itself and holds only the Tool objects it is handed (for the
token-cost math) plus its own enabled flags.

Pack membership was finalized in Phase 4 against the consolidated
108-tool surface (plus enable_tools/disable_tools, registered under lite
so they are visible in every mode); server.py's @_tool decorator is the
single source of membership truth.
"""

from __future__ import annotations

import json
import os
import sys
from typing import Callable

from .core.errors import WordMcpError

# v2 packs in menu order, per V2_DESIGN Section 14.8 (author-firmed seven
# packs; granularity per the 2026-09-02 addendum). "lite" is the always-on
# core, not a pack. "everything" is a convenience alias for all seven.
PACK_SUMMARIES: dict[str, str] = {
    "references": (
        "citations and bibliography: Word-native sources, bibliography "
        "styles, Zotero search/cite, parity checks, style conversion"
    ),
    "review": (
        "tracked changes (read, accept/reject), comments (add, reply, "
        "resolve), reviewer reports, structured diff, anonymize"
    ),
    "academic": (
        "notes, TOC, index, captions, cross-references, bookmarks, "
        "headers/footers, sections, styles, word counts, validation, "
        "submission prep, accessibility"
    ),
    "assembly": (
        "multi-document work: insert/split documents, move sections, "
        "copy tables across files, templates, mail merge"
    ),
    "media-forms": (
        "images, charts, equations, text boxes, hyperlinks, table "
        "structure (merge/split cells, rows, columns, sort), forms, "
        "content controls, field codes"
    ),
    "com-live": (
        "drives the Word app: PDF import/export, compare/combine, "
        "proofing, readability, live editing of open documents"
    ),
    "protection-io": (
        "protection, watermarks, redaction with verification, table "
        "CSV/JSON import/export"
    ),
}
EVERYTHING = "everything"

# pack -> {tool_name: fastmcp Tool}; "lite" holds the always-on core.
_REGISTRY: dict[str, dict[str, object]] = {"lite": {}}

# tool_name -> currently enabled? Authoritative bookkeeping (fastmcp 3.x
# has no per-tool enabled flag to read back; see module docstring).
_ENABLED: dict[str, bool] = {}

# Injected by server.py: called as _visibility_hook(names, enabled) after
# every state change so the FastMCP surface mirrors the registry. None
# means bookkeeping only (unit tests, measurement scripts).
_visibility_hook: Callable[[set[str], bool], None] | None = None


def set_visibility_hook(hook: Callable[[set[str], bool], None] | None) -> None:
    """server.py wires this to the fastmcp 3.x visibility API."""
    global _visibility_hook
    _visibility_hook = hook


def _sync(names: set[str], enabled: bool) -> None:
    if _visibility_hook is not None and names:
        _visibility_hook(names, enabled)


def register(tool_name: str, pack: str | None, tool: object) -> None:
    """Called by server.py once per tool at import time. pack=None means
    the lite core (enabled at startup); anything else starts disabled."""
    key = pack or "lite"
    if key != "lite" and key not in PACK_SUMMARIES:
        raise ValueError(f"unknown pack {key!r} for tool {tool_name}")
    _REGISTRY.setdefault(key, {})[tool_name] = tool
    _ENABLED[tool_name] = key == "lite"


def pack_names() -> list[str]:
    return list(PACK_SUMMARIES)


def pack_tools(pack: str) -> list[str]:
    return sorted(_REGISTRY.get(pack, {}))


def pack_of(tool_name: str) -> str | None:
    for pack, tools in _REGISTRY.items():
        if tool_name in tools:
            return pack
    return None


def is_tool_enabled(tool_name: str) -> bool:
    return _ENABLED.get(tool_name, False)


def tool_names() -> dict[str, list[str]]:
    return {pack: sorted(tools) for pack, tools in _REGISTRY.items()}


def approx_tokens(tool: object) -> int:
    """Rough per-tool client cost: description + JSON schema at ~4 chars per
    token. Honest enough for the informed-approval report; not a billing
    meter."""
    desc = getattr(tool, "description", "") or ""
    try:
        schema = json.dumps(getattr(tool, "parameters", {}) or {})
    except (TypeError, ValueError):
        schema = ""
    return round((len(desc) + len(schema)) / 4)


def pack_cost(pack: str) -> int:
    return sum(approx_tokens(t) for t in _REGISTRY.get(pack, {}).values())


def surface_report() -> dict:
    """Current active surface: enabled tool count and approx token bill."""
    active = 0
    tokens = 0
    per_pack: dict[str, str] = {}
    for pack, tools in _REGISTRY.items():
        enabled = [t for n, t in tools.items() if _ENABLED.get(n, False)]
        active += len(enabled)
        tokens += sum(approx_tokens(t) for t in enabled)
        per_pack[pack] = f"{len(enabled)}/{len(tools)} enabled"
    return {
        "active_tools": active,
        "approx_active_tokens": tokens,
        "packs": per_pack,
    }


#: The startup surface env for power users: "lite", "full", or a pack list.
ENV_MODE = "KS4W_MODE"

#: The surface-lock env for power users: "auto" (default) or "locked".
ENV_PACK_POLICY = "KS4W_PACK_POLICY"

#: The positively-named boolean toggles behind the .mcpb user_config
#: checkboxes. Claude Desktop writes the LITERAL strings "true" and "false"
#: for a user_config boolean, so both are honored with the meaning the
#: checkbox shows the human: ticked means the thing the name says.
ENV_ALL_TOOLS = "KS4W_ALL_TOOLS"
ENV_LOCK_TOOLS = "KS4W_LOCK_TOOLS"

_TRUE = ("1", "true", "on", "yes")
_FALSE = ("0", "false", "off", "no")


def parse_toggle(name: str, value: str | bool | None) -> bool:
    """Resolve one boolean launch toggle, POSITIVE polarity: the value reads
    the way the Desktop checkbox does, so "true" means the thing the env name
    says.

    An EMPTY value FAILS CLOSED to False, never to the enabling side: an
    empty env is what an unconfigured host writes, and an empty string that
    silently turned a setting ON would be a fail-open defect. Unrecognized
    values are an ERROR, never a shrug, because guessing at a typo would
    silently change the tool surface the operator asked for."""
    if value is None or value is False:
        return False
    if value is True:
        return True
    text = str(value).strip().lower()
    if not text:
        return False  # empty NEVER enables
    if text in _TRUE:
        return True
    if text in _FALSE:
        return False
    raise WordMcpError(
        f"unknown {name} value {value!r}: use 'true' or 'false' "
        f"(also accepted: {', '.join(_TRUE)} / {', '.join(_FALSE)}). "
        f"Refusing to start rather than guessing, because guessing here "
        f"would silently change which tools this server offers."
    )


def toggle(name: str) -> bool:
    """Read one boolean toggle from the environment. Absent or empty is
    False; garbage raises."""
    return parse_toggle(name, os.environ.get(name))


def _explicit(name: str) -> str:
    """The power-user env's value when it is set to something non-empty.
    An unset or blank value is not a choice and defers to the toggle."""
    return os.environ.get(name, "").strip()


def toggle_env_names() -> list[str]:
    """Every boolean launch env this server reads.

    There is deliberately NO per-pack toggle (author ruling, 2026-09-05).
    Claude Desktop's own per-tool permissions already own the consent
    layer, and enable_tools already lets the agent load a pack mid-session
    on demand, so a startup checkbox per pack would duplicate both. The
    two toggles here are the ones neither of those can express: how big the
    surface starts, and whether it may change at all."""
    return [ENV_ALL_TOOLS, ENV_LOCK_TOOLS]


def validate_toggles() -> None:
    """Parse every boolean env at startup so a typo refuses LOUDLY, even in
    a toggle the precedence rules end up ignoring. Ignored means ignored for
    the DECISION, not unchecked: a misspelled value is an operator mistake
    whichever variable it lands in."""
    for name in toggle_env_names():
        toggle(name)


def resolve_lock() -> bool:
    """Is the tool surface fixed at startup?

    Precedence: an explicit KS4W_PACK_POLICY beats KS4W_LOCK_TOOLS beats
    the unlocked default."""
    explicit = _explicit(ENV_PACK_POLICY)
    if explicit:
        return explicit.lower() == "locked"
    return toggle(ENV_LOCK_TOOLS)


def _policy_locked() -> bool:
    return resolve_lock()


def _validate(packs: list[str]) -> list[str]:
    if isinstance(packs, str):
        packs = [packs]
    if not isinstance(packs, list) or not packs:
        raise WordMcpError(
            f"packs must be a non-empty list from {pack_names()} "
            f"(or ['{EVERYTHING}'])"
        )
    out: list[str] = []
    for p in packs:
        name = str(p).strip().lower()
        if name == EVERYTHING:
            return list(PACK_SUMMARIES)
        if name == "lite":
            raise WordMcpError(
                "the lite core is always on; it cannot be enabled or "
                "disabled as a pack"
            )
        if name not in PACK_SUMMARIES:
            raise WordMcpError(
                f"unknown pack {p!r}; valid packs: {pack_names()} "
                f"(or '{EVERYTHING}' for all of them)"
            )
        if name not in out:
            out.append(name)
    return out


def enable(packs: list[str]) -> dict:
    """Idempotent enable. Reports what changed, the approx token cost added,
    and the resulting total surface."""
    if _policy_locked():
        err = WordMcpError(
            "the tool surface is fixed at startup by the host "
            "(KS4W_PACK_POLICY=locked, or the 'Lock the tool set at "
            "startup' setting). Only a human can change it: untick that "
            "setting, or restart the server with a different KS4W_MODE."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    enabled_now: list[str] = []
    already: list[str] = []
    tokens_added = 0
    flipped: set[str] = set()
    for pack in wanted:
        newly = False
        for name, tool in _REGISTRY.get(pack, {}).items():
            if not _ENABLED.get(name, False):
                _ENABLED[name] = True
                flipped.add(name)
                tokens_added += approx_tokens(tool)
                newly = True
        (enabled_now if newly else already).append(pack)
    _sync(flipped, True)
    result = {
        "enabled": enabled_now,
        "already_enabled": already,
        "approx_tokens_added": tokens_added,
        **surface_report(),
    }
    if enabled_now:
        result["note"] = (
            "tools/list_changed was sent; re-fetch the tool list if your "
            "client does not refresh automatically"
        )
    return result


def disable(packs: list[str]) -> dict:
    """Idempotent disable; the lite core always stays on."""
    if _policy_locked():
        err = WordMcpError(
            "KS4W_PACK_POLICY=locked: the tool surface is fixed at startup "
            "by the host."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    disabled_now: list[str] = []
    already: list[str] = []
    tokens_removed = 0
    flipped: set[str] = set()
    for pack in wanted:
        newly = False
        for name, tool in _REGISTRY.get(pack, {}).items():
            if _ENABLED.get(name, False):
                _ENABLED[name] = False
                flipped.add(name)
                tokens_removed += approx_tokens(tool)
                newly = True
        (disabled_now if newly else already).append(pack)
    _sync(flipped, False)
    return {
        "disabled": disabled_now,
        "already_disabled": already,
        "approx_tokens_removed": tokens_removed,
        **surface_report(),
    }


def resolve_startup_mode() -> str:
    """The startup surface in force, before any pack is flipped.

    Precedence, highest first: KS4W_MODE set to something non-empty (the
    power-user pin wins outright and the checkbox is ignored, because a
    host that spells out a pack list has said exactly what it wants), then
    KS4W_ALL_TOOLS=true, then the lite default."""
    explicit = _explicit(ENV_MODE)
    if explicit:
        return explicit.lower()
    return "full" if toggle(ENV_ALL_TOOLS) else "lite"


def startup_note() -> str:
    """One line naming what decided the startup surface, written to stderr
    at startup. A surprising tool list should name its own cause, and the
    case that most needs saying out loud is a KS4W_MODE pin silently
    overriding checkboxes a human ticked in an installer."""
    explicit = _explicit(ENV_MODE)
    if explicit:
        note = f"startup surface from {ENV_MODE}={explicit!r}"
        if toggle(ENV_ALL_TOOLS):
            return (
                f"{note}; the 'Load every tool at startup' setting "
                f"({ENV_ALL_TOOLS}) is IGNORED while it is set"
            )
        return note
    if toggle(ENV_ALL_TOOLS):
        return f"startup surface: every pack, from {ENV_ALL_TOOLS}=true"
    return "startup surface: lite core only (no startup toggle set)"


def apply_startup_mode() -> str:
    """Apply the resolved startup surface at server start (before the event
    loop; no client is connected yet, so the visibility hook runs without a
    session and the server-side wiring must use global transforms, not
    session state). Returns the mode applied, for logging.

    Every boolean toggle is parsed here so a typo in any of them refuses
    LOUDLY before the server serves a single request, and the resolution is
    announced on stderr (stdout carries the protocol and stays clean)."""
    validate_toggles()
    resolve_lock()
    mode = resolve_startup_mode()
    sys.stderr.write(f"[kitchensink4word] {startup_note()}\n")
    if not mode or mode == "lite":
        return "lite"
    # "lite" and "full"/"everything" are mode tokens, tolerated inside
    # comma lists alike: lite is always on anyway, full means every pack.
    # Refusing lite bricked the server at startup (pptx round 1 M5);
    # refusing full did the same the other way (pptx round 2 L4). Typos
    # still fail LOUDLY via _validate below.
    tokens = [p.strip() for p in mode.split(",") if p.strip()]
    wants_full = any(t in ("full", EVERYTHING) for t in tokens)
    named = [t for t in tokens if t not in ("lite", "full", EVERYTHING)]
    if named:
        _validate(named)  # raises on typos so a bad env fails LOUDLY
    packs = list(PACK_SUMMARIES) if wants_full else named
    if not packs:
        return "lite"
    valid = _validate(packs)
    flipped: set[str] = set()
    for pack in valid:
        for name in _REGISTRY.get(pack, {}):
            if not _ENABLED.get(name, False):
                _ENABLED[name] = True
                flipped.add(name)
    _sync(flipped, True)
    return mode


def menu() -> dict:
    """The full pack menu with per-pack tool lists and approx token costs."""
    return {
        pack: {
            "summary": PACK_SUMMARIES[pack],
            "tools": pack_tools(pack),
            "approx_tokens": pack_cost(pack),
        }
        for pack in PACK_SUMMARIES
    }
