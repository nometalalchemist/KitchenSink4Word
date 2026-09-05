"""The .mcpb launch toggles: KS4W_ALL_TOOLS and KS4W_LOCK_TOOLS.

Claude Desktop writes the LITERAL strings "true" and "false" for a
user_config boolean, so the parser is tested against exactly those first.
The three properties that matter, and that this file pins:

1. POLARITY. "true" means the thing the env name says. The checkbox and
   the variable read the same direction, so a human ticking "Load every
   tool at startup" gets every tool.
2. FAIL CLOSED. An absent or empty value resolves to False, never to the
   enabling side. An empty string is what an unconfigured host writes, and
   an empty value that silently switched something on is a fail-open
   defect (the sibling web server shipped that bug once and it is not
   being reproduced here).
3. REFUSE GARBAGE. An unrecognized value raises at STARTUP, naming the
   accepted values, rather than shrugging into a default the operator did
   not ask for.

Plus precedence: KS4W_MODE beats the master checkbox and KS4W_PACK_POLICY
beats the lock checkbox, in both directions, so a power user's pin survives
whatever an installer wrote. Section 5 pins the author's ruling that there
are exactly two toggles, and section 7 pins the manifest against the code.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from word_mcp import packs, server  # noqa: F401  (server populates the registry)
from word_mcp.core.errors import WordMcpError

TOGGLES = (packs.ENV_ALL_TOOLS, packs.ENV_LOCK_TOOLS)


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    for name in (packs.ENV_MODE, packs.ENV_PACK_POLICY,
                 *packs.toggle_env_names()):
        monkeypatch.delenv(name, raising=False)


@pytest.fixture(autouse=True)
def _restore_enabled():
    """The enabled bookkeeping is process-global; apply_startup_mode writes
    to it, so snapshot and restore around every test."""
    saved = dict(packs._ENABLED)
    saved_hook = packs._visibility_hook
    packs.set_visibility_hook(None)
    yield
    packs._ENABLED.clear()
    packs._ENABLED.update(saved)
    packs.set_visibility_hook(saved_hook)


# --------------------------------------------------------- 1. polarity


@pytest.mark.parametrize("name", TOGGLES)
def test_literal_true_is_true(monkeypatch, name):
    """The exact string Claude Desktop writes for a ticked checkbox."""
    monkeypatch.setenv(name, "true")
    assert packs.toggle(name) is True


@pytest.mark.parametrize("name", TOGGLES)
def test_literal_false_is_false(monkeypatch, name):
    """The exact string Claude Desktop writes for an unticked checkbox."""
    monkeypatch.setenv(name, "false")
    assert packs.toggle(name) is False


@pytest.mark.parametrize("value", ["true", "TRUE", " True ", "1", "on", "yes"])
def test_true_spellings(value):
    assert packs.parse_toggle("KS4W_ALL_TOOLS", value) is True


@pytest.mark.parametrize("value", ["false", "FALSE", " False ", "0", "off",
                                   "no"])
def test_false_spellings(value):
    assert packs.parse_toggle("KS4W_ALL_TOOLS", value) is False


# ------------------------------------------------------ 2. fail closed


@pytest.mark.parametrize("value", [None, "", "   ", "\t"])
def test_empty_fails_closed(value):
    """Empty NEVER enables. The whole point of the property."""
    assert packs.parse_toggle("KS4W_ALL_TOOLS", value) is False


@pytest.mark.parametrize("name", TOGGLES)
def test_absent_is_false(name):
    assert packs.toggle(name) is False


@pytest.mark.parametrize("name", TOGGLES)
def test_empty_env_is_false(monkeypatch, name):
    monkeypatch.setenv(name, "")
    assert packs.toggle(name) is False


def test_empty_all_tools_stays_lite(monkeypatch):
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "")
    assert packs.resolve_startup_mode() == "lite"


def test_empty_lock_tools_stays_unlocked(monkeypatch):
    monkeypatch.setenv(packs.ENV_LOCK_TOOLS, "")
    assert packs.resolve_lock() is False


# --------------------------------------------------- 3. refuse garbage


@pytest.mark.parametrize("value", ["ture", "enabled", "full", "2", "-1",
                                   "y e s", "null", "none"])
def test_garbage_refuses(value):
    with pytest.raises(WordMcpError) as exc:
        packs.parse_toggle("KS4W_ALL_TOOLS", value)
    message = str(exc.value)
    assert "KS4W_ALL_TOOLS" in message
    assert "true" in message and "false" in message


@pytest.mark.parametrize("name", TOGGLES)
def test_garbage_refuses_at_startup(monkeypatch, name):
    """Not at first use. A typo must stop the server before it serves."""
    monkeypatch.setenv(name, "yess")
    with pytest.raises(WordMcpError):
        packs.apply_startup_mode()


# ------------------------------------------------------- 4. precedence


def test_all_tools_true_means_full():
    assert packs.resolve_startup_mode() == "lite"


def test_all_tools_toggle_selects_full(monkeypatch):
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    assert packs.resolve_startup_mode() == "full"


def test_explicit_mode_beats_toggle(monkeypatch):
    """A power user's pack list is never overridden by the checkbox."""
    monkeypatch.setenv(packs.ENV_MODE, "references,com-live")
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    assert packs.resolve_startup_mode() == "references,com-live"


def test_explicit_mode_beats_toggle_downward(monkeypatch):
    monkeypatch.setenv(packs.ENV_MODE, "lite")
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    assert packs.resolve_startup_mode() == "lite"


def test_blank_mode_defers_to_toggle(monkeypatch):
    """A blank KS4W_MODE is not a choice; the checkbox still decides."""
    monkeypatch.setenv(packs.ENV_MODE, "  ")
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    assert packs.resolve_startup_mode() == "full"


def test_lock_toggle_locks(monkeypatch):
    monkeypatch.setenv(packs.ENV_LOCK_TOOLS, "true")
    assert packs.resolve_lock() is True


def test_explicit_policy_beats_lock_toggle(monkeypatch):
    monkeypatch.setenv(packs.ENV_PACK_POLICY, "auto")
    monkeypatch.setenv(packs.ENV_LOCK_TOOLS, "true")
    assert packs.resolve_lock() is False


def test_explicit_policy_locks_over_false_toggle(monkeypatch):
    monkeypatch.setenv(packs.ENV_PACK_POLICY, "locked")
    monkeypatch.setenv(packs.ENV_LOCK_TOOLS, "false")
    assert packs.resolve_lock() is True


def test_lock_toggle_refuses_enable_tools(monkeypatch):
    """The toggle reaches the same refusal the power-user env does."""
    monkeypatch.setenv(packs.ENV_LOCK_TOOLS, "true")
    with pytest.raises(WordMcpError) as exc:
        packs.enable(["references"])
    assert getattr(exc.value, "code", None) == "CONFLICT"


# ------------------------------- 5. no per-pack toggles, by decision


def test_only_two_toggles_exist():
    """Author ruling, 2026-09-05: no per-pack startup toggle. Claude
    Desktop's own per-tool permissions own the consent layer and
    enable_tools already loads a pack mid-session on demand, so a checkbox
    per pack would duplicate both. This asserts the decision so a future
    session does not quietly re-add them."""
    assert packs.toggle_env_names() == [packs.ENV_ALL_TOOLS,
                                        packs.ENV_LOCK_TOOLS]
    assert not hasattr(packs, "pack_env_names")


# ------------------------------------------------- 6. the startup note


def test_note_names_the_default():
    assert "lite core only" in packs.startup_note()


def test_note_names_the_master_toggle(monkeypatch):
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    assert packs.ENV_ALL_TOOLS in packs.startup_note()


def test_note_names_the_mode_pin(monkeypatch):
    monkeypatch.setenv(packs.ENV_MODE, "academic")
    assert "academic" in packs.startup_note()


def test_note_says_the_mode_pin_ignored_the_checkbox(monkeypatch):
    """The case that most needs saying out loud: a human ticked the box in
    an installer and a KS4W_MODE pin quietly overrode it."""
    monkeypatch.setenv(packs.ENV_MODE, "lite")
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    note = packs.startup_note()
    assert packs.ENV_MODE in note
    assert "IGNORED" in note


def test_master_toggle_applies_for_real(monkeypatch):
    """End to end through apply_startup_mode: the master checkbox alone
    opens every pack."""
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    for name in packs._ENABLED:
        packs._ENABLED[name] = packs.pack_of(name) == "lite"
    assert packs.apply_startup_mode() == "full"
    for pack in packs.PACK_SUMMARIES:
        tools = packs.pack_tools(pack)
        assert tools, f"{pack} registered no tools; the check would be vacuous"
        for name in tools:
            assert packs.is_tool_enabled(name), f"{pack}/{name}"


# ------------------------------------------- 7. manifest <-> code parity


def _manifest() -> dict:
    path = Path(__file__).resolve().parents[2] / "bundle" / "manifest.json"
    return json.loads(path.read_text(encoding="utf-8"))


def test_manifest_declares_every_toggle():
    """A pack added in code without a checkbox is invisible to installers,
    and a checkbox with no env behind it does nothing. Neither drifts
    silently."""
    manifest = _manifest()
    declared = set(manifest["server"]["mcp_config"]["env"])
    assert declared == set(packs.toggle_env_names())


def test_manifest_env_values_point_at_real_user_config():
    manifest = _manifest()
    config = manifest["user_config"]
    for env, ref in manifest["server"]["mcp_config"]["env"].items():
        key = ref.removeprefix("${user_config.").removesuffix("}")
        assert key in config, f"{env} points at a missing checkbox {key!r}"
        assert config[key]["type"] == "boolean"
        assert config[key]["default"] is False, "every checkbox ships off"


def test_manifest_checkbox_copy_is_a_sentence():
    for key, entry in _manifest()["user_config"].items():
        assert entry["title"], key
        assert entry["description"].endswith("."), key
        assert "_" not in entry["description"], f"{key}: no env names in copy"


def test_note_reaches_stderr(monkeypatch, capsys):
    monkeypatch.setenv(PACK_ENVS["review"], "true")
    packs.apply_startup_mode()
    captured = capsys.readouterr()
    assert captured.out == "", "stdout carries the protocol and stays clean"
    assert PACK_ENVS["review"] in captured.err


def test_note_reaches_stderr(monkeypatch, capsys):
    monkeypatch.setenv(packs.ENV_ALL_TOOLS, "true")
    packs.apply_startup_mode()
    captured = capsys.readouterr()
    assert captured.out == "", "stdout carries the protocol and stays clean"
    assert packs.ENV_ALL_TOOLS in captured.err
