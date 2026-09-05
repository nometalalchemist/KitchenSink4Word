"""v2 Phase 0: the ported pack registry and enable/disable machinery.

Pack MEMBERSHIP is wired in Phase 4; these tests drive the machinery with
dummy tools so the registry, env handling, and toggle bookkeeping are
proven importable and correct now, independent of server.py.
"""

from __future__ import annotations

import pytest

from word_mcp import packs
from word_mcp.core.errors import WordMcpError


class DummyTool:
    def __init__(self, description="x" * 400, parameters=None):
        self.description = description
        self.parameters = parameters or {
            "properties": {"file_path": {"type": "string"}}
        }


@pytest.fixture(autouse=True)
def _clean_registry(monkeypatch):
    """Registry state is process-global; snapshot and restore around each
    test, and neutralize the env contract."""
    monkeypatch.delenv("KS4W_MODE", raising=False)
    monkeypatch.delenv("KS4W_PACK_POLICY", raising=False)
    monkeypatch.delenv("KS4W_ALL_TOOLS", raising=False)
    monkeypatch.delenv("KS4W_LOCK_TOOLS", raising=False)
    saved_reg = {p: dict(t) for p, t in packs._REGISTRY.items()}
    saved_en = dict(packs._ENABLED)
    saved_hook = packs._visibility_hook
    packs._REGISTRY.clear()
    packs._REGISTRY.update({"lite": {}})
    packs._ENABLED.clear()
    packs.set_visibility_hook(None)
    yield
    packs._REGISTRY.clear()
    packs._REGISTRY.update({p: dict(t) for p, t in saved_reg.items()})
    packs._ENABLED.clear()
    packs._ENABLED.update(saved_en)
    packs._visibility_hook = saved_hook


@pytest.fixture
def populated():
    """A lite core plus two packs of dummy tools."""
    packs.register("lite_alpha", None, DummyTool())
    packs.register("lite_beta", None, DummyTool())
    packs.register("ref_one", "references", DummyTool())
    packs.register("ref_two", "references", DummyTool())
    packs.register("rev_one", "review", DummyTool())


def test_seven_packs_and_summaries():
    """The Section 14.8 seven packs, in menu order, summaries clean."""
    assert packs.pack_names() == [
        "references", "review", "academic", "assembly",
        "media-forms", "com-live", "protection-io",
    ]
    for name, summary in packs.PACK_SUMMARIES.items():
        assert summary, f"pack {name} has an empty summary"
        assert "—" not in summary, f"pack {name} summary has an em dash"


def test_register_unknown_pack_fails_loudly():
    with pytest.raises(ValueError):
        packs.register("tool_x", "typo-pack", DummyTool())


def test_register_tracks_enabled_state(populated):
    assert packs.is_tool_enabled("lite_alpha")
    assert not packs.is_tool_enabled("ref_one")
    assert packs.pack_of("ref_one") == "references"
    assert packs.pack_of("lite_alpha") == "lite"
    assert packs.pack_of("nope") is None


def test_enable_disable_idempotent(populated):
    r1 = packs.enable(["references"])
    assert r1["enabled"] == ["references"]
    assert r1["approx_tokens_added"] > 0
    assert packs.is_tool_enabled("ref_one")
    r2 = packs.enable(["references"])
    assert r2["enabled"] == []
    assert r2["already_enabled"] == ["references"]
    assert r2["approx_tokens_added"] == 0

    d1 = packs.disable(["references"])
    assert d1["disabled"] == ["references"]
    assert not packs.is_tool_enabled("ref_one")
    d2 = packs.disable(["references"])
    assert d2["already_disabled"] == ["references"]
    assert d2["approx_tokens_removed"] == 0


def test_enable_reports_surface(populated):
    before = packs.surface_report()["active_tools"]
    r = packs.enable(["references"])
    assert r["active_tools"] == before + 2
    assert r["approx_active_tokens"] > 0
    assert "packs" in r
    assert "note" in r  # list_changed advisory on a real change


def test_everything_alias(populated):
    r = packs.enable(["everything"])
    assert r["active_tools"] == 5  # every dummy registered above


def test_unknown_pack_lists_valid_names(populated):
    with pytest.raises(WordMcpError) as exc:
        packs.enable(["refrences"])
    msg = str(exc.value)
    for name in packs.pack_names():
        assert name in msg
    assert "everything" in msg


def test_lite_cannot_be_toggled(populated):
    with pytest.raises(WordMcpError):
        packs.enable(["lite"])
    with pytest.raises(WordMcpError):
        packs.disable(["lite"])


def test_locked_policy_refuses(populated, monkeypatch):
    monkeypatch.setenv("KS4W_PACK_POLICY", "locked")
    with pytest.raises(WordMcpError) as exc:
        packs.enable(["references"])
    assert getattr(exc.value, "code", None) == "CONFLICT"
    with pytest.raises(WordMcpError):
        packs.disable(["references"])


def test_startup_mode_lite_default(populated, monkeypatch):
    monkeypatch.delenv("KS4W_MODE", raising=False)
    assert packs.apply_startup_mode() == "lite"
    assert packs.surface_report()["active_tools"] == 2  # lite only


def test_startup_mode_full(populated, monkeypatch):
    monkeypatch.setenv("KS4W_MODE", "full")
    packs.apply_startup_mode()
    assert packs.surface_report()["active_tools"] == 5


def test_startup_mode_pack_list(populated, monkeypatch):
    monkeypatch.setenv("KS4W_MODE", "references, review")
    packs.apply_startup_mode()
    assert packs.is_tool_enabled("ref_one")
    assert packs.is_tool_enabled("rev_one")


def test_startup_mode_tolerates_lite_and_full_in_comma_lists(
    populated, monkeypatch
):
    """The two startup-mode fixes carried from pptx (round 1 M5, round 2
    L4): lite and full inside a comma list must not brick the server."""
    monkeypatch.setenv("KS4W_MODE", "lite, references")
    packs.apply_startup_mode()
    assert packs.is_tool_enabled("ref_one")
    assert not packs.is_tool_enabled("rev_one")

    monkeypatch.setenv("KS4W_MODE", "full, references")
    packs.apply_startup_mode()
    assert packs.is_tool_enabled("rev_one")


def test_startup_mode_bad_pack_fails_loudly(populated, monkeypatch):
    monkeypatch.setenv("KS4W_MODE", "references,typo-pack")
    with pytest.raises(WordMcpError):
        packs.apply_startup_mode()


def test_visibility_hook_mirrors_changes(populated):
    calls: list[tuple[set, bool]] = []
    packs.set_visibility_hook(lambda names, enabled: calls.append(
        (set(names), enabled)
    ))
    packs.enable(["references"])
    assert calls == [({"ref_one", "ref_two"}, True)]
    packs.disable(["references"])
    assert calls[-1] == ({"ref_one", "ref_two"}, False)
    # idempotent re-enable of nothing flips nothing: no hook call
    calls.clear()
    packs.disable(["references"])
    assert calls == []


def test_menu_and_costs(populated):
    menu = packs.menu()
    assert set(menu) == set(packs.pack_names())
    assert menu["references"]["tools"] == ["ref_one", "ref_two"]
    assert menu["references"]["approx_tokens"] > 0
    assert menu["academic"]["tools"] == []  # membership lands in Phase 4
