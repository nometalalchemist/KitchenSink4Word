"""v2 migration map: structure now, completeness in Phase 5.

The skeleton's structural checks are enforced from Phase 0; the
completeness assertion (every shipped v1 tool name has an entry, which is
how "no functionality lost" is enforced mechanically) is xfail until the
map is finished. Phase 5 flip: delete the pytest.mark.xfail line.
"""

from __future__ import annotations

import asyncio
import json
from pathlib import Path

import pytest

from word_mcp import server

MAP_PATH = (
    Path(__file__).resolve().parents[2] / "migration" / "v1_to_v2.json"
)


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
    (identity rows aside, a typo here would silently strand migrators)."""
    shipped = {t.name for t in asyncio.run(server.mcp.list_tools())}
    for v1_name in _load()["tools"]:
        assert v1_name in shipped, (
            f"map entry {v1_name!r} is not a shipped v1 tool"
        )


@pytest.mark.xfail(reason="map completed in Phase 5", strict=False)
def test_map_completeness():
    """All 189 shipped v1 tools appear in the map. This is the mechanical
    'no functionality lost' gate; it flips to required in Phase 5."""
    shipped = {t.name for t in asyncio.run(server.mcp.list_tools())}
    mapped = set(_load()["tools"])
    missing = sorted(shipped - mapped)
    assert not missing, (
        f"{len(missing)} shipped v1 tools missing from the migration map: "
        f"{missing[:10]}..."
    )
