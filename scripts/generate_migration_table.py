"""Render the MIGRATION_V2.md rename table from migration/v1_to_v2.json.

The table in docs/MIGRATION_V2.md between the BEGIN/END markers is
GENERATED, never hand-written (V2_PLAN Phase 5 rule). Modes:

    python scripts/generate_migration_table.py            # print the block
    python scripts/generate_migration_table.py --write    # patch the guide

The docs-consistency test regenerates the block and diffs it against the
guide, so a stale table fails the suite.
"""

from __future__ import annotations

import io
import json
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GUIDE = ROOT / "docs" / "MIGRATION_V2.md"
BEGIN = "<!-- BEGIN GENERATED RENAME TABLE (scripts/generate_migration_table.py) -->"
END = "<!-- END GENERATED RENAME TABLE -->"


def _v2_call(entry: dict) -> str:
    """Render the v2 destination with its injected discriminators, e.g.
    manage_note(action="insert", note_type="endnote")."""
    inject = entry.get("inject") or {}
    args = ", ".join(f"{k}={json.dumps(v)}" for k, v in inject.items())
    return f"{entry['to']}({args})" if args else entry["to"]


def _params(entry: dict) -> str:
    moves = entry.get("params") or {}
    parts = [
        f"`{old}` -> `{new}`" for old, new in moves.items() if old != new
    ]
    kept = [old for old, new in moves.items() if old == new]
    if kept:
        parts.append("unchanged: " + ", ".join(f"`{k}`" for k in kept))
    return "; ".join(parts) if parts else "-"


def render() -> str:
    data = json.loads(
        (ROOT / "migration" / "v1_to_v2.json").read_text(encoding="utf-8")
    )
    tools: dict[str, dict] = data["tools"]

    groups: dict[str, list[str]] = defaultdict(list)
    for v1_name in sorted(tools):
        groups[tools[v1_name]["to"]].append(v1_name)

    buf = io.StringIO()
    w = buf.write
    w(BEGIN + "\n\n")
    w(f"All {len(tools)} v1.6 tools, grouped by their v2 destination "
      f"({len(groups)} v2 tools receive them). `inject` values are the "
      "literal v2 parameters that reproduce the v1 tool's fixed behavior; "
      "parameter moves use dot paths into nested objects (so "
      "`location.search.text` means `location={\"search\": {\"text\": "
      "...}}`).\n\n")
    for target in sorted(groups):
        members = groups[target]
        w(f"### -> `{target}`\n\n")
        w("| v1 tool | v2 call | parameter moves | notes |\n")
        w("|---|---|---|---|\n")
        for v1_name in members:
            entry = tools[v1_name]
            note = entry.get("notes", "-")
            w(f"| {v1_name} | `{_v2_call(entry)}` | {_params(entry)} "
              f"| {note} |\n")
        w("\n")
    w(END + "\n")
    return buf.getvalue()


def main() -> int:
    block = render()
    if "--write" in sys.argv:
        text = GUIDE.read_text(encoding="utf-8")
        start = text.find(BEGIN)
        stop = text.find(END)
        if start < 0 or stop < 0:
            raise SystemExit(f"markers not found in {GUIDE}")
        new = text[:start] + block + text[stop + len(END) + 1:]
        GUIDE.write_text(new, encoding="utf-8", newline="\n")
        print(f"patched {GUIDE}")
    else:
        sys.stdout.write(block)
    return 0


if __name__ == "__main__":
    sys.exit(main())
