"""Phase 5 docs consistency: the generated rename table in
docs/MIGRATION_V2.md must be current against migration/v1_to_v2.json
(regenerate-and-diff, so a stale guide fails red), and the live-parity
table must agree with the server's actual routing."""

from __future__ import annotations

import ast
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
GUIDE = ROOT / "docs" / "MIGRATION_V2.md"
PARITY = ROOT / "docs" / "internal" / "live_parity_v2.md"
BEGIN = ("<!-- BEGIN GENERATED RENAME TABLE "
         "(scripts/generate_migration_table.py) -->")
END = "<!-- END GENERATED RENAME TABLE -->"


def _routed_tools() -> set[str]:
    """v2 tools that call _route_live, from the server AST."""
    src = (ROOT / "src" / "word_mcp" / "server.py").read_text(
        encoding="utf-8"
    )
    routed = set()
    for node in ast.walk(ast.parse(src)):
        if not isinstance(node, ast.FunctionDef):
            continue
        deco = any(
            isinstance(d, ast.Call) and isinstance(d.func, ast.Name)
            and d.func.id == "_tool"
            for d in node.decorator_list
        )
        if not deco:
            continue
        if any(
            isinstance(sub, ast.Call) and (
                (isinstance(sub.func, ast.Name)
                 and sub.func.id == "_route_live")
            )
            for sub in ast.walk(node)
        ):
            routed.add(node.name)
    return routed


def test_rename_table_is_current():
    out = subprocess.run(
        [sys.executable, "-X", "utf8",
         str(ROOT / "scripts" / "generate_migration_table.py")],
        capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
    )
    assert out.returncode == 0, out.stderr
    fresh = out.stdout.replace("\r\n", "\n").strip()
    text = GUIDE.read_text(encoding="utf-8")
    start, stop = text.find(BEGIN), text.find(END)
    assert start >= 0 and stop >= 0, "generated-table markers missing"
    embedded = text[start:stop + len(END)].replace("\r\n", "\n").strip()
    assert embedded == fresh, (
        "docs/MIGRATION_V2.md rename table is stale; rerun "
        "scripts/generate_migration_table.py --write"
    )


def test_guide_has_no_em_dashes():
    assert "—" not in GUIDE.read_text(encoding="utf-8")


def test_guide_live_list_matches_server_routing():
    """The migration guide's live-routed tool list stays truthful."""
    routed = _routed_tools()
    text = GUIDE.read_text(encoding="utf-8")
    m = re.search(
        r"Live-routed v2 tools: (.*?)\.\n\n", text, re.S
    )
    assert m, "live-routed list missing from the guide"
    listed = set(re.findall(r"`(\w+)`", m.group(1)))
    listed.discard("replace_paragraph_text")  # named as the old v1 name
    assert listed == routed, (
        f"guide says {sorted(listed)}, server routes {sorted(routed)}"
    )


def test_parity_table_matches_server_routing():
    """Spot checks on docs/internal/live_parity_v2.md: routed set current,
    row per tool, no regression entries."""
    routed = _routed_tools()
    text = PARITY.read_text(encoding="utf-8")
    m = re.search(r"- v2 live-routed tools \((\d+)\): (.+)", text)
    assert m, "routed-tools summary line missing"
    listed = {t.strip() for t in m.group(2).split(",")}
    assert int(m.group(1)) == len(routed) and listed == routed, (
        "parity table routed set is stale; rerun "
        "scripts/generate_live_parity.py"
    )
    rows = re.findall(r"^\| (\w+) \|", text, re.M)
    # one row per registered tool (minus the header row's 'v2' cell)
    tool_rows = [r for r in rows if r != "v2"]
    assert len(tool_rows) == len(set(tool_rows)) >= 100
    assert "DEFECT:" not in text, "parity table records a live regression"
    assert "—" not in text
