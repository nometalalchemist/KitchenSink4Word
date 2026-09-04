"""Measure the v1.6 operation count with the v2 definition, for comparison.

The public v1.6 headline was 189 TOOLS, which is not the same yardstick as
the v2 headline of distinct OPERATIONS. This script produces the missing
number: what v1.6 measures when counted exactly the way v2 is counted, so
the two figures can sit in one sentence honestly.

How the methodology carries over, detector by detector:

  1. SOURCE DISPATCH SCAN: identical. The same AST scan runs against
     v1.6's server.py and its ops modules, reading the same validation
     shapes (membership tests, equality chains, table lookups, values
     validated one layer down in an ops module, and dispatch keys read
     out of caller-supplied dicts). Only the decorator changes: v1.6
     registers with @mcp.tool(), v2 with @_tool("pack").

  2. CAPABILITY GROUND TRUTH: on v2 this is the migration map, which
     credits each v2 tool with the v1.6 tools folded into it. On v1.6
     that reduces to 1 per tool, because the map's source side is one
     entry per v1.6 registration: its 189 keys are exactly the 189
     functions v1.6 decorates with @mcp.tool (verified by this script
     before it reports). So each v1.6 tool contributes 1, expanded by
     its own dispatch values where it has them. That is --migration unit.

The result is symmetric. One capability counts once on the tree where it
was registered and once on the tree where it survives, and neither tree
gets credit the other is denied. v1.6's pack toggles do not exist (packs
are a v2 feature), so the exclusion list is empty there and removes two
tools on the v2 side only.

Run against an existing v1.6 checkout:
    .venv/Scripts/python.exe -X utf8 scripts/count_operations_v16.py PATH

Or let it make and remove its own worktree from a ref (default: main):
    .venv/Scripts/python.exe -X utf8 scripts/count_operations_v16.py --ref main
"""

from __future__ import annotations

import argparse
import ast
import json
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
COUNTER = ROOT / "scripts" / "count_operations.py"
MIGRATION_JSON = ROOT / "migration" / "v1_to_v2.json"


def _v16_tool_names(checkout: Path) -> list[str]:
    tree = ast.parse(
        (checkout / "src" / "word_mcp" / "server.py").read_text(encoding="utf-8"))
    names = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef):
            continue
        for dec in node.decorator_list:
            target = dec.func if isinstance(dec, ast.Call) else dec
            if isinstance(target, ast.Attribute) and target.attr == "tool":
                names.append(node.name)
    return names


def _verify_ground_truth(checkout: Path) -> None:
    """The comparison only holds if the map's source side IS v1.6's surface."""
    tools = set(_v16_tool_names(checkout))
    mapped = set(json.loads(MIGRATION_JSON.read_text(encoding="utf-8"))["tools"])
    missing = sorted(tools - mapped)
    extra = sorted(mapped - tools)
    print(f"v1.6 registrations: {len(tools)}   migration map entries: {len(mapped)}")
    if missing or extra:
        print("GROUND TRUTH MISMATCH, the comparison is not apples to apples:")
        if missing:
            print(f"  registered in v1.6, absent from the map: {missing}")
        if extra:
            print(f"  in the map, not registered in v1.6:      {extra}")
        raise SystemExit(2)
    print("ground truth verified: every v1.6 tool is one map entry, and back\n")


def _measure(checkout: Path, json_out: str | None) -> int:
    cmd = [sys.executable, "-X", "utf8", str(COUNTER),
           "--root", str(checkout),
           "--decorator", "mcp.tool",
           "--migration", "unit",
           "--label", "v1.6.0 measured with the v2 definition"]
    if json_out:
        cmd += ["--json", json_out]
    return subprocess.call(cmd)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("checkout", nargs="?",
                    help="path to a v1.6 checkout (skips the worktree dance)")
    ap.add_argument("--ref", default="main",
                    help="git ref holding the v1.6 tree (default: main)")
    ap.add_argument("--json", dest="json_out",
                    help="write the per-tool snapshot to this path")
    args = ap.parse_args()

    if args.checkout:
        checkout = Path(args.checkout).resolve()
        _verify_ground_truth(checkout)
        return _measure(checkout, args.json_out)

    tmp = Path(tempfile.mkdtemp(prefix="ks4w_v16_")) / "tree"
    subprocess.check_call(
        ["git", "-C", str(ROOT), "worktree", "add", "--detach",
         str(tmp), args.ref])
    try:
        _verify_ground_truth(tmp)
        return _measure(tmp, args.json_out)
    finally:
        subprocess.call(
            ["git", "-C", str(ROOT), "worktree", "remove", "--force", str(tmp)])


if __name__ == "__main__":
    raise SystemExit(main())
