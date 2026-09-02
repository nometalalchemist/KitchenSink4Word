"""Generate docs/internal/live_parity_v2.md: the v2 live-routing parity
table (V2_PLAN Phase 5, V2_DESIGN 5.4).

For every v2 tool: its live behavior class (derived from the server.py
AST), the v1.6 constituents that fold into it (from migration/v1_to_v2.json)
with each constituent's v1 live class (from the v1.6 tree via git show),
and a REGRESSION verdict: any v1 live-routed capability whose v2 home is
not live-routed is a defect.

Run from the repo root:
    ./.venv/Scripts/python.exe -X utf8 scripts/generate_live_parity.py
"""

from __future__ import annotations

import ast
import io
import json
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Verified per-tool live-mode caveats (sub-features that stay file-mode,
# guards, and bespoke refusal text). These annotate the scripted
# classification; the class itself always comes from the AST.
NOTES: dict[str, str] = {
    "insert_paragraphs": (
        "live: heading_level maps to built-in Heading styles by numeric "
        "constant (outline-based docs get direct outline levels); "
        "inherit_format/copy_format_from file-mode only; text-selector "
        "locations run the snapshot staleness guard"
    ),
    "set_paragraph_text": (
        "live: expect guard honored; tracked-revision paragraphs refuse; "
        "text-selector locations run the snapshot staleness guard"
    ),
    "delete_paragraphs": (
        "live: section-break/field-crossing ranges refuse; range={} "
        "location endpoints run the snapshot staleness guard"
    ),
    "search_and_replace": (
        "live: plain items route live (255-char finds handled); preview "
        "and find_formatting modes are file-mode only (find_formatting "
        "with live='force' refuses)"
    ),
    "format_text": (
        "live: formatting mode routes live; case mode file-mode only "
        "(live='force' refuses); range={} locations run the snapshot "
        "staleness guard"
    ),
    "set_paragraph_format": (
        "live: shading/borders/tab_stops keys refuse live (XML-level); "
        "raw indices keep the v1 index-trust contract"
    ),
    "set_cells": "live: plain text cells; vertical merges refuse live",
    "get_text": (
        "live: body shape identical; include_textboxes/textbox modes "
        "file-mode only"
    ),
    "find_text": (
        "live: plain queries, 500-match cap; formatting mode and "
        "include_textboxes file-mode only"
    ),
    "word_count": (
        "live: Word's own statistics engine; exclusions mode file-mode "
        "only"
    ),
    "get_document_info": (
        "live: same key names; adds words/track_revisions, omits part "
        "list"
    ),
    "get_comments": "live: same shape",
    "get_outline": "live: same flat-list shape",
    "apply_edits": (
        "live: one COM undo group; markdown lists/pipe tables in insert "
        "ops refuse up front with a close-the-file hint; replace verifies "
        "matched text in the live range; all other index-addressed ops "
        "run the snapshot staleness guard against the last saved state"
    ),
    "get_document_view": (
        "reads the last SAVED state of an open document and says so in "
        "the result (live: true + note); stamp_anchors refuses while the "
        "file is open (mutation)"
    ),
    "diagnose_document": (
        "no live route BY DESIGN: bespoke DocumentLocked message names "
        "the stale-XML reason and points to com_validate_opens_clean / "
        "live get_document_info"
    ),
}


def _load_v1_source() -> str:
    out = subprocess.run(
        ["git", "show", "main:src/word_mcp/server.py"],
        cwd=ROOT, capture_output=True, text=True, encoding="utf-8",
    )
    if out.returncode != 0:
        raise SystemExit(f"git show failed: {out.stderr}")
    return out.stdout


def _is_tool(fn: ast.AST) -> str | None:
    """Pack tag for @_tool("pack") / marker for @mcp.tool, else None."""
    for d in fn.decorator_list:
        if isinstance(d, ast.Call):
            f = d.func
            if isinstance(f, ast.Name) and f.id == "_tool":
                if d.args and isinstance(d.args[0], ast.Constant):
                    return str(d.args[0].value)
                return "?"
            if isinstance(f, ast.Attribute) and f.attr == "tool":
                return "v1"
        elif isinstance(d, ast.Attribute) and d.attr == "tool":
            return "v1"
    return None


def _analyze(src: str) -> dict[str, dict]:
    tree = ast.parse(src)
    tools: dict[str, dict] = {}
    for node in ast.walk(tree):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        pack = _is_tool(node)
        if pack is None:
            continue
        routed = any(
            isinstance(sub, ast.Call)
            and (
                (isinstance(sub.func, ast.Name)
                 and sub.func.id == "_route_live")
                or (isinstance(sub.func, ast.Attribute)
                    and sub.func.attr == "_route_live")
            )
            for sub in ast.walk(node)
        )
        doc = ast.get_docstring(node) or ""
        reader = "Read-only" in doc
        tools[node.name] = {
            "pack": pack,
            "routed": routed,
            "reader": reader,
        }
    return tools


# Tools whose behavior class the AST heuristic cannot see (verified by
# reading the implementations).
CLASS_OVERRIDES: dict[str, str] = {
    "get_document_view": (
        "open-document reader (no refusal: reads the last SAVED state via "
        "a shared-read snapshot)"
    ),
    "diagnose_document": (
        "file-only reader (locked file refuses, stale-XML honesty)"
    ),
    "get_workflows": "not applicable (static recipes, no document input)",
    "create_document": "not applicable (creates a new file)",
    "copy_document": (
        "not applicable (byte copy; Word grants shared reads, so it works "
        "while the document is open)"
    ),
}


def _classify(name: str, info: dict) -> str:
    if name in CLASS_OVERRIDES:
        return CLASS_OVERRIDES[name]
    if name.startswith("live_"):
        return "LIVE-NATIVE (live tier: requires the open document)"
    if name.startswith("com_"):
        return "COM tier (drives the Word application directly)"
    if info["routed"]:
        return "LIVE-ROUTED (live='auto' dual mode)"
    if info["reader"]:
        return "file-only reader (locked file refuses, stale-XML honesty)"
    return "file-only writer (locked file refuses until closed)"


def main() -> int:
    v2 = _analyze(
        (ROOT / "src" / "word_mcp" / "server.py").read_text(encoding="utf-8")
    )
    v1 = _analyze(_load_v1_source())
    migration = json.loads(
        (ROOT / "migration" / "v1_to_v2.json").read_text(encoding="utf-8")
    )["tools"]

    constituents: dict[str, list[str]] = defaultdict(list)
    for v1_name, entry in migration.items():
        constituents[entry["to"]].append(v1_name)

    regressions: list[str] = []
    buf = io.StringIO()
    w = buf.write
    w("# v2 live-routing parity table\n\n")
    w("GENERATED by scripts/generate_live_parity.py. Do not hand-edit;\n")
    w("regenerate after any live-routing change. Sources: the v2 server\n")
    w("AST (live class), migration/v1_to_v2.json (constituents), and the\n")
    w("v1.6 tree at main (constituent live class).\n\n")
    w("Policy (V2_DESIGN 5.4): every merged/renamed tool preserves EVERY\n")
    w("constituent's live behavior; every read tool has a live route or a\n")
    w("documented stale-XML refusal reason. File-only tools refuse locked\n")
    w("files with the standard DocumentLocked message, which names the\n")
    w("live alternates (com_save_document / com_close_open_document) and\n")
    w("states that the call has no live route.\n\n")

    counts = defaultdict(int)
    w("| v2 tool | pack | v2 live behavior | v1 constituents "
      "(v1-live starred) | regression |\n")
    w("|---|---|---|---|---|\n")
    for name in sorted(v2):
        info = v2[name]
        cls = _classify(name, info)
        counts[cls.split(" (")[0]] += 1
        cons = sorted(constituents.get(name, []))
        cons_cells = []
        regression = "none"
        for c in cons:
            v1_info = v1.get(c)
            if v1_info is None:
                cons_cells.append(f"{c} (NOT IN v1 TREE?)")
                regression = "UNMAPPED CONSTITUENT"
                continue
            if v1_info["routed"]:
                cons_cells.append(f"**{c}\\***")
                if not info["routed"] and not name.startswith(
                        ("com_", "live_")):
                    regression = f"LOST LIVE ROUTE from {c}"
            else:
                cons_cells.append(c)
        if regression != "none":
            regressions.append(f"{name}: {regression}")
        note = NOTES.get(name)
        behavior = cls if not note else f"{cls}. {note}"
        w(f"| {name} | {info['pack']} | {behavior} | "
          f"{', '.join(cons_cells) or '(new in v2)'} | {regression} |\n")

    w("\n## Totals (scripted)\n\n")
    for k in sorted(counts):
        w(f"- {k}: {counts[k]}\n")
    w(f"- v2 tools total: {len(v2)}\n")
    v1_routed = sorted(k for k, v in v1.items() if v["routed"])
    w(f"- v1 live-routed tools ({len(v1_routed)}): "
      f"{', '.join(v1_routed)}\n")
    v2_routed = sorted(k for k, v in v2.items() if v["routed"])
    w(f"- v2 live-routed tools ({len(v2_routed)}): "
      f"{', '.join(v2_routed)}\n")
    w(f"\n## Regressions\n\n")
    if regressions:
        for r in regressions:
            w(f"- DEFECT: {r}\n")
    else:
        w("None. Every v1 live-routed tool's capability has a live-routed "
          "v2 home; the one rename (replace_paragraph_text to "
          "set_paragraph_text) stays routed; apply_edits adds a live "
          "route v1 never had.\n")

    out = ROOT / "docs" / "internal" / "live_parity_v2.md"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(buf.getvalue(), encoding="utf-8", newline="\n")
    print(f"wrote {out} ({len(v2)} v2 tools, "
          f"{len(regressions)} regression(s))")
    return 1 if regressions else 0


if __name__ == "__main__":
    sys.exit(main())
