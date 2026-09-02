"""Measure the real tool surface: counts and approx token bills per pack.

Ported from KitchenSink4PPT. Imports the server (which registers every
tool) and prints, from the live registry, the per-pack tool count and token
estimate (description + JSON schema at ~4 chars/token, the same math
packs.approx_tokens uses for the informed-approval report), plus the lite
and full totals. README numbers come from running this, never from
hand-math.

Pack membership is wired (Phase 4); the "(unpacked)" row prints only if
a registration ever bypasses the packs registry, which the membership
integrity test forbids.

Run:  .venv/Scripts/python.exe -X utf8 scripts/measure_surface.py
"""

from __future__ import annotations

import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from word_mcp import packs, server  # noqa: E402


def _fmt_tokens(n: int) -> str:
    return f"{n / 1000:.1f}k"


def main() -> None:
    all_tools = {t.name: t for t in asyncio.run(server.mcp.list_tools())}
    names = packs.tool_names()
    packed = {n for members in names.values() for n in members}
    unpacked = sorted(set(all_tools) - packed)

    total_tools = 0
    total_tokens = 0
    print(f"{'pack':<24} {'tools':>5} {'~tokens':>8}")
    print("-" * 40)
    order = ["lite", *packs.pack_names()]
    for pack in order:
        members = names.get(pack, [])
        if not members:
            continue
        cost = sum(
            packs.approx_tokens(all_tools.get(n) or packs._REGISTRY[pack][n])
            for n in members
        )
        total_tools += len(members)
        total_tokens += cost
        print(f"{pack:<24} {len(members):>5} {_fmt_tokens(cost):>8}")
    if unpacked:
        cost = sum(packs.approx_tokens(all_tools[n]) for n in unpacked)
        total_tools += len(unpacked)
        total_tokens += cost
        print(f"{'(unpacked v1 surface)':<24} {len(unpacked):>5} "
              f"{_fmt_tokens(cost):>8}")
    print("-" * 40)
    print(f"{'TOTAL (full)':<24} {total_tools:>5} "
          f"{_fmt_tokens(total_tokens):>8}")

    lite = names.get("lite", [])
    if lite:
        lite_cost = sum(
            packs.approx_tokens(all_tools.get(n) or packs._REGISTRY['lite'][n])
            for n in lite
        )
        print()
        print(f"lite startup surface: {len(lite)} tools, "
              f"~{_fmt_tokens(lite_cost)} tokens")
    print(f"full surface (KS4W_MODE=full): {total_tools} tools, "
          f"~{_fmt_tokens(total_tokens)} tokens")

    # Description vs schema split, for the efficiency-accounting sections.
    desc = sum(len(t.description or "") for t in all_tools.values())
    import json
    schema = sum(
        len(json.dumps(t.parameters or {})) for t in all_tools.values()
    )
    print(f"split: ~{_fmt_tokens(round(desc / 4))} descriptions + "
          f"~{_fmt_tokens(round(schema / 4))} schemas")


if __name__ == "__main__":
    main()
