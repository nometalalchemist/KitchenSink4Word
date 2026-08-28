"""Smoke test for the KitchenSink4Word .mcpb bundle.

Checks, in order:
  1. The .mcpb unzips and contains exactly manifest.json + icon.png.
  2. manifest.json parses, carries the required fields, declares the
     binary/uvx launcher shape, pins the PyPI version to the manifest
     version, and has NO tools array (Smithery cli bug #787).
  3. If the mcpb CLI is on PATH, `mcpb validate` passes on the extracted
     manifest.
  4. If uvx is available (PATH or next to this interpreter), launches
     `uvx kitchensink4word==<ver>` and performs a real MCP stdio
     initialize handshake, proving the launch command Claude Desktop
     will run actually starts the server.

Usage: python -X utf8 smoke_test.py [path/to/kitchensink4word.mcpb]
Exit code 0 = all mandatory checks passed (steps 3-4 skip with a note
when the tool is absent).
"""

import json
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
FAILURES = []


def check(label: str, ok: bool, detail: str = "") -> None:
    mark = "PASS" if ok else "FAIL"
    print(f"[{mark}] {label}" + (f" - {detail}" if detail else ""))
    if not ok:
        FAILURES.append(label)


def find_uvx() -> str | None:
    hit = shutil.which("uvx")
    if hit:
        return hit
    beside = Path(sys.executable).parent / ("uvx.exe" if os.name == "nt" else "uvx")
    return str(beside) if beside.exists() else None


def png_size(data: bytes) -> tuple[int, int]:
    w, h = struct.unpack(">II", data[16:24])
    return w, h


def main() -> int:
    mcpb_path = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "kitchensink4word.mcpb"
    print(f"Bundle: {mcpb_path}\n")
    if not mcpb_path.exists():
        check("bundle file exists", False, str(mcpb_path))
        return 1
    check("bundle file exists", True, f"{mcpb_path.stat().st_size} bytes")

    with tempfile.TemporaryDirectory() as td:
        # 1. unzip + contents
        with zipfile.ZipFile(mcpb_path) as zf:
            names = sorted(zf.namelist())
            zf.extractall(td)
        check("archive contents", names == ["icon.png", "manifest.json"], str(names))

        icon = (Path(td) / "icon.png").read_bytes()
        check("icon is a PNG", icon[:8] == b"\x89PNG\r\n\x1a\n")
        w, h = png_size(icon)
        check("icon is 512x512", (w, h) == (512, 512), f"{w}x{h}")

        # 2. manifest shape
        mpath = Path(td) / "manifest.json"
        m = json.loads(mpath.read_text(encoding="utf-8"))
        for field in ("manifest_version", "name", "version", "description", "author", "server"):
            check(f"manifest has {field}", field in m)
        check("name", m.get("name") == "kitchensink4word", m.get("name"))
        srv = m.get("server", {})
        check("server.type is binary", srv.get("type") == "binary", srv.get("type"))
        cfg = srv.get("mcp_config", {})
        check("command is uvx", cfg.get("command") == "uvx", str(cfg.get("command")))
        ver = m.get("version", "")
        expected_arg = f"kitchensink4word=={ver}"
        check(
            "args pin PyPI version to manifest version",
            cfg.get("args") == [expected_arg],
            str(cfg.get("args")),
        )
        check("no tools array (Smithery bug #787)", "tools" not in m)
        check("license is AGPL-3.0-only",
              m.get("license") == "AGPL-3.0-only", str(m.get("license")))

        # 3. mcpb validate, if available
        mcpb_cli = shutil.which("mcpb")
        if mcpb_cli:
            r = subprocess.run(
                [mcpb_cli, "validate", str(mpath)],
                capture_output=True, text=True, timeout=60, shell=False,
            )
            check("mcpb validate", r.returncode == 0,
                  (r.stdout + r.stderr).strip().splitlines()[-1] if (r.stdout or r.stderr) else "")
        else:
            print("[SKIP] mcpb validate - mcpb CLI not on PATH")

        # 4. uvx launch + MCP initialize handshake
        uvx = find_uvx()
        if not uvx:
            print("[SKIP] uvx handshake - uvx not found (PATH or venv Scripts)")
        else:
            print(f"\nLaunching: {uvx} {expected_arg}  (first run may download from PyPI)")
            init = {
                "jsonrpc": "2.0", "id": 1, "method": "initialize",
                "params": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {},
                    "clientInfo": {"name": "smoke-test", "version": "0"},
                },
            }
            proc = subprocess.Popen(
                [uvx, expected_arg],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                stderr=subprocess.PIPE, text=True, encoding="utf-8",
            )
            try:
                out, err = proc.communicate(json.dumps(init) + "\n", timeout=120)
            except subprocess.TimeoutExpired:
                proc.kill()
                out, err = proc.communicate()
            resp = None
            for line in out.splitlines():
                line = line.strip()
                if line.startswith("{"):
                    try:
                        cand = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    if cand.get("id") == 1 and "result" in cand:
                        resp = cand
                        break
            ok = resp is not None
            info = ""
            if ok:
                si = resp["result"].get("serverInfo", {})
                info = f"serverInfo={si.get('name')} {si.get('version')}"
            else:
                info = "no initialize response; stderr tail: " + " | ".join(err.strip().splitlines()[-3:])
            check("uvx stdio initialize handshake", ok, info)

    print()
    if FAILURES:
        print(f"RESULT: {len(FAILURES)} failure(s): {FAILURES}")
        return 1
    print("RESULT: all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
