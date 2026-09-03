"""Phase 7 public-copy guards. Stale marketing numbers shipped twice in
launch week (llms.txt claiming the wrong version, the site claiming a
retired tool count). These tests make the failure mode mechanical: the
published operation/tool/token figures must match what the measurement
scripts produce, no retired figure may survive in any public file, and no
public file may carry an em dash in any language.

Public files guarded: README.md, docs/index.html, docs/llms.txt.
Em-dash sweep also covers docs/MIGRATION_V2.md.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PUBLIC = [
    ROOT / "README.md",
    ROOT / "docs" / "index.html",
    ROOT / "docs" / "llms.txt",
]
EM_DASH_FILES = PUBLIC + [ROOT / "docs" / "MIGRATION_V2.md"]


def _run(script: str) -> str:
    out = subprocess.run(
        [sys.executable, "-X", "utf8", str(ROOT / "scripts" / script)],
        capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
    )
    assert out.returncode == 0, f"{script} failed: {out.stderr}"
    return out.stdout


def _measured():
    """Authoritative figures, straight from the scripts."""
    surf = _run("measure_surface.py")
    ops = _run("count_operations.py")
    lite = int(re.search(r"lite startup surface: (\d+) tools", surf).group(1))
    full = int(re.search(r"full surface .*?: (\d+) tools", surf).group(1))
    n_ops = int(re.search(r"TOTAL DISTINCT OPERATIONS:\s+(\d+)", ops).group(1))
    return lite, full, n_ops


def test_published_numbers_match_scripts():
    """Every headline number in the public copy is the script's number."""
    lite, full, n_ops = _measured()
    assert (lite, full, n_ops) == (28, 110, 187), (
        f"scripts now report lite={lite} full={full} ops={n_ops}; update the "
        "public copy AND this test together (that is the whole point)."
    )
    docs = full - 2  # document tools, excluding enable/disable_tools
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    llms = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    index = (ROOT / "docs" / "index.html").read_text(encoding="utf-8")
    for name, text in (("README", readme), ("llms.txt", llms),
                       ("index.html", index)):
        assert str(n_ops) in text, f"{name} is missing the operations count"
        assert str(docs) in text, f"{name} is missing the tool count"
        assert ("7.5k" in text or "7,500" in text), (
            f"{name} is missing the lite token figure"
        )


# Retired figures that must never reappear as current claims. Each entry is
# a compiled pattern; a match in any public file fails the guard.
_STALE = [
    (r"\b107\b", "retired tool count 107"),
    (r"\b182\b", "retired operations count 182 (undercounted delegate "
                 "dispatch; corrected to 187 in the Fable re-audit)"),
    (r"27[,.   ]?000", "retired token bill 27,000"),
    (r"~?27k\b", "retired token bill 27k"),
    (r"2\.5\s*[-–to ]{1,4}3\s*x", "retired get_text 2.5-3x claim"),
    (r"\b2\.5x\b", "retired 2.5x claim"),
    (r"\b946\b", "retired test count 946"),
    (r"\b910\b", "retired test count 910"),
    (r"\b852\b", "retired test count 852"),
    (r"\b4,400\b", "retired stress-call count 4,400"),
    (r"\b4\.400\b", "retired stress-call count 4.400"),
    (r"144\s*token", "retired per-tool average 144 tokens"),
    (r"\b20 tools\b", "retired lite estimate 20 tools"),
    (r"\b4,000\b", "retired lite token estimate 4,000"),
    (r"\b4\.000\b", "retired lite token estimate 4.000"),
]


def test_no_stale_figures_in_public_copy():
    for path in PUBLIC:
        text = path.read_text(encoding="utf-8")
        for pat, why in _STALE:
            m = re.search(pat, text)
            assert m is None, (
                f"{path.name}: {why} still present near "
                f"{text[max(0, m.start()-25):m.start()+25]!r}"
            )


def test_189_only_as_v1_history():
    """189 was the v1.6 tool count. It may appear only where a line names
    v1.6/v1.x/migration context, never as a current claim."""
    for path in PUBLIC:
        for i, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if "189" not in line:
                continue
            low = line.lower()
            assert ("v1.6" in low or "v1.x" in low or "migrat" in low), (
                f"{path.name}:{i}: bare 189 not qualified as v1 history: "
                f"{line.strip()[:90]!r}"
            )


def test_readme_pack_table_matches_measurement():
    """Every row of the README pack table carries the tool count and token
    figure that measure_surface.py currently reports, so a re-measured
    surface cannot ship under stale pack numbers."""
    surf = _run("measure_surface.py")
    rows = re.findall(r"^([a-z][\w-]*)\s+(\d+)\s+(~?[\d.]+k)\s*$", surf, re.M)
    assert len(rows) >= 8, f"measure_surface pack table not parsed: {rows!r}"
    lines = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()
    for pack, tools, toks in rows:
        cands = [ln for ln in lines
                 if ln.startswith(f"| {pack} ") or ln.startswith(f"| **{pack}**")]
        assert cands, f"README pack table has no row for {pack!r}"
        ln = cands[0]
        assert f" {tools} " in ln and toks.lstrip("~") in ln, (
            f"README row for {pack!r} does not carry measured "
            f"{tools} tools / {toks}: {ln!r}"
        )


def test_token_headlines_match_measurement():
    """The lite/full token figures in every public file are the measured
    ones, in some accepted thousands format."""
    surf = _run("measure_surface.py")
    lite_k = re.search(r"lite startup surface: \d+ tools, ~([\d.]+)k",
                       surf).group(1)
    full_k = re.search(r"full surface .*?: \d+ tools, ~([\d.]+)k",
                       surf).group(1)

    def variants(k: str) -> list[str]:
        n = int(round(float(k) * 1000))
        comma = f"{n:,}"
        seps = [".", " ", " ", " "]  # locale thousands separators
        return [f"{k}k", comma] + [comma.replace(",", s) for s in seps]

    for path in PUBLIC:
        text = path.read_text(encoding="utf-8")
        for label, k in (("lite", lite_k), ("full", full_k)):
            assert any(v in text for v in variants(k)), (
                f"{path.name}: measured {label} figure ~{k}k appears in no "
                f"accepted format {variants(k)}"
            )


def test_i18n_dictionaries_carry_current_figures():
    """Each of the seven language dictionaries in index.html must carry the
    operations count and the measured token figures in a locale-accepted
    format; the whole-file substring checks cannot see a single stale
    language."""
    text = (ROOT / "docs" / "index.html").read_text(encoding="utf-8")
    langs = ["en", "ko", "ja", "zh", "de", "fr", "es"]
    spans = []
    for lang in langs:
        start = text.index(f"\n{lang}: {{")
        spans.append(start)
    spans.append(text.index("/*I18N-END*/"))
    def forms_of(base: str) -> list[str]:
        return [base] + [base.replace(",", s)
                         for s in (".", " ", " ", " ")]

    sep = forms_of("7,500")
    full = forms_of("26,400")
    old = forms_of("34,400")
    for i, lang in enumerate(langs):
        block = text[spans[i]:spans[i + 1]]
        assert "187" in block, f"i18n {lang}: operations count 187 missing"
        for name, forms in (("lite 7.5k", sep), ("full 26.4k", full),
                            ("v1.6 34.4k", old)):
            assert any(f in block for f in forms), (
                f"i18n {lang}: {name} figure missing in all accepted formats"
            )


def test_no_em_dashes_in_any_public_file():
    for path in EM_DASH_FILES:
        text = path.read_text(encoding="utf-8")
        assert "—" not in text, f"{path.name} contains an em dash"


def test_mcp_name_marker_survives():
    first = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()[0]
    assert first == (
        "<!-- mcp-name: io.github.nometalalchemist/kitchensink4word -->"
    ), "README line 1 mcp-name marker was lost in the rewrite"


def test_non_affiliation_disclaimer_present():
    for path in (ROOT / "README.md", ROOT / "docs" / "index.html",
                 ROOT / "docs" / "llms.txt"):
        text = path.read_text(encoding="utf-8")
        assert "Not affiliated with or endorsed by Microsoft" in text, (
            f"{path.name} is missing the non-affiliation disclaimer"
        )
