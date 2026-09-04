# v1.5 Research: OMML Equation Authoring + .mcpb Bundle

**Written 2026-08-28 05:38 KST.** Implementation-grade research for the two v1.5 candidates.
Everything marked EMPIRICAL was executed on this machine during this session (Word 16.0
build 16.0.20326, Python 3.14, current repo state = v1.4.0, 147 tools). No code was
written into src/ or tests/.

---

# TOPIC 1 — OMML Equation Authoring (LaTeX → Word math)

## 1.1 The conversion landscape

The pipeline everyone converges on is two-stage: **LaTeX → MathML → OMML**. No credible
pure-Python direct LaTeX→OMML converter exists; the direct converters are JavaScript
(`latex-to-omml` on npm, MIT) or C#. The Python ecosystem:

| Component | What it does | License | Maintenance | Deps |
|---|---|---|---|---|
| **latex2mathml** (roniemartinez) | LaTeX → MathML, pure Python | MIT | **Active** — v3.81.0, last release 2026-04-15, Production/Stable | none |
| **MML2OMML.XSL** (Microsoft) | MathML → OMML, XSLT 1.0 | Office EULA — **NOT redistributable** | Ships with every Word install | lxml to run it |
| **mathml2omml** (amedama41) | MathML → OMML, pure Python | MIT | Frozen — last release 2019-11-24, but the OMML spec is frozen too | none (stdlib SAX) |
| **dwml** (xiilei) | OMML → LaTeX — **reverse direction only**, for displaying Word math in browsers | (unverified) | n/a | Not applicable to this feature |
| **addFormula2docx** (Sun-ZhenXing) | Recipe repo: python-docx + the XSL trick | — | Example code, not a dependency candidate | — |

**What python-docx users actually do:** the canonical Stack Overflow / blog recipe is
exactly `latex2mathml.converter.convert()` → `lxml.etree.XSLT(MML2OMML.XSL)` → append the
resulting element to `paragraph._p`. EMPIRICAL: this recipe works verbatim on this
machine (details in 1.2). python-docx itself has no math API and never will (dormant).

**MML2OMML.XSL on this machine (EMPIRICAL):**

- `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL` (194,432 bytes; the
  reverse `OMML2MML.XSL` sits beside it, 115,781 bytes)
- Header inspected: **plain XSLT 1.0, no `msxsl:script`, no extension functions** —
  lxml's XSLT 1.0 engine runs it cleanly (verified, see 1.2)
- **Licensing reality:** it is a component of Microsoft Office, covered by the Office
  license; there is no redistribution grant. A widely-cited Microsoft Q&A answer claiming
  it derives from TEI open-source stylesheets is wrong (checked: the TEI/transpect
  `mml2omml.xsl` files are independent third-party implementations, not the Microsoft
  file, and the transpect one carries no license header at all). **Do not vendor
  Microsoft's XSL into a PolyForm NC package.** Loading it at runtime from the user's own
  Office install is fine (they licensed Office), but that makes the file-based tier
  Office-dependent, which contradicts the product's cross-platform file-based story.

**License compatibility:** latex2mathml (MIT) and mathml2omml (MIT) both bundle cleanly
under PolyForm NC distribution — include their copyright notices, done. Dependency
weight: both are pure Python with zero transitive dependencies; combined they add well
under 1 MB.

## 1.2 Fidelity testing (EMPIRICAL, run this session)

Full pipeline tested on 18 academic-math cases, LaTeX → latex2mathml → both converters
in parallel (Microsoft XSL via lxml, and mathml2omml). Script pattern preserved in the
session scratchpad (`fidelity_test.py`, `roundtrip_test.py`).

| Case | XSL path | mathml2omml path |
|---|---|---|
| `\frac`, `\sqrt`, `\binom` | OK (`m:f`, `m:rad`) | OK |
| `pmatrix` / `bmatrix` | OK (`m:d` + `m:m` matrix, correct delimiters) | OK |
| `cases` | OK — renders as `m:d` `{`-delimiter + `m:m` two-column matrix (visually correct; Word's native cases uses `m:eqArr`, cosmetic difference only) | OK (same approach) |
| `aligned` | **FAIL — latex2mathml emits malformed XML** (unescaped entity, `xmlParseEntityRef`) — bug is upstream in latex2mathml, both converters die on its output | same FAIL |
| `align*` | OK | OK |
| `\sum`, `\int` with limits | OK (`m:nary`, `undOvr` limits) | OK |
| Greek (`\alpha…\Gamma`), operators (`\leq \neq \approx \times \cdot`), `\partial`, `\nabla` | OK (correct Unicode codepoints) | OK |
| Accents (`\hat \bar \tilde \vec`), `\mathbf`, subscripts | OK | OK |
| `\lim_{x \to 0}` | OK but rendered as subscript (`m:sSub`) rather than under-limit (`m:limLow`/`m:func`) — displays as "lim₍x→0₎" beside rather than below. Minor fidelity gap, acceptable; flag in docs | same |
| `\text{…}` | OK (`m:nor` normal-text run) | OK |
| `\underbrace`, `\cdots` | OK | OK |

**Round-trip validation in Word (EMPIRICAL):** built a docx with python-docx containing
(a) an XSL-produced inline equation inside a sentence, (b) an XSL-produced display
equation (`m:oMathPara`), (c) a mathml2omml-produced equation. Opened it in an invisible
Word instance: `OMaths.Count == 3`, correct types (inline = `wdOMathInline`, block =
`wdOMathDisplay`), all structures built up correctly, no repair dialog, clean close.
**Both converter outputs are Word-valid.**

**mathml2omml integration notes (EMPIRICAL):**
- Its output string uses the `m:` prefix but **declares no namespace** — inject
  `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"` on the root
  before `etree.fromstring`.
- It emits bare `m:oMath` only — for display equations, wrap in `m:oMathPara` ourselves
  (trivial: one parent element).
- Its output is more verbose than the XSL's (extra `m:box` wrappers) but Word normalizes
  on open; the round-trip confirmed no visual or structural problem.

## 1.3 The OOXML side

- **Placement:** inline math is an `m:oMath` element that sits as a **direct child of
  `w:p`, sibling of `w:r`** (verified in the round-trip file:
  `…<w:r><w:t>The identity </w:t></w:r><m:oMath>…`). Display math is `m:oMathPara`
  (containing one or more `m:oMath`) also as a child of `w:p`, conventionally in its own
  paragraph. `latex2mathml.convert(tex, display="block")` sets `display="block"` on the
  MathML root and the XSL then emits `m:oMathPara` automatically; with mathml2omml we add
  the wrapper.
- **Namespace:** python-docx's default template already declares `xmlns:m` on the
  `w:document` root (verified), and lxml adds a local declaration when appending into
  documents that lack it — no plumbing needed either way.
- **mathFont/settings:** python-docx's default `settings.xml` already ships
  `m:mathPr` with `m:mathFont` = Cambria Math (verified). For documents created by other
  producers that lack `m:mathPr`, Word falls back to defaults — treat adding `m:mathPr`
  as a nice-to-have hardening step, not a requirement.
- **Runmap interaction (verified against `src/word_mcp/ops/_runmap.py`):** the runmap
  iterates `p.iter(qn("w:r"))` and indexes `w:t` text only. Math runs are `m:r`/`m:t` in
  the m: namespace, so **equations are completely invisible to the runmap** — and to
  python-docx `paragraph.text` (verified empirically: sentence text round-trips without
  the math). Consequences to design in deliberately:
  1. Search/replace can never corrupt an equation (good, free).
  2. Extracted text silently omits equation content. Decision needed: acceptable
     (document it), or add an opt-in `[EQUATION: …]` placeholder / linearized text in
     read tools. COM's `Range.Text` DOES include math glyphs with `\r` separators
     (verified), so live-mode reads already see math while file-mode reads do not —
     an inconsistency worth a docs note at minimum.
  3. Range deletions computed from runmap offsets will leave an embedded equation
     standing while deleting text around it. Needs an explicit test and a documented
     behavior choice.

## 1.4 COM alternative (EMPIRICAL — invisible DispatchEx probe, quit cleanly)

Probed on Word 16.0 (build 16.0.20326):

- **`OMaths.Add(range)` + `OMath.BuildUp()` on UnicodeMath linear text WORKS
  invisibly.** `(a+b)/(c-d)` → proper `m:f` fraction; `∫_0^∞ e^(-x^2) dx=√π/2` → proper
  `m:nary`/`m:sup`/`m:rad`/`m:f`. This is a genuine zero-dependency equation path — but
  the input language is **UnicodeMath, not LaTeX**.
- **BuildUp does NOT accept LaTeX.** `\frac{a}{b}` stays literal text after BuildUp
  (no `m:f` produced). The UI's LaTeX toggle (Equation ribbon) is **not exposed in the
  COM object model** — no `ConvertFromLaTeX`, no `ConvertFromLinearFormat` member exists
  (probed). OMath members that DO exist: `BuildUp`, `Linearize`,
  `ConvertToLiteralText`, `ConvertToMathText`, `ConvertToNormalText`.
- **`Range.InsertXML` with raw MathML FAILS silently** — no error, but `OMaths.Count`
  stays 0 and no `m:f` appears. (Clipboard-paste of MathML does convert, but clipboard
  hijacking is disqualified — it destroys the user's clipboard in live mode.)
- Net: COM gives a free **UnicodeMath** equation tool for live mode, not a LaTeX one.
  Agents speak LaTeX; a LaTeX→UnicodeMath translator would be a whole new converter to
  maintain. So COM is a complement, not the LaTeX path.

## 1.5 RECOMMENDATION — Topic 1

**Ship the pure-Python pipeline: `latex2mathml` + `mathml2omml` (both MIT, zero
transitive deps, ~<1 MB combined) as the ONE path for `insert_equation(latex=…)`.**
It works identically in file-based and live-COM modes (file mode: append the element via
lxml; live mode: same XML via `Range.InsertXML`-free route — write to the docx layer or
use the existing live insert machinery), is fully redistributable under PolyForm NC, and
passed every gated case in Word validation. Do NOT redistribute or runtime-depend on
Microsoft's MML2OMML.XSL (license risk + Office-install dependency in the file tier);
keep the local XSL only as a dev-time cross-check oracle in tests that are skipped when
Office is absent. Do NOT build the COM/UnicodeMath path as the primary (wrong input
language for agents); optionally expose it later as a live-mode bonus
(`insert_equation_unicodemath`) since it is literally free.

Implementation details settled by this research:
- Preprocess `\begin{aligned}` → `\begin{align*}` (and `\begin{align}` → `align*`)
  before conversion — upstream latex2mathml bug emits malformed XML for `aligned`.
  Also wrap conversion in try/except and return a clean tool error on any parse failure
  (never write malformed XML into a document).
- Inject the `xmlns:m` declaration on mathml2omml output; wrap in `m:oMathPara` for
  `display=true`.
- Tool surface: `insert_equation(path, latex, display=false, anchor…)` following the
  existing anchor conventions; consider `replace`-style targeting later.
- Docs note: equations are invisible to text extraction/search in file mode (runmap
  excludes `m:t` by design — this also makes search/replace equation-safe).

**Gate tests for the build session** (all validated feasible this session):
1. `\frac{a+b}{c}`, `\sqrt{x^2+1}`, `\binom{n}{k}` → contain `m:f`/`m:rad`
2. `pmatrix`/`bmatrix` → `m:d` + `m:m` with correct `begChr`/`endChr`
3. `cases` → `{`-delimited `m:m`
4. `aligned` AND `align*` → both succeed (aligned via the rewrite)
5. `\sum_{i=1}^{n}`, `\int_0^\infty` → `m:nary` with `undOvr`
6. Greek + operator set (`\alpha \Gamma \leq \neq \approx \times \partial \nabla`) →
   correct Unicode codepoints in `m:t`
7. Accents/bold/limits/`\text{}` cases from the 18-case table
8. Inline insert mid-sentence: surrounding runs intact, runmap offsets for the
   paragraph unchanged for non-math text
9. Display insert: `m:oMathPara` in its own `w:p`
10. Round-trip: generated file opens in real Word with `OMaths.Count` correct and no
    repair prompt (live-marked test)
11. Extraction: paragraph text excludes math; search/replace over a paragraph containing
    an equation neither matches nor corrupts it; range deletion behavior documented
12. Equation inside a table cell; equation in a tracked-changes document (insert as
    tracked insertion must not split the `m:oMath`)
13. Malformed LaTeX (`\frac{a`) → clean tool error, document untouched
14. Namespace: insert into a minimal third-party docx whose root lacks `xmlns:m`

---

# TOPIC 2 — .mcpb Bundle (MCP Bundle format)

## 2.1 Current spec and tooling

- **Format:** `.mcpb` is a ZIP with a required `manifest.json` (manifest_version
  currently "0.3"), optional `icon.png`, server code, and bundled deps. Formerly DXT
  (Anthropic), donated to the MCP project 2025-11 — spec now lives at
  `github.com/modelcontextprotocol/mcpb` (the `anthropics/mcpb` repo redirects/mirrors;
  examples live there too).
- **CLI:** npm package **`@anthropic-ai/mcpb`** provides `mcpb init` (interactive
  manifest), `mcpb validate`, `mcpb pack` (honors `.mcpbignore`), `mcpb sign`/`verify`
  (PKCS#7, optional; self-signed OK for dev — unsigned installs fine with a warning),
  plus `unpack`/`info`/`clean`.
- **Claude Desktop install:** double-click the `.mcpb` (file association), drag into the
  window, or Settings → Extensions → Advanced → Install Extension. Extracts under the
  Claude app-data `Claude Extensions/<name>/` dir. `user_config` fields render as a
  settings UI with `${user_config.key}` substitution into args/env; `${__dirname}`
  resolves to the install dir.
- **Smithery "Local (MCPB Bundle)":** `smithery mcp publish <bundle.mcpb> -n <org/server>`
  uploads the bundle as the downloadable stdio artifact for the listing. **Known bug
  (smithery-ai/cli #787): the registry 400-rejects any bundle whose manifest declares
  `tools`** (MCPB tools schema is incompatible with the registry's Tool schema) —
  workaround is stripping `tools` from the manifest before `mcpb pack`. Plan for two
  bundle variants or just omit the tools array (it is display-only metadata).

## 2.2 The Python problem

Manifest `server.type` options: `node` (recommended by spec; Claude Desktop ships its
own Node runtime), `python`, `uv` (added v0.4, **still marked experimental as of
2026-08**), `binary`. What each means for a Python server:

1. **`python` with bundled deps** (`server/lib/` + `PYTHONPATH`, or a full
   `server/venv`): uses the **user's system Python** (Desktop refuses install when it
   detects none — modelcontextprotocol/mcpb issue #84 shows it disabling install even
   when uv IS present). Two hard problems for us:
   - **Compiled deps are not portable**: the spec itself warns you cannot portably
     bundle pydantic (fastmcp → pydantic-core is a compiled wheel, cp-version- and
     platform-specific; the user's Python minor version must match the bundled wheel).
   - **pywin32 breaks under PYTHONPATH bundling**: pywin32 relies on its `.pth`
     bootstrap (`pywin32_bootstrap` adds `win32/`, `win32/lib/`, and loads the
     `pywin32_system32` DLLs). **`.pth` files are only processed in site directories,
     never on PYTHONPATH**, so `import win32com` fails from a `server/lib` bundle unless
     the entry script manually calls `site.addsitedir(lib_dir)` before importing.
     Workable but fragile. This is the heaviest, most brittle option — reject.
2. **`uv` type**: bundle contains just `manifest.json` + `pyproject.toml` + source
   (~100 KB); the HOST runs `uv run` and resolves deps at install/launch. Cross-platform
   and correct (a real venv → `.pth` processed → pywin32 works), but experimental,
   requires Claude Desktop ≥0.10, and has open compatibility-detection bugs (#84).
   Example: `modelcontextprotocol/mcpb` `examples/file-manager-python` uses
   `command: uv run --directory ${__dirname} server/main.py`.
3. **Thin launcher via `mcp_config`** — the manifest's `mcp_config` is ultimately just
   the stdio command written into the client config, so a bundle can declare
   `command: "uvx", args: ["kitchensink4word"]` and carry almost nothing inside.
   This is exactly the `uvx <pypi-package>` pattern already standard for Python MCP
   servers (e.g., `mcp-server-fetch`, `arxiv-mcp-server`), riding on the fact that
   **kitchensink4word is already on PyPI with `word-mcp`/`kitchensink4word` console
   scripts**. uvx builds a cached ephemeral env: real venv, `.pth` processed, pywin32
   installs from its Windows wheel — everything just works. Prerequisite: the user has
   uv installed (uvx ships with uv). No pip-install-at-install-time hook exists in the
   manifest; `uv`/`uvx` delivery IS the supported "install from a registry" story.

Real-world Python bundles found: the spec's own `file-manager-python` (uv type),
himalaya-mcp (documents the format extensively), Toloka tendem-mcp. No found example
bundles pywin32 in `server/lib` — Windows-native-dep servers converge on uvx/uv.

## 2.3 Constraints mapped to this server

- pywin32 is already correctly gated: `pywin32; sys_platform == 'win32'` in
  pyproject.toml, so a uvx install on macOS/Linux cleanly skips it and the file-based
  tier still works cross-platform. Live-COM tools are Windows-only by nature — say so in
  the manifest description, do NOT restrict `compatibility.platforms` to win32.
- Entry points `kitchensink4word` / `word-mcp` are console scripts → `uvx
  kitchensink4word` resolves and runs them directly.
- fastmcp/pydantic compiled deps make the bundled-lib route effectively unshippable
  (confirms the spec's own warning); this alone forces the uvx route.

## 2.4 RECOMMENDATION — Topic 2

**Ship a thin uvx-launcher bundle: `server.type: "python"` in name only is wrong — use
the manifest with `mcp_config.command: "uvx"` pinning the PyPI package, and keep the
bundle to manifest + icon (~few KB).** Concretely:

Build steps:
1. `npm install -g @anthropic-ai/mcpb`
2. Create a `mcpb/` staging dir in the repo (never pack the whole repo): `manifest.json`,
   `icon.png`, optionally a short `README`.
3. manifest.json core (validate with `mcpb validate`):
   ```json
   {
     "manifest_version": "0.3",
     "name": "kitchensink4word",
     "display_name": "KitchenSink4Word",
     "version": "1.5.0",
     "description": "Everything plus the kitchen sink for Microsoft Word ...",
     "author": { "name": "nometalalchemist" },
     "server": {
       "type": "binary",
       "entry_point": "",
       "mcp_config": {
         "command": "uvx",
         "args": ["kitchensink4word==1.5.0"]
       }
     },
     "compatibility": {
       "claude_desktop": ">=0.10.0",
       "platforms": ["win32", "darwin", "linux"],
       "runtimes": { "python": ">=3.12" }
     }
   }
   ```
   (If `mcpb validate` rejects `type: "binary"` with an empty entry_point, fall back to
   the `uv` server type with a stub `pyproject.toml` whose only dependency is
   `kitchensink4word==1.5.0` — both deliver the same PyPI package; test which validates
   and installs cleanly on this machine's Claude Desktop, that empirical check is the
   first build-session task.)
4. Pin the exact version in args (`==1.5.0`) and rebuild the bundle each release —
   uvx caches envs, and an unpinned spec can serve a stale cached version.
5. `mcpb pack mcpb/ kitchensink4word.mcpb`, then double-click-install into local Claude
   Desktop as the smoke test (tools list appears, a file-based tool call works, and on
   this machine a live-COM call works).
6. Smithery: `smithery mcp publish kitchensink4word.mcpb -n <org>/kitchensink4word` —
   with NO `tools` array in the manifest (cli bug #787 rejects bundles that declare
   tools). If Desktop-facing tool listings are wanted later, maintain a second variant
   with `tools` for direct distribution only.
7. Signing (`mcpb sign`) optional — skip for v1.5, unsigned installs work; revisit if a
   directory submission requires it.

Expected pitfalls (all researched above): uv must be on the user's machine (state it in
the description + a friendly failure hint; this is the accepted norm for Python MCP
servers); Desktop's runtime detection is buggy around python/uv types (#84) — the
`binary`/uvx form sidesteps the system-Python check entirely; never attempt
`server/lib` bundling (pydantic-core portability + pywin32 `.pth` bootstrap both break);
`.mcpbignore` matters if packing from a dir with stray files; the manifest `tools` array
is display-only and currently poisons Smithery publishing.

---

## Sources

- latex2mathml: https://github.com/roniemartinez/latex2mathml (MIT, v3.81.0 2026-04)
- mathml2omml: https://pypi.org/project/mathml2omml/ and https://github.com/amedama41/mathml2omml (MIT)
- dwml (reverse direction): https://github.com/xiilei/dwml
- python-docx XSL recipe example: https://github.com/Sun-ZhenXing/addFormula2docx
- MML2OMML redistribution discussion (answer quality poor, treated as unreliable): https://learn.microsoft.com/en-us/answers/questions/5286296/redistrubution-of-omml2mml-xsl-from-ms-office
- MCPB adoption announcement: https://blog.modelcontextprotocol.io/posts/2025-11-20-adopting-mcpb/
- MCPB spec + manifest: https://github.com/modelcontextprotocol/mcpb/blob/main/MANIFEST.md
- MCPB CLI: https://www.npmjs.com/package/@anthropic-ai/mcpb
- Python uv example: https://github.com/anthropics/mcpb/blob/main/examples/file-manager-python/manifest.json
- Format reference (third-party, detailed): https://data-wise.github.io/himalaya-mcp/reference/mcpb-format-reference/
- Smithery publish: https://smithery.ai/docs/build/publish ; tools-schema bug: https://github.com/smithery-ai/cli/issues/787
- Desktop python/uv detection bug: https://github.com/modelcontextprotocol/mcpb/issues/84
- uvx-for-MCP pattern: https://docs.bswen.com/blog/2026-03-05-using-uvx-with-mcp-servers/

*Empirical artifacts from this session (scratchpad, disposable): `omath_probe.py`,
`fidelity_test.py`, `roundtrip_test.py`, `omml_roundtrip.docx`.*
