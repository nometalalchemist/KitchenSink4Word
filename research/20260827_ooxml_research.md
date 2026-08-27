# OOXML / Word MCP Server — Implementation Research

**Date:** 2026-08-27 (KST)
**Purpose:** Implementation reference for building a custom Microsoft Word MCP server in Python (FastMCP + python-docx + lxml direct OOXML manipulation). Each topic covers: XML structures, algorithms, repair-prompt pitfalls, and reference implementations with licenses.
**Method:** Four parallel research passes over GitHub source (fetched and read directly), python-docx docs/issues, MS-DOCX open specs, ECMA-376 schema references (datypic, c-rex, liquid-technologies mirrors — officeopenxml.com was unreachable via HTTPS during research; its content was corroborated via mirrors), plus empirical verification on this machine's Python 3.14.3 environment.

---

## Local environment verification (empirical, this machine, 2026-08-27)

Verified directly on `the local Python 3.14 interpreter` (Python 3.14.3, MSC v.1944 64-bit, Windows 11):

- **Installed and importing cleanly on 3.14.3:** fastmcp 2.14.7, mcp 1.27.2, pydantic 2.13.4 / pydantic_core 2.46.4, lxml 6.0.2 (cp314 win_amd64 wheel), python-docx 1.2.0 (dist-info), bayoo-docx 0.2.20, anyio 4.13.0, starlette 1.3.1, uvicorn 0.49.0. `import fastmcp, mcp, docx, lxml.etree, pydantic` all succeed (fastmcp emits a harmless authlib.jose deprecation warning).
- **CRITICAL PACKAGING CONFLICT:** `docx.__version__` reports **0.2.20** — bayoo-docx has **shadowed python-docx 1.2.0** in the shared site-packages. Both projects install the `docx` package; whichever installed last owns the files. The active `docx/` tree is bayoo's fork (old python-docx 0.8.x layout with bayoo's `oxml/footnotes.py` and `oxml/comments.py` additions). Consequence: **python-docx 1.2.0's native comments API is NOT importable in the shared environment.** The word-mcp server MUST use its own venv with python-docx 1.2.0 only (do not install bayoo-docx alongside it; port bayoo's footnote patterns as code instead — it is MIT, see Topic 9).
- The installed fastmcp (2.14.7) already runs on 3.14 here, but note PyPI's current line is 3.4.x/4.0-beta (see Topic 8) — a fresh venv will resolve to a newer fastmcp than the globally installed one.

---

# TOPIC 1 — Table column insert/delete with merged cells

## 1.1 Data model: tblGrid, gridSpan, vMerge

A WordprocessingML table has no column objects; it has a grid plus rows of cells:

- `w:tbl` → `w:tblPr`, `w:tblGrid` (one `w:gridCol` per **grid column**, optional `w:w` width in twips; schema: empty element, `w:w` is ST_TwipsMeasure — https://www.datypic.com/sc/ooxml/e-w_gridCol-1.html), then `w:tr` rows.
- A row's Nth `w:tc` does NOT correspond to grid column N. Mapping is **cumulative**: a cell's starting grid column = `grid_before + sum(grid_span of preceding w:tc siblings)`, where `grid_span` = `./w:tcPr/w:gridSpan/@w:val`, **defaulting to 1 when absent**.
- **Horizontal merge:** ONE `w:tc` carries `<w:gridSpan w:val="N"/>`; the spanned-over cells **do not exist as elements** (row has fewer `w:tc` than grid columns).
- **Vertical merge:** every row of the span keeps its `w:tc`. Top cell: `<w:vMerge w:val="restart"/>`. Continuation cells: `<w:vMerge/>` (omitted `w:val` = "continue" per ECMA-376 §17.4.84; absent element = not merged). Continuation cells **must carry the same gridSpan as the restart cell** (python-docx cell-merge analysis: https://github.com/python-openxml/python-docx/blob/master/docs/dev/analysis/features/table/cell-merge.rst) and conventionally contain one empty `w:p`.

python-docx (MIT, https://github.com/python-openxml/python-docx, `src/docx/oxml/table.py`) implements the mapping — reuse these as building blocks:

```python
@property
def grid_offset(self) -> int:
    """Starting offset of `tc` in the layout-grid columns of its table."""
    grid_before = self._tr.grid_before
    preceding_tc_grid_spans = sum(
        tc.grid_span for tc in self.xpath("./preceding-sibling::w:tc")
    )
    return grid_before + preceding_tc_grid_spans

@property
def grid_span(self) -> int:
    """Determined by ./w:tcPr/w:gridSpan/@val, it defaults to 1."""
    tcPr = self.tcPr
    return 1 if tcPr is None else tcPr.grid_span
```

`CT_Row.tc_at_grid_offset(grid_offset)` walks tcs decrementing a `remaining_offset` by each `tc.grid_span` and raises `ValueError` if no cell *starts* at that offset (offset falls mid-span).

Example — 3-column grid; row 1 has a 2-col horizontal merge plus a vMerge restart in col 3; row 2 continues the vertical merge:

```xml
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblLayout w:type="autofit"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="2880"/><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="5760" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>
      <w:p/>
    </w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>
      <w:p/>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p/></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p/></w:tc>
    <w:tc>
      <w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:vMerge/></w:tcPr>
      <w:p/>
    </w:tc>
  </w:tr>
</w:tbl>
```

## 1.2 Width and layout semantics

- `w:tcW` (`w:tcPr` child): `w:type` ∈ `dxa` (twips), `pct` (**fiftieths of a percent** of table width — 1667 = 33.3%; Word also accepts "33.3%"), `auto`, `nil`. Omitted = auto. (ECMA-376 §17.4.71; http://webapp.docx4java.org/OnlineDemo/ecma376/WordML/tcW.html)
- `w:tblLayout w:type="fixed|autofit"` (absent = autofit): `fixed` renders purely from tblGrid/tcW widths; `autofit` recomputes from content, so Word silently re-derives mildly inconsistent grids on open/save, while `fixed` renders errors visibly.
- Word rewrites tblGrid on save as the union of computed cell boundaries. The invariant to maintain when editing: for every row, `grid_before + Σ(grid_span) + grid_after == len(gridCol list)`, and ideally `Σ(gridCol/@w:w) ≈ tblW` under fixed layout.

## 1.3 Algorithm — DELETE grid column N (merge-aware)

No public API exists in python-docx (feature request open since 2017: https://github.com/python-openxml/python-docx/issues/441; maintainer scanny's design notes there identify the two hard parts — dangling references inside deleted cells, and merge interpretation). Correct algorithm, per row with a running grid cursor:

```python
from docx.oxml.ns import qn

def delete_grid_column(tbl, n):
    """Delete grid column index n (0-based) from a CT_Tbl, merge-aware."""
    tblGrid = tbl.find(qn('w:tblGrid'))
    gridCols = tblGrid.findall(qn('w:gridCol'))
    ncols = len(gridCols)
    assert 0 <= n < ncols

    for tr in tbl.findall(qn('w:tr')):
        trPr = tr.find(qn('w:trPr'))
        grid_before = _int_val(trPr, 'w:gridBefore', 0)
        grid_after  = _int_val(trPr, 'w:gridAfter', 0)

        # Case A: column n falls inside the gridBefore/gridAfter dead zones
        if n < grid_before:
            _decrement_val(trPr, 'w:gridBefore')   # and shrink w:wBefore if tracked
            continue
        row_span_total = sum(_grid_span(tc) for tc in tr.findall(qn('w:tc')))
        if n >= grid_before + row_span_total:      # inside gridAfter zone
            if grid_after > 0:
                _decrement_val(trPr, 'w:gridAfter')
            continue

        # Case B: find the tc covering grid column n
        cursor = grid_before
        for tc in tr.findall(qn('w:tc')):
            span = _grid_span(tc)
            if cursor <= n < cursor + span:
                if span > 1:
                    # Cell spans multiple grid cols: SHRINK, don't remove.
                    _set_grid_span(tc, span - 1)   # span-1 == 1 -> remove w:gridSpan entirely
                    _shrink_tcW(tc, gridCols[n])   # subtract deleted col width if type=dxa
                else:
                    # Cell occupies exactly this grid column: remove the w:tc.
                    # vMerge restart/continue cells removed the same way -- the
                    # whole vertical stack of width-1 cells goes, so no orphan
                    # continuation is possible.
                    tr.remove(tc)
                break
            cursor += span

    # Update the grid last
    tblGrid.remove(gridCols[n])
```

Key decisions:

1. **Shrink vs remove.** Covering cell with `grid_span > 1`: decrement `w:gridSpan/@w:val`; when new span is 1, **remove the `w:gridSpan` element entirely** (val="1" is legal but pointless). `grid_span == 1`: remove the whole `w:tc`, content included (optionally salvage content into a neighbor first).
2. **vMerge needs no special casing** if the same rule applies in every row: width-1 vMerge stacks lose their `w:tc` in every row (restart + continuations); multi-width stacks get gridSpan decremented in every row, keeping spans identical (the ECMA requirement). The danger is only *asymmetric* handling (remove restart's tc but shrink a continuation → continuation with no restart above, or mismatched spans in one merge).
3. **Widths:** for `tcW type="dxa"`, subtract the deleted `gridCol/@w:w`; for `pct`/`auto`, leave `tcW` alone. Remove the Nth `w:gridCol`; if `w:tblW` is dxa, subtract from it too.
4. **Plan grid-first, mutate-last:** compute every row's action against the ORIGINAL grid before touching anything (removing tcs changes sibling offsets).
5. **Dangling references:** removed cells may contain `w:hyperlink r:id`, `w:drawing` (image rels), comment range markers, bookmarks referenced by fields. Orphaned relationships are tolerated by Word; **half-deleted bookmarkStart/End or commentRange pairs are a repair-prompt vector** — scan the doomed content and remove the counterpart markers.

**Naive community recipe (merge-FREE tables only)** — python-docx issue #441 (Dreamsorcerer 2019, generalized by seahawks8):

```python
def delete_columns(table, columns):
    columns.sort(reverse=True)
    grid = table._tbl.find("w:tblGrid", table._tbl.nsmap)
    for ci in columns:
        for cell in table.column_cells(ci):
            cell._tc.getparent().remove(cell._tc)
        grid.remove(grid[ci])
```

**Why it corrupts merged tables:** python-docx's `column_cells`/`row.cells` *repeat the same `_tc` object* for every grid position a merged cell covers, so it deletes whole multi-column cells instead of shrinking them (and can double-remove a detached element); newer python-docx raises on irregular tables (issues #422, #992, #1434). **Never index through the `cells` convenience API when mutating** — use the grid-cursor algorithm.

## 1.4 Algorithm — INSERT a column at grid index N

python-docx only appends at right (`Table.add_column`: adds `gridCol`, then one `tc` per `tr` with width). For insert-at-N with merges, per row with the same grid cursor:

1. **N strictly inside a horizontal span** (covering cell starts before N and ends after): the new column passes *through* the merged cell → increment its `w:gridSpan` (create the element going 1→2), add the new width to `tcW` if dxa. Do NOT insert a `w:tc`.
2. **N at a cell boundary** (a cell starts exactly at N, or N == row width = append): insert a fresh `w:tc` via `existing_tc.addprevious(new_tc)` (or append). The new tc must be `<w:tc><w:tcPr><w:tcW w:w="..." w:type="dxa"/></w:tcPr><w:p/></w:tc>` — **the trailing `w:p` is mandatory** (python-docx `CT_Tc.new()` builds exactly `<w:tc><w:p/></w:tc>`; its `clear_content()` docstring warns the element is invalid without block content).
3. **Crossing a vertical merge:** apply the same decision to every row of the stack. Span cut through → ALL rows (restart + continuations) get gridSpan incremented identically. Boundary → each row gets an independent new tc (or give continuation rows `<w:vMerge/>` if the new column should join the vertical merge; independent unmerged cells are the simplest consistent choice).
4. **gridBefore/gridAfter zones:** increment `w:gridBefore/@w:val` (and `w:wBefore`) or `w:gridAfter` instead of adding a tc.
5. **Grid update:** insert `<w:gridCol w:w="..."/>` **at position N** (`gridCols[N].addprevious(...)`). gridCol matching is positional — appending the gridCol while inserting tcs mid-row silently misassigns every width at and after N.

**Cautionary reference — Rookie0x80/docx-mcp** (`src/docx_mcp/operations/tables/table_operations.py`; pyproject declares MIT but **no LICENSE file exists — see Topic 9, treat as all-rights-reserved; do not copy**): its column insert appends via `table.add_column()` then repositions cells with `row.cells[insert_at]` / `addprevious`. Two defects: it never moves the appended `w:gridCol` (positional width misassignment), and `row.cells[...]` returns duplicate tc references under merges (moves the wrong element). Works only for uniform merge-free tables. It has **no column delete at all**. Its `unmerge_cells` removes `gridSpan` and `vMerge` from `w:tcPr` — correct for purely vertical merges (continuation tcs still exist and revive), but **wrong for horizontal merges**: the spanned-over `w:tc`s were deleted at merge time, so a proper horizontal unmerge must insert `span-1` fresh `<w:tc><w:tcPr><w:tcW .../></w:tcPr><w:p/></w:tc>` cells after the unmerged cell and split the dxa width.

**Gold-standard merge code to mine — python-docx's own `CT_Tc.merge` chain** (MIT, `src/docx/oxml/table.py`): `_span_dimensions` (rejects non-rectangular regions) → `_grow_to` (recurses down rows setting vMerge None/"restart"/continue) → `_span_to_width` → `_swallow_next_tc`:

```python
def _swallow_next_tc(self, grid_width: int, top_tc: CT_Tc):
    def raise_on_invalid_swallow(next_tc):
        if next_tc is None:
            raise InvalidSpanError("not enough grid columns")
        if self.grid_span + next_tc.grid_span > grid_width:
            raise InvalidSpanError("span is not rectangular")
    next_tc = self._next_tc
    raise_on_invalid_swallow(next_tc)
    next_tc._move_content_to(top_tc)
    self._add_width_of(next_tc)          # tcW widths summed only when BOTH present
    self.grid_span += next_tc.grid_span
    next_tc._remove()
```

`_add_width_of` "does nothing if either tc does not have a specified width" — follow the same conservative width policy.

## 1.5 gridBefore / gridAfter

`w:trPr/w:gridBefore` and `w:gridAfter` (`w:val` = count; `w:wBefore`/`w:wAfter` = skipped width) declare unpopulated grid cells at a row's start/end (https://www.datypic.com/sc/ooxml/e-w_gridBefore-1.html). Rare in Word UI output but present in the wild. Every offset computation must add `grid_before`; insert/delete inside those zones adjusts counts, not the tc list. Forgetting them shifts every cell in affected rows — silent rendering corruption.

## 1.6 Table pitfalls → repair prompt / corruption

1. **`w:tc` with no block-level child** (must end with `w:p` or `w:tbl`+trailing `w:p`) → hard repair prompt. The single most common hand-built-tc corruption. Confirmed by MS-OI29500 §2.1.168: "if a tc element does not contain at least one p element as the last child block-level element, Word will fail to open the file."
2. **Row span sum ≠ grid count:** Word tolerates *short* rows (ragged right edge) but overrunning the grid, or missing `w:tblGrid` entirely, produces repair/mis-render (python-docx #547; pandoc nested-table bug https://github.com/jgm/pandoc/issues/6983).
3. **`gridSpan w:val="0"` or negative** → repair.
4. **Orphan vMerge continuation** (no restart above at the same grid columns): not schema-invalid but Word rewrites unpredictably; mismatched gridSpan within one vertical merge → staircase rendering, sometimes repair.
5. **Positional gridCol drift** (tc inserted without matching gridCol at the same index) → all widths shifted.
6. **Width sum mismatch under `tblLayout fixed`:** renders wrong (no repair); Word rewrites the grid on save.
7. **Dangling r:id / half-deleted bookmark or comment-range pairs** inside deleted cells → repair (rels merely orphaned are tolerated).
8. **Element order inside `w:tcPr`** is an xsd:sequence — `w:tcW` before `w:gridSpan` before `w:vMerge` etc. (full order in Topic 7). Raw lxml `append` of `w:gridSpan` after `w:shd` violates it.

---

# TOPIC 2 — Safe find-and-replace across fragmented runs

## 2.1 Why runs fragment

A paragraph's text splits into multiple `w:r` at every run-level boundary: (a) **rsid revision-save IDs** — Word stamps `w:rsidR`/`w:rsidRPr` per editing session, so identically formatted text typed in different sessions lands in different runs; (b) **`w:proofErr`** spell/grammar markers (`spellStart/spellEnd/gramStart/gramEnd` siblings force breaks around flagged words); (c) real formatting changes (`w:rPr` differences); (d) tracked changes (`w:ins`/`w:del` wrappers); (e) `w:hyperlink` wrapping its own runs; (f) fields, footnote refs, comment markers. So `${name}` can be stored as `${` + `name` + `}` across three runs. Naive per-run replace (i) misses matches spanning boundaries and (ii) `paragraph.text = ...` flattens to a single unformatted run, destroying all formatting.

**Cautionary real-world example — GongRzhe/Office-Word-MCP-Server** (MIT, `word_document_server/utils/document_utils.py::find_and_replace_text`, the code behind the currently-installed `search_and_replace` tool):

```python
for para in doc.paragraphs:
    if para.style and para.style.name.startswith("TOC"):
        continue
    if old_text in para.text:
        for run in para.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                count += 1
```

It detects matches on concatenated `para.text` but replaces only within single runs — **fragmented matches are detected and silently not replaced**; headers/footers/footnotes/comments/text boxes are missed. Do not copy this pattern.

## 2.2 The correct algorithm: character-offset map + run-preserving rewrite

**Step 1 — enumerate text atoms.** Walk run inner content in document order; build the concatenated string plus a per-character back-map `(run_element, w:t_element, char_index_within_t)`. Count as text: `w:t`, `w:tab` → `"\t"`, `w:br`/`w:cr` → `"\n"`, `w:noBreakHyphen` → `"\u2011"`, `w:ptab`. Model: python-docx `CT_R.text` getter:

```python
@property
def text(self) -> str:
    return "".join(
        str(e) for e in self.xpath("w:br | w:cr | w:noBreakHyphen | w:ptab | w:t | w:tab")
    )
```

Gather runs NOT from `paragraph.runs` (direct children only) but from `./w:r | ./w:hyperlink/w:r | ./w:ins/w:r | ./w:smartTag//w:r | ./w:fldSimple/w:r` (or `paragraph.iter_inner_content()` in modern python-docx). **Exclude** `w:del/w:r/w:delText` (tracked-deleted text) and runs containing `w:instrText` (field codes).

**Step 2 — find matches** with `str.find`/`re.finditer` on the concatenated string.

**Step 3 — rewrite affected runs.** For a match at `[s, e)`:

1. Map `s` and `e-1` to (first_run, first_t, i) and (last_run, last_t, j).
2. **First affected run:** keep text before the match, append the **entire replacement** — it inherits that run's `w:rPr` untouched (never rebuild rPr; just don't touch it).
3. **Last affected run:** keep only text after the match.
4. Delete or empty runs strictly between (empty via `run.text = ""` is safest; removal must not drop non-text siblings like `w:drawing`, and mixed-atom runs may only lose consumed atoms).
5. first == last → splice within the single `w:t`.
6. **Matches spanning a `w:hyperlink` boundary:** affected runs have different parents; safe policy — replacement goes into the first run (outside the link), in-link portion emptied, wrapper left in place — or refuse/flag such matches. Same for `w:ins` boundaries. Never move a run across the hyperlink boundary (changes what text is linked).
7. **Fields:** skip runs between `fldChar begin` and `end`; never touch `w:instrText` (editing ` HYPERLINK ` / ` REF ` codes corrupts field behavior). Replacing in result runs (after `separate`) is transient — blown away on update.
8. Iterate matches **right-to-left**, or re-derive the map after each splice.

**Step 4 — `xml:space="preserve"`.** Any written `w:t` with leading/trailing (or all-)whitespace MUST carry it or Word silently strips the spaces — exactly what splitting a match across runs produces. python-docx's own rule (`CT_R.add_t`):

```python
def add_t(self, text: str) -> CT_Text:
    t = self._add_t(text=text)
    if len(text.strip()) < len(text):
        t.set(qn("xml:space"), "preserve")
    return t
```

Going through python-docx's `run.text` setter handles this plus `\t`→`w:tab`, `\n`→`w:br` conversion (which also means a replacement containing `\n` becomes a line-break element, not literal text). Raw lxml writers must set the attribute themselves.

## 2.3 Reference implementations

**Best minimal-and-correct: ivanbicalho/python-docx-replace** (MIT, https://github.com/ivanbicalho/python-docx-replace). Driver tries the cheap path, falls back to the map; the map (`src/python_docx_replace/key_changer.py`, verbatim):

```python
class KeyChanger:
    def __init__(self, p, key, value) -> None:
        self.p = p
        self.key = key
        self.value = value
        self.run_text = ""
        self.runs_indexes: List = []
        self.run_char_indexes: List = []
        self.runs_to_change: Dict = {}

    def _initialize(self) -> None:
        run_index = 0
        for run in self.p.runs:
            self.run_text += run.text
            self.runs_indexes += [run_index for _ in run.text]
            self.run_char_indexes += [char_index for char_index, char in enumerate(run.text)]
            run_index += 1

    def replace(self) -> None:
        self._initialize()
        parsed_key_length = len(self.key)
        index_to_replace = self.run_text.find(self.key)

        for i in range(parsed_key_length):
            index = index_to_replace + i
            run_index = self.runs_indexes[index]
            run = self.p.runs[run_index]
            run_char_index = self.run_char_indexes[index]

            if not self.runs_to_change.get(run_index):
                self.runs_to_change[run_index] = [char for char_index, char in enumerate(run.text)]

            run_to_change: Dict = self.runs_to_change.get(run_index)
            if index == index_to_replace:
                run_to_change[run_char_index] = self.value   # first matched char -> whole replacement
            else:
                run_to_change[run_char_index] = ""           # other matched chars -> deleted

        for index, text in self.runs_to_change.items():
            run = self.p.runs[index]
            run.text = "".join(text)
```

The trick: per-character run ownership map; the first matched character *becomes* the whole replacement, every other matched character becomes `""` — surviving characters keep their original run/rPr automatically, and assignment through `run.text` handles xml:space and escaping. **Limitations to fix in a port:** uses `p.runs` (blind to `w:hyperlink`/`w:ins` content); first-occurrence-per-pass `while` loop (quadratic on many hits); doesn't sweep footnotes/endnotes/comments parts; leaves emptied `w:r` elements (harmless).

**docxtpl's different strategy** (elapouya/python-docx-template, **LGPL-2.1** — mind the license): XML-level regex preprocessing (`patch_xml`) that deletes markup Word inserted *between* known jinja2 delimiters (`{{`…`}}`), e.g. collapsing `{</w:t></w:r><w:r><w:t>{`. Works only because template tags are known delimiters; the right model for a placeholder-syntax tool, not general find/replace. Merged-region formatting collapses to the first run's.

**adejones gist** (https://gist.github.com/adejones/a6d42984f66ea9990d78974531863bee, **no license — reference only**): substring-based `(run_index, char_start, match_len)` bookkeeping; buggier partial-match scanning; prefer the KeyChanger design.

## 2.4 Find/replace pitfalls checklist

1. `xml:space="preserve"` on every whitespace-edged `w:t` (silent space-eating otherwise).
2. Never replace inside `w:instrText` / `w:fldSimple/@w:instr`; skip fldChar begin→end territory.
3. Exclude `w:delText` (or replacements resurrect deleted content); matches inside `w:ins` should stay inside the wrapper to keep revision attribution.
4. Sweep ALL story parts: `header*.xml`, `footer*.xml`, `footnotes.xml`, `endnotes.xml`, `comments.xml`, and text boxes (`//w:txbxContent//w:p`).
5. Never "reconstruct" runs copying selected properties — preserve whole `w:rPr` (or don't touch surviving runs at all). rStyle + direct formatting are both preserved automatically that way.
6. Mixed-atom runs (`w:t` + `w:tab`/`w:br`): python-docx `run.text` setter rebuilds correctly; raw lxml editing only the first `w:t` drops tabs/breaks.
7. Empty leftover runs and stale `w:proofErr` are harmless; rsids need no maintenance.
8. Never build XML by string substitution for replacement text (escaping); assign through lxml `.text`.
9. Process matches right-to-left or re-map after each splice.

---

# TOPIC 3 — Footnotes and endnotes at the OOXML level

## 3.1 Package plumbing (three registrations required)

**(a) `[Content_Types].xml` overrides:**

```xml
<Override PartName="/word/footnotes.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml" />
<Override PartName="/word/endnotes.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml" />
```

**(b) Relationships in `word/_rels/document.xml.rels`:**

```xml
<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
  Id="rId7" Target="footnotes.xml" />
```

Endnotes type: `.../relationships/endnotes`, `Target="endnotes.xml"`. Any unique Id. python-docx's `Part.relate_to(part, RT.FOOTNOTES)` handles both the rel and the content-type override on save — which is why bayoo-docx never touches `[Content_Types].xml` manually.

**(c) settings.xml special-note declarations** (3.3). Footnote references in the body with no related footnotes part = repair prompt.

## 3.2 word/footnotes.xml structure

Root `w:footnotes` containing flat `w:footnote` elements. Attributes: `w:id` (required, ST_DecimalNumber — signed; negatives legal and used for the separator), `w:type` (optional: `normal` default / `separator` / `continuationSeparator` / `continuationNotice`). Content: ≥1 block element (`w:p`, tables allowed). Verbatim shape (pandoc's Word-derived template, https://github.com/jgm/pandoc/blob/main/data/docx/word/footnotes.xml — pandoc is GPL-2.0-or-later, but these are spec-constant strings mirroring Word output):

```xml
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator /></w:r></w:p></w:footnote>
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator /></w:r></w:p></w:footnote>
  <w:footnote w:id="31">
    <w:p>
      <w:pPr><w:pStyle w:val="FootnoteText" /></w:pPr>
      <w:r>
        <w:rPr><w:rStyle w:val="FootnoteReference" /></w:rPr>
        <w:footnoteRef />
      </w:r>
      <w:r><w:t>Footnote Text.</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
```

`word/endnotes.xml` is structurally identical: `w:endnotes`/`w:endnote`, `w:endnoteRef`, `EndnoteText`/`EndnoteReference` styles, `w:endnoteReference` in the body.

## 3.3 The REQUIRED special notes — verified id convention

**Word-native convention (Word 2007→365, pandoc, bayoo-docx): separator = `w:id="-1"`, continuationSeparator = `w:id="0"`; real footnotes start at id 1.** The ECMA-376 Primer's 0/1/2/3 positive-id examples also work because **`w:type` is what determines the role, not the id** (proof: SecurityRonin/docx-mcp's endnote bootstrap assigns the types to ids 0/-1 *swapped* relative to Word and still validates). Emit the Word-native −1/0 pair; when reading third-party files, identify specials by `w:type`, never id; never renumber or delete them.

With the paragraph formatting Word emits (bayoo-docx `docx/templates/default-footnotes.xml`, MIT):

```xml
<w:footnote w:type="separator" w:id="-1">
  <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:separator/></w:r></w:p>
</w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0">
  <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:continuationSeparator/></w:r></w:p>
</w:footnote>
```

The empty `<w:separator/>` / `<w:continuationSeparator/>` run children render the horizontal rule. **A footnotes part with real notes but missing the special pair is a classic repair-prompt cause.**

**settings.xml, exact:** document-level `w:footnotePr` (CT_FtnDocProps) carries `w:footnote` children (0..3, one per non-normal type) referencing the specials by id:

```xml
<w:footnotePr>
  <w:footnote w:id="-1" />
  <w:footnote w:id="0" />
</w:footnotePr>
```

Endnotes analog: `w:endnotePr` with `<w:endnote w:id="-1"/><w:endnote w:id="0"/>`. Schema order inside footnotePr: `w:pos?`, `w:numFmt?`, `w:numStart?`, `w:numRestart?`, then the `w:footnote` refs (https://www.datypic.com/sc/ooxml/e-w_footnotePr-2.html). Only declare parts that exist — a settings id pointing to a nonexistent special is itself an inconsistency. Word repairs a missing settings declaration more gracefully than missing footnotes.xml entries, but emit both.

## 3.4 Body side, the two marks, and required styles

Body anchor — a run carrying the `FootnoteReference` character style and the empty `w:footnoteReference` element:

```xml
<w:r>
  <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
  <w:footnoteReference w:id="1"/>
</w:r>
```

Inside footnotes.xml, the note's first paragraph: `pStyle FootnoteText`; first run: `rStyle FootnoteReference` + `w:footnoteRef` (the self-referencing mark rendering the note's own number); Word also emits a plain run containing a single space after it.

**Superscript is NOT intrinsic to the reference elements — it comes entirely from the character style.** Missing styles = full-size marks (the #1 cosmetic bug), not corruption. Inject into styles.xml if absent (shape from pandoc styles.xml; load-bearing line is `vertAlign`):

```xml
<w:style w:type="paragraph" w:styleId="FootnoteText">
  <w:name w:val="Footnote Text"/><w:basedOn w:val="Normal"/>
  <w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
  <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
  <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
</w:style>
<w:style w:type="character" w:styleId="FootnoteReference">
  <w:name w:val="Footnote Reference"/><w:basedOn w:val="DefaultParagraphFont"/>
  <w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
  <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
</w:style>
```

(Plus `EndnoteText`/`EndnoteReference` with the same rPr. This is the "style-based superscript approach" attributed to bayoo-docx — bayoo only *assigns* the style names; the styles ship in its default-styles template for new documents. Adding footnotes to an arbitrary existing docx requires checking styles.xml and injecting.)

## 3.5 Numbering

- Real-note `w:id` values: unique, conventionally positive, monotonically increasing. **They do NOT determine displayed numbers** — Word assigns numbers dynamically by the order `w:footnoteReference` elements appear in the body. Inserting between existing notes requires no renumbering; there is no stored number anywhere.
- `w:footnotePr` numbering children: `w:numFmt` (`decimal|lowerRoman|chicago|...`), `w:numStart`, `w:numRestart` (`continuous|eachSect|eachPage`; endnotes: `continuous|eachSect`), `w:pos` (`pageBottom|beneathText`; endnotePr: `sectEnd|docEnd`). **Restart-per-section = `<w:numRestart w:val="eachSect"/>` inside each `w:sectPr`'s `w:footnotePr`** (section-level), not the settings-level one.
- Custom marks: `<w:footnoteReference w:customMarkFollows="true" w:id="4"/>` followed in the same run by literal text (`<w:t>*</w:t>`) as the visible mark; the note is skipped by auto-numbering; inside the note the `w:footnoteRef` position holds the same literal.

## 3.6 Add-footnote algorithm

1. Ensure `word/footnotes.xml` exists; if not, create with ONLY the two specials (−1/0), plus content-type override, document.xml.rels relationship, and settings.xml `w:footnotePr` refs.
2. Ensure `FootnoteText` + `FootnoteReference` styles exist; inject if missing.
3. `next_id = max(all w:footnote/@w:id) + 1` — **max, not last+1** (bayoo-docx's `_next_id` reads `ids[-1] + 1`, breaking on unsorted files; SecurityRonin's `max(existing | {0}) + 1` is the correct form).
4. Append `<w:footnote w:id="N">` with FootnoteText paragraph, ref run (rStyle + `w:footnoteRef`), a space run, then the text runs.
5. Insert the body reference run at the anchor point.

## 3.7 Delete-footnote algorithm

Both halves mandatory, atomically:
1. Remove `w:footnote[@w:id=N]` from footnotes.xml (refuse for id < 1 — protects the specials).
2. Remove every `w:footnoteReference[@w:id=N]`'s containing `w:r` from the body; if the run's parent is a now-empty `w:hyperlink`, remove the wrapper too.

**A body reference whose id has no definition = repair prompt.** The reverse (orphan definition, never referenced) does NOT trigger repair — just never displays. SecurityRonin's `validate_footnotes()` reports both directions (`missing_definitions` = corrupting; `orphan_definitions` = dead weight).

## 3.8 bayoo-docx implementation (github.com/BayooG/bayoo-docx — MIT, © 2019 Obay Daba, forked from python-docx)

Key patterns (quoted from master):

`docx/parts/footnotes.py` — lazy part creation from template:
```python
class FootnotesPart(XmlPart):
    @classmethod
    def default(cls, package):
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = parse_xml(cls._default_footnotes_xml())
        return cls(partname, content_type, element, package)
```

`docx/parts/document.py` — part registration via python-docx OPC machinery (rel + content type automatic):
```python
@property
def _footnotes_part(self):
    try:
        return self.part_related_by(RT.FOOTNOTES)
    except KeyError:
        footnotes_part = FootnotesPart.default(self)
        self.relate_to(footnotes_part, RT.FOOTNOTES)
        return footnotes_part
```

`docx/oxml/text/run.py` — the whole superscript mechanism is a style name:
```python
def add_footnote_reference(self, _id):
    rPr = self.get_or_add_rPr()
    rstyle = rPr.get_or_add_rStyle()
    rstyle.val = 'FootnoteReference'
    reference = OxmlElement('w:footnoteReference')
    reference._id = _id
    self.append(reference)
    return reference
```

Limitations to NOT copy: no delete support; no settings.xml footnotePr maintenance for pre-existing docs; no style injection into existing docs; `_next_id = ids[-1] + 1` ordering assumption.

## 3.9 SecurityRonin/docx-mcp implementation (github.com/SecurityRonin/docx-mcp — **MIT**, verified via GitHub API)

Files: `docx_mcp/document/footnotes.py` (full CRUD + validation), `docx_mcp/document/endnotes.py`. Verbatim highlights:

Next-id + definition build (note extras beyond the minimum: w14 paraId/textId and an `_FnN` bookmark so the body mark can be a real internal hyperlink, surviving PDF export):

```python
existing = {int(f.get(f"{W}id", "0")) for f in fn_tree.findall(f"{W}footnote")}
next_id = max(existing | {0}) + 1
anchor = f"_Fn{next_id}"
fn_el = etree.SubElement(fn_tree, f"{W}footnote")
fn_el.set(f"{W}id", str(next_id))
fn_para = etree.SubElement(fn_el, f"{W}p")
ppr = etree.SubElement(fn_para, f"{W}pPr")
ps = etree.SubElement(ppr, f"{W}pStyle"); ps.set(f"{W}val", "FootnoteText")
ref_run = etree.SubElement(fn_para, f"{W}r")
ref_rpr = etree.SubElement(ref_run, f"{W}rPr")
ref_style = etree.SubElement(ref_rpr, f"{W}rStyle"); ref_style.set(f"{W}val", "FootnoteReference")
etree.SubElement(ref_run, f"{W}footnoteRef")
sp_run = etree.SubElement(fn_para, f"{W}r")     # Word-style space after the mark
sp_t = etree.SubElement(sp_run, f"{W}t"); _preserve(sp_t, " ")
```

Delete (two-phase, hyperlink-wrapper aware):

```python
fn_tree.remove(target)                       # from footnotes.xml
for ref_el in doc.iter(f"{W}footnoteReference"):
    if ref_el.get(f"{W}id") == str(footnote_id):
        ref_run = ref_el.getparent(); container = ref_run.getparent()
        if container.tag == f"{W}hyperlink":
            container.getparent().remove(container)   # whole wrapper
        else:
            container.remove(ref_run)
```

Other reusable pieces: `_next_global_markup_id()` (scans bookmarkStart/End, ins/del, commentRange ids across ALL parts for a collision-free id); `add_footnote_ref()` re-cites an existing note elsewhere via bookmark hyperlink (a second real `w:footnoteReference` with the same id is also legal — shows the same number twice); `update_footnote()` rejects ids < 1; `validate_footnotes()` per 3.7. No renumber operation exists — none is needed (Word numbers by body order). Caveats: it never bootstraps `word/footnotes.xml`/content-type/rel itself and doesn't maintain settings.xml footnotePr — our server must cover both.

## 3.10 Footnote pitfalls → repair prompt / breakage

1. `w:footnoteReference` pointing to an undefined id (the delete-orphan case) → repair.
2. Real notes present but no separator/continuationSeparator entries → repair or lost rules; always emit.
3. footnotes.xml present but no relationship or no content-type override → part ignored / repair.
4. Duplicate `w:id` within footnotes.xml → repair.
5. `w:footnote` with no `w:p` child (content model requires ≥1 block).
6. `w:footnoteReference` inside footnotes.xml itself (no nesting) or in headers/footers/textboxes (disallowed).
7. Renumber passes touching ids −1/0, or renumbering definitions without renumbering body references.
8. Missing `FootnoteReference` style / forgotten rStyle — cosmetic (full-size marks), not corruption.

---

# TOPIC 4 — TOC insertion via field codes

## 4.1 Simple vs complex fields

**Simple:** `<w:fldSimple w:instr=" TOC \o &quot;1-3&quot; \h \z \u ">…cached result runs…</w:fldSimple>` — run-level, lives inside ONE paragraph, so the cached result cannot span paragraphs. A real TOC result is a paragraph sequence → **the complex form is effectively mandatory for TOC** (Word always writes TOCs as complex fields).

**Complex:** three run-level markers, legally spread across paragraphs:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<!-- cached result: zero or more runs / whole paragraphs -->
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

`w:fldChar` attributes: `fldCharType` (required), `dirty`, `fldLock` (https://schemas.liquid-technologies.com/OfficeOpenXML/2006/fldchar.html). `separate` is optional — without it the field has no cached result and displays empty until updated (no corruption). Instruction text may split across multiple `w:instrText` runs.

## 4.2 Switches

Word default set: ` TOC \o "1-3" \h \z \u `:
- `\o "1-3"` — build from Heading1..Heading3 paragraph styles
- `\h` — entries as hyperlinks
- `\z` — hide tab leader/page numbers in Web Layout
- `\u` — include paragraphs with applied outline levels (`w:outlineLvl`, covers custom styles)
- `\t "StyleName,1,Other,2"` — custom styles at given levels (combinable with `\o`)
- Others: `\a`, `\b` (bookmark-scoped), `\c` (caption tables), `\f`/`\l` (TC fields), `\n` (omit page numbers), `\p`, `\w`, `\x` (full table: https://github.com/dolanmiu/docx/blob/master/docs/usage/table-of-contents.md, MIT)

## 4.3 Word-native SDT wrapping — exact XML

Word wraps the TOC in a block-level SDT for the "Update Table…" UI chrome (structure is spec-defined; template shape corroborated at harvard-lil/h2o — that repo is **AGPL-3.0**, do not copy its code verbatim, but the XML shape below is standard Word output):

```xml
<w:sdt>
  <w:sdtPr>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>
      <w:r><w:t>Table of Contents</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    </w:p>
    <!-- optional cached entry paragraphs, each pStyle TOC1/TOC2/... -->
    <w:p>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
```

Notes: the `w:sdt` sits at BODY level (sibling of `w:p`, **before** the final `w:sectPr`); `w:sdtPr` must precede `w:sdtContent`; Word's real output adds `<w:id w:val="..."/>` and rsid noise (optional). `TOCHeading` style (based on Heading1) must exist or the title renders plain. **The SDT wrapper is entirely optional for correctness** — a bare field in ordinary paragraphs works and updates fine.

## 4.4 Forcing refresh on open

`word/settings.xml`:

```xml
<w:updateFields w:val="true"/>
```

Word recalculates ALL fields on next open — the only reliable no-Word way to get a populated TOC — but shows the modal **"This document contains fields that may refer to other files. Do you want to update the fields in this document?"** (confirmed: https://github.com/dolanmiu/docx/issues/1212). Word removes the element after updating (one-shot). Alternatives, none reliable without opening Word: COM (`doc.TablesOfContents(1).Update()` via win32com — viable on this Windows+Word machine), LibreOffice headless `updateall`, or pre-computing entries (page numbers impossible without a layout engine — omit via `\n` or wrap in PAGEREF fields Word fixes later).

**`dirty` flag alternative:** `<w:fldChar w:fldCharType="begin" w:dirty="true"/>` (or `w:fldSimple w:dirty`). ECMA: application "shall update this field when its contents are displayed." Per-field, usually no document-wide prompt; behavior less uniform across versions; LibreOffice ignores it. dolanmiu/docx defaults to this. **Recommendation: set `dirty="true"` on begin AND expose an option for `w:updateFields`** (docx4j pairs them the same way).

## 4.5 What a TOC looks like BEFORE Word populates it

- Nothing between `separate` and `end` (or no `separate`): renders as empty space — valid, but users think insertion failed. Put a placeholder run between separate and end: canonical text `"Right-click to update field."` (classic SO recipe).
- **"No table of contents entries found."** is Word's OWN update-result when switches match nothing (e.g. `\o "1-3"` with non-Heading styles, no outline levels, no `\u`/`\t`) — if you see it, the switches don't match the doc's heading strategy; you never write it.
- Cached-entry paragraphs, if pre-populated: `pStyle TOC1`/`TOC2`/… (+ `w:noProof` runs), optionally `w:hyperlink w:anchor="_Toc…"` + nested PAGEREF fields.
- **Bookmarks: a fresh, never-updated TOC field needs NO `_Toc` bookmarks — Word creates them around headings when IT updates the field. Do not fabricate `_Toc` bookmarks** (collision with Word's own on update). Only self-made pre-cached hyperlink entries need self-made anchors (use a distinct prefix like `_auto_toc_N`).

## 4.6 Reference implementation (python-docx recipe)

Classic recipe (Stack Overflow, CC BY-SA 4.0, https://stackoverflow.com/questions/18595864/ mirrored in python-docx issue #36 — python-docx has no TOC API):

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

paragraph = document.add_paragraph()
run = paragraph.add_run()
fldChar = OxmlElement('w:fldChar')
fldChar.set(qn('w:fldCharType'), 'begin')
instrText = OxmlElement('w:instrText')
instrText.set(qn('xml:space'), 'preserve')
instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
fldChar2 = OxmlElement('w:fldChar')
fldChar2.set(qn('w:fldCharType'), 'separate')
fldChar4 = OxmlElement('w:fldChar')
fldChar4.set(qn('w:fldCharType'), 'end')
r = run._r
r.append(fldChar); r.append(instrText); r.append(fldChar2)
# CLEAN placeholder form (the SO original nests w:t inside fldChar -- Word
# tolerates it but validators flag it; use a separate run instead):
r2 = paragraph.add_run("Right-click to update field.")._r
r3 = paragraph.add_run()._r
r3.append(fldChar4)
```

SDT wrapping: build `w:sdt`/`w:sdtPr`/`w:docPartObj`/`w:docPartGallery val="Table of Contents"`/`w:docPartUnique` via OxmlElement, then `document._element.body.insert_element_before(sdt, 'w:sectPr')`. settings.xml side:

```python
settings = document.settings.element
upd = OxmlElement('w:updateFields')
upd.set(qn('w:val'), 'true')
settings.append(upd)
```

GongRzhe/Office-Word-MCP-Server has **no TOC-insertion tool** (read-only `get_document_outline`) — nothing to borrow there.

## 4.7 TOC pitfalls

1. **`w:instrText` without `xml:space="preserve"`** → padding spaces eaten, instruction can malform. Word writes the instruction with surrounding spaces.
2. **fldChar begin/end imbalance = repair prompt.** Every `begin` needs one `end` in the same story. Paragraph-range deletions cutting through a field are the classic cause — any delete tool must scan the doomed range for `w:fldChar` and delete whole fields or refuse.
3. `separate` outside begin/end, or improperly closed nested fields (PAGEREF inside TOC results must each balance).
4. Wrong SDT placement: block-level sdtContent (containing `w:p`) inside a run-level sdt position = invalid → repair; `w:sdtPr` before `w:sdtContent`.
5. Body children appended after `w:sectPr` — the final sectPr must stay last (`insert_element_before(sdt, 'w:sectPr')`).
6. The SO recipe's `w:t` nested inside the separate `w:fldChar` — tolerated by Word, flagged by strict validators; use a separate run.
7. `\o` quoting: inside `w:fldSimple/@w:instr` the quotes are XML-escaped `&quot;`; inside `w:instrText` they are literal characters.
8. `w:updateFields` left set on every save = modal prompt on every open (reads like a virus warning). Use once, or prefer `dirty`.

---

# TOPIC 5 — Comments XML including modern threading (Word 2019+/365)

Note: officeopenxml.com was unreachable during research; details below verified against MS-DOCX specs (learn.microsoft.com), datypic's ECMA-376 browser, python-docx source, docx4j source, and two working implementations read directly.

## 5.1 Part inventory: exact content types and relationship types

Verified against docx4j (`ContentTypes.java`, `Namespaces.java` — Apache-2.0) and pablospe/docx-editor (MIT), which agree exactly:

| Part | Content type (Override) | Relationship type (from document.xml) | Root |
|---|---|---|---|
| `word/comments.xml` | `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` | `w:comments` |
| `word/commentsExtended.xml` | `...wordprocessingml.commentsExtended+xml` | `http://schemas.microsoft.com/office/2011/relationships/commentsExtended` | `w15:commentsEx` |
| `word/commentsIds.xml` | `...wordprocessingml.commentsIds+xml` | `http://schemas.microsoft.com/office/2016/09/relationships/commentsIds` | `w16cid:commentsIds` |
| `word/commentsExtensible.xml` | `...wordprocessingml.commentsExtensible+xml` | `http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible` | `w16cex:commentsExtensible` |
| `word/people.xml` | `...wordprocessingml.people+xml` | `http://schemas.microsoft.com/office/2011/relationships/people` | `w15:people` |

Namespaces: `w14` = `http://schemas.microsoft.com/office/word/2010/wordml` (paraId); `w15` = `.../word/2012/wordml`; `w16cid` = `.../word/2016/wordml/cid`; `w16cex` = `.../word/2018/wordml/cex`; `w16du` = `.../word/2023/wordml/word16du` (dateUtc). All relationships go in `word/_rels/document.xml.rels`. Extension-part roots must carry `mc:Ignorable` listing the extension prefixes (Word's own roots declare ~35 namespaces and `mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"`; docx-editor's templates reproduce this — minimal single-namespace roots also work, the full battery is what Word round-trips).

## 5.2 word/comments.xml

```xml
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Steve Canny" w:initials="SJC"
             w:date="2025-06-10T22:27:56Z">
    <w:p>
      <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
        <w:annotationRef/>
      </w:r>
      <w:r><w:t>I have this to say about that</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>
```

- `w:id` required, decimal; Word starts at 0. python-docx allocates `max(used_ids) + 1`.
- `w:author` required by schema (python-docx writes "" when absent); `w:initials` optional (Word always writes it); `w:date` optional, UTC `...Z`, seconds resolution.
- A comment is a full story: multiple `w:p`/`w:tbl`, hyperlinks, images allowed.
- First run conventionally holds `w:annotationRef` with `rStyle CommentReference` (cosmetic, "doesn't affect behavior"); python-docx also puts `pStyle CommentText` on comment paragraphs.

## 5.3 Anchoring in document.xml — exact placement

```xml
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>Hello, world!</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r>
    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
    <w:commentReference w:id="0"/>
  </w:r>
</w:p>
```

1. `commentRangeStart` immediately before the first run of the range; `commentRangeEnd` immediately after the last run. Ranges start/end at run boundaries — split runs as needed.
2. The reference run comes **after** `commentRangeEnd`: order is rangeStart … content … rangeEnd → reference run.
3. **`w:commentReference` is authoritative; range markers are optional** and only control highlight extent. No reference run anywhere = comment invisible.
4. All three share the `w:id` matching `w:comment/@w:id`.
5. Ranges may span paragraphs (EG_RangeMarkupElements members, block level allowed).
6. Comments cannot live in headers/footers or nest inside other comments — silently removed or repair.

python-docx 1.2.0's anchoring code (`src/docx/oxml/text/run.py`, MIT — note the two `addnext` in reverse order producing rangeEnd → reference run):

```python
def insert_comment_range_end_and_reference_below(self, comment_id: int) -> None:
    self.addnext(self._new_comment_reference_run(comment_id))
    self.addnext(OxmlElement("w:commentRangeEnd", attrs={qn("w:id"): str(comment_id)}))

def insert_comment_range_start_above(self, comment_id: int) -> None:
    self.addprevious(OxmlElement("w:commentRangeStart", attrs={qn("w:id"): str(comment_id)}))
```

## 5.4 Threading: commentsExtended.xml (w15)

MS-DOCX CT_CommentEx:
- **`w15:paraId`** (required, ST_LongHexNumber): "the paraId of the **LAST paragraph** in the associated comment" — the link key is the `w14:paraId` attribute on the comment's last `w:p` in comments.xml (Word also writes `w14:textId`).
- **`w15:paraIdParent`** (optional): the parent comment's paraId → makes this a reply.
- **`w15:done`** (optional ST_OnOff, default 0): resolved flag.

**paraId constraints** (MS-DOCX, https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/a0e7d2e2-2246-44c6-96e8-1cf009823615): unique within the part; **values MUST be > 0 and < 0x80000000**. `00000000` invalid; ≥ `80000000` invalid. Any element with paraId must also carry textId. docx-editor's generator is a good model (constrains to < 0x7FFFFFFF, valid for durableId too):

```python
def _generate_hex_id() -> str:
    return f"{random.randint(1, 0x7FFFFFFE):08X}"
```

Example commentsExtended body:

```xml
<w15:commentEx w15:paraId="6B21A3F0" w15:done="0"/>
<w15:commentEx w15:paraId="1E44C2A7" w15:paraIdParent="6B21A3F0" w15:done="0"/>
<!-- resolved: done="1" on the thread root (Word resolves whole threads) -->
```

The reply's anchor in document.xml duplicates the parent's range (docx-editor inserts the reply's rangeStart right after the parent's rangeStart, and the reply's rangeEnd + reference run right after the parent's reference run).

## 5.5 commentsIds.xml (w16cid) and commentsExtensible.xml (w16cex)

**commentsIds.xml:** `w16cid:commentId` entries with `w16cid:paraId` (matching the comment's last-paragraph paraId) and `w16cid:durableId` (ST_LongHexNumber, **> 0 and < 0x7FFFFFFF** per MS-DOCX CT_CommentId). Purpose: stable identity surviving paraId regeneration (co-authoring/Word online).

```xml
<w16cid:commentId w16cid:paraId="6B21A3F0" w16cid:durableId="3A7F2C11"/>
```

**Required?** Not for threading to render — threading + resolved state live entirely in commentsExtended (paraId/paraIdParent/done). No authoritative source states Word 365 requires commentsIds; Word tolerates its loss. **FLAG: "commentsExtended without commentsIds in Word 365" is not verified against a primary source — run an empirical round-trip test in Word during implementation.** Safe engineering position (what Word and docx-editor both do): write commentsIds + commentsExtensible whenever writing commentsExtended, one entry per comment.

**commentsExtensible.xml:** `w16cex:commentExtensible` keyed by `w16cex:durableId`, optionally `w16cex:dateUtc` (modern Word stores true-UTC here) + extLst. Optional; write for Word parity.

## 5.6 word/people.xml (w15)

Optional for display — comments render fine with just `w:author`:

```xml
<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:person w15:author="Jane Smith">
    <w15:presenceInfo w15:providerId="None" w15:userId="Jane Smith"/>
  </w15:person>
</w15:people>
```

`providerId`: "None" for local, "AD"/"Windows Live" for real identities. `w15:author` must exactly match the `w:author` string on comments/revisions.

## 5.7 MINIMAL writer sets

**Flat comment (all Word versions):** comments.xml with the `w:comment` + anchor triplet (or minimally just the reference run) + content-type override + relationship. That is literally all python-docx 1.2.0 produces; Word opens it fine and regenerates modern parts on next save.

**Threaded reply (renders as thread in 2019/365):** the above, PLUS valid unique `w14:paraId` on every comment's last content paragraph; commentsExtended.xml with one `w15:commentEx` per comment (replies carrying `paraIdParent`); override + relationship for commentsExtended. Strongly recommended for parity: commentsIds + commentsExtensible with matching durableIds, and people.xml.

**Resolved thread:** threaded set with `w15:done="1"` on the thread root's commentEx (setting it on replies too matches Word's resolve-whole-thread behavior).

## 5.8 python-docx 1.2.0 comment support — exact coverage

Added in 1.2.0: `document.add_comment(runs, text, author, initials)`, `document.comments`, `Comments.add_comment()`, rich content (paragraphs AND tables) via BlockItemContainer. Anchoring "only on an even run boundary". **NOT covered: no w14:paraId on comment paragraphs, no commentsExtended/commentsIds/commentsExtensible/people at all (opc/constants.py has no constants for them), no replies, no resolve** — docs state threading/resolution are "not currently planned." A custom server uses python-docx for base comment + anchoring and lxml for the extension parts (pattern 5.4–5.7). Docs: https://python-docx.readthedocs.io/en/latest/user/comments.html

## 5.9 Reference implementations

**GongRzhe/Office-Word-MCP-Server (MIT):** extraction-only (get_all_comments etc.); finds the comments part via `document_part.rels`, XPaths `.//w:comment`. No creation, no threading. Reading pattern only.

**pablospe/docx-editor (MIT, https://github.com/pablospe/docx-editor) — the best writing model found.** `CommentManager` (docx_editor/comments.py) manages all 5 parts: `add_comment` (comment id + paraId + durableId generation, marker placement, writes all four comment parts), `reply_to_comment` (adjacent markers + paraIdParent), `resolve_comment` (`w15:done="1"`), `delete_comment` (removes markers + entries from every part). Its `DocxXMLEditor._inject_attributes_to_nodes` auto-stamps `w:author`/`w:date`/`w:initials` on `w:comment`, `w14:paraId`/`w14:textId` on new `w:p`, `w16cex:dateUtc` on commentExtensible. Its `ooxml/templates/` directory has copy-safe part templates with the full namespace battery.

## 5.10 Comment pitfalls

1. Mismatched `w:id` across rangeStart/rangeEnd/reference/comment → orphaned markers or unanchored comment; unbalanced rangeStart/End = markup-range imbalance → repair.
2. Missing reference run → comment never displays (ranges alone don't attach it).
3. Missing content-type Override for any added part → repair; part present with no relationship → orphaned, dropped or repair.
4. Duplicate/invalid paraIds (zero, ≥ 0x80000000, duplicated) → threading breaks; has caused repair (docx4j bug history).
5. commentsExtended keyed to a paraId that is not the comment's LAST paragraph → thread linkage silently fails on multi-paragraph comments.
6. Unescaped `&`/`<` in comment text → part corruption (assign via lxml `.text`, never string-build).
7. Comments in headers/footers or nested → silently removed or repair.
8. **`w:commentRangeStart` shares the annotation-id space with bookmarks and move ranges** — real-world Word rejection over duplicated low ids across bookmark/tracked-change ranges (https://github.com/anthropics/skills/issues/489). Allocate from `max(all w:id across ALL id-bearing elements) + 1`.

---

# TOPIC 6 — Tracked changes (revisions) XML

## 6.1 Run-level insertions and deletions

```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-07-30T23:05:00Z">
  <w:r w:rsidR="00792858"><w:t>inserted text</w:t></w:r>
</w:ins>

<w:del w:id="2" w:author="Claude" w:date="2025-07-30T23:05:00Z">
  <w:r w:rsidDel="00792858"><w:delText>deleted text</w:delText></w:r>
</w:del>
```

Rules:
- `w:ins`/`w:del` wrap **complete `w:r` elements** at paragraph-content level. **Never nest revision tags inside `w:r`** — invalid XML, documented non-deterministic corruption (anthropics/skills#489).
- Text inside `w:del` **must** be `w:delText` (deleted field code text: `w:delInstrText`). `w:t` inside `w:del` = invalid → repair, or deleted text resurrects on accept. Symmetrically, `w:delText` outside a `w:del` is invalid.
- Attributes (CT_TrackChange): `w:id` required (decimal), `w:author` required, `w:date` optional (modern Word adds `w16du:dateUtc`).
- Runs inside `w:del` carry `w:rsidDel` instead of `w:rsidR`; swap back on reject-restore.
- Keep `xml:space="preserve"` on whitespace-edged `w:t`/`w:delText`.
- Block level: inserted paragraph = `w:p` inside body-level `w:ins`, or more commonly an inserted paragraph *mark* (6.3).

## 6.2 Property changes

Each `*PrChange` sits **inside** the corresponding properties element, carries `w:id`/`w:author`/`w:date`, and contains the **PREVIOUS** properties:

| Element | Lives in | Holds previous |
|---|---|---|
| `w:rPrChange` | `w:rPr` | old `w:rPr` |
| `w:pPrChange` | `w:pPr` | old `w:pPr` (incl. old `w:numPr` for numbering changes) |
| `w:tblPrChange` | `w:tblPr` | old `w:tblPr` |
| `w:tblGridChange` | `w:tblGrid` | old `w:tblGrid` |
| `w:tblPrExChange` | `w:tblPrEx` | old exceptions |
| `w:trPrChange` | `w:trPr` | old `w:trPr` |
| `w:tcPrChange` | `w:tcPr` | old `w:tcPr` |
| `w:sectPrChange` | `w:sectPr` | old `w:sectPr` |

```xml
<w:r>
  <w:rPr>
    <w:b/>
    <w:rPrChange w:id="7" w:author="Jane" w:date="2026-08-01T10:00:00Z">
      <w:rPr/>   <!-- previously: not bold -->
    </w:rPrChange>
  </w:rPr>
  <w:t>now bold</w:t>
</w:r>
```

Accept = delete the `*PrChange` (current props stand). Reject = replace parent properties with the stored previous, delete the element. Mind pPr child order on rebuild: `w:rPr` before `w:sectPr`/`w:pPrChange` (schema sequence).

## 6.3 Paragraph-mark revisions

Empty `w:ins` or `w:del` inside `w:pPr/w:rPr` revises the paragraph MARK itself:

```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <w:del w:id="0" w:author="Eric White" w:date="2009-08-21T15:38:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:r><w:t>You can use these.</w:t></w:r>
</w:p>
```

Semantics (Eric White, "Accepting Revisions in Open XML Word-Processing Documents" — https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/ee836138(v=office.12), the canonical algorithm reference; its code became Open-Xml-PowerTools RevisionAccepter, Ms-PL):

- **Deleted mark, accept:** group consecutive marked paragraphs, merge their content with the following paragraph; **the merged paragraph takes the FOLLOWING paragraph's pPr.** Content controls must keep enclosing the same runs (block-level `w:sdt` demotes to run level when necessary); block-level customXml stays block-level. Reject = drop the `w:del` from the mark's rPr.
- **Inserted mark** (tracked paragraph *split*): accept = drop the `w:ins` (split permanent). Reject = remove the mark AND merge the following paragraph back (un-split), migrating any downstream pending mark.

## 6.4 Moves

`w:moveFrom` (source, text stays `w:t` not delText in Word output), `w:moveTo` (destination), plus `w:moveFromRangeStart/End` and `w:moveToRangeStart/End`. Schema for the range starts (datypic): `w:id` required (pairs start↔end on the same side), **`w:name` required — the `w:name` value is what pairs the moveFrom range with its moveTo range** (both sides share e.g. `move240623472`), `w:author`/`w:date` required.

```xml
<w:p>
  <w:moveFromRangeStart w:id="2" w:author="Eric White"
        w:date="2009-09-13T16:42:00Z" w:name="move240623472" />
  <w:moveFrom w:id="3" w:author="Eric White" w:date="2009-09-13T16:42:00Z">
    <w:r><w:t>Goodbye.</w:t></w:r>
  </w:moveFrom>
</w:p>
<w:moveFromRangeEnd w:id="2" />
```

Accept: moveFrom content **removed**, moveTo **collapsed** (unwrapped), all four markers removed; a paragraph whose end tag only is inside a moveFrom range → its mark is treated as a deleted paragraph mark. `w:moveFrom` on a paragraph mark behaves like a deleted mark (grouping + merge-forward). Reject: exact inverse (moveFrom unwrapped in place, moveTo content removed, markers removed).

## 6.5 Table revisions

- **Inserted row:** `w:tr/w:trPr/w:ins`. Accept = drop the `w:ins`; reject = remove the `w:tr`.
- **Deleted row:** `w:tr/w:trPr/w:del` (cell text is delText, marks carry del). Accept = remove the `w:tr`; reject = drop the del + restore descendants (delText→t, drop mark dels).
- **`w:cellIns`** (`w:tcPr`): accept = remove element; reject = remove the cell (with gridSpan bookkeeping mirrored from cellDel).
- **`w:cellDel`:** accept = remove the deleted cells AND add their gridSpans (default 1) to the `w:gridSpan` of the cell immediately before the group; reject = drop the element.
- **`w:cellMerge`** (tracked vertical merge from Compare/legal-blackline output): accept transforms attribute `w:vMerge="rest"` → `<w:vMerge w:val="restart"/>`, `"cont"` → `w:val="continue"`; reject = drop the element. Live Word editing doesn't produce cellIns/cellDel/cellMerge (Compare does) — a processor must still handle them.

## 6.6 Nested revisions (author B deletes author A's insertion)

The `w:del` nests **inside** the other author's `w:ins`, wrapping the affected runs, whose text converts to `w:delText`:

```xml
<w:ins w:author="Jane Smith" w:id="16">
  <w:del w:author="Claude" w:id="40">
    <w:r><w:delText>monthly</w:delText></w:r>
  </w:del>
</w:ins>
<w:ins w:author="Claude" w:id="41">
  <w:r><w:t>weekly</w:t></w:r>
</w:ins>
```

Preserve the outer `w:ins`'s author. To "restore" another author's deletion, never edit their `w:del` — add your own `w:ins` re-inserting the text after it.

Matrix for `ins(A) ⊃ del(B)`:

| Action | Result |
|---|---|
| Accept del(B) | remove inner del + content; a now-empty ins collapses to nothing |
| Reject del(B) | unwrap del, delText→t; content remains a pending insertion by A |
| Accept ins(A) | unwrap outer ins; inner del(B) remains a normal pending deletion |
| Reject ins(A) | remove everything |
| Accept all | text disappears |
| Reject all | text disappears |

Process nested structures with an iterate-until-stable loop (descending id order within each pass; repeat until no listed revision can be processed).

## 6.7 numberingChange

`w:numberingChange` (id/author/date/`original`) is a legacy Word-2003-era mechanism kept in ISO 29500 Transitional; modern Word tracks numbering via `w:pPrChange` holding the old `w:numPr`. Accept = remove; reject = prefer honoring pPrChange (`original` is a display string, not reliably reversible).

## 6.8 FULL accept/reject matrix

Accept column grounded in Eric White's article; reject is the inverse, matching pablospe/docx-editor's implementation.

| Revision markup | ACCEPT | REJECT |
|---|---|---|
| `w:ins` (run content) | Unwrap, keep children | Remove element + content |
| `w:del` (run content) | Remove element + content | Unwrap; every `w:delText`→`w:t`, `w:delInstrText`→`w:instrText`; restore `w:rsidR` from `w:rsidDel` |
| `w:ins` in `w:pPr/w:rPr` (inserted ¶ mark) | Remove the `w:ins` (split stands) | Remove mark AND merge following ¶ into this one; migrate downstream pending mark |
| `w:del` in `w:pPr/w:rPr` (deleted ¶ mark) | Group consecutive marked ¶s; merge content with following ¶; new ¶ takes the FOLLOWING ¶'s pPr; honor sdt/customXml rules | Remove the `w:del` |
| `w:ins` wrapping whole `w:p` | Unwrap (promote the `w:p`) | Remove `w:ins` + `w:p` |
| `w:del` wrapping whole `w:p` | Remove `w:del` + `w:p` | Unwrap; delText→t inside |
| `w:rPrChange` / `w:pPrChange` / `w:tblPrChange` / `w:tblGridChange` / `w:tblPrExChange` / `w:trPrChange` / `w:tcPrChange` / `w:sectPrChange` | Delete the element | Replace parent props with stored previous; delete element |
| `w:numberingChange` | Remove | Remove after restoring numbering via pPrChange if present |
| `w:moveFrom` (run content) | Remove content | Unwrap (content stays at source) |
| `w:moveTo` (run content) | Unwrap | Remove content |
| `w:moveFrom` on ¶ mark | Same as deleted ¶ mark | Remove the mark element |
| `w:moveFromRangeStart/End` | Remove markers; contained content removed; ¶-end-only-inside → deleted-mark handling | Remove markers; content stays |
| `w:moveToRangeStart/End` | Remove markers | Remove markers + contained moveTo content |
| `w:trPr/w:ins` (inserted row) | Remove the `w:ins` | Remove the `w:tr` |
| `w:trPr/w:del` (deleted row) | Remove the `w:tr` | Remove the `w:del`; restore descendants |
| `w:cellIns` | Remove element | Remove cell (gridSpan bookkeeping) |
| `w:cellDel` | Remove deleted cells; add their spans to preceding cell's gridSpan | Remove element |
| `w:cellMerge` | `vMerge="rest"`→restart, `"cont"`→continue | Remove element |
| `w:ins`/`w:del` in `m:ctrlPr` (math) | del: remove the `m:f` construct; ins: collapse | inverse |
| `w:customXmlInsRangeStart/End` | Remove markers (control stays) | Collapse the matched sdt/customXml + remove markers |
| `w:customXmlDelRangeStart/End` | Collapse the sdt whose start AND end fall within a matched (same id) pair; remove markers | Remove markers only |

## 6.9 Working accept/reject code — pablospe/docx-editor (MIT)

`docx_editor/track_changes.py`, verbatim core:

```python
def accept_revision(self, revision_id, element_index=None) -> bool:
    elem = self._find_revision_element(revision_id, element_index)
    if elem is None:
        return False
    if elem.tagName == "w:ins":
        self._unwrap_element(elem)      # Accept insertion: unwrap
    else:  # w:del
        self._remove_element(elem)      # Accept deletion: remove
    return True

def reject_revision(self, revision_id, element_index=None) -> bool:
    elem = self._find_revision_element(revision_id, element_index)
    if elem is None:
        return False
    if elem.tagName == "w:ins":
        if _is_paragraph_mark_ins(elem):
            self._rejoin_paragraph(elem)   # inverse of the tracked split
        else:
            self._remove_element(elem)
    else:  # w:del
        self._restore_deletion(elem)
    return True

def _unwrap_element(self, elem) -> None:
    parent = elem.parentNode
    while elem.firstChild:
        parent.insertBefore(elem.firstChild, elem)
    parent.removeChild(elem)

def _restore_deletion(self, del_elem) -> None:
    """Restore deleted content by converting w:delText back to w:t."""
    for del_text in list(del_elem.getElementsByTagName("w:delText")):
        t_elem = self.editor.dom.createElement("w:t")
        while del_text.firstChild:
            t_elem.appendChild(del_text.firstChild)
        for i in range(del_text.attributes.length):        # keeps xml:space
            attr = del_text.attributes.item(i)
            t_elem.setAttribute(attr.name, attr.value)
        del_text.parentNode.replaceChild(t_elem, del_text)
    for run in del_elem.getElementsByTagName("w:r"):       # rsidDel -> rsidR
        if run.hasAttribute("w:rsidDel"):
            run.setAttribute("w:rsidR", run.getAttribute("w:rsidDel"))
            run.removeAttribute("w:rsidDel")
    self._unwrap_element(del_elem)
```

`accept_all(author=None)`/`reject_all(author=None)`: list revisions (author-filtered), process in **descending id order**, repeat passes until no progress (handles nesting). Writing side: a central injector auto-stamps `w:id` (monotonic, folded against pre-existing ids to never collide), `w:author`/`w:date`/`w16du:dateUtc`, `w:rsidR` (or `rsidDel` inside dels), `w14:paraId`/`textId` on new `w:p`.

Other implementations: **Open-Xml-PowerTools RevisionAccepter** (C#, **Ms-PL**, https://github.com/EricWhiteDev/Open-Xml-PowerTools) — the canonical full implementation of Eric White's algorithms including sdt/customXml edge cases. **JSv4/Python-Redlines** (MIT) — generates redlines by shelling to a bundled C# comparer (produce tracked changes from two versions; no Python accept/reject). **balalofernandez/docx-revisions** — extends python-docx with ins/del read/write.

## 6.10 Author filtering

Plain exact string match on `w:author`. Caveats: (a) Word's "Remove personal information" rewrites all authors to literal "Author"; (b) match exactly (case/whitespace) — build filter lists from the document's own distinct author values, not user input; (c) same-author adjacent revisions each have their own `w:id` — group by (author, date) or contiguity, never by id; (d) Word can emit duplicate ids across authors — when resolving by id, verify the author attribute on the found element too.

## 6.11 Tracked-changes pitfalls

1. `w:t` inside `w:del`, or orphan `w:delText` outside one → invalid; corruption or wrong accept behavior. Convert on wrap AND unwrap.
2. Removing a `w:del` but leaving `w:delText` — same bug class; `_restore_deletion` shows the correct pairing.
3. `w:ins`/`w:del` nested inside `w:r`/`w:t` → invalid XML (real, non-deterministic generation failure mode).
4. **`w:id` collisions across the shared annotation-id space** (bookmarks, comment ranges, move ranges, tracked changes): Word rejects files where tracked-change ids duplicate bookmark ids (anthropics/skills#489). Scan max `w:id` over ALL id-bearing elements, allocate max+1, fold in pre-existing ids.
5. Unbalanced moveFromRangeStart/End (or moveTo) pairs; moveFrom range with no same-`w:name` moveTo → repair / broken moves.
6. Mismatched wrappers when string-templating XML (`</w:ins>` closing a `<w:del>`).
7. Wrap only the delta — marking unchanged text as changed.
8. Missing `<w:trackRevisions/>` in settings.xml — not corruption; future Word edits just won't be tracked.
9. RSIDs must be 8-digit hex.
10. pPr child order after pPrChange operations: `w:rPr` before `w:sectPr`/`w:pPrChange` or Word may repair.
11. Deleted-mark acceptance with content controls: sdt must keep enclosing the same runs; block-level customXml stays block-level.

---

# TOPIC 7 — Docx integrity: what triggers Word's repair prompt, and validation

## 7.1 Cause catalog

**Relationship problems (top cause for programmatic writers):**
- `r:id`/`r:embed`/`r:link` referencing a relationship ID absent from that part's `.rels` → treated as corruption at load. **Relationship IDs are scoped PER PART** — an image referenced from a header needs its rel in `header1.xml.rels`, not `document.xml.rels`. Copying runs between parts without migrating relationships is the classic MCP-server bug.
- Orphaned rels (entry → missing part) usually tolerated; wrong `TargetMode` on external targets can trigger repair.

**[Content_Types].xml:** any added part without a matching `<Override>` (or `<Default>` for its extension) → repair or silent part-dropping. Adding a PNG when only `jpeg` has a Default is a common failure.

**ID collisions / malformed IDs:**
- Duplicate comment ids, duplicate bookmark ids, and — documented in the wild — **tracked-change `w:id` values colliding with existing bookmark IDs** corrupt files (https://github.com/anthropics/skills/issues/489). Safe practice: scan ALL existing `w:id` values in the part and start above the max (or high random).
- `bookmarkStart` without `bookmarkEnd` (and vice versa; also misplaced across table boundaries — pandoc #8825): sometimes silently repaired, sometimes prompts; references break either way.
- `w14:paraId`/`w14:textId`: ST_LongHexNumber, unique per part, **> 0 and < 0x80000000**; paraId requires textId alongside. Duplicates break co-authoring/change-tracking even when the file opens (Open-XML-SDK #245/#925/#962 — the SDK ships no generator; Word uses random values; docx4j historically generated out-of-range ones). **Safest for the MCP server: omit paraId/textId on new paragraphs entirely** (optional; Word regenerates) — EXCEPT comment paragraphs feeding commentsExtended, which need them (Topic 5).

**Element-order violations (xsd:sequence):** `w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr`, `w:tcPr` and most CT_* property containers have fixed child order. Word is strict about several (e.g. `w:pStyle` first in pPr; `w:rPr` first child of `w:r`; `w:tcPr` first child of `w:tc`; `w:sectPr` in pPr after everything except `w:pPrChange`). Wrong order = repair or silently dropped formatting. Required orders: see 7.2.

**Structural block rules:**
- Every `w:tc` MUST have a `w:p` as its last child block element — otherwise **Word fails to open the file** (Microsoft's own implementer note, MS-OI29500 §2.1.168). Applies after deleting a cell's last paragraph and to nested-table-last cells. python-docx's `CT_Tc.new()` templates `<w:tc><w:p/></w:tc>`; `clear_content()` admits it leaves the cell invalid until a paragraph is re-added.
- Body-level `w:sectPr` must be the LAST child of `w:body`; mid-document section breaks live inside the pPr of the section's last paragraph, never as sibling blocks. Two adjacent tables with no paragraph between get merged by Word — keep an empty trailing paragraph after tables.
- Empty `w:sdtContent` corrupts (must contain ≥1 element).

**Field character imbalance:** unmatched `fldChar begin`/`end` (typically from deleting a paragraph range containing one endpoint) → repair, or the rest of the document gets eaten into the field. **Any range-deletion tool must count fldChar balance — and commentRangeStart/End and bookmarkStart/End pairing — across the deleted range.**

**Text content:**
- XML 1.0 forbids control chars 0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F in `w:t`. lxml refuses to serialize them directly (protective), but text smuggled via pre-encoded bytes or numeric character references (`&#x02;`) produces an unreadable file. Strip/replace `[\x00-\x08\x0B\x0C\x0E-\x1F]` on every text write.
- Missing `xml:space="preserve"` on whitespace-edged `w:t`: not corruption — silent data loss on load.

**Namespace/mc:** declaring a prefix in `mc:Ignorable` with no matching namespace declaration on the root corrupts (seen when copying root attributes between docs from different Word versions — docx4j #601, the `w16sdtfl` case). Clone the nsmap when cloning root attributes.

**ZIP-level:**
- Duplicate entry names (easy with `zipfile` append-as-replace logic — `ZipFile` happily writes two `word/document.xml` entries). Never append a same-named entry; rebuild the archive.
- Non-UTF-8 entry names without the UTF-8 flag, absolute paths, backslash separators → repair.
- **Entry order does NOT matter functionally** (unlike ODF's mimetype rule); convention puts `[Content_Types].xml` first — free to do, keeps forensic/AV tools happy (SANS ISC diary on OPC: https://isc.sans.edu/diary/26662).
- Use `zipfile.ZIP_DEFLATED`, default settings; no encryption flags/spanning.

## 7.2 WHERE to find the required element orders

1. **ECMA-376 schemas** (authoritative): https://ecma-international.org/publications-and-standards/standards/ecma-376/ — Part 4 5th ed. zip contains **`OfficeOpenXML-XMLSchema-Transitional.zip` with `wml.xsd`** (the set matching what Word actually writes; Part 1 has Strict). Browsable annotated rendering: https://schemas.liquid-technologies.com/officeopenxml/2006/wml_xsd.html. Microsoft deviations: MS-OE376, MS-OI29500; `w14:*`/`w15:*` extensions are in MS-DOCX, NOT the ECMA schemas.

2. **python-docx's oxml layer** encodes the orders as `_tag_seq` tuples (verified verbatim from master):
   - `CT_PPr._tag_seq` (`src/docx/oxml/text/parfmt.py`): `("w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")`
   - `CT_SectPr._tag_seq` (`src/docx/oxml/section.py`): `("w:footnotePr", "w:endnotePr", "w:type", "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange")`
   - `CT_TblPr._tag_seq` (`src/docx/oxml/table.py`): `("w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription", "w:tblPrChange")`
   - `CT_TcPr._tag_seq` (same file): `("w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange")`
   - rPr order (`src/docx/oxml/text/font.py` / schema): `rStyle, rFonts, b, bCs, i, iCs, caps, smallCaps, strike, dstrike, outline, shadow, emboss, imprint, noProof, snapToGrid, vanish, webHidden, color, spacing, w, kern, position, sz, szCs, highlight, u, effect, bdr, shd, fitText, vertAlign, rtl, cs, em, lang, eastAsianLayout, specVanish, oMath` (+ `rPrChange` last).

3. **The insertion machinery** (`src/docx/oxml/xmlchemy.py`): each child declared via descriptors like `ZeroOrOne("w:jc", successors=_tag_seq[27:])`; the metaclass generates `_insert_x()`/`get_or_add_x()` methods calling:

   ```python
   def insert_element_before(self, elm: ElementBase, *tagnames: str):
       successor = self.first_child_found_in(*tagnames)
       if successor is not None:
           successor.addprevious(elm)
       else:
           self.append(elm)
       return elm
   ```

   **For custom lxml manipulation: never hand-position property children.** Either (a) route through python-docx's registered CT_* classes and their generated `get_or_add_x()` methods, or (b) keep one `_tag_seq` tuple per container (copied from python-docx/XSD) and implement `insert_in_sequence(parent, new_el, tag_seq)` = find the new tag's index i, call `insert_element_before(parent, new_el, *tag_seq[i+1:])`. This single helper eliminates the entire order-corruption bug class.

4. **Open XML SDK docs** — each class page on learn.microsoft.com lists child elements in schema order (cross-check).

## 7.3 Validation approaches (layered strategy)

**(a) lxml XMLSchema against ECMA-376 transitional wml.xsd** — workable with caveats. `etree.XMLSchema(etree.parse("wml.xsd"))` resolves imports via relative paths if the schema zip's layout is kept intact. Problems: real Word files contain `mc:AlternateContent`/`mc:Ignorable` (Part 3 MCE) and post-2007 namespaces (`w14`, `w15`, `w16*`) — **none validate against ECMA wml.xsd**. Practical recipe: preprocess a COPY before validating — apply the MCE transform yourself (drop attributes/elements in `mc:Ignorable`-listed namespaces, replace `mc:AlternateContent` with its `mc:Fallback`, strip `w14:paraId`/`textId`), then validate against transitional (never strict — Word writes transitional). Use to catch order/cardinality mistakes in parts YOU generated; expect noise on arbitrary user documents. https://lxml.de/validation.html

**(b) python-docx round-trip** (`Document(path)` → `save(buffer)`): catches unparseable XML, missing required parts, broken rels it traverses. Does NOT catch element order, dangling r:ids (lazy-loaded), ID dupes, fldChar imbalance, missing overrides. A floor, not a validator.

**(c) Relationship-graph + invariants audit script — write this; highest value/cost.** Unzip; for every part: collect every attribute in the relationships namespace (`r:id|embed|link|pict|dm|lo|qs|cs`), verify each resolves in that part's `.rels`; verify every internal rel Target resolves to a zip entry; verify every part has a content type. Then per-part XML checks: `w:id` uniqueness per category, bookmarkStart/End + commentRangeStart/End pairing, fldChar begin/separate/end stack balance, every `w:tc` ends with `w:p`, body ends with `w:sectPr`, control-char scan of all `w:t`. ~200 lines of Python covering essentially the whole cause catalog. Run after every mutating tool call (milliseconds).

**(d) Open XML SDK validator (dotnet) in CI** — gold standard short of Word. Prebuilt CLI: **mikeebowen/OOXML-Validator** (MIT, .NET, JSON output, Office 2007–2021 + M365 profiles; also a VS Code extension). Classic write-ups: Eric White "Validate Open XML Documents using the Open XML SDK"; Brian Jones "Finding Open XML errors with Open XML SDK validation"; trailmax "Validating OpenXml generated documents". Catches order/cardinality/attribute-range with exact part+XPath; not a perfect Word-repair oracle in either direction.

**(e) LibreOffice headless smoke test** — `soffice --headless --convert-to pdf` — weak detector (LO is more forgiving than Word) but a good "did we destroy the file" canary.

**Recommended layered strategy:** (1) prevention — all property-child inserts through `_tag_seq` logic; (2) in-process audit script (c) after every mutating call + control-char/xml:space guards on write paths (~90% of real-world causes); (3) python-docx round-trip as cheap secondary; (4) OOXML-Validator in CI against golden outputs; (5) LibreOffice optionally in CI. **Keep a copy-on-write backup of the input file before every edit session** (also mandated by this vault's file-versioning rules).

Write-ups: MS Q&A "What is invalid about this docx?" and "Word document is corrupt for unknown reason" (learn.microsoft.com/answers); dolanmiu/docx #3314; docx4j #601; anthropics/skills #489; Stefan Sommarsjö "Structure of Docx Files" (Medium).

---

# TOPIC 8 — FastMCP / Python 3.14 state (as of 2026-08-27)

**Empirical result from THIS machine (see header):** the full stack imports and runs on Python 3.14.3 / Windows 11 today — fastmcp 2.14.7, mcp 1.27.2, pydantic 2.13.4, lxml 6.0.2 (cp314 wheel), python-docx 1.2.0. The remaining questions are about current PyPI versions, not feasibility.

**PyPI current state:**
- **fastmcp** (jlowin/fastmcp, Apache-2.0): stable **3.4.7** (2026-08-10); 4.0 beta line active (4.0.0b4, 2026-08-26). `requires-python >=3.10`; classifiers list 3.10–3.13 only — **no 3.14 classifier in 3.4.x**. Changelog shows "Python 3.14 compatibility hardening" landing in **4.0.0b3**, i.e. official 3.14 support is a FastMCP 4 feature. Historical 3.14 issues on the 2.x line: fastmcp 2.10.6 vs pydantic 2.12+ Field API (ii-agent #165); poetry marker weirdness in 2.13.0 (jlowin/fastmcp #2257). Note the locally installed 2.14.7 predates the 3.x line and runs fine here.
- **mcp** official SDK (modelcontextprotocol/python-sdk, MIT): **2.1.1** (2026-08-25) — the reworked 2.x line targeting the 2026-07-28 MCP spec; classifiers **include 3.14**. (Local install is 1.27.2.)
- **pydantic 2.13.4** supports 3.14 (practical floor: ≥2.12); **pydantic-core 2.48.0** ships `cp314-cp314-win_amd64` (and cp314t) wheels.
- **anyio/starlette/uvicorn**: all green on pyreadiness.org/3.14 (irrelevant for stdio-only transport anyway).
- **python-docx 1.2.0** (2025-06-16, MIT, pure Python, `lxml>=3.1.0`): **added the comments API** (`Document.add_comment()`, `Comments`); **footnotes/endnotes still NOT supported in any release** → footnote work means raw lxml against word/footnotes.xml (Topic 3) or porting bayoo-docx patterns.
- **lxml 6.1.2** (2026-08-19, BSD-3-Clause): **`cp314-cp314-win_amd64` wheel exists on PyPI** (plus cp314t, cp315).

**Verdict:** Windows/3.14 binary wheels are fully solved; the official `mcp` SDK explicitly supports 3.14; python-docx is pure Python. The one soft spot is fastmcp 3.4.7 not declaring 3.14 (hardening lands in 4.0). Options, in order of preference for this project: (1) **run Python 3.14 with fastmcp pinned + the smoke test already passing on this machine** (empirically fine), planning a move to FastMCP 4.0 stable when released; (2) build on the official `mcp` SDK's FastMCP-style API instead (clean 3.14 today, fewer conveniences); (3) pin 3.13 for zero risk. Avoid 3.14t (free-threaded) regardless. **Regardless of choice: use a dedicated venv — the shared site-packages has the bayoo-docx/python-docx shadowing conflict (see header).**

---

# TOPIC 9 — License audit for code reuse

Target: an MIT-style private project. Licenses verified via LICENSE files / GitHub API.

| Repo | License (verified) | Reuse into MIT/private project |
|---|---|---|
| **SecurityRonin/docx-mcp** | **MIT** (GitHub API spdx: MIT) | Copy/port/adapt freely; retain copyright + license notice for copied substantial portions. Best footnote-CRUD reference (Topic 3.9). |
| **Rookie0x80/docx-mcp** | **NO LICENSE FILE** — GitHub API `/license` 404, repo license field none. (Its pyproject.toml declares MIT text, but with no LICENSE file the legal default is **all rights reserved**.) | **Do NOT copy code.** GitHub ToS grants viewing/forking only. Learn at the idea/algorithm level only (and its table code has bugs anyway — Topic 1.4). Or ask the author to add a license. |
| **BayooG/bayoo-docx** | **MIT in substance** — LICENSE is standard MIT text, © 2019 Obay Daba, noting the python-docx fork; GitHub API says "NOASSERTION" only because header edits break the automatic matcher | Copy/port/adapt with attribution to Obay Daba AND the upstream python-docx MIT notice. The useful one for footnote patterns. Reproduce its LICENSE verbatim in NOTICE if copying substantial code. |
| **GongRzhe/Office-Word-MCP-Server** | **MIT** (GitHub API spdx: MIT) | Copy/port/adapt freely with notice. (The server behind the currently installed `office-word` tools; its search_and_replace is naive — Topic 2.1 — but other tool shapes are fair game.) |
| **ykarapazar/word-mcp-live** | **MIT** (GitHub API spdx: MIT; repo name verified exact) | Copy/port/adapt freely with notice. NB: drives live Word via COM (win32com) — a different architecture from file-based editing. |

**Additional repos encountered during research:**

| Repo | License | Notes |
|---|---|---|
| python-openxml/python-docx | MIT | Foundation library; oxml layer is the element-order oracle |
| pablospe/docx-editor | MIT | Best comments-threading + tracked-changes accept/reject reference (Topics 5, 6) |
| ivanbicalho/python-docx-replace | MIT | KeyChanger fragment-safe replace (Topic 2.3) |
| elapouya/python-docx-template (docxtpl) | **LGPL-2.1** | Depend on, don't vendor; its patch_xml approach can be reimplemented from the idea |
| dolanmiu/docx (TypeScript) | MIT | TOC switch documentation |
| jgm/pandoc | **GPL-2.0-or-later** | Its data/docx templates mirror Word's own output — the XML strings are spec constants, but don't vendor pandoc code |
| harvard-lil/h2o | **AGPL-3.0** | TOC lua template read for the SDT shape only — the XML structure is spec-standard; do NOT copy code |
| EricWhiteDev/Open-Xml-PowerTools | **Ms-PL** | Canonical revision accept/reject (C#); Ms-PL is permissive but NOT MIT-compatible for relicensing copied code — treat as algorithm reference (the algorithms are documented in the MSDN article anyway) |
| JSv4/Python-Redlines | MIT | Redline generation via bundled C# comparer |
| plutext/docx4j | Apache-2.0 | Content-type/namespace constant verification |
| mikeebowen/OOXML-Validator | MIT | CI validation CLI |
| adejones replace gist | unstated | Reference-only; reimplement |
| anthropics/skills docx OOXML reference | license not verified | Treat as documentation reference, not vendored code |

**Foundation deps:** python-docx MIT; lxml BSD-3-Clause; fastmcp Apache-2.0 (depend-on only → nothing required; if vendoring source, keep license/NOTICE — it also carries an express patent grant); official `mcp` SDK MIT. All permissive, all compatible.

**Copyleft flags:** none of the five target repos is GPL/AGPL; the only red flag is **Rookie0x80/docx-mcp's missing license — legally MORE restrictive than GPL** (no permission at all). AGPL (h2o) and GPL (pandoc) repos: XML *shapes* are spec facts and freely usable; their *code* is not.

**Practical setup:** own MIT LICENSE + `THIRD_PARTY_NOTICES.md` reproducing the MIT notices of python-docx, bayoo-docx, SecurityRonin/docx-mcp, GongRzhe, ykarapazar, pablospe/docx-editor, ivanbicalho for any ported files, and lxml's BSD-3 notice if redistributing. Belt-and-suspenders for a private tool; future-proofs open-sourcing.

---

# Cross-cutting build recommendations

1. **Architecture:** python-docx 1.2.0 for the OPC container, part plumbing (`relate_to` auto-handles rels + content types), comments base API, and the oxml `_tag_seq` machinery; raw lxml for footnotes/endnotes, comment threading parts, TOC fields, tracked changes, and merge-aware table surgery. Dedicated venv (bayoo-docx conflict).
2. **One shared helper eliminates the biggest corruption class:** `insert_in_sequence(parent, new_el, tag_seq)` per 7.2 — never hand-position property children.
3. **One shared ID allocator:** scan max `w:id` across ALL id-bearing elements in all parts (bookmarks, comments, revisions, move ranges), allocate max+1 (SecurityRonin's `_next_global_markup_id` pattern) — prevents the documented Word-rejection collision bug.
4. **Range-deletion guard:** any tool deleting paragraph/run ranges must balance-check fldChar, bookmarkStart/End, commentRangeStart/End across the doomed range (delete whole constructs or refuse).
5. **Post-edit audit:** run the Topic 7.3(c) invariants script after every mutating tool call; python-docx round-trip as secondary; OOXML-Validator + LibreOffice in the test suite, not per-request.
6. **Copy-on-write:** back up the input docx before every edit session (aligns with the vault's never-destroy-user-edits rule).
7. **No library exists for merge-aware column ops or full comment threading in Python — build them from the algorithms in Topics 1 and 5**; the accept/reject matrix in 6.8 is complete enough to code from directly.
8. **Empirical tests to run in Word during implementation** (flagged unverified): (a) threading render with commentsExtended but no commentsIds; (b) the chosen fastmcp version's full tool round-trip on 3.14.

---

*Research compiled 2026-08-27 from four parallel research passes; all code quoted under the licenses noted inline. Local environment facts verified directly on this machine the same day.*
