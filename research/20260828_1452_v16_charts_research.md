# v1.6 Native Charts — Research Document

**Date:** 2026-08-28 14:52 KST. **Status:** research-only deliverable for the builder agent.
**Scope:** add_chart (bar/line/pie/scatter from JSON/CSV data) + update_chart_data for existing charts, at the OOXML level via lxml, no rendering. Built on the existing `DocxPackage` layer (`src/word_mcp/core/package.py`) and the house part-creation pattern from `src/word_mcp/ops/media.py`.

---

## Q1. Chart part anatomy

### Parts and identifiers (verified against ECMA-376 fundamentals text)

| Item | Value |
|---|---|
| Chart part name (convention) | `word/charts/chart1.xml` (Word uses `charts/chartN.xml`; any name works, convention keeps Word-native shape) |
| Chart part content type | `application/vnd.openxmlformats-officedocument.drawingml.chart+xml` (Override in `[Content_Types].xml`, NOT a Default — it is a per-part XML content type) |
| Relationship type (document.xml.rels -> chart part) | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart` |
| Chart XML root | `c:chartSpace` |
| Chart namespace (`c:`) | `http://schemas.openxmlformats.org/drawingml/2006/chart` |
| DrawingML main (`a:`) | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| Relationships (`r:`) | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |

Source: ECMA-376 Part 1 Fundamentals, Chart Part (c-rex mirror): https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Chart_topic_ID0ELZLM.html — also confirms: the chart part may itself be the SOURCE of relationships to Chart Drawing parts and Embedded Package parts (the xlsx), must be targeted internally (same package), and "for WordprocessingML... the data for a chart is stored in an embedded SpreadsheetML package targeted by an Embedded Package part specified by that Chart part" (normative wording; see Q2 for what Word actually tolerates).

### How the chart hangs off document.xml

Identical drawing scaffold to images (see `add_image` in `src/word_mcp/ops/media.py` lines 144–188), except `a:graphicData/@uri` is the chart namespace and the payload is a single `c:chart` element carrying `r:id`, instead of `pic:pic`:

```xml
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="5486400" cy="3200400"/>
        <wp:docPr id="7" name="Chart 7"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                     r:id="rId8"/>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
```

Notes:
- `rId8` lives in `word/_rels/document.xml.rels` with the chart relationship type above and `Target="charts/chart1.xml"`.
- No `pic:` namespace involved. No `wp:effectExtent`/`wp:cNvGraphicFramePr` needed (Word writes them, tolerates their absence — same as our image inserter, which already omits them and opens clean).
- `wp:docPr/@id` must be unique across the document — the repo already has the helper `_next_docpr_id(pkg)` in `media.py` (lines 61–66), which scans every `wp:docPr` in document.xml and returns max+1. Reuse it (move to a shared home or import from media).

Sources: ECMA/officeopenxml inline drawing anatomy http://officeopenxml.com/drwPicInline.php ; python-docx add-chart PR showing the exact chart graphicData shape https://github.com/python-openxml/python-docx/pull/392/commits/6ac9b73d67c3f2d133550c3f86a9697a605c41d6 ; a:chart reference https://ooxml.org/drawingml/a-chart/

### Chart part rels

`word/charts/_rels/chart1.xml.rels` — needed only if the chart references other parts. For us: one relationship to the embedded workbook (Q2), optionally two more to the MS chart-style parts (Q4/Q7, we will NOT emit them). If we emit no externalData and no style parts, the chart part needs no .rels file at all.

---

## Q2. The embedded workbook

### Is it required to render? NO.

The caches (and/or literal data) inside chart1.xml are what render — confirmed from multiple independent directions:

- The `externalData` element is schema-optional (`[0..1]` in CT_ChartSpace — datypic content model, and ExternalData class docs https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.externaldata?view=openxml-2.8.1 ).
- Docxtemplater's chart module ships charts with no embedded workbook: "the externalData element is optional and Word can render charts without it and an embedded spreadsheet. However, without an embedded spreadsheet, you cannot right-click in Word to edit the chart data." https://docxtemplater.com/modules/chart/
- PHPWord has generated cache-only/literal-only charts for years; they render fine, and the long-standing complaint (issue #956) is exactly and only that Edit Data does nothing: charts "still render visually in the document," but right-click > Edit Data silently no-ops because there is no `externalData` rel and no embedded xlsx. https://github.com/PHPOffice/PHPWord/issues/956
- The strict ECMA wording ("shall be stored in an embedded SpreadsheetML package") describes what conforming producers like Word write, not what Word's consumer requires. Word is lenient here.

### What breaks when it's missing or stale

| Situation | Render | Edit Data (right-click) | "Refresh Data" / chart refresh | Repair prompt |
|---|---|---|---|---|
| No externalData + no xlsx (caches/literals only) | OK | Silently does nothing (PHPWord #956) or Word offers to create data — either way degraded UX, never corruption | Nothing to refresh | None |
| externalData r:id present but target part missing | — | — | — | DANGEROUS: a relationship pointing at a nonexistent part is a package-integrity error; Word can declare unreadable content. Never emit the rel without the part. |
| xlsx present but stale (caches updated, xlsx not) | Renders from caches (Word paints from cache first) | Opens the OLD numbers — user edits stale data and the chart snaps back to it on refresh | Refresh reloads FROM the xlsx, silently reverting the chart to stale data (botched-deployments post: "if you then click the refresh button on the chart, it will reload from the embedded excel file") | None, but it is a silent-data-loss trap |
| xlsx updated, caches not | Renders the OLD cached numbers until a refresh | Opens correct data | Refresh fixes the picture | None |

Conclusion: cache and workbook are a redundant pair with cache winning at paint time and workbook winning at edit/refresh time. Both must always be written together (Q5). Sources: Eric White, "Update Cached Data and Embedded XLSX for Charts in DOCX, PPTX" http://www.ericwhite.com/blog/update-cached-data-and-embedded-xlsx-for-charts-in-docx-pptx/ ; https://botched-deployments.com/posts/python-docx-charts ; https://docxtemplater.com/modules/chart/

### DECISION: emit the embedded workbook. Always.

Rationale: the server's brand is "engineered not to corrupt" AND Word-native behavior. A chart where Edit Data silently fails is the #1 complaint against every library that skipped the workbook (PHPWord #956, early docx4j threads). Cost is low (Q2 minimal workbook below).

### Wiring

- Part name: `word/embeddings/Microsoft_Excel_Worksheet1.xlsx` (Word's own name; increment the digit per chart — scan existing `word/embeddings/` names the way media.py scans `imageN`).
- Content type: **Default** entry for extension `xlsx` -> `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` in `[Content_Types].xml` (an embedded package keeps its native content type; a Default by extension is what Word itself writes for embedded xlsx).
- Relationship: in `word/charts/_rels/chart1.xml.rels` (source = the CHART part, not document.xml):
  `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"` `Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"`.
  (Embedded Package part relationship; this is what Word writes for chart data workbooks. Eric White forum thread on embeddedPackagePart concurs: http://www.ericwhite.com/blog/forums/topic/adding-workbook-to-chart-as-an-embeddedpackagepart/ )
- In chart1.xml: `<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>` — `autoUpdate` is a CHILD element (CT_Boolean), val 0 = do not auto-refresh from the workbook on open (Word's own default for embedded chart data; also protects our cache-is-truth model). externalData's position in chartSpace is fixed by schema: AFTER `c:chart`/`c:spPr`/`c:txPr` (see Q3 ordering).

### Minimal valid embedded workbook, WITHOUT openpyxl

Five parts, hand-writable strings + `zipfile` (repo already builds zips this way in `DocxPackage.save`). Verified minimal structure (Brendan Long "minimum viable XLSX", MS Technet PowerShell-generated xlsx, professor-excel structure walkthrough):

```
[Content_Types].xml
_rels/.rels
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/worksheets/sheet1.xml
```

`[Content_Types].xml`:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
```

`_rels/.rels`:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>
```

`xl/workbook.xml`:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>
```

`xl/_rels/workbook.xml.rels`:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>
```

`xl/worksheets/sheet1.xml` — lay data out Word-style: A1 blank, categories (or X values) in column A starting A2, one series per column starting B1 (B1 = series name, B2.. = values). Use `t="inlineStr"` for strings so no sharedStrings part is needed, bare `<v>` for numbers:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="B1" t="inlineStr"><is><t>Series 1</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Alpha</t></is></c>
      <c r="B2"><v>4.3</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Beta</t></is></c>
      <c r="B3"><v>2.5</v></c>
    </row>
  </sheetData>
</worksheet>
```

What is safely OMITTED (all optional per schema; Excel and Word's embedded-edit both open such files): `docProps/*`, `xl/styles.xml`, `xl/theme/theme1.xml`, `xl/sharedStrings.xml` (avoided via inlineStr), `dimension`, `sheetViews`, `sheetFormatPr`, `calcChain`. Sources: https://www.brendanlong.com/the-minimum-viable-xlsx-reader.html ; https://social.technet.microsoft.com/wiki/contents/articles/19601.powershell-generate-real-excel-xlsx-files-without-excel.aspx ; https://professor-excel.com/xml-zip-excel-file-structure/

### Dependency recommendation: do NOT add openpyxl

- Current deps (`pyproject.toml`): fastmcp, latex2mathml, mathml2omml, python-docx, lxml, regex, pywin32. No openpyxl, no spreadsheet stack.
- The workbook we need is a fixed 5-part shell plus one trivially generated sheetData block. openpyxl would add a dependency (plus its et_xmlfile dep) to write ~40 lines of XML we fully control, and openpyxl's own chart writer is irrelevant here (it writes xlsx-native charts with bare numRef and NO caches, relying on Excel to populate them — the exact opposite of what a Word chart needs, so nothing to reuse).
- For update_chart_data we must also PARSE the existing embedded workbook; for our regenerate-in-full strategy (Q5) we never parse it, we rewrite it, so no reader dependency either.
- Verdict: hand-rolled writer module (`_chart_xlsx.py` or a function inside the ops module), stdlib `zipfile` + string templates or lxml. Zero new dependencies.

---

## Q3. Minimal valid chart XML per type (the load-bearing section)

### 3.0 Global rules that decide clean-open vs repair prompt

1. **Element ORDER inside every CT_* is fixed by xsd:sequence and Word's chart consumer enforces it.** Out-of-order children in chart1.xml (e.g. externalData before chart, c:cat after c:val, axId before ser) is the classic "unreadable content" trigger reported across docx4j/pptx4j threads (blank chart or repair dialog): https://www.docx4java.org/forums/pptx-java-f14/need-help-to-create-chart-in-ppt-slide-t1645.html . Build with fixed literal templates, never by appending in call order.
2. **Paired axes are mandatory for bar/line/scatter:** the plot-group element must carry exactly the `c:axId` values of axes that exist in plotArea, each axis's `c:crossAx` must point at the other. A dangling axId = repair. Pie has NO axes (schema: CT_PieChart has no axId child — confirmed via OpenXML SDK PieChart child list https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.piechart ).
3. **axId values:** any 32-bit int, unique within the chart part, both references consistent. Word uses big negative/positive ints; value is arbitrary. We should generate two constants or derive from a counter.
4. **Caches are what render.** With `numRef`/`strRef`, the `c:f` formula points at the workbook and the `numCache`/`strCache` sibling holds the literal points; Word paints from the cache without touching the workbook (Q2 evidence). `ptCount/@val` MUST equal the logical point count; `c:pt/@idx` are 0-based and may be sparse (gaps = missing data points) but must be < ptCount. A ptCount that disagrees with the pts is tolerated visually in some versions but is exactly the kind of half-valid state we refuse to produce — always write it exactly.
5. **Ref-vs-Lit:** `strLit`/`numLit` (no `c:f`, no workbook link) are schema-valid and render (PHPWord ships them), but Word-NATIVE charts always use `strRef`/`numRef` + cache + `c:f` ("Sheet1!$A$2:$A$4"), and Edit Data only round-trips sensibly when the refs match the embedded workbook layout. DECISION: emit Ref+cache with `c:f` matching the Q2 worksheet layout.
6. **Number formatting:** `c:formatCode` inside numCache is optional; safe default is to emit `<c:formatCode>General</c:formatCode>`.
7. `c:chart` requires only `c:plotArea` inside it; `c:plotVisOnly`, `c:autoTitleDeleted`, `c:dispBlanksAs`, legend, title are all optional. plotArea requires `c:layout` first child in Word-native files — schema allows omitting it, Word tolerates `<c:layout/>` empty; EMIT the empty `<c:layout/>` (it is what Word writes and costs nothing).
8. Root `c:chartSpace` sequence (verified, datypic CT_ChartSpace http://www.datypic.com/sc/ooxml/e-draw-chart_chartSpace.html ): `date1904? lang? roundedCorners? style? clrMapOvr? pivotSource? protection? chart! spPr? txPr? externalData? printSettings? userShapes? extLst?` — `chart` is the only required child; **externalData comes AFTER chart/spPr/txPr**.
9. Declare exactly `xmlns:c`, `xmlns:a`, `xmlns:r` on `c:chartSpace`. No mc/c14/c16 needed when we don't emit any extension content (Q7).

### 3.1 Shared skeleton (all four types)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <!-- optional: <c:title>…</c:title><c:autoTitleDeleted val="0"/> -->
    <c:plotArea>
      <c:layout/>
      <!-- ONE plot-group element from 3.2–3.5 -->
      <!-- axes for bar/line/scatter; nothing for pie -->
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:externalData r:id="rId1">
    <c:autoUpdate val="0"/>
  </c:externalData>
</c:chartSpace>
```

Optional title block (only when the tool's `title` param is given), placed as FIRST child of `c:chart`:
```xml
<c:title>
  <c:tx><c:rich>
    <a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:t>TITLE TEXT</a:t></a:r></a:p>
  </c:rich></c:tx>
  <c:overlay val="0"/>
</c:title>
<c:autoTitleDeleted val="0"/>
```

Shared series data blocks (category types: bar, line, pie). Series i, categories in A2:A(n+1), values in column letter L (B for series 0, C for 1, ...):

```xml
<c:ser>
  <c:idx val="0"/>
  <c:order val="0"/>
  <c:tx>
    <c:strRef>
      <c:f>Sheet1!$B$1</c:f>
      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache>
    </c:strRef>
  </c:tx>
  <c:cat>
    <c:strRef>
      <c:f>Sheet1!$A$2:$A$4</c:f>
      <c:strCache>
        <c:ptCount val="3"/>
        <c:pt idx="0"><c:v>Alpha</c:v></c:pt>
        <c:pt idx="1"><c:v>Beta</c:v></c:pt>
        <c:pt idx="2"><c:v>Gamma</c:v></c:pt>
      </c:strCache>
    </c:strRef>
  </c:cat>
  <c:val>
    <c:numRef>
      <c:f>Sheet1!$B$2:$B$4</c:f>
      <c:numCache>
        <c:formatCode>General</c:formatCode>
        <c:ptCount val="3"/>
        <c:pt idx="0"><c:v>4.3</c:v></c:pt>
        <c:pt idx="1"><c:v>2.5</c:v></c:pt>
        <c:pt idx="2"><c:v>3.5</c:v></c:pt>
      </c:numCache>
    </c:numRef>
  </c:val>
</c:ser>
```

**Series child ORDER (CT_BarSer, from the SDK ChildElementInfo sequence, https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.barchartseries ):**
`idx! order! tx? spPr? invertIfNegative? pictureOptions? dPt* dLbls? trendline* errBars? cat? val? shape? extLst?`
Line series inserts `marker? ... smooth?` (marker after spPr, smooth after val); pie series has `explosion?` (after spPr, no shape/invertIfNegative applies differently); scatter series replaces cat/val with `xVal? yVal? smooth?`. idx and order are the only REQUIRED children of every series type. Word requires idx==order==list position for sane behavior; emit both = series position.

### 3.2 Bar (and column) — minimal plot group + axes

Content model refs: python-pptx bar-chart analysis (working minimal specimen validated against PowerPoint's chart engine, same engine as Word): https://python-pptx.readthedocs.io/en/latest/dev/analysis/cht-bar-chart.html

```xml
<c:barChart>
  <c:barDir val="col"/>              <!-- "col" = vertical bars, "bar" = horizontal; REQUIRED -->
  <c:grouping val="clustered"/>      <!-- clustered | stacked | percentStacked -->
  <c:varyColors val="0"/>
  <!-- c:ser* (3.1 block) -->
  <c:gapWidth val="150"/>            <!-- Word default 150; python-pptx uses 300 -->
  <!-- stacked variants additionally need <c:overlap val="100"/> AFTER gapWidth -->
  <c:axId val="111111111"/>
  <c:axId val="222222222"/>
</c:barChart>
<c:catAx>
  <c:axId val="111111111"/>
  <c:scaling><c:orientation val="minMax"/></c:scaling>
  <c:delete val="0"/>
  <c:axPos val="b"/>
  <c:crossAx val="222222222"/>
</c:catAx>
<c:valAx>
  <c:axId val="222222222"/>
  <c:scaling><c:orientation val="minMax"/></c:scaling>
  <c:delete val="0"/>
  <c:axPos val="l"/>
  <c:crossAx val="111111111"/>
</c:valAx>
```

Axis child order (CT_CatAx / CT_ValAx): `axId! scaling! delete? axPos! [gridlines/title/numFmt/tickmarks/txPr...]? crossAx! crosses? ...` — axId, scaling, axPos, crossAx are the required four; `delete val="0"` should always be emitted (absent `delete` has version-dependent visibility behavior). `scaling` may be empty per schema but Word writes `<c:orientation val="minMax"/>`; emit it.

### 3.3 Line — minimal plot group

CT_LineChart children in order (SDK LineChart page): `grouping! varyColors? ser* dLbls? dropLines? hiLowLines? upDownBars? marker? smooth? axId axId extLst?` — grouping REQUIRED, exactly two axId (cat + val, same axes block as bar).

```xml
<c:lineChart>
  <c:grouping val="standard"/>       <!-- standard | stacked | percentStacked -->
  <c:varyColors val="0"/>
  <!-- c:ser* — line series may add <c:marker><c:symbol val="none"/></c:marker> after spPr,
       and <c:smooth val="0"/> after c:val -->
  <c:marker val="1"/>                <!-- chart-level: show markers -->
  <c:axId val="111111111"/>
  <c:axId val="222222222"/>
</c:lineChart>
<!-- same catAx + valAx pair as bar -->
```

Emit `<c:smooth val="0"/>` per series (prevents Word defaulting curves on) — last child before extLst.

### 3.4 Pie — minimal plot group, NO axes

CT_PieChart children (SDK PieChart page): `varyColors? ser* dLbls? firstSliceAng? extLst?` — nothing required except the element itself; no axId children exist in the model.

```xml
<c:pieChart>
  <c:varyColors val="1"/>            <!-- 1 = one color per slice; Word-native for pie -->
  <!-- exactly ONE c:ser (3.1 block; multiple series in a pieChart are legal but only the first renders) -->
  <c:firstSliceAng val="0"/>
</c:pieChart>
```

REFUSE multi-series data for pie (only series 0 would render — conservative-refusal case, not a silent drop).

### 3.5 Scatter — minimal plot group + TWO value axes

python-pptx XY analysis (validated specimen): https://python-pptx.readthedocs.io/en/latest/dev/analysis/cht-xy-chart.html — CT_ScatterChart: `scatterStyle! varyColors? ser* dLbls? axId axId` — scatterStyle REQUIRED (`none|line|lineMarker|marker|smooth|smoothMarker`), two axId, and both axes are `c:valAx`.

```xml
<c:scatterChart>
  <c:scatterStyle val="lineMarker"/>
  <c:varyColors val="0"/>
  <c:ser>
    <c:idx val="0"/>
    <c:order val="0"/>
    <c:tx><!-- strRef as 3.1 --></c:tx>
    <!-- markers-only: suppress the connecting line with
         <c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr> -->
    <c:xVal>
      <c:numRef>
        <c:f>Sheet1!$A$2:$A$4</c:f>
        <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
          <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
        </c:numCache>
      </c:numRef>
    </c:xVal>
    <c:yVal>
      <c:numRef>
        <c:f>Sheet1!$B$2:$B$4</c:f>
        <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
          <c:pt idx="0"><c:v>4.3</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt><c:pt idx="2"><c:v>3.5</c:v></c:pt>
        </c:numCache>
      </c:numRef>
    </c:yVal>
    <c:smooth val="0"/>
  </c:ser>
  <c:axId val="111111111"/>
  <c:axId val="222222222"/>
</c:scatterChart>
<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>
  <c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:valAx>
<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>
  <c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
```

Scatter x-values are numeric (`numRef`); non-numeric X input is a refusal (offer line chart instead in the error message).

### 3.6 "What Word tolerates missing" table

| Element / part | Missing => | Verdict for us |
|---|---|---|
| Embedded xlsx + externalData (both absent) | Renders; Edit Data dead | Tolerated, but we EMIT both |
| externalData rel present, xlsx part absent | Package integrity error, repair risk | NEVER produce |
| numCache/strCache under numRef/strRef | Word tries the workbook; with autoUpdate 0 and no open, some versions render blank until refresh | Always emit caches — the caches are what render |
| c:f (using Lit instead of Ref) | Renders; Edit Data disconnected from data (PHPWord #956) | Don't use Lit |
| c:formatCode in numCache | Fine (treated as General) | Emit "General" anyway |
| c:layout in plotArea | Fine (auto layout) | Emit empty `<c:layout/>` |
| c:legend, c:title | Fine (no legend / auto or no title) | Legend on by default (param), title only if given |
| c:plotVisOnly, c:dispBlanksAs | Fine (defaults) | Emit (Word-native) |
| c:varyColors, c:gapWidth, c:overlap, c:firstSliceAng, c:marker, c:smooth | Fine (engine defaults) | Emit for determinism |
| c:catAx/c:valAx pair for bar/line/scatter, or axId mismatch with plot group | REPAIR / unreadable content | Mandatory, template-fixed |
| axId on pieChart | Schema violation (element not in model) => repair risk | Never |
| c:idx / c:order in ser | Schema-required => repair risk | Mandatory |
| Wrong child ORDER anywhere in the c: tree | Repair / silently dropped chart (docx4j/pptx4j reports) | Fixed templates only |
| ptCount disagreeing with pt list | Undefined; some versions render partially | Always exact |
| style1.xml/colors1.xml (MS chart-style parts) | Fine — fallback formatting from theme + c:style (Q4/Q7) | Don't emit |
| wp:docPr duplicate id | Word repairs/renumbers, selection bugs | Reuse `_next_docpr_id` |
| Override content-type entry for chart1.xml | REPAIR (part unreadable) | Mandatory |

---

## Q4. Theme colors vs hardcoded

- Word-native charts get their series colors from the chart color style part (`colors1.xml`, cycling `<a:schemeClr val="accent1"/>` … `accent6` with `lumMod/lumOff` variations) and shape formatting from `style1.xml` — both MICROSOFT EXTENSION parts, not ECMA (Q7). When those parts are absent, the chart engine falls back to the built-in default chart style, which ALSO cycles the document theme's accent1–accent6. Net effect: a chart with no explicit `c:spPr` fills at all follows the document theme automatically. Sources: https://www.brandwares.com/bestpractices/2021/08/ooxml-hacking-chart-template-colors/ ; https://learn.microsoft.com/en-us/answers/questions/289550/colors-for-chart ; https://learn.microsoft.com/en-gb/answers/questions/324319/how-can-i-set-the-fallback-chart-format-and-then-f (fallback mapping note: `c14:style val="102"` maps to `c:style val="2"` when style/colors parts are absent).
- RECOMMENDATION: emit NO per-series `c:spPr` fills by default — that is the maximally theme-following choice and the minimal-XML choice at once. Optionally emit `<c:style val="2"/>` (chartSpace child before c:chart, value 1–48) to pin the default style family; safe but not required — SKIP it for v1.6 (fewer knobs).
- If a `colors` parameter is offered (hex list), emit per-series `<c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>`; document that hardcoded colors do NOT retheme. For a `theme_colors` option, `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>` per series is valid chart XML and retheme-safe. Default remains: nothing.

---

## Q5. update_chart_data for existing charts

### Where the data lives (all copies)

1. `c:strCache`/`c:numCache` (and any `c:strLit`/`c:numLit`) under each series' `c:tx`/`c:cat`/`c:val`/`c:xVal`/`c:yVal` in the chart part — what renders.
2. The `c:f` range formulas — must span the new point count (e.g. 5 categories => `Sheet1!$A$2:$A$6`).
3. The embedded workbook's `sheetData` — what Edit Data shows and what refresh reloads.
4. `ptCount` in every touched cache.

All four in sync or the chart lies to somebody (Q2 stale table). This is exactly Eric White's ChartUpdater contract ("updates cached data AND the embedded XLSX"): http://www.ericwhite.com/blog/update-cached-data-and-embedded-xlsx-for-charts-in-docx-pptx/

### Strategy (recommended)

- Identify charts: enumerate document.xml `//a:graphicData[@uri = chart-ns]/c:chart/@r:id`, resolve through document.xml.rels to part names; expose `list_charts` returning index, part name, detected plot-group type(s), series names, point counts, size.
- For update: parse the chart part with lxml (`pkg.tree("word/charts/chart1.xml")` works today — DocxPackage is part-name agnostic; `mark_dirty(name)` and save already handle any XML part).
- EDIT THE CACHES IN PLACE, do not regenerate the chart part. Word-authored charts are full of formatting, dLbls, extLst/c14/c16 content we must not touch (Q7). In-place cache surgery: for each series, replace children of numCache/strCache (ptCount + pts), rewrite the `c:f` text, leave everything else byte-identical. lxml keeps sibling order — order risk is zero if we only replace cache children and text.
- Regenerate the embedded workbook WHOLE (find it via the chart part's rels, rel type `.../package`; overwrite the part bytes with a fresh Q2 minimal workbook containing the new data). Rationale: parsing arbitrary Word-written xlsx (styles, sharedStrings, dimension refs) to patch cells is a corruption surface; a full clean regenerate is deterministic. The workbook's only job is Edit Data, and a minimal sheet fulfils it. If the existing chart has NO embedded workbook rel (e.g. PHPWord-made), update caches only and report `"embedded_workbook": "none"` in the result.
- Data-shape rule (v1.6): update accepts new values for the EXISTING series/category structure or a new structure with the same series count; changing series COUNT on an existing Word chart means cloning/deleting `c:ser` trees with their formatting — defer, refuse with a clear message ("delete and re-add the chart to change series count").
- `c:tx` series names: update the tx strCache + B1-row cell when the caller supplies names.

### Refusal list for update_chart_data

- Plot group not in {barChart, lineChart, pieChart, scatterChart} — refuse: bar3DChart/line3DChart/pie3DChart, area/area3D, doughnut, radar, surface/surface3D, stock, bubble, ofPieChart. (Recognize-then-refuse with the found type named.)
- Combo charts: more than one plot-group element in plotArea — refuse (ambiguous series mapping).
- Series data held by `c:extLst`-based modern references only (chartex / `cx:` charts — see Q7: a `cx:` chart is a DIFFERENT part type with relationship `.../2014/relationships/chartEx`; our enumerator simply won't match it, but list_charts should detect and label it "chartex (unsupported)").
- Point-count mismatch between supplied series (ragged rows) unless the type is scatter-with-shared-X? — no: v1.6 requires rectangular data everywhere; refuse ragged.
- Scatter chart + non-numeric X.
- `c:multiLvlStrRef` in c:cat (multi-level categories) — refuse (structure change we can't faithfully cache-patch).
- Chart part not well-formed / cache nodes absent where expected — refuse rather than synthesize.
- Pie with >1 series (add path) / update pie with >1 supplied series.

---

## Q6. Sizing and positioning

- EMU constants already in `media.py`: `EMU_PER_INCH = 914400`, `EMU_PER_PT = 12700`. Word's default new-chart size is 6.0 x 3.94 in ≈ 5486400 x 3600000 EMU (15.24 x 10 cm); python-pptx/PowerPoint default and common practice ~5486400 x 3200400 (6.0 x 3.5 in). RECOMMEND default `width_pt=432` (6.0 in), `height_pt=252` (3.5 in), both parameterizable; cap width at 6.5 in like images.
- `wp:extent` cx/cy is the whole sizing story for inline charts (no aspect lock needed — charts reflow). Emit the same `distT/B/L/R="0"` quartet as images.
- docPr id: reuse `_next_docpr_id(pkg)` (media.py lines 61–66; scans all `wp:docPr` ids in document.xml, max+1). `name="Chart N"`. NOTE for builder: extract this helper to a shared module (e.g. `ops/_drawing.py`) rather than importing ops.media from ops.charts, or simply `from .media import _next_docpr_id, EMU_PER_PT` — house precedent exists for cross-ops imports (media.py imports from .text).
- Positioning: mirror add_image's anchor set exactly — `after_index`, `after_anchor`, `at_end`, `alignment` via `w:jc` on the wrapping paragraph, insert before `w:sectPr` at document end.

---

## Q7. Known landmines

1. **c14/c16 extension namespaces.** Word-authored charts carry `mc:Ignorable="c14 c16 c16r3"` plus `c14:style`, `c16:uniqueId` per series inside `c:extLst`, namespaces `http://schemas.microsoft.com/office/drawing/2007/8/2/chart` (c14), `.../2014/chart` (c16), `.../2017/03/chart` (c16r3). They are ignorable by design. RULES: (a) never emit them ourselves — then we need no `mc:` declaration at all; (b) in update_chart_data NEVER strip or reorder existing extLst/c14/c16 content — in-place cache surgery preserves it automatically; (c) `c16:uniqueId` values are opaque; do not clone a `c:ser` (would duplicate uniqueIds) — consistent with the refuse-series-count-change rule. Sources: wordarticles OOXML internals http://www.wordarticles.com/Articles/Formats/OOXML/OOXML.php ; https://learn.microsoft.com/en-gb/answers/questions/324319/how-can-i-set-the-fallback-chart-format-and-then-f
2. **Chart style parts.** `style1.xml` (`application/vnd.ms-office.chartstyle+xml`, rel `http://schemas.microsoft.com/office/2011/relationships/chartStyle`) and `colors1.xml` (`application/vnd.ms-office.chartcolorstyle+xml`, rel `.../2011/relationships/chartColorStyle`), hanging off the CHART part's rels. MS-only, optional; absence = built-in default style (Q4). We don't create them; update must not disturb them. Note the reverse quirk: Excel/Word can DROP style/colors parts of very old charts on re-save ( https://learn.microsoft.com/en-us/answers/questions/2286194/style1-xml-and-color1-xml-files-not-recognized-by ) — not our problem, but don't be surprised in round-trip tests.
3. **externalData r:id integrity.** The single most dangerous state is a dangling r:id (rel without part, or externalData without rel). Emission order in code: write xlsx part -> write chart rels -> write externalData element -> register content types -> hook into document. The existing save validator (`_validate_payload`) checks ZIP+XML well-formedness only; EXTEND the chart op with its own invariant check before `pkg.save()`: every `r:id` used in chart parts resolves in that part's rels, and every rel target exists in the package.
4. **Chartex (`cx:`) charts.** Modern Word chart types (treemap, sunburst, waterfall, funnel...) live in a different part (`word/charts/chartEx1.xml`, root `cx:chartSpace`, ns `http://schemas.microsoft.com/office/drawing/2014/chartex`, rel `http://schemas.microsoft.com/office/drawing/2014/relationships/chartEx`). Not ECMA c: charts at all. list_charts must label them; update refuses.
5. **Word round-trip.** When the user opens the doc and activates/edits our chart, Word rewrites chart1.xml Word-style: adds mc:Ignorable + c14/c16 blocks, may add style1/colors1 parts, renumbers rels, replaces our minimal workbook with a fuller one (styles.xml, theme, dimension). This is EXPECTED and lossless for data. Consequence for tests: round-trip assertions must compare semantics (series values), not bytes. Consequence for update: after a Word round-trip our own chart looks like a Word-authored chart — the in-place cache-surgery path (Q5) already handles that, so create-then-update-after-Word-edit works.
6. **`&` and text escaping in c:v / a:t** — lxml handles; never string-splice user text into templates (parse template, set `.text`). (Okapi found real-world `&amp;` corruption in chart.xml pipelines: https://gitlab.com/okapiframework/okapi/-/issues/526 )
7. **1900/1904 dates, c:date1904** — only relevant if we ever cache date axes; v1.6 treats categories as strings, values as floats. Date-typed input: pass through as its string form (documented).
8. **Decimal formatting in c:v** — write floats with `repr`-style minimal form, never locale-dependent formatting; integers without trailing `.0` (cosmetic, Word accepts both).
9. **ZIP ordering** — `DocxPackage` preserves `_order` and appends new parts at the end; Office does not care about member order (only that `[Content_Types].xml` is present); no action needed.

---

## Recommended implementation plan

### Files

- `src/word_mcp/ops/charts.py` — all chart ops:
  - `add_chart(pkg, chart_type, data, *, title=None, width_pt=None, height_pt=None, after_index=None, after_anchor=None, at_end=False, alignment="center", legend=True, colors=None) -> dict`
  - `list_charts(pkg) -> list[dict]`
  - `update_chart_data(pkg, chart_index, data, *, series_names=None) -> dict`
  - Private: `_build_chart_xml(chart_type, parsed) -> bytes`, `_build_workbook(parsed) -> bytes` (stdlib zipfile + templates), `_parse_chart_data(data) -> ParsedData` (normalizes the three input shapes), `_next_axid()`, cache-surgery helpers.
- Optionally split `_chart_xlsx.py` if charts.py crosses ~600 lines; builder's call.
- `integration/` registration snippet per house discipline; `@mcp.tool` wrappers in `server.py` following the `add_image`/`import_table` pattern (`_edit(file_path, lambda pkg: ..., backup=backup)`).
- Tests: `tests/test_charts.py` — build each type, reopen with zipfile+lxml and assert: content-type override present, rels resolve, chartSpace child order, cache counts; update tests including update-after-simulated-Word-markup (a fixture chart1.xml WITH c14/c16 extLst to prove preservation); refusal tests (radar fixture, combo fixture, chartex fixture, ragged data, pie multi-series, scatter string X). Validation with `com_validate_opens_clean` on Windows CI/manual pass (real-Word gate).

### Tool signatures (MCP layer)

```python
@mcp.tool
def add_chart(
    file_path: str,
    chart_type: str,              # "bar" | "column" | "line" | "pie" | "scatter"
    data: list | dict | str,      # rows [["","S1"],["Alpha",4.3],...] like import_table,
                                  # or {"categories":[...], "series":[{"name":..,"values":[...]},...]}
                                  # (scatter: {"series":[{"name":..,"x":[...],"y":[...]}]}),
                                  # or a .csv/.json path (reuse dataio's loaders)
    title: str | None = None,
    width_pt: float | None = None,     # default 432 (6.0in)
    height_pt: float | None = None,    # default 252 (3.5in)
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    alignment: str = "center",
    legend: bool = True,
    colors: list | None = None,        # optional hex list; default = theme-following (no spPr)
    backup: bool = True,
) -> dict: ...
    # returns {"chart_added": "word/charts/chart1.xml", "chart_index": N,
    #          "type": ..., "series": n, "points": m, "embedded_workbook": "word/embeddings/..."}

@mcp.tool
def list_charts(file_path: str) -> list: ...
    # [{"index":0,"part":"word/charts/chart1.xml","type":"barChart","series":["S1",...],
    #   "points":4,"width_pt":...,"height_pt":...,"embedded_workbook":bool,
    #   "supported_for_update":bool,"reason":None|"radar chart"|...}]

@mcp.tool
def update_chart_data(
    file_path: str,
    chart_index: int,
    data: list | dict | str,      # same shapes as add_chart
    series_names: list | None = None,
    backup: bool = True,
) -> dict: ...
    # returns {"updated": index, "type":..., "series": n, "points_before": a, "points_after": b,
    #          "embedded_workbook": "regenerated" | "none"}
```

`chart_type="bar"` => `barDir val="bar"` (horizontal); `"column"` => `barDir val="col"`. Both map to barChart.

### Build-order inside add_chart (integrity-safe sequence)

1. Parse/validate data (all refusals fire here, before any package mutation).
2. Write xlsx bytes -> `pkg.set_raw_part("word/embeddings/Microsoft_Excel_WorksheetN.xlsx", ...)`.
3. Write chart1.xml bytes (template-built, lxml-serialized) -> `set_raw_part("word/charts/chartN.xml", ...)`.
4. Write `word/charts/_rels/chartN.xml.rels` (rId1 -> package/embeddings).
5. `[Content_Types].xml`: Override for `/word/charts/chartN.xml`; Default `xlsx` (add-if-absent); Default `rels` exists already.
6. document.xml.rels: new rId -> chart part (reuse media.py's next-rId scan).
7. Insert the `w:p`/`w:drawing` block (docPr id via `_next_docpr_id`).
8. Self-check rels/parts closure, then normal `pkg.save()` (atomic + validated + backup).

### Refusal list (add_chart)

- chart_type outside the five aliases.
- Empty data, ragged rows, non-numeric values where numbers required, scatter non-numeric X.
- Pie with more than one series.
- More than 6? — no: any series count allowed for bar/line/scatter (colors cycle per Q4); no cap.
- `colors` list shorter than series count (refuse, name both counts, house style).
- Positioning conflicts (more than one of after_index/after_anchor/at_end) — same rule as add_image/add_equation.

### Dependency recommendation

**No new dependencies.** Hand-written 5-part xlsx (Q2) via stdlib zipfile; chart XML via lxml templates. openpyxl rejected: heavyweight for a fixed shell, and its chart writer emits cache-less charts useless as a reference for Word parts. If a future feature needs to READ arbitrary workbooks (true xlsx import), revisit then.

---

## Sources

- ECMA-376 Fundamentals, Chart Part: https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Chart_topic_ID0ELZLM.html
- CT_ChartSpace content model: http://www.datypic.com/sc/ooxml/e-draw-chart_chartSpace.html
- python-pptx chart analyses (validated minimal specimens): https://python-pptx.readthedocs.io/en/latest/dev/analysis/cht-bar-chart.html , https://python-pptx.readthedocs.io/en/latest/dev/analysis/cht-xy-chart.html , https://python-pptx.readthedocs.io/en/latest/dev/analysis/cht-access-xlsx.html , https://python-pptx.readthedocs.io/en/latest/dev/analysis/cht-plot-data.html
- OpenXML SDK schema-ordered child lists: PieChart https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.piechart , LineChart https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.linechart , BarChartSeries https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.barchartseries , ExternalData https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.externaldata?view=openxml-2.8.1
- Cache/xlsx sync: Eric White http://www.ericwhite.com/blog/update-cached-data-and-embedded-xlsx-for-charts-in-docx-pptx/ ; embeddedPackagePart thread http://www.ericwhite.com/blog/forums/topic/adding-workbook-to-chart-as-an-embeddedpackagepart/
- Render-without-workbook + Edit Data failure: https://docxtemplater.com/modules/chart/ ; https://github.com/PHPOffice/PHPWord/issues/956 ; https://botched-deployments.com/posts/python-docx-charts
- document.xml hookup: http://officeopenxml.com/drwPicInline.php ; https://github.com/python-openxml/python-docx/pull/392/commits/6ac9b73d67c3f2d133550c3f86a9697a605c41d6 ; https://ooxml.org/drawingml/a-chart/
- Theme colors / chart style parts: https://www.brandwares.com/bestpractices/2021/08/ooxml-hacking-chart-template-colors/ ; https://learn.microsoft.com/en-us/answers/questions/289550/colors-for-chart ; https://learn.microsoft.com/en-gb/answers/questions/324319/how-can-i-set-the-fallback-chart-format-and-then-f ; https://www.brandwares.com/bestpractices/2016/01/xml-hacking-linked-excel-charts/ ; https://learn.microsoft.com/en-us/answers/questions/2286194/style1-xml-and-color1-xml-files-not-recognized-by
- c14/c16/mc: http://www.wordarticles.com/Articles/Formats/OOXML/OOXML.php ; https://learn.microsoft.com/en-us/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml
- Order/corruption reports: https://www.docx4java.org/forums/pptx-java-f14/need-help-to-create-chart-in-ppt-slide-t1645.html ; https://www.docx4java.org/forums/docx-java-f6/add-chart-t872.html ; https://gitlab.com/okapiframework/okapi/-/issues/526
- Minimal xlsx: https://www.brendanlong.com/the-minimum-viable-xlsx-reader.html ; https://social.technet.microsoft.com/wiki/contents/articles/19601.powershell-generate-real-excel-xlsx-files-without-excel.aspx ; https://professor-excel.com/xml-zip-excel-file-structure/
- python-docx edit-chart-data feature request (state of the ecosystem): https://github.com/python-openxml/python-docx/issues/1141
