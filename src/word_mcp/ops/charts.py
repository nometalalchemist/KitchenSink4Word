"""Native charts: add, list, update. ECMA `c:chartSpace` parts plus an
embedded minimal workbook, built from fixed templates only.

Design rules (from the v1.6 charts research, 2026-08-28):
- Child order in every chart CT_* is xsd:sequence-fixed and Word's consumer
  enforces it; all XML here is emitted in template order, never call order.
- The literal caches (numCache/strCache) are what render; the embedded xlsx
  is what Edit Data and refresh use. The two are ALWAYS written together --
  an externalData r:id without its target part is a repair-prompt trigger.
- No per-series spPr by default so the document theme's accent cycle applies.
- update_chart_data performs in-place cache surgery (preserving any
  c14/c16 extLst content) and regenerates the embedded workbook whole.
- Anything outside the supported shapes is refused with a typed error,
  never guessed at.
"""

from __future__ import annotations

import io
import json
import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree

from ..core.errors import (
    TargetNotFound,
    UnsupportedStructure,
    ValidationFailed,
    WordMcpError,
)
from ..core.package import DocxPackage, qn
from ..core.sandbox import check_path
from .media import EMU_PER_INCH, EMU_PER_PT, _next_docpr_id

_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"

_CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
_XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_CHART_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)
_PACKAGE_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
)

# Fixed axis ids: arbitrary 32-bit ints, unique within one chart part (each
# chart lives in its own part, so constants are safe).
_CAT_AX_ID = "111111111"
_VAL_AX_ID = "222222222"

# Every plot-group element CT_PlotArea can hold (for combo detection and for
# naming the found type in refusals).
_ALL_PLOT_GROUPS = {
    "areaChart", "area3DChart", "lineChart", "line3DChart", "stockChart",
    "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart",
    "barChart", "bar3DChart", "ofPieChart", "surfaceChart", "surface3DChart",
    "bubbleChart",
}
_SUPPORTED_PLOT_GROUPS = {"barChart", "lineChart", "pieChart", "scatterChart"}

_CHART_TYPES = {"bar", "column", "line", "pie", "scatter"}

_MAX_ROWS = 100000  # sanity ceiling; a document chart never needs more


def _qc(name: str) -> str:
    return f"{{{_C}}}{name}"


def _c(parent: etree._Element, name: str, val=None) -> etree._Element:
    """Append a c: child in call order (callers are the fixed templates)."""
    el = etree.SubElement(parent, _qc(name))
    if val is not None:
        el.set("val", str(val))
    return el


def _fmt_num(x: float) -> str:
    """Minimal locale-free decimal form; integers without trailing .0."""
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return repr(x)


def _col_letter(idx0: int) -> str:
    """0-based column index -> spreadsheet letters (0 -> A, 26 -> AA)."""
    letters = ""
    n = idx0 + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


# ------------------------------------------------------------- data parsing


def _num(v, where: str, *, scatter_x: bool = False) -> float:
    if isinstance(v, bool) or v is None:
        raise WordMcpError(f"non-numeric value {v!r} {where}")
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip())
        except ValueError:
            pass
    msg = f"non-numeric value {v!r} {where}"
    if scatter_x:
        msg += (
            "; scatter x-values must be numeric -- for text categories "
            "use a line chart instead"
        )
    raise WordMcpError(msg)


def _series_display_name(name, i: int) -> str:
    text = "" if name is None else str(name).strip()
    return text or f"Series {i + 1}"


def _load_payload(data):
    """Normalize a file path (.csv/.json) or inline payload into either a
    dict (categories/series form) or a list of rows."""
    if isinstance(data, str):
        check_path(data, "read chart data file")
        p = Path(data)
        if not p.exists():
            raise TargetNotFound(f"data file not found: {data}")
        suffix = p.suffix.lower()
        if suffix == ".csv":
            from .dataio import _load_rows

            return _load_rows(data)
        if suffix == ".json":
            loaded = json.loads(p.read_text(encoding="utf-8-sig"))
            if isinstance(loaded, dict) and "rows" in loaded:
                return loaded["rows"]
            return loaded
        raise WordMcpError(
            f"data file must be .csv or .json, got {suffix or 'no extension'}"
        )
    return data


def _parse_rows(rows, xy: bool) -> dict:
    if (
        not isinstance(rows, list)
        or len(rows) < 2
        or not all(isinstance(r, (list, tuple)) for r in rows)
    ):
        raise WordMcpError(
            "row data needs a header row plus at least one data row, e.g. "
            '[["", "Series 1"], ["Alpha", 4.3], ["Beta", 2.5]]'
        )
    widths = {len(r) for r in rows}
    if len(widths) != 1:
        raise WordMcpError(
            f"ragged data refused: row lengths vary ({sorted(widths)}); "
            "chart data must be rectangular"
        )
    width = widths.pop()
    if width < 2:
        raise WordMcpError(
            "row data needs a label/x column plus at least one series column"
        )
    header, body = rows[0], rows[1:]
    names = [_series_display_name(h, i) for i, h in enumerate(header[1:])]
    if xy:
        xs = [
            _num(r[0], f"at row {i + 1}, column 1 (x)", scatter_x=True)
            for i, r in enumerate(body)
        ]
        series = [
            {
                "name": names[s],
                "x": list(xs),
                "y": [
                    _num(r[s + 1], f"at row {i + 1}, column {s + 2}")
                    for i, r in enumerate(body)
                ],
            }
            for s in range(width - 1)
        ]
        return {"kind": "xy", "categories": None, "series": series}
    categories = ["" if r[0] is None else str(r[0]) for r in body]
    series = [
        {
            "name": names[s],
            "values": [
                _num(r[s + 1], f"at row {i + 1}, column {s + 2}")
                for i, r in enumerate(body)
            ],
        }
        for s in range(width - 1)
    ]
    return {"kind": "category", "categories": categories, "series": series}


def _parse_dict(payload: dict, xy: bool) -> dict:
    raw_series = payload.get("series")
    if not isinstance(raw_series, list) or not raw_series:
        raise WordMcpError('data dict needs a non-empty "series" list')
    if xy:
        series = []
        for i, s in enumerate(raw_series):
            if not isinstance(s, dict) or "x" not in s or "y" not in s:
                raise WordMcpError(
                    f'scatter series {i} must be {{"name", "x", "y"}} '
                    "with parallel x/y value lists"
                )
            xs = [
                _num(v, f"in series {i} x[{j}]", scatter_x=True)
                for j, v in enumerate(s["x"])
            ]
            ys = [_num(v, f"in series {i} y[{j}]") for j, v in enumerate(s["y"])]
            if len(xs) != len(ys):
                raise WordMcpError(
                    f"series {i}: x has {len(xs)} points but y has "
                    f"{len(ys)}; x and y must be the same length"
                )
            if not xs:
                raise WordMcpError(f"series {i} has no data points")
            series.append(
                {"name": _series_display_name(s.get("name"), i), "x": xs, "y": ys}
            )
        counts = {len(s["x"]) for s in series}
        if len(counts) != 1:
            raise WordMcpError(
                f"ragged data refused: series point counts differ "
                f"({sorted(counts)}); all series must have the same length"
            )
        return {"kind": "xy", "categories": None, "series": series}
    categories = payload.get("categories")
    if not isinstance(categories, list) or not categories:
        raise WordMcpError('data dict needs a non-empty "categories" list')
    categories = ["" if v is None else str(v) for v in categories]
    series = []
    for i, s in enumerate(raw_series):
        if not isinstance(s, dict) or "values" not in s:
            raise WordMcpError(
                f'series {i} must be {{"name", "values"}}; for scatter '
                'charts use {"name", "x", "y"}'
            )
        values = [
            _num(v, f"in series {i} values[{j}]")
            for j, v in enumerate(s["values"])
        ]
        if len(values) != len(categories):
            raise WordMcpError(
                f"ragged data refused: series {i} has {len(values)} values "
                f"but there are {len(categories)} categories"
            )
        series.append(
            {"name": _series_display_name(s.get("name"), i), "values": values}
        )
    return {"kind": "category", "categories": categories, "series": series}


def _parse_chart_data(data, chart_type: str) -> dict:
    """Normalize all accepted input shapes; every data refusal fires here,
    before any package mutation."""
    xy = chart_type == "scatter"
    payload = _load_payload(data)
    if isinstance(payload, dict):
        parsed = _parse_dict(payload, xy)
    elif isinstance(payload, list):
        parsed = _parse_rows(payload, xy)
    else:
        raise WordMcpError(
            "data must be a dict (categories/series), a list of rows, or a "
            ".csv/.json file path"
        )
    npoints = (
        len(parsed["series"][0]["x"]) if xy else len(parsed["categories"])
    )
    if npoints > _MAX_ROWS:
        raise WordMcpError(f"{npoints} data points exceeds the {_MAX_ROWS} cap")
    if chart_type == "pie" and len(parsed["series"]) > 1:
        raise WordMcpError(
            f"pie charts show exactly one series; got "
            f"{len(parsed['series'])}. Word would silently render only the "
            "first -- pass a single series, or use a bar chart"
        )
    return parsed


def _point_count(parsed: dict) -> int:
    if parsed["kind"] == "xy":
        return len(parsed["series"][0]["x"])
    return len(parsed["categories"])


# ------------------------------------------------- embedded workbook builder

_XLSX_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType='
    '"application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType='
    '"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType='
    '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    "</Types>"
)

_XLSX_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type='
    '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="xl/workbook.xml"/>'
    "</Relationships>"
)

_XLSX_WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
    "</workbook>"
)

_XLSX_WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type='
    '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
    ' Target="worksheets/sheet1.xml"/>'
    "</Relationships>"
)

_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _sheet_cells(parsed: dict) -> dict[int, list[tuple[str, object]]]:
    """row number (1-based) -> [(cell ref, value)] in Word's chart layout.
    str values become inline strings, floats become numbers."""
    rows: dict[int, list[tuple[str, object]]] = {}

    def put(row: int, col0: int, value) -> None:
        rows.setdefault(row, []).append((f"{_col_letter(col0)}{row}", value))

    if parsed["kind"] == "xy":
        # Series i occupies its own column pair (x, y); name over the y col.
        for i, s in enumerate(parsed["series"]):
            xcol, ycol = 2 * i, 2 * i + 1
            put(1, ycol, s["name"])
            for j, (x, y) in enumerate(zip(s["x"], s["y"])):
                put(j + 2, xcol, x)
                put(j + 2, ycol, y)
    else:
        # A1 blank, categories in column A from A2, one series per column.
        for i, s in enumerate(parsed["series"]):
            put(1, i + 1, s["name"])
        for j, cat in enumerate(parsed["categories"]):
            put(j + 2, 0, cat)
            for i, s in enumerate(parsed["series"]):
                put(j + 2, i + 1, s["values"][j])
    return rows


def _build_worksheet_xml(parsed: dict) -> bytes:
    ws = etree.Element(f"{{{_SML_NS}}}worksheet", nsmap={None: _SML_NS})
    sheet_data = etree.SubElement(ws, f"{{{_SML_NS}}}sheetData")
    cells = _sheet_cells(parsed)
    for row_num in sorted(cells):
        row_el = etree.SubElement(sheet_data, f"{{{_SML_NS}}}row")
        row_el.set("r", str(row_num))
        for ref, value in cells[row_num]:
            cell = etree.SubElement(row_el, f"{{{_SML_NS}}}c")
            cell.set("r", ref)
            if isinstance(value, str):
                cell.set("t", "inlineStr")
                is_el = etree.SubElement(cell, f"{{{_SML_NS}}}is")
                etree.SubElement(is_el, f"{{{_SML_NS}}}t").text = value
            else:
                etree.SubElement(cell, f"{{{_SML_NS}}}v").text = _fmt_num(value)
    return etree.tostring(
        ws, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _build_workbook(parsed: dict) -> bytes:
    """The 5-part minimal xlsx (research Q2): renders nothing itself, exists
    so right-click > Edit Data works and matches the caches exactly."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _XLSX_CONTENT_TYPES)
        zf.writestr("_rels/.rels", _XLSX_ROOT_RELS)
        zf.writestr("xl/workbook.xml", _XLSX_WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", _XLSX_WORKBOOK_RELS)
        zf.writestr("xl/worksheets/sheet1.xml", _build_worksheet_xml(parsed))
    return buf.getvalue()


# ----------------------------------------------------- chart XML construction


def _range_f(col0: int, first_row: int, last_row: int) -> str:
    letter = _col_letter(col0)
    if first_row == last_row:
        return f"Sheet1!${letter}${first_row}"
    return f"Sheet1!${letter}${first_row}:${letter}${last_row}"


def _write_str_ref(parent, wrap_name: str, f: str, values: list[str]) -> None:
    wrap = _c(parent, wrap_name)
    ref = _c(wrap, "strRef")
    _c(ref, "f").text = f
    cache = _c(ref, "strCache")
    _c(cache, "ptCount", len(values))
    for i, v in enumerate(values):
        pt = _c(cache, "pt")
        pt.set("idx", str(i))
        _c(pt, "v").text = v


def _write_num_ref(parent, wrap_name: str, f: str, values: list[float]) -> None:
    wrap = _c(parent, wrap_name)
    ref = _c(wrap, "numRef")
    _c(ref, "f").text = f
    cache = _c(ref, "numCache")
    _c(cache, "formatCode").text = "General"
    _c(cache, "ptCount", len(values))
    for i, v in enumerate(values):
        pt = _c(cache, "pt")
        pt.set("idx", str(i))
        _c(pt, "v").text = _fmt_num(v)


def _write_series_color(ser, color: str, line_style: bool) -> None:
    """Optional per-series spPr (only when the colors param is given).
    Line-family charts color the stroke; bar fills the shape."""
    sppr = _c(ser, "spPr")
    if line_style:
        ln = etree.SubElement(sppr, f"{{{_A}}}ln")
        fill = etree.SubElement(ln, f"{{{_A}}}solidFill")
    else:
        fill = etree.SubElement(sppr, f"{{{_A}}}solidFill")
    etree.SubElement(fill, f"{{{_A}}}srgbClr").set("val", color)


def _normalize_colors(colors, n_series: int, chart_type: str) -> list[str] | None:
    if colors is None:
        return None
    if chart_type == "pie":
        raise WordMcpError(
            "colors is not supported for pie charts (slices are colored per "
            "point, not per series); omit colors to follow the document theme"
        )
    if not isinstance(colors, list) or len(colors) < n_series:
        given = len(colors) if isinstance(colors, list) else 0
        raise WordMcpError(
            f"colors has {given} entries but the data has {n_series} "
            "series; supply one hex color per series or omit colors to "
            "follow the document theme"
        )
    out = []
    for c in colors[:n_series]:
        h = str(c).lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", h):
            raise WordMcpError(
                f"invalid color {c!r}; use 6-digit hex like '4472C4'"
            )
        out.append(h)
    return out


def _category_series(
    group,
    parsed: dict,
    chart_type: str,
    colors: list[str] | None,
) -> None:
    """Emit c:ser blocks for bar/column/line/pie in CT_*Ser child order."""
    n = len(parsed["categories"])
    for i, s in enumerate(parsed["series"]):
        ser = _c(group, "ser")
        _c(ser, "idx", i)
        _c(ser, "order", i)
        _write_str_ref(ser, "tx", _range_f(i + 1, 1, 1), [s["name"]])
        if colors:
            _write_series_color(ser, colors[i], line_style=chart_type == "line")
        _write_str_ref(ser, "cat", _range_f(0, 2, n + 1), parsed["categories"])
        _write_num_ref(ser, "val", _range_f(i + 1, 2, n + 1), s["values"])
        if chart_type == "line":
            _c(ser, "smooth", 0)


def _xy_series(group, parsed: dict, colors: list[str] | None) -> None:
    for i, s in enumerate(parsed["series"]):
        n = len(s["x"])
        xcol, ycol = 2 * i, 2 * i + 1
        ser = _c(group, "ser")
        _c(ser, "idx", i)
        _c(ser, "order", i)
        _write_str_ref(ser, "tx", _range_f(ycol, 1, 1), [s["name"]])
        if colors:
            _write_series_color(ser, colors[i], line_style=True)
        _write_num_ref(ser, "xVal", _range_f(xcol, 2, n + 1), s["x"])
        _write_num_ref(ser, "yVal", _range_f(ycol, 2, n + 1), s["y"])
        _c(ser, "smooth", 0)


def _write_axis(plot_area, kind: str, ax_id: str, cross_id: str, pos: str) -> None:
    ax = _c(plot_area, kind)
    _c(ax, "axId", ax_id)
    scaling = _c(ax, "scaling")
    _c(scaling, "orientation", "minMax")
    _c(ax, "delete", 0)
    _c(ax, "axPos", pos)
    _c(ax, "crossAx", cross_id)


def _build_chart_xml(
    chart_type: str,
    parsed: dict,
    *,
    title: str | None,
    legend: bool,
    colors: list[str] | None,
) -> bytes:
    root = etree.Element(
        _qc("chartSpace"), nsmap={"c": _C, "a": _A, "r": _R_NS}
    )
    chart = _c(root, "chart")
    if title is not None:
        t = _c(chart, "title")
        tx = _c(t, "tx")
        rich = _c(tx, "rich")
        etree.SubElement(rich, f"{{{_A}}}bodyPr")
        etree.SubElement(rich, f"{{{_A}}}lstStyle")
        p = etree.SubElement(rich, f"{{{_A}}}p")
        r = etree.SubElement(p, f"{{{_A}}}r")
        etree.SubElement(r, f"{{{_A}}}t").text = title
        _c(t, "overlay", 0)
        _c(chart, "autoTitleDeleted", 0)
    plot_area = _c(chart, "plotArea")
    _c(plot_area, "layout")

    if chart_type in ("bar", "column"):
        group = _c(plot_area, "barChart")
        _c(group, "barDir", "bar" if chart_type == "bar" else "col")
        _c(group, "grouping", "clustered")
        _c(group, "varyColors", 0)
        _category_series(group, parsed, chart_type, colors)
        _c(group, "gapWidth", 150)
        _c(group, "axId", _CAT_AX_ID)
        _c(group, "axId", _VAL_AX_ID)
        _write_axis(plot_area, "catAx", _CAT_AX_ID, _VAL_AX_ID, "b")
        _write_axis(plot_area, "valAx", _VAL_AX_ID, _CAT_AX_ID, "l")
    elif chart_type == "line":
        group = _c(plot_area, "lineChart")
        _c(group, "grouping", "standard")
        _c(group, "varyColors", 0)
        _category_series(group, parsed, chart_type, colors)
        _c(group, "marker", 1)
        _c(group, "axId", _CAT_AX_ID)
        _c(group, "axId", _VAL_AX_ID)
        _write_axis(plot_area, "catAx", _CAT_AX_ID, _VAL_AX_ID, "b")
        _write_axis(plot_area, "valAx", _VAL_AX_ID, _CAT_AX_ID, "l")
    elif chart_type == "pie":
        group = _c(plot_area, "pieChart")
        _c(group, "varyColors", 1)
        _category_series(group, parsed, chart_type, colors)
        _c(group, "firstSliceAng", 0)
    else:  # scatter
        group = _c(plot_area, "scatterChart")
        _c(group, "scatterStyle", "lineMarker")
        _c(group, "varyColors", 0)
        _xy_series(group, parsed, colors)
        _c(group, "axId", _CAT_AX_ID)
        _c(group, "axId", _VAL_AX_ID)
        _write_axis(plot_area, "valAx", _CAT_AX_ID, _VAL_AX_ID, "b")
        _write_axis(plot_area, "valAx", _VAL_AX_ID, _CAT_AX_ID, "l")

    if legend:
        leg = _c(chart, "legend")
        _c(leg, "legendPos", "b")
        _c(leg, "overlay", 0)
    _c(chart, "plotVisOnly", 1)
    _c(chart, "dispBlanksAs", "gap")
    ext = _c(root, "externalData")
    ext.set(f"{{{_R_NS}}}id", "rId1")
    _c(ext, "autoUpdate", 0)
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


# --------------------------------------------------------- package plumbing


def _next_part_number(pkg: DocxPackage, pattern: str) -> int:
    """Next free N for a part-name pattern with one %d slot."""
    n = 1
    while pkg.has_part(pattern % n):
        n += 1
    return n


def _ensure_content_types(pkg: DocxPackage, chart_part: str) -> None:
    ct_root = pkg.root("[Content_Types].xml")
    part_name = "/" + chart_part
    if not any(
        o.get("PartName") == part_name
        for o in ct_root.findall(f"{{{_CT_NS}}}Override")
    ):
        override = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", _CHART_CT)
    if not any(
        d.get("Extension") == "xlsx"
        for d in ct_root.findall(f"{{{_CT_NS}}}Default")
    ):
        default = etree.SubElement(ct_root, f"{{{_CT_NS}}}Default")
        default.set("Extension", "xlsx")
        default.set("ContentType", _XLSX_CT)
    pkg.mark_dirty("[Content_Types].xml")


def _build_chart_rels(workbook_part: str) -> bytes:
    root = etree.Element(f"{{{_REL_NS}}}Relationships", nsmap={None: _REL_NS})
    rel = etree.SubElement(root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", "rId1")
    rel.set("Type", _PACKAGE_REL)
    rel.set("Target", "../embeddings/" + workbook_part.rsplit("/", 1)[1])
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _add_document_rel(pkg: DocxPackage, chart_part: str) -> str:
    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    existing = {r.get("Id") for r in rels_root}
    i = 1
    while f"rId{i}" in existing:
        i += 1
    rid = f"rId{i}"
    rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", rid)
    rel.set("Type", _CHART_REL)
    rel.set("Target", chart_part.split("word/", 1)[1])
    pkg.mark_dirty(rels_part)
    return rid


def _rels_part_for(part: str) -> str:
    folder, name = part.rsplit("/", 1)
    return f"{folder}/_rels/{name}.rels"


def _check_chart_closure(pkg: DocxPackage, chart_part: str) -> None:
    """Every r:id used in the chart part resolves in its rels, and every rel
    target exists in the package (the dangling-r:id repair trigger)."""
    rels_part = _rels_part_for(chart_part)
    rels = {}
    if pkg.has_part(rels_part):
        for rel in pkg.root(rels_part):
            rels[rel.get("Id")] = rel.get("Target")
    chart_root = etree.fromstring(pkg.raw_part(chart_part))
    used = set()
    for el in chart_root.iter():
        for attr, value in el.attrib.items():
            if attr.startswith(f"{{{_R_NS}}}"):
                used.add(value)
    missing = used - set(rels)
    if missing:
        raise ValidationFailed(
            f"{chart_part} references undefined relationship id(s) "
            f"{sorted(missing)}; document not saved"
        )
    base = chart_part.rsplit("/", 1)[0]
    for rid, target in rels.items():
        resolved = posixpath.normpath(posixpath.join(base, target))
        if not pkg.has_part(resolved):
            raise ValidationFailed(
                f"{rels_part} {rid} targets missing part {resolved}; "
                "document not saved"
            )


def _insert_chart_paragraph(
    pkg: DocxPackage,
    rid: str,
    *,
    cx: int,
    cy: int,
    alignment: str,
    after_index: int | None,
    after_anchor: str | None,
    at_end: bool,
) -> int:
    docpr_id = _next_docpr_id(pkg)
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    jc = etree.SubElement(ppr, qn("w:jc"))
    jc.set(
        qn("w:val"),
        {"left": "left", "center": "center", "right": "right"}[alignment],
    )
    r = etree.SubElement(p, qn("w:r"))
    drawing = etree.SubElement(r, qn("w:drawing"))
    inline = etree.SubElement(drawing, f"{{{_WP}}}inline")
    for attr in ("distT", "distB", "distL", "distR"):
        inline.set(attr, "0")
    extent = etree.SubElement(inline, f"{{{_WP}}}extent")
    extent.set("cx", str(cx))
    extent.set("cy", str(cy))
    docpr = etree.SubElement(inline, f"{{{_WP}}}docPr")
    docpr.set("id", str(docpr_id))
    docpr.set("name", f"Chart {docpr_id}")
    graphic = etree.SubElement(inline, f"{{{_A}}}graphic")
    gdata = etree.SubElement(graphic, f"{{{_A}}}graphicData")
    gdata.set("uri", _C)
    chart_el = etree.SubElement(
        gdata, _qc("chart"), nsmap={"c": _C, "r": _R_NS}
    )
    chart_el.set(f"{{{_R_NS}}}id", rid)

    from .text import _body_paragraph, _resolve_anchor

    body = pkg.body()
    if at_end or (after_index is None and after_anchor is None):
        sectpr = body.find(qn("w:sectPr"))
        if sectpr is not None:
            sectpr.addprevious(p)
        else:
            body.append(p)
    elif after_anchor is not None:
        _resolve_anchor(pkg, after_anchor).addnext(p)
    else:
        _body_paragraph(pkg, after_index).addnext(p)
    pkg.mark_dirty()
    return docpr_id


# ------------------------------------------------------------------ add_chart


def add_chart(
    pkg: DocxPackage,
    chart_type: str,
    data,
    *,
    title: str | None = None,
    width_pt: float | None = None,
    height_pt: float | None = None,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    alignment: str = "center",
    legend: bool = True,
    colors: list | None = None,
) -> dict:
    """Insert a native Word chart built from JSON/CSV data. Emits the chart
    part with full literal caches AND a matching embedded workbook, so the
    chart both renders and supports right-click Edit Data."""
    if chart_type not in _CHART_TYPES:
        raise WordMcpError(
            f"unsupported chart_type {chart_type!r}; use one of "
            f"{sorted(_CHART_TYPES)} (3D, area, doughnut, radar, and other "
            "types are deliberately not generated)"
        )
    positioners = sum(
        x is not None for x in (after_index, after_anchor)
    ) + bool(at_end)
    if positioners > 1:
        raise WordMcpError(
            "specify at most one of after_index, after_anchor, at_end "
            "(default is document end)"
        )
    if alignment not in ("left", "center", "right"):
        raise WordMcpError(
            f"alignment must be left, center, or right, got {alignment!r}"
        )
    if width_pt is not None and width_pt <= 0:
        raise WordMcpError("width_pt must be positive")
    if height_pt is not None and height_pt <= 0:
        raise WordMcpError("height_pt must be positive")

    parsed = _parse_chart_data(data, chart_type)
    color_list = _normalize_colors(colors, len(parsed["series"]), chart_type)

    # Integrity-safe build order (research Q7.3): workbook part first, then
    # chart part, then its rels, then content types, then the document hook.
    chart_n = _next_part_number(pkg, "word/charts/chart%d.xml")
    chart_part = f"word/charts/chart{chart_n}.xml"
    wb_n = _next_part_number(
        pkg, "word/embeddings/Microsoft_Excel_Worksheet%d.xlsx"
    )
    workbook_part = f"word/embeddings/Microsoft_Excel_Worksheet{wb_n}.xlsx"

    pkg.set_raw_part(workbook_part, _build_workbook(parsed))
    pkg.set_raw_part(
        chart_part,
        _build_chart_xml(
            chart_type, parsed, title=title, legend=legend, colors=color_list
        ),
    )
    pkg.set_raw_part(_rels_part_for(chart_part), _build_chart_rels(workbook_part))
    _ensure_content_types(pkg, chart_part)
    rid = _add_document_rel(pkg, chart_part)

    cx = int((width_pt if width_pt is not None else 432.0) * EMU_PER_PT)
    cy = int((height_pt if height_pt is not None else 252.0) * EMU_PER_PT)
    _insert_chart_paragraph(
        pkg,
        rid,
        cx=cx,
        cy=cy,
        alignment=alignment,
        after_index=after_index,
        after_anchor=after_anchor,
        at_end=at_end,
    )
    _check_chart_closure(pkg, chart_part)

    index = next(
        (
            e["index"]
            for e in _document_charts(pkg)
            if e.get("part") == chart_part
        ),
        None,
    )
    return {
        "chart_added": chart_part,
        "chart_index": index,
        "type": chart_type,
        "series": len(parsed["series"]),
        "points": _point_count(parsed),
        "embedded_workbook": workbook_part,
        "width_pt": round(cx / EMU_PER_PT, 1),
        "height_pt": round(cy / EMU_PER_PT, 1),
    }


# ---------------------------------------------------------------- enumeration


def _document_charts(pkg: DocxPackage) -> list[dict]:
    """Charts hanging off document.xml, in document order. Each entry:
    index, kind ('chart' | 'chartex'), part (resolved or None), extent."""
    rels_part = "word/_rels/document.xml.rels"
    rid_target = {}
    if pkg.has_part(rels_part):
        rid_target = {
            r.get("Id"): r.get("Target") for r in pkg.root(rels_part)
        }

    def resolve(rid: str | None) -> str | None:
        target = rid_target.get(rid)
        if not target:
            return None
        return posixpath.normpath(posixpath.join("word", target.lstrip("/")))

    out = []
    for gdata in pkg.root().iter(f"{{{_A}}}graphicData"):
        uri = gdata.get("uri")
        if uri == _C:
            chart_el = gdata.find(_qc("chart"))
            kind = "chart"
        elif uri == _CX:
            chart_el = gdata.find(f"{{{_CX}}}chart")
            kind = "chartex"
        else:
            continue
        rid = (
            chart_el.get(f"{{{_R_NS}}}id") if chart_el is not None else None
        )
        entry: dict = {
            "index": len(out),
            "kind": kind,
            "part": resolve(rid),
        }
        container = gdata.getparent()
        while container is not None and etree.QName(
            container
        ).localname not in ("inline", "anchor"):
            container = container.getparent()
        if container is not None:
            extent = container.find(f"{{{_WP}}}extent")
            if extent is not None:
                entry["width_pt"] = round(int(extent.get("cx")) / EMU_PER_PT, 1)
                entry["height_pt"] = round(int(extent.get("cy")) / EMU_PER_PT, 1)
        out.append(entry)
    return out


def _plot_groups(chart_root) -> list:
    plot_area = chart_root.find(_qc("chart") + "/" + _qc("plotArea"))
    if plot_area is None:
        return []
    return [
        el
        for el in plot_area
        if etree.QName(el).localname in _ALL_PLOT_GROUPS
    ]


def _series_name(ser) -> str | None:
    tx = ser.find(_qc("tx"))
    if tx is None:
        return None
    v = tx.find(_qc("v"))
    if v is not None:
        return v.text or ""
    pts = tx.findall(_qc("strRef") + "/" + _qc("strCache") + "/" + _qc("pt"))
    if pts:
        v = pts[0].find(_qc("v"))
        return (v.text or "") if v is not None else None
    return None


def _series_point_count(ser) -> int | None:
    for wrap_name in ("val", "yVal"):
        wrap = ser.find(_qc(wrap_name))
        if wrap is None:
            continue
        for holder in ("numRef/numCache", "numLit"):
            path = "/".join(_qc(part) for part in holder.split("/"))
            data_el = wrap.find(path)
            if data_el is not None:
                ptc = data_el.find(_qc("ptCount"))
                if ptc is not None:
                    return int(ptc.get("val", "0"))
                return len(data_el.findall(_qc("pt")))
    return None


def _chart_title(chart_root) -> str | None:
    title = chart_root.find(_qc("chart") + "/" + _qc("title"))
    if title is None:
        return None
    text = "".join(t.text or "" for t in title.iter(f"{{{_A}}}t"))
    return text or None


def _has_embedded_workbook(pkg: DocxPackage, chart_part: str) -> bool:
    rels_part = _rels_part_for(chart_part)
    if not pkg.has_part(rels_part):
        return False
    base = chart_part.rsplit("/", 1)[0]
    for rel in pkg.root(rels_part):
        if rel.get("Type") == _PACKAGE_REL:
            target = posixpath.normpath(posixpath.join(base, rel.get("Target")))
            if pkg.has_part(target):
                return True
    return False


def _update_supportability(chart_root) -> tuple[bool, str | None]:
    """(supported_for_update, reason) mirroring update_chart_data's refusals."""
    groups = _plot_groups(chart_root)
    if not groups:
        return False, "no plot group found in the chart part"
    if len(groups) > 1:
        names = [etree.QName(g).localname for g in groups]
        return False, f"combo chart with multiple plot groups ({names})"
    name = etree.QName(groups[0]).localname
    if name not in _SUPPORTED_PLOT_GROUPS:
        return False, f"unsupported chart type ({name})"
    for ser in groups[0].findall(_qc("ser")):
        cat = ser.find(_qc("cat"))
        if cat is not None and cat.find(_qc("multiLvlStrRef")) is not None:
            return False, "multi-level categories (multiLvlStrRef)"
    return True, None


def list_charts(pkg: DocxPackage) -> list[dict]:
    """Enumerate every chart in the document body with type, title, series
    names, size, and whether update_chart_data can act on it."""
    out = []
    for entry in _document_charts(pkg):
        info: dict = {
            "index": entry["index"],
            "part": entry["part"],
            "width_pt": entry.get("width_pt"),
            "height_pt": entry.get("height_pt"),
        }
        if entry["kind"] == "chartex":
            info.update(
                {
                    "type": "chartex",
                    "title": None,
                    "series": [],
                    "points": None,
                    "embedded_workbook": False,
                    "supported_for_update": False,
                    "reason": (
                        "modern chart (chartex cx: format, e.g. treemap/"
                        "sunburst/waterfall) -- not an ECMA c: chart; "
                        "unsupported"
                    ),
                }
            )
            out.append(info)
            continue
        part = entry["part"]
        if part is None or not pkg.has_part(part):
            info.update(
                {
                    "type": None,
                    "title": None,
                    "series": [],
                    "points": None,
                    "embedded_workbook": False,
                    "supported_for_update": False,
                    "reason": "chart relationship target missing",
                }
            )
            out.append(info)
            continue
        chart_root = pkg.root(part)
        groups = _plot_groups(chart_root)
        supported, reason = _update_supportability(chart_root)
        sers = groups[0].findall(_qc("ser")) if len(groups) == 1 else []
        info.update(
            {
                "type": "+".join(etree.QName(g).localname for g in groups)
                or None,
                "title": _chart_title(chart_root),
                "series": [
                    _series_name(s) or f"Series {i + 1}"
                    for i, s in enumerate(sers)
                ],
                "points": _series_point_count(sers[0]) if sers else None,
                "embedded_workbook": _has_embedded_workbook(pkg, part),
                "supported_for_update": supported,
                "reason": reason,
            }
        )
        out.append(info)
    return out


# ---------------------------------------------------------- update_chart_data


def _rebuild_cache(cache, values, *, numeric: bool) -> None:
    """Replace a cache/lit element's point data in place. formatCode text and
    any extLst children survive; everything else is rewritten exactly."""
    fmt_text = None
    fmt = cache.find(_qc("formatCode"))
    if fmt is not None:
        fmt_text = fmt.text
    exts = cache.findall(_qc("extLst"))
    for child in list(cache):
        cache.remove(child)
    if numeric:
        _c(cache, "formatCode").text = fmt_text or "General"
    _c(cache, "ptCount", len(values))
    for i, v in enumerate(values):
        pt = _c(cache, "pt")
        pt.set("idx", str(i))
        _c(pt, "v").text = _fmt_num(v) if numeric else v
    for ext in exts:
        cache.append(ext)


def _update_data_node(
    ser,
    wrap_name: str,
    values,
    f: str,
    *,
    numeric: bool,
    label: str,
) -> None:
    """In-place surgery on one c:cat/c:val/c:xVal/c:yVal/c:tx data node."""
    wrap = ser.find(_qc(wrap_name))
    if wrap is None:
        raise UnsupportedStructure(
            f"chart series has no c:{wrap_name} node; cannot update {label} "
            "-- delete and re-add the chart instead"
        )
    ref_name, cache_name = (
        ("numRef", "numCache") if numeric else ("strRef", "strCache")
    )
    other_ref = wrap.find(_qc("strRef" if numeric else "numRef"))
    if other_ref is not None:
        stored = "text" if numeric else "numbers"
        raise UnsupportedStructure(
            f"chart stores {label} as {stored} but the supplied data is "
            f"{'numeric' if numeric else 'text'}; refusing a type change -- "
            "delete and re-add the chart instead"
        )
    ref = wrap.find(_qc(ref_name))
    if ref is not None:
        f_el = ref.find(_qc("f"))
        if f_el is None:
            raise UnsupportedStructure(
                f"chart {label} reference has no c:f formula; part is not a "
                "shape this tool can update safely"
            )
        f_el.text = f
        cache = ref.find(_qc(cache_name))
        if cache is None:
            raise UnsupportedStructure(
                f"chart {label} reference has no {cache_name}; refusing to "
                "synthesize missing cache structure"
            )
        _rebuild_cache(cache, values, numeric=numeric)
        return
    lit = wrap.find(_qc("numLit" if numeric else "strLit"))
    if lit is not None:
        _rebuild_cache(lit, values, numeric=numeric)
        return
    v_el = wrap.find(_qc("v")) if wrap_name == "tx" else None
    if v_el is not None:
        v_el.text = values[0]
        return
    raise UnsupportedStructure(
        f"chart {label} holds no recognized data structure "
        f"({ref_name}/{'numLit' if numeric else 'strLit'}); refusing to guess"
    )


def update_chart_data(
    pkg: DocxPackage,
    chart_index: int,
    data,
    *,
    series_names: list | None = None,
) -> dict:
    """Replace an existing chart's data in place: cache surgery on the chart
    part (c14/c16 extension content preserved) plus full regeneration of the
    embedded workbook so Edit Data matches what renders."""
    charts = _document_charts(pkg)
    if not 0 <= chart_index < len(charts):
        raise TargetNotFound(
            f"chart index {chart_index} out of range "
            f"({len(charts)} chart(s) in the document)"
        )
    entry = charts[chart_index]
    if entry["kind"] == "chartex":
        raise UnsupportedStructure(
            "this is a modern chart (chartex cx: format, e.g. treemap/"
            "sunburst/waterfall), not an ECMA chart; updating it is not "
            "supported"
        )
    part = entry["part"]
    if part is None or not pkg.has_part(part):
        raise UnsupportedStructure(
            "the chart's relationship target part is missing from the "
            "package; the document needs repair before its charts can be "
            "edited"
        )
    chart_root = pkg.root(part)
    supported, reason = _update_supportability(chart_root)
    if not supported:
        raise UnsupportedStructure(f"cannot update this chart: {reason}")
    group = _plot_groups(chart_root)[0]
    group_name = etree.QName(group).localname

    chart_type = {
        "barChart": "bar",
        "lineChart": "line",
        "pieChart": "pie",
        "scatterChart": "scatter",
    }[group_name]
    parsed = _parse_chart_data(data, chart_type)

    sers = group.findall(_qc("ser"))
    if len(sers) != len(parsed["series"]):
        raise UnsupportedStructure(
            f"the chart has {len(sers)} series but the data has "
            f"{len(parsed['series'])}; changing series count on an existing "
            "chart is not supported -- delete and re-add the chart instead"
        )
    if series_names is not None and len(series_names) != len(sers):
        raise WordMcpError(
            f"series_names has {len(series_names)} entries but the chart "
            f"has {len(sers)} series"
        )

    points_before = _series_point_count(sers[0])
    n = _point_count(parsed)

    for i, (ser, new) in enumerate(zip(sers, parsed["series"])):
        if series_names is not None:
            name = _series_display_name(series_names[i], i)
            new["name"] = name
            ycol = 2 * i + 1 if parsed["kind"] == "xy" else i + 1
            _update_data_node(
                ser,
                "tx",
                [name],
                _range_f(ycol, 1, 1),
                numeric=False,
                label=f"series {i} name",
            )
        if parsed["kind"] == "xy":
            _update_data_node(
                ser,
                "xVal",
                new["x"],
                _range_f(2 * i, 2, n + 1),
                numeric=True,
                label=f"series {i} x-values",
            )
            _update_data_node(
                ser,
                "yVal",
                new["y"],
                _range_f(2 * i + 1, 2, n + 1),
                numeric=True,
                label=f"series {i} y-values",
            )
        else:
            _update_data_node(
                ser,
                "cat",
                parsed["categories"],
                _range_f(0, 2, n + 1),
                numeric=False,
                label="categories",
            )
            _update_data_node(
                ser,
                "val",
                new["values"],
                _range_f(i + 1, 2, n + 1),
                numeric=True,
                label=f"series {i} values",
            )
    pkg.mark_dirty(part)

    # Regenerate the embedded workbook whole (never patch Word-written xlsx).
    workbook_state = "none"
    rels_part = _rels_part_for(part)
    if pkg.has_part(rels_part):
        base = part.rsplit("/", 1)[0]
        for rel in pkg.root(rels_part):
            if rel.get("Type") == _PACKAGE_REL:
                target = posixpath.normpath(
                    posixpath.join(base, rel.get("Target"))
                )
                if pkg.has_part(target):
                    pkg.set_raw_part(target, _build_workbook(parsed))
                    workbook_state = "regenerated"
                break
    _check_chart_closure(pkg, part)
    return {
        "updated": chart_index,
        "type": chart_type,
        "series": len(parsed["series"]),
        "points_before": points_before,
        "points_after": n,
        "embedded_workbook": workbook_state,
    }
