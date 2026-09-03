"""Table operations: full CRUD including the ecosystem gaps — column
insert/delete on merged tables, bulk cell writes, merge/unmerge.

Coordinate systems (documented on every tool):
- ROW index: position of w:tr in the table, 0-based.
- CELL index: position of a w:tc within its row (what get_table shows), 0-based.
- GRID column: column of the underlying w:tblGrid, 0-based. Merged cells span
  several grid columns. Column operations use GRID coordinates; cell text
  operations use CELL coordinates.
"""

from __future__ import annotations

import copy

from lxml import etree

from ..core.errors import TargetNotFound, UnsupportedStructure, WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap
from .read import body_items


# ------------------------------------------------------------------ grid model


class CellSpan:
    __slots__ = ("tc", "grid_start", "grid_end", "vmerge")

    def __init__(self, tc, grid_start, grid_end, vmerge):
        self.tc = tc
        self.grid_start = grid_start  # inclusive
        self.grid_end = grid_end  # exclusive
        self.vmerge = vmerge  # None | 'restart' | 'continue'


def _find_table(pkg: DocxPackage, index: int) -> etree._Element:
    for kind, idx, el in body_items(pkg):
        if kind == "table" and idx == index:
            return el
    raise TargetNotFound(f"no body-level table with index {index}")


def _grid_cols(tbl: etree._Element) -> list[etree._Element]:
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        raise UnsupportedStructure("table has no w:tblGrid; cannot address columns")
    return grid.findall(qn("w:gridCol"))


def _row_zones(tr: etree._Element) -> tuple[int, int]:
    """(gridBefore, gridAfter) skip counts for a row (0 when absent)."""
    trpr = tr.find(qn("w:trPr"))
    if trpr is None:
        return 0, 0
    before = trpr.find(qn("w:gridBefore"))
    after = trpr.find(qn("w:gridAfter"))
    return (
        int(before.get(qn("w:val"), "0")) if before is not None else 0,
        int(after.get(qn("w:val"), "0")) if after is not None else 0,
    )


def _set_row_zone(
    tbl: etree._Element, tr: etree._Element, which: str, count: int
) -> None:
    """Set gridBefore/gridAfter (+ matching wBefore/wAfter width) on a row,
    removing the elements when count reaches 0. Keeps CT_TrPr order (grid
    elements first, then widths)."""
    trpr = tr.find(qn("w:trPr"))
    if trpr is None:
        trpr = etree.Element(qn("w:trPr"))
        tr.insert(0, trpr)
    grid_tag = qn(f"w:grid{which.capitalize()}")
    w_tag = qn(f"w:w{which.capitalize()}")
    grid_el = trpr.find(grid_tag)
    w_el = trpr.find(w_tag)
    if count <= 0:
        if grid_el is not None:
            trpr.remove(grid_el)
        if w_el is not None:
            trpr.remove(w_el)
        return
    if grid_el is None:
        grid_el = etree.Element(grid_tag)
        trpr.insert(0, grid_el)
    grid_el.set(qn("w:val"), str(count))
    cols = _grid_cols(tbl)
    n = len(cols)
    if which == "before":
        width = sum(int(cols[i].get(qn("w:w"), "0")) for i in range(min(count, n)))
    else:
        width = sum(
            int(cols[i].get(qn("w:w"), "0")) for i in range(max(n - count, 0), n)
        )
    if w_el is None:
        w_el = etree.Element(w_tag)
        grid_el.addnext(w_el)
    w_el.set(qn("w:w"), str(width))
    w_el.set(qn("w:type"), "dxa")


def _row_spans(tr: etree._Element) -> list[CellSpan]:
    before, _after = _row_zones(tr)
    spans = []
    pos = before
    for tc in tr.findall(qn("w:tc")):
        tcpr = tc.find(qn("w:tcPr"))
        span = 1
        vmerge = None
        if tcpr is not None:
            gs = tcpr.find(qn("w:gridSpan"))
            if gs is not None:
                span = int(gs.get(qn("w:val"), "1"))
            vm = tcpr.find(qn("w:vMerge"))
            if vm is not None:
                vmerge = vm.get(qn("w:val"), "continue")
        spans.append(CellSpan(tc, pos, pos + span, vmerge))
        pos += span
    return spans


def _table_model(tbl: etree._Element) -> tuple[list[etree._Element], list[list[CellSpan]]]:
    rows = tbl.findall(qn("w:tr"))
    model = [_row_spans(tr) for tr in rows]
    n_grid = len(_grid_cols(tbl))
    for r_i, (tr, row) in enumerate(zip(rows, model)):
        _before, after = _row_zones(tr)
        width = (row[-1].grid_end if row else _before) + after
        if width != n_grid:
            raise UnsupportedStructure(
                f"row {r_i} covers {width} grid columns (incl. skip zones) but "
                f"the grid defines {n_grid}; refusing to operate on an "
                "inconsistent table"
            )
    return rows, model


# ------------------------------------------------------------- cell utilities


def _cell_set_text(tc: etree._Element, text: str) -> None:
    """Replace a cell's content with `text` ('\n' = paragraph break), keeping
    the first paragraph's properties and first run's formatting."""
    from .text import _check_storable_text

    _check_storable_text(text, "cell text")
    paras = tc.findall(qn("w:p"))
    template_ppr = None
    template_rpr = None
    if paras:
        ppr = paras[0].find(qn("w:pPr"))
        if ppr is not None:
            template_ppr = copy.deepcopy(ppr)
        first_r = paras[0].find(qn("w:r"))
        if first_r is not None:
            rpr = first_r.find(qn("w:rPr"))
            if rpr is not None:
                template_rpr = copy.deepcopy(rpr)
    for p in paras:
        tc.remove(p)
    for i, line in enumerate(text.split("\n")):
        p = etree.SubElement(tc, qn("w:p"))
        if template_ppr is not None:
            p.append(copy.deepcopy(template_ppr))
        run = etree.SubElement(p, qn("w:r"))
        if template_rpr is not None:
            run.append(copy.deepcopy(template_rpr))
        if line:
            t = etree.SubElement(run, qn("w:t"))
            t.text = line
            _runmap._preserve_space(t)


def _cell_set_text_tracked(pkg, tc: etree._Element, text: str, author: str) -> None:
    """Tracked cell rewrite: existing runs wrapped in w:del, new text appended
    as w:ins in the first paragraph."""
    from . import _tracked

    root = pkg.root()
    date = _tracked.now_iso()
    first_p = tc.find(qn("w:p"))
    if first_p is None:
        first_p = etree.SubElement(tc, qn("w:p"))
    for p in tc.findall(qn("w:p")):
        _tracked.wrap_paragraph_content_deleted(
            p, author=author, rev_id=_tracked.next_rev_id(root), date=date
        )
    if text:
        ins = _tracked.make_ins(
            text=text,
            author=author,
            rev_id=_tracked.next_rev_id(root),
            date=date,
        )
        first_p.append(ins)


def _cell_text(tc: etree._Element) -> str:
    from .read import paragraph_text

    return "\n".join(paragraph_text(p) for p in tc.findall(qn("w:p")))


def _tcpr(tc: etree._Element) -> etree._Element:
    tcpr = tc.find(qn("w:tcPr"))
    if tcpr is None:
        tcpr = etree.Element(qn("w:tcPr"))
        tc.insert(0, tcpr)
    return tcpr


def _set_grid_span(tc: etree._Element, span: int) -> None:
    tcpr = _tcpr(tc)
    gs = tcpr.find(qn("w:gridSpan"))
    if span <= 1:
        if gs is not None:
            tcpr.remove(gs)
    else:
        if gs is None:
            gs = etree.SubElement(tcpr, qn("w:gridSpan"))
        gs.set(qn("w:val"), str(span))


def _set_tc_width(tc: etree._Element, twips: int) -> None:
    tcpr = _tcpr(tc)
    tcw = tcpr.find(qn("w:tcW"))
    if tcw is None:
        tcw = etree.SubElement(tcpr, qn("w:tcW"))
    tcw.set(qn("w:w"), str(twips))
    tcw.set(qn("w:type"), "dxa")


def _new_cell(width_twips: int | None = None) -> etree._Element:
    tc = etree.Element(qn("w:tc"))
    if width_twips:
        _set_tc_width(tc, width_twips)
    etree.SubElement(tc, qn("w:p"))
    return tc


def _span_width(tbl: etree._Element, span: CellSpan) -> int:
    cols = _grid_cols(tbl)
    return sum(
        int(cols[i].get(qn("w:w"), "0"))
        for i in range(span.grid_start, min(span.grid_end, len(cols)))
    )


# ------------------------------------------------------------------ column ops


def delete_columns(pkg: DocxPackage, table_index: int, columns: list[int]) -> dict:
    """Delete GRID columns (0-based). Merged cells shrink; single cells vanish.
    Deleting every column is refused."""
    tbl = _find_table(pkg, table_index)
    cols = _grid_cols(tbl)
    n = len(cols)
    targets = sorted(set(columns), reverse=True)
    if not targets:
        raise WordMcpError("no columns given")
    for c in targets:
        if not 0 <= c < n:
            raise TargetNotFound(f"grid column {c} out of range (table has {n})")
    if len(targets) >= n:
        raise WordMcpError(
            "refusing to delete every column; use delete_table instead"
        )

    for col in targets:
        rows, model = _table_model(tbl)
        n_now = len(_grid_cols(tbl))
        for tr, row in zip(rows, model):
            before, after = _row_zones(tr)
            if col < before:
                _set_row_zone(tbl, tr, "before", before - 1)
                continue
            if col >= n_now - after:
                _set_row_zone(tbl, tr, "after", after - 1)
                continue
            for span in row:
                if span.grid_start <= col < span.grid_end:
                    width = span.grid_end - span.grid_start
                    if width > 1:
                        _set_grid_span(span.tc, width - 1)
                        _set_tc_width(
                            span.tc,
                            max(_span_width(tbl, span) - _col_width(tbl, col), 0),
                        )
                    else:
                        span.tc.getparent().remove(span.tc)
                    break
        grid = tbl.find(qn("w:tblGrid"))
        grid.remove(_grid_cols(tbl)[col])
        # Widths of skip zones changed with the grid; refresh them.
        for tr in tbl.findall(qn("w:tr")):
            b, a = _row_zones(tr)
            if b:
                _set_row_zone(tbl, tr, "before", b)
            if a:
                _set_row_zone(tbl, tr, "after", a)
    pkg.mark_dirty()
    from .notes import purge_orphans

    result = {"deleted_columns": targets[::-1], "remaining": n - len(targets)}
    purged = purge_orphans(pkg)["purged"]
    if purged:
        result["note_definitions_purged"] = purged
    return result


def _col_width(tbl: etree._Element, col: int) -> int:
    return int(_grid_cols(tbl)[col].get(qn("w:w"), "0"))


def insert_columns(
    pkg: DocxPackage,
    table_index: int,
    *,
    at: int,
    count: int = 1,
    width_pt: float | None = None,
) -> dict:
    """Insert `count` GRID columns before grid position `at` (at = column
    count appends at the right edge). Insertion inside a horizontally merged
    cell widens that cell instead of splitting it."""
    tbl = _find_table(pkg, table_index)
    cols = _grid_cols(tbl)
    n = len(cols)
    if not 0 <= at <= n:
        raise TargetNotFound(f"insert position {at} out of range 0..{n}")
    if count < 1:
        raise WordMcpError("count must be >= 1")
    if width_pt is not None and width_pt <= 0:
        raise WordMcpError("width_pt must be positive")
    width_twips = (
        int(width_pt * 20)
        if width_pt
        else (sum(_col_width(tbl, i) for i in range(n)) // n if n else 1440)
    )

    for _ in range(count):
        rows, model = _table_model(tbl)
        n_now = len(_grid_cols(tbl))
        for tr, row_spans in zip(rows, model):
            before, after = _row_zones(tr)
            if at < before or (at == before and before > 0 and not row_spans):
                _set_row_zone(tbl, tr, "before", before + 1)
                continue
            if at > n_now - after or (after > 0 and at == n_now):
                _set_row_zone(tbl, tr, "after", after + 1)
                continue
            placed = False
            for span in row_spans:
                if span.grid_start == at:
                    span.tc.addprevious(_new_cell(width_twips))
                    placed = True
                    break
                if span.grid_start < at < span.grid_end:
                    _set_grid_span(
                        span.tc, (span.grid_end - span.grid_start) + 1
                    )
                    _set_tc_width(
                        span.tc, _span_width(tbl, span) + width_twips
                    )
                    placed = True
                    break
            if not placed:
                # Append at the right edge (after the last cell element).
                if row_spans:
                    row_spans[-1].tc.addnext(_new_cell(width_twips))
        grid = tbl.find(qn("w:tblGrid"))
        new_gc = etree.Element(qn("w:gridCol"))
        new_gc.set(qn("w:w"), str(width_twips))
        existing = _grid_cols(tbl)
        if at >= len(existing):
            grid.append(new_gc)
        else:
            existing[at].addprevious(new_gc)
        for tr in tbl.findall(qn("w:tr")):
            b, a = _row_zones(tr)
            if b:
                _set_row_zone(tbl, tr, "before", b)
            if a:
                _set_row_zone(tbl, tr, "after", a)
    pkg.mark_dirty()
    result = {"inserted_at": at, "count": count}
    if count > _LARGE_INSERT_WARN:
        result["warning"] = (
            f"inserted {count} columns in one call; operations above "
            f"{_LARGE_INSERT_WARN} columns produce very large documents "
            "and slow every subsequent read and save"
        )
    return result


# --------------------------------------------------------------------- row ops


_LARGE_INSERT_WARN = 1000


def insert_rows(
    pkg: DocxPackage,
    table_index: int,
    *,
    at: int,
    count: int = 1,
    copy_format_from: int | None = None,
) -> dict:
    """Insert rows before row `at` (at = row count appends). Structure
    (cell count, gridSpans, widths, cell formatting) is copied from
    `copy_format_from` (default: the row at/above the insertion point).
    New rows normally stand alone (vertical-merge markers stripped), but a
    row inserted INSIDE an existing vertical merge joins it: its cells at
    the merged columns get vMerge continue, so the chain below is never
    orphaned (field test, 2026-09-03)."""
    tbl = _find_table(pkg, table_index)
    rows, model = _table_model(tbl)
    if not 0 <= at <= len(rows):
        raise TargetNotFound(f"row position {at} out of range 0..{len(rows)}")
    template_i = (
        copy_format_from
        if copy_format_from is not None
        else min(at, len(rows) - 1)
    )
    if not 0 <= template_i < len(rows):
        raise TargetNotFound(f"copy_format_from {template_i} out of range")
    template = rows[template_i]

    # Grid ranges where the insertion point sits INSIDE a vertical merge:
    # the row above belongs to a chain that the row at `at` continues.
    # New cells covering exactly those ranges must continue the chain.
    inside_merge: set[tuple[int, int]] = set()
    if 0 < at < len(rows):
        above = {
            (s.grid_start, s.grid_end): s.vmerge for s in model[at - 1]
        }
        for s in model[at]:
            if s.vmerge == "continue" and above.get(
                (s.grid_start, s.grid_end)
            ) in ("restart", "continue"):
                inside_merge.add((s.grid_start, s.grid_end))

    template_spans = model[template_i]
    cells_continued = 0
    new_rows = []
    for _ in range(count):
        tr = copy.deepcopy(template)
        # Strip row-level height/header repeat? Keep them: format copy.
        for tc, span in zip(tr.findall(qn("w:tc")), template_spans):
            tcpr = tc.find(qn("w:tcPr"))
            if tcpr is not None:
                vm = tcpr.find(qn("w:vMerge"))
                if vm is not None:
                    tcpr.remove(vm)
            # Empty the content but keep paragraph formatting of first para.
            _cell_set_text(tc, "")
            if (span.grid_start, span.grid_end) in inside_merge:
                vm = etree.SubElement(_tcpr(tc), qn("w:vMerge"))
                vm.set(qn("w:val"), "continue")
                cells_continued += 1
        new_rows.append(tr)

    if at >= len(rows):
        for tr in new_rows:
            tbl.append(tr)
    else:
        for tr in reversed(new_rows):
            rows[at].addprevious(tr)
    pkg.mark_dirty()
    result = {"inserted_rows_at": at, "count": count}
    if cells_continued:
        result["cells_continued_vertical_merges"] = cells_continued
        result["note"] = (
            "the insertion point lies inside a vertical merge; the new "
            "rows' cells at the merged columns joined the merge (vMerge "
            "continue) so the chain below stays intact"
        )
    if count > _LARGE_INSERT_WARN:
        result["warning"] = (
            f"inserted {count} rows in one call; operations above "
            f"{_LARGE_INSERT_WARN} rows produce very large documents and "
            "slow every subsequent read and save"
        )
    return result


def delete_rows(pkg: DocxPackage, table_index: int, start: int, end: int | None = None) -> dict:
    """Delete rows [start, end] inclusive. Vertical merges that begin in a
    deleted row are re-rooted onto the first surviving row of the chain, with
    the merged cell's content moved there."""
    tbl = _find_table(pkg, table_index)
    rows, model = _table_model(tbl)
    end = start if end is None else end
    if not (0 <= start <= end < len(rows)):
        raise TargetNotFound(
            f"row range {start}-{end} out of range (table has {len(rows)} rows)"
        )
    if end - start + 1 >= len(rows):
        raise WordMcpError("refusing to delete every row; use delete_table")

    # Re-root vertical merges whose restart lies in the deleted range.
    for r_i in range(start, end + 1):
        for span in model[r_i]:
            if span.vmerge == "restart":
                # Find the first surviving continuation below.
                for r_j in range(end + 1, len(rows)):
                    cont = next(
                        (
                            s
                            for s in model[r_j]
                            if s.grid_start == span.grid_start
                            and s.grid_end == span.grid_end
                            and s.vmerge == "continue"
                        ),
                        None,
                    )
                    if cont is None:
                        break
                    # Promote to restart and move content down.
                    vm = _tcpr(cont.tc).find(qn("w:vMerge"))
                    vm.set(qn("w:val"), "restart")
                    text = _cell_text(span.tc)
                    if text.strip():
                        _cell_set_text(cont.tc, text)
                    break

    for r_i in range(end, start - 1, -1):
        tbl.remove(rows[r_i])
    pkg.mark_dirty()
    from .notes import purge_orphans

    result = {"deleted_rows": list(range(start, end + 1))}
    purged = purge_orphans(pkg)["purged"]
    if purged:
        result["note_definitions_purged"] = purged
    return result


# -------------------------------------------------------------- bulk cell ops


def set_cells(
    pkg: DocxPackage,
    table_index: int,
    edits: list[dict],
    *,
    track: bool = False,
    author: str = "Claude",
) -> dict:
    """Bulk cell text write. Each edit: {row, cell, text} using ROW and CELL
    (not grid) coordinates as reported by get_table. One call, any number of
    cells. With track, old content is marked deleted and new content inserted,
    attributed to `author`."""
    tbl = _find_table(pkg, table_index)
    rows = tbl.findall(qn("w:tr"))
    applied = 0
    for edit in edits:
        r, c = edit["row"], edit["cell"]
        if not 0 <= r < len(rows):
            raise TargetNotFound(f"row {r} out of range (table has {len(rows)})")
        tcs = rows[r].findall(qn("w:tc"))
        if not 0 <= c < len(tcs):
            raise TargetNotFound(
                f"cell {c} out of range (row {r} has {len(tcs)} cells)"
            )
        tcpr = tcs[c].find(qn("w:tcPr"))
        vmerge = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
        if vmerge is not None and vmerge.get(qn("w:val"), "continue") == "continue":
            raise UnsupportedStructure(
                f"cell ({r},{c}) is a vertically merged CONTINUATION — text "
                "written there is invisible in Word; write to the restart "
                "cell at the top of the merge instead"
            )
        if track:
            _cell_set_text_tracked(pkg, tcs[c], str(edit["text"]), author)
        else:
            _cell_set_text(tcs[c], str(edit["text"]))
        applied += 1
    pkg.mark_dirty()
    from .notes import purge_orphans

    result = {"cells_written": applied}
    purged = purge_orphans(pkg)["purged"]
    if purged:
        result["note_definitions_purged"] = purged
    return result


def set_cells_grid(
    pkg: DocxPackage, table_index: int, *, origin_row: int, origin_cell: int, data: list[list]
) -> dict:
    """Write a 2D block starting at (origin_row, origin_cell), row-major."""
    edits = [
        {"row": origin_row + r, "cell": origin_cell + c, "text": val}
        for r, row in enumerate(data)
        for c, val in enumerate(row)
    ]
    return set_cells(pkg, table_index, edits)


_VALIGN = {"top": "top", "center": "center", "bottom": "bottom"}


def format_cells(
    pkg: DocxPackage,
    table_index: int,
    targets: list[dict],
    formatting: dict,
) -> dict:
    """Bulk cell formatting. targets: [{row, cell}] or [{row}] for whole rows.
    formatting keys: shading (hex fill), bold/italic (applied to all runs),
    alignment (left|center|right|justify), valign (top|center|bottom),
    padding_pt (uniform cell margin)."""
    from .text import _ALIGN, _check_keys

    _check_keys(
        formatting,
        {"shading", "bold", "italic", "alignment", "valign", "padding_pt",
         "text_direction"},
        "cell-formatting",
    )

    tbl = _find_table(pkg, table_index)
    rows = tbl.findall(qn("w:tr"))
    cells: list[etree._Element] = []
    for tgt in targets:
        r = tgt["row"]
        if not 0 <= r < len(rows):
            raise TargetNotFound(f"row {r} out of range")
        tcs = rows[r].findall(qn("w:tc"))
        if "cell" in tgt and tgt["cell"] is not None:
            if not 0 <= tgt["cell"] < len(tcs):
                raise TargetNotFound(f"cell {tgt['cell']} out of range in row {r}")
            cells.append(tcs[tgt["cell"]])
        else:
            cells.extend(tcs)

    for tc in cells:
        tcpr = _tcpr(tc)
        if "shading" in formatting:
            shd = tcpr.find(qn("w:shd"))
            if shd is None:
                shd = etree.SubElement(tcpr, qn("w:shd"))
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), formatting["shading"].lstrip("#"))
        if "valign" in formatting:
            if formatting["valign"] not in _VALIGN:
                raise WordMcpError(f"valign must be one of {list(_VALIGN)}")
            va = tcpr.find(qn("w:vAlign"))
            if va is None:
                va = etree.SubElement(tcpr, qn("w:vAlign"))
            va.set(qn("w:val"), _VALIGN[formatting["valign"]])
        if "text_direction" in formatting:
            directions = {"horizontal": None, "btLr": "btLr", "tbRl": "tbRl"}
            if formatting["text_direction"] not in directions:
                raise WordMcpError(
                    "text_direction must be horizontal | btLr | tbRl"
                )
            td = tcpr.find(qn("w:textDirection"))
            val = directions[formatting["text_direction"]]
            if val is None:
                if td is not None:
                    tcpr.remove(td)
            else:
                if td is None:
                    td = etree.SubElement(tcpr, qn("w:textDirection"))
                td.set(qn("w:val"), val)
        if "padding_pt" in formatting:
            mar = tcpr.find(qn("w:tcMar"))
            if mar is None:
                mar = etree.SubElement(tcpr, qn("w:tcMar"))
            tw = str(int(formatting["padding_pt"] * 20))
            for side in ("top", "left", "bottom", "right"):
                el = mar.find(qn(f"w:{side}"))
                if el is None:
                    el = etree.SubElement(mar, qn(f"w:{side}"))
                el.set(qn("w:w"), tw)
                el.set(qn("w:type"), "dxa")
        if "alignment" in formatting:
            val = _ALIGN.get(formatting["alignment"])
            if val is None:
                raise WordMcpError(f"alignment must be one of {list(_ALIGN)}")
            for p in tc.findall(qn("w:p")):
                ppr = p.find(qn("w:pPr"))
                if ppr is None:
                    ppr = etree.Element(qn("w:pPr"))
                    p.insert(0, ppr)
                jc = ppr.find(qn("w:jc"))
                if jc is None:
                    jc = etree.SubElement(ppr, qn("w:jc"))
                jc.set(qn("w:val"), val)
        for key in ("bold", "italic"):
            if key in formatting:
                tag = qn("w:b") if key == "bold" else qn("w:i")
                for r_el in tc.iter(qn("w:r")):
                    rpr = r_el.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = etree.Element(qn("w:rPr"))
                        r_el.insert(0, rpr)
                    existing = rpr.find(tag)
                    if formatting[key] and existing is None:
                        etree.SubElement(rpr, tag)
                    elif not formatting[key] and existing is not None:
                        rpr.remove(existing)
    pkg.mark_dirty()
    return {"cells_formatted": len(cells)}


# ------------------------------------------------------------------- merge ops


def merge_cells(
    pkg: DocxPackage,
    table_index: int,
    *,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> dict:
    """Merge the rectangle [start_row..end_row] x [start_col..end_col] (GRID
    coordinates, inclusive). The rectangle must align with existing cell
    boundaries in every row; content of absorbed cells is appended to the
    top-left cell's text."""
    tbl = _find_table(pkg, table_index)
    rows, model = _table_model(tbl)
    if end_row < start_row:
        raise WordMcpError(
            f"merge range is inverted: end_row {end_row} precedes "
            f"start_row {start_row}"
        )
    if not (0 <= start_row <= end_row < len(rows)):
        raise TargetNotFound(
            f"row range {start_row}..{end_row} out of bounds "
            f"(table has {len(rows)} row(s), 0-based)"
        )
    n_grid = len(_grid_cols(tbl))
    if end_col < start_col:
        raise WordMcpError(
            f"merge range is inverted: end_col {end_col} precedes "
            f"start_col {start_col}"
        )
    if not (0 <= start_col <= end_col < n_grid):
        raise TargetNotFound(
            f"column range {start_col}..{end_col} out of bounds "
            f"(grid has {n_grid} column(s), 0-based)"
        )
    if start_row == end_row and start_col == end_col:
        raise WordMcpError(
            "merge range is a single cell; nothing to merge (widen the range)"
        )

    collected_text: list[str] = []
    kept_cells: list[etree._Element] = []
    for r_i in range(start_row, end_row + 1):
        spans = [
            s
            for s in model[r_i]
            if s.grid_start >= start_col and s.grid_end <= end_col + 1
        ]
        covered = sum(s.grid_end - s.grid_start for s in spans)
        if covered != end_col - start_col + 1 or any(
            s.grid_start < start_col or s.grid_end > end_col + 1
            for s in model[r_i]
            if s.grid_start < end_col + 1 and s.grid_end > start_col
        ):
            raise UnsupportedStructure(
                f"merge rectangle does not align with cell boundaries in row {r_i}"
            )
        for s in spans:
            txt = _cell_text(s.tc).strip()
            if txt:
                collected_text.append(txt)
        # Horizontal merge within the row: keep first, absorb the rest.
        first = spans[0]
        for s in spans[1:]:
            s.tc.getparent().remove(s.tc)
        _set_grid_span(first.tc, end_col - start_col + 1)
        _set_tc_width(
            first.tc,
            sum(_col_width(tbl, i) for i in range(start_col, end_col + 1)),
        )
        kept_cells.append(first.tc)

    # Vertical merge across the kept cells.
    if end_row > start_row:
        for i, tc in enumerate(kept_cells):
            tcpr = _tcpr(tc)
            vm = tcpr.find(qn("w:vMerge"))
            if vm is None:
                vm = etree.SubElement(tcpr, qn("w:vMerge"))
            vm.set(qn("w:val"), "restart" if i == 0 else "continue")
            if i > 0:
                _cell_set_text(tc, "")

    _cell_set_text(kept_cells[0], "\n".join(collected_text))
    pkg.mark_dirty()

    # Absorbing an existing vertical merge can EXTEND the result past the
    # requested range: a pre-existing continuation chain hanging below
    # end_row now continues the new merge. Report the actual extent.
    actual_end = end_row
    if end_row > start_row:
        _rows2, model2 = _table_model(tbl)
        for r_j in range(end_row + 1, len(model2)):
            cont = next(
                (
                    s
                    for s in model2[r_j]
                    if s.grid_start == start_col
                    and s.grid_end == end_col + 1
                    and s.vmerge == "continue"
                ),
                None,
            )
            if cont is None:
                break
            actual_end = r_j
    result = {
        "merged": {
            "rows": [start_row, actual_end],
            "grid_cols": [start_col, end_col],
        }
    }
    if actual_end != end_row:
        result["requested"] = {
            "rows": [start_row, end_row],
            "grid_cols": [start_col, end_col],
        }
        result["note"] = (
            "the requested range overlapped an existing vertical merge "
            "whose continuation cells extend below it; the resulting merge "
            f"absorbed them and runs to row {actual_end}"
        )
    return result


def unmerge_cells(pkg: DocxPackage, table_index: int, *, row: int, cell: int) -> dict:
    """Undo merging at (row, cell) [CELL coordinates]: a horizontal merge is
    split into single cells (content stays in the first); a vertical merge
    chain is dissolved into independent cells."""
    tbl = _find_table(pkg, table_index)
    rows, model = _table_model(tbl)
    if not 0 <= row < len(rows):
        raise TargetNotFound(f"row {row} out of range")
    spans = model[row]
    if not 0 <= cell < len(spans):
        raise TargetNotFound(f"cell {cell} out of range in row {row}")
    span = spans[cell]
    did = []

    def _split_horizontal(s: CellSpan) -> bool:
        width = s.grid_end - s.grid_start
        if width <= 1:
            return False
        _set_grid_span(s.tc, 1)
        _set_tc_width(s.tc, _col_width(tbl, s.grid_start))
        for i in range(s.grid_start + 1, s.grid_end):
            s.tc.addnext(_new_cell(_col_width(tbl, i)))
        return True

    if span.vmerge == "restart":
        for r_j in range(row + 1, len(rows)):
            cont = next(
                (
                    s
                    for s in model[r_j]
                    if s.grid_start == span.grid_start
                    and s.grid_end == span.grid_end
                    and s.vmerge == "continue"
                ),
                None,
            )
            if cont is None:
                break
            _tcpr(cont.tc).remove(_tcpr(cont.tc).find(qn("w:vMerge")))
            _split_horizontal(cont)
            did.append(f"row {r_j} released")
        _tcpr(span.tc).remove(_tcpr(span.tc).find(qn("w:vMerge")))
        did.append("vertical merge dissolved")
    elif span.vmerge == "continue":
        raise WordMcpError(
            "this is a continuation cell; unmerge the restart cell of the chain"
        )

    if _split_horizontal(span):
        did.append(
            f"horizontal merge split into {span.grid_end - span.grid_start} cells"
        )

    if not did:
        raise WordMcpError("cell is not merged")
    pkg.mark_dirty()
    return {"unmerged": did}


# ------------------------------------------------------------ table create/etc


def create_table(
    pkg: DocxPackage,
    data: list[list[str]],
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    header_row: bool = True,
    width_pt: float | None = None,
) -> dict:
    """Create a table from 2D data. Borders single-line by default; first row
    optionally repeated as header on page breaks."""
    tbl = build_table_element(data, header_row=header_row, width_pt=width_pt)

    from .text import _body_paragraph, _resolve_anchor

    body = pkg.body()
    if at_end or (after_index is None and after_anchor is None):
        sectpr = body.find(qn("w:sectPr"))
        if sectpr is not None:
            sectpr.addprevious(tbl)
        else:
            body.append(tbl)
    elif after_anchor is not None:
        _resolve_anchor(pkg, after_anchor).addnext(tbl)
    else:
        _body_paragraph(pkg, after_index).addnext(tbl)
    # Word requires a paragraph between two tables and at body end; add a
    # spacer after the table unconditionally (harmless, removable).
    tbl.addnext(etree.Element(qn("w:p")))
    pkg.mark_dirty()
    n_cols = max(len(row) for row in data)
    return {"created": {"rows": len(data), "columns": n_cols}}


def build_table_element(
    data: list[list[str]],
    *,
    header_row: bool = True,
    width_pt: float | None = None,
) -> etree._Element:
    """Build (without inserting) the w:tbl element create_table produces;
    extracted so the batch layer's markdown table inserts share one
    construction path."""
    if not data or not data[0]:
        raise WordMcpError("data must be a non-empty 2D list")
    if width_pt is not None and width_pt <= 0:
        raise WordMcpError("width_pt must be positive")
    n_cols = max(len(row) for row in data)
    total_twips = int((width_pt or 468) * 20)  # default ~6.5in usable width
    col_w = total_twips // n_cols

    tbl = etree.Element(qn("w:tbl"))
    tblpr = etree.SubElement(tbl, qn("w:tblPr"))
    tblw = etree.SubElement(tblpr, qn("w:tblW"))
    tblw.set(qn("w:w"), str(total_twips))
    tblw.set(qn("w:type"), "dxa")
    borders = etree.SubElement(tblpr, qn("w:tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = etree.SubElement(borders, qn(f"w:{side}"))
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
    grid = etree.SubElement(tbl, qn("w:tblGrid"))
    for _ in range(n_cols):
        gc = etree.SubElement(grid, qn("w:gridCol"))
        gc.set(qn("w:w"), str(col_w))
    for r_i, row in enumerate(data):
        tr = etree.SubElement(tbl, qn("w:tr"))
        if r_i == 0 and header_row:
            trpr = etree.SubElement(tr, qn("w:trPr"))
            etree.SubElement(trpr, qn("w:tblHeader"))
        for c_i in range(n_cols):
            tc = _new_cell(col_w)
            _cell_set_text(tc, str(row[c_i]) if c_i < len(row) else "")
            if r_i == 0 and header_row:
                for r_el in tc.iter(qn("w:r")):
                    rpr = r_el.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = etree.Element(qn("w:rPr"))
                        r_el.insert(0, rpr)
                    if rpr.find(qn("w:b")) is None:
                        etree.SubElement(rpr, qn("w:b"))
            tr.append(tc)
    return tbl


def delete_table(pkg: DocxPackage, table_index: int) -> dict:
    tbl = _find_table(pkg, table_index)
    pkg.body().remove(tbl)
    pkg.mark_dirty()
    from .notes import purge_orphans

    result = {"deleted_table": table_index}
    purged = purge_orphans(pkg)["purged"]
    if purged:
        result["note_definitions_purged"] = purged
    return result


def set_column_widths(
    pkg: DocxPackage, table_index: int, widths_pt: list[float]
) -> dict:
    """Set every grid column width (list length must equal grid column count).
    Cell widths are recomputed from spans."""
    tbl = _find_table(pkg, table_index)
    cols = _grid_cols(tbl)
    if len(widths_pt) != len(cols):
        raise WordMcpError(
            f"expected {len(cols)} widths, got {len(widths_pt)}"
        )
    if any(w <= 0 for w in widths_pt):
        raise WordMcpError("column widths must all be positive points")
    for gc, w in zip(cols, widths_pt):
        gc.set(qn("w:w"), str(int(w * 20)))
    rows, model = _table_model(tbl)
    for row in model:
        for span in row:
            _set_tc_width(span.tc, _span_width(tbl, span))
    tblpr = tbl.find(qn("w:tblPr"))
    if tblpr is not None:
        tblw = tblpr.find(qn("w:tblW"))
        if tblw is not None and tblw.get(qn("w:type")) == "dxa":
            tblw.set(qn("w:w"), str(sum(int(w * 20) for w in widths_pt)))
    pkg.mark_dirty()
    return {"widths_set": len(widths_pt)}


# ----------------------------------------------------- Phase B table features


def apply_table_style(
    pkg: DocxPackage,
    table_index: int,
    style: str = "TableGrid",
    *,
    banded_rows: bool = True,
    first_row_header: bool = True,
) -> dict:
    """Apply a named table style (definition injected if the document lacks
    it). style: TableGrid (bordered) | PlainTable (borderless) | BandedTable
    (alternating row shading)."""
    known = {"TableGrid", "PlainTable", "BandedTable"}
    if style not in known:
        raise WordMcpError(f"style must be one of {sorted(known)}")
    tbl = _find_table(pkg, table_index)
    _ensure_table_style(pkg, style)
    tblpr = tbl.find(qn("w:tblPr"))
    if tblpr is None:
        tblpr = etree.Element(qn("w:tblPr"))
        tbl.insert(0, tblpr)
    st = tblpr.find(qn("w:tblStyle"))
    if st is None:
        st = etree.Element(qn("w:tblStyle"))
        tblpr.insert(0, st)
    st.set(qn("w:val"), style)
    # Direct borders on the table would fight the style; drop them.
    borders = tblpr.find(qn("w:tblBorders"))
    if borders is not None and style != "TableGrid":
        tblpr.remove(borders)
    look = tblpr.find(qn("w:tblLook"))
    if look is None:
        look = etree.SubElement(tblpr, qn("w:tblLook"))
    look.set(qn("w:val"), "04A0")
    look.set(qn("w:firstRow"), "1" if first_row_header else "0")
    look.set(qn("w:lastRow"), "0")
    look.set(qn("w:firstColumn"), "0")
    look.set(qn("w:lastColumn"), "0")
    look.set(qn("w:noHBand"), "0" if banded_rows else "1")
    look.set(qn("w:noVBand"), "1")
    pkg.mark_dirty()
    return {"table": table_index, "style": style}


def _ensure_table_style(pkg: DocxPackage, style_id: str) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    if style_id in have:
        return
    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), "table")
    s.set(qn("w:styleId"), style_id)
    names = {
        "TableGrid": "Table Grid",
        "PlainTable": "Plain Table",
        "BandedTable": "Banded Table",
    }
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), names[style_id])
    etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "TableNormal")
    tblpr = etree.SubElement(s, qn("w:tblPr"))
    if style_id == "TableGrid":
        borders = etree.SubElement(tblpr, qn("w:tblBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b = etree.SubElement(borders, qn(f"w:{side}"))
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), "4")
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "auto")
    if style_id == "BandedTable":
        band = etree.SubElement(s, qn("w:tblStylePr"))
        band.set(qn("w:type"), "band1Horz")
        tcpr = etree.SubElement(band, qn("w:tcPr"))
        shd = etree.SubElement(tcpr, qn("w:shd"))
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F2F2F2")
        first = etree.SubElement(s, qn("w:tblStylePr"))
        first.set(qn("w:type"), "firstRow")
        frpr = etree.SubElement(first, qn("w:tcPr"))
        fshd = etree.SubElement(frpr, qn("w:shd"))
        fshd.set(qn("w:val"), "clear")
        fshd.set(qn("w:color"), "auto")
        fshd.set(qn("w:fill"), "D9E2F3")
    if "TableNormal" not in have:
        tn = etree.SubElement(root, qn("w:style"))
        tn.set(qn("w:type"), "table")
        tn.set(qn("w:styleId"), "TableNormal")
        tn.set(qn("w:default"), "1")
        etree.SubElement(tn, qn("w:name")).set(qn("w:val"), "Normal Table")
        tnpr = etree.SubElement(tn, qn("w:tblPr"))
        mar = etree.SubElement(tnpr, qn("w:tblCellMar"))
        for side, w in (("top", "0"), ("left", "108"), ("bottom", "0"), ("right", "108")):
            el = etree.SubElement(mar, qn(f"w:{side}"))
            el.set(qn("w:w"), w)
            el.set(qn("w:type"), "dxa")
    pkg.mark_dirty("word/styles.xml")


def sort_table(
    pkg: DocxPackage,
    table_index: int,
    *,
    column: int,
    numeric: bool = False,
    descending: bool = False,
    has_header: bool = True,
) -> dict:
    """Sort data rows by a column (CELL index). Refused when vertical merges
    make rows interdependent."""
    tbl = _find_table(pkg, table_index)
    rows, model = _table_model(tbl)
    if any(s.vmerge for row in model for s in row):
        raise UnsupportedStructure(
            "table has vertical merges; sorting would break them"
        )
    start = 1 if has_header else 0
    data_rows = rows[start:]
    if not data_rows:
        raise WordMcpError("no data rows to sort")
    max_cells = max(len(tr.findall(qn("w:tc"))) for tr in data_rows)
    if not 0 <= column < max_cells:
        raise TargetNotFound(
            f"column {column} out of range (rows have up to {max_cells} cells)"
        )

    def key(tr):
        tcs = tr.findall(qn("w:tc"))
        if column >= len(tcs):
            return (1, "")  # rows lacking the column sort last
        text = _cell_text(tcs[column]).strip()
        if numeric:
            import re as _re

            m = _re.search(r"-?\d+(?:[.,]\d+)?", text.replace(",", ""))
            return (0, float(m.group(0)) if m else float("inf"))
        return (0, text.lower())

    non_numeric = 0
    if numeric:
        import re as _re

        for tr in data_rows:
            tcs = tr.findall(qn("w:tc"))
            if column < len(tcs):
                txt = _cell_text(tcs[column]).strip()
                if not _re.search(r"-?\d", txt):
                    non_numeric += 1
    ordered = sorted(data_rows, key=key, reverse=descending)
    for tr in ordered:
        tbl.append(tr)  # append moves the element to the end, in order
    pkg.mark_dirty()
    result = {"sorted_by_column": column, "rows": len(ordered)}
    if non_numeric:
        result["non_numeric_cells_sorted_last"] = non_numeric
    return result


def split_table(pkg: DocxPackage, table_index: int, *, at_row: int) -> dict:
    """Split a table into two at row `at_row` (that row starts the new
    table). Refused if a vertical merge crosses the boundary."""
    import copy as _copy

    tbl = _find_table(pkg, table_index)
    rows, model = _table_model(tbl)
    if not 1 <= at_row < len(rows):
        raise WordMcpError(
            f"at_row must be 1..{len(rows) - 1} (row 0 cannot start a split)"
        )
    if any(s.vmerge == "continue" for s in model[at_row]):
        raise UnsupportedStructure(
            "a vertical merge crosses the split point; unmerge first"
        )
    new_tbl = etree.Element(qn("w:tbl"))
    tblpr = tbl.find(qn("w:tblPr"))
    grid = tbl.find(qn("w:tblGrid"))
    if tblpr is not None:
        new_tbl.append(_copy.deepcopy(tblpr))
    if grid is not None:
        new_tbl.append(_copy.deepcopy(grid))
    for tr in rows[at_row:]:
        new_tbl.append(tr)  # moves
    spacer = etree.Element(qn("w:p"))
    tbl.addnext(spacer)
    spacer.addnext(new_tbl)
    pkg.mark_dirty()
    return {"split_at_row": at_row, "new_table_index": table_index + 1}


def set_header_row_repeat(
    pkg: DocxPackage, table_index: int, *, rows: int = 1, on: bool = True
) -> dict:
    """Repeat the first N rows as the table header on every page."""
    tbl = _find_table(pkg, table_index)
    trs = tbl.findall(qn("w:tr"))
    if rows > len(trs):
        raise WordMcpError(f"table has only {len(trs)} rows")
    changed = 0
    for i, tr in enumerate(trs):
        trpr = tr.find(qn("w:trPr"))
        header_el = (
            trpr.find(qn("w:tblHeader")) if trpr is not None else None
        )
        should = on and i < rows
        if should and header_el is None:
            if trpr is None:
                trpr = etree.Element(qn("w:trPr"))
                tr.insert(0, trpr)
            etree.SubElement(trpr, qn("w:tblHeader"))
            changed += 1
        elif not should and header_el is not None:
            trpr.remove(header_el)
            changed += 1
    pkg.mark_dirty()
    return {"header_rows": rows if on else 0, "rows_changed": changed}


def get_nested_table(
    pkg: DocxPackage,
    table_index: int,
    *,
    row: int,
    cell: int,
    nested_index: int = 0,
) -> dict:
    """Read a table nested inside a cell of a body-level table."""
    tbl = _find_table(pkg, table_index)
    trs = tbl.findall(qn("w:tr"))
    if not 0 <= row < len(trs):
        raise TargetNotFound(f"row {row} out of range")
    tcs = trs[row].findall(qn("w:tc"))
    if not 0 <= cell < len(tcs):
        raise TargetNotFound(f"cell {cell} out of range")
    nested = tcs[cell].findall(qn("w:tbl"))
    if not nested:
        raise TargetNotFound(f"no nested table in ({row},{cell})")
    if not 0 <= nested_index < len(nested):
        raise TargetNotFound(
            f"nested_index {nested_index} out of range ({len(nested)} nested)"
        )
    from .read import _table_dict

    d = _table_dict(nested[nested_index], nested_index)
    d["nested_in"] = {"table": table_index, "row": row, "cell": cell}
    return d


def set_nested_cells(
    pkg: DocxPackage,
    table_index: int,
    *,
    row: int,
    cell: int,
    edits: list[dict],
    nested_index: int = 0,
) -> dict:
    """Bulk cell writes inside a NESTED table."""
    tbl = _find_table(pkg, table_index)
    trs = tbl.findall(qn("w:tr"))
    if not 0 <= row < len(trs):
        raise TargetNotFound(f"row {row} out of range")
    tcs = trs[row].findall(qn("w:tc"))
    if not 0 <= cell < len(tcs):
        raise TargetNotFound(f"cell {cell} out of range")
    nested = tcs[cell].findall(qn("w:tbl"))
    if not nested or not 0 <= nested_index < len(nested):
        raise TargetNotFound("nested table not found")
    ntbl = nested[nested_index]
    nrows = ntbl.findall(qn("w:tr"))
    applied = 0
    for edit in edits:
        r, c = edit["row"], edit["cell"]
        if not 0 <= r < len(nrows):
            raise TargetNotFound(f"nested row {r} out of range")
        ntcs = nrows[r].findall(qn("w:tc"))
        if not 0 <= c < len(ntcs):
            raise TargetNotFound(f"nested cell {c} out of range")
        _cell_set_text(ntcs[c], str(edit["text"]))
        applied += 1
    pkg.mark_dirty()
    return {"nested_cells_written": applied}
