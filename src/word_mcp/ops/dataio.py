"""Data plumbing: table <-> CSV/JSON, bulk image extraction, document split.

Conventions:
- Table data travels as a rectangular ROW x GRID-COLUMN matrix. Merged cells
  put their value in the anchor position (top-left of the merge); every
  covered position holds an empty string. A separate merges list
  [{row, col, rowspan, colspan}] describes the merge topology so a round trip
  loses no information.
- All file-producing operations refuse existing output files BEFORE writing
  anything (atomicity: a refusal never leaves partial output behind).
- split_document is the inverse of the chapter merge: one saved .docx in,
  one standalone .docx per heading section out. The source is never touched.
"""

from __future__ import annotations

import csv
import io
import json
import posixpath
import re
import shutil
from pathlib import Path

from lxml import etree

from ..core.errors import (
    TargetNotFound,
    ValidationFailed,
    WordMcpError,
)
from ..core.package import DocxPackage, qn
from . import media as _media
from . import notes as _notes
from . import tables as _tables
from .read import (
    _outline_level,
    _style_outline_map,
    body_items,
    paragraph_text,
)

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_IMAGE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


# ----------------------------------------------------------- table grid model


def _flat_cell_text(tc: etree._Element) -> tuple[str, int]:
    """(text, nested_table_count) for a cell. Direct paragraphs first, then
    each nested table flattened row-by-row (cells joined with ' | ')."""
    parts = [paragraph_text(p) for p in tc.findall(qn("w:p"))]
    nested = tc.findall(qn("w:tbl"))
    count = len(nested)
    for ntbl in nested:
        for tr in ntbl.findall(qn("w:tr")):
            row_texts = []
            for ntc in tr.findall(qn("w:tc")):
                t, n = _flat_cell_text(ntc)
                count += n
                row_texts.append(t)
            parts.append(" | ".join(row_texts))
    return "\n".join(parts).strip("\n"), count


def _table_grid_data(
    tbl: etree._Element,
) -> tuple[list[list[str]], list[dict], list[dict]]:
    """(grid, merges, nested_notes) for a table.

    grid: rows x grid-columns matrix; merged-cell values sit at the anchor,
    covered positions are ''. merges: [{row, col, rowspan, colspan}] for every
    anchor spanning more than one grid position. nested_notes: cells that
    contained nested tables (their text was flattened into the grid)."""
    _rows, model = _tables._table_model(tbl)
    n_grid = len(_tables._grid_cols(tbl))
    grid: list[list[str]] = []
    merges: list[dict] = []
    nested_notes: list[dict] = []
    for r_i, row_spans in enumerate(model):
        row = [""] * n_grid
        for span in row_spans:
            if span.vmerge == "continue":
                continue  # covered by a vertical merge above
            text, nested = _flat_cell_text(span.tc)
            row[span.grid_start] = text
            if nested:
                nested_notes.append(
                    {
                        "row": r_i,
                        "col": span.grid_start,
                        "nested_tables": nested,
                        "note": "nested table content flattened into cell text",
                    }
                )
            colspan = span.grid_end - span.grid_start
            rowspan = 1
            if span.vmerge == "restart":
                for r_j in range(r_i + 1, len(model)):
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
                    rowspan += 1
            if colspan > 1 or rowspan > 1:
                merges.append(
                    {
                        "row": r_i,
                        "col": span.grid_start,
                        "rowspan": rowspan,
                        "colspan": colspan,
                    }
                )
        grid.append(row)
    return grid, merges, nested_notes


# ---------------------------------------------------------------- export_table


def export_table(
    pkg: DocxPackage,
    table_index: int,
    *,
    format: str = "csv",
    output_path: str | None = None,
    include_merges: bool = True,
) -> dict:
    """Export a body-level table to CSV or JSON.

    The output is a rows x grid-columns matrix. Merged cells keep their value
    in the anchor (top-left) position; covered positions are empty strings.
    With include_merges, the merge topology is reported as a list of
    {row, col, rowspan, colspan} entries — embedded in the JSON document, and
    returned as a report field alongside CSV (CSV itself cannot carry it).
    Nested tables are flattened into their host cell's text and flagged.

    output_path=None returns the data inline: a 'csv' string for CSV, a
    'rows' matrix for JSON. With output_path, the file must not already
    exist (refused before anything is written)."""
    if format not in ("csv", "json"):
        raise WordMcpError("format must be csv or json")
    out = Path(output_path) if output_path else None
    if out is not None and out.exists():
        raise WordMcpError(
            f"output file already exists: {out} — refusing to overwrite; "
            "delete it or choose another path"
        )

    tbl = _tables._find_table(pkg, table_index)
    grid, merges, nested_notes = _table_grid_data(tbl)

    result: dict = {
        "table_index": table_index,
        "format": format,
        "rows": len(grid),
        "columns": len(grid[0]) if grid else 0,
    }
    if include_merges:
        result["merges"] = merges
    if nested_notes:
        result["nested_table_cells"] = nested_notes

    if format == "csv":
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerows(grid)
        payload = buf.getvalue()
        if out is None:
            result["csv"] = payload
        else:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(payload, encoding="utf-8", newline="")
            result["output_path"] = str(out)
    else:
        doc: dict = {"table_index": table_index, "rows": grid}
        if include_merges:
            doc["merges"] = merges
        if out is None:
            result["data"] = grid
        else:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(
                json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            result["output_path"] = str(out)
    return result


# ---------------------------------------------------------------- import_table


def _load_rows(data) -> list[list[str]]:
    """Normalize the import payload (CSV path / JSON path / inline list of
    lists) into a list of string rows."""
    if isinstance(data, str):
        p = Path(data)
        if not p.exists():
            raise TargetNotFound(f"data file not found: {data}")
        suffix = p.suffix.lower()
        if suffix == ".csv":
            with open(p, encoding="utf-8-sig", newline="") as fh:
                rows = [list(row) for row in csv.reader(fh)]
        elif suffix == ".json":
            loaded = json.loads(p.read_text(encoding="utf-8-sig"))
            if isinstance(loaded, dict) and "rows" in loaded:
                rows = loaded["rows"]
            elif isinstance(loaded, list):
                rows = loaded
            else:
                raise WordMcpError(
                    "JSON data must be a list of rows or an object with a "
                    "'rows' key (the export_table JSON form)"
                )
        else:
            raise WordMcpError(
                f"data file must be .csv or .json, got {suffix or 'no extension'}"
            )
    elif isinstance(data, list):
        rows = data
    else:
        raise WordMcpError(
            "data must be a CSV/JSON file path or an inline list of rows"
        )
    if not rows or not isinstance(rows[0], (list, tuple)) or not rows[0]:
        raise WordMcpError("data must be a non-empty 2D list of rows")
    for r in rows:
        if not isinstance(r, (list, tuple)):
            raise WordMcpError("every data row must be a list")
    return [
        ["" if v is None else str(v) for v in row] for row in rows
    ]


def _shape_desc(rows: list[list[str]]) -> str:
    lens = {len(r) for r in rows}
    if len(lens) == 1:
        return f"{len(rows)}x{lens.pop()}"
    return f"{len(rows)}x{min(lens)}-{max(lens)}"


def import_table(
    pkg: DocxPackage,
    data,
    *,
    table_index: int | None = None,
    at_end: bool = False,
    after_anchor: str | None = None,
    has_header: bool = True,
) -> dict:
    """Fill a document table from CSV / JSON / inline data.

    data: a .csv file path, a .json file path (export_table's JSON form, or a
    bare list of rows), or an inline list of lists.

    Without table_index, a NEW table is created (positioned by at_end /
    after_anchor; default end of document; has_header bolds and repeats the
    first row). With table_index, the EXISTING table's cell texts are
    overwritten in place: data dimensions must exactly match the table's
    rows x grid-columns shape (mismatches are refused, listing both shapes),
    merged cells take their value from the anchor position, and every
    merge-covered position in the data must be empty (a value there would be
    silently invisible in Word, so it is refused instead)."""
    rows = _load_rows(data)

    if table_index is None:
        result = _tables.create_table(
            pkg,
            rows,
            at_end=at_end,
            after_anchor=after_anchor,
            header_row=has_header,
        )
        result["mode"] = "created"
        return result

    if at_end or after_anchor is not None:
        raise WordMcpError(
            "give either table_index (overwrite in place) or a placement "
            "(at_end / after_anchor) for a new table, not both"
        )
    tbl = _tables._find_table(pkg, table_index)
    _trs, model = _tables._table_model(tbl)
    n_grid = len(_tables._grid_cols(tbl))

    uniform = len({len(r) for r in rows}) == 1
    if len(rows) != len(model) or not uniform or len(rows[0]) != n_grid:
        raise WordMcpError(
            f"dimension mismatch: table {table_index} is "
            f"{len(model)}x{n_grid} (rows x grid columns) but data is "
            f"{_shape_desc(rows)} — refusing to overwrite; fix the data or "
            "omit table_index to create a new table"
        )

    # Anchor positions take values; every other grid position must be empty.
    anchors: set[tuple[int, int]] = set()
    edits: list[dict] = []
    for r_i, row_spans in enumerate(model):
        for c_i, span in enumerate(row_spans):
            if span.vmerge == "continue":
                continue
            anchors.add((r_i, span.grid_start))
            edits.append(
                {"row": r_i, "cell": c_i, "text": rows[r_i][span.grid_start]}
            )
    offenders = [
        {"row": r_i, "col": c_i, "value": rows[r_i][c_i]}
        for r_i in range(len(rows))
        for c_i in range(n_grid)
        if (r_i, c_i) not in anchors and rows[r_i][c_i].strip()
    ]
    if offenders:
        raise WordMcpError(
            "data has values in merge-covered positions, which Word would "
            f"never display: {offenders[:10]} — move each value to its merge "
            "anchor (top-left of the merge) or blank it"
        )

    result = _tables.set_cells(pkg, table_index, edits)
    result["mode"] = "overwritten"
    result["table_index"] = table_index
    return result


# -------------------------------------------------------------- extract_images


def _rels_base(rels_part: str) -> str:
    """'word/_rels/document.xml.rels' -> 'word'; '_rels/.rels' -> ''."""
    head = rels_part.rsplit("_rels/", 1)[0]
    return head.rstrip("/")


def _resolve_rel_target(rels_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(_rels_base(rels_part), target))


def _body_position_of(pkg: DocxPackage, el: etree._Element) -> dict:
    """Where a document.xml element sits: its top-level body child, described
    as {'body_paragraph': i} or {'body_table': i}."""
    body = pkg.body()
    node = el
    while node is not None and node.getparent() is not body:
        node = node.getparent()
    if node is None:
        return {"location": "outside body"}
    for kind, idx, child in body_items(pkg):
        if child is node:
            key = "body_paragraph" if kind == "paragraph" else "body_table"
            return {key: idx}
    return {"location": "body (unindexed element)"}


def extract_images(
    pkg: DocxPackage, output_dir: str, *, prefix: str | None = None
) -> dict:
    """Extract every image part of the document to files in output_dir.

    Body images are named by their list_images index plus the original
    extension (image0.png, image1.jpg, ...; prefix replaces 'image' when
    given). Media parts referenced only from headers, footers, or notes get
    the next indices after the body images, with their referencing part
    reported. Every image entry reports the output file, native pixel
    dimensions (PNG/JPEG/GIF; None for other formats), and where it appears
    in the document.

    Any name collision with an existing file is refused BEFORE anything is
    written — no partial extraction."""
    out_dir = Path(output_dir)
    if out_dir.exists() and not out_dir.is_dir():
        raise WordMcpError(
            f"output_dir points at an existing FILE, not a directory: "
            f"{out_dir}"
        )
    base = prefix if prefix else "image"

    # Body images, in list_images order (same blip iteration order).
    entries: list[dict] = []
    body_targets: set[str] = set()
    listed = _media.list_images(pkg)
    _A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    blips = list(pkg.root().iter(f"{{{_A}}}blip"))
    for info, blip in zip(listed, blips):
        target = info.get("target")
        if not target:
            entries.append(
                {
                    "index": info["index"],
                    "error": "image has no relationship target; not extracted",
                }
            )
            continue
        part = _resolve_rel_target("word/_rels/document.xml.rels", target)
        body_targets.add(part)
        entries.append(
            {
                "index": info["index"],
                "part": part,
                "appears": _body_position_of(pkg, blip),
            }
        )

    # Media parts referenced from elsewhere (headers, footers, notes) or
    # referenced from nothing at all.
    referenced_by: dict[str, list[str]] = {}
    for name in pkg.part_names():
        if not name.endswith(".rels"):
            continue
        for rel in pkg.root(name):
            if rel.get("TargetMode") == "External":
                continue
            resolved = _resolve_rel_target(name, rel.get("Target", ""))
            if resolved.startswith("word/media/"):
                base_dir = _rels_base(name)
                source_part = name.rsplit("_rels/", 1)[1][: -len(".rels")]
                referenced_by.setdefault(resolved, []).append(
                    f"{base_dir}/{source_part}" if base_dir else source_part
                )
    next_index = len(listed)
    for part in pkg.part_names():
        if not part.startswith("word/media/") or part in body_targets:
            continue
        sources = referenced_by.get(part)
        entries.append(
            {
                "index": next_index,
                "part": part,
                "appears": {
                    "referenced_from": sources if sources else "unreferenced"
                },
            }
        )
        next_index += 1

    # Assign output names and refuse collisions before any write.
    planned: list[tuple[dict, Path, bytes]] = []
    collisions: list[str] = []
    for entry in entries:
        if "part" not in entry:
            continue
        part = entry["part"]
        if not pkg.has_part(part):
            entry["error"] = f"relationship targets missing part {part}"
            continue
        ext = "." + part.rsplit(".", 1)[1].lower() if "." in part else ""
        dest = out_dir / f"{base}{entry['index']}{ext}"
        if dest.exists():
            collisions.append(str(dest))
        data = pkg.raw_part(part)
        px = _media._image_size_px(data, ext)
        entry["width_px"], entry["height_px"] = px if px else (None, None)
        entry["file"] = str(dest)
        planned.append((entry, dest, data))
    if collisions:
        raise WordMcpError(
            f"output files already exist: {collisions} — refusing to "
            "overwrite; nothing was written. Clear them or use another "
            "output_dir/prefix"
        )
    if not planned:
        raise TargetNotFound("document contains no image parts")

    out_dir.mkdir(parents=True, exist_ok=True)
    for _entry, dest, data in planned:
        dest.write_bytes(data)
    return {
        "output_dir": str(out_dir),
        "extracted": len(planned),
        "images": entries,
    }


# -------------------------------------------------------------- split_document

_INVALID_FS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _sanitize_filename(text: str) -> str:
    clean = _INVALID_FS.sub("_", text)
    clean = re.sub(r"\s+", " ", clean).strip(" ._")
    return clean[:80].rstrip(" .") or "section"


def _remove_part(pkg: DocxPackage, name: str) -> None:
    """Drop a part from the package (DocxPackage keeps no public remover;
    the save loop iterates _order, so removing from the internals is the
    supported-by-construction path)."""
    if name in pkg._raw:
        pkg._raw.pop(name, None)
        pkg._order.remove(name)
        pkg._trees.pop(name, None)
        pkg._dirty.discard(name)


def _prune_unused_images(pkg: DocxPackage) -> list[str]:
    """After body pruning: drop image relationships document.xml no longer
    references, then drop media parts no remaining relationship targets
    (header/footer/notes rels keep their images alive)."""
    used_rids: set[str] = set()
    for el in pkg.root().iter():
        for key, value in el.attrib.items():
            if key.startswith("{" + _R_NS + "}"):
                used_rids.add(value)
    rels_part = "word/_rels/document.xml.rels"
    if pkg.has_part(rels_part):
        rels_root = pkg.root(rels_part)
        doomed = [
            rel
            for rel in rels_root.findall(f"{{{_REL_NS}}}Relationship")
            if rel.get("Type") == _IMAGE_REL_TYPE
            and rel.get("Id") not in used_rids
        ]
        if doomed:
            for rel in doomed:
                rels_root.remove(rel)
            pkg.mark_dirty(rels_part)

    referenced: set[str] = set()
    for name in pkg.part_names():
        if not name.endswith(".rels"):
            continue
        for rel in pkg.root(name):
            if rel.get("TargetMode") == "External":
                continue
            referenced.add(_resolve_rel_target(name, rel.get("Target", "")))
    removed = [
        name
        for name in pkg.part_names()
        if name.startswith("word/media/") and name not in referenced
    ]
    for name in removed:
        _remove_part(pkg, name)
    return removed


def _block_has_content(pkg: DocxPackage, children: list) -> bool:
    for child in children:
        local = etree.QName(child).localname
        if local == "tbl":
            return True
        if local == "p":
            if paragraph_text(child).strip():
                return True
            if child.find(f".//{qn('w:drawing')}") is not None:
                return True
            if child.find(f".//{qn('w:pict')}") is not None:
                return True
    return False


def split_document(
    path: str,
    output_dir: str,
    *,
    level: int = 1,
    filename_from: str = "heading",
) -> dict:
    """Split a saved document into one standalone .docx per heading section —
    the inverse of chapter merge. The source file is never modified.

    Sections start at each heading of `level` (or higher, so a level-1
    heading always closes a level-2 section) and run to the next such
    heading, carrying everything in between: paragraphs, tables, images,
    footnotes. Content before the first section heading becomes
    00_front_matter.docx (only when it actually contains something).

    Every output is a full standalone document: styles, numbering, settings,
    fonts, and themes carry over from the source; footnote/endnote
    definitions are kept only for the references that survive in that
    section (note ids never determine displayed numbers — Word numbers by
    reference order — so no renumbering is needed); image parts referenced
    only by other sections are dropped; headers and footers stay with the
    document's governing section properties.

    filename_from: 'heading' names files 01_Heading Text.docx (sanitized);
    'index' names them 01.docx, 02.docx, ... Existing output files are
    refused before anything is written, and a failure mid-split removes the
    outputs already produced."""
    if filename_from not in ("heading", "index"):
        raise WordMcpError("filename_from must be heading or index")
    if not 1 <= level <= 9:
        raise WordMcpError("level must be 1..9")
    src = Path(path)
    pkg = DocxPackage(src)  # read-only reference for partitioning

    style_outline = _style_outline_map(pkg)
    body = pkg.body()
    children = list(body)
    boundaries: list[tuple[int, str]] = []  # (child position, heading text)
    for pos, child in enumerate(children):
        if etree.QName(child).localname != "p":
            continue
        lvl = _outline_level(child, style_outline)
        if lvl is not None and lvl <= level:
            boundaries.append((pos, paragraph_text(child).strip()))
    if not boundaries:
        raise TargetNotFound(
            f"document has no headings at level {level} (or higher); "
            "nothing to split on — check get_outline for the actual levels"
        )

    sections: list[dict] = []
    first_pos = boundaries[0][0]
    front_children = [
        c
        for c in children[:first_pos]
        if etree.QName(c).localname != "sectPr"
    ]
    if front_children and _block_has_content(pkg, front_children):
        sections.append(
            {
                "heading": None,
                "positions": set(range(first_pos)),
                "filename": "00_front_matter.docx",
            }
        )
    for i, (pos, heading) in enumerate(boundaries):
        end = boundaries[i + 1][0] if i + 1 < len(boundaries) else len(children)
        if filename_from == "index":
            fname = f"{i + 1:02d}.docx"
        else:
            fname = f"{i + 1:02d}_{_sanitize_filename(heading)}.docx"
        sections.append(
            {
                "heading": heading,
                "positions": set(range(pos, end)),
                "filename": fname,
            }
        )

    out_dir = Path(output_dir)
    targets = [out_dir / s["filename"] for s in sections]
    existing = [str(t) for t in targets if t.exists()]
    if existing:
        raise WordMcpError(
            f"output files already exist: {existing} — refusing to "
            "overwrite; nothing was written"
        )
    out_dir.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    report_files: list[dict] = []
    try:
        for section, target in zip(sections, targets):
            shutil.copyfile(src, target)
            written.append(target)
            opkg = DocxPackage(target)
            obody = opkg.body()
            keep = section["positions"]
            for pos, child in enumerate(list(obody)):
                if pos in keep:
                    continue
                if etree.QName(child).localname == "sectPr":
                    continue  # governing section properties stay
                obody.remove(child)
            opkg.mark_dirty()
            purged = _notes.purge_orphans(opkg)["purged"]
            images_dropped = _prune_unused_images(opkg)
            opkg.save(do_backup=False)

            # Round-trip validation: reopen and check note integrity.
            vpkg = DocxPackage(target)
            vn = _notes.validate_notes(vpkg)
            problems = [
                k
                for k, v in vn.items()
                if not v["ok"] or v["needs_cleanup"]
            ]
            if problems:
                raise ValidationFailed(
                    f"{target.name} failed note validation ({problems}): {vn}"
                )
            from .read import get_document_info

            info = get_document_info(vpkg)
            report_files.append(
                {
                    "file": str(target),
                    "heading": section["heading"] or "(front matter)",
                    "paragraphs": info["paragraphs"],
                    "tables": info["tables"],
                    "footnotes": info["footnotes"],
                    "endnotes": info["endnotes"],
                    "images": info["images"],
                    "note_definitions_dropped": purged,
                    "image_parts_dropped": images_dropped,
                }
            )
    except Exception:
        for w in written:
            w.unlink(missing_ok=True)
        raise

    return {
        "source": str(src),
        "output_dir": str(out_dir),
        "level": level,
        "sections": len(report_files),
        "files": report_files,
    }
