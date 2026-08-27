"""Mail merge / template population.

Placeholders are ``{{name}}`` markers typed in the document plus legacy Word
``MERGEFIELD`` fields. All matching runs through the runmap layer, so a
placeholder Word has fragmented across several runs (spell-check state,
formatting boundaries) is still found and replaced as one unit.

Nothing here writes to disk except mail_merge, which manages whole files;
fill_template and list_template_placeholders follow the standard ops contract
(take a DocxPackage, mark parts dirty, caller saves).
"""

from __future__ import annotations

import copy
import csv
import json
import re
from pathlib import Path

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap
from .text import _replace_parts

# {{ name }} — name may carry internal spaces; trimmed before matching keys.
_PLACEHOLDER_RE = re.compile(r"\{\{([^{}\n]{1,200}?)\}\}")
_MERGEFIELD_RE = re.compile(r"MERGEFIELD\s+(?:\"([^\"]+)\"|([^\s\\\"]+))")
# XML 1.0 forbids these outright (\x07, Word's cell separator, included);
# they would also poison every downstream save.
_BAD_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

_MISSING_MODES = ("error", "skip", "empty")


def _check_value(name: str, value) -> str:
    text = str(value)
    if _BAD_CHARS_RE.search(text):
        raise WordMcpError(
            f"value for {name!r} contains control characters (e.g. \\x07) "
            "that cannot be stored in document text"
        )
    return text


def _scan_complex_fields(p: etree._Element) -> list[dict]:
    """Complex fields (begin..end) fully contained in this paragraph.

    Each: {instr, runs, result_runs, nested}. Fields that span paragraphs or
    nest other fields are dropped/flagged rather than guessed at.
    """
    fields: list[dict] = []
    cur: dict | None = None
    depth = 0
    for r in p.iter(qn("w:r")):
        fc = r.find(qn("w:fldChar"))
        if fc is not None:
            ftype = fc.get(qn("w:fldCharType"))
            if ftype == "begin":
                if cur is None:
                    cur = {
                        "instr": "",
                        "runs": [r],
                        "result_runs": [],
                        "stage": "instr",
                        "nested": False,
                    }
                    depth = 1
                else:
                    cur["nested"] = True
                    cur["runs"].append(r)
                    depth += 1
                continue
            if cur is None:
                continue
            cur["runs"].append(r)
            if ftype == "separate" and depth == 1:
                cur["stage"] = "result"
            elif ftype == "end":
                depth -= 1
                if depth == 0:
                    fields.append(cur)
                    cur = None
            continue
        if cur is None:
            continue
        cur["runs"].append(r)
        if cur["stage"] == "instr" and depth == 1:
            for it in r.findall(qn("w:instrText")):
                cur["instr"] += it.text or ""
        elif cur["stage"] == "result":
            cur["result_runs"].append(r)
    # A field still open at paragraph end spans paragraphs: unsupported, drop.
    return fields


def _mergefield_name(instr: str) -> str | None:
    m = _MERGEFIELD_RE.search(instr)
    if not m:
        return None
    return (m.group(1) or m.group(2)).strip()


def _iter_mergefields(pkg: DocxPackage, parts: list[str]):
    """Yield (part, paragraph, kind, name, payload) for every MERGEFIELD.

    kind 'simple' -> payload is the w:fldSimple element;
    kind 'complex' -> payload is the field dict from _scan_complex_fields.
    """
    for part in parts:
        for p in pkg.root(part).iter(qn("w:p")):
            for fs in p.iter(qn("w:fldSimple")):
                name = _mergefield_name(fs.get(qn("w:instr"), ""))
                if name:
                    yield part, p, "simple", name, fs
            for field in _scan_complex_fields(p):
                if field["nested"]:
                    continue
                name = _mergefield_name(field["instr"])
                if name:
                    yield part, p, "complex", name, field


def _context(p: etree._Element, limit: int = 80) -> str:
    text, _ = _runmap.build_map(p)
    return text[:limit]


def list_template_placeholders(pkg: DocxPackage) -> dict:
    """Find every {{name}} placeholder and legacy MERGEFIELD in the document
    (body, tables, headers/footers, footnotes/endnotes). Returns names with
    counts and locations — the keys fill_template expects in its data dict."""
    parts = _replace_parts(pkg, "all")
    placeholders: dict[str, dict] = {}
    for part in parts:
        for p in pkg.root(part).iter(qn("w:p")):
            text, _ = _runmap.build_map(p)
            for m in _PLACEHOLDER_RE.finditer(text):
                name = m.group(1).strip()
                if not name:
                    continue
                entry = placeholders.setdefault(
                    name, {"name": name, "count": 0, "locations": []}
                )
                entry["count"] += 1
                if len(entry["locations"]) < 10:
                    entry["locations"].append(
                        {"part": part, "context": _context(p)}
                    )
    mergefields: dict[str, dict] = {}
    for part, p, kind, name, _payload in _iter_mergefields(pkg, parts):
        entry = mergefields.setdefault(
            name, {"name": name, "count": 0, "kind": kind, "locations": []}
        )
        entry["count"] += 1
        if len(entry["locations"]) < 10:
            entry["locations"].append({"part": part, "context": _context(p)})
    names = sorted(set(placeholders) | set(mergefields))
    return {
        "names": names,
        "placeholders": sorted(placeholders.values(), key=lambda e: e["name"]),
        "mergefields": sorted(mergefields.values(), key=lambda e: e["name"]),
    }


def _replace_mergefield(kind: str, payload, value: str) -> None:
    """Swap a whole MERGEFIELD (instruction + result) for plain text, keeping
    the formatting of its result run."""
    if kind == "simple":
        fs = payload
        ref_run = fs.find(qn("w:r"))
        new_run = etree.Element(qn("w:r"))
        if ref_run is not None:
            rpr = ref_run.find(qn("w:rPr"))
            if rpr is not None:
                new_run.append(copy.deepcopy(rpr))
        t = etree.SubElement(new_run, qn("w:t"))
        t.text = value
        _runmap._preserve_space(t)
        fs.addprevious(new_run)
        fs.getparent().remove(fs)
        return
    field = payload
    ref_run = field["result_runs"][0] if field["result_runs"] else field["runs"][0]
    new_run = etree.Element(qn("w:r"))
    rpr = ref_run.find(qn("w:rPr"))
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    t = etree.SubElement(new_run, qn("w:t"))
    t.text = value
    _runmap._preserve_space(t)
    field["runs"][0].addprevious(new_run)
    for r in field["runs"]:
        parent = r.getparent()
        if parent is not None:
            parent.remove(r)


def fill_template(pkg: DocxPackage, data: dict, *, missing: str = "error") -> dict:
    """Replace every {{name}} placeholder with data[name] (runmap-safe, first
    run's formatting preserved) and set MERGEFIELDs to their values as plain
    text. missing: 'error' refuses and changes nothing if the document needs a
    key data lacks; 'skip' leaves those markers and reports them; 'empty'
    fills them with empty strings. Values are coerced to str."""
    if missing not in _MISSING_MODES:
        raise WordMcpError(f"missing must be one of {_MISSING_MODES}")
    values = {str(k): _check_value(str(k), v) for k, v in data.items()}

    # Pass 1 (read-only): what does the document need? Refusal must precede
    # any mutation so a partial fill can never be saved.
    found = list_template_placeholders(pkg)
    needed = set(found["names"])
    missing_names = sorted(needed - set(values))
    if missing_names and missing == "error":
        raise WordMcpError(
            "template needs values for "
            f"{missing_names} that data does not provide; nothing was changed "
            "— add the keys or call with missing='skip' or 'empty'"
        )

    def value_for(name: str) -> str | None:
        if name in values:
            return values[name]
        if missing == "empty":
            return ""
        return None  # skip

    parts = _replace_parts(pkg, "all")
    ph_counts: dict[str, int] = {}
    mf_counts: dict[str, int] = {}
    for part in parts:
        dirty = False
        for p in pkg.root(part).iter(qn("w:p")):
            text, segments = _runmap.build_map(p)
            matches = []
            for m in _PLACEHOLDER_RE.finditer(text):
                name = m.group(1).strip()
                if not name:
                    continue
                repl = value_for(name)
                if repl is not None:
                    matches.append((m.start(), m.end(), name, repl))
            # Right-to-left on one snapshot: earlier offsets stay valid and
            # replacement text is never re-matched.
            for start, end, name, repl in reversed(matches):
                _runmap.replace_range(p, segments, start, end, repl)
                ph_counts[name] = ph_counts.get(name, 0) + 1
                dirty = True
        # Materialized first: replacement removes elements the scan yields.
        for _part, _p, kind, name, payload in list(
            _iter_mergefields(pkg, [part])
        ):
            repl = value_for(name)
            if repl is None:
                continue
            _replace_mergefield(kind, payload, repl)
            mf_counts[name] = mf_counts.get(name, 0) + 1
            dirty = True
        if dirty:
            pkg.mark_dirty(part)

    result = {
        "placeholders_replaced": ph_counts,
        "mergefields_replaced": mf_counts,
        "total": sum(ph_counts.values()) + sum(mf_counts.values()),
    }
    if missing_names:
        key = "skipped" if missing == "skip" else "filled_empty"
        result[key] = missing_names
    unused = sorted(set(values) - needed)
    if unused:
        result["unused_keys"] = unused
    return result


# ------------------------------------------------------------------ mail merge


def _load_rows(data_rows: list[dict] | str) -> list[dict]:
    if isinstance(data_rows, list):
        if not all(isinstance(r, dict) for r in data_rows):
            raise WordMcpError("data_rows list items must all be dicts")
        return data_rows
    path = Path(data_rows)
    if not path.is_file():
        raise WordMcpError(f"data file not found: {path}")
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = [
                {k: ("" if v is None else v) for k, v in row.items() if k is not None}
                for row in csv.DictReader(fh)
            ]
    elif suffix == ".json":
        with open(path, encoding="utf-8-sig") as fh:
            rows = json.load(fh)
        if not isinstance(rows, list) or not all(
            isinstance(r, dict) for r in rows
        ):
            raise WordMcpError(
                f"{path.name} must contain a JSON array of objects "
                "(one object per output document)"
            )
    else:
        raise WordMcpError(
            f"data file must be .csv or .json, got {path.suffix!r}"
        )
    if not rows:
        raise WordMcpError(f"{path.name} contains no data rows")
    return rows


def _output_name(pattern: str, row: dict, row_index: int) -> str:
    mapping = {str(k): str(v) for k, v in row.items()}
    mapping["row_index"] = str(row_index)
    try:
        name = pattern.format_map(mapping)
    except KeyError as exc:
        raise WordMcpError(
            f"filename_pattern references {{{exc.args[0]}}} which is not a "
            f"data column (available: row_index, {sorted(row)})"
        ) from None
    except (IndexError, ValueError) as exc:
        raise WordMcpError(f"bad filename_pattern {pattern!r}: {exc}") from None
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().rstrip(". ")
    if not name:
        raise WordMcpError(
            f"filename_pattern produced an empty name for row {row_index}"
        )
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name


def _cap_names(names: list, out_dir) -> list:
    """Windows path-length preflight: an over-long full path fails the OS
    write MID-RUN, after earlier rows already landed on disk. Budget covers
    the output dir, the atomic-save temp suffix, and MAX_PATH headroom."""
    budget = 240 - len(str(out_dir)) - len(".word-mcp-tmp")
    if budget < 15:
        raise WordMcpError(
            f"output directory path is too long ({len(str(out_dir))} chars) "
            "for safe Windows file writes; choose a shorter output_dir"
        )
    capped = []
    for name in names:
        if len(name) > budget:
            name = name[: budget - 6] + "….docx"
        capped.append(name)
    return capped


def mail_merge(
    template_path: str,
    data_rows: list[dict] | str,
    output_dir: str,
    *,
    filename_pattern: str = "{row_index}.docx",
    missing: str = "error",
) -> dict:
    """One filled document per data row. data_rows: list of dicts, or a path
    to a .csv (header row = placeholder names) or .json (array of objects).
    filename_pattern supports {row_index} (1-based) and any {column}; hostile
    filename characters are replaced with '_'. Refuses BEFORE writing anything
    if any output file already exists, if two rows collide on one name, or
    (missing='error') if any row lacks a value the template needs. The
    template file itself is never modified."""
    if missing not in _MISSING_MODES:
        raise WordMcpError(f"missing must be one of {_MISSING_MODES}")
    template = Path(template_path)
    rows = _load_rows(data_rows)

    # Template must load, and its required names drive the pre-flight check.
    needed = set(list_template_placeholders(DocxPackage(template))["names"])
    if missing == "error":
        problems = []
        for i, row in enumerate(rows, start=1):
            lacking = sorted(needed - {str(k) for k in row})
            if lacking:
                problems.append(f"row {i} lacks {lacking}")
        if problems:
            raise WordMcpError(
                "refusing to merge (nothing written): "
                + "; ".join(problems[:20])
                + ("; ..." if len(problems) > 20 else "")
                + " — fix the data or call with missing='skip' or 'empty'"
            )

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    names = _cap_names(
        [
            _output_name(filename_pattern, row, i)
            for i, row in enumerate(rows, start=1)
        ],
        out_dir,
    )
    dupes = sorted({n for n in names if names.count(n) > 1})
    if dupes:
        raise WordMcpError(
            f"filename_pattern maps multiple rows to the same file {dupes}; "
            "nothing was written — include {row_index} or a unique column"
        )
    collisions = [str(out_dir / n) for n in names if (out_dir / n).exists()]
    if collisions:
        raise WordMcpError(
            "refusing to overwrite existing files (nothing was written): "
            + ", ".join(collisions)
        )

    per_row = []
    outputs = []
    for i, (row, name) in enumerate(zip(rows, names), start=1):
        out_path = out_dir / name
        try:
            pkg = DocxPackage(template)
            fill = fill_template(pkg, row, missing=missing)
            pkg.save(out_path, do_backup=False)
        except OSError as exc:
            # honest partial-run reporting: earlier rows are on disk
            raise WordMcpError(
                f"row {i} ({name}) failed to write: {exc}. "
                f"{len(outputs)} earlier output(s) were already written and "
                f"remain on disk: {outputs or 'none'}"
            ) from exc
        outputs.append(str(out_path))
        per_row.append(
            {"row_index": i, "output": str(out_path), "total": fill["total"]}
        )
    return {
        "rows": len(rows),
        "template": str(template),
        "outputs": outputs,
        "per_row": per_row,
    }
