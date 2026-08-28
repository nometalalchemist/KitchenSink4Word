"""Batch operations across document sets.

One DocxPackage, one backup, one atomic save per file: a file either receives
ALL requested operations or none of them (an op failure skips that file's
save, leaving it untouched on disk). Only an explicit allowlist of ops-level
callables can run — batch mode multiplies mistakes across files, so anything
destructive-by-surprise stays out.
"""

from __future__ import annotations

from pathlib import Path

from ..core.errors import WordMcpError
from ..core.package import DocxPackage
from ..core.safesave import write_lock
from . import furniture as _fu, structure as _sx, tables as _tb, text as _tx


def _set_header(pkg, p):
    return _fu.set_header_footer(pkg, "header", p.pop("text", ""), **p)


def _set_footer(pkg, p):
    return _fu.set_header_footer(pkg, "footer", p.pop("text", ""), **p)


def _add_watermark(pkg, p):
    return _fu.add_watermark(pkg, p.pop("text", "DRAFT"), **p)


# tool name (as exposed by the server) -> callable(pkg, params).
# Params dicts arrive pre-copied; adapters may pop from them freely.
_ALLOWED = {
    "search_and_replace": lambda pkg, p: _tx.search_and_replace(pkg, **p),
    "insert_paragraphs": lambda pkg, p: _tx.insert_paragraphs(pkg, **p),
    "delete_paragraphs": lambda pkg, p: _tx.delete_paragraphs(pkg, **p),
    "replace_paragraph_text": lambda pkg, p: _tx.replace_paragraph_text(pkg, **p),
    "format_text": lambda pkg, p: _tx.format_text(pkg, **p),
    "set_paragraph_format": lambda pkg, p: _tx.set_paragraph_format(pkg, **p),
    "apply_style": lambda pkg, p: _tx.apply_style(pkg, **p),
    "set_header": _set_header,
    "set_footer": _set_footer,
    "add_page_numbers": lambda pkg, p: _fu.add_page_numbers(pkg, **p),
    "set_page_number_format": lambda pkg, p: _fu.set_page_number_format(pkg, **p),
    "set_document_properties": lambda pkg, p: _sx.set_document_properties(pkg, **p),
    "set_cells": lambda pkg, p: _tb.set_cells(pkg, **p),
    "add_watermark": _add_watermark,
    "remove_watermark": lambda pkg, p: _fu.remove_watermark(pkg, **p),
}


def _validate_operations(operations: list[dict]) -> None:
    if not operations:
        raise WordMcpError("operations list is empty")
    for i, op in enumerate(operations):
        if not isinstance(op, dict) or "tool" not in op:
            raise WordMcpError(
                f"operation {i} must be a dict with 'tool' and optional 'params'"
            )
        if op["tool"] not in _ALLOWED:
            raise WordMcpError(
                f"tool {op['tool']!r} is not batchable; allowed: "
                f"{sorted(_ALLOWED)}"
            )
        params = op.get("params", {})
        if not isinstance(params, dict):
            raise WordMcpError(f"operation {i} ({op['tool']}): params must be a dict")


def batch_apply(
    file_paths: list[str],
    operations: list[dict],
    *,
    stop_on_error: bool = True,
    backup: bool = True,
) -> dict:
    """Apply the same operation list to many documents. Each operation:
    {'tool': <name>, 'params': {...}} using the tool's normal parameters
    minus file_path. Per file: all operations run in order, then ONE save
    (one backup); if any operation fails, that file is not saved and stays
    exactly as it was. stop_on_error=True additionally skips the remaining
    files — files already saved stay saved and are reported as such."""
    if not file_paths:
        raise WordMcpError("file_paths is empty")
    _validate_operations(operations)
    # sandbox check BEFORE the existence probe: a blocked path must refuse
    # identically whether or not the file exists (no existence oracle)
    from ..core.sandbox import check_path

    file_paths = [check_path(f, "batch edit") for f in file_paths]
    absent = [f for f in file_paths if not Path(f).is_file()]
    if absent:
        raise WordMcpError(
            "refusing to start (no file was touched): missing files "
            + ", ".join(absent)
        )

    files: list[dict] = []
    saved: list[str] = []
    failed: list[str] = []
    aborted = False
    for pos, file_path in enumerate(file_paths):
        if aborted:
            break
        entry: dict = {"path": file_path, "ok": False, "operations": []}
        files.append(entry)
        # Per-file write lock across the full read-modify-save cycle, same as
        # the single-file _edit path: a concurrent mutation of the same file
        # must not interleave with this one.
        with write_lock(file_path):
            try:
                pkg = DocxPackage(file_path)
            except Exception as exc:
                entry["error"] = f"{type(exc).__name__}: {exc}"
                failed.append(file_path)
                if stop_on_error:
                    aborted = True
                continue
            ops_ok = True
            for op in operations:
                tool = op["tool"]
                try:
                    result = _ALLOWED[tool](pkg, dict(op.get("params", {})))
                    entry["operations"].append(
                        {"tool": tool, "ok": True, "result": result}
                    )
                except TypeError as exc:
                    entry["operations"].append(
                        {"tool": tool, "ok": False,
                         "error": f"bad params for {tool}: {exc}"}
                    )
                    ops_ok = False
                    break
                except Exception as exc:
                    entry["operations"].append(
                        {"tool": tool, "ok": False,
                         "error": f"{type(exc).__name__}: {exc}"}
                    )
                    ops_ok = False
                    break
            if not ops_ok:
                entry["error"] = (
                    "an operation failed; this file was NOT saved and is unchanged"
                )
                failed.append(file_path)
                if stop_on_error:
                    aborted = True
                continue
            try:
                entry["saved"] = str(pkg.save(do_backup=backup))
                entry["ok"] = True
                saved.append(file_path)
            except Exception as exc:
                entry["error"] = (
                    f"save failed, file unchanged — {type(exc).__name__}: {exc}"
                )
                failed.append(file_path)
                if stop_on_error:
                    aborted = True

    not_attempted = file_paths[len(files):]
    result = {
        "files": files,
        "saved": saved,
        "failed": failed,
        "not_attempted": not_attempted,
    }
    if failed and saved:
        result["note"] = (
            f"{len(saved)} file(s) were already saved before the failure and "
            "keep their changes; failed and unattempted files are unchanged"
        )
    return result
