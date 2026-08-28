"""Backup management: list / restore / purge for the slot-based backup system
(core.safesave) plus the legacy per-mutation ``*.bak-*`` files it replaced.

Restore discipline: prev rotates FIRST (from the document's current content),
so a restore is itself undoable via prev. Restores refuse documents that are
open in Word (same detection the editing path uses) and validate the backup
payload before touching the target; the target is replaced atomically and is
never absent from its own path.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
import shutil
import uuid
from pathlib import Path

from ..core import safesave
from ..core.errors import DocumentLocked, DocumentNotFound, WordMcpError
from ..core.package import DocxPackage
from ..core.safesave import (
    ANCHOR_SLOT,
    BACKUP_DIR_NAME,
    PREV_SLOT,
    SLOT_POLICY,
    write_lock,
)
from ..core.sandbox import check_path

_LEGACY_GLOB = "*.bak-*.docx"


def _refuse_if_word_locked(path: Path) -> None:
    """Same detection DocxPackage uses: Word holds an exclusive lock."""
    owner_file = path.with_name("~$" + path.name[-153:])
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        hint = " (Word owner file present)" if owner_file.exists() else ""
        raise DocumentLocked(
            f"{path.name} is open in Word or locked by another process{hint}. "
            "Close it in Word before restoring a backup over it."
        ) from None


def _stat_entry(p: Path) -> dict:
    st = p.stat()
    return {
        "path": str(p),
        "size_bytes": st.st_size,
        "modified": _dt.datetime.fromtimestamp(st.st_mtime).isoformat(
            timespec="seconds"
        ),
    }


def _legacy_for_doc(doc: Path) -> list[Path]:
    return sorted(doc.parent.glob(f"{doc.stem}.bak-*{doc.suffix}"))


def _orphan_folders(directory: Path) -> list[Path]:
    """Slot folders under directory/.ks4w-backups whose source doc is gone."""
    root = directory / BACKUP_DIR_NAME
    if not root.is_dir():
        return []
    orphans = []
    for folder in sorted(p for p in root.iterdir() if p.is_dir()):
        if not safesave.source_doc_for(folder).exists():
            orphans.append(folder)
    return orphans


def _dir_size(folder: Path) -> int:
    return sum(p.stat().st_size for p in folder.rglob("*") if p.is_file())


# ---------------------------------------------------------------------- list


def list_backups(
    file_path: str | None = None, directory: str | None = None
) -> dict:
    """Backups for one document (file_path) or a whole folder (directory):
    slot files with sizes + mtimes, legacy *.bak-* files, and orphaned slot
    folders whose source document no longer exists."""
    if file_path is None and directory is None:
        raise WordMcpError("provide file_path (one document) or directory")
    if file_path is not None:
        check_path(file_path, "list backups")
        doc = Path(file_path).resolve()
        d = safesave.slot_dir(doc)
        slots = []
        for slot in SLOT_POLICY:
            sp = d / slot
            if sp.exists():
                slots.append({"slot": slot.split(".")[0], **_stat_entry(sp)})
        result = {
            "document": str(doc),
            "document_exists": doc.exists(),
            "slots": slots,
            "legacy_backups": [_stat_entry(p) for p in _legacy_for_doc(doc)],
            "orphaned_folders": [
                {"folder": str(f), "size_bytes": _dir_size(f),
                 "missing_document": str(safesave.source_doc_for(f))}
                for f in _orphan_folders(doc.parent)
            ],
        }
        return result

    check_path(directory, "list backups")
    base = Path(directory).resolve()
    if not base.is_dir():
        raise DocumentNotFound(f"no directory at {base}")
    root = base / BACKUP_DIR_NAME
    documents = []
    if root.is_dir():
        for folder in sorted(p for p in root.iterdir() if p.is_dir()):
            src = safesave.source_doc_for(folder)
            if not src.exists():
                continue  # reported under orphaned_folders below
            slots = []
            for slot in SLOT_POLICY:
                sp = folder / slot
                if sp.exists():
                    slots.append({"slot": slot.split(".")[0], **_stat_entry(sp)})
            documents.append({"document": str(src), "slots": slots})
    return {
        "directory": str(base),
        "documents": documents,
        "legacy_backups": [
            _stat_entry(p) for p in sorted(base.glob(_LEGACY_GLOB))
        ],
        "orphaned_folders": [
            {"folder": str(f), "size_bytes": _dir_size(f),
             "missing_document": str(safesave.source_doc_for(f))}
            for f in _orphan_folders(base)
        ],
    }


# ------------------------------------------------------------------- restore


def restore_backup(file_path: str, source: str) -> dict:
    """Replace the document's content with a backup. source: 'prev', 'anchor',
    or a path to a legacy *.bak-* file. Rotates prev from the current content
    FIRST, so the restore itself can be undone by restoring prev again."""
    check_path(file_path, "restore backup over document")
    doc = Path(file_path).resolve()
    if source in ("prev", "anchor"):
        src = safesave.slot_dir(doc) / (
            PREV_SLOT if source == "prev" else ANCHOR_SLOT
        )
        label = source
    else:
        check_path(source, "read backup file")
        src = Path(source).resolve()
        label = f"legacy file {src.name}"
    if not src.is_file():
        raise DocumentNotFound(
            f"no backup to restore: {src} does not exist. "
            "Use manage_backups action='list' to see what is available."
        )

    with write_lock(doc):
        target_existed = doc.exists()
        if target_existed:
            _refuse_if_word_locked(doc)

        # Validate the backup payload BEFORE touching anything.
        payload = src.read_bytes()
        DocxPackage._validate_payload(payload)

        # Copy (not hardlink) the source to a temp beside the target first:
        # rotating prev below may clobber the very slot being restored from,
        # and the restored target must own its bytes outright.
        d = safesave.slot_dir(doc, create=True)
        tmp = d / f".restore-{uuid.uuid4().hex}.tmp"
        try:
            shutil.copy2(src, tmp)
            rotated_prev = False
            if target_existed:
                safesave._place_onto_slot(doc, d / PREV_SLOT)
                rotated_prev = True
            safesave.replace_with_retry(tmp, doc)
        except BaseException:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            raise

    result = {
        "restored": str(doc),
        "from": label,
        "bytes": len(payload),
        "prev_rotated": rotated_prev,
    }
    if rotated_prev:
        result["undo"] = "restore source='prev' brings back the pre-restore content"
    else:
        result["note"] = "document did not exist; nothing rotated into prev"
    return result


# --------------------------------------------------------------------- purge


def _collect_purge_targets(
    scope: str, file_path: str | None, directory: str | None
) -> tuple[list[Path], Path | None]:
    """Returns (targets, slot_folder_for_slots_scope)."""
    if file_path is not None:
        check_path(file_path, "purge backups")
    if directory is not None:
        check_path(directory, "purge backups")
    if scope == "slots":
        if not file_path:
            raise WordMcpError("scope='slots' needs file_path (whose slots to purge)")
        doc = Path(file_path).resolve()
        d = safesave.slot_dir(doc)
        targets = [d / slot for slot in SLOT_POLICY if (d / slot).exists()]
        return targets, d
    if scope == "legacy":
        if file_path:
            return [Path(p) for p in _legacy_for_doc(Path(file_path).resolve())], None
        if directory:
            base = Path(directory).resolve()
            return sorted(base.glob(_LEGACY_GLOB)), None
        raise WordMcpError("scope='legacy' needs file_path or directory")
    if scope == "orphans":
        base = (
            Path(file_path).resolve().parent
            if file_path
            else Path(directory).resolve()
            if directory
            else None
        )
        if base is None:
            raise WordMcpError("scope='orphans' needs directory (or file_path)")
        return _orphan_folders(base), None
    raise WordMcpError(
        f"unknown purge scope {scope!r}; use 'legacy', 'orphans', or 'slots'"
    )


def purge_backups(
    scope: str,
    file_path: str | None = None,
    directory: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Delete backups. scope: 'legacy' (*.bak-* files), 'orphans' (slot
    folders whose source document is gone), or 'slots' (one document's
    prev/anchor). dry_run=True (the default) only reports what WOULD be
    deleted; pass dry_run=False to actually delete."""
    targets, slot_folder = _collect_purge_targets(scope, file_path, directory)
    report = []
    for t in targets:
        size = _dir_size(t) if t.is_dir() else t.stat().st_size
        report.append({"path": str(t), "size_bytes": size})
    total = sum(e["size_bytes"] for e in report)
    result = {
        "scope": scope,
        "dry_run": dry_run,
        ("would_delete" if dry_run else "deleted"): report,
        "total_bytes": total,
        "count": len(report),
    }
    if dry_run:
        result["note"] = "nothing was deleted; pass dry_run=False to delete"
        return result

    for t in targets:
        if t.is_dir():
            shutil.rmtree(t, ignore_errors=False)
        else:
            t.unlink()
    # After purging a document's slots, drop its now-empty folder (and the
    # breadcrumb, if any); best-effort, a leftover lockfile just stays.
    if scope == "slots" and slot_folder is not None and slot_folder.is_dir():
        try:
            crumb = slot_folder / safesave._SOURCE_NAME_FILE
            crumb.unlink(missing_ok=True)
            os.rmdir(slot_folder)
        except OSError:
            pass
    return result


# ----------------------------------------------------------------- snapshots

_DTG_PREFIX = re.compile(r"^\d{8}_\d{4}_")
_LABEL_BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def create_snapshot(
    file_path: str,
    *,
    label: str | None = None,
    dest_dir: str | None = None,
) -> dict:
    """DTG-stamped permanent copy of a document: YYYYMMDD_HHMM_<name>.docx
    (an existing leading DTG on the name is replaced, not stacked). These are
    the PERMANENT keepers that complement the automatic prev/anchor slots:
    slots rotate on every mutation, snapshots are never touched by the backup
    system and never auto-pruned. Never overwrites; collisions get a numeric
    suffix. Returns the created path."""
    check_path(file_path, "snapshot document")
    doc = Path(file_path).resolve()
    if not doc.is_file():
        raise DocumentNotFound(f"no document at {doc}")
    if label is not None:
        label = label.strip()
        if not label:
            label = None
        elif _LABEL_BAD.search(label):
            raise WordMcpError(
                "label contains characters not allowed in filenames "
                '(< > : " / \\ | ? * or control characters)'
            )
        elif len(label) > 60:
            raise WordMcpError("label must be 60 characters or fewer")

    if dest_dir:
        check_path(dest_dir, "write snapshot")
    target_dir = Path(dest_dir).resolve() if dest_dir else doc.parent
    if not target_dir.is_dir():
        raise DocumentNotFound(f"no directory at {target_dir}")

    dtg = _dt.datetime.now().strftime("%Y%m%d_%H%M")
    stem = _DTG_PREFIX.sub("", doc.stem)
    base = f"{dtg}_{stem}" + (f"_{label}" if label else "")
    dest = target_dir / f"{base}{doc.suffix}"
    n = 2
    while dest.exists():
        dest = target_dir / f"{base} ({n}){doc.suffix}"
        n += 1

    shutil.copy2(doc, dest)
    result = {"snapshot": str(dest), "source": str(doc), "label": label}
    owner_file = doc.with_name("~$" + doc.name[-153:])
    if owner_file.exists():
        result["note"] = (
            "the document appears to be open in Word; unsaved changes are "
            "NOT in this snapshot (save in Word first for a current copy)"
        )
    return result


# ------------------------------------------------------------------ dispatch


def manage_backups(
    action: str,
    file_path: str | None = None,
    directory: str | None = None,
    source: str | None = None,
    scope: str | None = None,
    dry_run: bool = True,
) -> dict:
    """Single entry point the server tool exposes. action: list|restore|purge."""
    if action == "list":
        return list_backups(file_path=file_path, directory=directory)
    if action == "restore":
        if not file_path:
            raise WordMcpError("restore needs file_path (the document to restore)")
        if not source:
            raise WordMcpError(
                "restore needs source: 'prev', 'anchor', or a legacy backup path"
            )
        return restore_backup(file_path, source)
    if action == "purge":
        if not scope:
            raise WordMcpError(
                "purge needs scope: 'legacy', 'orphans', or 'slots'"
            )
        return purge_backups(
            scope, file_path=file_path, directory=directory, dry_run=dry_run
        )
    raise WordMcpError(
        f"unknown action {action!r}; use 'list', 'restore', or 'purge'"
    )
