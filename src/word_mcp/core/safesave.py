"""Slot-based backup rotation + per-file write serialization for the save path.

Backup model (v1.6 redesign, replacing per-mutation ``*.bak-*`` accumulation):

- A hidden ``.ks4w-backups/`` folder sits next to each mutated document
  (dot-prefix everywhere; FILE_ATTRIBUTE_HIDDEN is additionally set on
  Windows, non-fatal if that fails).
- Inside it, one subfolder per document (the document's own file name,
  hash-suffixed when the name is very long; Korean/unicode names work).
- Exactly two stable slots per document (SLOT_POLICY below):
    ``prev.docx``   - state before the most recent mutation, rotates every call
    ``anchor.docx`` - session-start state, rotates only when the document has
                      been idle ANCHOR_IDLE_SECONDS or more, measured from the
                      prev slot's mtime (no state database).
- Rotation mechanism (crash-window-free): HARDLINK the current target file to
  a temp name inside the slot folder, then os.replace that link onto the slot.
  The document is NEVER absent from its own path at any instant; the old
  content survives the caller's final target replace under the slot link.
  Hardlink failure (non-NTFS, cloud placeholders, cross-volume) falls back to
  shutil.copy2. Link/replace calls are retried with backoff to ride out
  transient AV sharing violations.

Write serialization (fixes the parallel read-modify-save race):

- ``write_lock(path)`` context manager, held across the FULL
  read-modify-validate-save cycle of a mutation.
- In-process: one threading.RLock per file, keyed on
  normcase(realpath(path)).
- Cross-process: an advisory lockfile inside the document's slot folder
  carrying PID + timestamp. Stale locks (dead PID, or older than
  LOCK_STALE_SECONDS) are broken; otherwise acquisition waits up to
  LOCK_WAIT_SECONDS and then refuses with MutationLockTimeout naming the
  holder. Two server processes on one machine is the normal case.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from .errors import WordMcpError
from .sandbox import check_path

# ---------------------------------------------------------------- constants

BACKUP_DIR_NAME = ".ks4w-backups"

#: Slot policy: stable slot file names, at most these per document ever.
#: (User-chosen 2-slot design; a third prior-anchor slot was considered and
#: declined - see V1.6 kickoff.)
SLOT_POLICY: tuple[str, ...] = ("prev.docx", "anchor.docx")
PREV_SLOT = "prev.docx"
ANCHOR_SLOT = "anchor.docx"

#: Idle gap (seconds) after which the next mutation is a "new session" and
#: the anchor slot rotates to the current pre-mutation state.
ANCHOR_IDLE_SECONDS = 60 * 60

#: Advisory lockfile name inside a document's slot folder.
LOCK_FILE_NAME = "write.lock"
#: How long an acquirer waits for a live holder before refusing.
LOCK_WAIT_SECONDS = 10.0
#: A lockfile older than this is broken regardless of PID liveness
#: (generous: no single mutation legitimately holds the lock this long).
LOCK_STALE_SECONDS = 10 * 60

#: Slot folder names longer than this are truncated + hash-suffixed
#: (keeps total path length sane for Windows even without long-path opt-in).
_MAX_FOLDER_NAME = 80

#: Breadcrumb written when the folder name had to be truncated, so orphan
#: detection can still map the folder back to its source document name.
_SOURCE_NAME_FILE = "source-name.txt"

_RETRY_DELAYS = (0.05, 0.1, 0.2, 0.4)  # seconds, between retries
_TRANSIENT_WINERRORS = {5, 32, 33}  # access denied / sharing violations


class MutationLockTimeout(WordMcpError):
    """Another process holds the write lock on this file and did not release
    it within the wait window; the mutation was refused, nothing changed."""


# ------------------------------------------------------------ path plumbing


def canonical_key(path: str | os.PathLike) -> str:
    """Stable per-file identity: normcase(realpath). Lock keys use this."""
    return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(path))))


def _folder_name_for(doc_name: str) -> str:
    """Slot subfolder name for a document file name. The document's own name
    when it fits (trivial reverse mapping); truncated + 8-hex-hash when long."""
    if len(doc_name) <= _MAX_FOLDER_NAME:
        return doc_name
    digest = hashlib.sha1(
        os.path.normcase(doc_name).encode("utf-8", "surrogatepass")
    ).hexdigest()[:8]
    return f"{doc_name[: _MAX_FOLDER_NAME - 9]}-{digest}"


def backup_root(doc_path: str | os.PathLike) -> Path:
    """The .ks4w-backups folder next to a document (not created)."""
    return Path(doc_path).resolve().parent / BACKUP_DIR_NAME


def slot_dir(doc_path: str | os.PathLike, *, create: bool = False) -> Path:
    """This document's slot folder inside .ks4w-backups (created on demand)."""
    p = Path(doc_path).resolve()
    root = p.parent / BACKUP_DIR_NAME
    folder = _folder_name_for(p.name)
    d = root / folder
    if create:
        d.mkdir(parents=True, exist_ok=True)
        _hide(root)  # dot-prefix everywhere + real hidden bit on Windows
        if folder != p.name:
            crumb = d / _SOURCE_NAME_FILE
            if not crumb.exists():
                try:
                    crumb.write_text(p.name, encoding="utf-8")
                except OSError:
                    pass  # breadcrumb is best-effort
    return d


def source_doc_for(folder: Path) -> Path:
    """Map a slot folder back to its source document path (may not exist)."""
    name = folder.name
    crumb = folder / _SOURCE_NAME_FILE
    if crumb.exists():
        try:
            name = crumb.read_text(encoding="utf-8").strip() or folder.name
        except OSError:
            pass
    return folder.parent.parent / name


def _hide(path: Path) -> None:
    """Set FILE_ATTRIBUTE_HIDDEN on Windows. Non-fatal on any failure."""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        target = str(path)
        if len(target) > 240 and not target.startswith("\\\\?\\"):
            target = "\\\\?\\" + target
        FILE_ATTRIBUTE_HIDDEN = 0x2
        attrs = ctypes.windll.kernel32.GetFileAttributesW(target)
        if attrs != -1 and not (attrs & FILE_ATTRIBUTE_HIDDEN):
            ctypes.windll.kernel32.SetFileAttributesW(
                target, attrs | FILE_ATTRIBUTE_HIDDEN
            )
    except Exception:
        pass  # hidden bit is cosmetic; the dot-prefix stands on its own


# ------------------------------------------------------- retry / placement


def _is_transient(exc: OSError) -> bool:
    winerr = getattr(exc, "winerror", None)
    return isinstance(exc, PermissionError) or winerr in _TRANSIENT_WINERRORS


def _with_retry(fn, *args):
    """Run an os-level file op, retrying briefly on AV sharing violations."""
    for delay in _RETRY_DELAYS:
        try:
            return fn(*args)
        except OSError as exc:
            if not _is_transient(exc):
                raise
            time.sleep(delay)
    return fn(*args)  # final attempt, exceptions propagate


def replace_with_retry(src: str | os.PathLike, dst: str | os.PathLike) -> None:
    """os.replace with transient-error backoff (used by the save path too)."""
    _with_retry(os.replace, os.fspath(src), os.fspath(dst))


def _link_or_copy(src: Path, dst: Path) -> None:
    """Hardlink src as dst; copy2 fallback for non-NTFS / cloud placeholders /
    anything the link call rejects. copy2 preserves mtime, which the anchor
    idle measurement relies on."""
    try:
        _with_retry(os.link, str(src), str(dst))
    except OSError:
        _with_retry(shutil.copy2, str(src), str(dst))


def _place_onto_slot(current: Path, slot_path: Path) -> None:
    """Capture `current`'s present content into a slot without the document
    ever leaving its own path: link (or copy) to a temp name in the slot
    folder, then atomically replace the slot."""
    tmp = slot_path.parent / f".slot-{uuid.uuid4().hex}.tmp"
    try:
        _link_or_copy(current, tmp)
        replace_with_retry(tmp, slot_path)
    except BaseException:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


# ------------------------------------------------------------- slot rotation


def rotate_slots(doc_path: str | os.PathLike) -> dict:
    """Rotate backup slots for `doc_path` BEFORE its content is replaced.

    Called by the save path after the new payload is validated and written to
    a temp file, immediately before the final os.replace onto the target.
    Captures the CURRENT (pre-mutation) content:

    - anchor.docx: created if absent (first mutation ever / after purge), or
      rotated when the prev slot says the document has been idle 60+ minutes
      (new session).
    - prev.docx: rotated on every call.

    Never unlinks or renames the document itself.
    """
    doc = Path(doc_path).resolve()
    d = slot_dir(doc, create=True)
    prev = d / PREV_SLOT
    anchor = d / ANCHOR_SLOT
    rotated = {"prev": True, "anchor": False}

    if not anchor.exists():
        _place_onto_slot(doc, anchor)
        rotated["anchor"] = True
    elif prev.exists():
        try:
            idle = time.time() - prev.stat().st_mtime
        except OSError:
            idle = 0.0
        if idle >= ANCHOR_IDLE_SECONDS:
            _place_onto_slot(doc, anchor)
            rotated["anchor"] = True

    _place_onto_slot(doc, prev)
    return rotated


# ------------------------------------------------------------------ locking

_MUTEXES: dict[str, threading.RLock] = {}
_MUTEX_GUARD = threading.Lock()


def _mutex_for(key: str) -> threading.RLock:
    with _MUTEX_GUARD:
        lock = _MUTEXES.get(key)
        if lock is None:
            lock = _MUTEXES[key] = threading.RLock()
        return lock


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        # NEVER os.kill(pid, 0) on Windows - it terminates the process.
        try:
            import ctypes

            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION, False, pid
            )
            if not handle:
                # Access denied means it exists (another user's process).
                return ctypes.get_last_error() == 5 or kernel32.GetLastError() == 5
            try:
                code = ctypes.c_ulong()
                if kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                    return code.value == STILL_ACTIVE
                return True  # could not query; assume alive (conservative)
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return True  # cannot tell; assume alive (never break a live lock)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except OSError:
        return True


def _read_lock_info(lock_path: Path) -> dict:
    try:
        info = json.loads(lock_path.read_text(encoding="utf-8"))
        if isinstance(info, dict):
            return info
    except (OSError, ValueError):
        pass
    return {}


def _acquire_lockfile(lock_path: Path, doc_name: str) -> bool:
    """Create the advisory lockfile. Returns True when this call created it
    (and must therefore remove it); False for a re-entrant same-process hold.
    Raises MutationLockTimeout when a live holder does not release in time."""
    deadline = time.monotonic() + LOCK_WAIT_SECONDS
    payload = json.dumps(
        {"pid": os.getpid(), "time": time.time(), "host": os.environ.get("COMPUTERNAME", "")}
    )
    while True:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, payload.encode("utf-8"))
            finally:
                os.close(fd)
            return True
        except FileExistsError:
            info = _read_lock_info(lock_path)
            pid = info.get("pid", -1)
            stamp = info.get("time", 0.0)
            age = time.time() - stamp if isinstance(stamp, (int, float)) else None
            if pid == os.getpid():
                # Same process: the in-process mutex (already held) is the
                # real serializer; a nested acquisition must not deadlock.
                return False
            stale = (
                not isinstance(pid, int)
                or not _pid_alive(pid)
                or age is None
                or age > LOCK_STALE_SECONDS
            )
            if stale:
                try:
                    lock_path.unlink(missing_ok=True)
                except OSError:
                    pass
                continue
            if time.monotonic() > deadline:
                holder = f"PID {pid}"
                if age is not None:
                    holder += f", held for {int(age)}s"
                raise MutationLockTimeout(
                    f"{doc_name} is being modified by another kitchensink4word "
                    f"process ({holder}). Waited {int(LOCK_WAIT_SECONDS)}s; "
                    "retry once that operation finishes."
                )
            time.sleep(0.1)


@contextmanager
def write_lock(doc_path: str | os.PathLike):
    """Serialize the full read-modify-validate-save cycle of one document.

    In-process mutex (per resolved path) + cross-process advisory lockfile in
    the document's slot folder. Hold this around DocxPackage(...) through
    pkg.save() so every mutation sees the previous one's result and response
    metadata (paragraph indices etc.) is computed against settled state.
    """
    # Public entry that takes an arbitrary path (server.py copy_document and
    # ops/batch call it directly); slot/backup paths all derive from doc_path,
    # so gating it here covers them too. No-op unless KS4W_ALLOWED_ROOTS is set.
    check_path(doc_path, "modify document")
    key = canonical_key(doc_path)
    mutex = _mutex_for(key)
    mutex.acquire()
    owns_lockfile = False
    lock_path: Path | None = None
    try:
        try:
            d = slot_dir(doc_path, create=True)
            lock_path = d / LOCK_FILE_NAME
        except OSError:
            lock_path = None  # cannot host a lockfile; mutex still serializes
        if lock_path is not None:
            owns_lockfile = _acquire_lockfile(lock_path, Path(doc_path).name)
        yield
    finally:
        if owns_lockfile and lock_path is not None:
            try:
                lock_path.unlink(missing_ok=True)
            except OSError:
                pass
        mutex.release()
