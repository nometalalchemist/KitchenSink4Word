"""Cross-PROCESS serialization for live COM sessions.

Why this exists (concurrency matrix, 2026-09-05, findings H1/H2/H4):

``com/serial.py`` holds a process-wide RLock, and it says so honestly: it
serializes THIS server process. The matrix ran two independent server
processes against one Word instance and measured 49 overlapping operation
pairs across 50 operations. Word's own STA serializes individual COM
CALLS, but a tool call is a SEQUENCE of calls -- resolve a target, then
write it -- and nothing kept two processes from interleaving inside one
sequence:

- H1: two live ``search_and_replace`` calls on the same token both ran
  Find, both got a Range over the same characters, and both wrote. Six
  tokens produced twelve claimed replacements and text welded together
  from two edits, with both callers reporting success.
- H2: ``DisplayAlerts`` is APPLICATION state. Process B snapshotted the
  ``wdAlertsNone`` that process A was holding, A restored the real value,
  then B restored ``wdAlertsNone``. Word was left with alerts suppressed
  after both servers were gone, so the user's next close-a-dirty-document
  answered itself and discarded the work.
- H4: an index resolved in one tool call and deleted in the next is a
  CROSS-CALL race that no lock can close; ``expect=`` guards handle that
  one (live_ops), and this lock closes the within-call half.

Design, ported from the Excel sibling's ``core/safesave.write_lock``
(xlsx-mcp commits 31934c2 and 20bb3fe), which learned these lessons the
expensive way:

- **Atomic publish.** The payload is written to a private temp file and
  hard-linked into place, so the lockfile is never observable
  half-written. The old two-syscall create (O_CREAT|O_EXCL, then write)
  let a concurrent acquirer read an EMPTY file, parse no pid, call it
  stale, delete a LIVE lock and take its own; two holders then unlinked
  each other's files and the lock outlived every writer.
- **Per-process-instance token.** A bare PID is not an identity. PIDs are
  recycled, so a process inheriting a dead writer's number read that
  writer's leaked lockfile as its own forever. Re-entrancy is TOKEN
  equality; our own PID under a foreign token is stale by definition.
- **Release only if ours.** Unconditional unlink is how a leaked lock
  outlives every writer.
- **Stale detection via process create time.** xlsx-mcp's PidJournal fix
  (M-1) proved that "is some process holding that number right now" is
  the wrong question: Windows recycles PIDs, and the answer can be the
  user's own application. The lock payload records the holder's process
  creation time (ctypes ``GetProcessTimes``, no psutil dependency), and a
  live PID whose creation time does not match is a recycled number, which
  makes the lock stale.

Scope: APPLICATION, not document. H2 leaked between two processes editing
DIFFERENT documents, because the state they corrupted belongs to
Word.Application, and H1's two callers reach one Application anyway. A
per-document lock would leave H2 open. The cost is that two servers
driving two different open documents now queue instead of overlapping,
which costs nothing real: Word's STA already admitted one COM call at a
time, and the matrix measured a 1191 ms median per live operation under
load either way.

Deliberate deviation from the Excel port: the lockfile lives in a local
per-user directory, NOT beside the document. The live route never writes
the document's folder (it has no backup slots, by design), the target may
sit on OneDrive or a network share where a stray lockfile would sync, and
COM automation is machine-local, so machine-local is exactly the right
scope for the lock.
"""

from __future__ import annotations

import contextlib
import json
import os
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path

from ..core.errors import LiveLockTimeout

#: Directory holding live-session lockfiles (per user, machine-local).
_LOCK_DIR_ENV = "KS4W_LIVE_LOCK_DIR"

#: One lock covers the Word application tier (see the scope note above).
APP_SCOPE = "word-app"

#: How long an acquirer waits for a live holder before refusing. Live COM
#: operations are SLOW -- the matrix measured a 1191 ms median under load,
#: and a 20-mutation batch is one held session -- so the wait is generous.
#: A refusal here must mean "the other server is genuinely stuck", never
#: "the other server was doing ordinary work".
LOCK_WAIT_SECONDS = 120.0

#: A lockfile older than this is broken regardless of PID liveness. No
#: legitimate live session holds Word for ten minutes.
LOCK_STALE_SECONDS = 10 * 60

#: A lockfile we cannot parse is only assumed abandoned after this long
#: (covers a torn write from any source; our own publish is atomic).
_UNREADABLE_GRACE_SECONDS = 30.0

#: Poll interval while waiting for a live holder.
_POLL_SECONDS = 0.05


def _lock_dir() -> Path:
    override = os.environ.get(_LOCK_DIR_ENV)
    if override:
        return Path(override)
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return Path(base) / "word-mcp" / "live-locks"


def _process_create_time(pid: int) -> float | None:
    """Process creation time as a float, or None when it cannot be read.

    Windows only (ctypes GetProcessTimes); returns None elsewhere, which
    downgrades staleness detection to plain PID liveness rather than
    breaking it.
    """
    if pid <= 0 or sys.platform != "win32":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION, False, pid
        )
        if not handle:
            return None
        try:
            creation = wintypes.FILETIME()
            exit_t = wintypes.FILETIME()
            kernel_t = wintypes.FILETIME()
            user_t = wintypes.FILETIME()
            ok = kernel32.GetProcessTimes(
                handle,
                ctypes.byref(creation),
                ctypes.byref(exit_t),
                ctypes.byref(kernel_t),
                ctypes.byref(user_t),
            )
            if not ok:
                return None
            return float(
                (creation.dwHighDateTime << 32) | creation.dwLowDateTime
            )
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return None


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
                return (
                    ctypes.get_last_error() == 5
                    or kernel32.GetLastError() == 5
                )
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


#: Identity of THIS process instance, minted once at import. PID + a random
#: token: re-entrancy is token equality, and a lock carrying our PID under a
#: foreign token is a recycled-PID leftover (stale by definition).
_OWNER_TOKEN = f"{os.getpid()}-{uuid.uuid4().hex}"
_OWNER_CREATED = _process_create_time(os.getpid())

_MUTEXES: dict[str, threading.RLock] = {}
_MUTEX_GUARD = threading.Lock()


def _mutex_for(key: str) -> threading.RLock:
    with _MUTEX_GUARD:
        lock = _MUTEXES.get(key)
        if lock is None:
            lock = _MUTEXES[key] = threading.RLock()
        return lock


def _read_lock_info(lock_path: Path) -> dict:
    try:
        info = json.loads(lock_path.read_text(encoding="utf-8"))
        if isinstance(info, dict):
            return info
    except (OSError, ValueError):
        pass
    return {}


def _publish_lockfile(lock_path: Path, holder: str) -> bool:
    """Make a COMPLETE lockfile appear atomically, or report that one is
    already there. The payload is written to a private temp file first and
    linked into place, so the lockfile is never observable half-written and
    its timestamp is minted at the instant it becomes visible."""
    payload = json.dumps({
        "pid": os.getpid(),
        "token": _OWNER_TOKEN,
        "pid_created": _OWNER_CREATED,
        "time": time.time(),
        "holder": holder,
        "host": os.environ.get("COMPUTERNAME", ""),
    })
    tmp = lock_path.parent / f".lock-{uuid.uuid4().hex}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            fh.write(payload)
        try:
            os.link(str(tmp), str(lock_path))
            return True
        except FileExistsError:
            return False
        except OSError:
            # No hardlink support (non-NTFS, some network shares): fall back
            # to exclusive-create, still writing the payload before the
            # handle closes.
            try:
                fd = os.open(
                    str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY
                )
            except FileExistsError:
                return False
            try:
                os.write(fd, payload.encode("utf-8"))
            finally:
                os.close(fd)
            return True
    finally:
        with contextlib.suppress(OSError):
            tmp.unlink(missing_ok=True)


def _is_ours(info: dict) -> bool:
    return info.get("token") == _OWNER_TOKEN


def _break_lock(lock_path: Path) -> None:
    with contextlib.suppress(OSError):
        lock_path.unlink(missing_ok=True)


def _is_stale(info: dict) -> bool:
    """A lock nobody can still be holding."""
    pid = info.get("pid", -1)
    stamp = info.get("time", 0.0)
    age = time.time() - stamp if isinstance(stamp, (int, float)) else None
    if not isinstance(pid, int) or not _pid_alive(pid):
        return True
    if age is None or age > LOCK_STALE_SECONDS:
        return True
    if pid == os.getpid():
        # Our own PID under a foreign token (checked before this call): the
        # number was recycled and the writer is gone.
        return True
    born = info.get("pid_created")
    now_born = _process_create_time(pid)
    if born is not None and now_born is not None and born != now_born:
        # Same number, different process: Windows recycled the PID. Asking
        # only "is that PID alive" would treat a stranger's process as the
        # live holder and wait the full window for a lock nobody holds.
        return True
    return False


def _acquire_lockfile(lock_path: Path, holder: str, wait: float) -> bool:
    """Create the advisory lockfile. Returns True when this call created it
    (and must therefore remove it); False for a re-entrant same-process
    hold. Raises LiveLockTimeout when a live holder does not release."""
    deadline = time.monotonic() + wait
    while True:
        if _publish_lockfile(lock_path, holder):
            return True
        info = _read_lock_info(lock_path)
        if _is_ours(info):
            # This process instance already holds it; the in-process mutex
            # (held by the caller) is the real serializer and a nested
            # acquisition must not deadlock.
            return False
        if not info:
            # Unparseable: give a torn write a moment to settle, then treat
            # a persistently unreadable lock as abandoned.
            try:
                born = lock_path.stat().st_mtime
            except OSError:
                continue                      # it vanished; try again
            if time.time() - born > _UNREADABLE_GRACE_SECONDS:
                _break_lock(lock_path)
            elif time.monotonic() > deadline:
                raise LiveLockTimeout(
                    "Word is held by a live-session lock this process "
                    f"cannot read. Waited {int(wait)}s; retry, or delete "
                    f"{lock_path} if no other kitchensink4word process is "
                    "running."
                )
            else:
                time.sleep(_POLL_SECONDS)
            continue
        if _is_stale(info):
            _break_lock(lock_path)
            continue
        if time.monotonic() > deadline:
            pid = info.get("pid")
            stamp = info.get("time", 0.0)
            age = (
                time.time() - stamp
                if isinstance(stamp, (int, float))
                else None
            )
            detail = f"PID {pid}"
            if info.get("holder"):
                detail += f", running {info['holder']!r}"
            if age is not None:
                detail += f", held for {int(age)}s"
            raise LiveLockTimeout(
                "another kitchensink4word server process is driving Word "
                f"right now ({detail}). Live COM sessions are serialized "
                "across processes so two servers cannot interleave inside "
                f"one edit. Waited {int(wait)}s; nothing was changed — "
                "retry once that operation finishes."
            )
        time.sleep(_POLL_SECONDS)


def _release_lockfile(lock_path: Path) -> None:
    """Remove the lockfile ONLY while it is still ours. Unconditional
    unlink is how a leaked lock outlives every writer."""
    info = _read_lock_info(lock_path)
    if info and not _is_ours(info):
        return
    _break_lock(lock_path)


def holder_info(scope: str = APP_SCOPE) -> dict | None:
    """Who holds the cross-process live lock right now, or None. Read-only;
    used by status reporting so a caller can see WHY it would queue."""
    lock_path = _lock_dir() / f"{scope}.lock"
    info = _read_lock_info(lock_path)
    if not info or _is_stale(info):
        return None
    return {
        "pid": info.get("pid"),
        "holder": info.get("holder"),
        "held_for_s": round(time.time() - info.get("time", time.time()), 1),
        "ours": _is_ours(info),
    }


@contextlib.contextmanager
def cross_process_lock(
    holder: str,
    *,
    scope: str = APP_SCOPE,
    wait: float = LOCK_WAIT_SECONDS,
):
    """Serialize a whole live COM session against every other server process.

    Hold this around the FULL attach-resolve-probe-mutate-restore sequence.
    Yields True when this call published the lockfile, False when the hold
    is re-entrant within this process instance (the in-process mutex is the
    real serializer there).

    Degrades to the in-process mutex alone when the lock directory cannot
    be created; it never fails a live edit because a lockfile could not be
    hosted, and it never silently claims cross-process coverage it does not
    have (see ``lock_state``).
    """
    mutex = _mutex_for(scope)
    mutex.acquire()
    owns = False
    lock_path: Path | None = None
    try:
        try:
            d = _lock_dir()
            d.mkdir(parents=True, exist_ok=True)
            lock_path = d / f"{scope}.lock"
        except Exception:
            # Broad on purpose: a lock directory that cannot be hosted must
            # never be the reason a live edit fails. The in-process mutex
            # still serializes this process, and lock_state() reports the
            # reduced coverage rather than letting a caller assume it.
            lock_path = None
        if lock_path is not None:
            owns = _acquire_lockfile(lock_path, holder, wait)
        yield owns
    finally:
        if owns and lock_path is not None:
            _release_lockfile(lock_path)
        mutex.release()


def lock_state() -> dict:
    """Honest description of the coverage this module is providing, for
    status output: cross-process when a lockfile can be hosted, in-process
    only when it cannot."""
    try:
        d = _lock_dir()
        d.mkdir(parents=True, exist_ok=True)
        return {"cross_process": True, "lock_dir": str(d)}
    except Exception as exc:
        return {"cross_process": False, "reason": str(exc)}
