"""Live COM layer: edit documents while they are OPEN in the user's Word.

Contract (the inverse of bridge.py, which owns invisible instances):
- NEVER creates a Word instance, never quits, never closes or saves a
  document without an explicit flag. The user's unsaved work is sacred.
- Attaches per call via GetActiveObject (ROT-scan fallback for multi-instance
  setups); no COM pointer is ever cached across tool calls — STA pointers are
  thread-affine and FastMCP does not pin tool calls to one thread.
- All content addressing goes through Range objects. Selection is never
  written and (outside live_ops' explicit cursor tools) never read, so the
  user's cursor and scroll position are untouched.
- Every tool call becomes ONE entry in the user's Ctrl+Z stack via
  Application.UndoRecord; on failure the record still closes in finally.
- Any app/doc state a tool mutates goes through StateGuard and is restored
  LIFO in finally; restore failures are reported, never masked.
"""

from __future__ import annotations

import contextlib
import threading
import time
from pathlib import Path

from ..core.errors import (
    DocumentNotOpenInWord,
    DocumentProtected,
    ProtectedViewRefused,
    WordBlocked,
    WordBusy,
    WordDisconnected,
    WordMcpError,
    WordNotRunning,
)
from . import serial as _serial

# HRESULTs (as signed ints, the way pywin32 surfaces them)
RPC_E_CALL_REJECTED = -2147418111        # modal dialog / call rejected
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # busy, retry later
RPC_E_DISCONNECTED = -2147417848         # server died under a live proxy
CO_E_OBJNOTCONNECTED = -2147220995       # proxy no longer connected
MK_E_UNAVAILABLE = -2147221021           # nothing in the ROT
RPC_S_CALL_FAILED = -2147023170          # 0x800706BE — Word killed mid-call
RPC_S_SERVER_UNAVAILABLE = -2147023174   # 0x800706BA — RPC server gone

BUSY_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
GONE_HRESULTS = {
    RPC_E_DISCONNECTED,
    CO_E_OBJNOTCONNECTED,
    RPC_S_CALL_FAILED,
    RPC_S_SERVER_UNAVAILABLE,
}
# Word's "this is protected / not allowed to edit" command errors
PROTECTED_HRESULTS = {-2146823683, -2146822164}

# wdProtectionType values
_WD_NO_PROTECTION = -1
_WD_ALLOW_ONLY_REVISIONS = 0
_PROTECTION_NAMES = {
    0: "trackedChanges",
    1: "comments",
    2: "formFields",
    3: "readOnly",
}

_WD_ALERTS_ALL = -1
_WD_ALERTS_NONE = 0

# Word COM rejects single string arguments much beyond ~32K; stay well under.
TEXT_CHUNK = 30000
# The table-cell separator; inserting it as text corrupts table structure.
CELL_SEPARATOR = "\x07"


def _com_modules():
    try:
        import pythoncom
        import pywintypes
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise WordMcpError("pywin32 is not available; live editing needs it") from exc
    return pythoncom, pywintypes, win32com.client


def _hresults(exc) -> set:
    """Every HRESULT pywin32 might have stashed in a com_error."""
    out = set()
    hr = getattr(exc, "hresult", None)
    if hr is not None:
        out.add(hr)
    args = getattr(exc, "args", ())
    if len(args) >= 3 and args[2]:
        with contextlib.suppress(Exception):
            scode = args[2][5]
            if scode is not None:
                out.add(scode)
    return out


def _classify(exc):
    """Map a com_error to our typed errors, or return None if unrecognized."""
    hrs = _hresults(exc)
    if hrs & GONE_HRESULTS:
        return WordDisconnected(
            "Word or the document closed while the tool was running — the edit "
            "may be partially applied. If Word is still open, one Ctrl+Z undoes "
            "the partial step."
        )
    if hrs & BUSY_HRESULTS:
        return WordBusy(
            "Word is busy or has a dialog open (a dialog box, the File "
            "menu/Backstage, or a running command). Close it and retry."
        )
    if hrs & PROTECTED_HRESULTS:
        return DocumentProtected(
            "Word refused the edit because the document's editing "
            "restrictions forbid it; remove the protection in Word (or with "
            "remove_document_protection on the closed file) and retry."
        )
    return None


def _ensure_com(pythoncom):
    """COM apartment handling: initialize freely, NEVER uninitialize.

    pywin32's CoInitialize is a no-op on an already-initialized thread
    while CoUninitialize ALWAYS decrements, so a paired call on a
    host-initialized thread destroys the host's apartment and disconnects
    every COM proxy it holds (verified empirically 2026-08-28). Calling
    CoInitialize unconditionally is safe (no-op when alive) and self-heals
    an apartment some OTHER code tore down; skipping CoUninitialize means
    this module can never be the one that kills the thread's COM state.
    The apartment persists for the thread's lifetime — intended."""
    pythoncom.CoInitialize()


class StateGuard:
    """Snapshot-on-mutate for interactive-instance state; LIFO restore.

    Tools must change app/doc state ONLY through set(); restore() runs in the
    session finally and reports (not raises) anything it could not put back.
    """

    def __init__(self):
        self._stack = []

    def set(self, obj, attr, value):
        self._stack.append((obj, attr, getattr(obj, attr)))
        setattr(obj, attr, value)

    def restore(self) -> list:
        failed = []
        for obj, attr, saved in reversed(self._stack):
            try:
                setattr(obj, attr, saved)
            except Exception:
                failed.append(attr)
        self._stack.clear()
        return failed


@contextlib.contextmanager
def undo_group(app, name: str, doc=None):
    """Group everything inside into one Ctrl+Z step. Yields whether grouping
    is active; degrades HONESTLY (undo_grouped=false) rather than failing
    the edit or lying.

    Active-document constraint (L5 finding, 2026-08-28):
    Application.UndoRecord only records into the ACTIVE document. When the
    target doc is open but another doc is foremost, a custom record groups
    NOTHING while claiming success — and activating the target would steal
    the user's window, which is banned. So grouping is only attempted when
    the target IS the active document; otherwise the edits are individually
    undoable and the result says so.

    Crash reality (verified 2026-08-28): if a client dies mid-record, Word
    does NOT auto-end the record on this build — the user's own edits keep
    merging into the orphan until something ends it. We therefore close any
    orphaned record before starting ours, and word_live_repair also closes
    them."""
    undo = app.UndoRecord
    started = False
    if doc is not None:
        try:
            if app.ActiveDocument.FullName.lower() != doc.FullName.lower():
                yield False
                return
        except Exception:
            yield False
            return
    try:
        # orphans from crashed clients NEST (CustomRecordLevel grows with
        # every new Start under them) — drain ALL open levels, bounded
        for _ in range(16):
            if not undo.IsRecordingCustomRecord:
                break
            undo.EndCustomRecord()
    except Exception:
        pass
    try:
        undo.StartCustomRecord(("word-mcp: " + name)[:64])
        started = True
    except Exception:
        pass
    try:
        yield started
    finally:
        if started:
            with contextlib.suppress(Exception):
                if undo.IsRecordingCustomRecord:
                    undo.EndCustomRecord()


def _winword_running() -> bool:
    import subprocess

    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=15,
        )
    except Exception:
        return False
    return "WINWORD.EXE" in result.stdout.upper()


def _attach_app(win32com, pythoncom):
    try:
        return win32com.GetActiveObject("Word.Application")
    except Exception as exc:
        if _winword_running():
            raise WordNotRunning(
                "Word is running but not attachable (it may have just "
                "launched and not yet registered, or one side is elevated). "
                "Click into another window once, then retry."
            ) from exc
        raise WordNotRunning(
            "Word is not running; live tools need the document open in Word. "
            "For closed files use the regular file-based tools."
        ) from exc


def _find_doc_via_rot(pythoncom, win32com, target_lower: str):
    """Multi-instance fallback: open documents register their full path as a
    file moniker; bind to the Document directly and reach its Application."""
    rot = pythoncom.GetRunningObjectTable()
    for moniker in rot.EnumRunning():
        ctx = pythoncom.CreateBindCtx(0)
        try:
            name = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue
        if name.lower() == target_lower:
            with contextlib.suppress(Exception):
                obj = rot.GetObject(moniker)
                doc = win32com.Dispatch(
                    obj.QueryInterface(pythoncom.IID_IDispatch)
                )
                return doc.Application, doc
    return None, None


def _resolve_document(pythoncom, pywintypes, win32com, app, path: str):
    target = str(Path(path).resolve())
    target_lower = target.lower()
    open_names = []
    primary_error = None
    try:
        for doc in app.Documents:
            full = doc.FullName
            open_names.append(full)
            if full.lower() == target_lower:
                return app, doc
    except pywintypes.com_error as exc:
        # The GetActiveObject instance may be busy or mid-shutdown while a
        # DIFFERENT instance holds the target — fall through to the ROT
        # scan before giving up on it.
        primary_error = _classify(exc) or exc
    other_app, other_doc = _find_doc_via_rot(pythoncom, win32com, target_lower)
    if other_doc is not None:
        return other_app, other_doc
    if primary_error is not None:
        # the target was nowhere else and the primary instance is unusable
        if isinstance(primary_error, WordMcpError):
            raise primary_error
        raise WordBusy(
            "Word did not answer while resolving the document; retry shortly"
        ) from primary_error
    pv_hit = False
    with contextlib.suppress(Exception):
        for i in range(1, app.ProtectedViewWindows.Count + 1):
            pv = app.ProtectedViewWindows(i).Document.FullName
            if pv.lower() == target_lower:
                pv_hit = True
                break
    if pv_hit:
        raise ProtectedViewRefused(
            f"{Path(path).name} is open in Protected View; click "
            "'Enable Editing' in Word, then retry."
        )
    hint = f" Open documents: {open_names}" if open_names else ""
    raise DocumentNotOpenInWord(
        f"{Path(path).name} is not open in the running Word instance — live "
        f"tools only work on open documents.{hint} For closed files use the "
        "regular file-based tools."
    )


def probe_ready(pywintypes, app, doc, retries: int = 3, delay: float = 0.25):
    """Cheap round-trip into Word's STA; refuse BEFORE any mutation."""
    for attempt in range(retries):
        try:
            _ = app.Name
            _ = doc.Name
            return
        except pywintypes.com_error as exc:
            typed = _classify(exc)
            if isinstance(typed, WordDisconnected):
                raise typed from exc
            if isinstance(typed, WordBusy) and attempt < retries - 1:
                time.sleep(delay)
                continue
            raise typed or exc from exc


def probe_with_timeout(timeout: float = 5.0) -> str:
    """'ready' | 'busy' | 'blocked' | 'not_running' — fresh attach on a helper
    thread, so a Word stuck in a long synchronous op cannot hang the server."""
    result = {}

    def _worker():
        pythoncom, pywintypes, win32com = _com_modules()
        pythoncom.CoInitialize()
        try:
            app = win32com.GetActiveObject("Word.Application")
            _ = app.Name
            result["state"] = "ready"
        except pywintypes.com_error as exc:
            if _hresults(exc) & BUSY_HRESULTS:
                result["state"] = "busy"
            else:
                result["state"] = "not_running"
        except Exception:
            result["state"] = "not_running"
        finally:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    t.join(timeout)
    return result.get("state", "blocked")


def _screen_repair_async():
    """Last-resort daemon: fresh attach can always restore ScreenUpdating."""

    def _worker():
        pythoncom, _, win32com = _com_modules()
        for _ in range(3):
            time.sleep(1.0)
            if not _serial.acquire(timeout=10.0):
                continue
            pythoncom.CoInitialize()
            try:
                app = win32com.GetActiveObject("Word.Application")
                app.ScreenUpdating = True
                app.ScreenRefresh()
                return
            except Exception:
                pass
            finally:
                with contextlib.suppress(Exception):
                    pythoncom.CoUninitialize()
                _serial.release()

    threading.Thread(target=_worker, daemon=True).start()


class LiveSession:
    """What a tool body receives: the attached objects plus result plumbing."""

    def __init__(self, app, doc, guard: StateGuard, undo_grouped: bool):
        self.app = app
        self.doc = doc
        self.guard = guard
        self.undo_grouped = undo_grouped
        self.had_unsaved_user_changes = None
        self.screen_toggled = False
        self.state_restore_failed: list = []
        self.enforced_tracking = False   # trackedChanges protection active
        # False = Word said "not read-only"; True = it said "read-only";
        # None = it did not answer, and read_only_probe_failed says why
        self.opened_read_only = False
        self.read_only_probe_failed = None

    def batch_screen_off(self):
        """Only for >20-mutation batches; auto-restored, watchdog-backed."""
        self.guard.set(self.app, "ScreenUpdating", False)
        self.screen_toggled = True

    def result(self, payload: dict) -> dict:
        out = dict(payload)
        out["live"] = True
        out["undo_grouped"] = self.undo_grouped
        if not self.undo_grouped:
            out["undo_note"] = (
                "edits are individually undoable — one-step undo grouping "
                "only works when the target is Word's active document"
            )
        if self.enforced_tracking:
            out["enforced_tracking"] = True
        if self.opened_read_only:
            out["opened_read_only"] = True
        if self.read_only_probe_failed:
            # never report a confident False for a question Word refused to
            # answer: say the probe failed and hand back the reason
            out["opened_read_only"] = None
            out["read_only_probe_failed"] = self.read_only_probe_failed
        with contextlib.suppress(Exception):
            out["document_dirty"] = not self.doc.Saved
        try:
            out["autosave_on"] = bool(self.doc.AutoSaveOn)
        except Exception:
            out["autosave_on"] = None
        if self.had_unsaved_user_changes is not None:
            out["had_unsaved_user_changes"] = self.had_unsaved_user_changes
        return out


def _check_protection(doc, path: str, session: "LiveSession"):
    """Mutating tools: refuse restricted docs up front — Word otherwise
    either raises mid-edit or (for some formatting writes) silently ignores
    the write while we report success. trackedChanges protection allows
    edits but forces them tracked; that is flagged, not refused."""
    try:
        ptype = doc.ProtectionType
    except Exception:
        return
    if ptype == _WD_NO_PROTECTION:
        return
    if ptype == _WD_ALLOW_ONLY_REVISIONS:
        session.enforced_tracking = True
        return
    mode = _PROTECTION_NAMES.get(ptype, str(ptype))
    raise DocumentProtected(
        f"{Path(path).name} has editing restrictions (mode: {mode}); Word "
        "ignores or rejects programmatic edits under this protection. "
        "Remove the protection in Word (Review > Restrict Editing) or use "
        "remove_document_protection on the closed file."
    )


def probe_read_only(session: "LiveSession", doc) -> None:
    """Ask the document whether it opened read-only, and RECORD a refusal.

    This probe used to sit under a bare contextlib.suppress, so a COM fault
    on doc.ReadOnly (a wedged instance, a proxy that answers some properties
    and not others, a late-bound getattr that surfaces as AttributeError)
    became a confident opened_read_only=False. That is the one failure mode
    a read-only flag must not have: the caller edits a document it was told
    was writable and cannot tell that the check never ran. On a fault the
    session records None plus the reason, and the result carries
    read_only_probe_failed."""
    try:
        session.opened_read_only = bool(doc.ReadOnly)
    except Exception as exc:
        session.opened_read_only = None
        session.read_only_probe_failed = f"{type(exc).__name__}: {exc}"


@contextlib.contextmanager
def live_session(path: str, tool_name: str, *, mutating: bool = True):
    """attach → resolve → probe → guard+undo → yield → restore, per call.

    The WHOLE session runs under the process-wide COM serialization lock
    (com/serial.py): exactly one tool call reaches Word at a time, which
    is what makes concurrent multi-agent live editing safe (2026-09-03
    stress report — unserialized Range writes interleave per character).
    Alerts are suppressed (DisplayAlerts = wdAlertsNone) for the session
    and restored by the StateGuard, so Word cannot raise a modal dialog
    mid-operation; live_repair restores alerts if a client crashes."""
    pythoncom, pywintypes, win32com = _com_modules()
    _ensure_com(pythoncom)
    with _serial.com_operation(f"live:{tool_name}"):
        app = doc = None
        guard = StateGuard()
        restore_failed: list = []
        try:
            app = _attach_app(win32com, pythoncom)
            app, doc = _resolve_document(
                pythoncom, pywintypes, win32com, app, path
            )
            probe_ready(pywintypes, app, doc)
            with contextlib.suppress(Exception):
                guard.set(app, "DisplayAlerts", _WD_ALERTS_NONE)
            with undo_group(app, tool_name, doc=doc) as grouped:
                session = LiveSession(app, doc, guard, grouped)
                with contextlib.suppress(Exception):
                    session.had_unsaved_user_changes = not doc.Saved
                probe_read_only(session, doc)
                if mutating:
                    _check_protection(doc, path, session)
                try:
                    yield session
                except pywintypes.com_error as exc:
                    typed = _classify(exc)
                    if typed:
                        raise typed from exc
                    raise
                finally:
                    restore_failed = guard.restore()
                    session.state_restore_failed = restore_failed
                    if restore_failed and session.screen_toggled:
                        _screen_repair_async()
        finally:
            # release proxies promptly; the thread's apartment deliberately
            # stays initialized (see _ensure_com)
            doc = None
            app = None


def run_live(path: str, tool_name: str, body, *, mutating: bool = True) -> dict:
    """Run body(session) inside a live session; the session's post-restore
    report (state_restore_failed) is merged into the returned result.
    mutating=False skips the protection refusal (reads work everywhere)."""
    with live_session(path, tool_name, mutating=mutating) as session:
        result = session.result(body(session))
    result["state_restore_failed"] = session.state_restore_failed
    return result


def check_text_safe(text: str):
    if CELL_SEPARATOR in text:
        raise WordMcpError(
            "text contains \\x07 (Word's internal table-cell separator); "
            "inserting it would corrupt table structure"
        )


def insert_text_chunked(rng, text: str, *, before: bool = False):
    """Insert respecting Word COM's ~32K per-call string limit.

    InsertAfter extends the range to cover what was inserted, so sequential
    InsertAfter calls on the same range append in order.
    """
    check_text_safe(text)
    chunks = [text[i : i + TEXT_CHUNK] for i in range(0, len(text), TEXT_CHUNK)]
    if before:
        for chunk in reversed(chunks):
            rng.InsertBefore(chunk)
    else:
        for chunk in chunks:
            rng.InsertAfter(chunk)


def live_repair() -> dict:
    """Recovery tool: a fresh attach can always fix a Word left frozen or
    mid-record by a crashed client. Restores ScreenUpdating and alerts and
    closes orphaned undo records; it CANNOT know a crashed client's
    pre-crash TrackRevisions value, so that one is reported, not reverted."""
    pythoncom, pywintypes, win32com = _com_modules()
    _ensure_com(pythoncom)
    with _serial.com_operation("live_repair"):
        return _live_repair_locked(pythoncom, win32com)


def _live_repair_locked(pythoncom, win32com) -> dict:
    app = None
    actions = []
    try:
        try:
            app = win32com.GetActiveObject("Word.Application")
        except Exception:
            return {"word_running": False, "actions": []}
        with contextlib.suppress(Exception):
            if app.ScreenUpdating is False:
                app.ScreenUpdating = True
                app.ScreenRefresh()
                actions.append("screen_updating_restored")
        with contextlib.suppress(Exception):
            if app.DisplayAlerts != _WD_ALERTS_ALL:
                app.DisplayAlerts = _WD_ALERTS_ALL
                actions.append("display_alerts_restored")
        with contextlib.suppress(Exception):
            undo = app.UndoRecord
            closed = 0
            for _ in range(16):
                if not undo.IsRecordingCustomRecord:
                    break
                undo.EndCustomRecord()
                closed += 1
            if closed:
                actions.append(
                    f"orphaned_undo_record_closed (x{closed} nested levels)"
                )
        with contextlib.suppress(Exception):
            for doc in app.Documents:
                if doc.TrackRevisions:
                    actions.append(
                        f"track_revisions_still_on:{doc.Name} (a crashed "
                        "client's pre-crash value is unknowable — turn it "
                        "off with live_set_track_changes if unwanted)"
                    )
        return {"word_running": True, "actions": actions}
    finally:
        app = None


def _doc_entry(doc) -> dict:
    entry = {"path": doc.FullName}
    with contextlib.suppress(Exception):
        entry["dirty"] = not doc.Saved
    try:
        entry["autosave_on"] = bool(doc.AutoSaveOn)
    except Exception:
        entry["autosave_on"] = None
    return entry


def interactive_status() -> dict:
    """Extended word_status: responsiveness + per-document dirty/autosave.

    Documents are collected from the primary (GetActiveObject) instance AND
    from every other Word instance found via the running-object table, so
    multi-instance setups report completely.

    Contention honesty (stress report item 7): the result always carries
    com_serialization (lock_snapshot). When another COM operation holds
    the serialization lock, this tool does NOT queue behind it — it
    reports interactive_state "serving" with the running operation's name
    so callers can back off instead of piling on. The probe itself runs
    lock-free on a helper thread (read-only; a stuck Word cannot hang the
    server). Limits: Word does not expose modal-dialog state directly —
    a dialog usually surfaces as "busy" (RPC call rejected), but simple
    property reads can still succeed with a dialog pending, so "ready"
    is evidence, not proof, of a dialog-free Word."""
    if not _serial.acquire(timeout=2.0):
        return {
            "interactive_state": "serving",
            "open_documents": [],
            "com_serialization": _serial.lock_snapshot(),
            "note": (
                "another COM operation holds the serialization lock; "
                "Word state was not probed to avoid queuing — retry "
                "after the running operation finishes"
            ),
        }
    try:
        out = _interactive_status_locked()
    finally:
        _serial.release()
    out["com_serialization"] = _serial.lock_snapshot()
    return out


def _interactive_status_locked() -> dict:
    state = probe_with_timeout()
    out = {"interactive_state": state, "open_documents": []}
    if state != "ready":
        return out
    pythoncom, pywintypes, win32com = _com_modules()
    _ensure_com(pythoncom)
    app = None
    seen = set()
    try:
        app = win32com.GetActiveObject("Word.Application")
        for doc in app.Documents:
            entry = _doc_entry(doc)
            seen.add(entry["path"].lower())
            out["open_documents"].append(entry)
        with contextlib.suppress(Exception):
            out["protected_view_documents"] = [
                app.ProtectedViewWindows(i).Document.FullName
                for i in range(1, app.ProtectedViewWindows.Count + 1)
            ]
        # Other instances: open docs register file monikers in the ROT.
        # Template monikers are excluded — binding a STARTUP add-in
        # template (e.g. Zotero.dotm) and touching any attribute is a hard
        # ACCESS VIOLATION (native crash, not a catchable exception), and
        # loaded add-in templates are not open documents anyway.
        with contextlib.suppress(Exception):
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning():
                ctx = pythoncom.CreateBindCtx(0)
                try:
                    name = moniker.GetDisplayName(ctx, None)
                except Exception:
                    continue
                low = name.lower()
                if (
                    low in seen
                    or not low.endswith((".docx", ".docm", ".doc", ".rtf"))
                    or "\\microsoft\\word\\startup\\" in low
                ):
                    continue
                with contextlib.suppress(Exception):
                    obj = rot.GetObject(moniker)
                    doc = win32com.Dispatch(
                        obj.QueryInterface(pythoncom.IID_IDispatch)
                    )
                    if doc.FullName.lower() == low:
                        seen.add(low)
                        out["open_documents"].append(_doc_entry(doc))
    except Exception:
        out["interactive_state"] = "busy"
    finally:
        app = None
    return out
