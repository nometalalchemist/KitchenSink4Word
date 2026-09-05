"""Regressions for the 2026-09-05 production concurrency matrix.

Every test here is a finding from that run, and the live ones use the
report's own repro shape: TWO independent server processes driving one
Word instance, started from a shared epoch so their operations actually
overlap. That two-process shape is the whole point — every one of these
failures was invisible to a single-process test, which is why the suite
was green while the matrix was destroying paragraphs.

Findings covered:

- H1  same-target replace double-applies, both callers report success
- H2  DisplayAlerts leaks to wdAlertsNone whenever two processes overlap
- H3  one wedged Word instance disables live editing in every instance
- H4  index TOCTOU deletes the wrong paragraph and reports deleted: 1
- M1  the status probe says "ready" during a dialog storm
- L1  backup=True is a silent no-op on the live route
- L2  the f9 test's own named-argument bug (fixed in test_l4_fixes)

The pure-Python tests run everywhere, CI included. The live tests spawn
and reclaim their own visible Word instance and skip where Word is absent.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import textwrap
import time
from pathlib import Path

import pytest

import word_mcp.server as srv
from word_mcp.com import live
from word_mcp.com import xproc
from word_mcp.core.errors import (
    LiveLockTimeout,
    TargetNotFound,
    WordBusy,
)

from test_live_core import _word_available, quit_instance_holding

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)


# ===================================================== pure python: the lock


@pytest.fixture()
def lock_dir(tmp_path, monkeypatch):
    d = tmp_path / "locks"
    monkeypatch.setenv(xproc._LOCK_DIR_ENV, str(d))
    return d


def test_lock_is_published_complete_and_released_when_ours(lock_dir):
    """Atomic publish: a lockfile is never observable half-written, and
    release removes it. (xlsx-mcp H-5: a two-syscall create let a
    concurrent acquirer read an EMPTY file, call it stale, and delete a
    LIVE lock.)"""
    path = lock_dir / f"{xproc.APP_SCOPE}.lock"
    with xproc.cross_process_lock("unit", wait=1.0) as owns:
        assert owns is True
        assert path.exists()
        info = json.loads(path.read_text(encoding="utf-8"))
        # every field present the moment the file exists
        assert info["pid"] == os.getpid()
        assert info["token"] == xproc._OWNER_TOKEN
        assert info["holder"] == "unit"
        assert isinstance(info["time"], float)
    assert not path.exists()


def test_reentrant_hold_does_not_deadlock_or_delete(lock_dir):
    path = lock_dir / f"{xproc.APP_SCOPE}.lock"
    with xproc.cross_process_lock("outer", wait=1.0) as outer:
        with xproc.cross_process_lock("inner", wait=1.0) as inner:
            assert inner is False       # re-entrant, not a second publish
            assert path.exists()
        # the inner exit must NOT have removed the outer's lockfile
        assert path.exists()
        assert outer is True
    assert not path.exists()


def test_release_only_removes_our_own_lockfile(lock_dir):
    """Unconditional unlink is how a leaked lock outlives every writer."""
    lock_dir.mkdir(parents=True, exist_ok=True)
    path = lock_dir / f"{xproc.APP_SCOPE}.lock"
    path.write_text(json.dumps({
        "pid": os.getpid(), "token": "someone-else", "time": time.time(),
    }), encoding="utf-8")
    xproc._release_lockfile(path)
    assert path.exists(), "released a lockfile belonging to another holder"


def test_our_pid_under_a_foreign_token_is_stale(lock_dir):
    """A recycled PID must not grant amnesty: the lock carries our number
    but not our token, so its writer is gone. The old design read that as
    'this is mine' and never cleaned it up."""
    assert xproc._is_stale({
        "pid": os.getpid(), "token": "foreign", "time": time.time(),
    })


def test_dead_pid_and_aged_lock_are_stale(lock_dir):
    dead = {"pid": 999_999_999, "token": "x", "time": time.time()}
    assert xproc._is_stale(dead)
    aged = {
        "pid": os.getpid() + 0, "token": "x",
        "time": time.time() - (xproc.LOCK_STALE_SECONDS + 60),
    }
    assert xproc._is_stale(aged)


def test_recycled_pid_detected_by_process_create_time(lock_dir):
    """Stale detection via create_time (xlsx-mcp M-1): 'is that PID alive'
    is the wrong question, because Windows recycles numbers."""
    if sys.platform != "win32":
        pytest.skip("create_time probe is Windows-only")
    real = xproc._process_create_time(os.getpid())
    if real is None:
        pytest.skip("process create time unavailable")
    forged = {
        "pid": os.getpid(), "token": "foreign",
        "pid_created": real + 1.0, "time": time.time(),
    }
    assert xproc._is_stale(forged)


_HOLDER = textwrap.dedent(
    """
    import os, sys, time
    os.environ["KS4W_LIVE_LOCK_DIR"] = sys.argv[1]
    sys.path.insert(0, sys.argv[3])
    from word_mcp.com import xproc
    with xproc.cross_process_lock("holder", wait=5.0):
        print("HELD", flush=True)
        time.sleep(float(sys.argv[2]))
    print("RELEASED", flush=True)
    """
)


def _spawn_holder(lock_dir: Path, seconds: float) -> subprocess.Popen:
    src = str(Path(srv.__file__).resolve().parents[1])
    proc = subprocess.Popen(
        [sys.executable, "-c", _HOLDER, str(lock_dir), str(seconds), src],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
    )
    deadline = time.monotonic() + 30.0
    while time.monotonic() < deadline:
        line = proc.stdout.readline()
        if line.strip() == "HELD":
            return proc
        if proc.poll() is not None:
            raise AssertionError(
                f"holder died: {proc.stderr.read()[:2000]}"
            )
    proc.kill()
    raise AssertionError("holder never reported HELD")


def test_second_process_is_refused_while_the_first_holds(lock_dir):
    """H1/H4/2b, at the lock level: the second PROCESS does not get in.
    The in-process RLock could never have produced this refusal — that is
    exactly the gap the matrix measured (49 overlapping operation pairs
    out of 50)."""
    proc = _spawn_holder(lock_dir, 6.0)
    try:
        # the holder is visible to status reporting while it holds
        held = xproc.holder_info()
        assert held is not None
        assert held["ours"] is False
        assert held["pid"] != os.getpid()
        assert held["holder"] == "holder"
        t0 = time.monotonic()
        with pytest.raises(LiveLockTimeout) as excinfo:
            with xproc.cross_process_lock("late", wait=1.0):
                pass
        waited = time.monotonic() - t0
        assert 0.8 <= waited < 5.0, "refusal did not respect the wait window"
        msg = str(excinfo.value)
        assert "another kitchensink4word server process" in msg
        assert f"PID {held['pid']}" in msg    # names the real holder
        assert "nothing was changed" in msg
    finally:
        proc.wait(timeout=30)
    # once released, the lock is free again
    with xproc.cross_process_lock("after", wait=2.0) as owns:
        assert owns is True


def test_lock_survives_a_missing_lock_directory(monkeypatch, tmp_path):
    """Degrade to the in-process mutex rather than failing a live edit,
    and never claim cross-process coverage that is not there."""
    monkeypatch.setattr(
        xproc, "_lock_dir", lambda: tmp_path / "nul\x00bad"
    )
    with xproc.cross_process_lock("degraded", wait=0.5) as owns:
        assert owns is False
    assert xproc.lock_state()["cross_process"] is False


# =========================================== pure python: H2, H3, M1, L1, H4


def test_state_guard_restores_to_an_override_not_the_snapshot():
    """H2: snapshot-and-restore re-leaks another process's suppression."""
    class App:
        DisplayAlerts = live._WD_ALERTS_NONE   # leaked by someone else

    app = App()
    guard = live.StateGuard()
    recovered = live._suppress_alerts_owned(guard, app)
    assert recovered is True
    assert app.DisplayAlerts == live._WD_ALERTS_NONE   # suppressed in-session
    assert guard.restore() == []
    assert app.DisplayAlerts == live._WD_ALERTS_ALL, (
        "a leaked wdAlertsNone was restored as if it were the user's setting"
    )


def test_state_guard_keeps_a_genuine_user_value():
    class App:
        DisplayAlerts = -2      # wdAlertsMessageBox, a real user setting

    app = App()
    guard = live.StateGuard()
    assert live._suppress_alerts_owned(guard, app) is False
    guard.restore()
    assert app.DisplayAlerts == -2


def test_state_guard_reports_a_restore_that_did_not_take():
    """A setattr that raised nothing is not proof the value went back."""
    class Sticky:
        def __init__(self):
            self._v = 1

        @property
        def attr(self):
            return self._v

        @attr.setter
        def attr(self, value):
            self._v = 99        # accepts the write, keeps its own value

    o = Sticky()
    g = live.StateGuard()
    g.set(o, "attr", 5)
    failed = g.restore()
    assert len(failed) == 1
    assert "reads back" in failed[0]


def test_late_bound_attribute_error_classifies_as_busy():
    """H3: pywin32's dynamic dispatch turns a refused property call into a
    plain AttributeError with the HRESULT discarded."""
    typed = live._classify(AttributeError("<unknown>.FullName"))
    assert isinstance(typed, WordBusy)
    assert "dialog" in str(typed)


def test_resolve_document_falls_back_to_the_rot_when_the_primary_wedges(
    monkeypatch
):
    """H3, the consequence that mattered: the ROT fallback exists precisely
    for a wedged instance, and catching only com_error meant it was never
    reached, so the healthy document in the OTHER instance stayed
    unreachable."""
    class WedgedDocs:
        def __iter__(self):
            raise AttributeError("<unknown>.FullName")

    class WedgedApp:
        Documents = WedgedDocs()

    healthy_app, healthy_doc = object(), object()
    monkeypatch.setattr(
        live, "_find_doc_via_rot",
        lambda *a, **k: (healthy_app, healthy_doc),
    )

    class _PW:
        class com_error(Exception):
            pass

    app, doc = live._resolve_document(
        None, _PW, None, WedgedApp(), r"C:\any\where.docx"
    )
    assert app is healthy_app and doc is healthy_doc


def test_resolve_document_raises_typed_busy_when_the_rot_has_nothing(
    monkeypatch
):
    """...and the caller gets WordBusy naming a dialog, not a raw
    AttributeError: <unknown>.FullName."""
    class WedgedDocs:
        def __iter__(self):
            raise AttributeError("<unknown>.FullName")

    class WedgedApp:
        Documents = WedgedDocs()

    monkeypatch.setattr(live, "_find_doc_via_rot", lambda *a, **k: (None, None))

    class _PW:
        class com_error(Exception):
            pass

    with pytest.raises(WordBusy):
        live._resolve_document(
            None, _PW, None, WedgedApp(), r"C:\any\where.docx"
        )


def _fake_com_modules():
    """A Word that answers a property read perfectly well — which is
    exactly the state M1 is about: a modal can be up and simple reads
    still succeed."""
    class _PW:
        class com_error(Exception):
            pass

    class _W32:
        @staticmethod
        def GetActiveObject(_name):
            return type("App", (), {"Name": "Microsoft Word"})()

    class _PC:
        @staticmethod
        def CoInitialize():
            return None

        @staticmethod
        def CoUninitialize():
            return None

    return _PC, _PW, _W32


def test_probe_does_not_report_ready_while_a_dialog_is_up(monkeypatch):
    """M1: the window layer already knew the truth and was not consulted.
    The matrix's probe answered 'ready' while every document resolution in
    the same instance was failing."""
    monkeypatch.setattr(live, "_com_modules", _fake_com_modules)
    monkeypatch.setattr(
        live._dialogs, "pending_dialogs",
        lambda pids=None: [{"title": "", "class": "NUIDialog"}],
    )
    assert live.probe_with_timeout(timeout=5.0) == "blocked"


def test_probe_reports_ready_when_no_dialog_is_up(monkeypatch):
    monkeypatch.setattr(live, "_com_modules", _fake_com_modules)
    monkeypatch.setattr(live._dialogs, "pending_dialogs", lambda pids=None: [])
    assert live.probe_with_timeout(timeout=5.0) == "ready"


def test_live_route_reports_that_no_backup_was_taken():
    """L1: accepting backup=True and silently doing nothing."""
    from word_mcp.core.errors import DocumentLocked

    def _file():
        raise DocumentLocked("open in Word")

    out = srv._route_live(
        "auto", _file, lambda: {"live": True, "replaced": 1}, backup=True
    )
    assert out["backup"] is False
    assert "no file backup was taken" in out["backup_skipped"]

    out = srv._route_live(
        "auto", _file, lambda: {"live": True}, backup=False
    )
    assert out["backup"] is False
    assert "backup_skipped" not in out

    # the file route still reports whatever the file layer reported
    out = srv._route_live("auto", lambda: {"saved": "x"}, None, backup=True)
    assert "backup" not in out


def test_expect_guard_refuses_a_drifted_index():
    """H4's proven mitigation, generalized to the destructive paths."""
    class P:
        def __init__(self, text):
            self.Range = type("R", (), {"Text": text + "\r"})()

    from word_mcp.com import live_ops

    paras = [P("LINE-018 body text here."), P("LINE-019 body text here.")]
    live_ops._expect_guard(paras, 0, "LINE-018", "delete_paragraphs")  # ok
    with pytest.raises(TargetNotFound) as excinfo:
        live_ops._expect_guard(paras, 0, "LINE-022", "delete_paragraphs")
    msg = str(excinfo.value)
    assert "does not contain the expected text 'LINE-022'" in msg
    assert "LINE-018" in msg          # says what is actually there
    assert "Nothing was changed" in msg


def test_delete_paragraphs_expect_guard_on_the_file_route(tmp_path):
    """Same guard, same refusal, on the closed-file route: a tool argument
    that only worked on one of two routes would be its own silent no-op."""
    path = tmp_path / "guard.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": f"LINE-{i:03d} body text here."} for i in range(6)],
        backup=False, live="off",
    )
    with pytest.raises(TargetNotFound):
        srv.delete_paragraphs(
            str(path), start=2, expect_start="LINE-999",
            backup=False, live="off",
        )
    # nothing was deleted
    paras = srv.get_text(str(path), live="off")
    assert len(paras) == 6
    # and the guard passes when the index really is what the caller thinks
    r = srv.delete_paragraphs(
        str(path), start=2, expect_start="LINE-002", backup=False, live="off",
    )
    assert r["deleted"] == 1
    assert len(srv.get_text(str(path), live="off")) == 5


# ================================================ live: the two-process repros

_WORKER = textwrap.dedent(
    """
    import json, sys, time
    sys.path.insert(0, sys.argv[1])
    import word_mcp.server as srv

    path, who, epoch, mode = sys.argv[2], sys.argv[3], float(sys.argv[4]), sys.argv[5]
    while time.time() < epoch:
        time.sleep(0.002)
    out = {"who": who, "calls": [], "errors": []}
    try:
        if mode == "replace":
            for i in range(6):
                tok = "TOKEN-%02d" % i
                try:
                    r = srv.search_and_replace(
                        path, [{"find": tok, "replace": "HIT-%s-%02d" % (who, i)}],
                        live="force",
                    )
                    out["calls"].append(r.get("total"))
                except Exception as exc:
                    out["errors"].append(repr(exc))
        elif mode == "edit_only":
            for i in range(6):
                try:
                    srv.insert_paragraphs(
                        path, [{"text": "EDIT-%s-%02d" % (who, i)}],
                        live="force",
                    )
                    out["calls"].append(1)
                except Exception as exc:
                    out["errors"].append(repr(exc))
    finally:
        print("RESULT" + json.dumps(out), flush=True)
    """
)


def _run_workers(path: str, mode_a: str, mode_b: str, *, lead: float = 2.0,
                 timeout: float = 90.0) -> list[dict]:
    """Two independent server processes started from a shared epoch, the
    matrix harness's own shape."""
    src = str(Path(srv.__file__).resolve().parents[1])
    epoch = time.time() + lead
    procs = [
        subprocess.Popen(
            [sys.executable, "-c", _WORKER, src, path, who, str(epoch), mode],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
        )
        for who, mode in (("A", mode_a), ("B", mode_b))
    ]
    out = []
    try:
        for proc in procs:
            stdout, stderr = proc.communicate(timeout=timeout)
            line = next(
                (ln for ln in stdout.splitlines() if ln.startswith("RESULT")),
                None,
            )
            assert line, f"worker produced no result; stderr: {stderr[:2000]}"
            out.append(json.loads(line[len("RESULT"):]))
    finally:
        for proc in procs:
            if proc.poll() is None:
                proc.kill()
    return out


@pytest.fixture()
def two_process_doc(tmp_path_factory):
    """A visible Word instance this test OWNS, holding one document."""
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path_factory.mktemp("xproc") / "shared.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": f"TOKEN-{i:02d} sits on this line."} for i in range(6)],
        backup=False, live="off",
    )
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.DisplayAlerts = live._WD_ALERTS_ALL   # the real user default
    app.Documents.Open(str(path))
    yield str(path)
    app = None
    try:
        quit_instance_holding(str(path))
    finally:
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_h1_same_target_replace_no_longer_double_applies(two_process_doc):
    """H1's exact repro: two processes, one open document, each replacing
    the SAME six tokens, started from a shared epoch. The matrix got six
    tokens, twelve claimed replacements, and text welded from two edits
    ('HIT-C1-00HIT-C2-00 sits on this line.') with both callers reporting
    success. Serialized across processes, the second caller finds nothing
    left to replace and says so."""
    results = _run_workers(two_process_doc, "replace", "replace")
    assert not any(r["errors"] for r in results), [r["errors"] for r in results]
    claimed = sum(sum(c or 0 for c in r["calls"]) for r in results)
    assert claimed == 6, f"six tokens, {claimed} claimed replacements"

    paras = srv.get_text(two_process_doc, live="force")
    texts = [p["text"] if isinstance(p, dict) else str(p) for p in paras]
    body = "\n".join(texts)
    for line in texts:
        markers = line.count("HIT-")
        assert markers <= 1, (
            f"two replacements welded into one line: {line!r}\n{body}"
        )
    assert body.count("HIT-") == 6


@live_mark
@needs_word
def test_h2_display_alerts_survives_two_overlapping_processes(
    two_process_doc
):
    """H2's three-way control, at the C3 setting that failed: two
    processes. The matrix measured -1 -> 0, and then a dirty document
    closing in 0.09 s with the unsaved sentence silently discarded."""
    import win32com.client

    results = _run_workers(two_process_doc, "edit_only", "edit_only")
    assert not any(r["errors"] for r in results), [r["errors"] for r in results]
    app = win32com.client.GetActiveObject("Word.Application")
    try:
        assert app.DisplayAlerts == live._WD_ALERTS_ALL, (
            f"DisplayAlerts left at {app.DisplayAlerts} after two server "
            "processes overlapped; a suppressed Word answers its own save "
            "prompts and discards unsaved work"
        )
    finally:
        app = None


_SHIFTER = textwrap.dedent(
    """
    import sys
    sys.path.insert(0, sys.argv[1])
    import word_mcp.server as srv
    for i in range(6):
        srv.insert_paragraphs(
            sys.argv[2], [{"text": "SHIFT-%02d" % i}],
            location={"paragraph": 0, "position": "before"}, live="force")
    print("SHIFTED", flush=True)
    """
)


def _shift_document_from_another_process(path: str) -> None:
    """A SECOND server process inserts six paragraphs at the top, which
    shifts every index below it. This is the other agent in H4."""
    src = str(Path(srv.__file__).resolve().parents[1])
    proc = subprocess.run(
        [sys.executable, "-c", _SHIFTER, src, path],
        capture_output=True, text=True, timeout=180,
    )
    assert "SHIFTED" in proc.stdout, f"shifter failed: {proc.stderr[:2000]}"


@live_mark
@needs_word
def test_h4_guarded_delete_refuses_instead_of_destroying(tmp_path_factory):
    """H4's exact shape: an index is resolved in ONE tool call, the caller
    pauses the way an agent pauses to think, a second process shifts the
    document underneath, and the delete lands in the next call. No lock
    closes this — the resolve and the delete are different calls — which is
    why the matrix got three destroyed paragraphs and six confident
    successes. Both arms run here: unguarded still destroys (the race is
    real, not an artefact), guarded refuses and names the drift."""
    import pythoncom
    import win32com.client

    def _fresh(name: str):
        path = tmp_path_factory.mktemp(name) / "shift.docx"
        srv.create_document(str(path))
        srv.insert_paragraphs(
            str(path),
            [{"text": f"LINE-{i:03d} body text here."} for i in range(30)],
            backup=False, live="off",
        )
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = True
        app.Documents.Open(str(path))
        return str(path)

    def _body(path: str) -> str:
        return "\n".join(
            p["text"] if isinstance(p, dict) else str(p)
            for p in srv.get_text(path, live="force")
        )

    pythoncom.CoInitialize()
    try:
        # --- arm A: no guard. The race is still there and still destroys.
        path = _fresh("toctou_unguarded")
        try:
            assert "LINE-020" in _body(path).splitlines()[20]
            _shift_document_from_another_process(path)
            r = srv.delete_paragraphs(path, start=20, live="force")
            assert r["deleted"] == 1          # confident success...
            body = _body(path)
            assert "LINE-014" not in body     # ...on the wrong paragraph
            assert "LINE-020" in body         # target untouched
        finally:
            quit_instance_holding(path)

        # --- arm B: the guard. Same race, nothing destroyed.
        path = _fresh("toctou_guarded")
        try:
            assert "LINE-020" in _body(path).splitlines()[20]
            _shift_document_from_another_process(path)
            with pytest.raises(TargetNotFound) as excinfo:
                srv.delete_paragraphs(
                    path, start=20, expect_start="LINE-020", live="force"
                )
            msg = str(excinfo.value)
            assert "does not contain the expected text 'LINE-020'" in msg
            assert "LINE-014" in msg          # names what is actually there
            body = _body(path)
            for i in range(14, 21):
                assert f"LINE-{i:03d}" in body, (
                    f"LINE-{i:03d} was destroyed despite the guard\n{body}"
                )
        finally:
            quit_instance_holding(path)
    finally:
        pythoncom.CoUninitialize()
