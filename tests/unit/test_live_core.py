"""Live COM layer core tests (v2.0 L1).

The `live` tests spawn their own VISIBLE Word instance with a throwaway
document and quit it afterward; they skip cleanly where Word is unavailable
(CI). The classification tests are pure Python and run everywhere.
"""

from pathlib import Path

import pytest

from word_mcp.com import live
from word_mcp.core.errors import DocumentNotOpenInWord, WordMcpError


# ---------------------------------------------------------------- pure python


class _FakeComError(Exception):
    def __init__(self, hresult, scode=None):
        self.hresult = hresult
        self.args = (hresult, "msg", (0, None, None, None, None, scode), None)


def test_classify_busy_and_gone():
    assert isinstance(live._classify(_FakeComError(live.RPC_E_CALL_REJECTED)),
                      type(live._classify(_FakeComError(-2147418111))))
    from word_mcp.core.errors import WordBusy, WordDisconnected

    assert isinstance(live._classify(_FakeComError(-2147418111)), WordBusy)
    assert isinstance(live._classify(_FakeComError(-2147417846)), WordBusy)
    assert isinstance(live._classify(_FakeComError(-2147417848)), WordDisconnected)
    assert isinstance(live._classify(_FakeComError(-2147220995)), WordDisconnected)
    assert live._classify(_FakeComError(-1)) is None


def test_classify_nested_scode():
    # server-raised errors bury the HRESULT in excepinfo scode
    from word_mcp.core.errors import WordBusy

    err = _FakeComError(-2147352567, scode=-2147418111)  # DISP_E_EXCEPTION shell
    assert isinstance(live._classify(err), WordBusy)


def test_check_text_safe_rejects_cell_separator():
    with pytest.raises(WordMcpError):
        live.check_text_safe("before\x07after")
    live.check_text_safe("plain text is fine")


def test_state_guard_lifo_restore():
    class Obj:
        a = 1
        b = 2

    o = Obj()
    g = live.StateGuard()
    g.set(o, "a", 10)
    g.set(o, "b", 20)
    g.set(o, "a", 100)          # second mutation of the same attr
    assert (o.a, o.b) == (100, 20)
    failed = g.restore()
    assert failed == []
    assert (o.a, o.b) == (1, 2)  # LIFO unwinds the double-set correctly


def test_state_guard_reports_restore_failure():
    class Obj:
        locked = False

        @property
        def x(self):
            return 1

        @x.setter
        def x(self, v):
            if self.locked:
                raise RuntimeError("nope")

    o = Obj()
    g = live.StateGuard()
    g.set(o, "x", 5)
    o.locked = True
    assert g.restore() == ["x"]


# --------------------------------------------------------------------- live


def _word_available():
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    import winreg

    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
        return True
    except OSError:
        return False


live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)


def quit_instance_holding(path: str, wait: float = 8.0):
    """Quit exactly the Word instance holding `path` (via its ROT file
    moniker), then WAIT until its moniker leaves the ROT — a half-dead
    instance still answers GetActiveObject and poisons the next fixture.
    Assumes COM is already initialized on this thread."""
    import time

    import pythoncom
    import win32com.client

    def _find():
        rot = pythoncom.GetRunningObjectTable()
        for mk in rot.EnumRunning():
            ctx = pythoncom.CreateBindCtx(0)
            try:
                name = mk.GetDisplayName(ctx, None)
            except Exception:
                continue
            if name.lower() == str(path).lower():
                return rot, mk
        return None, None

    try:
        rot, mk = _find()
        if mk is None:
            return False
        obj = rot.GetObject(mk)
        d = win32com.client.Dispatch(
            obj.QueryInterface(pythoncom.IID_IDispatch)
        )
        d.Application.Quit(SaveChanges=0)
        d = None
        obj = None
        deadline = time.monotonic() + wait
        while time.monotonic() < deadline:
            _, still = _find()
            if still is None:
                return True
            time.sleep(0.25)
        return True
    except Exception:
        return False


@pytest.fixture(scope="module")
def probe_doc(tmp_path_factory):
    """A visible Word instance of our own, holding one throwaway document."""
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    doc = app.Documents.Add()
    doc.Content.InsertAfter("live test probe. " * 40)
    path = tmp_path_factory.mktemp("live") / "live_probe.docx"
    doc.SaveAs2(str(path))
    yield str(path)
    import sys

    # The proxies cached at setup are typically DEAD by now (the tests'
    # CoInitialize/CoUninitialize cycles disconnect them) — the same reason
    # live.py never caches COM pointers. Quit via a fresh ROT-scoped attach.
    doc = None
    app = None
    try:
        if not quit_instance_holding(str(path)):
            print("probe_doc teardown: instance not found — a WINWORD.EXE "
                  "may be left behind", file=sys.stderr)
    finally:
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_interactive_status_ready(probe_doc):
    status = live.interactive_status()
    assert status["interactive_state"] == "ready"
    assert any(
        d["path"].lower() == probe_doc.lower() for d in status["open_documents"]
    )


@live_mark
@needs_word
def test_run_live_result_fields_and_grouped_undo(probe_doc):
    def body(s):
        rng = s.doc.Range(0, 0)
        live.insert_text_chunked(rng, "ONE ")
        live.insert_text_chunked(s.doc.Range(0, 0), "TWO ")
        live.insert_text_chunked(s.doc.Range(0, 0), "THREE ")
        return {"ok": True}

    result = live.run_live(probe_doc, "test edits", body)
    assert result["live"] is True
    assert result["undo_grouped"] is True
    assert result["document_dirty"] is True
    assert result["state_restore_failed"] == []
    assert "autosave_on" in result

    def undo_body(s):
        s.doc.Undo(1)
        return {"text": s.doc.Range(0, 20).Text}

    after = live.run_live(probe_doc, "undo", undo_body)
    assert after["text"].startswith("live test probe.")


@live_mark
@needs_word
def test_state_guard_restores_track_revisions(probe_doc):
    def body(s):
        original = s.doc.TrackRevisions
        s.guard.set(s.doc, "TrackRevisions", not original)
        return {"original": bool(original), "during": bool(s.doc.TrackRevisions)}

    r = live.run_live(probe_doc, "guard test", body)
    assert r["during"] != r["original"]

    def check(s):
        return {"now": bool(s.doc.TrackRevisions)}

    r2 = live.run_live(probe_doc, "guard check", check)
    assert r2["now"] == r["original"]


@live_mark
@needs_word
def test_not_open_refusal(probe_doc):
    with pytest.raises(DocumentNotOpenInWord):
        live.run_live(str(Path(probe_doc).parent / "ghost.docx"), "x", lambda s: {})


@live_mark
@needs_word
def test_chunked_insert_over_32k(probe_doc):
    big = "x" * 65000

    def body(s):
        end = s.doc.Content.End
        live.insert_text_chunked(s.doc.Range(end - 1, end - 1), big)
        return {"len": s.doc.Content.End}

    r = live.run_live(probe_doc, "big insert", body)
    assert r["len"] > 65000

    def undo_body(s):
        s.doc.Undo(1)
        return {}

    live.run_live(probe_doc, "undo big", undo_body)


@live_mark
@needs_word
def test_live_repair_restores_screen_updating(probe_doc):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("Word.Application")
        app.ScreenUpdating = False
        app = None
    finally:
        pythoncom.CoUninitialize()

    report = live.live_repair()
    assert report["word_running"] is True
    assert "screen_updating_restored" in report["actions"]
