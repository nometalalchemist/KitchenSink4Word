"""v2.0.0 live-mode fix batch (2026-09-03 live COM stress report).

Non-live regression coverage for the five MUST-FIX items:
1. COM serialization: lock mechanics, no interleaving under threads, and
   an entry-point audit proving every COM path takes the lock.
2. Tracked-replace loop: an emulated-COM Word 2016 harness reproduces the
   report's failure mechanism (tracked assignment inserts BEFORE the
   deletion markup and the deleted copy stays findable) and proves the
   three-layer fix: revision skip fails closed, resume advances past own
   markup, and tracked mode caps at the pre-edit match count.
3. apply_edits atomicity: preflight conflict simulation and rollback
   note contract.
4. Dialog prevention: alerts-suppression context manager contract.
5. Bounded timeouts: fast path, error propagation, queue-vs-stuck
   distinction, and parameter validation.
6. Name-collision guard and OS-layer dialog detection (synthetic #32770).

The live halves (real Word) are in test_live_concurrency.py, marked live
for the next zero-foreign-WINWORD round.
"""

from __future__ import annotations

import inspect
import sys
import threading
import time

import pytest

from word_mcp.com import bridge, convert, dialogs
from word_mcp.com import serial as com_serial
from word_mcp.com import live_batch, live_ops
from word_mcp.core.errors import (
    TargetNotFound,
    WordBlocked,
    WordBusy,
    WordMcpError,
)

# ------------------------------------------------------- 1. lock mechanics


def test_com_operation_records_state_and_reenters():
    snap0 = com_serial.lock_snapshot()
    assert snap0["held"] is False
    with com_serial.com_operation("outer-op"):
        snap = com_serial.lock_snapshot()
        assert snap["held"] is True
        assert snap["current_op"]["name"] == "outer-op"
        with com_serial.com_operation("nested-op"):  # RLock: no deadlock
            assert com_serial.lock_snapshot()["held"] is True
    snap2 = com_serial.lock_snapshot()
    assert snap2["held"] is False
    assert snap2["last_op"]["name"] == "outer-op"
    assert snap2["last_op"]["duration_ms"] >= 0


def test_threads_serialize_no_interleaving():
    """Two threads running COM-op bodies must never overlap in time."""
    spans = []

    def op(name):
        with com_serial.com_operation(name):
            t0 = time.monotonic()
            time.sleep(0.15)
            spans.append((t0, time.monotonic()))

    threads = [
        threading.Thread(target=op, args=(f"t{i}",)) for i in range(4)
    ]
    for t in threads:
        t.start()
    for t in threads:
        t.join(10)
    assert len(spans) == 4
    spans.sort()
    for (s1, e1), (s2, _e2) in zip(spans, spans[1:]):
        assert e1 <= s2 + 1e-4, "COM operations overlapped in time"


def test_bounded_acquire_reports_busy_from_other_thread():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("holding-op"):
            held.set()
            release.wait(5)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        assert com_serial.acquire(timeout=0.05) is False
        snap = com_serial.lock_snapshot()
        assert snap["held"] is True
        assert snap["current_op"]["name"] == "holding-op"
    finally:
        release.set()
        t.join(5)


# ---------------------------------------------- 1b. entry-point coverage


BRIDGE_EXEMPT = {
    # tasklist only, no COM:
    "zombie_check",
    # bounded try-acquire form, marked manually and asserted below:
    "word_status",
}


def _public_functions(module):
    return {
        name: fn
        for name, fn in vars(module).items()
        if callable(fn)
        and not name.startswith("_")
        and inspect.isfunction(inspect.unwrap(fn))
        and getattr(fn, "__module__", "") == module.__name__
    }


def test_every_bridge_entry_point_is_serialized():
    for name, fn in _public_functions(bridge).items():
        if name in BRIDGE_EXEMPT and name != "word_status":
            continue
        assert getattr(fn, "_com_serialized", None), (
            f"bridge.{name} does not take the COM serialization lock "
            "(missing @_serial.serialized / @_bounded_op)"
        )


def test_convert_entry_point_is_serialized():
    assert getattr(convert.import_pdf, "_com_serialized", None)


def test_live_session_acquires_the_lock(monkeypatch):
    """live_session (every live tool and dual-mode auto-route funnels
    through it) must hold the lock before touching COM."""
    from word_mcp.com import live

    if sys.platform != "win32":  # pragma: no cover
        pytest.skip("live layer is Windows-only")
    try:
        import pythoncom  # noqa: F401
    except ImportError:  # pragma: no cover
        pytest.skip("pywin32 not installed")

    seen = {}

    class Sentinel(Exception):
        pass

    def fake_attach(win32com, pythoncom):
        seen["snapshot"] = com_serial.lock_snapshot()
        raise Sentinel()

    monkeypatch.setattr(live, "_attach_app", fake_attach)
    with pytest.raises(Sentinel):
        live.run_live("C:/nonexistent.docx", "probe tool", lambda s: {})
    assert seen["snapshot"]["held"] is True
    assert seen["snapshot"]["current_op"]["name"] == "live:probe tool"


def test_interactive_status_reports_serving_when_lock_held():
    from word_mcp.com import live

    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("long-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        out = live.interactive_status()
        assert out["interactive_state"] == "serving"
        assert out["com_serialization"]["held"] is True
        assert out["com_serialization"]["current_op"]["name"] == "long-op"
    finally:
        release.set()
        t.join(5)


def test_word_status_skips_probe_when_lock_held():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("long-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        out = bridge.word_status()
        assert "note" in out and "not probed" in out["note"]
    finally:
        release.set()
        t.join(5)


# ------------------------------------------------- 5. bounded timeouts


def test_run_bounded_fast_path_and_error_propagation():
    assert bridge._run_bounded("fast", 10, lambda: {"ok": 1}) == {"ok": 1}
    with pytest.raises(TargetNotFound):
        bridge._run_bounded(
            "err", 10, lambda: (_ for _ in ()).throw(TargetNotFound("x"))
        )


def test_run_bounded_stuck_op_raises_word_blocked():
    def stuck():
        time.sleep(3)
        return {}

    t0 = time.monotonic()
    with pytest.raises(WordBlocked, match="did not finish within"):
        bridge._run_bounded("stuck-op", 0.3, stuck)
    # must not have waited for the sleep to end before raising... the
    # implementation waits up to 10s for the worker to unwind after the
    # (no-op) kill; the worker finishes its sleep in 3s, well under that
    assert time.monotonic() - t0 < 8


def test_run_bounded_queued_behind_lock_raises_word_busy():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("blocking-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        with pytest.raises(WordBusy, match="blocking-op"):
            bridge._run_bounded("queued-op", 0.3, lambda: {})
    finally:
        release.set()
        t.join(5)


def test_bounded_op_timeout_validation():
    @bridge._bounded_op("val-op", default=60.0)
    def sample():
        return {"ok": True}

    assert sample() == {"ok": True}
    assert sample._com_serialized == "val-op"
    with pytest.raises(WordMcpError, match="between 5 and 3600"):
        sample(timeout=1)
    with pytest.raises(WordMcpError, match="number of seconds"):
        sample(timeout="soon")


# ---------------------------------------- 2. tracked-replace loop (fake COM)


class _FakeRev:
    def __init__(self, start, end):
        self.Type = 2  # wdRevisionDelete

        class _R:
            pass

        self.Range = _R()
        self.Range.Start = start
        self.Range.End = end


class _FakeFind:
    def __init__(self, rng):
        self._rng = rng
        self.Text = ""
        self.Forward = True
        self.Wrap = 0
        self.MatchWildcards = False
        self.MatchCase = True

    def ClearFormatting(self):
        pass

    def Execute(self):
        story = self._rng._story
        i = story.text.find(self.Text, self._rng.Start)
        if i < 0:
            return False
        self._rng.SetRange(i, i + len(self.Text))
        return True


class _FakeRange:
    def __init__(self, story, start, end):
        self._story = story
        self.Start = start
        self.End = end

    def SetRange(self, start, end):
        self.Start, self.End = start, end

    @property
    def Duplicate(self):
        return _FakeRange(self._story, self.Start, self.End)

    @property
    def Find(self):
        return _FakeFind(self)

    @property
    def Text(self):
        return self._story.text[self.Start:self.End]

    @Text.setter
    def Text(self, value):
        self._story.assign(self, value)

    @property
    def Revisions(self):
        return self._story.revisions_for(self.Start, self.End)

    @property
    def ParentContentControl(self):
        return None


class FakeStory(_FakeRange):
    """Emulates the Word 2016 behavior the stress report hit: a tracked
    Range.Text assignment INSERTS the replacement at the range start and
    keeps the old text findable in the story as a tracked deletion AFTER
    the insertion (so a naive resume point sits right before the deleted
    copy of the find text)."""

    def __init__(self, text, *, tracked=False, revisions_visible=True,
                 revisions_raise=False, deletions=None):
        self._story = self
        self.text = text
        self.tracked = tracked
        self.revisions_visible = revisions_visible
        self.revisions_raise = revisions_raise
        self.deletions = list(deletions or [])
        self.assignments = 0

    @property
    def Start(self):
        return 0

    @Start.setter
    def Start(self, v):  # story range is the whole story
        pass

    @property
    def End(self):
        return len(self.text)

    @End.setter
    def End(self, v):
        pass

    @property
    def Fields(self):
        return []

    def assign(self, rng, value):
        self.assignments += 1
        if not self.tracked:
            old = self.text
            self.text = old[:rng.Start] + value + old[rng.End:]
            rng.End = rng.Start + len(value)
            return
        old_len = rng.End - rng.Start
        self.text = (
            self.text[:rng.Start] + value + self.text[rng.Start:]
        )
        shift = len(value)
        self.deletions = [
            (s + shift if s >= rng.Start else s,
             e + shift if e > rng.Start else e)
            for s, e in self.deletions
        ]
        self.deletions.append(
            (rng.Start + shift, rng.Start + shift + old_len)
        )
        rng.End = rng.Start + shift  # range covers the insertion only

    def revisions_for(self, start, end):
        if self.revisions_raise:
            raise RuntimeError("COM call rejected (emulated contention)")
        if not self.revisions_visible:
            return []
        return [
            _FakeRev(s, e) for s, e in self.deletions
            if s < end and e > start
        ]


ERP_TEXT = (
    "The enterprise resource planning rollout continues. "
    "Other text follows the single occurrence."
)


def test_tracked_replace_does_not_rematch_own_markup():
    """The report's bug 1 mechanism: 1 occurrence, tracked replace,
    insertion lands before the still-findable deleted copy. Must replace
    exactly once (the pre-fix loop produced ERP x51)."""
    story = FakeStory(ERP_TEXT, tracked=True)
    done, skipped, skipped_del = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 1
    assert story.text.count("ERP") == 1
    assert "ERPERP" not in story.text


def test_tracked_replace_two_occurrences_both_replaced_once():
    """Guarantee 2 in action: after replacing occurrence 1, the resume
    point advances past its own deletion markup, so occurrence 2 (the
    real one) is the next match, not the deleted copy of occurrence 1."""
    text = ERP_TEXT + " Later the enterprise resource planning gets cut."
    story = FakeStory(text, tracked=True)
    done, _, _ = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 2
    assert story.text.count("ERP") == 2
    assert "ERPERP" not in story.text


def test_tracked_replace_capped_even_if_revisions_invisible():
    """Defense in depth: if COM misreports the deletion markup as
    revision-free (the concurrency failure mode), the pre-edit match
    count caps replacements at 1 — garbage growth is impossible."""
    story = FakeStory(ERP_TEXT, tracked=True, revisions_visible=False)
    done, _, _ = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 1
    assert story.text.count("ERP") == 1
    assert "ERPERP" not in story.text


def test_tracked_replace_fails_closed_on_revision_read_error():
    story = FakeStory(ERP_TEXT, tracked=True, revisions_raise=True)
    with pytest.raises(WordBusy, match="fail-closed"):
        live_ops._replace_literal(
            story, "enterprise resource planning", "ERP", tracked=True
        )
    assert story.assignments == 0  # nothing replaced unverified


def test_tracked_replace_skips_preexisting_deletion():
    """A document that ALREADY contains tracked changes (the report's
    single-agent hypothesis): a match inside an existing tracked deletion
    is skipped, the live one is replaced."""
    text = "old enterprise resource planning gone. " + ERP_TEXT
    start = text.find("enterprise")
    end = start + len("enterprise resource planning")
    story = FakeStory(text, tracked=True, deletions=[(start, end)])
    done, _, skipped_del = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 1
    assert skipped_del == 1
    assert story.text.count("ERP") == 1


def test_untracked_self_referencing_replacement_still_terminates():
    story = FakeStory("alliance one alliance two", tracked=False)
    done, _, _ = live_ops._replace_literal(
        story, "alliance", "alliance-x", tracked=False
    )
    assert done == 2
    assert story.text == "alliance-x one alliance-x two"


# ------------------------------------ 3. apply_edits atomicity (unit level)


def _spec(op, **kw):
    out = {"op": op, "index": kw.pop("index", None)}
    out.update(kw)
    return out


def test_preflight_conflicts_catches_deleted_target():
    specs = [
        _spec("delete", indices=[2]),
        _spec("set_text", index=2, text="x"),
    ]
    with pytest.raises(WordMcpError, match="Nothing was applied"):
        live_batch._preflight_conflicts(specs, n_paras=5)


def test_preflight_conflicts_accepts_valid_sequence():
    specs = [
        _spec("insert", index=1, items=[{"text": "a"}], mode="after"),
        _spec("set_text", index=3, text="x"),
        _spec("delete", indices=[0, 4]),
    ]
    live_batch._preflight_conflicts(specs, n_paras=5)  # no raise


class _FakeUndo:
    def __init__(self, recording=True):
        self.IsRecordingCustomRecord = recording
        self.ended = 0

    def EndCustomRecord(self):
        self.ended += 1
        self.IsRecordingCustomRecord = False


class _FakeSession:
    def __init__(self, grouped=True, undo_raises=False):
        self.undo_grouped = grouped

        class _App:
            pass

        class _Doc:
            def __init__(self):
                self.undone = 0
                self._raises = undo_raises

            def Undo(self):
                if self._raises:
                    raise RuntimeError("no undo")
                self.undone += 1

        self.app = _App()
        self.app.UndoRecord = _FakeUndo()
        self.doc = _Doc()


def test_rollback_via_undo_group():
    s = _FakeSession(grouped=True)
    note = live_batch._rollback(s)
    assert "ROLLED BACK" in note
    assert s.app.UndoRecord.ended == 1
    assert s.doc.undone == 1


def test_rollback_honest_when_ungrouped_or_failing():
    assert "PARTIALLY APPLIED" in live_batch._rollback(
        _FakeSession(grouped=False)
    )
    assert "PARTIALLY APPLIED" in live_batch._rollback(
        _FakeSession(grouped=True, undo_raises=True)
    )


# ------------------------------------------- 4. alerts suppression contract


def test_alerts_suppressed_sets_and_restores():
    class _App:
        DisplayAlerts = -1

    app = _App()
    with bridge._alerts_suppressed(app):
        assert app.DisplayAlerts == 0
    assert app.DisplayAlerts == -1


def test_alerts_suppressed_restores_on_error():
    class _App:
        DisplayAlerts = -1

    app = _App()
    with pytest.raises(RuntimeError):
        with bridge._alerts_suppressed(app):
            assert app.DisplayAlerts == 0
            raise RuntimeError("boom")
    assert app.DisplayAlerts == -1


def test_retry_word_call_backs_off_and_reports():
    calls = []

    def flaky():
        calls.append(1)
        if len(calls) < 3:
            raise RuntimeError("transient")
        return "ok"

    value, retries = bridge._retry_word_call(flaky, first_delay=0.01)
    assert value == "ok" and retries == 2

    def always_fails():
        raise RuntimeError("permission error")

    with pytest.raises(WordBusy, match="no dialog is pending"):
        bridge._retry_word_call(
            always_fails, attempts=2, first_delay=0.01
        )


# --------------------------------------- 6. dialog detection (OS layer)


def test_pending_dialogs_empty_without_word():
    assert dialogs.pending_dialogs(pids=set()) == []
    assert dialogs.pending_dialogs(pids={999999999}) == []


@pytest.mark.skipif(sys.platform != "win32", reason="Win32 only")
def test_pending_dialogs_sees_synthetic_dialog_window():
    """Create a hidden-offscreen but WS_VISIBLE #32770 window in this
    process and detect it via the same enumeration com_word_status uses."""
    import ctypes
    import os

    user32 = ctypes.windll.user32
    user32.CreateWindowExW.restype = ctypes.c_void_p
    WS_POPUP = 0x80000000
    WS_VISIBLE = 0x10000000
    title = "ks4w dialog probe: file permission error"
    hwnd = user32.CreateWindowExW(
        0, "#32770", title, WS_POPUP | WS_VISIBLE,
        -32000, -32000, 1, 1, None, None, None, None,
    )
    assert hwnd, "could not create the synthetic dialog window"
    try:
        found = dialogs.pending_dialogs(pids={os.getpid()})
        assert any(
            d["title"] == title and d["class"] == "#32770" for d in found
        ), f"synthetic dialog not detected: {found}"
    finally:
        user32.DestroyWindow(hwnd)
    assert not any(
        d["title"] == title
        for d in dialogs.pending_dialogs(pids={os.getpid()})
    )


# ----------------------------------------- name-collision guard (unit)


def test_proofing_refuses_open_document(monkeypatch, tmp_path):
    doc = tmp_path / "open_doc.docx"
    doc.write_bytes(b"PK\x03\x04stub")
    monkeypatch.setattr(bridge, "_open_in_running_word", lambda p: True)
    with pytest.raises(WordBusy, match="same-name dialog"):
        bridge.proofing_errors(str(doc))
    with pytest.raises(WordBusy, match="same-name dialog"):
        bridge.readability_statistics(str(doc))
