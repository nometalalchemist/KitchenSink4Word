"""LIVE round: concurrency + dialog-prevention verification for the
v2.0.0 live-mode fix batch (2026-09-03 stress report).

Written non-live on 2026-09-04 (the author's Word was open, so no live
run was possible); these run in the next zero-foreign-WINWORD live
round. They spawn their own visible Word instance, fire concurrent live
writes through the SERVER layer from multiple threads, and assert the
serialization lock produced intact, non-interleaved text; plus the
single-agent tracked-replace repro the report asked for (a document that
already contains tracked changes).
"""

from __future__ import annotations

import threading

import pytest

import word_mcp.server as srv
from word_mcp.com import live
from word_mcp.com import serial as com_serial

from test_live_core import _word_available, quit_instance_holding

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)

N_PARAS = 12


@pytest.fixture()
def busy_doc(tmp_path_factory):
    """A visible Word instance holding one document with numbered
    paragraphs for the concurrency hammer."""
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path_factory.mktemp("live_conc") / "concurrency.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": f"paragraph {i:02d} marker-{i:02d} stable tail."}
         for i in range(N_PARAS)],
        backup=False,
    )
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.Documents.Open(str(path))
    yield str(path)
    app = None
    try:
        quit_instance_holding(str(path))
    finally:
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_concurrent_live_writes_serialize_without_interleaving(busy_doc):
    """The stress report's failure mode, reproduced deliberately: many
    threads writing through the server layer at once. Under the lock,
    every write must land intact (no character-level interleaving, no
    garbled text) and every call must succeed or refuse cleanly."""
    errors: list = []
    results: list = []

    def replace_worker(i):
        try:
            r = srv.search_and_replace(
                busy_doc,
                [{"find": f"marker-{i:02d}", "replace": f"HIT-{i:02d}"}],
            )
            results.append((i, r["total"]))
        except Exception as exc:  # noqa: BLE001 - collected for assert
            errors.append((i, repr(exc)))

    def set_text_worker(i):
        try:
            srv.set_paragraph_text(
                busy_doc, {"paragraph": i},
                f"rewritten {i:02d} clean full sentence body.",
            )
        except Exception as exc:  # noqa: BLE001
            errors.append((i, repr(exc)))

    threads = [
        threading.Thread(target=replace_worker, args=(i,))
        for i in range(0, 6)
    ] + [
        threading.Thread(target=set_text_worker, args=(i,))
        for i in range(6, 12)
    ]
    for t in threads:
        t.start()
    for t in threads:
        t.join(120)
    assert not errors, f"concurrent live calls failed: {errors}"
    assert all(total == 1 for _i, total in results)
    text = srv.get_text(busy_doc)
    joined = "\n".join(p["text"] for p in text)
    for i in range(0, 6):
        assert f"HIT-{i:02d}" in joined
    for i in range(6, 12):
        assert f"rewritten {i:02d} clean full sentence body." in joined
    # the report's signature corruptions must be absent
    assert "HIT-HIT" not in joined
    for p in text:
        for frag in ("monitoment", "\x07"):
            assert frag not in p["text"]
    snap = com_serial.lock_snapshot()
    assert snap["ops_serialized"] >= 12


@live_mark
@needs_word
def test_tracked_replace_single_agent_on_pretracked_doc(busy_doc):
    """Report bug 1, single-agent honest repro: a document that ALREADY
    contains tracked changes, then a track=true replace. Must replace
    exactly once, never ERPERP..."""
    # seed a pre-existing tracked change plus the target phrase
    srv.search_and_replace(
        busy_doc, [{"find": "stable tail", "replace": "steady tail"}],
        track=True, author="Seed Author", max_replacements=None,
    )
    srv.insert_paragraphs(
        busy_doc,
        [{"text": "The enterprise resource planning rollout continues."}],
        location=None,
    )
    r = srv.search_and_replace(
        busy_doc,
        [{"find": "enterprise resource planning", "replace": "ERP"}],
        track=True, author="Live Fix Tester",
    )
    assert r["total"] == 1
    text = srv.get_text(busy_doc, contains="ERP")
    joined = "\n".join(p["text"] for p in text)
    assert "ERPERP" not in joined

    def cleanup(s):
        s.doc.Revisions.RejectAll()
        return {}

    live.run_live(busy_doc, "reject revisions", cleanup)


@live_mark
@needs_word
def test_word_status_reports_dialog_free_and_lock_state(busy_doc):
    """Live checklist item: after the fix batch, a normal session shows
    no pending dialogs, and com_word_status carries the serialization
    block."""
    out = srv.com_word_status()
    assert "pending_dialogs" not in out, (
        f"unexpected dialogs: {out.get('pending_dialogs')}"
    )
    assert out.get("blocked") is not True
    assert "com_serialization" in out
    assert out["com_serialization"]["held"] is False


@live_mark
@needs_word
def test_save_retry_succeeds_under_repeated_saves(busy_doc):
    """The report's 11/11 save failure: repeated com_save_document calls
    (serialized, alerts suppressed, retry with backoff) must all
    succeed with no dialog left behind."""
    srv.search_and_replace(
        busy_doc, [{"find": "paragraph 00", "replace": "paragraph zero"}]
    )
    for _ in range(3):
        out = srv.com_save_document(busy_doc)
        assert "saved" in out
    status = srv.com_word_status()
    assert "pending_dialogs" not in status
