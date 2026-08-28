"""Live-specific tools (v2.0 L3): cursor insert, scroll-to, track toggle."""

import pytest

import word_mcp.server as srv
from word_mcp.com import live, live_ops
from word_mcp.core.errors import TargetNotFound, WordMcpError

from test_live_core import _word_available

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)


@pytest.fixture(scope="module")
def cursor_doc(tmp_path_factory):
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path_factory.mktemp("live3") / "cursor.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": f"Paragraph number {i} with filler text."} for i in range(40)],
        at_end=True,
        backup=False,
    )
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    doc = app.Documents.Open(str(path))
    # park the "user's cursor" at the start of paragraph 10's text
    target = doc.Paragraphs(11).Range
    doc.Range(target.Start, target.Start).Select()
    doc = None
    yield str(path)
    app = None
    from test_live_core import quit_instance_holding

    try:
        quit_instance_holding(str(path))
    finally:
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_insert_at_cursor(cursor_doc):
    r = srv.live_insert_at_cursor(cursor_doc, "CURSOR-INSERTED ")
    assert r["live"] is True and r["chars"] == len("CURSOR-INSERTED ")
    text = srv.get_text(cursor_doc, contains="CURSOR-INSERTED")
    assert len(text) == 1
    assert text[0]["text"].startswith(
        "CURSOR-INSERTED Paragraph number 10"
    )

    def check_selection(s):
        sel = s.app.Selection
        return {"sel_start": sel.Range.Start, "sel_text_after": s.doc.Range(
            sel.Range.Start, sel.Range.Start + 16).Text}

    # A collapsed cursor stays at the insertion boundary (Word treats it
    # like a bookmark): still exactly where the user left it, with the
    # inserted text flowing after it. It must NOT have jumped elsewhere.
    read = live.run_live(cursor_doc, "check sel", check_selection)
    assert read["sel_start"] == r["inserted_at"]
    assert read["sel_text_after"].startswith("CURSOR-INSERTED")

    def undo(s):
        s.doc.Undo(1)
        return {}

    live.run_live(cursor_doc, "undo", undo)


@live_mark
@needs_word
def test_scroll_to_find_and_index(cursor_doc):
    r = srv.live_scroll_to(cursor_doc, find="Paragraph number 35")
    assert r["live"] is True and r["found"] == "Paragraph number 35"
    assert r["page"] >= 1
    r2 = srv.live_scroll_to(cursor_doc, paragraph_index=1)
    assert r2["paragraph"] == 1
    with pytest.raises(TargetNotFound):
        srv.live_scroll_to(cursor_doc, find="zzz_not_present")
    with pytest.raises(WordMcpError):
        srv.live_scroll_to(cursor_doc)  # neither target given


@live_mark
@needs_word
def test_set_track_changes_persists(cursor_doc):
    r = srv.live_set_track_changes(cursor_doc, True)
    assert r["track_changes"] is True and r["was"] is False

    def check(s):
        return {"on": bool(s.doc.TrackRevisions)}

    assert live.run_live(cursor_doc, "check", check)["on"] is True
    r2 = srv.live_set_track_changes(cursor_doc, False)
    assert r2["was"] is True
    assert live.run_live(cursor_doc, "check", check)["on"] is False


@live_mark
@needs_word
def test_com_word_status_extended(cursor_doc):
    status = srv.com_word_status()
    assert status["interactive_state"] == "ready"
    entry = next(
        d for d in status["open_documents"]
        if d["path"].lower() == cursor_doc.lower()
    )
    assert "dirty" in entry and "autosave_on" in entry


@live_mark
@needs_word
def test_word_live_repair_reports(cursor_doc):
    report = srv.word_live_repair()
    assert report["word_running"] is True
    assert isinstance(report["actions"], list)
