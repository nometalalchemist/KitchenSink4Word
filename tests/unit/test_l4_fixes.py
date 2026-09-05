"""Regressions for the L4 interaction bug-hunt findings (F1-F11).

V2 STAGED REWRITE (Wave E): apply_style moved to range kwargs; everything
else in this file already matches the v2 surface. DO NOT RUN in the wave
phase; the integrator runs the live rounds after applying the briefs.
Imports from test_live_core resolve once this file replaces tests/unit's copy.
"""

import subprocess
import sys
import textwrap

import pytest

import word_mcp.server as srv
from word_mcp.com import live, live_ops
from word_mcp.core.errors import DocumentProtected, UnsupportedStructure

from test_live_core import _word_available, quit_instance_holding

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)

PY = sys.executable


def _open_in_new_word(path):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.Documents.Open(str(path))
    return app


@pytest.fixture()
def plain_doc(tmp_path):
    if not _word_available():
        pytest.skip("Word not available")
    path = tmp_path / "plain.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": f"plain paragraph {i} charlie delta."} for i in range(6)],
        backup=False,
    )
    app = _open_in_new_word(path)
    app = None
    yield str(path)
    import pythoncom

    quit_instance_holding(str(path))
    pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f1_replace_inside_field_result_terminates(tmp_path):
    """Literal replace matching a CITATION field's rendered text must skip
    it (bounded), not hang forever."""
    path = tmp_path / "field.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path), [{"text": "Cited here: "}], backup=False
    )
    srv.manage_source(
        str(path), action="add", tag="hurd99",
        authors=[{"last": "Hurd", "first": "Ian"}],
        title="Legitimacy", year="1999", source_type="JournalArticle",
    )
    srv.insert_citation(
        str(path), tag="hurd99",
        location={"search": {"text": "Cited here: "}},
    )
    app = _open_in_new_word(path)
    app = None
    try:
        r = srv.search_and_replace(
            str(path), [{"find": "Hurd", "replace": "HURDX"}], live="force"
        )
        # L7 parity shape: skip counters are top-level dicts keyed by find
        assert r.get("skipped_inside_fields", {}).get("Hurd", 0) >= 1
        assert "HURDX" not in str(
            srv.get_text(str(path))
        ) or r["replaced"]["Hurd"] == 0
    finally:
        import pythoncom

        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f2_paragraph_index_parity_with_sdt_toc(tmp_path):
    """A TOC (SDT) must not shift live body indices vs the file layer."""
    path = tmp_path / "toc.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "Alpha bravo charlie."}, {"text": "Heading One"},
         {"text": "Delta echo foxtrot."}], backup=False,
    )
    srv.apply_style(
        str(path), style="Heading 1", range={"start": 2, "end": 2},
        backup=False,
    )
    srv.insert_reference_list(str(path), type="toc", backup=False)
    file_read = srv.get_text(str(path), live="off")
    file_idx = {
        p["text"]: p["index"] for p in file_read
        if p.get("index") is not None
    }
    app = _open_in_new_word(path)
    app = None
    try:
        live_read = srv.get_text(str(path), live="force")
        live_idx = {
            p["text"]: p["index"] for p in live_read
            if p.get("index") is not None
        }
        for text, idx in file_idx.items():
            if text in live_idx:
                assert live_idx[text] == idx, f"index parity broke at {text!r}"
        assert "Alpha bravo charlie." in live_idx
    finally:
        import pythoncom

        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f3_f4_protected_doc_typed_refusal(tmp_path):
    """Mutations on a readOnly-protected open doc raise DocumentProtected
    (typed), never silent success or raw com_error; reads still work."""
    path = tmp_path / "prot.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path), [{"text": "protected charlie text."}], backup=False,
    )
    srv.set_document_protection(str(path), protection="readOnly", password="pw1")
    app = _open_in_new_word(path)
    app = None
    try:
        with pytest.raises(DocumentProtected):
            srv.format_text(
                str(path), {"bold": True}, find="charlie", live="force"
            )
        with pytest.raises(DocumentProtected):
            srv.search_and_replace(
                str(path), [{"find": "charlie", "replace": "x"}], live="force"
            )
        with pytest.raises(DocumentProtected):
            srv.live_set_track_changes(str(path), True)
        read = srv.get_text(str(path), live="force")
        assert any(
            "protected charlie" in p["text"] for p in read
        )
    finally:
        import pythoncom

        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f5_host_apartment_survives_live_call(plain_doc):
    """A live call on a host-initialized thread must NOT tear down the
    host's COM apartment or disconnect its proxies."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    host_app = win32com.client.GetActiveObject("Word.Application")
    host_name_before = host_app.Name

    srv.search_and_replace(
        plain_doc, [{"find": "charlie", "replace": "charly"}], live="force"
    )
    srv.search_and_replace(
        plain_doc, [{"find": "charly", "replace": "charlie"}], live="force"
    )

    # the host's proxy must still answer
    assert host_app.Name == host_name_before
    host_app = None
    pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f6_author_effective_diffs_not_position(tmp_path):
    """author_effective must come from the NEW revisions, not whichever
    revision happens to sit last in document order."""
    path = tmp_path / "revs.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "early hotel text."}, {"text": "late india text."}], backup=False,
    )
    # pre-existing tracked change LATE in the document by a distinct author
    srv.search_and_replace(
        str(path), [{"find": "india", "replace": "INDIA"}],
        track=True, author="FileLayer", backup=False, live="off",
    )
    app = _open_in_new_word(path)
    app = None
    try:
        r = srv.search_and_replace(
            str(path), [{"find": "hotel", "replace": "HOTEL"}],
            track=True, author="Zeta", live="force",
        )
        assert r["author_requested"] == "Zeta"
        assert r.get("author_effective") != "FileLayer"
    finally:
        import pythoncom

        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f7_vmerge_table_typed_refusal(tmp_path):
    path = tmp_path / "vmerge.docx"
    srv.create_document(str(path))
    srv.create_table(
        str(path), [["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]], backup=False,
    )
    srv.modify_table_structure(
        str(path), 0, action="merge",
        range={"start_row": 0, "end_row": 1, "start_col": 0, "end_col": 0},
        backup=False,
    )
    app = _open_in_new_word(path)
    app = None
    try:
        with pytest.raises(UnsupportedStructure, match="vertically merged"):
            srv.set_cells(
                str(path), 0, [{"row": 2, "cell": 2, "text": "X"}],
                live="force",
            )
    finally:
        import pythoncom

        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f8_tracked_changes_protection_flagged(tmp_path):
    path = tmp_path / "tcprot.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path), [{"text": "juliet kilo lima."}], backup=False
    )
    srv.set_document_protection(
        str(path), protection="trackedChanges", password="pw2"
    )
    app = _open_in_new_word(path)
    app = None
    try:
        r = srv.search_and_replace(
            str(path), [{"find": "kilo", "replace": "KILO"}], live="force"
        )
        assert r.get("enforced_tracking") is True
    finally:
        import pythoncom

        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f9_read_only_open_flagged(tmp_path):
    import pythoncom
    import win32com.client

    path = tmp_path / "ro.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path), [{"text": "mike november oscar."}], backup=False,
    )
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    # POSITIONAL, not ReadOnly=True (concurrency matrix L2). pywin32's
    # late-bound dispatch cannot resolve named arguments without a type
    # library, so it drops them silently: the document opened READ-WRITE,
    # doc.ReadOnly answered False, the server correctly reported nothing,
    # and the test failed asserting a flag the product had no reason to
    # set. Signature: Open(FileName, ConfirmConversions, ReadOnly).
    app.Documents.Open(str(path), False, True)
    app = None
    try:
        r = srv.search_and_replace(
            str(path), [{"find": "november", "replace": "NOVEMBER"}],
            live="force",
        )
        assert r.get("opened_read_only") is True
    finally:
        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f10_deleted_text_and_section_break_excluded(tmp_path):
    path = tmp_path / "del.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path), [{"text": "ordinary prose with no fields"}], backup=False,
    )
    srv.search_and_replace(
        str(path), [{"find": "ordinary", "replace": "ordinary tracked"}],
        track=True, author="FileLayer", backup=False, live="off",
    )
    app = _open_in_new_word(path)
    try:
        # default view (Simple Markup): deletion text absent from Range.Text
        # and nothing may be wrongly subtracted
        live_read = srv.get_text(str(path), live="force")
        texts = [p["text"] for p in live_read]
        assert any("ordinary tracked prose" in t for t in texts), texts

        # All-Markup view: deletion text IS in Range.Text and must be
        # excluded to match the file layer
        doc = app.Documents(1)
        doc.ActiveWindow.View.RevisionsFilter.Markup = 2  # wdRevisionsMarkupAll
        live_read = srv.get_text(str(path), live="force")
        texts = [p["text"] for p in live_read]
        assert any("ordinary tracked prose" in t for t in texts), texts
        assert not any("ordinaryordinary" in t for t in texts)
        assert not any("\x0c" in t for t in texts)
        doc = None
    finally:
        import pythoncom

        app = None
        quit_instance_holding(str(path))
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_f11_orphaned_undo_record_closed_at_next_session(plain_doc):
    """A crashed client's still-recording undo record is closed before the
    next tool call starts its own record."""
    script = textwrap.dedent(f"""
        import os, pythoncom, win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.GetActiveObject("Word.Application")
        app.UndoRecord.StartCustomRecord("orphaned by crash")
        for d in app.Documents:
            if d.FullName.lower() == r"{plain_doc}".lower():
                d.Content.InsertAfter("crash-edit ")
        os._exit(1)
    """)
    subprocess.run([PY, "-X", "utf8", "-c", script], timeout=60)

    r = srv.search_and_replace(
        plain_doc, [{"find": "delta", "replace": "DELTA"}], live="force"
    )
    assert r["undo_grouped"] is True

    # Inspect OUTSIDE any live session (a session's own record is open
    # while inside it, so run_live cannot observe this)
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("Word.Application")
    undo = app.UndoRecord
    assert bool(undo.IsRecordingCustomRecord) is False
    app = None
    pythoncom.CoUninitialize()
    srv.search_and_replace(
        plain_doc, [{"find": "DELTA", "replace": "delta"}], live="force"
    )
