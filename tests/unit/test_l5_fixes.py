"""Regressions for the L5 insane-stress findings.

V2 STAGED REWRITE (Wave E): apply_style moved to range kwargs; format_text
paragraph_index moved to single-paragraph range. DO NOT RUN in the wave
phase; the integrator runs the live rounds after applying the briefs.
Imports from test_live_core resolve once this file replaces tests/unit's copy.
"""

import pytest

import word_mcp.server as srv
from word_mcp.com import live
from word_mcp.core.errors import UnsupportedStructure, WordMcpError

from test_live_core import _word_available, quit_instance_holding

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)


# ------------------------------------------------------------- file-layer


def _table_doc(tmp_path, name="t.docx"):
    path = tmp_path / name
    srv.create_document(str(path))
    srv.create_table(
        str(path), [["a", "b"], ["c", "d"], ["e", "f"]], backup=False,
    )
    return str(path)


def test_f5_file_set_cells_refuses_vmerge_continuation(tmp_path):
    path = _table_doc(tmp_path)
    srv.modify_table_structure(
        str(path), 0, action="merge",
        range={"start_row": 0, "end_row": 1, "start_col": 0, "end_col": 0},
        backup=False,
    )
    with pytest.raises(UnsupportedStructure, match="CONTINUATION"):
        srv.set_cells(
            str(path), 0, [{"row": 1, "cell": 0, "text": "invisible"}],
            backup=False, live="off",
        )
    # the restart cell still writes fine
    r = srv.set_cells(
        str(path), 0, [{"row": 0, "cell": 0, "text": "visible"}],
        backup=False, live="off",
    )
    assert r["cells_written"] == 1


def test_f6_file_find_text_refuses_empty_and_empty_matching(tmp_path):
    path = tmp_path / "f.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(str(path), [{"text": "abc"}],
                          backup=False)
    with pytest.raises(WordMcpError, match="non-empty"):
        srv.find_text(str(path), "", live="off")
    with pytest.raises(WordMcpError, match="empty string"):
        srv.find_text(str(path), "z*", regex=True, live="off")


def test_f7_file_get_text_slice_excludes_sdt(tmp_path):
    path = tmp_path / "s.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "Body zero"}, {"text": "Heading"}, {"text": "Body two"}], backup=False,
    )
    srv.apply_style(
        str(path), style="Heading 1", range={"start": 1, "end": 1},
        backup=False,
    )
    srv.insert_reference_list(str(path), type="toc", backup=False)
    whole = srv.get_text(str(path), live="off")
    assert any(p.get("in_sdt") for p in whole)
    sliced = srv.get_text(str(path), start=0, end=2, live="off")
    assert not any(p.get("in_sdt") for p in sliced)
    assert [p["index"] for p in sliced] == [0, 1]  # end is exclusive


# ------------------------------------------------------------------ live


@pytest.fixture()
def live_doc(tmp_path):
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path / "live.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": f"live paragraph {i} qbert."} for i in range(8)], backup=False,
    )
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.Documents.Open(str(path))
    app = None
    yield str(path)
    quit_instance_holding(str(path))


@live_mark
@needs_word
def test_f1_inactive_document_reports_ungrouped(live_doc, tmp_path):
    """With a second doc ACTIVE, edits on the background doc must report
    undo_grouped=false (Word can only group into the active document)."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("Word.Application")
    second = app.Documents.Add()   # becomes the active document
    second.Content.InsertAfter("decoy active document")
    app = None
    try:
        r = srv.search_and_replace(
            live_doc, [{"find": "qbert", "replace": "QBERT"}], live="force"
        )
        assert r["total"] == 8  # L7 parity: file-mode key
        assert r["undo_grouped"] is False
        assert "undo_note" in r
        srv.search_and_replace(
            live_doc, [{"find": "QBERT", "replace": "qbert"}], live="force"
        )
    finally:
        import pythoncom as pc

        app2 = win32com.client.GetActiveObject("Word.Application")
        for d in app2.Documents:
            if not d.Path:          # the unsaved decoy
                d.Close(SaveChanges=0)
                break
        app2 = None


@live_mark
@needs_word
def test_f1_active_document_still_groups(live_doc):
    r = srv.search_and_replace(
        live_doc, [{"find": "qbert", "replace": "QBERT"}], live="force"
    )
    assert r["undo_grouped"] is True

    def undo(s):
        s.doc.Undo(1)
        return {"text": s.doc.Content.Text[:40]}

    read = live.run_live(live_doc, "undo", undo)
    assert "qbert" in read["text"]


@live_mark
@needs_word
def test_f2_zero_length_regex_refused_live(live_doc):
    before = [p["text"] for p in srv.get_text(live_doc)]
    with pytest.raises(WordMcpError, match="empty string"):
        srv.search_and_replace(
            live_doc, [{"find": "z*", "replace": "Q", "regex": True}],
            live="force",
        )
    after = [p["text"] for p in srv.get_text(live_doc)]
    assert before == after  # nothing was shredded


@live_mark
@needs_word
def test_f4_enforced_tracking_uses_deletion_counter(live_doc):
    srv.live_set_track_changes(live_doc, True)
    try:
        r = srv.search_and_replace(
            live_doc, [{"find": "paragraph 3", "replace": "PARAGRAPH 3"}],
            track=True, author="L5",
        )
        # no fields exist in this doc: the field counter must stay silent
        assert "skipped_inside_fields" not in r
    finally:
        def cleanup(s):
            s.doc.Revisions.RejectAll()
            return {}

        live.run_live(live_doc, "cleanup", cleanup)
        srv.live_set_track_changes(live_doc, False)


@live_mark
@needs_word
def test_f9_live_format_text_extended_keys(live_doc):
    r = srv.format_text(
        live_doc,
        {"small_caps": True, "char_spacing_pt": 1.5},
        range={"start": 1, "end": 1}, find="live",
    )
    assert r["live"] is True

    def check(s):
        from word_mcp.com import live_ops

        p = live_ops._body_paragraphs(s.doc)[1]
        rng = s.doc.Range(p.Range.Start, p.Range.Start + 4)
        return {"sc": bool(rng.Font.SmallCaps), "sp": float(rng.Font.Spacing)}

    read = live.run_live(live_doc, "verify", check)
    assert read["sc"] is True and read["sp"] == 1.5
    # WS-L: language keys now WORK live via the LCID map. Latin-script tags
    # go to Range.LanguageID; East-Asian tags (ko/ja/zh) go to
    # LanguageIDFarEast (Word silently ignores them in the latin slot).
    r2 = srv.format_text(
        live_doc, {"language": "fr-FR", "east_asian_language": "ja-JP"},
        range={"start": 1, "end": 1}, live="force",
    )
    assert r2["live"] is True

    def check_lang(s):
        from word_mcp.com import live_ops

        p = live_ops._body_paragraphs(s.doc)[1]
        rng = s.doc.Range(p.Range.Start, p.Range.End - 1)
        return {
            "latin": int(rng.LanguageID),
            "far_east": int(rng.LanguageIDFarEast),
        }

    langs = live.run_live(live_doc, "verify lang", check_lang)
    assert langs["latin"] == 1036      # fr-FR
    assert langs["far_east"] == 1041   # ja-JP
    with pytest.raises(WordMcpError, match="no live LCID"):
        srv.format_text(
            live_doc, {"language": "xx-XX"}, range={"start": 1, "end": 1},
            live="force",
        )
    # ko-KR in the latin slot would be silently ignored by Word; the tool
    # must refuse with routing guidance rather than fake success
    with pytest.raises(WordMcpError, match="east_asian_language"):
        srv.format_text(
            live_doc, {"language": "ko-KR"}, range={"start": 1, "end": 1},
            live="force",
        )
