"""v2 Phase 5 live guards: the snapshot staleness guard (verify-or-refuse
at the COM boundary when live-branch locations resolved against a stale
disk snapshot) and live heading_level support on insert_paragraphs.

Scenario: a document open in Word with UNSAVED changes that shift
paragraph indices. Text-selector locations resolve against the saved
file, so without the guard the live edit would hit the wrong paragraph;
with it, the call refuses STALE_ANCHOR and saving the document clears
the refusal.
"""

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import StaleAnchor

from test_live_core import _word_available, quit_instance_holding

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)

TARGET = "Guarded target paragraph with unique payload text."


@pytest.fixture(scope="module")
def guard_doc(tmp_path_factory):
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path_factory.mktemp("live_guards") / "guards.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": f"Filler paragraph number {i}."} for i in range(5)]
        + [{"text": TARGET}]
        + [{"text": f"Trailing paragraph number {i}."} for i in range(3)],
        backup=False,
    )
    # durable anchors on disk for the apply_edits leg
    srv.get_document_view(str(path), stamp_anchors=True)
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.Documents.Open(str(path))
    yield str(path), app
    app = None
    try:
        quit_instance_holding(str(path))
    finally:
        pythoncom.CoUninitialize()


def _make_dirty(app, path):
    """Insert an unsaved paragraph at the very top of the open document,
    shifting every body index by one."""
    doc = app.Documents(1)
    doc.Paragraphs(1).Range.InsertBefore("Unsaved shift paragraph.\r")
    assert not doc.Saved
    return doc


@live_mark
@needs_word
def test_stale_search_selector_refuses(guard_doc):
    path, app = guard_doc
    _make_dirty(app, path)
    with pytest.raises(StaleAnchor) as exc:
        srv.set_paragraph_text(
            path,
            {"search": {"text": "unique payload"}},
            "This must not land.",
        )
    assert "UNSAVED" in str(exc.value)
    # nothing changed: the target text is still present live
    live = srv.get_text(path, contains="unique payload")
    assert len(live) == 1 and TARGET in live[0]["text"]


@live_mark
@needs_word
def test_stale_apply_edits_set_text_refuses(guard_doc):
    path, app = guard_doc
    doc = app.Documents(1)
    assert not doc.Saved  # still dirty from the previous test
    view = srv.get_document_view(path)
    assert view.get("live") is True and "SAVED" in view.get("note", "")
    anchor = next(
        line.split("]")[0].lstrip("[")
        for line in view["view"].splitlines()
        if "unique payload" in line and line.startswith("[")
    )
    with pytest.raises(StaleAnchor):
        srv.apply_edits(
            path,
            [{"op": "set_text", "anchor": anchor,
              "text": "This must not land either."}],
        )
    live = srv.get_text(path, contains="unique payload")
    assert len(live) == 1 and TARGET in live[0]["text"]


@live_mark
@needs_word
def test_saving_clears_the_refusal(guard_doc):
    path, app = guard_doc
    doc = app.Documents(1)
    doc.Save()
    assert doc.Saved
    r = srv.set_paragraph_text(
        path,
        {"search": {"text": "unique payload"}},
        "Replacement landed after save.",
    )
    assert r.get("live") is True
    live = srv.get_text(path, contains="Replacement landed")
    assert len(live) == 1


@live_mark
@needs_word
def test_heading_level_inserts_live(guard_doc):
    path, app = guard_doc
    doc = app.Documents(1)
    doc.Save()
    r = srv.insert_paragraphs(
        path,
        [{"text": "Live Heading Two", "heading_level": 2}],
    )
    assert r.get("live") is True and r.get("inserted") == 1
    # built-in Heading 2 style implies outline level 2 (locale-free check)
    found = [
        p for p in doc.Paragraphs
        if "Live Heading Two" in p.Range.Text
    ]
    assert found and found[0].OutlineLevel == 2


@live_mark
@needs_word
def test_paragraph_index_keeps_v1_trust(guard_doc):
    """Raw {paragraph: N} addressing carries no snapshot text dependency:
    it stays index-trusting even on a dirty document (the v1 contract)."""
    path, app = guard_doc
    doc = app.Documents(1)
    doc.Save()
    _make_dirty(app, path)
    r = srv.set_paragraph_text(
        path, {"paragraph": 0}, "Index-addressed overwrite of paragraph 0."
    )
    assert r.get("live") is True
    doc.Save()
