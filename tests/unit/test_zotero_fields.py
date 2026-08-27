"""Zotero field-code preservation (community feedback, 2026-08-28).

Zotero's Word integration stores citations as `ADDIN ZOTERO_ITEM
CSL_CITATION {json}` fields and the bibliography as `ADDIN ZOTERO_BIBL`.
Structural edits that damage these codes silently break the user's citation
management. Both layers must leave them intact through nearby edits.
"""

import shutil

import pytest

import word_mcp.server as srv
from word_mcp.core.package import DocxPackage

from test_live_core import _word_available, quit_instance_holding

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)

ZOTERO_CODE_MARK = "ZOTERO_ITEM CSL_CITATION"
_WD_FIELD_ADDIN = 81


def _zotero_codes(path: str) -> list[str]:
    """All ADDIN ZOTERO field instruction texts in document.xml. Parses a
    read-copy so it works while the doc is still open (locked) in Word."""
    import tempfile
    from pathlib import Path

    copy_path = Path(tempfile.mkdtemp()) / Path(path).name
    shutil.copy(path, copy_path)
    pkg = DocxPackage(copy_path)
    root = pkg.root()
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    codes = []
    current = None
    for el in root.iter():
        if el.tag == f"{ns}fldChar":
            t = el.get(f"{ns}fldCharType")
            if t == "begin":
                current = []
            elif t == "end" and current is not None:
                joined = "".join(current)
                if "ZOTERO" in joined:
                    codes.append(joined.strip())
                current = None
        elif el.tag == f"{ns}instrText" and current is not None:
            current.append(el.text or "")
    return codes


@pytest.fixture(scope="module")
def zotero_doc(tmp_path_factory):
    """A doc with a real ADDIN ZOTERO_ITEM field, built via Word COM."""
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path_factory.mktemp("zotero") / "cited.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [
            {"text": "Before the citation paragraph."},
            {"text": "The alliance argument rests on "},
            {"text": "After the citation paragraph."},
        ],
        at_end=True,
        backup=False,
    )
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    doc = app.Documents.Open(str(path))
    target = None
    for p in doc.Paragraphs:
        if "alliance argument" in p.Range.Text:
            target = p.Range
            break
    assert target is not None, "citation anchor paragraph not found"
    point = doc.Range(target.End - 1, target.End - 1)
    field = doc.Fields.Add(
        point,
        _WD_FIELD_ADDIN,
        'ZOTERO_ITEM CSL_CITATION {"citationID":"abc123",'
        '"properties":{"formattedCitation":"(Hurd, 1999)"},'
        '"citationItems":[{"id":1}]} CSL_CITATION',
        False,
    )
    field.Result.Text = "(Hurd, 1999)"
    doc.Save()
    doc = None
    yield str(path), app
    app = None
    quit_instance_holding(str(path))
    pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_live_edits_preserve_zotero_field(zotero_doc):
    path, _ = zotero_doc
    # live edits AROUND the citation (doc is open in Word)
    srv.search_and_replace(path, [{"find": "alliance", "replace": "coalition"}])
    srv.insert_paragraphs(
        path, [{"text": "Inserted near the citation."}],
        after_anchor="Before the citation",
    )
    srv.format_text(path, {"bold": True}, find="After the citation")
    srv.search_and_replace(path, [{"find": "coalition", "replace": "alliance"}])
    srv.com_save_open_document(path)
    codes = _zotero_codes(path)
    assert len(codes) == 1
    assert ZOTERO_CODE_MARK in codes[0]
    assert '"citationID":"abc123"' in codes[0]


@live_mark
@needs_word
def test_file_edits_preserve_zotero_field(zotero_doc, tmp_path):
    path, _ = zotero_doc
    # copy the SAVED doc and run file-based structural edits on the copy
    work = tmp_path / "cited_copy.docx"
    shutil.copy(path, work)
    srv.search_and_replace(
        str(work), [{"find": "argument", "replace": "claim"}],
        backup=False, live="off",
    )
    srv.insert_paragraphs(
        str(work), [{"text": "New paragraph after everything."}],
        at_end=True, backup=False, live="off",
    )
    codes = _zotero_codes(str(work))
    assert len(codes) == 1
    assert ZOTERO_CODE_MARK in codes[0]
    assert '"citationID":"abc123"' in codes[0]


@live_mark
@needs_word
def test_live_delete_of_citation_paragraph_removes_field_cleanly(zotero_doc):
    path, _ = zotero_doc
    # deleting the WHOLE citation-bearing paragraph is a legitimate edit; the
    # field must go with it (no orphaned half-field), other codes untouched
    paras = srv.get_text(path)["paragraphs"]
    idx = next(
        p["index"] for p in paras if "alliance argument" in p["text"]
    )
    srv.delete_paragraphs(path, idx, idx)
    srv.com_save_open_document(path)
    assert _zotero_codes(path) == []
    # restore for other tests (module fixture is shared)
    srv.insert_paragraphs(
        path, [{"text": "The alliance argument rests on (plain now)."}],
        after_anchor="Before the citation",
    )
    srv.com_save_open_document(path)
