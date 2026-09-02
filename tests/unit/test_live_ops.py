"""Live routing layer tests (v2.0 L2). V2 STAGED REWRITE (Wave E): calls
updated to the v2 surface (location objects, range shapes). DO NOT RUN in the
wave phase; the integrator runs the live rounds after applying the briefs.
Imports from test_live_core resolve once this file replaces tests/unit's copy.

A document with known structure is built by the FILE-BASED tools, its
file-based reads are captured, then it is opened in a spawned visible Word.
The routed tools must (a) auto-detect the lock and go live, (b) mirror the
file-based semantics, (c) leave the user's Word state untouched.
"""

import pytest

import word_mcp.server as srv
from word_mcp.com import live, live_ops
from word_mcp.core.errors import DocumentLocked, WordMcpError

from test_live_core import _word_available

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)

PARAS = [
    "Introduction",                       # 0  (heading)
    "The alliance framework matters.",    # 1
    "Second paragraph with alliance terms.",  # 2
    "Methods",                            # 3  (heading)
    "Cell data follows in the table.",    # 4
    "Closing paragraph.",                 # 5
]


@pytest.fixture(scope="module")
def routed_doc(tmp_path_factory):
    """(path, file_read): built file-based, then opened in a spawned Word."""
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path_factory.mktemp("live_ops") / "routed.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path), [{"text": t} for t in PARAS], backup=False
    )
    # v2: apply_style takes range/target; [0, 3] is non-contiguous, so one
    # call per paragraph.
    srv.apply_style(
        str(path), style="Heading 1", range={"start": 0, "end": 0},
        backup=False,
    )
    srv.apply_style(
        str(path), style="Heading 1", range={"start": 3, "end": 3},
        backup=False,
    )
    srv.create_table(
        str(path), [["a", "b"], ["c", "d"]], backup=False
    )
    file_read = srv.get_text(str(path), live="off")
    file_outline = srv.get_outline(str(path), live="off")

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.Documents.Open(str(path))
    yield str(path), file_read, file_outline
    app = None
    from test_live_core import quit_instance_holding

    try:
        quit_instance_holding(str(path))
    finally:
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_live_off_still_refuses(routed_doc):
    path, _, _ = routed_doc
    with pytest.raises(DocumentLocked):
        srv.get_text(path, live="off")


@live_mark
@needs_word
def test_get_text_routes_live_and_matches_file_read(routed_doc):
    path, file_read, _ = routed_doc
    result = srv.get_text(path)          # auto → live
    # L9 parity: live returns the FILE-mode flat list shape
    assert isinstance(result, list)
    # empty doc starts with one empty paragraph; both layers must agree on
    # indexing, text, and heading levels for every non-empty paragraph
    file_by_idx = {p["index"]: p for p in file_read}
    for lp in result:
        fp = file_by_idx.get(lp["index"])
        assert fp is not None, f"live index {lp['index']} missing file-side"
        assert lp["text"] == fp["text"]


@live_mark
@needs_word
def test_get_outline_parity(routed_doc):
    path, _, file_outline = routed_doc
    result = srv.get_outline(path)
    # L9 parity: live returns the FILE-mode flat list shape
    assert isinstance(result, list)
    live_heads = [(h["paragraph_index"], h["level"], h["text"])
                  for h in result]
    file_heads = [(h["paragraph_index"], h["level"], h["text"])
                  for h in file_outline]
    assert live_heads == file_heads
    assert all(h["detected_via"] == "heading_style" for h in result)


@live_mark
@needs_word
def test_search_and_replace_routes_live(routed_doc):
    path, _, _ = routed_doc
    r = srv.search_and_replace(
        path, [{"find": "framework", "replace": "structure"}]
    )
    # L7 parity: canonical file shape {replaced, total} + live key
    assert r["live"] is True and r["total"] == 1
    assert r["replaced"] == {"framework": 1}
    text = srv.get_text(path, contains="structure")
    assert any("alliance structure matters" in p["text"] for p in text)
    srv.search_and_replace(path, [{"find": "structure", "replace": "framework"}])


@live_mark
@needs_word
def test_self_referencing_replacement_terminates(routed_doc):
    path, _, _ = routed_doc
    r = srv.search_and_replace(
        path, [{"find": "alliance", "replace": "alliance-x"}]
    )
    assert r["total"] == 2
    back = srv.search_and_replace(
        path, [{"find": "alliance-x", "replace": "alliance"}]
    )
    assert back["total"] == 2


@live_mark
@needs_word
def test_regex_replacement_with_groups(routed_doc):
    path, _, _ = routed_doc
    r = srv.search_and_replace(
        path,
        [{"find": r"(Second) (paragraph)", "replace": r"\2 \1", "regex": True}],
    )
    assert r["total"] == 1
    text = srv.get_text(path, contains="paragraph Second")
    assert text
    srv.search_and_replace(
        path,
        [{"find": r"(paragraph) (Second)", "replace": r"\2 \1", "regex": True}],
    )


@live_mark
@needs_word
def test_max_replacements_aborts_atomically(routed_doc):
    path, _, _ = routed_doc
    with pytest.raises(WordMcpError, match="max_replacements"):
        srv.search_and_replace(
            path,
            [{"find": "paragraph", "replace": "XXX"}],
            max_replacements=1,
        )
    assert not srv.get_text(path, contains="XXX")


@live_mark
@needs_word
def test_insert_and_delete_paragraphs_live(routed_doc):
    path, _, _ = routed_doc
    before = len(srv.get_text(path))
    # v2: the search selector replaces after_anchor (default position: after)
    r = srv.insert_paragraphs(
        path, [{"text": "LIVE INSERTED A"}, {"text": "LIVE INSERTED B"}],
        location={"search": {"text": "Closing paragraph"}},
    )
    assert r["live"] is True and r["inserted"] == 2
    now = srv.get_text(path)
    assert len(now) == before + 2
    idxs = [p["index"] for p in now
            if p["text"].startswith("LIVE INSERTED")]
    assert len(idxs) == 2
    d = srv.delete_paragraphs(path, idxs[0], idxs[1])
    assert d["live"] is True and d["deleted"] == 2
    assert len(srv.get_text(path)) == before


@live_mark
@needs_word
def test_set_cells_live(routed_doc):
    path, _, _ = routed_doc
    r = srv.set_cells(
        path, 0,
        [{"row": 0, "cell": 0, "text": "H1"}, {"row": 0, "cell": 1, "text": "H2"},
         {"row": 1, "cell": 0, "text": "v1"}, {"row": 1, "cell": 1, "text": "v2"}],
    )
    # L7 parity: file-mode key name
    assert r["live"] is True and r["cells_written"] == 4
    found = srv.find_text(path, "v2")
    # L9 parity: flat list; table matches carry table/row/cell addressing
    assert isinstance(found, list)
    assert any(
        m.get("table_index") == 0 and m.get("row") == 1 and m.get("cell") == 1
        for m in found
    )


@live_mark
@needs_word
def test_format_text_live_applies_and_reads_back(routed_doc):
    path, _, _ = routed_doc
    # v2: single-paragraph range replaces paragraph_index
    r = srv.format_text(
        path, {"bold": True, "color": "FF0000"},
        range={"start": 5, "end": 5}, find="Closing",
    )
    assert r["live"] is True

    def check(s):
        paras = live_ops._body_paragraphs(s.doc)
        p = paras[5]
        base = p.Range.Start
        rng = s.doc.Range(base, base + len("Closing"))
        return {"bold": bool(rng.Font.Bold), "color": int(rng.Font.Color)}

    read = live.run_live(path, "verify format", check)
    assert read["bold"] is True
    assert read["color"] == 0x0000FF  # BGR for red


@live_mark
@needs_word
def test_format_text_rejects_unknown_key(routed_doc):
    path, _, _ = routed_doc
    with pytest.raises(WordMcpError, match="allowed"):
        srv.format_text(path, {"bolded": True}, range={"start": 1, "end": 1})


@live_mark
@needs_word
def test_tracked_live_replace_creates_revisions(routed_doc):
    path, _, _ = routed_doc
    r = srv.search_and_replace(
        path, [{"find": "Closing", "replace": "Final"}],
        track=True, author="Live Tester",
    )
    assert r["total"] == 1
    assert r["tracked_as"] == "Live Tester"  # file-mode parity key

    def check(s):
        authors = {rev.Author for rev in s.doc.Revisions}
        return {"count": s.doc.Revisions.Count, "authors": sorted(authors),
                "track_now": bool(s.doc.TrackRevisions)}

    read = live.run_live(path, "verify revisions", check)
    assert read["count"] >= 1
    # Word signed into an Office account overrides UserName; the tool must
    # report the effective author honestly rather than claim the requested one
    assert r["author_requested"] == "Live Tester"
    assert r.get("author_effective") in read["authors"]
    if r.get("author_effective") != "Live Tester":
        assert "author_note" in r
    assert read["track_now"] is False  # guard restored the OFF state

    def cleanup(s):
        s.doc.Revisions.RejectAll()
        return {}

    live.run_live(path, "reject revisions", cleanup)


@live_mark
@needs_word
def test_delete_spanning_table_refused(routed_doc):
    """Body-level indices skip tables, so a start..end range can silently
    cross one; deleting it once deleted the table itself. Must refuse."""
    from word_mcp.core.errors import UnsupportedStructure

    path, _, _ = routed_doc
    paras = srv.get_text(path)
    n = len([p for p in paras if p["index"] is not None])
    # paragraph 5 is before the table, the trailing paragraph is after it
    with pytest.raises(UnsupportedStructure, match="table"):
        srv.delete_paragraphs(path, 5, n - 1)
    info = srv.get_document_info(path)
    assert info["tables"] == 1  # table survived the refusal


@live_mark
@needs_word
def test_document_info_live(routed_doc):
    path, _, _ = routed_doc
    info = srv.get_document_info(path)
    assert info["live"] is True
    assert info["tables"] == 1
    assert info["sections"] >= 1
