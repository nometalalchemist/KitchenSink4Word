"""Workstream L regressions: the Ch5 live-ops bug batch (L1-L9 + sweep).

V2 STAGED REWRITE (Wave E): replace_paragraph_text renamed to
set_paragraph_text (location object), apply_style moved to range kwargs. DO
NOT RUN in the wave phase; the integrator runs the live rounds after applying
the briefs. Imports from test_live_core resolve once this file replaces
tests/unit's copy.

File-mode tests (LCID map, outlineLvl outline detection) run everywhere.
The live-marked tests spawn their OWN Word instance via the standard fixture
infrastructure (DispatchEx + ROT-moniker quit) and never touch documents
they did not create.
"""

import pytest

import word_mcp.server as srv
from word_mcp.com import live, live_ops
from word_mcp.core.errors import DocumentLocked, WordMcpError
from word_mcp.core.package import DocxPackage, qn

from test_live_core import _word_available, quit_instance_holding

live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)

PARA_CURLY = "The Delta Model’s core claim holds."
LONG_SENT = (
    "This is a deliberately long paragraph engineered to exceed the two "
    "hundred and fifty five character ceiling of the COM Find interface by "
    "a comfortable margin, because the chunked long-string pathway must "
    "locate it via a prefix search, extend the located range to the full "
    "find length, verify the complete match character for character, and "
    "then replace it through direct text assignment without ever raising "
    "the string parameter too long error that used to crash this tool."
)
LONG_REPLACEMENT = (
    "This replacement paragraph is likewise far beyond the two hundred and "
    "fifty five character ceiling of the COM Find and Replacement text "
    "properties, which proves the replacement side of the chunked pathway "
    "never routes through Word's Find machinery at all and instead assigns "
    "the text directly onto the verified range, where no such length "
    "limitation exists, completing the round trip both ways without error."
)


# ------------------------------------------------------------ pure python


def test_lcid_map_known_tags():
    assert live_ops._lcid("ko-KR") == 1042
    assert live_ops._lcid("KO-kr") == 1042  # case-insensitive
    assert live_ops._lcid("en-US") == 1033
    assert live_ops._lcid("ja-JP") == 1041


def test_lcid_map_unknown_tag_typed_error():
    with pytest.raises(WordMcpError, match="no live LCID"):
        live_ops._lcid("xx-XX")


def test_long_sentences_exceed_com_limit():
    # guard: the L5 fixtures must actually exercise the >255 path
    assert len(LONG_SENT) > live_ops._FIND_TEXT_LIMIT
    assert len(LONG_REPLACEMENT) > live_ops._FIND_TEXT_LIMIT


# ------------------------------------------------- L8 file-mode outline


def _new_doc(tmp_path, name, texts):
    path = tmp_path / name
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path), [{"text": t} for t in texts], backup=False,
        live="off",
    )
    return str(path)


def _apply_para_style(path, index, style):
    """v2 apply_style on a single paragraph."""
    srv.apply_style(
        path, style=style, range={"start": index, "end": index}, backup=False
    )


def _add_style(path, style_id, outline_val=None, based_on=None):
    """Append a paragraph style to styles.xml (raw surgery, test-only)."""
    from lxml import etree

    pkg = DocxPackage(path)
    root = pkg.root("word/styles.xml")
    st = etree.SubElement(root, qn("w:style"))
    st.set(qn("w:type"), "paragraph")
    st.set(qn("w:styleId"), style_id)
    name = etree.SubElement(st, qn("w:name"))
    name.set(qn("w:val"), style_id)
    if based_on is not None:
        b = etree.SubElement(st, qn("w:basedOn"))
        b.set(qn("w:val"), based_on)
    if outline_val is not None:
        ppr = etree.SubElement(st, qn("w:pPr"))
        ol = etree.SubElement(ppr, qn("w:outlineLvl"))
        ol.set(qn("w:val"), str(outline_val))
    pkg.mark_dirty("word/styles.xml")
    pkg.save(do_backup=False)


def test_outline_detects_direct_outlinelvl(tmp_path):
    """The NSU-template case: Normal-styled paragraph + direct w:outlineLvl
    must appear in get_outline, labeled outline_level (L8)."""
    path = _new_doc(tmp_path, "direct.docx", ["Chapter One", "Body text."])
    srv.set_paragraph_format(path, [0], {"outline_level": 0}, backup=False)
    outline = srv.get_outline(path, live="off")
    assert [(h["text"], h["level"], h["detected_via"]) for h in outline] == [
        ("Chapter One", 1, "outline_level")
    ]


def test_outline_detects_style_outlinelvl(tmp_path):
    path = _new_doc(tmp_path, "styled.docx", ["Section Head", "Body."])
    _add_style(path, "NsuHead", outline_val=1)
    _apply_para_style(path, 0, "NsuHead")
    outline = srv.get_outline(path, live="off")
    assert [(h["text"], h["level"], h["detected_via"]) for h in outline] == [
        ("Section Head", 2, "outline_level")
    ]


def test_outline_detects_basedon_chain_outlinelvl(tmp_path):
    """outlineLvl inherited through the style's basedOn chain (grandparent
    carries it): Word resolves the effective value this way and so must
    get_outline (L8)."""
    path = _new_doc(tmp_path, "chain.docx", ["Deep Head", "Body."])
    _add_style(path, "BaseHead", outline_val=2)
    _add_style(path, "MidHead", based_on="BaseHead")
    _add_style(path, "LeafHead", based_on="MidHead")
    _apply_para_style(path, 0, "LeafHead")
    outline = srv.get_outline(path, live="off")
    assert [(h["text"], h["level"], h["detected_via"]) for h in outline] == [
        ("Deep Head", 3, "outline_level")
    ]


def test_outline_lvl9_is_body_text(tmp_path):
    """w:outlineLvl val=9 is the explicit body-text value; it must NOT be
    reported as a level-10 heading, and a direct val=9 must override a
    heading style's level."""
    from lxml import etree

    path = _new_doc(tmp_path, "lvl9.docx", ["Not a heading", "Demoted head"])
    _apply_para_style(path, 1, "Heading 1")
    pkg = DocxPackage(path)
    for idx in (0, 1):
        p = [
            el for el in pkg.body().findall(qn("w:p"))
        ][idx]
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            p.insert(0, ppr)
        ol = etree.SubElement(ppr, qn("w:outlineLvl"))
        ol.set(qn("w:val"), "9")
    pkg.mark_dirty()
    pkg.save(do_backup=False)
    assert srv.get_outline(path, live="off") == []


def test_outline_heading_style_still_detected(tmp_path):
    path = _new_doc(tmp_path, "head.docx", ["Real Heading", "Body."])
    _apply_para_style(path, 0, "Heading 2")
    outline = srv.get_outline(path, live="off")
    assert [(h["text"], h["level"], h["detected_via"]) for h in outline] == [
        ("Real Heading", 2, "heading_style")
    ]


def test_word_count_sections_from_outlinelvl_headings(tmp_path):
    """File word_count's per-section logic must see outlineLvl-based
    headings (it shares the outline resolver with get_outline)."""
    path = _new_doc(
        tmp_path, "wc.docx",
        ["Template Chapter", "four words of body", "more body words here"],
    )
    _add_style(path, "NsuHead", outline_val=0)
    _apply_para_style(path, 0, "NsuHead")
    wc = srv.word_count(path, live="off")
    assert len(wc["sections"]) == 1
    sec = wc["sections"][0]
    assert sec["heading"] == "Template Chapter" and sec["level"] == 1
    assert sec["words"] == 8


# ------------------------------------------------------------------ live


@pytest.fixture(scope="module")
def ws_l_doc(tmp_path_factory):
    """Built file-based (curly quote, long paragraph, both heading systems,
    table, threaded comment), file reads captured, then opened in a spawned
    Word instance of our own."""
    if not _word_available():
        pytest.skip("Word not available")
    import pythoncom
    import win32com.client

    path = tmp_path_factory.mktemp("ws_l") / "ws_l.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [
            {"text": "Intro heading"},
            {"text": PARA_CURLY},
            {"text": LONG_SENT},
            {"text": "Template heading"},
            {"text": "Closing text paragraph."},
        ], backup=False, live="off",
    )
    read = srv.get_text(str(path), live="off")
    idx = {p["text"]: p["index"] for p in read if p["index"] is not None}
    srv.apply_style(
        str(path), style="Heading 1",
        range={"start": idx["Intro heading"], "end": idx["Intro heading"]},
        backup=False,
    )
    srv.set_paragraph_format(
        str(path), [idx["Template heading"]], {"outline_level": 0},
        backup=False,
    )
    srv.create_table(str(path), [["a", "b"], ["c", "d"]],
                     backup=False)
    srv.manage_comment(
        str(path), action="add",
        location={"search": {"text": "Closing text"}},
        text="First comment", author="Tester A", backup=False,
    )
    first = srv.get_comments(str(path), live="off")
    srv.manage_comment(
        str(path), action="reply", comment_id=first[0]["id"],
        text="A reply", author="Tester B", backup=False,
    )
    captured = {
        "path": str(path),
        "comments": srv.get_comments(str(path), live="off"),
        "outline": srv.get_outline(str(path), live="off"),
        "word_count": srv.word_count(str(path), live="off"),
        "text": srv.get_text(str(path), live="off"),
    }

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = True
    app.Documents.Open(str(path))
    app = None
    yield captured
    try:
        quit_instance_holding(str(path))
    finally:
        pythoncom.CoUninitialize()


@live_mark
@needs_word
def test_l1_l2_live_word_count_matches_words_own_number(ws_l_doc):
    path = ws_l_doc["path"]

    def check(s):
        return {
            "stats": int(s.doc.ComputeStatistics(0)),
            "raw_words_count": int(s.doc.Words.Count),
        }

    read = live.run_live(path, "wc oracle", check)
    wc = srv.word_count(path)          # auto → live (was a refusal, L2)
    assert wc["live"] is True
    assert wc["totals"]["words"] == read["stats"]
    # this doc has punctuation, so the old Words.Count number IS inflated;
    # the tool must not report it (L1)
    assert read["raw_words_count"] > read["stats"]
    assert wc["totals"]["words"] != read["raw_words_count"]
    # file-mode shape parity
    assert set(wc["totals"]) >= {
        "words", "characters", "paragraphs", "tables", "cjk_chars",
    }
    assert wc["counting"] == "spaces"
    # per-section attribution mirrors the file-mode sections
    live_heads = {s["heading"] for s in wc["sections"]}
    file_heads = {s["heading"] for s in ws_l_doc["word_count"]["sections"]}
    assert live_heads == file_heads
    intro = next(s for s in wc["sections"] if s["heading"] == "Intro heading")
    assert intro["words"] > 0
    assert "sections" not in srv.word_count(path, by_section=False)


@live_mark
@needs_word
def test_l1_get_document_info_words_accurate(ws_l_doc):
    path = ws_l_doc["path"]
    info = srv.get_document_info(path)
    assert info["live"] is True

    def check(s):
        return {
            "stats": int(s.doc.ComputeStatistics(0)),
            "raw": int(s.doc.Words.Count),
        }

    read = live.run_live(path, "info oracle", check)
    assert info["words"] == read["stats"]
    assert info["words"] < read["raw"]
    # L7 parity keys shared with file mode
    for key in ("path", "paragraphs", "tables", "sections", "footnotes",
                "endnotes", "comments", "revisions"):
        assert key in info


@live_mark
@needs_word
def test_l3_get_comments_live_parity(ws_l_doc):
    path = ws_l_doc["path"]
    result = srv.get_comments(path)    # auto → live (was a refusal)
    assert isinstance(result, list)
    file_comments = ws_l_doc["comments"]
    assert len(result) == len(file_comments) == 2
    for entry in result:
        assert set(entry) >= {
            "id", "author", "initials", "date", "text", "anchored_text",
            "resolved", "reply_to",
        }
    by_author = {e["author"]: e for e in result}
    assert by_author["Tester A"]["text"] == "First comment"
    assert by_author["Tester A"]["reply_to"] is None
    assert "Closing text" in by_author["Tester A"]["anchored_text"]
    assert by_author["Tester B"]["text"] == "A reply"
    assert by_author["Tester B"]["reply_to"] == by_author["Tester A"]["id"]
    assert by_author["Tester A"]["resolved"] is False
    # author filter
    only_a = srv.get_comments(path, author="Tester A")
    assert [e["author"] for e in only_a] == ["Tester A"]


@live_mark
@needs_word
def test_l3_diagnose_document_honest_refusal(ws_l_doc):
    path = ws_l_doc["path"]
    with pytest.raises(DocumentLocked, match="by design"):
        srv.diagnose_document(path)


@live_mark
@needs_word
def test_l4_set_paragraph_text_live(ws_l_doc):
    path = ws_l_doc["path"]
    paras = srv.get_text(path)
    target = next(p for p in paras if p["text"] == "Closing text paragraph.")
    r = srv.set_paragraph_text(
        path, location={"paragraph": target["index"]},
        new_text="Rewritten closing paragraph.",
    )
    assert r["live"] is True
    assert r["replaced_paragraph"] == target["index"]
    after = srv.get_text(path)
    entry = next(p for p in after if p["index"] == target["index"])
    assert entry["text"] == "Rewritten closing paragraph."
    assert entry["style"] == target["style"]  # paragraph mark untouched
    # restore for the other tests
    srv.set_paragraph_text(
        path, location={"paragraph": target["index"]},
        new_text=target["text"],
    )


@live_mark
@needs_word
def test_l5_long_find_replace_roundtrip(ws_l_doc):
    path = ws_l_doc["path"]
    r = srv.search_and_replace(
        path, [{"find": LONG_SENT, "replace": LONG_REPLACEMENT}]
    )
    assert r["live"] is True and r["total"] == 1
    assert any(
        LONG_REPLACEMENT in p["text"] for p in srv.get_text(path)
    )
    back = srv.search_and_replace(
        path, [{"find": LONG_REPLACEMENT, "replace": LONG_SENT}]
    )
    assert back["total"] == 1
    assert any(LONG_SENT in p["text"] for p in srv.get_text(path))


@live_mark
@needs_word
def test_l5_long_find_absent_and_prefix_only(ws_l_doc):
    path = ws_l_doc["path"]
    # not present at all: no crash, zero replacements
    absent = "Z" * 300
    r = srv.search_and_replace(path, [{"find": absent, "replace": "x"}])
    assert r["total"] == 0
    # prefix matches but the full string does not: must verify and skip
    prefix_only = LONG_SENT[:260] + " ENTIRELY DIFFERENT TAIL CONTENT HERE."
    r2 = srv.search_and_replace(path, [{"find": prefix_only, "replace": "x"}])
    assert r2["total"] == 0
    assert any(LONG_SENT in p["text"] for p in srv.get_text(path))


@live_mark
@needs_word
def test_l5_scroll_to_long_find_typed_error(ws_l_doc):
    with pytest.raises(WordMcpError, match="at most"):
        srv.live_scroll_to(ws_l_doc["path"], find="q" * 300)


@live_mark
@needs_word
def test_l6_curly_apostrophe_preserved_roundtrip(ws_l_doc):
    """The document stores U+2019; live get_text must return U+2019 (never a
    normalized straight apostrophe), and a find string copied from that
    output must actually match (the user's failing loop)."""
    path = ws_l_doc["path"]
    paras = srv.get_text(path)
    entry = next(p for p in paras if "Delta Model" in p["text"])
    assert "’" in entry["text"], (
        "live get_text lost the curly apostrophe: %r" % entry["text"]
    )
    assert "Model’s" in entry["text"]
    assert "Model's" not in entry["text"]  # no straight-apostrophe variant
    # find_text with the exact returned text must match
    found = srv.find_text(path, "Model’s")
    assert isinstance(found, list) and len(found) == 1
    # and search_and_replace built from get_text output must NOT no-op
    r = srv.search_and_replace(
        path, [{"find": "Model’s", "replace": "MODELXQ’s"}]
    )
    assert r["total"] == 1
    back = srv.search_and_replace(
        path, [{"find": "MODELXQ’s", "replace": "Model’s"}]
    )
    assert back["total"] == 1


@live_mark
@needs_word
def test_l7_l9_schema_parity_shapes(ws_l_doc):
    path = ws_l_doc["path"]
    # search_and_replace: file shape + live key; old live keys retired
    r = srv.search_and_replace(
        path, [{"find": "no-such-string-anywhere", "replace": "x"}]
    )
    assert set(r) >= {"replaced", "total", "live"}
    assert "items" not in r and "total_replacements" not in r
    assert r["replaced"] == {"no-such-string-anywhere": 0}
    # list-shaped reads: flat lists, same entry keys as file mode
    text = srv.get_text(path)
    assert isinstance(text, list)
    assert {"index", "text", "style"} <= set(text[0])
    found = srv.find_text(path, "Intro heading")
    assert isinstance(found, list)
    assert any("paragraph_index" in m for m in found)
    outline = srv.get_outline(path)
    assert isinstance(outline, list)


@live_mark
@needs_word
def test_l8_live_outline_sees_outlinelvl_headings(ws_l_doc):
    path = ws_l_doc["path"]
    outline = srv.get_outline(path)    # auto → live
    by_text = {h["text"]: h for h in outline}
    assert "Intro heading" in by_text
    assert by_text["Intro heading"]["detected_via"] == "heading_style"
    assert "Template heading" in by_text, (
        "live outline is blind to outlineLvl-based headings"
    )
    th = by_text["Template heading"]
    assert th["level"] == 1
    assert th["detected_via"] == "outline_level"
    # parity with the captured file outline
    file_by_text = {h["text"]: h for h in ws_l_doc["outline"]}
    assert set(by_text) == set(file_by_text)
    for k in by_text:
        assert by_text[k]["level"] == file_by_text[k]["level"]
        assert by_text[k]["paragraph_index"] == \
            file_by_text[k]["paragraph_index"]


@live_mark
@needs_word
def test_zombie_gate_own_fixture(ws_l_doc):
    """Our fixture's instance is still attached (moniker present) while the
    module runs; the teardown quits it. This is a canary that the document
    is really being served by an instance we control."""
    from word_mcp.com import live as live_mod

    status = live_mod.interactive_status()
    assert any(
        d["path"].lower() == ws_l_doc["path"].lower()
        for d in status["open_documents"]
    )
