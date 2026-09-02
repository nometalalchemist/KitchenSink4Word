"""v1.6 small-batch sweep gate: include_textboxes on get_text/find_text
(box content via ops/textboxes.py, never the body walk), replace_formatted
(find_formatted's mutation twin), and the Chicago-notes narrative-harvest
upgrade in styleconvert. Synthetic documents built with the server-layer
functions; no COM, no Word."""

import hashlib
from pathlib import Path

import pytest
from lxml import etree

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import read, stylefind as sf, styleconvert as sc

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def new_doc(tmp_path, name="doc.docx"):
    path = str(tmp_path / name)
    srv.create_document(path)
    return path


def md5(path):
    return hashlib.md5(open(path, "rb").read()).hexdigest()


def _txbx_xml(box_text, name="Text Box 1"):
    """Minimal AlternateContent shape (wps:txbx) WITH the mc:Fallback VML
    copy Word always writes alongside — the doubled-storage case the body
    walk must never re-include."""
    content = f"<w:txbxContent><w:p><w:r><w:t>{box_text}</w:t></w:r></w:p></w:txbxContent>"
    fallback = (
        f'<mc:Fallback><w:pict><v:shape id="{name}" '
        f'style="width:144pt;height:72pt">'
        f"<v:textbox>{content}</v:textbox></v:shape></w:pict></mc:Fallback>"
    )
    return f'''<w:r xmlns:w="{W}" xmlns:mc="{MC}" xmlns:wps="{WPS}"
 xmlns:a="{A}" xmlns:wp="{WP}" xmlns:v="{V}">
<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
 relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1"
 allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
<wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/><wp:docPr id="7" name="{name}"/>
<a:graphic><a:graphicData uri="{WPS}">
<wps:wsp><wps:cNvSpPr txBox="1"/>
<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
<wps:txbx>{content}</wps:txbx><wps:bodyPr/></wps:wsp>
</a:graphicData></a:graphic></wp:anchor>
</w:drawing></mc:Choice>{fallback}</mc:AlternateContent></w:r>'''


def inject_run_xml(path, host_text, run_xml):
    pkg = DocxPackage(path)
    for p in pkg.body().findall(qn("w:p")):
        joined = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if host_text in joined:
            p.append(etree.fromstring(run_xml))
            pkg.mark_dirty()
            pkg.save(do_backup=False)
            return
    raise AssertionError(f"host paragraph not found: {host_text!r}")


@pytest.fixture
def boxed_doc(tmp_path):
    path = new_doc(tmp_path, "box.docx")
    srv.insert_paragraphs(
        path,
        [{"text": "Before the box."}, {"text": "Host paragraph."},
         {"text": "After the box."}], backup=False, live="off",
    )
    inject_run_xml(
        path, "Host paragraph.", _txbx_xml("Callout text lives here.")
    )
    return path


# ================================================ item 1: include_textboxes


def test_get_text_default_excludes_boxes(boxed_doc):
    paras = srv.get_text(boxed_doc, live="off")
    joined = " ".join(p["text"] for p in paras)
    assert "Callout" not in joined
    assert all("source" not in p for p in paras)


def test_get_text_include_textboxes_appends_labeled_entries(boxed_doc):
    base = srv.get_text(boxed_doc, live="off")
    withbox = srv.get_text(boxed_doc, include_textboxes=True, live="off")
    # body entries (and their indices) are byte-identical to the default read
    assert withbox[: len(base)] == base
    extra = withbox[len(base):]
    assert len(extra) == 1
    box = extra[0]
    assert box["source"] == "textbox"
    assert box["index"] is None
    assert box["box_index"] == 0
    assert box["text"] == "Callout text lives here."
    assert box["shape_name"] == "Text Box 1"
    # the host body paragraph anchors the box
    host_idx = next(
        p["index"] for p in base if p["text"] == "Host paragraph."
    )
    assert box["anchor_paragraph_index"] == host_idx


def test_get_text_include_textboxes_contains_filter(boxed_doc):
    out = srv.get_text(
        boxed_doc, contains="Callout", include_textboxes=True, live="off"
    )
    assert [e["text"] for e in out] == ["Callout text lives here."]
    assert out[0]["source"] == "textbox"


def test_find_text_default_excludes_boxes(boxed_doc):
    assert srv.find_text(boxed_doc, "Callout", live="off") == []


def test_find_text_include_textboxes_single_labeled_match(boxed_doc):
    matches = srv.find_text(
        boxed_doc, "Callout", include_textboxes=True, live="off"
    )
    # exactly ONE match despite the mc:Fallback twin — no doubled text
    assert len(matches) == 1
    m = matches[0]
    assert m["source"] == "textbox"
    assert m["box_index"] == 0
    assert m["match"] == "Callout"
    assert "Callout text lives here." in m["context"]
    # body matches keep their shape and indices
    body = srv.find_text(
        boxed_doc, "Host", include_textboxes=True, live="off"
    )
    assert [m for m in body if "paragraph_index" in m]


def test_find_text_include_textboxes_regex(boxed_doc):
    matches = srv.find_text(
        boxed_doc, r"lives\s+here", regex=True,
        include_textboxes=True, live="off",
    )
    assert len(matches) == 1
    assert matches[0]["source"] == "textbox"


# ================================================ item 2: replace_formatted


def _replace(path, **kw):
    pkg = DocxPackage(path)
    result = sf.replace_formatted(pkg, **kw)
    if result["total"]:
        pkg.save(do_backup=False)
    return result


@pytest.fixture
def fmt_doc(tmp_path):
    path = new_doc(tmp_path, "fmt.docx")
    srv.insert_paragraphs(
        path,
        [{"text": "The delta term is bold here."},
         {"text": "The delta term is plain here."},
         {"text": "Entirely bold sentence."}], backup=False, live="off",
    )
    srv.format_text(
        path, {"bold": True}, range={"start": 0, "end": 0}, find="delta term",
        backup=False, live="off",
    )
    srv.format_text(
        path, {"bold": True}, range={"start": 2, "end": 2}, backup=False, live="off",
    )
    return path


def test_replace_only_inside_matching_formatting(fmt_doc):
    result = _replace(
        fmt_doc, formatting={"bold": True}, find="delta", replace="DELTA"
    )
    assert result["total"] == 1
    assert result["replaced"] == {"delta": 1}
    texts = [p["text"] for p in srv.get_text(fmt_doc, live="off")]
    assert "The DELTA term is bold here." in texts
    assert "The delta term is plain here." in texts  # non-bold untouched
    rep = result["replacements"][0]
    assert rep["matched_via"]["bold"] == "explicit"
    assert rep["paragraph_index"] == 0
    assert rep["replaced_text"] == "delta"


def test_replace_whole_stretch_when_find_none(fmt_doc):
    result = _replace(
        fmt_doc, formatting={"bold": True}, find=None, replace="[X]"
    )
    # two bold stretches: "delta term" and the whole bold sentence
    assert result["total"] == 2
    assert result["replaced"] == {"(formatted stretch)": 2}
    texts = [p["text"] for p in srv.get_text(fmt_doc, live="off")]
    assert "The [X] is bold here." in texts
    assert "[X]" in texts
    assert "The delta term is plain here." in texts


def test_replace_keeps_matched_formatting(fmt_doc):
    _replace(fmt_doc, formatting={"bold": True}, find="delta", replace="DELTA")
    hit = sf.find_formatted(
        DocxPackage(fmt_doc), "DELTA", formatting={"bold": True}
    )
    assert hit["total"] == 1  # replacement landed inside the bold run


def test_replace_across_fragmented_runs(tmp_path):
    path = new_doc(tmp_path, "frag.docx")
    srv.insert_paragraphs(
        path, [{"text": "alpha beta gamma"}],
        backup=False, live="off",
    )
    # two formatting calls fragment the paragraph into multiple bold runs
    srv.format_text(path, {"bold": True}, range={"start": 0, "end": 0}, find="alpha be",
                    backup=False, live="off")
    srv.format_text(path, {"bold": True}, range={"start": 0, "end": 0}, find="ta gamma",
                    backup=False, live="off")
    result = _replace(
        path, formatting={"bold": True}, find="beta", replace="BETA"
    )
    assert result["total"] == 1  # match spans the run boundary
    assert srv.get_text(path, live="off")[0]["text"] == "alpha BETA gamma"


def test_replace_by_style_criterion(tmp_path):
    path = new_doc(tmp_path, "styled.docx")
    srv.insert_paragraphs(
        path,
        [{"text": "Alpha section"}, {"text": "Alpha in body text."}], backup=False, live="off",
    )
    srv.apply_style(path, style="Heading1", range={"start": 0, "end": 0}, backup=False)
    result = _replace(
        path, formatting={"style": "Heading1"}, find="Alpha", replace="Beta"
    )
    assert result["total"] == 1
    texts = [p["text"] for p in srv.get_text(path, live="off")]
    assert "Beta section" in texts
    assert "Alpha in body text." in texts
    assert result["replacements"][0]["matched_via"]["style"] == "paragraph_style"


def test_replace_scope_all_reaches_footnotes(tmp_path):
    path = new_doc(tmp_path, "notes.docx")
    srv.insert_paragraphs(
        path, [{"text": "Body sentence with anchor."}],
        backup=False, live="off",
    )
    srv.manage_note(path, action="insert", note_type="footnote", text="Note says alpha.", location={"search": {"text": "anchor"}}, backup=False)
    body_only = _replace(
        path, formatting={"style": "FootnoteText"}, find="alpha",
        replace="beta", scope="body",
    )
    assert body_only["total"] == 0
    result = _replace(
        path, formatting={"style": "FootnoteText"}, find="alpha",
        replace="beta", scope="all",
    )
    assert result["total"] == 1
    assert result["replacements"][0]["part"] == "word/footnotes.xml"
    assert result["replacements"][0]["paragraph_index"] is None
    assert "beta" in read.list_footnotes(DocxPackage(path))[0]["text"]


def test_replace_never_touches_textboxes(boxed_doc):
    result = _replace(
        boxed_doc, formatting={"bold": False}, find="Callout", replace="XX"
    )
    assert result["total"] == 0
    box = srv.get_text(boxed_doc, include_textboxes=True, live="off")[-1]
    assert box["text"] == "Callout text lives here."


def test_replace_refusals(fmt_doc):
    before = md5(fmt_doc)
    with pytest.raises(WordMcpError, match="regex"):
        sf.replace_formatted(
            DocxPackage(fmt_doc), formatting={"bold": True},
            find=r"del.a", replace="x", regex=True,
        )
    with pytest.raises(WordMcpError, match="find must be non-empty"):
        sf.replace_formatted(
            DocxPackage(fmt_doc), formatting={"bold": True},
            find="", replace="x",
        )
    with pytest.raises(WordMcpError, match="scope"):
        sf.replace_formatted(
            DocxPackage(fmt_doc), formatting={"bold": True},
            find="delta", replace="x", scope="everything",
        )
    with pytest.raises(WordMcpError, match="unknown formatting key"):
        sf.replace_formatted(
            DocxPackage(fmt_doc), formatting={"bogus": True},
            find="delta", replace="x",
        )
    with pytest.raises(WordMcpError, match="max_replacements"):
        sf.replace_formatted(
            DocxPackage(fmt_doc), formatting={"bold": True}, find=None,
            replace="x", max_replacements=1,
        )
    assert md5(fmt_doc) == before  # every refusal left the file untouched


def test_replace_result_mirrors_search_and_replace_shape(fmt_doc):
    result = _replace(
        fmt_doc, formatting={"bold": True}, find="delta", replace="DELTA"
    )
    assert set(result) >= {
        "replaced", "total", "criteria", "find", "scope",
        "unresolved_theme_runs", "replacements",
    }
    assert result["criteria"] == {"bold": True}
    assert result["scope"] == "body"


# ===================================== item 3: notes narrative harvesting


CLEAN_BODY = [
    "Legitimacy shapes compliance in institutions (Hurd, 1999, p. 384).",
    "Lake (2009) argues that hierarchy persists.",
    "Alliance behavior tracks threat (Walt, 1987).",
    "References",
]
CLEAN_ENTRIES = [
    "Hurd, I. (1999). Legitimacy and authority in international politics. International Organization, 53(2), 379-408.",
    "Lake, D. A. (2009). Hierarchy in international relations. Cornell University Press.",
    "Snyder, G. H. (1997). Alliance politics. Cornell University Press.",
    "Walt, S. M. (1987). The origins of alliances. Cornell University Press.",
]


def _nb_doc(tmp_path, extra_notes=()):
    """Manuscript converted to Chicago notes by the tool itself, plus
    hand-written footnotes in the real-world forms under test. extra_notes:
    (anchor_text, note_text) pairs, appended in order."""
    path = str(tmp_path / "nb.docx")
    srv.create_document(path)
    srv.insert_paragraphs(
        path, [{"text": t} for t in CLEAN_BODY + CLEAN_ENTRIES], backup=False, live="off",
    )
    srv.apply_style(path, style="Heading1", range={"start": 3, "end": 3}, backup=False)
    pkg = DocxPackage(path)
    r = sc.convert_citation_style(pkg, "chicago17-notes")
    assert r["converted"]
    pkg.save(do_backup=False)
    for anchor, note in extra_notes:
        srv.manage_note(path, action="insert", note_type="footnote", text=note, location={"search": {"text": anchor}}, backup=False)
    return path


def _back_to_apa(path):
    pkg = DocxPackage(path)
    result = sc.convert_citation_style(
        pkg, "apa7", source_style="chicago17-notes"
    )
    if result.get("converted"):
        pkg.save(do_backup=False)
    return result


def _body_text(path):
    return " ".join(p["text"] for p in srv.get_text(path, live="off"))


def test_harvest_new_real_note_forms(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        # signal-prefixed narrative note
        ("Legitimacy", "See Hurd (1999, 384)."),
        # unquoted short BOOK note (title verified against the entry)
        ("hierarchy", "Lake, Hierarchy in international relations, 62."),
        # author-only short form (unique family name)
        ("threat", "Snyder, 12."),
    ])
    result = _back_to_apa(path)
    # 3 tool-created citation notes + 3 hand-written real-form notes
    assert result["footnotes_harvested"] == 6
    assert read.list_footnotes(DocxPackage(path)) == []
    body = _body_text(path)
    assert "(see Hurd, 1999, p. 384)" in body  # signal word carried
    assert "(Lake, 2009, p. 62)" in body
    assert "(Snyder, 1997, p. 12)" in body


def test_harvest_ibid_resolves_to_previous_note(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        ("Legitimacy", "Hurd (1999, 384)."),
        ("hierarchy", "Ibid., 390."),
    ])
    result = _back_to_apa(path)
    assert result["footnotes_harvested"] == 5
    assert read.list_footnotes(DocxPackage(path)) == []
    body = _body_text(path)
    assert "(Hurd, 1999, p. 384)" in body
    assert "(Hurd, 1999, p. 390)" in body  # Ibid. took the new page


def test_harvest_ibid_after_unharvestable_note_left_alone(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        ("Legitimacy", "This claim is contested in the literature."),
        ("hierarchy", "Ibid."),
    ])
    result = _back_to_apa(path)
    assert result["footnotes_harvested"] == 3  # only the tool-created notes
    remaining = [f["text"] for f in read.list_footnotes(DocxPackage(path))]
    assert "This claim is contested in the literature." in remaining
    assert "Ibid." in remaining
    problems = " ".join(
        f["problem"] for f in result["citations_flagged"]
    )
    assert "Ibid" in problems


def test_harvest_dangling_ibid_cancels_antecedent(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        ("Legitimacy", "Walt (1987)."),
        ("hierarchy", "Ibid.; but compare the appendix."),
    ])
    result = _back_to_apa(path)
    # the Walt note would resolve, but harvesting it would strand the Ibid
    assert result["footnotes_harvested"] == 3
    remaining = [f["text"] for f in read.list_footnotes(DocxPackage(path))]
    assert "Walt (1987)." in remaining
    assert "Ibid.; but compare the appendix." in remaining
    problems = " ".join(f["problem"] for f in result["citations_flagged"])
    assert "antecedent" in problems


def test_harvest_signal_refused_for_numeric_targets(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        ("Legitimacy", "See Hurd (1999, 384)."),
    ])
    pkg = DocxPackage(path)
    result = sc.convert_citation_style(
        pkg, "ieee", source_style="chicago17-notes", dry_run=True
    )
    flagged = " ".join(
        f["problem"] for f in result["plan"]["citations_flagged"]
    )
    assert "signal word" in flagged


def test_harvest_prose_narrative_still_left_alone(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        ("Legitimacy", "Hurd (1999) makes this argument at length."),
    ])
    result = _back_to_apa(path)
    assert result["footnotes_harvested"] == 3
    remaining = [f["text"] for f in read.list_footnotes(DocxPackage(path))]
    assert remaining == ["Hurd (1999) makes this argument at length."]


def test_short_book_note_without_title_match_left_alone(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        ("Legitimacy", "Lake, Some Entirely Different Title, 62."),
    ])
    result = _back_to_apa(path)
    assert result["footnotes_harvested"] == 3
    assert len(read.list_footnotes(DocxPackage(path))) == 1


def test_dry_run_still_byte_identical(tmp_path):
    path = _nb_doc(tmp_path, extra_notes=[
        ("Legitimacy", "See Hurd (1999, 384)."),
    ])
    before = md5(path)
    pkg = DocxPackage(path)
    result = sc.convert_citation_style(
        pkg, "apa7", source_style="chicago17-notes", dry_run=True
    )
    assert result["dry_run"] is True and result["converted"] is False
    assert md5(path) == before


# ============================= scope add 1: outline_level (Feature Gap #2)


@pytest.fixture
def template_doc(tmp_path):
    """NSU-template pattern: Normal-styled paragraph, direct bold formatting,
    heading hierarchy meant to live in w:outlineLvl (not Heading styles)."""
    path = new_doc(tmp_path, "tmpl.docx")
    srv.insert_paragraphs(
        path,
        [{"text": "Sub-Question 1: Regulation", "formatting": {"bold": True}},
         {"text": "Ordinary body text follows."}], backup=False, live="off",
    )
    return path


def test_set_outline_level_and_read_back(template_doc):
    srv.set_paragraph_format(
        template_doc, [0], {"outline_level": 1}, backup=False
    )
    rep = sf.get_paragraph_format(DocxPackage(template_doc), 0)
    fmt = rep["paragraphs"][0]["format"]["outline_level"]
    assert fmt == {"value": 1, "source": "explicit"}
    # visual formatting and style untouched
    para = srv.get_text(template_doc, live="off")[0]
    assert para["style"] == "Normal"
    hit = sf.find_formatted(
        DocxPackage(template_doc), "Sub-Question", formatting={"bold": True}
    )
    assert hit["total"] == 1  # direct bold survived
    # the paragraph now appears in the outline (nav pane / TOC behavior);
    # WS-L's get_outline additionally reports how it was detected
    outline = srv.get_outline(template_doc, live="off")
    assert any(
        h["paragraph_index"] == 0
        and h["level"] == 2
        and h["text"] == "Sub-Question 1: Regulation"
        and h["detected_via"] == "outline_level"
        for h in outline
    )


def test_outline_level_null_removes_override(template_doc):
    srv.set_paragraph_format(
        template_doc, [0], {"outline_level": 2}, backup=False
    )
    srv.set_paragraph_format(
        template_doc, [0], {"outline_level": None}, backup=False
    )
    rep = sf.get_paragraph_format(DocxPackage(template_doc), 0)
    assert rep["paragraphs"][0]["format"]["outline_level"] == {
        "value": None, "source": "none",
    }
    pkg = DocxPackage(template_doc)
    p = pkg.body().findall(qn("w:p"))[0]
    assert p.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}") is None
    assert srv.get_outline(template_doc, live="off") == []


def test_outline_level_out_of_range_refused(template_doc):
    before = md5(template_doc)
    for bad in (9, -1, True, "2"):
        with pytest.raises(WordMcpError, match="outline_level must be"):
            srv.set_paragraph_format(
                template_doc, [0], {"outline_level": bad}, backup=False
            )
    assert md5(template_doc) == before


def test_outline_level_style_inheritance_reported(template_doc):
    srv.apply_style(template_doc, style="Heading2", range={"start": 1, "end": 1}, backup=False)
    rep = sf.get_paragraph_format(DocxPackage(template_doc), 1)
    assert rep["paragraphs"][1 - 1]["format"]["outline_level"] == {
        "value": 1, "source": "paragraph_style",
    }


# ============================== scope add 2: anchor XML-entity hint (Bug 10)


def test_anchor_entity_hint_on_not_found(tmp_path):
    from word_mcp.core.errors import TargetNotFound

    path = new_doc(tmp_path, "anchor.docx")
    srv.insert_paragraphs(
        path, [{"text": "Wit, J. S., Poneman, D. B., & Gallucci, R. L."}], backup=False, live="off",
    )
    with pytest.raises(TargetNotFound, match="PLAIN text"):
        srv.insert_paragraphs(
            path, [{"text": "new entry"}],
            location={"search": {"text": "Poneman, D. B., &amp; Gallucci"}},
            backup=False, live="off",
        )
    # a miss WITHOUT entities keeps the plain error (no misleading hint)
    with pytest.raises(TargetNotFound) as exc:
        srv.insert_paragraphs(
            path, [{"text": "new entry"}], location={"search": {"text": "No Such Anchor"}},
            backup=False, live="off",
        )
    assert "PLAIN text" not in str(exc.value)
    # the literal character works
    srv.insert_paragraphs(
        path, [{"text": "new entry"}],
        location={"search": {"text": "Poneman, D. B., & Gallucci"}},
        backup=False, live="off",
    )
    texts = [p["text"] for p in srv.get_text(path, live="off")]
    assert "new entry" in texts


# ======================================= item 4: i18n alias verification


def test_corrected_aliases_resolve():
    from word_mcp.ops.localization import canonical_for_name

    # corrections applied from sourced research (v1.6 sweep)
    assert canonical_for_name("Cabeçalho do Sumário") == "toc_heading"  # pt
    assert canonical_for_name("Testo nota piè di pagina") == "footnote_text"
    assert canonical_for_name("ブロック") == "block_text"        # ja
    assert canonical_for_name("列出段落") == "list_paragraph"    # zh UI string
    assert canonical_for_name("列表段落") == "list_paragraph"    # zh secondary
    assert canonical_for_name("Puesto") == "title"               # es
    assert canonical_for_name("표준") == "normal"                # ko
    assert canonical_for_name("正文") == "normal"                # zh


def test_wrong_and_unconfirmed_aliases_removed():
    from word_mcp.ops.localization import canonical_for_name

    # WRONG former guesses no longer mis-detect styles
    assert canonical_for_name("Título do Sumário") is None
    assert canonical_for_name("Testo nota a piè di pagina") is None
    assert canonical_for_name("ブロック テキスト") is None
    # UNCONFIRMED Korean guesses removed pending a real-install check
    assert canonical_for_name("블록 텍스트") is None
    assert canonical_for_name("목록 단락") is None


# ================================================== registration snippet
