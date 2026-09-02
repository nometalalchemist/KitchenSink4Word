"""AGENT-SAFETY + JOURNAL bundle gate: text-box extraction, dry-run replace,
word-count exclusions, peer-review anonymizer. Synthetic documents built with
the server-layer functions plus hand-built drawing XML (lxml); checks are
exercised through the ops modules. No COM, no Word."""

import hashlib
import json

import pytest
from lxml import etree

import word_mcp.server as srv
from word_mcp.core.errors import (
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import anonymize as an
from word_mcp.ops import journalcount as jc
from word_mcp.ops import preview as pv
from word_mcp.ops import textboxes as tbx

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


def texts(path):
    return [p["text"] for p in srv.get_text(path, live="off")]


def index_of(path, needle):
    for p in srv.get_text(path, live="off"):
        if needle in p["text"]:
            return p["index"]
    raise AssertionError(f"paragraph not found: {needle!r}")


def md5(path):
    return hashlib.md5(open(path, "rb").read()).hexdigest()


# ------------------------------------------------------- textbox XML builder


def _txbx_xml(box_text, *, with_fallback=True, name="Text Box 1",
              with_table=False):
    """Minimal AlternateContent shape: wp:anchor > wps:wsp > wps:txbx, with
    the mc:Fallback VML copy Word always writes alongside."""
    if with_table:
        inner = (
            f'<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'
            f"<w:tr><w:tc><w:p><w:r><w:t>{box_text}</w:t></w:r></w:p>"
            f"</w:tc></w:tr></w:tbl>"
        )
    else:
        inner = f"<w:p><w:r><w:t>{box_text}</w:t></w:r></w:p>"
    content = f"<w:txbxContent>{inner}</w:txbxContent>"
    fallback = (
        f"<mc:Fallback><w:pict>"
        f'<v:shape id="{name}" style="width:144pt;height:72pt">'
        f"<v:textbox>{content}</v:textbox></v:shape>"
        f"</w:pict></mc:Fallback>"
        if with_fallback
        else ""
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


def _vml_xml(box_text, name="Legacy Box"):
    return f'''<w:r xmlns:w="{W}" xmlns:v="{V}">
<w:pict><v:shape id="{name}" style="width:100pt;height:50pt">
<v:textbox><w:txbxContent><w:p><w:r><w:t>{box_text}</w:t></w:r></w:p>
</w:txbxContent></v:textbox></v:shape></w:pict></w:r>'''


def inject_run_xml(path, host_text, run_xml):
    """Append a hand-built run to the body paragraph containing host_text."""
    pkg = DocxPackage(path)
    for p in pkg.body().findall(qn("w:p")):
        joined = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if host_text in joined:
            p.append(etree.fromstring(run_xml))
            pkg.mark_dirty()
            pkg.save(do_backup=False)
            return
    raise AssertionError(f"host paragraph not found: {host_text!r}")


def doc_with_box(tmp_path, **kw):
    path = new_doc(tmp_path, "box.docx")
    srv.insert_paragraphs(path, [{"text": "Before the box."}],
                          backup=False)
    srv.insert_paragraphs(path, [{"text": "Host paragraph."}],
                          backup=False)
    inject_run_xml(path, "Host paragraph.", _txbx_xml("Box body text", **kw))
    return path


# --------------------------------------------------------------- text boxes


def test_textbox_read_modern_dedupes_fallback(tmp_path):
    path = doc_with_box(tmp_path)
    out = tbx.get_textbox_text(DocxPackage(path))
    assert out["count"] == 1  # Choice + Fallback = ONE box, not two
    box = out["boxes"][0]
    assert box["text"] == "Box body text"
    assert box["paragraphs"] == ["Box body text"]
    assert box["part"] == "word/document.xml"
    assert box["kind"] == "drawing"
    assert box["shape_name"] == "Text Box 1"
    assert box["anchor_paragraph_index"] == index_of(path, "Host paragraph.")


def test_textbox_read_legacy_vml(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(path, [{"text": "VML host."}],
                          backup=False)
    inject_run_xml(path, "VML host.", _vml_xml("Legacy secret"))
    out = tbx.get_textbox_text(DocxPackage(path))
    assert out["count"] == 1
    assert out["boxes"][0]["text"] == "Legacy secret"
    assert out["boxes"][0]["kind"] == "vml"
    assert out["boxes"][0]["shape_name"] == "Legacy Box"


def test_generic_read_tools_exclude_box_text(tmp_path):
    """Box text is EXCLUDED from generic reads (the runmap skips
    w:txbxContent — before this fix it smeared into the host paragraph
    DOUBLED via the Choice + Fallback copies). get_textbox_text is the one
    sanctioned reader for box content."""
    path = doc_with_box(tmp_path)
    host = texts(path)[index_of(path, "Host paragraph.")]
    assert "Box body text" not in host
    assert host == "Host paragraph."
    boxes = tbx.get_textbox_text(DocxPackage(path))
    assert any("Box body text" in b["text"] for b in boxes["boxes"])


def test_set_textbox_text_roundtrip(tmp_path):
    from docx import Document

    path = doc_with_box(tmp_path)
    pkg = DocxPackage(path)
    result = tbx.set_textbox_text(pkg, 0, "New line one\nNew line two")
    pkg.save(do_backup=False)
    assert result["paragraphs_written"] == 2
    assert result["fallback_copy_updated"] is True

    out = tbx.get_textbox_text(DocxPackage(path))
    assert out["boxes"][0]["paragraphs"] == ["New line one", "New line two"]
    # Fallback twin really rewritten: the raw XML carries the new content in
    # BOTH compatibility copies (Choice + Fallback) and the old in neither.
    # (Host paragraph text no longer includes box content — runmap fix.)
    xml = DocxPackage(path).raw_part("word/document.xml").decode("utf-8")
    assert xml.count("New line one") == 2
    assert "Box body text" not in xml
    host = texts(path)[index_of(path, "Host paragraph.")]
    assert "New line one" not in host
    # The document still opens for Word-family consumers.
    assert Document(path).paragraphs is not None


def test_set_textbox_refuses_nested_table(tmp_path):
    path = doc_with_box(tmp_path, with_table=True)
    with pytest.raises(UnsupportedStructure):
        tbx.set_textbox_text(DocxPackage(path), 0, "flatten me")


def test_set_textbox_bad_index(tmp_path):
    path = doc_with_box(tmp_path)
    with pytest.raises(TargetNotFound):
        tbx.set_textbox_text(DocxPackage(path), 5, "nope")


# ------------------------------------------------------------ preview_replace


def build_replace_doc(tmp_path):
    path = new_doc(tmp_path, "prev.docx")
    srv.insert_paragraphs(
        path,
        [
            {"text": "alpha beta alpha gamma."},
            {"text": "no match here."},
            {"text": "final alpha stands."},
        ],
        backup=False,
    )
    # Fragment the first occurrence across runs — matching must still work.
    srv.format_text(path, find="lph", occurrence=1,
                    formatting={"bold": True}, backup=False)
    return path


def test_preview_matches_real_replace_exactly(tmp_path):
    path = build_replace_doc(tmp_path)
    items = [{"find": "alpha", "replace": "OMEGA"}]
    out = pv.preview_replace(DocxPackage(path), items)
    assert out["total"] == 3
    assert out["items"][0]["matches"] == 3
    assert out["file_untouched"] is True
    para_idx = index_of(path, "alpha beta")
    first = out["matches"][0]
    assert first["paragraph_index"] == para_idx
    assert first["match"] == "alpha"
    assert "alpha beta alpha" in first["before"]
    assert first["after"].startswith("OMEGA beta")

    real = srv.search_and_replace(path, items, max_replacements=out["total"],
                                  backup=False, live="off")
    assert real["total"] == out["total"]
    assert "OMEGA beta OMEGA gamma." in texts(path)


def test_preview_never_touches_file_or_package(tmp_path):
    path = build_replace_doc(tmp_path)
    before = md5(path)
    pkg = DocxPackage(path)
    pv.preview_replace(pkg, [{"find": "alpha", "replace": "X"}], scope="body")
    assert pkg._dirty == set()
    assert md5(path) == before  # byte-identical


def test_preview_chained_items_sequencing(tmp_path):
    """Item 2 must see item 1's output, exactly like the real engine."""
    path = new_doc(tmp_path, "chain.docx")
    srv.insert_paragraphs(path, [{"text": "cat dog"}],
                          backup=False)
    items = [
        {"find": "cat", "replace": "dog"},
        {"find": "dog", "replace": "bird"},
    ]
    out = pv.preview_replace(DocxPackage(path), items)
    assert [i["matches"] for i in out["items"]] == [1, 2]
    assert out["total"] == 3
    real = srv.search_and_replace(path, items, backup=False, live="off")
    assert real["total"] == 3
    assert "bird bird" in texts(path)


def test_preview_regex_group_expansion(tmp_path):
    path = new_doc(tmp_path, "rx.docx")
    srv.insert_paragraphs(path, [{"text": "Smith 1999 wrote."}],
                          backup=False)
    out = pv.preview_replace(
        DocxPackage(path),
        [{"find": r"(Smith) (\d{4})", "replace": r"\1 (\2)", "regex": True}],
    )
    assert out["total"] == 1
    assert "Smith (1999)" in out["matches"][0]["after"]


def test_preview_guards(tmp_path):
    path = new_doc(tmp_path, "guard.docx")
    srv.insert_paragraphs(path, [{"text": "xx yy"}], backup=False)
    pkg = DocxPackage(path)
    # Invalid regex raises — the identical error the real run would raise.
    with pytest.raises(WordMcpError):
        pv.preview_replace(pkg, [{"find": "(", "replace": "x", "regex": True}])
    # Unknown scope refused (same validator as the real tool).
    with pytest.raises(WordMcpError):
        pv.preview_replace(pkg, [{"find": "a", "replace": "b"}],
                           scope="bogus")
    # Zero-width-matchable regex: reported, zero-length matches skipped.
    out = pv.preview_replace(
        pkg, [{"find": "x*", "replace": "Z", "regex": True}]
    )
    assert any("empty string" in r["problem"] for r in out["refusals"])
    assert out["total"] == 1  # the single "xx" run, nothing zero-width
    # \x07 in the replacement: reported instead of corrupting anything.
    out = pv.preview_replace(
        pkg, [{"find": "yy", "replace": "a\x07b"}]
    )
    assert any("control characters" in r["problem"] for r in out["refusals"])


# --------------------------------------------------------- journal word count


def build_journal_doc(tmp_path):
    path = new_doc(tmp_path, "journal.docx")
    srv.insert_paragraphs(path, [{"text": "Front matter text here."}], backup=False)                  # 4 words
    srv.insert_paragraphs(path, [{"text": "Abstract", "heading_level": 1}], backup=False)   # 1
    srv.insert_paragraphs(path, [{"text": "Short abstract text."}], backup=False)                  # 3
    srv.insert_paragraphs(path, [{"text": "Introduction", "heading_level": 1}], backup=False)  # 1
    srv.insert_paragraphs(
        path, [{"text": "This is the body of the paper."}], backup=False)                                    # 7
    srv.insert_paragraphs(path, [{"text": "Quoted words here."}], backup=False)                  # 3
    srv.set_paragraph_format(
        path, [index_of(path, "Quoted words here.")],
        {"indent_left_pt": 36, "indent_right_pt": 36}, backup=False)
    srv.insert_paragraphs(path, [{"text": "Figure 1. A caption."}], backup=False)                  # 4
    # Style the caption paragraph by raw pStyle (Caption need not be defined).
    pkg = DocxPackage(path)
    from word_mcp.ops.read import body_items, paragraph_text
    for kind, idx, el in body_items(pkg):
        if kind == "paragraph" and "Figure 1." in paragraph_text(el):
            ppr = el.find(qn("w:pPr"))
            if ppr is None:
                ppr = etree.SubElement(el, qn("w:pPr"))
                el.insert(0, ppr)
            etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "Caption")
    pkg.mark_dirty()
    pkg.save(do_backup=False)
    srv.create_table(path, [["h1", "h2"], ["a b", "c"]],
                     backup=False)                                    # 5
    srv.manage_note(path, action="insert", note_type="footnote", text="A footnote note.", location={"search": {"text": "body"}}, backup=False)  # 3
    srv.insert_paragraphs(path, [{"text": "References", "heading_level": 1}], backup=False)  # 1
    srv.insert_paragraphs(
        path,
        [{"text": "Hurd, I. (1999). Legitimacy and authority."}], backup=False)                                    # 6
    return path


ALL_ZONES = ("references", "captions", "footnotes", "block_quotes",
             "front_matter", "tables", "headings", "abstract")


def test_journalcount_zone_arithmetic(tmp_path):
    path = build_journal_doc(tmp_path)
    out = jc.word_count_with_exclusions(DocxPackage(path), exclude=ALL_ZONES)
    assert out["excluded"] == {
        "front_matter": 4,
        "abstract": 4,        # heading + 3 (abstract precedence over headings)
        "headings": 1,        # "Introduction"
        "block_quotes": 3,
        "captions": 4,
        "tables": 5,
        "footnotes": 3,
        "references": 7,      # "References" heading + 6-token entry
    }
    assert out["included"] == 7
    assert out["total"] == out["included"] + out["excluded_total"] == 38
    assert out["zones_detected"]["references_section"]["heading"] == "References"
    assert out["zones_detected"]["abstract_section"]["heading"] == "Abstract"


def test_journalcount_default_exclusions(tmp_path):
    path = build_journal_doc(tmp_path)
    out = jc.word_count_with_exclusions(DocxPackage(path))
    assert out["exclusions_applied"] == ["references", "captions", "footnotes"]
    assert out["excluded"] == {"references": 7, "captions": 4, "footnotes": 3}
    assert out["total"] == out["included"] + out["excluded_total"]
    assert out["included"] == 38 - 14


def test_journalcount_unknown_zone_rejected(tmp_path):
    path = new_doc(tmp_path)
    with pytest.raises(WordMcpError) as exc:
        jc.word_count_with_exclusions(DocxPackage(path), exclude=("bogus",))
    assert "references" in str(exc.value)  # allowed list is shown


# -------------------------------------------------------------- anonymizer


def build_manuscript(tmp_path):
    path = new_doc(tmp_path, "ms.docx")
    srv.insert_paragraphs(
        path,
        [
            {"text": "Hurd (1999) shows that legitimacy matters."},
            {"text": "Some argue the opposite (Hurd, 1999, p. 4)."},
            {"text": "Joint work extends this (Hurd & Lake, 2005)."},
            {"text": "Hurd and Lake (2005) argue the point."},
            {"text": "In my previous work I claimed this."},
            {"text": "As Hurd argues, this is central."},
        ],
        backup=False,
    )
    srv.insert_paragraphs(path, [{"text": "References", "heading_level": 1}], backup=False)
    srv.insert_paragraphs(
        path,
        [
            {"text": "Hurd, I. (1999). Legitimacy and authority in "
                     "international politics."},
            {"text": "Lake, D. A. (2009). Hierarchy in international "
                     "relations."},
            {"text": "Hurd, I., & Lake, D. A. (2005). Jointly written "
                     "piece."},
        ],
        backup=False,
    )
    # Identifying metadata.
    pkg = DocxPackage(path)
    core = pkg.root("docProps/core.xml")
    creator = core.find("{http://purl.org/dc/elements/1.1/}creator")
    assert creator is not None
    creator.text = "Ian Hurd"
    pkg.mark_dirty("docProps/core.xml")
    pkg.save(do_backup=False)
    return path


def run_anonymize(path, tmp_path):
    mapping = str(tmp_path / "ms.anonymization.json")
    pkg = DocxPackage(path)
    result = an.anonymize_for_review(pkg, ["Ian Hurd"], mapping_path=mapping)
    pkg.save(do_backup=False)
    return result, mapping


def test_anonymize_masks_citations_and_references(tmp_path):
    path = build_manuscript(tmp_path)
    result, mapping = run_anonymize(path, tmp_path)
    body = texts(path)
    assert "Author (1999) shows that legitimacy matters." in body
    assert "Some argue the opposite (Author, 1999, p. 4)." in body
    assert "Joint work extends this (Author, 2005)." in body
    assert "Author (2005) argue the point." in body
    # Reference entries: the author's masked, the third party untouched.
    assert "Author (1999). [Details removed for peer review.]" in body
    assert "Author (2005). [Details removed for peer review.]" in body
    assert ("Lake, D. A. (2009). Hierarchy in international relations."
            in body)
    assert result["changed"]["self_citations"] == 4
    assert result["changed"]["reference_entries"] == 2
    # Metadata scrubbed and recorded.
    assert "creator" in result["changed"]["metadata_fields"]
    core = DocxPackage(path).root("docProps/core.xml")
    assert core.find("{http://purl.org/dc/elements/1.1/}creator").text in (
        "", None)
    # Prose flagged, never edited.
    kinds = {f["kind"] for f in result["flagged_not_changed"]}
    assert "self_identifying_phrase" in kinds
    assert "surname_outside_citation" in kinds
    assert "In my previous work I claimed this." in body  # untouched
    assert "As Hurd argues, this is central." in body     # untouched
    # Mapping written, and the honesty warning is on the result.
    data = json.loads(open(mapping, encoding="utf-8").read())
    assert data["format"] == an.MAPPING_FORMAT
    assert len(data["changes"]) == 6
    assert data["metadata"]["creator"] == "Ian Hurd"
    assert any("PRIVATE" in w for w in result["warnings"])


def test_anonymize_refuses_to_overwrite_mapping(tmp_path):
    path = build_manuscript(tmp_path)
    _, mapping = run_anonymize(path, tmp_path)
    with pytest.raises(WordMcpError):
        an.anonymize_for_review(DocxPackage(path), ["Hurd"],
                                mapping_path=mapping)


def test_deanonymize_roundtrip(tmp_path):
    path = build_manuscript(tmp_path)
    original = texts(path)
    _, mapping = run_anonymize(path, tmp_path)
    assert texts(path) != original

    pkg = DocxPackage(path)
    out = an.deanonymize(pkg, mapping_path=mapping)
    pkg.save(do_backup=False)
    assert out["restored_text_changes"] == 6
    assert out["restored_metadata_fields"] >= 1  # creator, at minimum
    assert texts(path) == original  # byte-visible text fully restored
    core = DocxPackage(path).root("docProps/core.xml")
    assert core.find(
        "{http://purl.org/dc/elements/1.1/}creator").text == "Ian Hurd"


def test_deanonymize_refuses_on_drift(tmp_path):
    path = build_manuscript(tmp_path)
    _, mapping = run_anonymize(path, tmp_path)
    # Drift: an edit BEFORE a masked span shifts its recorded position.
    srv.search_and_replace(
        path, [{"find": "Some argue the opposite", "replace": "X"}],
        backup=False, live="off")
    with pytest.raises(WordMcpError) as exc:
        pkg = DocxPackage(path)
        an.deanonymize(pkg, mapping_path=mapping)
    assert "drifted" in str(exc.value)
    # Atomicity: the untouched masked paragraphs were NOT restored either.
    assert "Author (1999) shows that legitimacy matters." in texts(path)
