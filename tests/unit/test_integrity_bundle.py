"""Academic-integrity bundle: integrity, cleanup, reffields ops.

Docs are built with the server's own file-based tools; reference-manager
fields (Zotero/EndNote/Mendeley) are hand-built via lxml because no server
tool creates ADDIN fields — the XML shape copies ops/fields.py:_field_run.
"""

import base64

import pytest
from docx import Document
from lxml import etree

import word_mcp.server as srv
from word_mcp.core.errors import DocumentProtected
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import cleanup, integrity, read, reffields

# 1x1 red PNG, for image-caption tests.
_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8"
    "z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def make_doc(tmp_path, paragraphs, name="doc.docx"):
    path = str(tmp_path / name)
    srv.create_document(path)
    srv.insert_paragraphs(
        path,
        [{"text": t} for t in paragraphs],
        backup=False,
        live="off",
    )
    return path


def _field_runs(instr, cached, *, broken=False):
    """Complex-field runs in the fields.py shape; broken=True omits the end
    marker (the damage list_reference_fields must catch)."""
    els = []
    r1 = etree.Element(qn("w:r"))
    etree.SubElement(r1, qn("w:fldChar")).set(qn("w:fldCharType"), "begin")
    els.append(r1)
    r2 = etree.Element(qn("w:r"))
    it = etree.SubElement(r2, qn("w:instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = f" {instr} "
    els.append(r2)
    r3 = etree.Element(qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    els.append(r3)
    r4 = etree.Element(qn("w:r"))
    t = etree.SubElement(r4, qn("w:t"))
    t.text = cached
    els.append(r4)
    if not broken:
        r5 = etree.Element(qn("w:r"))
        etree.SubElement(r5, qn("w:fldChar")).set(qn("w:fldCharType"), "end")
        els.append(r5)
    return els


def _append_field(pkg, para_contains, instr, cached, *, broken=False):
    for kind, _idx, el in read.body_items(pkg):
        if kind == "paragraph" and para_contains in read.paragraph_text(el):
            for run in _field_runs(instr, cached, broken=broken):
                el.append(run)
            pkg.mark_dirty()
            return
    raise AssertionError(f"no paragraph containing {para_contains!r}")


def _set_seq_placeholder(path, number):
    """Simulate Word updating SEQ fields: '#' placeholders -> a number."""
    pkg = DocxPackage(path)
    for t in pkg.root().iter(qn("w:t")):
        if t.text == "#":
            t.text = number
    pkg.mark_dirty()
    pkg.save(do_backup=False)


# --------------------------------------------------------- cross-references


def test_valid_cross_reference_ok(tmp_path):
    path = make_doc(
        tmp_path,
        ["The delta model shows change over time.", "See the later section."],
    )
    srv.insert_bookmark(path, "DeltaModel", anchor_text="delta model", backup=False)
    srv.insert_cross_reference(
        path, to_bookmark="DeltaModel", kind="page",
        location={"search": {"text": "See the later"}}, backup=False,
    )
    rep = integrity.validate_cross_references(DocxPackage(path))
    assert rep["ok"] is True
    assert rep["broken"] == []
    assert rep["ref_fields"] == 1
    assert not any(
        b["name"] == "DeltaModel" for b in rep["unreferenced_bookmarks"]
    )


def test_broken_cross_reference_detected(tmp_path):
    path = make_doc(
        tmp_path,
        ["The delta model shows change over time.", "See the later section."],
    )
    srv.insert_bookmark(path, "DeltaModel", anchor_text="delta model", backup=False)
    srv.insert_cross_reference(
        path, to_bookmark="DeltaModel", kind="page",
        location={"search": {"text": "See the later"}}, backup=False,
    )
    # Delete the bookmark out from under the field.
    pkg = DocxPackage(path)
    for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
        for el in list(pkg.root().iter(qn(tag))):
            el.getparent().remove(el)
    pkg.mark_dirty()
    pkg.save(do_backup=False)

    rep = integrity.validate_cross_references(DocxPackage(path))
    assert rep["ok"] is False
    assert len(rep["broken"]) == 1
    assert rep["broken"][0]["bookmark"] == "DeltaModel"
    assert rep["broken"][0]["field"] == "PAGEREF"
    assert rep["broken"][0]["paragraph_index"] == 1


def test_unreferenced_bookmark_reported(tmp_path):
    path = make_doc(tmp_path, ["The delta model shows change over time."])
    srv.insert_bookmark(path, "Orphan", anchor_text="delta model", backup=False)
    rep = integrity.validate_cross_references(DocxPackage(path))
    assert rep["ok"] is True  # informational, not an error
    assert any(b["name"] == "Orphan" for b in rep["unreferenced_bookmarks"])


def test_textual_reference_unverified_then_verified(tmp_path):
    path = make_doc(
        tmp_path,
        ["As shown in Figure 3, the trend holds.", "Results follow."],
    )
    rep = integrity.validate_cross_references(DocxPackage(path))
    unverified = rep["text_references"]["unverified"]
    assert any(u["text"] == "Figure 3" for u in unverified)
    assert unverified[0]["paragraph_index"] == 0

    # Add a Figure caption and simulate Word computing its number as 3.
    srv.insert_caption(
        path, text="Trend over time", after_anchor="Results follow",
        label="Figure", backup=False,
    )
    _set_seq_placeholder(path, "3")
    rep2 = integrity.validate_cross_references(DocxPackage(path))
    assert not any(
        u["text"] == "Figure 3" for u in rep2["text_references"]["unverified"]
    )


# ------------------------------------------------------------------ captions


def test_missing_table_caption_detected(tmp_path):
    path = make_doc(tmp_path, ["Intro paragraph for the tables."])
    srv.create_table(path, [["A", "B"], ["1", "2"]], backup=False)
    srv.create_table(path, [["C", "D"], ["3", "4"]], backup=False)
    srv.insert_caption(
        path, text="First table", table_index=0, label="Table", backup=False
    )
    rep = integrity.validate_captions(DocxPackage(path))
    assert rep["tables_checked"] == 2
    assert rep["ok"] is False
    assert rep["missing"] == [{"kind": "table", "table_index": 1}]


def test_missing_image_caption_then_resolved(tmp_path):
    png = tmp_path / "dot.png"
    png.write_bytes(base64.b64decode(_PNG_B64))
    path = make_doc(tmp_path, ["Diagram below."])
    srv.insert_image(path, str(png), backup=False)

    rep = integrity.validate_captions(DocxPackage(path))
    assert rep["images_checked"] == 1
    assert any(m["kind"] == "image" for m in rep["missing"])

    # Caption lands between the text and the image -> adjacent to the image.
    srv.insert_caption(
        path, text="A diagram", after_anchor="Diagram below", label="Figure",
        backup=False,
    )
    rep2 = integrity.validate_captions(DocxPackage(path))
    assert rep2["missing"] == []
    assert rep2["ok"] is True


def test_mixed_numbering_convention_flagged(tmp_path):
    path = make_doc(tmp_path, ["Intro paragraph for the figures."])
    srv.create_table(path, [["A"], ["1"]], backup=False)
    srv.create_table(path, [["B"], ["2"]], backup=False)
    srv.insert_caption(
        path, text="Sequential one", table_index=0, label="Figure",
        backup=False,
    )
    _set_seq_placeholder(path, "3")  # sequential: "Figure 3"

    # Hand-build a chapter-relative caption ("Figure 4.2") above table 1.
    pkg = DocxPackage(path)
    cap = etree.Element(qn("w:p"))
    ppr = etree.SubElement(cap, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "Caption")
    r = etree.SubElement(cap, qn("w:r"))
    t = etree.SubElement(r, qn("w:t"))
    t.text = "Figure "
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for el in _field_runs("SEQ Figure \\* ARABIC \\s 1", "4.2"):
        cap.append(el)
    r2 = etree.SubElement(cap, qn("w:r"))
    etree.SubElement(r2, qn("w:t")).text = ": Chapter-relative one"
    tables = [el for k, _i, el in read.body_items(pkg) if k == "table"]
    tables[1].addprevious(cap)
    pkg.mark_dirty()
    pkg.save(do_backup=False)

    rep = integrity.validate_captions(DocxPackage(path))
    assert rep["conventions"]["Figure"] == "mixed"
    assert rep["mixed_conventions"] == ["Figure"]
    assert rep["ok"] is False
    assert rep["missing"] == []


# ------------------------------------------------------------------- cleanup


def test_prepare_for_submission_end_to_end(tmp_path):
    path = make_doc(
        tmp_path,
        ["The committee reviews the draft.", "Another paragraph here."],
    )
    srv.set_document_properties(
        path, title="Keep Me", author="Secret Author", backup=False
    )
    srv.search_and_replace(
        path, [{"find": "draft", "replace": "manuscript"}],
        track=True, author="Reviewer", backup=False, live="off",
    )
    srv.manage_comment(
        path, action="add", text="fix this", author="Reviewer",
        location={"search": {"text": "Another paragraph"}}, backup=False,
    )
    assert read.revision_summary(DocxPackage(path))["total"] > 0
    assert len(read.get_comments(DocxPackage(path))) == 1

    pkg = DocxPackage(path)
    rep = cleanup.prepare_for_submission(pkg)
    pkg.save(do_backup=False)

    assert rep["revisions_accepted"] > 0
    assert rep["comments_removed"] == 1
    assert "word/comments.xml" in rep["comment_parts_removed"]
    assert any("creator" in s for s in rep["metadata_scrubbed"])
    assert rep["not_removed"]  # rsids stay, and the result says so
    assert rep["remaining"]["tracked_changes"] == 0
    assert rep["remaining"]["comments"] == 0

    pkg2 = DocxPackage(path)
    assert read.revision_summary(pkg2)["total"] == 0
    assert read.get_comments(pkg2) == []
    for part in (
        "word/comments.xml", "word/commentsExtended.xml",
        "word/commentsIds.xml", "word/people.xml",
    ):
        assert not pkg2.has_part(part)
    core = pkg2.root("docProps/core.xml")
    creator = core.find("{http://purl.org/dc/elements/1.1/}creator")
    assert creator is None or not (creator.text or "")
    title = core.find("{http://purl.org/dc/elements/1.1/}title")
    assert title is not None and title.text == "Keep Me"
    # Accepted replacement survived; comments-family rels/overrides are gone.
    text = "\n".join(p["text"] for p in read.get_paragraphs(pkg2))
    assert "manuscript" in text and "draft" not in text
    ct = pkg2.raw_part("[Content_Types].xml").decode("utf-8")
    assert "people" not in ct and "comments" not in ct
    Document(path)  # opens clean


def test_prepare_for_submission_refuses_protected(tmp_path):
    path = make_doc(tmp_path, ["A protected document."])
    srv.set_document_protection(path, protection="readOnly", backup=False)
    with pytest.raises(DocumentProtected):
        cleanup.prepare_for_submission(DocxPackage(path))


# ----------------------------------------------------------------- reffields


def _build_manager_doc(tmp_path):
    path = make_doc(
        tmp_path,
        [
            "Alpha paragraph.", "Bravo paragraph.", "Charlie paragraph.",
            "Delta paragraph.", "Echo paragraph.",
        ],
    )
    pkg = DocxPackage(path)
    _append_field(
        pkg, "Alpha",
        'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"a1"}',
        "(Hurd, 1999)",
    )
    _append_field(
        pkg, "Bravo",
        'ADDIN ZOTERO_BIBL {"uncited":[]} CSL_BIBLIOGRAPHY',
        "Hurd, I. (1999). Legitimacy and authority.",
    )
    _append_field(
        pkg, "Charlie",
        "ADDIN EN.CITE <EndNote><Cite><Author>Lake</Author></Cite></EndNote>",
        "(Lake, 2009)",
    )
    _append_field(
        pkg, "Delta",
        'ADDIN CSL_CITATION {"citationItems":[]}',
        "(Kelman, 2008)",
    )
    # Deliberately broken: begin + instr + separate, no end. Last in document
    # order so nothing after it gets swallowed by the open field.
    _append_field(
        pkg, "Echo",
        'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"b2"}',
        "(Broken, 2020)",
        broken=True,
    )
    pkg.save(do_backup=False)
    return path


def test_reference_field_inventory(tmp_path):
    path = _build_manager_doc(tmp_path)
    inv = reffields.list_reference_fields(DocxPackage(path))

    assert inv["total"] == 5
    assert inv["by_manager"] == {"zotero": 3, "endnote": 1, "mendeley": 1}
    kinds = {(f["manager"], f["kind"]) for f in inv["fields"]}
    assert ("zotero", "bibliography") in kinds
    assert ("endnote", "citation") in kinds
    assert ("mendeley", "citation") in kinds

    hurd = next(f for f in inv["fields"] if f["cached_text"] == "(Hurd, 1999)")
    assert hurd["manager"] == "zotero" and hurd["kind"] == "citation"
    assert hurd["intact"] is True
    # Location matches the paragraph the field lives in.
    alpha_idx = read.get_paragraphs(DocxPackage(path), contains="Alpha")[0]["index"]
    assert hurd["paragraph_index"] == alpha_idx

    assert len(inv["broken"]) == 1
    assert inv["broken"][0]["manager"] == "zotero"
    broken_field = next(f for f in inv["fields"] if not f["intact"])
    assert broken_field["cached_text"] == "(Broken, 2020)"


def test_reference_field_integrity_check(tmp_path):
    path = _build_manager_doc(tmp_path)
    chk = reffields.check_reference_field_integrity(DocxPackage(path))
    assert chk["ok"] is False
    assert chk["citations"] == 4
    assert chk["bibliographies"] == 1
    assert len(chk["broken"]) == 1
    assert "warning" in chk


def test_reference_field_integrity_clean_doc(tmp_path):
    # Non-manager fields (a PAGEREF cross-reference) must not trip the check.
    path = make_doc(tmp_path, ["The delta model target.", "See the target."])
    srv.insert_bookmark(path, "Target", anchor_text="delta model", backup=False)
    srv.insert_cross_reference(
        path, to_bookmark="Target", kind="page",
        location={"search": {"text": "See the target"}}, backup=False,
    )
    chk = reffields.check_reference_field_integrity(DocxPackage(path))
    assert chk["ok"] is True
    assert chk["total_reference_fields"] == 0
    assert chk["broken"] == []
