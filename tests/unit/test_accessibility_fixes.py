"""fix_accessibility gate: the audit's mutation twin. Per-category opt-ins,
conservative refusals, dry_run, and the audit -> fix -> audit round trip.
Documents are synthetic (server-layer builders); the tool itself is exercised
through the integration registration snippet, which doubles as its
paste-readiness smoke test. No COM, no Word."""

import importlib.util
import struct
import zlib
from pathlib import Path

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage
from word_mcp.ops import accessibility as ax

# Load the registration snippet exactly as server.py would host it: this both
# smoke-tests paste-readiness and gives the tests the real tool (dry_run
# routing + _edit save path included).
_REG = Path(__file__).resolve().parents[2] / "integration" / (
    "accessibility_fix_registrations.py"
)
if _REG.exists():
    _spec = importlib.util.spec_from_file_location("accessibility_fix_reg", _REG)
    _regmod = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(_regmod)
    fix_accessibility = _regmod.fix_accessibility
else:
    fix_accessibility = srv.fix_accessibility.fn if hasattr(srv.fix_accessibility, 'fn') else srv.fix_accessibility


def make_png(path, w=80, h=40, rgb=(200, 30, 30)):
    """Minimal valid PNG without external deps (same as the compliance gate)."""
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(
            ">I", zlib.crc32(c) & 0xFFFFFFFF
        )

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    path.write_bytes(png)
    return path


def new_doc(tmp_path, name="doc.docx", title=None):
    path = str(tmp_path / name)
    srv.create_document(path, title=title)
    return path


def audit(path):
    return ax.audit_accessibility(DocxPackage(path))


def outline_levels(path):
    return [h["level"] for h in srv.get_outline(path, live="off")]


def core_title(path):
    el = DocxPackage(path).root("docProps/core.xml").find(
        "{http://purl.org/dc/elements/1.1/}title"
    )
    return None if el is None else el.text


def build_flagged_doc(tmp_path):
    """One document carrying all four fixable finding types."""
    path = new_doc(tmp_path, "flagged.docx")  # no title
    srv.add_heading(path, "Report Title", 1, at_end=True, backup=False)
    srv.add_heading(path, "Deep Detail", 3, at_end=True, backup=False)  # skip
    srv.insert_paragraphs(
        path, [{"text": "Body text under the deep heading."}], at_end=True,
        backup=False,
    )
    png = make_png(tmp_path / "img.png")
    srv.add_image(path, str(png), at_end=True, backup=False)  # no alt text
    srv.create_table(
        path, [["Name", "Value"], ["a", "1"]], at_end=True, header_row=False,
        backup=False,
    )
    return path


# ------------------------------------------------------------- opt-in gating


def test_no_categories_selected_refuses(tmp_path):
    path = new_doc(tmp_path)
    with pytest.raises(WordMcpError, match="no fix categories selected"):
        ax.fix_accessibility(DocxPackage(path))


def test_only_opted_in_categories_run(tmp_path):
    path = build_flagged_doc(tmp_path)
    res = fix_accessibility(path, doc_title=True, backup=False)
    assert set(res["categories"]) == {"doc_title"}
    after = audit(path)["summary"]["counts"]
    # only the opted-in category was resolved; everything else untouched
    assert after["document_title_missing"] == 0
    assert after["images_missing_alt_text"] == 1
    assert after["tables_without_header_row"] == 1
    assert any(
        x["issue"] == "skipped_level"
        for x in audit(path)["findings"]["heading_hierarchy"]
    )


# ------------------------------------------------------- alt text placeholders


def test_alt_text_placeholder_marked_and_listed(tmp_path):
    path = new_doc(tmp_path, title="T")
    png = make_png(tmp_path / "img.png")
    srv.add_image(path, str(png), at_end=True, backup=False)
    res = fix_accessibility(path, alt_text_placeholders=True, backup=False)
    cat = res["categories"]["alt_text_placeholders"]
    assert len(cat["fixed"]) == 1
    assert cat["fixed"][0]["image_index"] == 0
    assert cat["fixed"][0]["placeholder"].startswith("IMAGE: needs description")
    assert cat["fixed"][0]["placeholder"].endswith(".png")  # filename hint
    # every touched image is flagged for the human replacement pass
    assert len(cat["needs_human_review"]) == 1
    assert "replace" in cat["needs_human_review"][0]["reason"].lower()
    assert audit(path)["summary"]["counts"]["images_missing_alt_text"] == 0


def test_alt_text_existing_never_touched(tmp_path):
    path = new_doc(tmp_path, title="T")
    png = make_png(tmp_path / "img.png")
    srv.add_image(path, str(png), at_end=True, backup=False)
    srv.set_image_alt_text(path, 0, "A real description", backup=False)
    res = fix_accessibility(path, alt_text_placeholders=True, backup=False)
    cat = res["categories"]["alt_text_placeholders"]
    assert cat["fixed"] == [] and cat["needs_human_review"] == []
    # the real description survives untouched
    docpr_descrs = [
        docpr.get("descr")
        for docpr in DocxPackage(path).root().iter(
            "{http://schemas.openxmlformats.org/drawingml/2006/"
            "wordprocessingDrawing}docPr"
        )
    ]
    assert docpr_descrs == ["A real description"]
    assert audit(path)["summary"]["counts"]["images_missing_alt_text"] == 0


# ------------------------------------------------------------- heading skips


def test_heading_promote_closes_gap_with_subtree(tmp_path):
    path = new_doc(tmp_path, title="T")
    srv.add_heading(path, "Top", 1, at_end=True, backup=False)
    srv.add_heading(path, "Deep", 3, at_end=True, backup=False)
    srv.add_heading(path, "Deeper", 4, at_end=True, backup=False)
    res = fix_accessibility(path, heading_skips=True, backup=False)
    cat = res["categories"]["heading_skips"]
    assert [(f["from_level"], f["to_level"]) for f in cat["fixed"]] == [
        (3, 2), (4, 3),
    ]
    assert outline_levels(path) == [1, 2, 3]
    assert not any(
        x["issue"] == "skipped_level"
        for x in audit(path)["findings"]["heading_hierarchy"]
    )


def test_heading_demote_following(tmp_path):
    path = new_doc(tmp_path, title="T")
    srv.add_heading(path, "A", 1, at_end=True, backup=False)
    srv.add_heading(path, "B", 2, at_end=True, backup=False)
    srv.add_heading(path, "C", 1, at_end=True, backup=False)  # over-promoted
    srv.add_heading(path, "D", 3, at_end=True, backup=False)  # skip after C
    res = fix_accessibility(
        path, heading_skips=True, heading_strategy="demote_following",
        backup=False,
    )
    cat = res["categories"]["heading_skips"]
    assert [(f["text"], f["from_level"], f["to_level"]) for f in cat["fixed"]] \
        == [("C", 1, 2)]
    assert outline_levels(path) == [1, 2, 2, 3]
    assert not any(
        x["issue"] == "skipped_level"
        for x in audit(path)["findings"]["heading_hierarchy"]
    )


def test_heading_nested_skip_refused_unchanged(tmp_path):
    path = new_doc(tmp_path, title="T")
    srv.add_heading(path, "One", 1, at_end=True, backup=False)
    srv.add_heading(path, "Three", 3, at_end=True, backup=False)
    srv.add_heading(path, "Five", 5, at_end=True, backup=False)
    res = fix_accessibility(path, heading_skips=True, backup=False)
    cat = res["categories"]["heading_skips"]
    assert cat["fixed"] == []
    (review,) = cat["needs_human_review"]
    assert "nothing was changed" in review["reason"]
    assert review["details"]  # says what was ambiguous
    assert outline_levels(path) == [1, 3, 5]  # untouched


def test_heading_demote_refuses_removing_only_h1(tmp_path):
    path = new_doc(tmp_path, title="T")
    srv.add_heading(path, "Only", 1, at_end=True, backup=False)
    srv.add_heading(path, "Deep", 3, at_end=True, backup=False)
    res = fix_accessibility(
        path, heading_skips=True, heading_strategy="demote_following",
        backup=False,
    )
    cat = res["categories"]["heading_skips"]
    assert cat["fixed"] == []
    (review,) = cat["needs_human_review"]
    assert any("Heading 1" in d for d in review["details"])
    assert outline_levels(path) == [1, 3]  # untouched


def test_heading_bad_strategy_rejected(tmp_path):
    path = new_doc(tmp_path, title="T")
    with pytest.raises(WordMcpError, match="heading_strategy"):
        ax.fix_accessibility(
            DocxPackage(path), heading_skips=True, heading_strategy="merge"
        )


# ------------------------------------------------------------- table headers


def test_table_headers_fix_and_numeric_skip(tmp_path):
    path = new_doc(tmp_path, title="T")
    srv.create_table(
        path, [["Name", "Value"], ["a", "1"]], at_end=True, header_row=False,
        backup=False,
    )
    srv.create_table(
        path, [["1", "2.5"], ["3", "4"]], at_end=True, header_row=False,
        backup=False,
    )
    res = fix_accessibility(path, table_headers=True, backup=False)
    cat = res["categories"]["table_headers"]
    assert [f["table_index"] for f in cat["fixed"]] == [0]
    (skip,) = cat["skipped"]
    assert skip["table_index"] == 1
    assert "all-numeric" in skip["reason"]
    flagged = audit(path)["findings"]["tables_without_header_row"]
    assert [f["location"]["table_index"] for f in flagged] == [1]


def test_table_single_row_skipped(tmp_path):
    path = new_doc(tmp_path, title="T")
    srv.create_table(
        path, [["only", "row"]], at_end=True, header_row=False, backup=False
    )
    res = fix_accessibility(path, table_headers=True, backup=False)
    cat = res["categories"]["table_headers"]
    assert cat["fixed"] == []
    assert "single-row" in cat["skipped"][0]["reason"]
    assert audit(path)["summary"]["counts"]["tables_without_header_row"] == 1


# ------------------------------------------------------------------ doc title


def test_doc_title_from_first_h1(tmp_path):
    path = new_doc(tmp_path)  # no title
    srv.add_heading(path, "Annual Report", 1, at_end=True, backup=False)
    srv.add_heading(path, "Later Heading", 1, at_end=True, backup=False)
    res = fix_accessibility(path, doc_title=True, backup=False)
    (fixedfix,) = res["categories"]["doc_title"]["fixed"]
    assert fixedfix["title"] == "Annual Report"
    assert core_title(path) == "Annual Report"
    assert audit(path)["summary"]["counts"]["document_title_missing"] == 0


def test_doc_title_never_overwritten(tmp_path):
    path = new_doc(tmp_path, title="Existing Title")
    srv.add_heading(path, "Different Heading", 1, at_end=True, backup=False)
    res = fix_accessibility(path, doc_title=True, backup=False)
    cat = res["categories"]["doc_title"]
    assert cat["fixed"] == []
    assert "never overwritten" in cat["skipped"][0]["reason"]
    assert core_title(path) == "Existing Title"


def test_doc_title_no_h1_needs_review(tmp_path):
    path = new_doc(tmp_path)  # no title
    srv.add_heading(path, "Only a Subsection", 2, at_end=True, backup=False)
    res = fix_accessibility(path, doc_title=True, backup=False)
    cat = res["categories"]["doc_title"]
    assert cat["fixed"] == []
    assert "Heading 1" in cat["needs_human_review"][0]["reason"]
    assert audit(path)["summary"]["counts"]["document_title_missing"] == 1


# --------------------------------------------------------------------- dry run


def test_dry_run_reports_but_changes_nothing(tmp_path):
    path = build_flagged_doc(tmp_path)
    before = Path(path).read_bytes()
    res = fix_accessibility(
        path, alt_text_placeholders=True, heading_skips=True,
        table_headers=True, doc_title=True, dry_run=True,
    )
    assert res["dry_run"] is True
    assert res["summary"]["applied"] is False
    assert "saved" not in res  # never went through the save path
    assert res["summary"]["fixed"] >= 4  # image, heading, table, title
    assert Path(path).read_bytes() == before  # byte-identical
    counts = audit(path)["summary"]["counts"]
    assert counts["images_missing_alt_text"] == 1
    assert counts["tables_without_header_row"] == 1
    assert counts["document_title_missing"] == 1


# ------------------------------------------------- audit -> fix -> audit gate


def test_audit_fix_audit_roundtrip(tmp_path):
    path = build_flagged_doc(tmp_path)
    before = audit(path)
    counts = before["summary"]["counts"]
    assert counts["images_missing_alt_text"] == 1
    assert counts["tables_without_header_row"] == 1
    assert counts["document_title_missing"] == 1
    assert any(
        x["issue"] == "skipped_level"
        for x in before["findings"]["heading_hierarchy"]
    )

    res = fix_accessibility(
        path, alt_text_placeholders=True, heading_skips=True,
        table_headers=True, doc_title=True,
    )
    assert res["summary"]["applied"] is True
    assert res["summary"]["fixed"] >= 4
    assert "saved" in res  # went through the real _edit save path

    after = audit(path)
    counts = after["summary"]["counts"]
    assert counts["images_missing_alt_text"] == 0
    assert counts["tables_without_header_row"] == 0
    assert counts["document_title_missing"] == 0
    assert not any(
        x["issue"] == "skipped_level"
        for x in after["findings"]["heading_hierarchy"]
    )
    # what remains needs a human: the placeholder alt text pass
    assert res["categories"]["alt_text_placeholders"]["needs_human_review"]
