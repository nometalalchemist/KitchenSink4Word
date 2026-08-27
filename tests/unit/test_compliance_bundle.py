"""COMPLIANCE bundle gate: template/brand compliance, accessibility audit,
image DPI. Synthetic documents are built with the server-layer functions
(decorated tools are plain functions); checks are exercised through the ops
modules. No COM, no Word."""

import struct
import zlib

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage
from word_mcp.ops import accessibility as ax, compliance as cp


def make_png(path, w=80, h=40, rgb=(200, 30, 30)):
    """Minimal valid PNG without external deps (same as the media gate)."""
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


def body_index_of(path, text):
    for p in srv.get_text(path, live="off"):
        if p["text"].strip() == text:
            return p["index"]
    raise AssertionError(f"paragraph not found: {text!r}")


def rules_hit(result, rule_prefix):
    return [v for v in result["violations"] if v["rule"].startswith(rule_prefix)]


# ------------------------------------------------------- template compliance


def build_passing_doc(tmp_path):
    path = new_doc(tmp_path, "pass.docx", title="Compliance Test")
    # python-docx's default template has 1.25in (90pt) left/right margins;
    # normalize to the ruleset's 1in all around.
    srv.set_section_properties(
        path, section=0,
        margins_pt={"top": 72, "bottom": 72, "left": 72, "right": 72},
        backup=False,
    )
    srv.add_heading(path, "Abstract", 1, at_end=True, backup=False)
    srv.insert_paragraphs(
        path,
        [{"text": "This dissertation examines compliance."}],
        at_end=True, backup=False,
    )
    srv.add_heading(path, "Acknowledgments", 1, at_end=True, backup=False)
    srv.insert_paragraphs(
        path, [{"text": "Thanks to the committee."}], at_end=True, backup=False
    )
    for t in (
        "Abstract",
        "This dissertation examines compliance.",
        "Acknowledgments",
        "Thanks to the committee.",
    ):
        srv.format_text(
            path, {"font": "Times New Roman", "size_pt": 12}, find=t,
            backup=False,
        )
    body_idx = [
        p["index"]
        for p in srv.get_text(path, live="off")
        if p["text"].strip() and "heading_level" not in p
    ]
    srv.set_paragraph_format(path, body_idx, {"line_spacing": 2.0}, backup=False)
    srv.set_page_number_format(
        path, section=0, number_format="decimal", start_at=1, backup=False
    )
    return path


PASS_RULES = {
    "page": {
        "margins_pt": {"top": 72, "bottom": 72, "left": 72, "right": 72},
        "tolerance_pt": 1,
        "size": "letter",
        "orientation": "portrait",
    },
    "fonts": {"allowed": ["Times New Roman"], "body_size_pt": 12},
    "line_spacing": {"body": 2.0},
    "headings": {"max_skip": 0, "required_first_level": 1},
    "page_numbering": [{"section": 0, "format": "decimal", "restart_at": 1}],
    "required_headings_in_order": ["Abstract", "Acknowledgments"],
}


def test_passing_document(tmp_path):
    path = build_passing_doc(tmp_path)
    result = cp.check_template_compliance(DocxPackage(path), PASS_RULES)
    assert result["compliant"], result["violations"]
    assert result["violation_count"] == 0
    assert set(result["rules_checked"]) == set(PASS_RULES)


def test_unknown_rule_key_rejected(tmp_path):
    path = new_doc(tmp_path)
    with pytest.raises(WordMcpError, match="allowed"):
        cp.check_template_compliance(DocxPackage(path), {"pages": {}})
    # colors is a brand rule, not a template rule
    with pytest.raises(WordMcpError, match="allowed"):
        cp.check_template_compliance(
            DocxPackage(path), {"colors": {"allowed_hex": ["FF0000"]}}
        )
    # unknown sub-key inside a known rule
    with pytest.raises(WordMcpError, match="allowed"):
        cp.check_template_compliance(
            DocxPackage(path), {"page": {"margin": {"top": 72}}}
        )


def test_margin_and_orientation_violations(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(path, [{"text": "Body."}], at_end=True, backup=False)
    srv.set_section_properties(
        path, section=0, orientation="landscape", margins_pt={"left": 90},
        backup=False,
    )
    result = cp.check_template_compliance(
        DocxPackage(path),
        {"page": {"margins_pt": {"left": 72}, "orientation": "portrait",
                  "size": "letter"}},
    )
    assert not result["compliant"]
    (mv,) = rules_hit(result, "page.margins.left")
    assert mv["expected"] == 72 and mv["found"] == 90
    assert mv["location"] == {"section": 0}
    (ov,) = rules_hit(result, "page.orientation")
    assert ov["found"] == "landscape"
    # size is orientation-agnostic: landscape letter is still letter
    assert not rules_hit(result, "page.size")


def test_page_size_violation(tmp_path):
    path = new_doc(tmp_path)  # python-docx default is US letter
    result = cp.check_template_compliance(
        DocxPackage(path), {"page": {"size": "a4"}}
    )
    (v,) = rules_hit(result, "page.size")
    assert v["expected"] == "a4"
    assert v["found"]["width_pt"] == 612


def test_font_and_size_violations(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Wrong font here."}], at_end=True, backup=False
    )
    srv.format_text(
        path, {"font": "Arial", "size_pt": 10}, find="Wrong font here.",
        backup=False,
    )
    result = cp.check_template_compliance(
        DocxPackage(path),
        {"fonts": {"allowed": ["Times New Roman"], "body_size_pt": 12}},
    )
    (fv,) = rules_hit(result, "fonts.allowed")
    assert fv["found"] == "Arial"
    (sv,) = rules_hit(result, "fonts.body_size_pt")
    assert sv["found"] == 10.0


def test_theme_font_reported_not_guessed(tmp_path):
    """A run with no explicit font resolves (via Normal/docDefaults) to a
    theme reference; that must land in 'unverified', not in violations."""
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Theme font text."}], at_end=True, backup=False
    )
    result = cp.check_template_compliance(
        DocxPackage(path), {"fonts": {"allowed": ["Times New Roman"]}}
    )
    assert not rules_hit(result, "fonts.allowed")
    assert result["compliant"]
    assert any(
        u["rule"] == "fonts.allowed" and "theme" in u["reason"]
        for u in result["unverified"]
    ), result["unverified"]


def test_line_spacing_violation(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Single spaced paragraph."}], at_end=True, backup=False
    )
    idx = body_index_of(path, "Single spaced paragraph.")
    srv.set_paragraph_format(path, [idx], {"line_spacing": 1.0}, backup=False)
    result = cp.check_template_compliance(
        DocxPackage(path), {"line_spacing": {"body": 2.0}}
    )
    (v,) = rules_hit(result, "line_spacing.body")
    assert v["expected"] == 2.0 and v["found"] == 1.0
    assert v["location"] == {"paragraph_index": idx}

    srv.set_paragraph_format(path, [idx], {"line_spacing": 2.0}, backup=False)
    result = cp.check_template_compliance(
        DocxPackage(path), {"line_spacing": {"body": 2.0}}
    )
    assert result["compliant"], result["violations"]


def test_heading_skip_and_first_level(tmp_path):
    path = new_doc(tmp_path)
    srv.add_heading(path, "Intro", 2, at_end=True, backup=False)
    srv.add_heading(path, "Very Deep", 4, at_end=True, backup=False)
    result = cp.check_template_compliance(
        DocxPackage(path),
        {"headings": {"max_skip": 0, "required_first_level": 1}},
    )
    (fl,) = rules_hit(result, "headings.required_first_level")
    assert fl["expected"] == 1 and fl["found"] == 2
    (sk,) = rules_hit(result, "headings.max_skip")
    assert sk["found"] == 4
    # max_skip 1 tolerates the 2 -> 4 jump
    result2 = cp.check_template_compliance(
        DocxPackage(path), {"headings": {"max_skip": 1}}
    )
    assert not rules_hit(result2, "headings.max_skip")


def test_required_headings_missing_and_out_of_order(tmp_path):
    path = new_doc(tmp_path)
    srv.add_heading(path, "Acknowledgments", 1, at_end=True, backup=False)
    srv.add_heading(path, "Abstract", 1, at_end=True, backup=False)
    result = cp.check_template_compliance(
        DocxPackage(path),
        {"required_headings_in_order": ["Abstract", "Acknowledgments"]},
    )
    (v,) = rules_hit(result, "required_headings_in_order")
    assert "out of order" in v["found"]
    assert v["severity"] == "warning"

    result2 = cp.check_template_compliance(
        DocxPackage(path),
        {"required_headings_in_order": ["Abstract", "Methodology"]},
    )
    (m,) = rules_hit(result2, "required_headings_in_order")
    assert m["expected"] == "Methodology"
    assert m["found"] == "heading not found"


def test_page_numbering_rules(tmp_path):
    path = new_doc(tmp_path)
    srv.set_page_number_format(
        path, section=0, number_format="lowerRoman", backup=False
    )
    result = cp.check_template_compliance(
        DocxPackage(path),
        {"page_numbering": [{"section": 0, "format": "decimal",
                             "restart_at": 1}]},
    )
    (fv,) = rules_hit(result, "page_numbering.format")
    assert fv["found"] == "lowerRoman"
    (rv,) = rules_hit(result, "page_numbering.restart_at")
    assert rv["found"] == "no restart (numbering continues)"

    ok = cp.check_template_compliance(
        DocxPackage(path),
        {"page_numbering": [{"section": 0, "format": "lowerRoman"}]},
    )
    assert ok["compliant"], ok["violations"]

    missing = cp.check_template_compliance(
        DocxPackage(path), {"page_numbering": [{"section": 3}]}
    )
    (sv,) = rules_hit(missing, "page_numbering.section")
    assert "3" in sv["expected"]

    with pytest.raises(WordMcpError, match="format"):
        cp.check_template_compliance(
            DocxPackage(path),
            {"page_numbering": [{"section": 0, "format": "roman"}]},
        )


# ----------------------------------------------------------- brand compliance


def test_brand_colors(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Red alert text."}], at_end=True, backup=False
    )
    srv.format_text(path, {"color": "FF0000"}, find="Red alert text.",
                    backup=False)
    bad = cp.check_brand_compliance(
        DocxPackage(path), {"colors": {"allowed_hex": ["1F4E79"]}}
    )
    (v,) = rules_hit(bad, "colors.allowed_hex")
    assert v["found"] == "FF0000"

    # hex normalization: '#' prefix and lowercase both accepted
    ok = cp.check_brand_compliance(
        DocxPackage(path), {"colors": {"allowed_hex": ["#ff0000"]}}
    )
    assert ok["compliant"], ok["violations"]

    # brand shares the whole template engine
    shared = cp.check_brand_compliance(
        DocxPackage(path), {"page": {"size": "letter"}}
    )
    assert shared["compliant"]


# -------------------------------------------------------------- accessibility


def build_flagged_doc(tmp_path):
    path = new_doc(tmp_path, "flagged.docx")  # no title
    srv.add_heading(path, "Introduction", 1, at_end=True, backup=False)
    srv.add_heading(path, "Deep Detail", 3, at_end=True, backup=False)  # skip
    srv.add_heading(path, "", 2, at_end=True, backup=False)  # empty heading
    srv.insert_paragraphs(
        path, [{"text": "For more, click here today."}], at_end=True,
        backup=False,
    )
    srv.add_hyperlink(path, "click here", "https://example.com", backup=False)
    srv.insert_paragraphs(
        path, [{"text": "low contrast sample"}], at_end=True, backup=False
    )
    srv.format_text(
        path, {"color": "FFFF00", "highlight": "yellow"},
        find="low contrast sample", backup=False,
    )
    png = make_png(tmp_path / "img.png")
    srv.add_image(path, str(png), at_end=True, backup=False)
    srv.create_table(
        path, [["A", "B"], ["1", "2"]], at_end=True, header_row=False,
        backup=False,
    )
    return path


def test_audit_flags_every_category(tmp_path):
    path = build_flagged_doc(tmp_path)
    audit = ax.audit_accessibility(DocxPackage(path))
    counts = audit["summary"]["counts"]
    f = audit["findings"]

    issues = {x["issue"] for x in f["heading_hierarchy"]}
    assert {"skipped_level", "empty_heading"} <= issues
    assert counts["images_missing_alt_text"] == 1
    assert f["images_missing_alt_text"][0]["location"] == {"image_index": 0}
    assert counts["tables_without_header_row"] == 1
    assert counts["document_title_missing"] == 1
    assert counts["link_text_generic"] == 1
    assert f["link_text_generic"][0]["text"] == "click here"
    assert counts["low_contrast_text"] == 1
    lc = f["low_contrast_text"][0]
    assert lc["color"] == "FFFF00" and lc["background"] == "FFFF00"
    assert lc["contrast_ratio"] == 1.0
    assert audit["summary"]["pass"] is False
    # every finding carries a fix hint
    for items in f.values():
        for item in items:
            assert item["fix"]


def test_audit_clean_document_passes(tmp_path):
    path = new_doc(tmp_path, "clean.docx", title="Accessible Doc")
    srv.add_heading(path, "Introduction", 1, at_end=True, backup=False)
    srv.add_heading(path, "Background", 2, at_end=True, backup=False)
    srv.insert_paragraphs(
        path, [{"text": "Read the full report at the project site."}],
        at_end=True, backup=False,
    )
    srv.add_hyperlink(path, "full report", "https://example.com/report",
                      backup=False)
    png = make_png(tmp_path / "ok.png")
    srv.add_image(path, str(png), at_end=True, backup=False)
    srv.set_image_alt_text(path, 0, "Bar chart of compliance results",
                           backup=False)
    srv.create_table(path, [["Name", "Value"], ["x", "1"]], at_end=True,
                     backup=False)  # header_row defaults to True
    audit = ax.audit_accessibility(DocxPackage(path))
    assert audit["summary"]["pass"] is True, audit["findings"]
    assert audit["summary"]["total"] == 0


def test_audit_no_heading_one(tmp_path):
    path = new_doc(tmp_path, title="T")
    srv.add_heading(path, "Only a Subsection", 2, at_end=True, backup=False)
    audit = ax.audit_accessibility(DocxPackage(path))
    issues = {x["issue"] for x in audit["findings"]["heading_hierarchy"]}
    assert "no_heading_1" in issues


# ----------------------------------------------------------------- image DPI


def test_dpi_low_then_ok(tmp_path):
    path = new_doc(tmp_path)
    png = make_png(tmp_path / "img.png", w=80, h=40)
    srv.add_image(path, str(png), at_end=True, width_pt=200, backup=False)

    # 80 px over 200pt (2.778 in) -> 28.8 DPI, way below 300
    res = ax.check_image_resolution(DocxPackage(path), min_dpi=300)
    (entry,) = res["images"]
    assert entry["status"] == "low"
    assert entry["pixels"] == [80, 40]
    assert abs(entry["dpi"][0] - 28.8) < 0.2
    assert abs(entry["dpi"][1] - 28.8) < 0.2
    assert res["low_resolution"] == [0]
    assert res["pass"] is False
    assert "28.8" in entry["fix"]

    # shrink the displayed size: 80 px over 19pt -> ~303 DPI
    srv.resize_image(path, 0, width_pt=19, backup=False)
    res2 = ax.check_image_resolution(DocxPackage(path), min_dpi=300)
    (entry2,) = res2["images"]
    assert entry2["status"] == "ok"
    assert entry2["dpi"][0] >= 300
    assert res2["pass"] is True


def test_dpi_vector_and_unchecked_formats(tmp_path):
    path = new_doc(tmp_path)
    bmp = tmp_path / "x.bmp"
    bmp.write_bytes(b"BM" + b"\x00" * 64)
    emf = tmp_path / "x.emf"
    emf.write_bytes(b"\x01\x00\x00\x00" + b"\x00" * 40)
    srv.add_image(path, str(bmp), at_end=True, backup=False)
    srv.add_image(path, str(emf), at_end=True, backup=False)
    res = ax.check_image_resolution(DocxPackage(path))
    statuses = {e["status"] for e in res["images"]}
    assert "unchecked (bmp)" in statuses
    assert "vector (not applicable)" in statuses
    # unverifiable formats are not counted as failures
    assert res["pass"] is True and res["low_resolution"] == []


def test_dpi_min_dpi_validated(tmp_path):
    path = new_doc(tmp_path)
    with pytest.raises(WordMcpError, match="min_dpi"):
        ax.check_image_resolution(DocxPackage(path), min_dpi=0)
