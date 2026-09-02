"""V2 STAGING copy of tests/unit/test_vb_adversarial_fixes.py (Wave A).
Only change: srv.check_defined_terms becomes srv.validate(
checks=["defined_terms"]) with the results/findings unwrap per the Wave
A brief. Other waves' tool calls (redact_text, add_bookmark,
add_cross_reference, apply_style, setup_chapter_headers, com_import_pdf)
are left at v1 names for their owning waves/integrator to align.
"""

"""Regressions for the Wave B adversarial findings (B1-B8)."""

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage

from test_live_core import _word_available


def test_b1_redaction_counts_attribute_correctly(tmp_path):
    """Body hits near a complex field must count under body, not
    field_results (dead lxml proxy ids used to misattribute them)."""
    path = tmp_path / "b1.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "secret alpha in plain body text."},
         {"text": "another paragraph with secret beta."}], backup=False,
    )
    srv.insert_bookmark(str(path), name="anchor", anchor_text="plain body",
                     backup=False)
    srv.insert_cross_reference(
        str(path), to_bookmark="anchor",
        location={"search": {"text": "another paragraph"}},
        backup=False,
    )
    r = srv.redact_text(str(path), [{"find": "secret"}], backup=False)
    counts = r["redacted"]
    assert r["verified_clean"] is True
    assert counts.get("body") == 2, counts
    assert counts.get("field_results", 0) == 0, counts


def test_b2_diagnose_malformed_document_not_ok(tmp_path):
    import zipfile

    path = tmp_path / "b2.docx"
    srv.create_document(str(path))
    broken = tmp_path / "broken.docx"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(
        broken, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = data[:-40]  # truncate: malformed XML
            zout.writestr(item, data)
    r = srv.diagnose_document(str(broken))
    assert r["ok"] is False
    assert any(p["severity"] == "error" for p in r["problems"])


def test_b3_b4_front_matter_spec_validation(tmp_path):
    path = tmp_path / "b3.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(str(path), [{"text": "Body"}],
                          backup=False)
    with pytest.raises(WordMcpError, match="levels"):
        srv.assemble_front_matter(
            str(path),
            {"sections": [{"kind": "toc", "levels": "banana"}]},
            backup=False,
        )
    with pytest.raises(WordMcpError, match="levels"):
        srv.assemble_front_matter(
            str(path), {"sections": [{"kind": "toc", "levels": "3-1"}]},
            backup=False,
        )
    with pytest.raises(WordMcpError, match="title"):
        srv.assemble_front_matter(
            str(path), {"sections": [{"kind": "toc", "title": 123}]},
            backup=False,
        )


def test_b6_pdf_magic_byte_refusal(tmp_path):
    fake = tmp_path / "fake.pdf"
    fake.write_text("this is not a pdf at all", encoding="utf-8")
    with pytest.raises(WordMcpError, match="not a PDF"):
        srv.com_import_pdf(str(fake))


def test_b7_shadowed_term_not_counted(tmp_path):
    path = tmp_path / "b7.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [
            {"text": 'The parties execute this instrument (the "Master '
                     'Agreement").'},
            {"text": 'Payment obligations arise under the Master Agreement '
                     'as stated.'},
            {"text": 'Delivery follows the Master Agreement schedule as '
                     'planned by the parties.'},
            {"text": '"Agreement" means this document together with all '
                     'exhibits.'},
            {"text": "Nothing here uses the shorter term on its own."},
        ], backup=False,
    )
    r = srv.validate(str(path), checks=["defined_terms"])
    findings = r["results"]["defined_terms"]["findings"]
    ubd = [e["term"] for e in findings.get("first_use_before_definition", [])]
    assert "Agreement" not in ubd, r


def test_b8_chapter_headers_skip_front_matter_section(tmp_path):
    if not _word_available():
        pytest.skip("indirectly builds on toc field write; still file-based")
    path = tmp_path / "b8.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "Chapter One"}, {"text": "Chapter body text."}], backup=False,
    )
    srv.apply_style(str(path), style="Heading 1", range={"start": 0, "end": 0}, backup=False)
    srv.assemble_front_matter(
        str(path),
        {"sections": [
            {"kind": "title_page", "lines": ["A Title"]},
            {"kind": "abstract", "title": "Abstract", "text": "Summary."},
        ]},
        backup=False,
    )
    r = srv.setup_chapter_headers(str(path), backup=False)
    touched = [
        s["section"] if isinstance(s, dict) else s for s in r["sections"]
    ]
    assert 0 not in touched, r
