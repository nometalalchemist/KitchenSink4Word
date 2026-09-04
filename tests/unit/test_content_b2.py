"""CONTENT B2 bundle gate: redaction, defined-terms audit, PDF import.

Synthetic documents are built with the server-layer functions (decorated
tools are plain functions); redaction and defined-terms are exercised
through the ops modules. The PDF-import test needs Word's converter, so it
is marked live and skips where Word is unavailable; the convert module
spawns and quits its own invisible instances.
"""

import shutil
from pathlib import Path

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import definedterms as dt, fields as fl, read as rd
from word_mcp.ops import redaction as rx

SECRET = "SECRET123"


def new_doc(tmp_path, name="doc.docx"):
    path = str(tmp_path / name)
    srv.create_document(path)
    return path


def doc_text(path):
    return " || ".join(p["text"] for p in srv.get_text(path, live="off"))


def idx_of(path, fragment):
    for p in srv.get_text(path, live="off"):
        if p["index"] is not None and fragment in p["text"]:
            return p["index"]
    raise AssertionError(f"paragraph not found: {fragment!r}")


# ------------------------------------------------------------------ redaction


def test_redact_fragmented_run(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": f"The launch code is {SECRET} today."}], backup=False,
    )
    # Bold half the secret: the match now spans two runs.
    srv.format_text(path, {"bold": True}, find="SECRET", backup=False)

    pkg = DocxPackage(path)
    res = rx.redact_text(pkg, [{"find": SECRET}])
    pkg.save(do_backup=False)

    assert res["redacted"].get("body") == 1
    assert res["total"] == 1
    assert res["verified_clean"] is True
    text = doc_text(path)
    assert SECRET not in text
    assert "The launch code is [REDACTED] today." in text


def test_redact_regex_target(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Codes ALPHA-7 and ALPHA-99 are sensitive."}], backup=False,
    )
    pkg = DocxPackage(path)
    res = rx.redact_text(
        pkg, [{"find": r"ALPHA-\d+", "regex": True}], replacement="X"
    )
    pkg.save(do_backup=False)
    assert res["redacted"].get("body") == 2
    assert "Codes X and X are sensitive." in doc_text(path)


def test_redact_all_location_classes_and_verify(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": f"Body mentions {SECRET} and a link here."}], backup=False,
    )
    srv.set_header_footer(path, part="header", text=f"Header {SECRET}", backup=False)
    srv.manage_note(path, action="insert", note_type="footnote", text=f"Footnote {SECRET}.", location={"search": {"text": "mentions"}}, backup=False)
    srv.insert_hyperlink(
        path, "link here", f"https://example.com/{SECRET}", backup=False
    )
    srv.manage_comment(path, action="add", text=f"Comment {SECRET}", location={"search": {"text": "Body"}}, backup=False)
    srv.set_document_properties(
        path, subject=f"About {SECRET}", keywords=f"{SECRET}, draft",
        backup=False,
    )
    # Hyperlink tooltip (no server tool sets tooltips; set the attribute).
    pkg = DocxPackage(path)
    link = pkg.root().find(f".//{qn('w:hyperlink')}")
    assert link is not None
    link.set(qn("w:tooltip"), f"tip {SECRET}")
    pkg.mark_dirty()
    pkg.save(do_backup=False)

    unredacted = str(tmp_path / "unredacted.docx")
    shutil.copy2(path, unredacted)

    pkg = DocxPackage(path)
    res = rx.redact_text(pkg, [{"find": SECRET}])
    pkg.save(do_backup=False)

    for cls in (
        "body", "headers_footers", "footnotes", "comments",
        "doc_properties", "hyperlink_tooltips", "hyperlink_urls",
    ):
        assert res["redacted"].get(cls, 0) >= 1, f"nothing redacted in {cls}"
    assert res["verified_clean"] is True, res.get("residual")
    assert res["scrubbed_location_classes"]
    assert isinstance(res["not_examined"], list)

    # Standalone verifier: redacted file clean, untouched copy not clean.
    assert rx.verify_redaction(DocxPackage(path), [{"find": SECRET}])[
        "clean"
    ] is True
    check = rx.verify_redaction(DocxPackage(unredacted), [{"find": SECRET}])
    assert check["clean"] is False
    assert check["residual"]


def test_redact_field_instruction_and_cached_result(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Field host paragraph."}], backup=False
    )
    pkg = DocxPackage(path)
    host = [
        el for k, _, el in rd.body_items(pkg) if k == "paragraph"
    ][-1]
    for el in fl._field_run(
        f'HYPERLINK "https://x.example/{SECRET}"', f"cached {SECRET} result"
    ):
        host.append(el)
    pkg.mark_dirty()
    pkg.save(do_backup=False)

    pkg = DocxPackage(path)
    res = rx.redact_text(pkg, [{"find": SECRET}])
    pkg.save(do_backup=False)
    assert res["redacted"].get("field_results") == 1
    assert res["redacted"].get("field_instructions") == 1
    assert res["verified_clean"] is True, res.get("residual")


def test_zero_length_regex_refused(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": f"keep {SECRET} safe"}], backup=False
    )
    before = doc_text(path)
    pkg = DocxPackage(path)
    with pytest.raises(WordMcpError):
        rx.redact_text(pkg, [{"find": "x*", "regex": True}])
    with pytest.raises(WordMcpError):
        rx.verify_redaction(pkg, [{"find": "x*", "regex": True}])
    assert doc_text(path) == before


def test_atomic_on_bad_target(tmp_path):
    """One valid target plus one refused target: NOTHING is applied —
    validation runs before any mutation, and nothing was saved."""
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": f"keep {SECRET} safe"}], backup=False
    )
    before = doc_text(path)
    pkg = DocxPackage(path)
    with pytest.raises(WordMcpError):
        rx.redact_text(pkg, [{"find": SECRET}, {"find": ""}])
    in_memory = " ".join(
        rd.paragraph_text(el)
        for k, _, el in rd.body_items(pkg)
        if k == "paragraph"
    )
    assert SECRET in in_memory  # in-memory tree untouched
    assert doc_text(path) == before  # file untouched


def test_unknown_scope_refused(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": SECRET}], backup=False
    )
    with pytest.raises(WordMcpError):
        rx.redact_text(DocxPackage(path), [{"find": SECRET}], scope="bogus")


# -------------------------------------------------------------- defined terms

CONTRACT = [
    'This Master Agreement (the "Agreement") is made by the parties.',
    "The parties expect that each License Fee will be invoiced monthly.",
    '"Confidential Information" means any nonpublic information disclosed '
    "under this Agreement.",
    '"Services" means the consulting work described in the statement of '
    "work.",
    '"Escrow Agent" means the bank named in the escrow schedule.',
    'The monthly charge for the work (the "License Fee") is due within '
    "thirty days.",
    '"Services" shall mean any additional work the parties agree to in '
    "writing.",
    "The receiving party shall protect Confidential Information and use "
    "the Services only as permitted.",
    "The obligations bind the disclosing party and the receiving party "
    '(each, a "Party"), and no Party may assign this Agreement.',
    "Each invoice covers one Delivery Milestone, and no Delivery Milestone "
    "may be billed twice.",
    "Acceptance of a Delivery Milestone requires written notice under this "
    "Agreement.",
    "Final Acceptance occurs when the committee approves the work.",
    "Final Acceptance may be delayed by written notice.",
]


@pytest.fixture()
def contract_doc(tmp_path):
    path = new_doc(tmp_path, "contract.docx")
    srv.insert_paragraphs(
        path, [{"text": t} for t in CONTRACT], backup=False
    )
    return path


def test_defined_terms_definitions_found(contract_doc):
    res = dt.check_defined_terms(DocxPackage(contract_doc))
    terms = {d["term"] for d in res["defined_terms"]}
    assert {
        "Agreement", "Confidential Information", "Services",
        "Escrow Agent", "License Fee", "Party",
    } <= terms


def test_defined_never_used(contract_doc):
    res = dt.check_defined_terms(DocxPackage(contract_doc))
    never = {d["term"] for d in res["defined_never_used"]}
    assert "Escrow Agent" in never
    for clean in ("Agreement", "Confidential Information", "Services",
                  "Party"):
        assert clean not in never


def test_used_never_defined_heuristic(contract_doc):
    res = dt.check_defined_terms(DocxPackage(contract_doc))
    flagged = {d["term"]: d for d in res["used_never_defined"]}
    assert "Delivery Milestone" in flagged
    entry = flagged["Delivery Milestone"]
    assert entry["count"] == 3
    assert idx_of(contract_doc, "Each invoice covers") in entry["paragraphs"]
    assert "review candidate" in entry["note"]
    # Only ever capitalized at sentence starts -> not a term.
    assert "Final Acceptance" not in flagged
    # Defined terms never appear in the heuristic list.
    assert "Confidential Information" not in flagged


def test_defined_multiple_times(contract_doc):
    res = dt.check_defined_terms(DocxPackage(contract_doc))
    dup = {d["term"]: d["defined_at"] for d in res["defined_multiple_times"]}
    assert set(dup) == {"Services"}
    assert dup["Services"] == [
        idx_of(contract_doc, '"Services" means'),
        idx_of(contract_doc, '"Services" shall mean'),
    ]


def test_first_use_before_definition(contract_doc):
    res = dt.check_defined_terms(DocxPackage(contract_doc))
    fub = {d["term"]: d for d in res["first_use_before_definition"]}
    assert set(fub) == {"License Fee"}
    entry = fub["License Fee"]
    assert entry["first_use_paragraph"] == idx_of(
        contract_doc, "invoiced monthly"
    )
    assert entry["first_definition_paragraph"] == idx_of(
        contract_doc, "monthly charge"
    )


def test_definition_pattern_without_group_refused(contract_doc):
    with pytest.raises(WordMcpError):
        dt.check_defined_terms(
            DocxPackage(contract_doc),
            definition_patterns=[r'"[A-Z]\w+" means'],
        )


def test_custom_definition_pattern(contract_doc):
    # Only the «shall mean» form: exactly the second Services definition.
    res = dt.check_defined_terms(
        DocxPackage(contract_doc),
        definition_patterns=[r'"([A-Z][^"\n]{0,80}?)"\s+shall\s+mean\b'],
    )
    assert {d["term"] for d in res["defined_terms"]} == {"Services"}
