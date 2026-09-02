"""v1.6 scope item 5: API-review absorptions.

Covers: get_workflows (including the every-named-tool-exists registry
assertion), detect_citation_system (single system, split-brain, plain-text
only), change_heading_level (single + subtree, 1-9 bounds refusals,
outlineLvl-only refusal, ambiguity), insert_field/list_fields (allowlist,
SEQ validation, complex + fldSimple + header scanning), create_snapshot
(DTG naming, restacking, labels, collisions), content-control coverage
(list/set/insert plain-text, lock and unwritable-type refusals), and
insert_glossary (harvested definitions, [DEFINITION NEEDED] markers,
alphabetization, placement).
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest
from lxml import etree

import word_mcp.server as srv
from word_mcp.core.errors import (
    AmbiguousTarget,
    DocumentNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import backups as bk
from word_mcp.ops import citesystem as cs
from word_mcp.ops import definedterms as dterms
from word_mcp.ops import fields as fl
from word_mcp.ops import forms as fm
from word_mcp.ops import structure as sx
from word_mcp.ops import workflows as wf
from word_mcp.ops.fields import _field_run

def _make_doc(tmp_path: Path, name: str = "t.docx") -> Path:
    doc = tmp_path / name
    srv.create_document(str(doc))
    return doc


def _add_paras(doc: Path, texts: list[str]) -> None:
    srv.insert_paragraphs(
        str(doc), [{"text": t} for t in texts], live="off"
    )


def _texts(doc: Path) -> list[str]:
    return [p["text"] for p in srv.get_text(str(doc), live="off")]


def _outline(doc: Path) -> list[dict]:
    return srv.get_outline(str(doc), live="off")


def _mutate(doc: Path, fn):
    pkg = DocxPackage(str(doc))
    result = fn(pkg)
    pkg.save(do_backup=False)
    return result


# ------------------------------------------------------------------ workflows


def test_workflows_listing_names_all_tasks():
    out = wf.get_workflows()
    tasks = {t["task"] for t in out["tasks"]}
    assert tasks == {
        "process-feedback",
        "prepare-submission",
        "format-citations",
        "build-from-template",
        "heavy-editing",
        "migrate-from-v1",
        "bulk-edit",
    }
    for t in out["tasks"]:
        assert t["summary"]


def test_workflows_unknown_task_refused():
    with pytest.raises(WordMcpError, match="heavy-editing"):
        wf.get_workflows("format-apa")


def test_workflows_task_shape_and_heavy_editing_pattern():
    out = wf.get_workflows("heavy-editing")
    assert out["task"] == "heavy-editing"
    tools = [s["tool"] for s in out["steps"]]
    assert tools[0] == "copy_document"  # tester-recommended DTG-copy pattern
    for step in out["steps"]:
        assert step["why"], f"step {step['tool']} lacks a why"
    assert any("open in Word" in n for n in out["notes"])


def test_every_workflow_tool_exists_in_registry():
    """Every tool named in any workflow step must be registered on the
    server (with the absorptions snippet loaded, since some steps name the
    new tools)."""
    import asyncio
    import inspect

    res = srv.mcp.list_tools()
    if inspect.iscoroutine(res):
        res = asyncio.run(res)
    registered = {getattr(t, "name", None) or t["name"] for t in res}
    referenced = {
        step["tool"]
        for wfd in wf.WORKFLOWS.values()
        for step in wfd["steps"]
    }
    missing = referenced - registered - wf.PENDING_TOOLS
    assert not missing, f"workflow steps name unregistered tools: {missing}"


# ------------------------------------------------------ detect_citation_system


def test_detect_none_on_empty_doc(tmp_path):
    doc = _make_doc(tmp_path)
    res = cs.detect_citation_system(DocxPackage(str(doc)))
    assert res["system"] == "none"
    assert res["systems"] == {}
    assert res["split_brain"] is False


def test_detect_plain_text_only(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(
        doc,
        [
            "Deterrence theory failed to predict this (Smith, 2020).",
            "Jones (2019) argued the opposite from archival evidence.",
        ],
    )
    res = cs.detect_citation_system(DocxPackage(str(doc)))
    assert res["system"] == "plain_text_only"
    assert res["plain_text_citations"]["total"] >= 2
    assert res["split_brain"] is False


def test_detect_split_brain_word_native_plus_zotero(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["First paragraph.", "Second paragraph."])

    def inject(pkg):
        paras = pkg.body().findall(qn("w:p"))
        for el in _field_run("CITATION Smi20 \\l 1033", "(Smith, 2020)"):
            paras[0].append(el)
        for el in _field_run(
            'ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"abc"}',
            "(Jones, 2019)",
        ):
            paras[1].append(el)
        pkg.mark_dirty()
        return {}

    _mutate(doc, inject)
    res = cs.detect_citation_system(DocxPackage(str(doc)))
    assert res["systems"]["word_native"]["citations"] == 1
    assert res["systems"]["zotero"]["citations"] == 1
    assert res["split_brain"] is True
    assert "split-brain" in res["warning"]
    assert res["system"].startswith("mixed:")


def test_detect_single_manager_no_split_brain(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Cited here."])

    def inject(pkg):
        p = pkg.body().findall(qn("w:p"))[0]
        for el in _field_run(
            "ADDIN EN.CITE <EndNote><Cite/></EndNote>", "(Doe, 2018)"
        ):
            p.append(el)
        pkg.mark_dirty()
        return {}

    _mutate(doc, inject)
    res = cs.detect_citation_system(DocxPackage(str(doc)))
    assert res["system"] == "endnote"
    assert res["split_brain"] is False
    assert "warning" not in res


def test_detect_sees_fldsimple_bibliography(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Body."])

    def inject(pkg):
        p = pkg.body().findall(qn("w:p"))[0]
        fs = etree.SubElement(p, qn("w:fldSimple"))
        fs.set(qn("w:instr"), " BIBLIOGRAPHY ")
        r = etree.SubElement(fs, qn("w:r"))
        etree.SubElement(r, qn("w:t")).text = "Smith, A. (2020). Title."
        pkg.mark_dirty()
        return {}

    _mutate(doc, inject)
    res = cs.detect_citation_system(DocxPackage(str(doc)))
    assert res["systems"]["word_native"]["bibliographies"] == 1


# ------------------------------------------------------- change_heading_level


def _heading_doc(tmp_path) -> Path:
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "Intro", "heading_level": 1}])
    _add_paras(doc, ["Some intro text."])
    srv.insert_paragraphs(str(doc), [{"text": "Background", "heading_level": 2}])
    _add_paras(doc, ["Background prose."])
    srv.insert_paragraphs(str(doc), [{"text": "Details", "heading_level": 3}])
    srv.insert_paragraphs(str(doc), [{"text": "Methods", "heading_level": 2}])
    return doc


def _levels(doc: Path) -> dict[str, int]:
    return {h["text"]: h["level"] for h in _outline(doc)}


def test_demote_single_heading(tmp_path):
    doc = _heading_doc(tmp_path)
    r = _mutate(
        doc,
        lambda pkg: sx.change_heading_level(
            pkg, delta=1, heading_text="Details"
        ),
    )
    assert r["changed"] == [
        {
            "paragraph_index": r["changed"][0]["paragraph_index"],
            "text": "Details",
            "from_level": 3,
            "to_level": 4,
        }
    ]
    assert _levels(doc)["Details"] == 4


def test_demote_subtree_moves_subordinates_only(tmp_path):
    doc = _heading_doc(tmp_path)
    r = _mutate(
        doc,
        lambda pkg: sx.change_heading_level(
            pkg, delta=1, heading_text="Background", subtree=True
        ),
    )
    assert [(c["text"], c["to_level"]) for c in r["changed"]] == [
        ("Background", 3),
        ("Details", 4),
    ]
    levels = _levels(doc)
    assert levels == {"Intro": 1, "Background": 3, "Details": 4, "Methods": 2}
    # Non-heading prose between the headings is untouched.
    assert "Background prose." in _texts(doc)


def test_promote_above_level_1_refused_naming_blocker(tmp_path):
    doc = _heading_doc(tmp_path)
    with pytest.raises(WordMcpError, match="Intro.*above level 1"):
        _mutate(
            doc,
            lambda pkg: sx.change_heading_level(
                pkg, delta=-1, heading_text="Intro"
            ),
        )
    assert _levels(doc)["Intro"] == 1  # nothing changed


def test_subtree_refusal_is_atomic(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "Top", "heading_level": 8}])
    srv.insert_paragraphs(str(doc), [{"text": "Deep", "heading_level": 9}])
    with pytest.raises(WordMcpError, match="Deep.*below level 9"):
        _mutate(
            doc,
            lambda pkg: sx.change_heading_level(
                pkg, delta=1, heading_text="Top", subtree=True
            ),
        )
    assert _levels(doc) == {"Top": 8, "Deep": 9}


def test_outline_lvl_only_heading_refused(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Before."])

    def inject(pkg):
        body = pkg.body()
        p = etree.Element(qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:outlineLvl")).set(qn("w:val"), "1")
        r = etree.SubElement(p, qn("w:r"))
        etree.SubElement(r, qn("w:t")).text = "Fake Heading"
        sectpr = body.find(qn("w:sectPr"))
        if sectpr is not None:
            sectpr.addprevious(p)
        else:
            body.append(p)
        pkg.mark_dirty()
        return {}

    _mutate(doc, inject)
    assert _levels(doc)["Fake Heading"] == 2  # readers see it
    with pytest.raises(WordMcpError, match="outlineLvl"):
        _mutate(
            doc,
            lambda pkg: sx.change_heading_level(
                pkg, delta=1, heading_text="Fake Heading"
            ),
        )


def test_ambiguous_heading_text_refused(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "Twin", "heading_level": 2}])
    srv.insert_paragraphs(str(doc), [{"text": "Twin", "heading_level": 2}])
    with pytest.raises(AmbiguousTarget, match="paragraph_index"):
        _mutate(
            doc,
            lambda pkg: sx.change_heading_level(
                pkg, delta=1, heading_text="Twin"
            ),
        )


def test_change_by_paragraph_index_and_bad_delta(tmp_path):
    doc = _heading_doc(tmp_path)
    idx = next(
        h["paragraph_index"] for h in _outline(doc) if h["text"] == "Methods"
    )
    _mutate(
        doc,
        lambda pkg: sx.change_heading_level(pkg, delta=1, paragraph_index=idx),
    )
    assert _levels(doc)["Methods"] == 3
    with pytest.raises(WordMcpError, match="non-zero"):
        _mutate(
            doc,
            lambda pkg: sx.change_heading_level(
                pkg, delta=0, paragraph_index=idx
            ),
        )
    with pytest.raises(WordMcpError, match="exactly one"):
        _mutate(doc, lambda pkg: sx.change_heading_level(pkg, delta=1))


# ------------------------------------------------- insert_field / list_fields


def test_insert_field_and_list_roundtrip(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Signed on DATEHERE by both parties."])
    _mutate(
        doc,
        lambda pkg: fl.insert_field(
            pkg,
            field_code='DATE \\@ "yyyy-MM-dd"',
            after_anchor="DATEHERE",
            placeholder="(date)",
        ),
    )
    inv = fl.list_fields(DocxPackage(str(doc)))
    assert inv["total"] == 1
    entry = inv["fields"][0]
    assert entry["type"] == "DATE"
    assert entry["kind"] == "complex"
    assert entry["cached_result"] == "(date)"
    assert entry["paragraph_index"] == 0
    assert "word/document.xml" in inv["parts_scanned"]


def test_insert_field_allowlist_refusal(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["anchor text here"])
    with pytest.raises(WordMcpError, match="allowlist"):
        _mutate(
            doc,
            lambda pkg: fl.insert_field(
                pkg,
                field_code='INCLUDETEXT "C:\\\\evil.docx"',
                after_anchor="anchor",
            ),
        )
    assert fl.list_fields(DocxPackage(str(doc)))["total"] == 0


def test_seq_field_validation(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Exhibit anchor point."])
    with pytest.raises(WordMcpError, match="SEQ needs an identifier"):
        _mutate(
            doc,
            lambda pkg: fl.insert_field(pkg, field_code="SEQ", after_anchor="anchor"),
        )
    _mutate(
        doc,
        lambda pkg: fl.insert_field(
            pkg, field_code="SEQ Exhibit \\* Arabic", after_anchor="anchor"
        ),
    )
    inv = fl.list_fields(DocxPackage(str(doc)))
    assert inv["fields"][0]["type"] == "SEQ"


def test_field_code_unbalanced_quote_refused(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["x anchor y"])
    with pytest.raises(WordMcpError, match="unbalanced"):
        _mutate(
            doc,
            lambda pkg: fl.insert_field(
                pkg, field_code='DATE \\@ "yyyy', after_anchor="anchor"
            ),
        )


def test_list_fields_sees_fldsimple_and_headers(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Body paragraph."])
    srv.set_header_footer(str(doc), part="header", text="Chapter header")

    def inject(pkg):
        p = pkg.body().findall(qn("w:p"))[0]
        fs = etree.SubElement(p, qn("w:fldSimple"))
        fs.set(qn("w:instr"), " PAGE ")
        r = etree.SubElement(fs, qn("w:r"))
        etree.SubElement(r, qn("w:t")).text = "7"
        pkg.mark_dirty()
        return {}

    _mutate(doc, inject)
    inv = fl.list_fields(DocxPackage(str(doc)))
    simple = [f for f in inv["fields"] if f["kind"] == "simple"]
    assert simple and simple[0]["type"] == "PAGE"
    assert simple[0]["cached_result"] == "7"
    assert any(p.startswith("word/header") for p in inv["parts_scanned"])


# ------------------------------------------------------------- create_snapshot


import datetime as _real_datetime


class _FrozenDatetime:
    """Stand-in for the datetime module with a fixed now()."""

    class datetime:
        @staticmethod
        def now():
            return _real_datetime.datetime(2026, 8, 28, 14, 5, 0)

        fromtimestamp = staticmethod(_real_datetime.datetime.fromtimestamp)


def test_snapshot_dtg_naming_and_content(tmp_path, monkeypatch):
    monkeypatch.setattr(bk, "_dt", _FrozenDatetime)
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Snapshot me."])
    r = bk.create_snapshot(str(doc))
    dest = Path(r["snapshot"])
    assert dest.name == "20260828_1405_t.docx"
    assert dest.parent == doc.parent
    assert "Snapshot me." in [p["text"] for p in srv.get_text(str(dest), live="off")]
    # Source untouched.
    assert doc.exists()


def test_snapshot_label_collision_and_restack(tmp_path, monkeypatch):
    monkeypatch.setattr(bk, "_dt", _FrozenDatetime)
    doc = _make_doc(tmp_path, "20250101_0101_Draft.docx")
    r1 = bk.create_snapshot(str(doc), label="final")
    # Old DTG replaced, not stacked; label appended.
    assert Path(r1["snapshot"]).name == "20260828_1405_Draft_final.docx"
    r2 = bk.create_snapshot(str(doc), label="final")
    assert Path(r2["snapshot"]).name == "20260828_1405_Draft_final (2).docx"
    assert Path(r1["snapshot"]).exists() and Path(r2["snapshot"]).exists()


def test_snapshot_refusals_and_word_open_note(tmp_path):
    with pytest.raises(DocumentNotFound):
        bk.create_snapshot(str(tmp_path / "missing.docx"))
    doc = _make_doc(tmp_path)
    with pytest.raises(WordMcpError, match="label"):
        bk.create_snapshot(str(doc), label="a/b")
    owner = doc.with_name("~$" + doc.name[-153:])
    owner.write_bytes(b"stub")
    r = bk.create_snapshot(str(doc))
    assert "open in Word" in r.get("note", "")


# ----------------------------------------------------------- content controls


def test_insert_list_set_plain_text_control(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Client name: ____ (signature)."])
    _mutate(
        doc,
        lambda pkg: fm.insert_content_control(
            pkg, tag="client", after_anchor="____", alias="Client", text="ACME"
        ),
    )
    out = fm.list_content_controls(DocxPackage(str(doc)))
    assert out["count"] == 1
    ctrl = out["controls"][0]
    assert ctrl["tag"] == "client"
    assert ctrl["alias"] == "Client"
    assert ctrl["type"] == "text"
    assert ctrl["value"] == "ACME"
    assert ctrl["content_locked"] is False

    _mutate(
        doc,
        lambda pkg: fm.set_content_control_value(pkg, "Beta LLC", tag="client"),
    )
    out = fm.list_content_controls(DocxPackage(str(doc)))
    assert out["controls"][0]["value"] == "Beta LLC"

    # The new control is a first-class form field too.
    r = srv.set_form_fields(str(doc), {"client": "Gamma"}, backup=False)
    assert r["filled"] == {"client": "sdt_text"}


def test_insert_control_duplicate_tag_refused(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["spot one and spot two"])
    _mutate(
        doc,
        lambda pkg: fm.insert_content_control(
            pkg, tag="dup", after_anchor="spot one", text="a"
        ),
    )
    with pytest.raises(WordMcpError, match="already exists"):
        _mutate(
            doc,
            lambda pkg: fm.insert_content_control(
                pkg, tag="dup", after_anchor="spot two", text="b"
            ),
        )


def test_locked_control_refused(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["field: X end"])
    _mutate(
        doc,
        lambda pkg: fm.insert_content_control(
            pkg, tag="locked1", after_anchor="X", text="v1"
        ),
    )

    def add_lock(pkg):
        pr = next(pkg.root().iter(qn("w:sdt"))).find(qn("w:sdtPr"))
        etree.SubElement(pr, qn("w:lock")).set(qn("w:val"), "sdtContentLocked")
        pkg.mark_dirty()
        return {}

    _mutate(doc, add_lock)
    ctrl = fm.list_content_controls(DocxPackage(str(doc)))["controls"][0]
    assert ctrl["content_locked"] is True and ctrl["control_locked"] is True
    with pytest.raises(WordMcpError, match="locked"):
        _mutate(
            doc,
            lambda pkg: fm.set_content_control_value(pkg, "v2", tag="locked1"),
        )
    assert (
        fm.list_content_controls(DocxPackage(str(doc)))["controls"][0]["value"]
        == "v1"
    )


def test_unwritable_type_refused_and_listed(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Citation goes here."])

    def inject_citation_sdt(pkg):
        p = pkg.body().findall(qn("w:p"))[0]
        sdt = etree.SubElement(p, qn("w:sdt"))
        pr = etree.SubElement(sdt, qn("w:sdtPr"))
        etree.SubElement(pr, qn("w:citation"))
        content = etree.SubElement(sdt, qn("w:sdtContent"))
        r = etree.SubElement(content, qn("w:r"))
        etree.SubElement(r, qn("w:t")).text = "(Smith 2020)"
        pkg.mark_dirty()
        return {}

    _mutate(doc, inject_citation_sdt)
    out = fm.list_content_controls(DocxPackage(str(doc)))
    assert out["controls"][0]["type"] == "citation"
    with pytest.raises(UnsupportedStructure, match="citation"):
        _mutate(
            doc,
            lambda pkg: fm.set_content_control_value(pkg, "text", index=0),
        )


def test_checkbox_control_set_bool_only(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["Agree: box here."])

    def inject_checkbox(pkg):
        p = pkg.body().findall(qn("w:p"))[0]
        sdt = etree.SubElement(p, qn("w:sdt"))
        pr = etree.SubElement(sdt, qn("w:sdtPr"))
        etree.SubElement(pr, qn("w:tag")).set(qn("w:val"), "agree")
        cb = etree.SubElement(pr, qn("w14:checkbox"))
        etree.SubElement(cb, qn("w14:checked")).set(qn("w14:val"), "0")
        content = etree.SubElement(sdt, qn("w:sdtContent"))
        r = etree.SubElement(content, qn("w:r"))
        etree.SubElement(r, qn("w:t")).text = "☐"
        pkg.mark_dirty()
        return {}

    _mutate(doc, inject_checkbox)
    ctrl = fm.list_content_controls(DocxPackage(str(doc)))["controls"][0]
    assert ctrl["type"] == "checkbox" and ctrl["value"] is False
    with pytest.raises(WordMcpError, match="true/false"):
        _mutate(
            doc,
            lambda pkg: fm.set_content_control_value(pkg, "yes", tag="agree"),
        )
    _mutate(
        doc, lambda pkg: fm.set_content_control_value(pkg, True, tag="agree")
    )
    ctrl = fm.list_content_controls(DocxPackage(str(doc)))["controls"][0]
    assert ctrl["value"] is True


def test_set_control_addressing_errors(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["anchor"])
    with pytest.raises(WordMcpError, match="exactly one"):
        _mutate(doc, lambda pkg: fm.set_content_control_value(pkg, "v"))
    with pytest.raises(WordMcpError, match="no content control"):
        _mutate(
            doc, lambda pkg: fm.set_content_control_value(pkg, "v", tag="nope")
        )


# ------------------------------------------------------------- insert_glossary


def _contract_doc(tmp_path) -> Path:
    doc = _make_doc(tmp_path)
    _add_paras(
        doc,
        [
            '"Confidential Information" means any nonpublic data shared '
            "under this contract between the parties.",
            'The parties execute this contract (the "Effective Date") today.',
            '"Term" shall mean the period of five years from signing.',
        ],
    )
    return doc


def test_insert_glossary_harvests_and_marks(tmp_path):
    doc = _contract_doc(tmp_path)
    r = _mutate(doc, lambda pkg: dterms.insert_glossary(pkg))
    assert r["terms"] == 3
    assert r["needing_definition"] == ["Effective Date"]
    texts = _texts(doc)
    gi = texts.index("Glossary")
    assert texts[gi + 1 :][:3] == [
        "Confidential Information: any nonpublic data shared under this "
        "contract between the parties.",
        "Effective Date: [DEFINITION NEEDED]",
        "Term: the period of five years from signing.",
    ]
    # Heading is a real heading.
    assert {"text": "Glossary", "level": 1} == {
        k: v
        for k, v in next(
            h for h in _outline(doc) if h["text"] == "Glossary"
        ).items()
        if k in ("text", "level")
    }


def test_insert_glossary_placement_after_index(tmp_path):
    doc = _contract_doc(tmp_path)
    r = _mutate(
        doc, lambda pkg: dterms.insert_glossary(pkg, after_index=0)
    )
    assert r["position"] == "after paragraph 0"
    assert _texts(doc)[1] == "Glossary"


def test_insert_glossary_refuses_without_terms(tmp_path):
    doc = _make_doc(tmp_path)
    _add_paras(doc, ["No defined terms in this prose at all."])
    with pytest.raises(WordMcpError, match="no defined terms"):
        _mutate(doc, lambda pkg: dterms.insert_glossary(pkg))
