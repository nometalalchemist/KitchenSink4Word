"""Post-merge formatting follow-ups (2026-08-28 real-dissertation merge
review, V1.6_KICKOFF "FOLLOW-UP BUILDER QUEUE"):

1. insert_document `formatting` param: "source" | "merge" | "destination"
   (Word paste modes, stripping applied to the carried COPIES only).
2. insert_paragraphs inherit_format / copy_format_from (clone the anchor's
   direct pPr + terminal-run rPr onto inserted paragraphs — the 66
   unformatted reference entries case).
3. copy_table: single-table transplant through the insert_document
   reconciliation pipeline.
4. find_formatted no longer walks text-box paragraphs (mc:Choice +
   mc:Fallback double-read fix).
"""

import hashlib
import importlib.util
from pathlib import Path

import pytest
from docx import Document
from lxml import etree

import word_mcp.server as srv
from word_mcp.core.errors import (
    AmbiguousTarget,
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import assembly as asm
from word_mcp.ops import lists as ls
from word_mcp.ops import stylefind as sf
from word_mcp.ops import tables as tb
from word_mcp.ops import text as tx
from word_mcp.ops.read import body_items, paragraph_text

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ------------------------------------------------------------------- helpers


def _md5(path) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


def _fresh(tmp_path, name, texts):
    f = tmp_path / name
    doc = Document()
    for t in texts:
        doc.add_paragraph(t)
    doc.save(str(f))
    return f


def _add_para(pkg, text, *, style=None, ppr_children=(), rpr_children=()):
    """Append a body paragraph with explicit direct pPr/rPr children.
    Children are (localTag, {attr: val}) pairs in w: namespace."""
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    if style:
        etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), style)
    for tag, attrs in ppr_children:
        el = etree.SubElement(ppr, qn(tag))
        for k, v in attrs.items():
            el.set(qn(k), v)
    if len(ppr) == 0:
        p.remove(ppr)
    r = etree.SubElement(p, qn("w:r"))
    if rpr_children:
        rpr = etree.SubElement(r, qn("w:rPr"))
        for tag, attrs in rpr_children:
            el = etree.SubElement(rpr, qn(tag))
            for k, v in attrs.items():
                el.set(qn(k), v)
    t = etree.SubElement(r, qn("w:t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    body = pkg.body()
    sectpr = body.find(qn("w:sectPr"))
    if sectpr is not None:
        sectpr.addprevious(p)
    else:
        body.append(p)
    pkg.mark_dirty()
    return p


def _find_para(pkg, needle):
    for kind, _idx, el in body_items(pkg):
        if kind == "paragraph" and needle in paragraph_text(el):
            return el
    raise AssertionError(f"paragraph not found: {needle!r}")


def _rpr_tags(p):
    """Local names of the FIRST run's rPr children (empty set if no rPr)."""
    r = p.find(qn("w:r"))
    assert r is not None, "paragraph has no run"
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return set()
    return {etree.QName(c).localname for c in rpr}


def _ppr_tags(p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return set()
    return {etree.QName(c).localname for c in ppr}


_DIRECT_RPR = (
    ("w:b", {}),
    ("w:rFonts", {"w:ascii": "Courier New", "w:hAnsi": "Courier New"}),
    ("w:color", {"w:val": "FF0000"}),
    ("w:spacing", {"w:val": "20"}),
    ("w:sz", {"w:val": "20"}),
    ("w:szCs", {"w:val": "20"}),
    ("w:highlight", {"w:val": "yellow"}),
)
_DIRECT_PPR = (
    ("w:keepNext", {}),
    ("w:spacing", {"w:before": "240", "w:line": "480", "w:lineRule": "auto"}),
    ("w:ind", {"w:left": "720", "w:hanging": "720"}),
    ("w:jc", {"w:val": "center"}),
    ("w:outlineLvl", {"w:val": "2"}),
)


def _fmt_source(tmp_path, name="fmt_src.docx"):
    """Source with one direct-formatted paragraph (incl. tab stops) and a
    numbered list."""
    f = _fresh(tmp_path, name, ["Plain source paragraph"])
    pkg = DocxPackage(f)
    p = _add_para(
        pkg,
        "Direct formatted paragraph",
        ppr_children=_DIRECT_PPR,
        rpr_children=_DIRECT_RPR,
    )
    tabs = etree.SubElement(p.find(qn("w:pPr")), qn("w:tabs"))
    tab = etree.SubElement(tabs, qn("w:tab"))
    tab.set(qn("w:val"), "left")
    tab.set(qn("w:pos"), "1440")
    ls.add_list(pkg, ["item one", "item two"], kind="number", at_end=True)
    pkg.save(do_backup=False)
    return f


# ============================== item 1: insert_document formatting modes ====


def test_source_mode_default_preserves_direct_formatting(tmp_path):
    src = _fmt_source(tmp_path)
    tgt = _fresh(tmp_path, "tgt.docx", ["T0"])
    res = srv.insert_document(str(tgt), str(src), at_end=True, backup=False)
    assert res["formatting_mode"] == "source"
    assert "runs_stripped" not in res and "paragraphs_stripped" not in res
    p = _find_para(DocxPackage(tgt), "Direct formatted paragraph")
    assert {"b", "rFonts", "sz", "color", "highlight"} <= _rpr_tags(p)
    assert {"spacing", "ind", "jc", "keepNext"} <= _ppr_tags(p)


def test_merge_mode_keeps_emphasis_strips_font_and_layout(tmp_path):
    src = _fmt_source(tmp_path)
    src_md5 = _md5(src)
    tgt = _fresh(tmp_path, "tgt.docx", ["T0"])
    res = srv.insert_document(
        str(tgt), str(src), at_end=True, formatting="merge", backup=False
    )
    assert res["formatting_mode"] == "merge"
    assert res["runs_stripped"] >= 1
    assert res["paragraphs_stripped"] >= 1
    pkg = DocxPackage(tgt)
    p = _find_para(pkg, "Direct formatted paragraph")
    rpr = _rpr_tags(p)
    assert {"b", "highlight"} <= rpr  # semantic emphasis kept
    assert not rpr & {"rFonts", "sz", "szCs", "color", "spacing"}
    ppr = _ppr_tags(p)
    assert not ppr & {"spacing", "ind"}
    assert {"jc", "keepNext", "outlineLvl", "tabs"} <= ppr  # not in strip set
    # Numbering survives merge mode and is remapped, not stripped.
    lp = _find_para(pkg, "item one")
    assert lp.find(qn("w:pPr")).find(qn("w:numPr")) is not None
    assert res["lists_carried"] >= 1
    # The source file is never modified.
    assert _md5(src) == src_md5


def test_destination_mode_strips_all_but_structural(tmp_path):
    src = _fmt_source(tmp_path)
    tgt = _fresh(tmp_path, "tgt.docx", ["T0"])
    res = srv.insert_document(
        str(tgt), str(src), at_end=True, formatting="destination", backup=False
    )
    assert res["formatting_mode"] == "destination"
    assert res["runs_stripped"] >= 1 and res["paragraphs_stripped"] >= 1
    pkg = DocxPackage(tgt)
    p = _find_para(pkg, "Direct formatted paragraph")
    assert _rpr_tags(p) == set()  # every direct run property gone
    ppr = _ppr_tags(p)
    assert not ppr & {"spacing", "ind", "jc", "keepNext"}
    assert {"outlineLvl", "tabs"} <= ppr  # structurally required, stays
    lp = _find_para(pkg, "item one")
    assert lp.find(qn("w:pPr")).find(qn("w:numPr")) is not None
    # Content itself is intact.
    assert paragraph_text(p) == "Direct formatted paragraph"


def test_invalid_formatting_mode_refused_untouched(tmp_path):
    src = _fmt_source(tmp_path)
    tgt = _fresh(tmp_path, "tgt.docx", ["T0"])
    before = _md5(tgt)
    with pytest.raises(WordMcpError, match="formatting"):
        srv.insert_document(
            str(tgt), str(src), at_end=True, formatting="merged", backup=False
        )
    assert _md5(tgt) == before


# ================== item 2: insert_paragraphs inherit_format / copy_from ====


_ANCHOR_PPR = (
    ("w:spacing", {"w:after": "240"}),
    ("w:ind", {"w:left": "720", "w:hanging": "720"}),
)
_ANCHOR_RPR = (
    ("w:rFonts", {"w:ascii": "Garamond", "w:hAnsi": "Garamond"}),
    ("w:sz", {"w:val": "20"}),
)


def _ref_doc(tmp_path, *, numbered_anchor=False):
    """Doc whose paragraph index 1 ('Entry one') carries the direct hanging
    indent + font a reference entry should inherit."""
    f = _fresh(tmp_path, "refs.docx", ["References"])
    pkg = DocxPackage(f)
    ppr = _ANCHOR_PPR
    if numbered_anchor:
        ppr = ppr + (("w:outlineLvl", {"w:val": "0"}),)
    p = _add_para(
        pkg, "Entry one", ppr_children=ppr, rpr_children=_ANCHOR_RPR
    )
    if numbered_anchor:
        numpr = etree.SubElement(p.find(qn("w:pPr")), qn("w:numPr"))
        etree.SubElement(numpr, qn("w:ilvl")).set(qn("w:val"), "0")
        etree.SubElement(numpr, qn("w:numId")).set(qn("w:val"), "1")
    pkg.save(do_backup=False)
    return f


def test_inherit_format_after_index_clones_ppr_and_terminal_rpr(tmp_path):
    f = _ref_doc(tmp_path)
    pkg = DocxPackage(f)
    res = tx.insert_paragraphs(
        pkg, [{"text": "Entry two"}], after_index=1, inherit_format=True
    )
    assert res["inserted"] == 1
    assert res["format_cloned_from"] == 1
    p = _find_para(pkg, "Entry two")
    ppr = p.find(qn("w:pPr"))
    assert ppr is not None
    ind = ppr.find(qn("w:ind"))
    assert ind.get(qn("w:hanging")) == "720"
    assert ppr.find(qn("w:spacing")).get(qn("w:after")) == "240"
    rpr = p.find(qn("w:r")).find(qn("w:rPr"))
    assert rpr.find(qn("w:rFonts")).get(qn("w:ascii")) == "Garamond"
    assert rpr.find(qn("w:sz")).get(qn("w:val")) == "20"


def test_inherit_format_excludes_anchor_numbering(tmp_path):
    f = _ref_doc(tmp_path, numbered_anchor=True)
    pkg = DocxPackage(f)
    tx.insert_paragraphs(
        pkg, [{"text": "Entry two"}], after_index=1, inherit_format=True
    )
    p = _find_para(pkg, "Entry two")
    ppr = p.find(qn("w:pPr"))
    assert ppr.find(qn("w:numPr")) is None  # numbering never cloned
    assert ppr.find(qn("w:ind")) is not None  # rest of the pPr still is


def test_inherit_format_after_anchor(tmp_path):
    f = _ref_doc(tmp_path)
    pkg = DocxPackage(f)
    res = tx.insert_paragraphs(
        pkg, [{"text": "Entry two"}], after_anchor="Entry one",
        inherit_format=True,
    )
    assert res["format_cloned_from"] == 1
    p = _find_para(pkg, "Entry two")
    assert p.find(qn("w:pPr")).find(qn("w:ind")) is not None


def test_copy_format_from_with_at_end(tmp_path):
    f = _ref_doc(tmp_path)
    pkg = DocxPackage(f)
    res = tx.insert_paragraphs(
        pkg, [{"text": "Entry two"}, {"text": "Entry three"}],
        at_end=True, copy_format_from=1,
    )
    assert res["format_cloned_from"] == 1
    for text in ("Entry two", "Entry three"):
        p = _find_para(pkg, text)
        assert p.find(qn("w:pPr")).find(qn("w:ind")) is not None


def test_inherit_and_copy_from_mutually_exclusive(tmp_path):
    pkg = DocxPackage(_ref_doc(tmp_path))
    with pytest.raises(WordMcpError, match="mutually exclusive"):
        tx.insert_paragraphs(
            pkg, [{"text": "x"}], after_index=1,
            inherit_format=True, copy_format_from=0,
        )


def test_inherit_format_needs_an_anchor_positioner(tmp_path):
    pkg = DocxPackage(_ref_doc(tmp_path))
    with pytest.raises(WordMcpError, match="copy_format_from"):
        tx.insert_paragraphs(
            pkg, [{"text": "x"}], before_index=1, inherit_format=True
        )
    with pytest.raises(WordMcpError, match="copy_format_from"):
        tx.insert_paragraphs(
            pkg, [{"text": "x"}], at_end=True, inherit_format=True
        )


def test_copy_format_from_out_of_range(tmp_path):
    pkg = DocxPackage(_ref_doc(tmp_path))
    with pytest.raises(TargetNotFound):
        tx.insert_paragraphs(
            pkg, [{"text": "x"}], at_end=True, copy_format_from=99
        )


def test_explicit_item_style_wins_over_clone(tmp_path):
    f = _ref_doc(tmp_path)
    pkg = DocxPackage(f)
    tx.insert_paragraphs(
        pkg, [{"text": "Styled entry", "style": "Heading1"}],
        after_index=1, inherit_format=True,
    )
    p = _find_para(pkg, "Styled entry")
    ppr = p.find(qn("w:pPr"))
    assert ppr.find(qn("w:pStyle")).get(qn("w:val")) == "Heading1"
    assert ppr.find(qn("w:ind")) is not None  # clone still applied around it


def test_explicit_item_formatting_wins_over_clone(tmp_path):
    f = _ref_doc(tmp_path)
    pkg = DocxPackage(f)
    tx.insert_paragraphs(
        pkg, [{"text": "Big entry", "formatting": {"size_pt": 14}}],
        after_index=1, inherit_format=True,
    )
    p = _find_para(pkg, "Big entry")
    rpr = p.find(qn("w:r")).find(qn("w:rPr"))
    assert rpr.find(qn("w:sz")).get(qn("w:val")) == "28"  # explicit 14pt
    assert rpr.find(qn("w:rFonts")).get(qn("w:ascii")) == "Garamond"  # clone


def test_inherit_format_via_server_tool(tmp_path):
    f = _ref_doc(tmp_path)
    res = srv.insert_paragraphs(
        str(f), [{"text": "Entry two"}], after_index=1,
        inherit_format=True, backup=False, live="off",
    )
    assert res["format_cloned_from"] == 1
    p = _find_para(DocxPackage(f), "Entry two")
    assert p.find(qn("w:pPr")).find(qn("w:ind")) is not None


def test_plain_insert_unchanged_without_new_params(tmp_path):
    """Regression: default path (no clone) still produces bare paragraphs."""
    pkg = DocxPackage(_ref_doc(tmp_path))
    res = tx.insert_paragraphs(pkg, [{"text": "Bare entry"}], after_index=1)
    assert "format_cloned_from" not in res
    p = _find_para(pkg, "Bare entry")
    assert p.find(qn("w:pPr")) is None


# ======================================== item 3: copy_table transplant ====


def _table_source(tmp_path):
    f = _fresh(tmp_path, "tbl_src.docx", ["Table doc intro"])
    pkg = DocxPackage(f)
    tb.create_table(pkg, [["A1", "B1"], ["A2", "B2"]], at_end=True)
    tb.create_table(pkg, [["X1"], ["X2"], ["X3"]], at_end=True)
    # A source-only style used inside the second table, to exercise the
    # style-reconciliation pass.
    root = pkg.root("word/styles.xml")
    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), "paragraph")
    s.set(qn("w:styleId"), "SrcCellStyle")
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), "Source Cell Style")
    pkg.mark_dirty("word/styles.xml")
    tbl2 = [
        el for el in pkg.body() if etree.QName(el).localname == "tbl"
    ][1]
    cell_p = next(tbl2.iter(qn("w:p")))
    ppr = etree.Element(qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "SrcCellStyle")
    cell_p.insert(0, ppr)
    pkg.save(do_backup=False)
    return f


def _target_tables(pkg):
    return [el for el in pkg.body() if etree.QName(el).localname == "tbl"]


def test_copy_table_basic_position_and_content(tmp_path):
    src = _table_source(tmp_path)
    src_md5 = _md5(src)
    tgt = _fresh(tmp_path, "tgt.docx", ["T0", "T1"])
    pkg = DocxPackage(tgt)
    res = asm.copy_table(pkg, str(src), 0, after_index=0)
    assert res["tables"] == 1 and res["paragraphs"] == 0
    assert res["source_table_index"] == 0
    assert res["rows"] == 2 and res["columns"] == 2
    assert res["position"]["starts_at_body_item"] == 1
    tables = _target_tables(pkg)
    assert len(tables) == 1
    cell_text = "".join(t.text or "" for t in tables[0].iter(qn("w:t")))
    assert "A1" in cell_text and "B2" in cell_text
    # Body order: T0, table, T1.
    kinds = [etree.QName(el).localname for el in pkg.body()][:3]
    assert kinds == ["p", "tbl", "p"]
    assert _md5(src) == src_md5  # source never modified


def test_copy_table_second_table_clones_style(tmp_path):
    src = _table_source(tmp_path)
    tgt = _fresh(tmp_path, "tgt.docx", ["T0"])
    pkg = DocxPackage(tgt)
    res = asm.copy_table(pkg, str(src), 1, at_end=True)
    assert res["rows"] == 3 and res["columns"] == 1
    cloned_names = {c["name"] for c in res["styles"]["cloned"]}
    assert "Source Cell Style" in cloned_names
    # The cloned style really exists in the target's styles part.
    names = {
        s.find(qn("w:name")).get(qn("w:val"))
        for s in pkg.root("word/styles.xml").findall(qn("w:style"))
        if s.find(qn("w:name")) is not None
    }
    assert "Source Cell Style" in names


def test_copy_table_index_out_of_range(tmp_path):
    src = _table_source(tmp_path)
    pkg = DocxPackage(_fresh(tmp_path, "tgt.docx", ["T0"]))
    with pytest.raises(TargetNotFound, match="out of range"):
        asm.copy_table(pkg, str(src), 2, at_end=True)


def test_copy_table_no_tables_in_source(tmp_path):
    src = _fresh(tmp_path, "empty_src.docx", ["No tables here"])
    pkg = DocxPackage(_fresh(tmp_path, "tgt.docx", ["T0"]))
    with pytest.raises(TargetNotFound, match="no top-level body tables"):
        asm.copy_table(pkg, str(src), 0, at_end=True)


def test_copy_table_positioner_contract(tmp_path):
    src = _table_source(tmp_path)
    pkg = DocxPackage(_fresh(tmp_path, "tgt.docx", ["T0"]))
    with pytest.raises(WordMcpError, match="exactly one positioner"):
        asm.copy_table(pkg, str(src), 0)
    with pytest.raises(WordMcpError, match="exactly one positioner"):
        asm.copy_table(pkg, str(src), 0, after_index=0, at_end=True)


def test_copy_table_ambiguous_anchor_refuses(tmp_path):
    src = _table_source(tmp_path)
    pkg = DocxPackage(_fresh(tmp_path, "tgt.docx", ["Dup", "Dup"]))
    with pytest.raises(AmbiguousTarget):
        asm.copy_table(pkg, str(src), 0, after_anchor="Dup")


def test_copy_table_blocked_content_refuses_untouched(tmp_path):
    src = _table_source(tmp_path)
    spkg = DocxPackage(src)
    tbl = _target_tables(spkg)[0]
    cell_p = next(tbl.iter(qn("w:p")))
    r = etree.SubElement(cell_p, qn("w:r"))
    etree.SubElement(r, qn("w:object"))
    spkg.mark_dirty()
    spkg.save(do_backup=False)
    pkg = DocxPackage(_fresh(tmp_path, "tgt.docx", ["T0"]))
    blocks_before = len(list(pkg.body()))
    with pytest.raises(UnsupportedStructure, match="OLE"):
        asm.copy_table(pkg, str(src), 0, at_end=True)
    assert len(list(pkg.body())) == blocks_before  # nothing half-applied


def test_copy_table_registration_snippet_loads():
    """Paste-readiness: the integration snippet imports and registers."""
    root = Path(__file__).parents[2]
    snippet = root / "integration" / "copytable_registrations.py"
    if not snippet.exists():
        pytest.skip("integration/ staging dir not present (gitignored)")
    spec = importlib.util.spec_from_file_location(
        "copytable_registrations_smoke", snippet,
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert hasattr(mod, "copy_table")


# ============================ item 4: find_formatted textbox exclusion ====


def _txbx_run_xml(box_text, name="Text Box 1"):
    """AlternateContent shape (wps:txbx) WITH the mc:Fallback VML copy Word
    always writes alongside — the doubled-storage case. The box paragraph's
    run is BOLD so it would satisfy a bold criterion if walked."""
    content = (
        "<w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr>"
        f"<w:t>{box_text}</w:t></w:r></w:p></w:txbxContent>"
    )
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


@pytest.fixture
def boxed_bold_doc(tmp_path):
    f = _fresh(tmp_path, "box.docx", ["Host paragraph."])
    pkg = DocxPackage(f)
    _add_para(pkg, "Body bold text", rpr_children=(("w:b", {}),))
    host = _find_para(pkg, "Host paragraph.")
    host.append(etree.fromstring(_txbx_run_xml("Boxed bold text")))
    pkg.save(do_backup=False)
    return f


def test_find_formatted_excludes_textbox_content(boxed_bold_doc):
    res = sf.find_formatted(
        DocxPackage(boxed_bold_doc), formatting={"bold": True}
    )
    texts = [m["text"] for m in res["matches"]]
    # Pre-fix this reported "Boxed bold text" TWICE (mc:Choice + mc:Fallback).
    assert all("Boxed" not in t for t in texts)
    assert texts.count("Body bold text") == 1
    assert res["total"] == 1


def test_find_formatted_query_ignores_box_text(boxed_bold_doc):
    res = sf.find_formatted(
        DocxPackage(boxed_bold_doc),
        "Boxed bold text",
        formatting={"bold": True},
    )
    assert res["total"] == 0
    assert res["matches"] == []
