"""EQUATIONS bundle gate: LaTeX -> OMML via latex2mathml + mathml2omml.

Documents are built with the server-layer functions (decorated tools are
plain functions); equations are exercised through the ops module. The gate
cases mirror the v1.5 research fidelity table. The Word round-trip test is
marked live and skips where Word is unavailable; it spawns and quits its own
invisible instance.
"""

import shutil
from pathlib import Path

import pytest
from lxml import etree

import word_mcp.server as srv
from word_mcp.core.errors import TargetNotFound, WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import equations as eq

M = eq.M_NS


def new_doc(tmp_path, name="doc.docx"):
    path = str(tmp_path / name)
    srv.create_document(path)
    return path


def add(path, latex, **kw):
    pkg = DocxPackage(path)
    result = eq.add_equation(pkg, latex, **kw)
    pkg.save(do_backup=False)
    return result


def doc_xml(path):
    return etree.fromstring(DocxPackage(path).raw_part("word/document.xml"))


def m_tags(path):
    """Set of local names of all math-namespace elements in the body."""
    return {
        etree.QName(el).localname
        for el in doc_xml(path).iter()
        if etree.QName(el).namespace == M
    }


def listed(path):
    return eq.list_equations(DocxPackage(path))


# ------------------------------------------------------------- gate: fidelity


def test_fraction_display(tmp_path):
    path = new_doc(tmp_path)
    r = add(path, r"\frac{a+b}{c}", at_end=True)
    assert r["equation_added"] and r["display"]
    tags = m_tags(path)
    assert "f" in tags and "num" in tags and "den" in tags
    assert "oMathPara" in tags
    out = listed(path)
    assert out["equation_count"] == 1
    assert out["equations"][0]["display"] is True


def test_sqrt(tmp_path):
    path = new_doc(tmp_path)
    add(path, r"\sqrt{x^2+1}", at_end=True)
    assert "rad" in m_tags(path)


def test_sum_with_limits(tmp_path):
    path = new_doc(tmp_path)
    add(path, r"\sum_{i=1}^{n} i^2", at_end=True)
    tags = m_tags(path)
    assert "nary" in tags and "sub" in tags and "sup" in tags


def test_pmatrix(tmp_path):
    path = new_doc(tmp_path)
    add(path, r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}", at_end=True)
    root = doc_xml(path)
    assert root.find(f".//{{{M}}}m") is not None  # matrix
    beg = root.find(f".//{{{M}}}begChr")
    assert beg is not None and beg.get(f"{{{M}}}val") == "("


def test_cases(tmp_path):
    path = new_doc(tmp_path)
    add(
        path,
        r"f(x) = \begin{cases} x^2 & x \ge 0 \\ -x & x < 0 \end{cases}",
        at_end=True,
    )
    root = doc_xml(path)
    assert root.find(f".//{{{M}}}m") is not None
    beg = root.find(f".//{{{M}}}begChr")
    assert beg is not None and beg.get(f"{{{M}}}val") == "{"


def test_greek_and_operators(tmp_path):
    path = new_doc(tmp_path)
    add(path, r"\alpha \Gamma \leq \neq \approx \times \partial \nabla", at_end=True)
    text = listed(path)["equations"][0]["text"]
    for ch in "αΓ≤≠≈×∂∇":
        assert ch in text


def test_subscript_superscript_chains(tmp_path):
    path = new_doc(tmp_path)
    add(path, r"x_i^2 + e^{-x^2}", at_end=True)
    tags = m_tags(path)
    assert "sSubSup" in tags and "sSup" in tags


def test_aligned_rewrite_still_needed_and_works(tmp_path):
    # Canary: the upstream latex2mathml bug (malformed XML for `aligned`)
    # must still be present — if this starts passing, the rewrite can go.
    import latex2mathml.converter as l2m
    import mathml2omml

    raw = r"\begin{aligned} a &= b \\ c &= d \end{aligned}"
    with pytest.raises(Exception):
        mathml2omml.convert(l2m.convert(raw, display="block"))

    # Through the rewrite it converts, and matches the align* structure.
    path = new_doc(tmp_path)
    r = add(path, raw, at_end=True)
    assert r["equation_added"]
    assert "m" in m_tags(path)  # the align block becomes a matrix

    path2 = new_doc(tmp_path, "star.docx")
    add(path2, r"\begin{align*} a &= b \\ c &= d \end{align*}", at_end=True)
    assert listed(path)["equations"][0]["text"] == (
        listed(path2)["equations"][0]["text"]
    )


# ---------------------------------------------------------- gate: placement


def test_display_positioning_modes(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path,
        [{"text": "First paragraph."}, {"text": "Second paragraph."}],
        at_end=True,
        backup=False,
    )
    add(path, r"a^2", after_anchor="First paragraph.")
    add(path, r"b^2", after_index=0)
    add(path, r"c^2", at_end=True)
    out = listed(path)
    assert out["equation_count"] == 3
    # Each display equation sits in its own paragraph holding only math.
    root = doc_xml(path)
    for opara in root.iter(f"{{{M}}}oMathPara"):
        p = opara.getparent()
        assert p.tag == qn("w:p")
        others = [
            c
            for c in p
            if c is not opara and etree.QName(c).localname != "pPr"
        ]
        assert others == []


def test_display_requires_exactly_one_position(tmp_path):
    path = new_doc(tmp_path)
    pkg = DocxPackage(path)
    with pytest.raises(WordMcpError):
        eq.add_equation(pkg, r"a^2")  # no position at all
    with pytest.raises(WordMcpError):
        eq.add_equation(pkg, r"a^2", at_end=True, after_index=0)
    with pytest.raises(WordMcpError):
        eq.add_equation(pkg, r"a^2", at_end=True, anchor_text="x")


def test_inline_insert_preserves_text(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path,
        [{"text": "The identity holds for all x in the domain."}],
        at_end=True,
        backup=False,
    )
    before = [p["text"] for p in srv.get_text(path, live="off")]
    r = add(
        path, r"e^{i\pi} + 1 = 0", display=False, anchor_text="The identity"
    )
    assert r["display"] is False
    # Non-math text is unchanged (math is invisible to the text layer).
    after = [p["text"] for p in srv.get_text(path, live="off")]
    assert after == before
    # The oMath is a direct child of the anchor's w:p, sibling of runs.
    root = doc_xml(path)
    omaths = list(root.iter(f"{{{M}}}oMath"))
    assert len(omaths) == 1
    assert omaths[0].getparent().tag == qn("w:p")
    out = listed(path)
    assert out["equation_count"] == 1
    assert out["equations"][0]["display"] is False


def test_inline_requires_anchor_and_refuses_positions(tmp_path):
    path = new_doc(tmp_path)
    pkg = DocxPackage(path)
    with pytest.raises(WordMcpError):
        eq.add_equation(pkg, r"a^2", display=False)
    with pytest.raises(WordMcpError):
        eq.add_equation(
            pkg, r"a^2", display=False, anchor_text="x", at_end=True
        )
    with pytest.raises(TargetNotFound):
        eq.add_equation(
            pkg, r"a^2", display=False, anchor_text="no such anchor text"
        )


# --------------------------------------------------- gate: read-side contract


def test_equation_invisible_to_get_text_visible_to_list(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Before the math."}], at_end=True, backup=False
    )
    add(path, r"\frac{dy}{dx} = 3x^2", at_end=True)
    all_text = " ".join(p["text"] for p in srv.get_text(path, live="off"))
    assert "dy" not in all_text and "dx" not in all_text
    assert srv.find_text(path, "dy", live="off") == []
    out = listed(path)
    assert out["equation_count"] == 1
    assert "dy" in out["equations"][0]["text"]


def test_list_reports_location_and_note_part(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Anchor paragraph here."}], at_end=True, backup=False
    )
    add(path, r"a^2", after_anchor="Anchor paragraph here.")
    entry = listed(path)["equations"][0]
    assert entry["display"] is True
    assert isinstance(entry["paragraph_index"], int)


# ------------------------------------------------------------- gate: delete


def test_delete_display_removes_paragraph(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Keep me."}], at_end=True, backup=False
    )
    n_before = len(srv.get_text(path, live="off"))
    add(path, r"\frac{a}{b}", at_end=True)
    assert len(srv.get_text(path, live="off")) == n_before + 1

    pkg = DocxPackage(path)
    r = eq.delete_equation(pkg, 0)
    pkg.save(do_backup=False)
    assert r["deleted"] and r["removed_paragraph"]
    assert listed(path)["equation_count"] == 0
    assert len(srv.get_text(path, live="off")) == n_before
    texts = [p["text"] for p in srv.get_text(path, live="off")]
    assert "Keep me." in texts


def test_delete_inline_keeps_paragraph_text(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Sentence with math after this."}],
        at_end=True, backup=False,
    )
    add(path, r"x^2", display=False, anchor_text="math")
    before = [p["text"] for p in srv.get_text(path, live="off")]

    pkg = DocxPackage(path)
    r = eq.delete_equation(pkg, 0)
    pkg.save(do_backup=False)
    assert r["deleted"] and not r["removed_paragraph"]
    assert listed(path)["equation_count"] == 0
    assert [p["text"] for p in srv.get_text(path, live="off")] == before


def test_delete_bad_index_refused(tmp_path):
    path = new_doc(tmp_path)
    pkg = DocxPackage(path)
    with pytest.raises(TargetNotFound):
        eq.delete_equation(pkg, 0)


def test_add_delete_roundtrip_restores_structure(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Stable content."}], at_end=True, backup=False
    )
    snapshot = [p["text"] for p in srv.get_text(path, live="off")]
    add(path, r"\int_0^\infty e^{-x^2} dx", at_end=True)
    add(path, r"y = mx + b", display=False, anchor_text="Stable")
    assert listed(path)["equation_count"] == 2
    for _ in range(2):
        pkg = DocxPackage(path)
        eq.delete_equation(pkg, 0)
        pkg.save(do_backup=False)
    assert listed(path)["equation_count"] == 0
    assert [p["text"] for p in srv.get_text(path, live="off")] == snapshot


# ---------------------------------------------------------- gate: atomicity


def test_conversion_failure_is_atomic(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Untouched content."}], at_end=True, backup=False
    )
    original = Path(path).read_bytes()
    pkg = DocxPackage(path)
    with pytest.raises(eq.EquationConversionError) as exc:
        eq.add_equation(pkg, r"\frac{a", at_end=True)
    # The error carries the offending LaTeX for the caller.
    assert r"\frac{a" in str(exc.value)
    # Nothing was saved; the file is byte-identical.
    assert Path(path).read_bytes() == original
    # And the in-memory package was never marked dirty.
    assert not pkg._dirty


def test_empty_latex_refused(tmp_path):
    path = new_doc(tmp_path)
    pkg = DocxPackage(path)
    with pytest.raises(WordMcpError):
        eq.add_equation(pkg, "   ", at_end=True)


# ------------------------------------------------------------------- live


def _word_available():
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    import winreg

    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
        return True
    except OSError:
        return False


live_mark = pytest.mark.live
needs_word = pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)


@live_mark
@needs_word
def test_word_roundtrip_three_equations(tmp_path):
    """Word validation: a doc with 3 equations opens clean and Word itself
    counts OMaths.Count == 3 (invisible DispatchEx instance, quit cleanly)."""
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "The identity appears mid-sentence."}],
        at_end=True, backup=False,
    )
    add(path, r"\frac{a+b}{c}", at_end=True)
    add(path, r"\sum_{i=1}^{n} i^2 = \frac{n(n+1)(2n+1)}{6}", at_end=True)
    add(path, r"e^{i\pi} + 1 = 0", display=False, anchor_text="identity")

    v = srv.com_validate_opens_clean(path)
    assert v["opens_clean"], v

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(Path(path).resolve()), ReadOnly=True)
        try:
            assert doc.OMaths.Count == 3
        finally:
            doc.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
