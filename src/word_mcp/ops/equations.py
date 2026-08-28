"""OMML equations: LaTeX -> Word math (m:oMath / m:oMathPara).

Conversion pipeline (decided by the v1.5 research, both deps MIT, pure
Python): latex2mathml (LaTeX -> MathML) then mathml2omml (MathML -> OMML).
The produced OMML is parsed and namespace-checked BEFORE anything in the
document is touched, so a conversion failure never modifies the file.

READ-SIDE REALITY (by design, documented here once): math runs are m:r/m:t
in the math namespace, which the run map and all plain-text extraction index
deliberately ignore (w:r/w:t only). Equations are therefore INVISIBLE to
get_text / find_text / search_and_replace — which also means search and
replace can never corrupt an equation. list_equations() is the read story
for math content.
"""

from __future__ import annotations

import re

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap
from .read import body_items

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_OMATH = f"{{{M_NS}}}oMath"
_OMATH_PARA = f"{{{M_NS}}}oMathPara"
_MT = f"{{{M_NS}}}t"

# Parts scanned for equations, in listing order (stable delete indices).
_EQ_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")


class EquationConversionError(WordMcpError):
    """LaTeX did not convert to valid OMML; the document was not modified."""


def _rewrite_aligned(latex: str) -> str:
    # Upstream latex2mathml bug: the `aligned` environment emits malformed
    # XML (unescaped entity). `align*` is semantically equivalent for our
    # purposes and converts cleanly, so rewrite before conversion.
    return latex.replace(r"\begin{aligned}", r"\begin{align*}").replace(
        r"\end{aligned}", r"\end{align*}"
    )


def _latex_to_omath(latex: str, *, display: bool) -> etree._Element:
    """Convert LaTeX to a parsed m:oMath element, or raise (nothing mutated).

    mathml2omml emits a bare m:oMath string that uses the m: prefix without
    declaring the namespace, so we parse it inside a wrapper that declares
    xmlns:m — that single step both validates the XML and binds the prefix.
    The returned element carries the namespace in its own nsmap, so lxml
    declares it locally when the element lands in a document whose root
    lacks xmlns:m.
    """
    if not latex or not latex.strip():
        raise WordMcpError("latex must be non-empty")
    # File/system macros are meaningless inside a Word equation and were
    # silently passing through as literal text (v1.5 adversarial F6) —
    # refuse them by name. (No execution risk either way; this is honesty.)
    _banned = re.search(
        r"\\(input|include|write\d*|openout|openin|read|immediate|usepackage)\b",
        latex,
    )
    if _banned:
        raise WordMcpError(
            f"\\{_banned.group(1)} is a file/preamble macro, not equation "
            "content; remove it — only math-mode LaTeX converts"
        )
    fixed = _rewrite_aligned(latex)
    try:
        import latex2mathml.converter as _l2m
        import mathml2omml as _m2o

        mathml = _l2m.convert(fixed, display="block" if display else "inline")
        omml = _m2o.convert(mathml)
        wrapper = etree.fromstring(
            f'<wrap xmlns:m="{M_NS}">{omml}</wrap>'.encode("utf-8")
        )
    except Exception as exc:  # converters raise assorted types; all mean "bad input"
        raise EquationConversionError(
            f"could not convert LaTeX to Word math: {exc} "
            f"(input was: {latex!r}); the document was not changed"
        ) from exc
    children = list(wrapper)
    if len(children) != 1 or children[0].tag != _OMATH:
        raise EquationConversionError(
            f"converter produced unexpected OMML root for {latex!r}; "
            "the document was not changed"
        )
    omath = children[0]
    wrapper.remove(omath)
    return omath


def _find_anchor_span(pkg: DocxPackage, anchor_text: str, occurrence: int):
    """(paragraph, start, end) of the Nth occurrence of anchor_text."""
    import re

    seen = 0
    for p in pkg.root().iter(qn("w:p")):
        text, _ = _runmap.build_map(p)
        for m in re.finditer(re.escape(anchor_text), text):
            seen += 1
            if seen == occurrence:
                return p, m.start(), m.end()
    raise TargetNotFound(
        f"anchor text not found: {anchor_text!r}"
        + (f" (occurrence {occurrence}, saw {seen})" if seen else "")
    )


def add_equation(
    pkg: DocxPackage,
    latex: str,
    *,
    display: bool = True,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    anchor_text: str | None = None,
    occurrence: int = 1,
) -> dict:
    """Insert a LaTeX equation as native Word math (editable in Word's
    equation editor, not an image).

    display=True: a block equation in its own paragraph; give exactly one of
    after_index / after_anchor / at_end to position it. display=False: an
    inline equation placed immediately after `anchor_text` (occurrence picks
    among repeats) inside that paragraph, surrounding text untouched.

    `\\begin{aligned}` is rewritten to `align*` before conversion (upstream
    converter bug). Any conversion failure raises with the LaTeX and the
    converter's message, and the document is not modified.

    NOTE: equations do not appear in get_text/find_text output (math runs
    are outside the text index by design); read them via list_equations.
    """
    positioners = sum(
        x is not None for x in (after_index, after_anchor)
    ) + bool(at_end)
    if display:
        if anchor_text is not None:
            raise WordMcpError(
                "anchor_text positions INLINE equations; for a display "
                "equation use after_index, after_anchor, or at_end"
            )
        if positioners != 1:
            raise WordMcpError(
                "specify exactly one of after_index, after_anchor, at_end"
            )
    else:
        if anchor_text is None:
            raise WordMcpError(
                "inline equations (display=False) require anchor_text: the "
                "equation is inserted right after that text"
            )
        if positioners:
            raise WordMcpError(
                "after_index/after_anchor/at_end position display equations; "
                "an inline equation is positioned by anchor_text alone"
            )

    # Convert FIRST: any failure raises before the tree is touched.
    omath = _latex_to_omath(latex, display=display)

    if display:
        para = etree.Element(qn("w:p"))
        opara = etree.SubElement(para, _OMATH_PARA)
        opara.append(omath)
        body = pkg.body()
        if at_end:
            sectpr = body.find(qn("w:sectPr"))
            if sectpr is not None:
                sectpr.addprevious(para)
            else:
                body.append(para)
        elif after_anchor is not None:
            ref, _, _ = _find_anchor_span(pkg, after_anchor, occurrence)
            ref.addnext(para)
        else:
            ref = None
            for kind, idx, el in body_items(pkg):
                if kind == "paragraph" and idx == after_index:
                    ref = el
                    break
            if ref is None:
                raise TargetNotFound(
                    f"no body paragraph with index {after_index}"
                )
            ref.addnext(para)
        location = _locate(pkg, "word/document.xml", para)
    else:
        p, _, end = _find_anchor_span(pkg, anchor_text, occurrence)
        # Split runs so the anchor's last character ends a run, then drop the
        # oMath in as its sibling — same placement mechanics as field anchors.
        covered = _runmap.split_for_range(p, end - 1, end)
        covered[-1].addnext(omath)
        location = _locate(pkg, "word/document.xml", p)

    pkg.mark_dirty()
    return {
        "equation_added": True,
        "display": display,
        "latex": latex,
        "text": _approx_text(omath),
        **location,
    }


def _approx_text(el: etree._Element) -> str:
    """Concatenated m:t content — a linear plain-text approximation only
    (structure like fraction bars and matrix layout is not represented)."""
    return "".join(t.text or "" for t in el.iter(_MT))


def _locate(pkg: DocxPackage, part: str, el: etree._Element) -> dict:
    """Location descriptor for the paragraph-or-element `el` in `part`."""
    if part != "word/document.xml":
        return {"part": part}
    # Containing w:p (el may itself be the w:p).
    p = el if el.tag == qn("w:p") else el.getparent()
    while p is not None and p.tag != qn("w:p"):
        p = p.getparent()
    if p is not None:
        for kind, idx, item in body_items(pkg):
            if kind == "paragraph" and item is p:
                return {"paragraph_index": idx}
    # Not a body-level paragraph: inside a table, SDT, or header-anchored shape.
    anc = el
    while anc is not None:
        if anc.tag == qn("w:tbl"):
            return {"in_table": True}
        anc = anc.getparent()
    return {"paragraph_index": None}


def _enumerate(pkg: DocxPackage) -> list[tuple[str, etree._Element, bool]]:
    """All equations in stable order: (part, element, is_display).

    One m:oMathPara = ONE display equation (a multi-line align block is a
    single unit); a bare m:oMath outside any oMathPara = one inline equation.
    """
    out: list[tuple[str, etree._Element, bool]] = []
    for part in _EQ_PARTS:
        if not pkg.has_part(part):
            continue
        for el in pkg.root(part).iter():
            if el.tag == _OMATH_PARA:
                out.append((part, el, True))
            elif el.tag == _OMATH:
                anc = el.getparent()
                inside_para = False
                while anc is not None:
                    if anc.tag == _OMATH_PARA:
                        inside_para = True
                        break
                    anc = anc.getparent()
                if not inside_para:
                    out.append((part, el, False))
    return out


def list_equations(pkg: DocxPackage) -> dict:
    """Every equation in the document (body, tables, footnotes, endnotes).

    THIS IS THE READ PATH FOR MATH: equations never appear in get_text or
    find_text output because math runs (m:r/m:t) are outside the plain-text
    index by design. Each entry: index (the handle delete_equation takes),
    display vs inline, location (body paragraph_index, in_table, or the note
    part), and `text`, a linear concatenation of the math text — an
    approximation only, structure (fraction bars, matrix layout) is lost.
    Also returns equation_count for document-info style summaries.
    """
    entries = []
    for i, (part, el, is_display) in enumerate(_enumerate(pkg)):
        entry = {
            "index": i,
            "display": is_display,
            "text": _approx_text(el),
            **_locate(pkg, part, el),
        }
        entries.append(entry)
    return {"equations": entries, "equation_count": len(entries)}


def delete_equation(pkg: DocxPackage, index: int) -> dict:
    """Delete the equation with the given list_equations index.

    Display equations: the whole paragraph is removed when it holds nothing
    but the equation (the usual case); otherwise only the m:oMathPara is
    removed and the paragraph's other content stays. Inline equations: only
    the m:oMath is removed, surrounding text untouched.
    """
    eqs = _enumerate(pkg)
    if not 0 <= index < len(eqs):
        raise TargetNotFound(
            f"no equation with index {index}; the document has {len(eqs)} "
            "(see list_equations)"
        )
    part, el, is_display = eqs[index]
    text = _approx_text(el)
    location = _locate(pkg, part, el)
    removed_paragraph = False
    if is_display:
        p = el.getparent()
        while p is not None and p.tag != qn("w:p"):
            p = p.getparent()
        if p is not None and _holds_only(p, el):
            p.getparent().remove(p)
            removed_paragraph = True
        else:
            el.getparent().remove(el)
    else:
        el.getparent().remove(el)
    pkg.mark_dirty(part)
    return {
        "deleted": True,
        "display": is_display,
        "text": text,
        "removed_paragraph": removed_paragraph,
        **location,
    }


_IGNORABLE_SIBLINGS = {"pPr", "proofErr", "bookmarkStart", "bookmarkEnd"}


def _holds_only(p: etree._Element, el: etree._Element) -> bool:
    """True if paragraph `p` has no content besides `el` and inert markers."""
    for child in p:
        if child is el:
            continue
        if etree.QName(child).localname not in _IGNORABLE_SIBLINGS:
            return False
    return True
