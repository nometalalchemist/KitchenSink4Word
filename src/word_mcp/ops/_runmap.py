"""Character-offset map over a paragraph's runs.

Word fragments logically continuous text into many w:r elements (revision-save
ids, spell-check state, formatting boundaries). Any edit addressed by character
position must be resolved through this map, never by per-run string operations:
a match can start in one run and end three runs later.

The map indexes w:t text only. Tabs/breaks contribute to *visible* text but are
atomic elements; edits may not start or end inside one (they occupy a single
character slot). Text inside w:del (deleted) is invisible and excluded; text
inside w:ins is visible and included — rewriting it keeps its insertion mark.
"""

from __future__ import annotations

import copy

from lxml import etree

from ..core.errors import UnsupportedStructure
from ..core.package import qn


class Segment:
    """One text-contributing child of a run: a w:t or an atomic (tab/br/etc.)."""

    __slots__ = ("el", "run", "start", "end", "atomic")

    def __init__(self, el, run, start, end, atomic):
        self.el = el
        self.run = run
        self.start = start
        self.end = end
        self.atomic = atomic


def _in_deleted(el) -> bool:
    parent = el.getparent()
    while parent is not None:
        if etree.QName(parent).localname in ("del", "moveFrom"):
            return True
        parent = parent.getparent()
    return False


def _in_textbox(el, stop) -> bool:
    """Inside w:txbxContent (either the mc:Choice wps or legacy v:textbox
    copy)? Text-box content is a SEPARATE story: counting it here smears it
    into the host paragraph DOUBLED (Word stores both compatibility copies).
    Dedicated access lives in ops/textboxes.py."""
    parent = el.getparent()
    while parent is not None and parent is not stop:
        if etree.QName(parent).localname == "txbxContent":
            return True
        parent = parent.getparent()
    return False


def build_map(p: etree._Element) -> tuple[str, list[Segment]]:
    """Visible text of the paragraph plus segments mapping offsets to
    elements. Text inside text boxes is EXCLUDED (see _in_textbox)."""
    segments: list[Segment] = []
    parts: list[str] = []
    pos = 0
    for r in p.iter(qn("w:r")):
        if _in_deleted(r):
            continue
        if _in_textbox(r, p):
            continue
        for child in r:
            tag = etree.QName(child).localname
            text = None
            atomic = False
            if tag == "t":
                text = child.text or ""
            elif tag == "tab":
                text, atomic = "\t", True
            elif tag == "br":
                if child.get(qn("w:type")) in ("page", "column"):
                    continue
                text, atomic = "\n", True
            elif tag == "cr":
                text, atomic = "\n", True
            elif tag == "noBreakHyphen":
                text, atomic = "-", True
            if text is None or text == "":
                continue
            segments.append(Segment(child, r, pos, pos + len(text), atomic))
            parts.append(text)
            pos += len(text)
    return "".join(parts), segments


def replace_range(
    p: etree._Element,
    segments: list[Segment],
    start: int,
    end: int,
    replacement: str,
) -> None:
    """Replace visible characters [start, end) with `replacement`.

    The replacement text lands in the first affected w:t (inheriting that run's
    formatting); remaining covered text is removed. Atomic elements fully inside
    the range are removed; a range boundary inside an atomic element is
    impossible (they are single characters), but a range may not *start or end
    mid-element* in any way that splits an atomic — single chars never split.
    """
    affected = [s for s in segments if s.start < end and s.end > start]
    if not affected:
        raise ValueError("range matches no text segments")

    first_t: etree._Element | None = None
    for seg in affected:
        if seg.atomic:
            # Atomic fully covered -> remove; partially covered cannot happen.
            _remove_child(seg)
            continue
        text = seg.el.text or ""
        lo = max(start - seg.start, 0)
        hi = min(end - seg.start, len(text))
        if first_t is None:
            seg.el.text = text[:lo] + replacement + text[hi:]
            _preserve_space(seg.el)
            first_t = seg.el
        else:
            remaining = text[:lo] + text[hi:]
            if remaining:
                seg.el.text = remaining
                _preserve_space(seg.el)
            else:
                _remove_child(seg)

    if first_t is None:
        # Every affected segment was atomic; put replacement in a new run
        # cloned from the first affected run's formatting.
        ref_run = affected[0].run
        new_run = _clone_run_shell(ref_run)
        t = etree.SubElement(new_run, qn("w:t"))
        t.text = replacement
        _preserve_space(t)
        ref_run.addprevious(new_run)


def _remove_child(seg: Segment) -> None:
    run = seg.run
    run.remove(seg.el)
    # A run left with no content children (only rPr) is dead weight.
    if not [c for c in run if etree.QName(c).localname != "rPr"]:
        parent = run.getparent()
        if parent is not None:
            parent.remove(run)


def _preserve_space(t: etree._Element) -> None:
    text = t.text or ""
    if text != text.strip() or text == "":
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:
        t.attrib.pop("{http://www.w3.org/XML/1998/namespace}space", None)


def _clone_run_shell(run: etree._Element) -> etree._Element:
    """New empty w:r carrying a copy of `run`'s rPr."""
    new_run = etree.Element(qn("w:r"))
    rpr = run.find(qn("w:rPr"))
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    return new_run


def split_for_range(
    p: etree._Element, start: int, end: int
) -> list[etree._Element]:
    """Split runs so that [start, end) is covered by whole runs; return them.

    Used by character-range formatting: after splitting, formatting can be
    applied to the returned runs without touching text outside the range.
    """
    _, segments = build_map(p)
    affected = [s for s in segments if s.start < end and s.end > start]
    if not affected:
        raise ValueError("range matches no text segments")

    covered_runs: list[etree._Element] = []
    for seg in affected:
        if seg.atomic:
            covered_runs.append(seg.run)
            continue
        run = seg.run
        # Metadata elements contribute no text and are safe to drop from a
        # run being split (lastRenderedPageBreak is a stale pagination hint
        # Word regenerates; proofErr never appears inside runs but guard it).
        _IGNORABLE = {"lastRenderedPageBreak", "proofErr"}
        for c in list(run):
            if etree.QName(c).localname in _IGNORABLE:
                run.remove(c)
        siblings = [c for c in run if etree.QName(c).localname != "rPr"]
        if len(siblings) > 1:
            kinds = sorted({etree.QName(c).localname for c in siblings})
            raise UnsupportedStructure(
                f"cannot split a run holding multiple content elements "
                f"({', '.join(kinds)}); narrow the range to avoid it"
            )
        text = seg.el.text or ""
        lo = max(start - seg.start, 0)
        hi = min(end - seg.start, len(text))
        pre, mid, post = text[:lo], text[lo:hi], text[hi:]
        if pre:
            pre_run = _clone_run_shell(run)
            t = etree.SubElement(pre_run, qn("w:t"))
            t.text = pre
            _preserve_space(t)
            run.addprevious(pre_run)
        if post:
            post_run = _clone_run_shell(run)
            t = etree.SubElement(post_run, qn("w:t"))
            t.text = post
            _preserve_space(t)
            run.addnext(post_run)
        seg.el.text = mid
        _preserve_space(seg.el)
        covered_runs.append(run)
    return covered_runs
