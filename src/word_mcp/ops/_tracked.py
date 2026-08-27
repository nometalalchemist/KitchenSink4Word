"""Helpers for WRITING edits as tracked changes (w:ins / w:del).

Rules (research topic 6): revision wrappers hold complete w:r elements, never
nest inside a run; deleted text uses w:delText; every revision element carries
a unique w:id plus author and date. The accept/reject engine in revisions.py
is the round-trip proof for everything produced here.
"""

from __future__ import annotations

import datetime

from lxml import etree

from ..core.errors import UnsupportedStructure
from ..core.package import qn
from . import _runmap


def now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")


def next_rev_id(root: etree._Element) -> int:
    ids = [
        int(el.get(qn("w:id"), "0"))
        for tag in ("w:ins", "w:del", "w:moveFrom", "w:moveTo")
        for el in root.iter(qn(tag))
    ]
    return max(ids, default=0) + 1


def _same_parent(runs: list[etree._Element]) -> etree._Element:
    parents = {id(r.getparent()): r.getparent() for r in runs}
    if len(parents) != 1:
        raise UnsupportedStructure(
            "the text range crosses a hyperlink or revision boundary; "
            "tracked editing cannot wrap it — narrow the range"
        )
    return next(iter(parents.values()))


def wrap_runs_deleted(
    runs: list[etree._Element], *, author: str, rev_id: int, date: str
) -> etree._Element:
    """Wrap whole runs in one w:del, converting their text to w:delText."""
    parent = _same_parent(runs)
    del_el = etree.Element(qn("w:del"))
    del_el.set(qn("w:id"), str(rev_id))
    del_el.set(qn("w:author"), author)
    del_el.set(qn("w:date"), date)
    runs[0].addprevious(del_el)
    for r in runs:
        del_el.append(r)  # moves the element
        for t in r.findall(qn("w:t")):
            t.tag = qn("w:delText")
        for it in r.findall(qn("w:instrText")):
            it.tag = qn("w:delInstrText")
    return del_el


def make_ins(
    *,
    text: str,
    author: str,
    rev_id: int,
    date: str,
    rpr_source: etree._Element | None = None,
) -> etree._Element:
    """w:ins containing one run with `text`, formatting cloned from rpr_source."""
    import copy

    ins = etree.Element(qn("w:ins"))
    ins.set(qn("w:id"), str(rev_id))
    ins.set(qn("w:author"), author)
    ins.set(qn("w:date"), date)
    run = etree.SubElement(ins, qn("w:r"))
    if rpr_source is not None:
        rpr = rpr_source.find(qn("w:rPr"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
    for i, line in enumerate(text.split("\n")):
        if i:
            etree.SubElement(run, qn("w:br"))
        if line:
            t = etree.SubElement(run, qn("w:t"))
            t.text = line
            _runmap._preserve_space(t)
    return ins


def mark_paragraph_mark(
    p: etree._Element, kind: str, *, author: str, rev_id: int, date: str
) -> None:
    """Mark the paragraph MARK inserted/deleted (w:pPr/w:rPr/w:ins|w:del)."""
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(ppr, qn("w:rPr"))
    marker = etree.SubElement(rpr, qn(f"w:{kind}"))
    marker.set(qn("w:id"), str(rev_id))
    marker.set(qn("w:author"), author)
    marker.set(qn("w:date"), date)


def tracked_replace_range(
    p: etree._Element,
    start: int,
    end: int,
    replacement: str,
    *,
    author: str,
    rev_id: int,
    date: str,
) -> None:
    """Replace [start, end) as a tracked change: old runs wrapped in w:del,
    replacement inserted as w:ins after them."""
    runs = _runmap.split_for_range(p, start, end)
    del_el = wrap_runs_deleted(runs, author=author, rev_id=rev_id, date=date)
    if replacement:
        ins = make_ins(
            text=replacement,
            author=author,
            rev_id=rev_id + 1,
            date=date,
            rpr_source=runs[0],
        )
        del_el.addnext(ins)


def wrap_paragraph_content_deleted(
    p: etree._Element, *, author: str, rev_id: int, date: str
) -> int:
    """Wrap ALL of a paragraph's runs in w:del (grouped by parent so hyperlink
    contents wrap inside their wrapper). Returns revision elements created."""
    created = 0
    groups: list[list[etree._Element]] = []
    current: list[etree._Element] = []
    current_parent = None
    for r in list(p.iter(qn("w:r"))):
        parent = r.getparent()
        if etree.QName(parent).localname in ("del", "moveFrom"):
            continue  # already deleted
        if parent is not current_parent and current:
            groups.append(current)
            current = []
        current_parent = parent
        current.append(r)
    if current:
        groups.append(current)
    for group in groups:
        wrap_runs_deleted(
            group, author=author, rev_id=rev_id + created, date=date
        )
        created += 1
    return created


def wrap_paragraph_content_inserted(
    p: etree._Element, *, author: str, rev_id: int, date: str
) -> None:
    """Wrap all of a (new) paragraph's runs in one w:ins."""
    runs = [r for r in p.findall(qn("w:r"))]
    if not runs:
        return
    ins = etree.Element(qn("w:ins"))
    ins.set(qn("w:id"), str(rev_id))
    ins.set(qn("w:author"), author)
    ins.set(qn("w:date"), date)
    runs[0].addprevious(ins)
    for r in runs:
        ins.append(r)
