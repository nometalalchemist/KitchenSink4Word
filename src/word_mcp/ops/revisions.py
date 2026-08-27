"""Accept/reject tracked changes, whole-document or per-author.

Matrix per Eric White's canonical algorithm (see research/20260827, topic 6).
Visible-text invariants used by the tests:
- accept all -> visible text unchanged (insertions shown, deletions hidden).
- reject all -> original text (deletions restored, insertions dropped).
"""

from __future__ import annotations

import copy

from lxml import etree

from ..core.package import DocxPackage, qn

_PARTS = ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")

_PRCHANGE = {
    "rPrChange",
    "pPrChange",
    "tblPrChange",
    "tblGridChange",
    "tblPrExChange",
    "trPrChange",
    "tcPrChange",
    "sectPrChange",
    "numberingChange",
}


def _author_matches(el: etree._Element, author: str | None) -> bool:
    return author is None or el.get(qn("w:author"), "") == author


def _unwrap(el: etree._Element) -> None:
    parent = el.getparent()
    for child in list(el):
        el.addprevious(child)
    parent.remove(el)


def _restore_deleted(el: etree._Element) -> None:
    """delText -> t, delInstrText -> instrText, rsidDel -> rsidR."""
    for dt in list(el.iter(qn("w:delText"))):
        dt.tag = qn("w:t")
    for di in list(el.iter(qn("w:delInstrText"))):
        di.tag = qn("w:instrText")
    for r in el.iter(qn("w:r")):
        rsid = r.attrib.pop(qn("w:rsidDel"), None)
        if rsid is not None:
            r.set(qn("w:rsidR"), rsid)


def _is_para_mark_revision(el: etree._Element) -> bool:
    """Empty w:ins/w:del inside w:pPr/w:rPr revises the paragraph mark."""
    parent = el.getparent()
    if parent is None or etree.QName(parent).localname != "rPr":
        return False
    gp = parent.getparent()
    return gp is not None and etree.QName(gp).localname == "pPr"


def _containing_paragraph(el: etree._Element) -> etree._Element | None:
    node = el
    while node is not None and etree.QName(node).localname != "p":
        node = node.getparent()
    return node


def _merge_forward(p: etree._Element, *, take_following_ppr: bool) -> bool:
    """Merge the FOLLOWING sibling paragraph into `p`. Returns False when no
    following paragraph exists."""
    nxt = p.getnext()
    while nxt is not None and etree.QName(nxt).localname != "p":
        nxt = nxt.getnext()
    if nxt is None:
        return False
    if take_following_ppr:
        own = p.find(qn("w:pPr"))
        if own is not None:
            p.remove(own)
        their = nxt.find(qn("w:pPr"))
        if their is not None:
            p.insert(0, copy.deepcopy(their))
    for child in list(nxt):
        if etree.QName(child).localname != "pPr":
            p.append(child)
    nxt.getparent().remove(nxt)
    return True


def _process_part(
    root: etree._Element, *, accept: bool, author: str | None
) -> int:
    """One pass over a story part; returns the number of revisions resolved."""
    count = 0

    # --- paragraph-mark revisions (before run-level, they change structure)
    for tag in ("w:ins", "w:del"):
        for el in list(root.iter(qn(tag))):
            if not _is_para_mark_revision(el) or not _author_matches(el, author):
                continue
            p = _containing_paragraph(el)
            local = etree.QName(el).localname
            rpr = el.getparent()
            if local == "ins":
                if accept:
                    rpr.remove(el)  # split stands
                else:
                    rpr.remove(el)
                    _merge_forward(p, take_following_ppr=False)  # un-split
            else:  # del
                if accept:
                    rpr.remove(el)
                    if not _merge_forward(p, take_following_ppr=True):
                        pass  # last paragraph: mark removal is enough
                else:
                    rpr.remove(el)
            count += 1

    # --- table row revisions
    for tr in list(root.iter(qn("w:tr"))):
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            continue
        ins = trpr.find(qn("w:ins"))
        if ins is not None and _author_matches(ins, author):
            if accept:
                trpr.remove(ins)
            else:
                tr.getparent().remove(tr)
            count += 1
            continue
        dele = trpr.find(qn("w:del"))
        if dele is not None and _author_matches(dele, author):
            if accept:
                tr.getparent().remove(tr)
            else:
                trpr.remove(dele)
                _restore_deleted(tr)
                for mark_del in list(tr.iter(qn("w:del"))):
                    if _is_para_mark_revision(mark_del):
                        mark_del.getparent().remove(mark_del)
            count += 1

    # --- run-content ins/del/moveFrom/moveTo (deepest first so nesting works)
    revs = [
        el
        for el in root.iter(qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"))
        if not _is_para_mark_revision(el)
        and etree.QName(el.getparent()).localname not in ("trPr", "rPr")
    ]
    for el in sorted(revs, key=lambda e: -len(list(e.iterancestors()))):
        if not _author_matches(el, author):
            continue
        local = etree.QName(el).localname
        if local in ("ins", "moveTo"):
            if accept:
                _unwrap(el)
            else:
                el.getparent().remove(el)
        else:  # del / moveFrom
            if accept:
                el.getparent().remove(el)
            else:
                _restore_deleted(el)
                _unwrap(el)
        count += 1

    # --- move range markers (dropped whenever we touch moves; author-scoped)
    for tag in (
        "w:moveFromRangeStart",
        "w:moveFromRangeEnd",
        "w:moveToRangeStart",
        "w:moveToRangeEnd",
    ):
        for el in list(root.iter(qn(tag))):
            if _author_matches(el, author) or el.get(qn("w:author")) is None:
                el.getparent().remove(el)

    # --- property changes
    for el in list(root.iter()):
        local = etree.QName(el).localname
        if local not in _PRCHANGE or not _author_matches(el, author):
            continue
        parent = el.getparent()
        if accept or local == "numberingChange":
            parent.remove(el)
        else:
            stored = next(iter(el), None)
            parent.remove(el)
            if stored is not None:
                for child in list(parent):
                    parent.remove(child)
                for child in list(stored):
                    parent.append(child)
        count += 1

    # --- cell revisions (Compare output)
    for tc in list(root.iter(qn("w:tc"))):
        tcpr = tc.find(qn("w:tcPr"))
        if tcpr is None:
            continue
        ci = tcpr.find(qn("w:cellIns"))
        if ci is not None and _author_matches(ci, author):
            if accept:
                tcpr.remove(ci)
            else:
                tc.getparent().remove(tc)
            count += 1
            continue
        cd = tcpr.find(qn("w:cellDel"))
        if cd is not None and _author_matches(cd, author):
            if accept:
                prev = tc.getprevious()
                tc.getparent().remove(tc)
                if prev is not None and etree.QName(prev).localname == "tc":
                    ptcpr = prev.find(qn("w:tcPr"))
                    if ptcpr is None:
                        ptcpr = etree.Element(qn("w:tcPr"))
                        prev.insert(0, ptcpr)
                    gs = ptcpr.find(qn("w:gridSpan"))
                    cur = int(gs.get(qn("w:val"), "1")) if gs is not None else 1
                    if gs is None:
                        gs = etree.SubElement(ptcpr, qn("w:gridSpan"))
                    gs.set(qn("w:val"), str(cur + 1))
            else:
                tcpr.remove(cd)
            count += 1
        cm = tcpr.find(qn("w:cellMerge"))
        if cm is not None and _author_matches(cm, author):
            if accept:
                val = cm.get(qn("w:vMerge"), "")
                vm = tcpr.find(qn("w:vMerge"))
                if vm is None:
                    vm = etree.SubElement(tcpr, qn("w:vMerge"))
                vm.set(
                    qn("w:val"), "restart" if val == "rest" else "continue"
                )
            tcpr.remove(cm)
            count += 1

    return count


def _apply(pkg: DocxPackage, *, accept: bool, author: str | None) -> dict:
    total = 0
    for part in _PARTS:
        if not pkg.has_part(part):
            continue
        root = pkg.root(part)
        part_count = 0
        # Iterate to stability: nested revisions can expose new ones.
        for _ in range(20):
            n = _process_part(root, accept=accept, author=author)
            part_count += n
            if n == 0:
                break
        if part_count:
            pkg.mark_dirty(part)
        total += part_count
    result = {
        "action": "accepted" if accept else "rejected",
        "author": author or "ALL",
        "revisions_resolved": total,
    }
    if total and accept:
        # Accepted deletions may have removed footnote/endnote references.
        from .notes import purge_orphans

        purged = purge_orphans(pkg)["purged"]
        if purged:
            result["note_definitions_purged"] = purged
    return result


def accept_revisions(pkg: DocxPackage, *, author: str | None = None) -> dict:
    return _apply(pkg, accept=True, author=author)


def reject_revisions(pkg: DocxPackage, *, author: str | None = None) -> dict:
    return _apply(pkg, accept=False, author=author)
