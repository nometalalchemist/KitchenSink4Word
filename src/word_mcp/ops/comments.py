"""Comment authoring: add, reply, resolve, delete.

Threading model (Word 2019+/365): comments.xml holds content; each comment's
last paragraph carries a w14:paraId; commentsExtended.xml maps paraId ->
{done, paraIdParent}. commentsIds.xml + people.xml are written for Word
parity (safe-engineering position; see research topic 5).
"""

from __future__ import annotations

import datetime
import random
import re

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import NSMAP, DocxPackage, qn
from . import _runmap

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_W14 = NSMAP["w14"]
_W15 = NSMAP["w15"]
_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"

_PARTS = {
    "word/comments.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml"
        ".comments+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    ),
    "word/commentsExtended.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml"
        ".commentsExtended+xml",
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    ),
    "word/commentsIds.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml"
        ".commentsIds+xml",
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
    ),
    "word/people.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml"
        ".people+xml",
        "http://schemas.microsoft.com/office/2011/relationships/people",
    ),
}

_ROOTS = {
    "word/comments.xml": ("w:comments", {"w": NSMAP["w"], "w14": _W14}),
    "word/commentsExtended.xml": (
        f"{{{_W15}}}commentsEx",
        {"w15": _W15},
    ),
    "word/commentsIds.xml": (
        f"{{{_W16CID}}}commentsIds",
        {"w16cid": _W16CID},
    ),
    "word/people.xml": (f"{{{_W15}}}people", {"w15": _W15}),
}


def _hex_id() -> str:
    return f"{random.randint(1, 0x7FFFFFFE):08X}"


def _ensure_part(pkg: DocxPackage, part: str) -> None:
    if pkg.has_part(part):
        return
    root_spec, nsmap = _ROOTS[part]
    tag = qn(root_spec) if root_spec.startswith("w:") else root_spec
    root = etree.Element(tag, nsmap=nsmap)
    pkg.set_raw_part(
        part,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )
    ct_root = pkg.root("[Content_Types].xml")
    content_type, rel_type = _PARTS[part]
    if not any(
        o.get("PartName") == "/" + part
        for o in ct_root.findall(f"{{{_CT_NS}}}Override")
    ):
        override = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        override.set("PartName", "/" + part)
        override.set("ContentType", content_type)
        pkg.mark_dirty("[Content_Types].xml")
    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    if not any(r.get("Type") == rel_type for r in rels_root):
        existing = {r.get("Id") for r in rels_root}
        i = 1
        while f"rId{i}" in existing:
            i += 1
        rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
        rel.set("Id", f"rId{i}")
        rel.set("Type", rel_type)
        rel.set("Target", part.split("/", 1)[1])
        pkg.mark_dirty(rels_part)


def _ensure_comment_style(pkg: DocxPackage) -> None:
    """Define the CommentReference character style if the document lacks it.
    _build_comment's reference run cites it; without the definition Word
    silently falls back to Normal (diagnose_document flagged this; field
    test, 2026-09-03)."""
    if not pkg.has_part("word/styles.xml"):
        return
    root = pkg.root("word/styles.xml")
    for s in root.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == "CommentReference":
            return
        name = s.find(qn("w:name"))
        if name is not None and name.get(qn("w:val")) == "annotation reference":
            return
    style = etree.SubElement(root, qn("w:style"))
    style.set(qn("w:type"), "character")
    style.set(qn("w:styleId"), "CommentReference")
    etree.SubElement(style, qn("w:name")).set(
        qn("w:val"), "annotation reference"
    )
    etree.SubElement(style, qn("w:uiPriority")).set(qn("w:val"), "99")
    etree.SubElement(style, qn("w:semiHidden"))
    etree.SubElement(style, qn("w:unhideWhenUsed"))
    rpr = etree.SubElement(style, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:sz")).set(qn("w:val"), "16")
    etree.SubElement(rpr, qn("w:szCs")).set(qn("w:val"), "16")
    pkg.mark_dirty("word/styles.xml")


def _next_comment_id(pkg: DocxPackage) -> int:
    if not pkg.has_part("word/comments.xml"):
        return 0
    ids = [
        int(c.get(qn("w:id"), "0"))
        for c in pkg.root("word/comments.xml").findall(qn("w:comment"))
    ]
    return max(ids, default=-1) + 1


def _build_comment(
    pkg: DocxPackage, cid: int, author: str, initials: str, text: str
) -> str:
    """Append the w:comment; returns the last paragraph's paraId."""
    root = pkg.root("word/comments.xml")
    c = etree.SubElement(root, qn("w:comment"))
    c.set(qn("w:id"), str(cid))
    c.set(qn("w:author"), author)
    c.set(qn("w:initials"), initials)
    c.set(qn("w:date"), datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"))
    para_id = _hex_id()
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = etree.SubElement(c, qn("w:p"))
        if i == len(lines) - 1:
            p.set(f"{{{_W14}}}paraId", para_id)
            p.set(f"{{{_W14}}}textId", _hex_id())
        if i == 0:
            ref_r = etree.SubElement(p, qn("w:r"))
            rpr = etree.SubElement(ref_r, qn("w:rPr"))
            etree.SubElement(rpr, qn("w:rStyle")).set(
                qn("w:val"), "CommentReference"
            )
            etree.SubElement(ref_r, qn("w:annotationRef"))
        if line:
            r = etree.SubElement(p, qn("w:r"))
            t = etree.SubElement(r, qn("w:t"))
            t.text = line
            _runmap._preserve_space(t)
    pkg.mark_dirty("word/comments.xml")
    return para_id


def _register_extended(
    pkg: DocxPackage, para_id: str, *, parent_para_id: str | None = None
) -> None:
    root = pkg.root("word/commentsExtended.xml")
    ex = etree.SubElement(root, f"{{{_W15}}}commentEx")
    ex.set(f"{{{_W15}}}paraId", para_id)
    if parent_para_id:
        ex.set(f"{{{_W15}}}paraIdParent", parent_para_id)
    ex.set(f"{{{_W15}}}done", "0")
    pkg.mark_dirty("word/commentsExtended.xml")

    ids_root = pkg.root("word/commentsIds.xml")
    cid_el = etree.SubElement(ids_root, f"{{{_W16CID}}}commentId")
    cid_el.set(f"{{{_W16CID}}}paraId", para_id)
    cid_el.set(f"{{{_W16CID}}}durableId", _hex_id())
    pkg.mark_dirty("word/commentsIds.xml")


def _register_person(pkg: DocxPackage, author: str) -> None:
    root = pkg.root("word/people.xml")
    if any(
        p.get(f"{{{_W15}}}author") == author
        for p in root.findall(f"{{{_W15}}}person")
    ):
        return
    person = etree.SubElement(root, f"{{{_W15}}}person")
    person.set(f"{{{_W15}}}author", author)
    info = etree.SubElement(person, f"{{{_W15}}}presenceInfo")
    info.set(f"{{{_W15}}}providerId", "None")
    info.set(f"{{{_W15}}}userId", author)
    pkg.mark_dirty("word/people.xml")


def _default_initials(author: str) -> str:
    parts = re.split(r"\s+", author.strip())
    return "".join(p[0].upper() for p in parts if p)[:3] or "C"


def add_comment(
    pkg: DocxPackage,
    *,
    anchor_text: str,
    text: str,
    author: str = "Claude",
    initials: str | None = None,
    occurrence: int = 1,
) -> dict:
    """Comment on `anchor_text` (the commented range in the body)."""
    for part in _PARTS:
        _ensure_part(pkg, part)
    _ensure_comment_style(pkg)
    seen = 0
    target = None
    for p in pkg.root().iter(qn("w:p")):
        body_text, _ = _runmap.build_map(p)
        for m in re.finditer(re.escape(anchor_text), body_text):
            seen += 1
            if seen == occurrence:
                target = (p, m.start(), m.end())
                break
        if target:
            break
    if target is None:
        raise TargetNotFound(
            f"anchor text not found: {anchor_text!r}"
            + (f" (occurrence {occurrence}, saw {seen})" if seen else "")
        )
    p, start, end = target
    runs = _runmap.split_for_range(p, start, end)

    cid = _next_comment_id(pkg)
    para_id = _build_comment(
        pkg, cid, author, initials or _default_initials(author), text
    )
    _register_extended(pkg, para_id)
    _register_person(pkg, author)

    crs = etree.Element(qn("w:commentRangeStart"))
    crs.set(qn("w:id"), str(cid))
    cre = etree.Element(qn("w:commentRangeEnd"))
    cre.set(qn("w:id"), str(cid))
    runs[0].addprevious(crs)
    runs[-1].addnext(cre)
    ref_r = etree.Element(qn("w:r"))
    ref = etree.SubElement(ref_r, qn("w:commentReference"))
    ref.set(qn("w:id"), str(cid))
    cre.addnext(ref_r)
    pkg.mark_dirty()
    return {"comment_id": str(cid), "author": author}


def reply_to_comment(
    pkg: DocxPackage,
    *,
    comment_id: str,
    text: str,
    author: str = "Claude",
    initials: str | None = None,
) -> dict:
    """Threaded reply to an existing comment."""
    parent = _find_comment(pkg, comment_id)
    parent_para_id = _last_para_id(parent)
    if parent_para_id is None:
        raise WordMcpError(
            "parent comment has no paraId; it predates threading — reply in Word"
        )
    for part in _PARTS:
        _ensure_part(pkg, part)
    _ensure_comment_style(pkg)
    cid = _next_comment_id(pkg)
    para_id = _build_comment(
        pkg, cid, author, initials or _default_initials(author), text
    )
    _register_extended(pkg, para_id, parent_para_id=parent_para_id)
    _register_person(pkg, author)

    # Anchor the reply beside the parent's reference mark.
    parent_ref = next(
        (
            r
            for r in pkg.root().iter(qn("w:commentReference"))
            if r.get(qn("w:id")) == comment_id
        ),
        None,
    )
    if parent_ref is None:
        raise TargetNotFound(f"comment {comment_id} has no body reference")
    parent_run = parent_ref.getparent()
    crs = etree.Element(qn("w:commentRangeStart"))
    crs.set(qn("w:id"), str(cid))
    cre = etree.Element(qn("w:commentRangeEnd"))
    cre.set(qn("w:id"), str(cid))
    parent_run.addprevious(crs)
    parent_run.addnext(cre)
    ref_r = etree.Element(qn("w:r"))
    ref = etree.SubElement(ref_r, qn("w:commentReference"))
    ref.set(qn("w:id"), str(cid))
    cre.addnext(ref_r)
    pkg.mark_dirty()
    return {"reply_id": str(cid), "parent_id": comment_id}


def _find_comment(pkg: DocxPackage, comment_id: str) -> etree._Element:
    if not pkg.has_part("word/comments.xml"):
        raise TargetNotFound("document has no comments")
    c = next(
        (
            c
            for c in pkg.root("word/comments.xml").findall(qn("w:comment"))
            if c.get(qn("w:id")) == str(comment_id)
        ),
        None,
    )
    if c is None:
        raise TargetNotFound(f"no comment with id {comment_id}")
    return c


def _last_para_id(comment: etree._Element) -> str | None:
    paras = comment.findall(qn("w:p"))
    return paras[-1].get(f"{{{_W14}}}paraId") if paras else None


def resolve_comment(pkg: DocxPackage, *, comment_id: str, done: bool = True) -> dict:
    """Mark a comment thread resolved (or reopen it)."""
    comment = _find_comment(pkg, comment_id)
    para_id = _last_para_id(comment)
    if para_id is None or not pkg.has_part("word/commentsExtended.xml"):
        raise WordMcpError(
            "comment lacks threading metadata; resolve it in Word instead"
        )
    root = pkg.root("word/commentsExtended.xml")
    ex = next(
        (
            e
            for e in root.findall(f"{{{_W15}}}commentEx")
            if e.get(f"{{{_W15}}}paraId") == para_id
        ),
        None,
    )
    if ex is None:
        ex = etree.SubElement(root, f"{{{_W15}}}commentEx")
        ex.set(f"{{{_W15}}}paraId", para_id)
    ex.set(f"{{{_W15}}}done", "1" if done else "0")
    pkg.mark_dirty("word/commentsExtended.xml")
    return {"comment_id": str(comment_id), "resolved": done}


def delete_comment(pkg: DocxPackage, *, comment_id: str) -> dict:
    """Delete a comment and its replies (cascade), plus all body markers and
    metadata entries."""
    root = pkg.root("word/comments.xml")
    victim = _find_comment(pkg, comment_id)
    doomed_ids = {str(comment_id)}
    doomed_para_ids = set()
    pid = _last_para_id(victim)
    if pid:
        doomed_para_ids.add(pid)

    # Cascade: find replies via commentsExtended parent links.
    if pkg.has_part("word/commentsExtended.xml") and pid:
        ex_root = pkg.root("word/commentsExtended.xml")
        para_to_comment = {}
        for c in root.findall(qn("w:comment")):
            p_id = _last_para_id(c)
            if p_id:
                para_to_comment[p_id] = c.get(qn("w:id"))
        frontier = {pid}
        while frontier:
            nxt = set()
            for ex in ex_root.findall(f"{{{_W15}}}commentEx"):
                if ex.get(f"{{{_W15}}}paraIdParent") in frontier:
                    child_pid = ex.get(f"{{{_W15}}}paraId")
                    if child_pid not in doomed_para_ids:
                        doomed_para_ids.add(child_pid)
                        nxt.add(child_pid)
                        if child_pid in para_to_comment:
                            doomed_ids.add(para_to_comment[child_pid])
            frontier = nxt

    # Remove comment elements.
    for c in list(root.findall(qn("w:comment"))):
        if c.get(qn("w:id")) in doomed_ids:
            root.remove(c)
    pkg.mark_dirty("word/comments.xml")

    # Remove body markers.
    body = pkg.root()
    removed_markers = 0
    for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
        for el in list(body.iter(qn(tag))):
            if el.get(qn("w:id")) in doomed_ids:
                el.getparent().remove(el)
                removed_markers += 1
    for ref in list(body.iter(qn("w:commentReference"))):
        if ref.get(qn("w:id")) in doomed_ids:
            run = ref.getparent()
            run.getparent().remove(run)
            removed_markers += 1
    pkg.mark_dirty()

    # Remove metadata entries.
    if pkg.has_part("word/commentsExtended.xml"):
        ex_root = pkg.root("word/commentsExtended.xml")
        for ex in list(ex_root.findall(f"{{{_W15}}}commentEx")):
            if ex.get(f"{{{_W15}}}paraId") in doomed_para_ids:
                ex_root.remove(ex)
        pkg.mark_dirty("word/commentsExtended.xml")
    if pkg.has_part("word/commentsIds.xml"):
        ids_root = pkg.root("word/commentsIds.xml")
        for ce in list(ids_root):
            if ce.get(f"{{{_W16CID}}}paraId") in doomed_para_ids:
                ids_root.remove(ce)
        pkg.mark_dirty("word/commentsIds.xml")

    return {
        "deleted_comments": sorted(doomed_ids),
        "body_markers_removed": removed_markers,
    }
