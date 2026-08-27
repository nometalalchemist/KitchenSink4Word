"""One-call submission prep: accept revisions, strip comments, scrub metadata.

Composes the existing revisions/comments ops plus a metadata scrub, and
reports exactly what was done and what deliberately was NOT touched. Content
is never removed — footnotes, endnotes, fields, and body text all survive.

What the metadata scrub covers: docProps/core.xml creator and lastModifiedBy
(emptied; title kept by default), docProps/app.xml Company and Manager
(emptied if present), and word/people.xml (removed with its content-type and
relationship entries — it lists comment authors by name). w:rsid* attributes
are left alone on purpose: they are weak fingerprints at most, and stripping
them rewrites every paragraph in document.xml for marginal benefit.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import DocumentProtected
from ..core.package import DocxPackage
from . import comments as _cm
from . import protection as _pr
from . import revisions as _rv
from .read import get_comments, list_endnotes, list_footnotes, revision_summary

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_DC = "http://purl.org/dc/elements/1.1/"
_CP = (
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
)
_EP = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)

_COMMENT_FAMILY = (
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/people.xml",
)


def _remove_part(pkg: DocxPackage, name: str) -> bool:
    """Drop a part plus its content-type override and document relationship.

    DocxPackage has no public delete-part API, so this reaches into its part
    tables directly; save() serializes strictly from _order/_raw/_dirty, so
    removing an entry from all of them removes the part from the output.
    (Candidate for promotion into DocxPackage as remove_part().)
    """
    if not pkg.has_part(name):
        return False
    pkg._raw.pop(name)
    pkg._order.remove(name)
    pkg._trees.pop(name, None)
    pkg._dirty.discard(name)

    ct_root = pkg.root("[Content_Types].xml")
    for o in list(ct_root.findall(f"{{{_CT_NS}}}Override")):
        if o.get("PartName") == "/" + name:
            ct_root.remove(o)
            pkg.mark_dirty("[Content_Types].xml")

    rels_part = "word/_rels/document.xml.rels"
    if name.startswith("word/") and pkg.has_part(rels_part):
        target = name.split("/", 1)[1]
        rels_root = pkg.root(rels_part)
        for r in list(rels_root):
            if r.get("Target") == target:
                rels_root.remove(r)
                pkg.mark_dirty(rels_part)

    # The part's own .rels, when it has one (e.g. images inside comments).
    own_rels = (
        name.rsplit("/", 1)[0] + "/_rels/" + name.rsplit("/", 1)[1] + ".rels"
    )
    if pkg.has_part(own_rels):
        pkg._raw.pop(own_rels)
        pkg._order.remove(own_rels)
        pkg._trees.pop(own_rels, None)
        pkg._dirty.discard(own_rels)
    return True


def _empty_text(root: etree._Element, ns: str, local: str) -> bool:
    el = root.find(f"{{{ns}}}{local}")
    if el is not None and (el.text or ""):
        el.text = ""
        return True
    return False


def prepare_for_submission(
    pkg: DocxPackage,
    *,
    accept_revisions: bool = True,
    remove_comments: bool = True,
    scrub_metadata: bool = True,
    keep_title: bool = True,
) -> dict:
    """Prepare a document for external submission in one pass: accept all
    tracked changes (every author), delete all comments including the
    comments-family parts, and scrub identifying metadata (author,
    last-modified-by, company; title kept unless keep_title=False).

    Refuses protected documents outright rather than delivering a half-clean
    file — lift the restriction first with remove_document_protection.

    Content is NEVER removed: footnotes, endnotes, citations, and fields all
    stay. The result lists what was done, what was deliberately left (rsids),
    and what remains in the document so nothing ships by surprise."""
    prot = _pr.get_protection(pkg)
    if prot.get("protected"):
        raise DocumentProtected(
            f"document has an enforced editing restriction "
            f"(edit={prot.get('edit')!r}); refusing to half-clean it — run "
            "remove_document_protection first, then retry"
        )

    result: dict = {"actions": []}

    if accept_revisions:
        rev = _rv.accept_revisions(pkg, author=None)
        result["revisions_accepted"] = rev["revisions_resolved"]
        if rev.get("note_definitions_purged"):
            result["note_definitions_purged"] = rev["note_definitions_purged"]
        result["actions"].append("accepted all tracked changes (all authors)")

    if remove_comments:
        removed = 0
        # delete_comment cascades through replies, so delete the head of the
        # list until none remain; the guard bounds a pathological cycle.
        for _ in range(10000):
            current = get_comments(pkg)
            if not current:
                break
            gone = _cm.delete_comment(pkg, comment_id=current[0]["id"])
            removed += len(gone["deleted_comments"])
        result["comments_removed"] = removed
        parts_removed = [
            part for part in _COMMENT_FAMILY if _remove_part(pkg, part)
        ]
        result["comment_parts_removed"] = parts_removed
        result["actions"].append(
            f"deleted {removed} comment(s) and removed "
            f"{len(parts_removed)} comment-family part(s)"
        )

    if scrub_metadata:
        scrubbed: list[str] = []
        if pkg.has_part("docProps/core.xml"):
            core = pkg.root("docProps/core.xml")
            if _empty_text(core, _DC, "creator"):
                scrubbed.append("core.xml creator")
            if _empty_text(core, _CP, "lastModifiedBy"):
                scrubbed.append("core.xml lastModifiedBy")
            if not keep_title and _empty_text(core, _DC, "title"):
                scrubbed.append("core.xml title")
            pkg.mark_dirty("docProps/core.xml")
        if pkg.has_part("docProps/app.xml"):
            app = pkg.root("docProps/app.xml")
            changed = False
            for local in ("Company", "Manager"):
                if _empty_text(app, _EP, local):
                    scrubbed.append(f"app.xml {local}")
                    changed = True
            if changed:
                pkg.mark_dirty("docProps/app.xml")
        if _remove_part(pkg, "word/people.xml"):
            scrubbed.append("word/people.xml removed")
        result["metadata_scrubbed"] = scrubbed
        result["actions"].append(
            "scrubbed metadata: " + (", ".join(scrubbed) or "nothing present")
        )

    result["not_removed"] = [
        "w:rsid* revision-save attributes (stripping them rewrites the whole "
        "document XML for marginal privacy benefit)",
    ]

    # Remaining-items report: normal content that still ships with the file.
    remaining: dict = {
        "footnotes": len(list_footnotes(pkg)),
        "endnotes": len(list_endnotes(pkg)),
        "tracked_changes": revision_summary(pkg)["total"],
        "comments": len(get_comments(pkg)),
    }
    result["remaining"] = remaining
    if remaining["footnotes"] or remaining["endnotes"]:
        result["actions"].append(
            f"document still contains {remaining['footnotes']} footnote(s) "
            f"and {remaining['endnotes']} endnote(s) — content is kept, this "
            "is normal"
        )
    return result
