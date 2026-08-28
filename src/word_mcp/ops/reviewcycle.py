"""Review-cycle analytics: reviewer matrix, revision analytics, structured diff.

Everything in this module is READ-ONLY: functions inspect packages and return
plain dict structures; nothing is ever marked dirty or saved. Comment and
revision extraction reuses the readers in ops/read.py (get_comments,
get_tracked_changes-style tag handling, outline detection) so the numbers here
always agree with the plain read tools.

Section attribution ("heading path") walks the body in document order keeping a
stack of the headings currently in force, so every comment, revision, and diff
change can say which part of the outline it falls under.
"""

from __future__ import annotations

import difflib

from lxml import etree

from ..core.package import DocxPackage, qn
from . import read as _read

_NO_SECTION = "(before first heading)"

# Revision wrapper tags counted by revision_analytics, bucketed by what they
# mean for the author's edit volume.
_INS_LIKE = {"ins"}
_DEL_LIKE = {"del"}
_MOVE_TAGS = {"moveFrom", "moveTo"}
_FORMAT_TAGS = {
    "rPrChange",
    "pPrChange",
    "tblPrChange",
    "sectPrChange",
    "cellIns",
    "cellDel",
}


# ------------------------------------------------------------- shared helpers


def _section_key(path: list[str]) -> str:
    return " > ".join(path) if path else _NO_SECTION


def _walk_body_with_sections(pkg: DocxPackage):
    """Yield (kind, index, element, heading_path) for body-level paragraphs and
    tables in document order. heading_path is the list of heading texts in
    force AT that element; a heading paragraph opens its own section."""
    style_outline = _read._style_outline_map(pkg)
    stack: list[tuple[int, str]] = []
    for kind, idx, el in _read.body_items(pkg):
        if kind == "paragraph":
            lvl = _read._outline_level(el, style_outline)
            if lvl is not None:
                text = _read.paragraph_text(el).strip()
                if text:
                    while stack and stack[-1][0] >= lvl:
                        stack.pop()
                    stack.append((lvl, text))
        yield kind, idx, el, [t for _, t in stack]


def _preview(text: str, limit: int = 120) -> str:
    text = text or ""
    return text if len(text) <= limit else text[:limit] + "..."


# ------------------------------------------------------------ comment report


def _comment_locations(pkg: DocxPackage) -> dict[str, dict]:
    """comment id -> body location: paragraph_index (None when the anchor sits
    inside a table), table_index when applicable, heading_path, section."""
    out: dict[str, dict] = {}
    for kind, idx, el, path in _walk_body_with_sections(pkg):
        for marker in el.iter(qn("w:commentRangeStart"), qn("w:commentReference")):
            cid = marker.get(qn("w:id"))
            if cid is None or cid in out:
                continue
            loc = {
                "heading_path": list(path),
                "section": _section_key(path),
            }
            if kind == "paragraph":
                loc["paragraph_index"] = idx
            else:
                loc["paragraph_index"] = None
                loc["table_index"] = idx
            out[cid] = loc
    return out


def comment_report(pkg: DocxPackage, *, include_resolved: bool = True) -> dict:
    """The reviewer matrix: every comment thread with author, initials, date,
    anchored text, comment text, nested replies, resolved flag, and a body
    locator (paragraph index plus the heading path of the section it falls
    under). Summary blocks give per-author, resolved/open, and per-section
    counts. Complete by design: an agent processing committee feedback should
    never need to re-read comments piecemeal after this one call.

    include_resolved=False drops threads whose ROOT comment is resolved (the
    Word model marks the thread on the root); the number excluded is reported.
    Read-only: never modifies the document.
    """
    flat = _read.get_comments(pkg)
    locations = _comment_locations(pkg)

    entries: dict[str, dict] = {}
    for c in flat:
        entry = dict(c)
        loc = locations.get(c["id"], {})
        entry["paragraph_index"] = loc.get("paragraph_index")
        if "table_index" in loc:
            entry["table_index"] = loc["table_index"]
        entry["heading_path"] = loc.get("heading_path", [])
        entry["section"] = loc.get("section", _NO_SECTION)
        entry["replies"] = []
        entries[c["id"]] = entry

    threads: list[dict] = []
    for c in flat:
        entry = entries[c["id"]]
        parent_id = entry.pop("reply_to", None)
        if parent_id is not None and parent_id in entries:
            entries[parent_id]["replies"].append(entry)
        else:
            threads.append(entry)

    excluded = 0
    if not include_resolved:
        kept = [t for t in threads if not t["resolved"]]
        excluded = len(threads) - len(kept)
        threads = kept

    def _all_comments(thread: dict):
        yield thread
        for r in thread["replies"]:
            yield from _all_comments(r)

    by_author: dict[str, int] = {}
    by_section: dict[str, int] = {}
    total_comments = 0
    open_threads = resolved_threads = 0
    for t in threads:
        if t["resolved"]:
            resolved_threads += 1
        else:
            open_threads += 1
        by_section[t["section"]] = by_section.get(t["section"], 0) + 1
        for c in _all_comments(t):
            total_comments += 1
            by_author[c["author"]] = by_author.get(c["author"], 0) + 1

    return {
        "comments": threads,
        "summary": {
            "threads": len(threads),
            "total_comments": total_comments,
            "open_threads": open_threads,
            "resolved_threads": resolved_threads,
            "by_author": by_author,
            "by_section": by_section,
        },
        "include_resolved": include_resolved,
        "resolved_threads_excluded": excluded,
    }


def comment_report_multi(paths: list[str]) -> dict:
    """The reviewer matrix across SEVERAL documents (e.g., three committee
    members' copies of the same draft), merged into one view keyed by
    (author, anchored text) with per-file provenance on every occurrence.
    Collision detection flags any anchored span commented on by two or more
    DIFFERENT reviewers (across files or within one file) so conflicting or
    overlapping feedback is surfaced instead of discovered mid-edit.

    Merging is by exact anchored-text match (whitespace-normalized), so the
    documents should be copies of the same draft; a span reworded between
    copies will not merge. Opens each file read-only; nothing is modified.
    """
    files: list[dict] = []
    merged: dict[tuple[str, str], dict] = {}
    span_authors: dict[str, dict[str, list[dict]]] = {}

    for path in paths:
        rep = comment_report(DocxPackage(path))
        file_authors: set[str] = set()
        for thread in rep["comments"]:
            author = thread["author"]
            file_authors.add(author)
            anchor = " ".join((thread["anchored_text"] or "").split())
            key = (author, anchor)
            slot = merged.setdefault(
                key,
                {"author": author, "anchored_text": anchor, "occurrences": []},
            )
            slot["occurrences"].append(
                {
                    "file": str(path),
                    "comment_id": thread["id"],
                    "date": thread["date"],
                    "text": thread["text"],
                    "resolved": thread["resolved"],
                    "paragraph_index": thread["paragraph_index"],
                    "heading_path": thread["heading_path"],
                    "section": thread["section"],
                    "replies": thread["replies"],
                }
            )
            if anchor:
                span_authors.setdefault(anchor, {}).setdefault(author, []).append(
                    {"file": str(path), "comment_id": thread["id"]}
                )
        files.append(
            {
                "path": str(path),
                "threads": rep["summary"]["threads"],
                "total_comments": rep["summary"]["total_comments"],
                "authors": sorted(file_authors),
            }
        )

    collisions = []
    for anchor, authors in sorted(span_authors.items()):
        if len(authors) < 2:
            continue
        entries = []
        for author, occs in sorted(authors.items()):
            for occ in occs:
                entries.append({"author": author, **occ})
        collisions.append(
            {
                "anchored_text": anchor,
                "authors": sorted(authors),
                "entries": entries,
            }
        )

    by_author: dict[str, int] = {}
    for (author, _), slot in merged.items():
        by_author[author] = by_author.get(author, 0) + len(slot["occurrences"])

    merged_list = [merged[k] for k in sorted(merged)]
    return {
        "files": files,
        "merged": merged_list,
        "collisions": collisions,
        "summary": {
            "files": len(files),
            "total_threads": sum(f["threads"] for f in files),
            "merged_entries": len(merged_list),
            "collision_count": len(collisions),
            "by_author": by_author,
        },
    }


# --------------------------------------------------------- revision analytics


def _new_author_bucket() -> dict:
    return {
        "insertions": 0,
        "deletions": 0,
        "moves": 0,
        "format_changes": 0,
        "words_added": 0,
        "words_removed": 0,
        "_dates": [],
        "by_section": {},
    }


def revision_analytics(pkg: DocxPackage) -> dict:
    """Tracked-change analytics per author: insertion and deletion counts,
    words added and removed, move and format-change counts, the date range of
    the author's edits (from w:date where present), and a per-section
    (heading-path) breakdown showing where each author's changes concentrate.
    Also returns the 10 heaviest body paragraphs by revision churn (characters
    inserted plus deleted), each with its section, authors, and a preview.

    Notes on the counting: moved text is counted under "moves", not as words
    added/removed (a move is not new prose); paragraph-mark and table-row
    revisions count as insertions/deletions with zero words. Footnote and
    endnote revisions are included in author totals under the pseudo-sections
    "[footnotes]" and "[endnotes]" but are not candidates for the heaviest-
    paragraphs list (that list is body paragraphs only). Read-only.
    """
    authors: dict[str, dict] = {}
    para_churn: dict[int, dict] = {}
    total = 0

    def account(el: etree._Element, section: str, para_idx: int | None) -> None:
        nonlocal total
        local = etree.QName(el).localname
        author = el.get(qn("w:author"), "")
        date = el.get(qn("w:date"), "")
        bucket = authors.setdefault(author, _new_author_bucket())
        chars = 0
        if local in _INS_LIKE:
            text = "".join(_read.run_text(r) for r in el.iter(qn("w:r")))
            bucket["insertions"] += 1
            bucket["words_added"] += len(text.split())
            chars = len(text)
        elif local in _DEL_LIKE:
            text = "".join(
                _read.run_text(r, include_deleted=True) for r in el.iter(qn("w:r"))
            )
            bucket["deletions"] += 1
            bucket["words_removed"] += len(text.split())
            chars = len(text)
        elif local in _MOVE_TAGS:
            bucket["moves"] += 1
        else:
            bucket["format_changes"] += 1
        if date:
            bucket["_dates"].append(date)
        bucket["by_section"][section] = bucket["by_section"].get(section, 0) + 1
        total += 1
        if para_idx is not None:
            churn = para_churn.setdefault(
                para_idx,
                {"revisions": 0, "chars_changed": 0, "authors": set()},
            )
            churn["revisions"] += 1
            churn["chars_changed"] += chars
            churn["authors"].add(author)

    watched = _INS_LIKE | _DEL_LIKE | _MOVE_TAGS | _FORMAT_TAGS
    para_meta: dict[int, tuple[str, str]] = {}  # idx -> (preview, section)
    for kind, idx, el, path in _walk_body_with_sections(pkg):
        section = _section_key(path)
        para_idx = idx if kind == "paragraph" else None
        if para_idx is not None:
            para_meta[para_idx] = (
                _preview(_read.paragraph_text(el).strip(), 100),
                section,
            )
        for sub in el.iter():
            if etree.QName(sub).localname in watched:
                account(sub, section, para_idx)

    for part, section in (
        ("word/footnotes.xml", "[footnotes]"),
        ("word/endnotes.xml", "[endnotes]"),
    ):
        if not pkg.has_part(part):
            continue
        for sub in pkg.root(part).iter():
            if etree.QName(sub).localname in watched:
                account(sub, section, None)

    by_author = {}
    for author, bucket in authors.items():
        dates = bucket.pop("_dates")
        bucket["date_range"] = (
            {"first": min(dates), "last": max(dates)} if dates else None
        )
        by_author[author] = bucket

    heaviest = sorted(
        para_churn.items(),
        key=lambda kv: (-kv[1]["chars_changed"], -kv[1]["revisions"], kv[0]),
    )[:10]
    heaviest_list = [
        {
            "paragraph_index": idx,
            "revisions": data["revisions"],
            "chars_changed": data["chars_changed"],
            "authors": sorted(data["authors"]),
            "section": para_meta.get(idx, ("", _NO_SECTION))[1],
            "text_preview": para_meta.get(idx, ("", ""))[0],
        }
        for idx, data in heaviest
    ]

    return {
        "total_revisions": total,
        "by_author": by_author,
        "heaviest_paragraphs": heaviest_list,
    }


# ------------------------------------------------------------ structured diff

_MOVE_MIN_CHARS = 15  # identical text shorter than this is not called a move


def _body_paragraphs(pkg: DocxPackage) -> tuple[list[str], list[str], dict[int, str]]:
    """(texts, normalized texts, paragraph index -> section key)."""
    texts: list[str] = []
    norms: list[str] = []
    sections: dict[int, str] = {}
    for kind, idx, el, path in _walk_body_with_sections(pkg):
        if kind != "paragraph":
            continue
        text = _read.paragraph_text(el)
        texts.append(text)
        norms.append(" ".join(text.split()))
        sections[idx] = _section_key(path)
    return texts, norms, sections


def _tables(pkg: DocxPackage) -> list[dict]:
    return [
        _read._table_dict(el, idx)
        for kind, idx, el in _read.body_items(pkg)
        if kind == "table"
    ]


def _intra_changes(old_text: str, new_text: str, snippet: int = 200) -> list[dict]:
    sm = difflib.SequenceMatcher(None, old_text, new_text, autojunk=False)
    out = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        out.append(
            {
                "op": op,
                "old": _preview(old_text[i1:i2], snippet),
                "new": _preview(new_text[j1:j2], snippet),
            }
        )
    return out


def _diff_tables(pkg_old: DocxPackage, pkg_new: DocxPackage) -> dict:
    old_tables = _tables(pkg_old)
    new_tables = _tables(pkg_new)
    changes = []
    cell_cap = 50
    for i in range(min(len(old_tables), len(new_tables))):
        ot, nt = old_tables[i], new_tables[i]
        old_dims = {"rows": ot["rows"], "columns": ot["grid_columns"]}
        new_dims = {"rows": nt["rows"], "columns": nt["grid_columns"]}
        changed_cells = []
        overflow = 0
        for r in range(min(len(ot["cells"]), len(nt["cells"]))):
            orow, nrow = ot["cells"][r], nt["cells"][r]
            for c in range(min(len(orow), len(nrow))):
                if orow[c]["text"] != nrow[c]["text"]:
                    if len(changed_cells) < cell_cap:
                        changed_cells.append(
                            {
                                "row": r,
                                "col": c,
                                "old": _preview(orow[c]["text"]),
                                "new": _preview(nrow[c]["text"]),
                            }
                        )
                    else:
                        overflow += 1
        if old_dims != new_dims or changed_cells or overflow:
            changes.append(
                {
                    "table_index": i,
                    "old_dims": old_dims,
                    "new_dims": new_dims,
                    "dimensions_changed": old_dims != new_dims,
                    "changed_cells": changed_cells,
                    "changed_cells_not_shown": overflow,
                }
            )
    return {
        "table_count": {
            "old": len(old_tables),
            "new": len(new_tables),
            "delta": len(new_tables) - len(old_tables),
        },
        "changed_tables": changes,
        "alignment": "positional (table N in the old document is compared to "
        "table N in the new one)",
    }


def _heading_changes(pkg_old: DocxPackage, pkg_new: DocxPackage) -> dict:
    old_outline = _read.get_outline(pkg_old)
    new_outline = _read.get_outline(pkg_new)
    a = [(h["level"], h["text"]) for h in old_outline]
    b = [(h["level"], h["text"]) for h in new_outline]
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    removed, added = [], []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op in ("replace", "delete"):
            removed.extend({"level": l, "text": t} for l, t in a[i1:i2])
        if op in ("replace", "insert"):
            added.extend({"level": l, "text": t} for l, t in b[j1:j2])
    return {"added": added, "removed": removed}


def structured_diff(
    old_path: str, new_path: str, *, detail_cap: int = 200
) -> dict:
    """Agent-readable structural diff between two saved drafts, computed
    without Word. Body paragraphs are aligned with difflib.SequenceMatcher
    over whitespace-normalized paragraph texts, then modified pairs get
    intra-paragraph opcodes. Reports: unchanged/modified/inserted/deleted
    paragraphs (with indices in both documents), MOVED paragraphs (identical
    text of 15+ characters at a different position), a per-section change
    summary keyed by heading path, table changes (positional table-to-table
    cell-grid comparison: dimension changes plus changed cells), footnote and
    endnote count deltas, and heading structure changes.

    This output is JSON-shaped data for programmatic consumption, deliberately
    NOT a redline: for a Word-rendered tracked-changes comparison use
    com_compare_documents instead. On huge diffs, per-change detail is capped
    at detail_cap entries (default 200, consumed in order modified, moved,
    inserted, deleted); the remainder is summarized in the counts and the cap
    is reported via detail_capped. Opens both files read-only.
    """
    pkg_old = DocxPackage(old_path)
    pkg_new = DocxPackage(new_path)

    old_texts, old_norms, old_sections = _body_paragraphs(pkg_old)
    new_texts, new_norms, new_sections = _body_paragraphs(pkg_new)

    sm = difflib.SequenceMatcher(None, old_norms, new_norms, autojunk=False)
    opcodes = sm.get_opcodes()

    # Pools of texts sitting in changed regions, both sides — used to spot
    # SWAPPED blocks: SequenceMatcher reports them as one big "replace",
    # whose positional pairing produced junk modified pairs while the move
    # pass (which only saw insert/delete leftovers) found nothing (v1.5
    # adversarial finding F3). A replace-pair member whose exact text exists
    # elsewhere in the other side's changed regions is deferred to the move
    # pass instead of being force-paired.
    from collections import Counter

    new_pool: Counter = Counter()
    old_pool: Counter = Counter()
    for op, i1, i2, j1, j2 in opcodes:
        if op in ("replace", "insert"):
            for j in range(j1, j2):
                new_pool[new_norms[j]] += 1
        if op in ("replace", "delete"):
            for i in range(i1, i2):
                old_pool[old_norms[i]] += 1

    unchanged = 0
    modified: list[dict] = []
    inserted: list[dict] = []
    deleted: list[dict] = []
    for op, i1, i2, j1, j2 in opcodes:
        if op == "equal":
            unchanged += i2 - i1
        elif op == "replace":
            pairs = min(i2 - i1, j2 - j1)
            for k in range(pairs):
                oi, ni = i1 + k, j1 + k
                o_norm, n_norm = old_norms[oi], new_norms[ni]
                if o_norm != n_norm and (
                    (len(o_norm) >= _MOVE_MIN_CHARS and new_pool[o_norm] > 0)
                    or (len(n_norm) >= _MOVE_MIN_CHARS and old_pool[n_norm] > 0)
                ):
                    deleted.append({"old_index": oi})
                    inserted.append({"new_index": ni})
                    continue
                modified.append(
                    {
                        "old_index": oi,
                        "new_index": ni,
                        "section": new_sections.get(ni, _NO_SECTION),
                        "old_text": _preview(old_texts[oi]),
                        "new_text": _preview(new_texts[ni]),
                        "changes": _intra_changes(old_texts[oi], new_texts[ni]),
                    }
                )
            for oi in range(i1 + pairs, i2):
                deleted.append({"old_index": oi})
            for ni in range(j1 + pairs, j2):
                inserted.append({"new_index": ni})
        elif op == "delete":
            for oi in range(i1, i2):
                deleted.append({"old_index": oi})
        elif op == "insert":
            for ni in range(j1, j2):
                inserted.append({"new_index": ni})

    # Moved detection: a deleted paragraph whose normalized text reappears
    # verbatim among the inserted ones is a move, not a delete+insert.
    moved: list[dict] = []
    remaining_inserted: list[dict] = []
    ins_by_norm: dict[str, list[dict]] = {}
    for entry in inserted:
        ins_by_norm.setdefault(new_norms[entry["new_index"]], []).append(entry)
    still_deleted: list[dict] = []
    for entry in deleted:
        norm = old_norms[entry["old_index"]]
        pool = ins_by_norm.get(norm)
        if pool and len(norm) >= _MOVE_MIN_CHARS:
            match = pool.pop(0)
            moved.append(
                {
                    "old_index": entry["old_index"],
                    "new_index": match["new_index"],
                    "section": new_sections.get(match["new_index"], _NO_SECTION),
                    "text": _preview(new_texts[match["new_index"]]),
                }
            )
        else:
            still_deleted.append(entry)
    for pool in ins_by_norm.values():
        remaining_inserted.extend(pool)
    remaining_inserted.sort(key=lambda e: e["new_index"])
    deleted = still_deleted
    inserted = remaining_inserted

    for entry in deleted:
        oi = entry["old_index"]
        entry["section"] = old_sections.get(oi, _NO_SECTION)
        entry["text"] = _preview(old_texts[oi])
    for entry in inserted:
        ni = entry["new_index"]
        entry["section"] = new_sections.get(ni, _NO_SECTION)
        entry["text"] = _preview(new_texts[ni])

    by_section: dict[str, dict] = {}

    def bump(section: str, field: str) -> None:
        slot = by_section.setdefault(
            section, {"modified": 0, "inserted": 0, "deleted": 0, "moved": 0}
        )
        slot[field] += 1

    for e in modified:
        bump(e["section"], "modified")
    for e in moved:
        bump(e["section"], "moved")
    for e in inserted:
        bump(e["section"], "inserted")
    for e in deleted:
        bump(e["section"], "deleted")

    counts = {
        "unchanged": unchanged,
        "modified": len(modified),
        "moved": len(moved),
        "inserted": len(inserted),
        "deleted": len(deleted),
        "old_paragraphs": len(old_texts),
        "new_paragraphs": len(new_texts),
    }
    total_changes = (
        counts["modified"] + counts["moved"] + counts["inserted"] + counts["deleted"]
    )
    detail_capped = total_changes > detail_cap
    if detail_capped:
        budget = detail_cap
        capped_lists = []
        for lst in (modified, moved, inserted, deleted):
            shown = lst[: max(budget, 0)]
            budget -= len(shown)
            capped_lists.append(shown)
        modified, moved, inserted, deleted = capped_lists

    fn_old = len(_read.list_footnotes(pkg_old))
    fn_new = len(_read.list_footnotes(pkg_new))
    en_old = len(_read.list_endnotes(pkg_old))
    en_new = len(_read.list_endnotes(pkg_new))
    tables = _diff_tables(pkg_old, pkg_new)

    return {
        "old_path": str(old_path),
        "new_path": str(new_path),
        "counts": counts,
        "modified": modified,
        "moved": moved,
        "inserted": inserted,
        "deleted": deleted,
        "by_section": by_section,
        "tables": tables,
        "footnotes": {"old": fn_old, "new": fn_new, "delta": fn_new - fn_old},
        "endnotes": {"old": en_old, "new": en_new, "delta": en_new - en_old},
        "headings": _heading_changes(pkg_old, pkg_new),
        "detail_cap": detail_cap,
        "detail_capped": detail_capped,
        "identical": total_changes == 0
        and fn_old == fn_new
        and en_old == en_new
        and not tables["changed_tables"]
        and tables["table_count"]["delta"] == 0,
    }
