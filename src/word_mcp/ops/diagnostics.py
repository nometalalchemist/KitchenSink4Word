"""One-call structural health report for a .docx package.

Read-only. Composes the existing validators (notes, cross-references) with
package-level checks: content-type coverage, relationship integrity, orphan
parts, field balance, undefined style/numbering references, SDT and bookmark
sanity, revision ids, image relationships, and a size profile.

Contract: diagnose_document NEVER raises on a weird-but-openable document —
every check degrades to a reported problem instead of an exception.

Severity model: "error" = will render broken or lose content in Word;
"warning" = latent bug (renders wrong or silently falls back);
"info" = harmless but worth knowing. ok = no "error" problems.
"""

from __future__ import annotations

import posixpath
import re

from lxml import etree

from ..core.package import DocxPackage, qn

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_REVISION_ID_TAGS = (
    "w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange", "w:pPrChange",
    "w:tblPrChange", "w:sectPrChange", "w:cellIns", "w:cellDel",
)


def _story_parts(pkg: DocxPackage) -> list[str]:
    parts = ["word/document.xml"]
    for name in pkg.part_names():
        if re.fullmatch(r"word/(header|footer)\d+\.xml", name):
            parts.append(name)
    for extra in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
        if pkg.has_part(extra):
            parts.append(extra)
    return parts


def _resolve_rel_target(rels_part: str, target: str) -> str:
    """Resolve a relationship Target to a package part name."""
    if target.startswith("/"):
        return posixpath.normpath(target[1:])
    # word/_rels/document.xml.rels -> base dir "word"; _rels/.rels -> ""
    base = posixpath.dirname(posixpath.dirname(rels_part))
    return posixpath.normpath(posixpath.join(base, target) if base else target)


def _part_rels_name(part: str) -> str:
    d, _, base = part.rpartition("/")
    return f"{d}/_rels/{base}.rels" if d else f"_rels/{base}.rels"


def diagnose_document(pkg: DocxPackage) -> dict:
    """Structural health report: {"ok", "problems": [{category, severity,
    detail, location?}], "info": {...}}. Read-only; see module docstring for
    the severity model."""
    problems: list[dict] = []
    info: dict = {}

    def add(category: str, severity: str, detail: str, location: str | None = None):
        entry = {"category": category, "severity": severity, "detail": detail}
        if location:
            entry["location"] = location
        problems.append(entry)

    # Checks whose FAILURE means the document itself is broken (a check
    # that cannot even parse document.xml is not a tooling hiccup — it is
    # the diagnosis) escalate to error severity and flip ok to False.
    _CORE_CHECKS = {"content_types", "relationships", "fields", "notes"}

    def run(name, fn):
        try:
            fn()
        except Exception as exc:  # degrade to reporting, never raise
            add(
                "diagnostic_error",
                "error" if name in _CORE_CHECKS else "warning",
                f"{name} check could not complete: {exc}",
            )

    # ---------------------------------------------- content-type coverage
    def check_content_types():
        if not pkg.has_part("[Content_Types].xml"):
            add("content_types", "error", "package has no [Content_Types].xml")
            return
        root = pkg.root("[Content_Types].xml")
        defaults = {
            (d.get("Extension") or "").lower()
            for d in root.findall(f"{{{_CT_NS}}}Default")
        }
        overrides = {
            o.get("PartName") for o in root.findall(f"{{{_CT_NS}}}Override")
        }
        for name in pkg.part_names():
            if name == "[Content_Types].xml":
                continue
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            if "/" + name not in overrides and ext not in defaults:
                add(
                    "content_types",
                    "error",
                    f"part has no content-type Override and no Default for "
                    f"extension {ext!r}",
                    name,
                )

    # -------------------------------------- relationships: dangling + orphan
    referenced_parts: set[str] = set()

    def check_relationships():
        for rels_part in pkg.part_names():
            if not rels_part.endswith(".rels"):
                continue
            for rel in pkg.root(rels_part).findall(f"{{{_REL_NS}}}Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target") or ""
                resolved = _resolve_rel_target(rels_part, target)
                referenced_parts.add(resolved)
                if not pkg.has_part(resolved):
                    add(
                        "relationships",
                        "error",
                        f"relationship {rel.get('Id')} targets missing part "
                        f"{resolved!r}",
                        rels_part,
                    )

    def check_orphan_parts():
        # Parts reachable from nothing. Word ignores them; dead weight only.
        exempt = {"[Content_Types].xml"}
        for name in pkg.part_names():
            if name in exempt or name.endswith(".rels"):
                continue
            if name not in referenced_parts:
                add(
                    "orphan_parts",
                    "info",
                    "part is not the target of any relationship",
                    name,
                )

    # ------------------------------------------------ field marker balance
    def check_fields():
        report = {}
        for part in _story_parts(pkg):
            begins = separates = ends = 0
            for fc in pkg.root(part).iter(qn("w:fldChar")):
                t = fc.get(qn("w:fldCharType"))
                if t == "begin":
                    begins += 1
                elif t == "separate":
                    separates += 1
                elif t == "end":
                    ends += 1
            report[part] = {"begin": begins, "separate": separates, "end": ends}
            if begins != ends:
                add(
                    "fields",
                    "error",
                    f"unbalanced field markers: {begins} begin vs {ends} end",
                    part,
                )
            elif separates > begins:
                add(
                    "fields",
                    "warning",
                    f"{separates} separate markers for {begins} begin markers",
                    part,
                )
        info["fields"] = report

    # ----------------------------------------------------- notes (existing)
    def check_notes():
        from .notes import validate_notes

        report = validate_notes(pkg)
        info["notes"] = report
        for kind, r in report.items():
            for nid in r["missing_definitions"]:
                add(
                    "notes",
                    "error",
                    f"{kind[:-1]} reference id {nid} has no definition",
                    "word/document.xml",
                )
            for nid in r["duplicate_references"]:
                add(
                    "notes",
                    "warning",
                    f"{kind[:-1]} id {nid} is referenced more than once",
                    "word/document.xml",
                )
            if r["orphan_definitions"]:
                add(
                    "notes",
                    "info",
                    f"{len(r['orphan_definitions'])} orphan {kind} "
                    "definitions (dead weight; cleanup available)",
                )

    # -------------------------------------- undefined style references
    def check_styles():
        defined: set[str] = set()
        if pkg.has_part("word/styles.xml"):
            defined = {
                s.get(qn("w:styleId"))
                for s in pkg.root("word/styles.xml").findall(qn("w:style"))
            }
        missing: dict[str, set[str]] = {}
        for part in _story_parts(pkg):
            root = pkg.root(part)
            for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
                for el in root.iter(qn(tag)):
                    val = el.get(qn("w:val"))
                    if val and val not in defined:
                        missing.setdefault(val, set()).add(part)
        for sid in sorted(missing):
            add(
                "styles",
                "warning",
                f"style {sid!r} is referenced but not defined in styles.xml "
                "(Word silently renders it as Normal/default)",
                ", ".join(sorted(missing[sid])),
            )

    # ---------------------------------- undefined numbering references
    def check_numbering():
        used: dict[str, set[str]] = {}
        for part in _story_parts(pkg):
            for numid in pkg.root(part).iter(qn("w:numId")):
                val = numid.get(qn("w:val"))
                if val and val != "0":  # numId 0 = "no numbering"
                    used.setdefault(val, set()).add(part)
        defined: set[str] = set()
        abstract_defined: set[str] = set()
        if pkg.has_part("word/numbering.xml"):
            nroot = pkg.root("word/numbering.xml")
            abstract_defined = {
                a.get(qn("w:abstractNumId"))
                for a in nroot.findall(qn("w:abstractNum"))
            }
            for num in nroot.findall(qn("w:num")):
                defined.add(num.get(qn("w:numId")))
                link = num.find(qn("w:abstractNumId"))
                aid = link.get(qn("w:val")) if link is not None else None
                if aid not in abstract_defined:
                    add(
                        "numbering",
                        "warning",
                        f"num {num.get(qn('w:numId'))} links to undefined "
                        f"abstractNum {aid}",
                        "word/numbering.xml",
                    )
        for nid in sorted(used, key=lambda x: (len(x), x)):
            if nid not in defined:
                add(
                    "numbering",
                    "warning",
                    f"numId {nid} is referenced but not defined "
                    "(list renders without numbers)",
                    ", ".join(sorted(used[nid])),
                )

    # -------------------------------------------------------- sdt sanity
    def check_sdt():
        for part in _story_parts(pkg):
            bad = sum(
                1
                for sdt in pkg.root(part).iter(qn("w:sdt"))
                if sdt.find(qn("w:sdtContent")) is None
            )
            if bad:
                add(
                    "sdt",
                    "warning",
                    f"{bad} content control(s) have no sdtContent "
                    "(empty shell; content may have been lost)",
                    part,
                )

    # ------------------------------------------------- bookmark pairing
    def check_bookmarks():
        names_seen: dict[str, int] = {}
        for part in _story_parts(pkg):
            root = pkg.root(part)
            starts = {}
            ends = set()
            for bs in root.iter(qn("w:bookmarkStart")):
                bid = bs.get(qn("w:id"))
                starts[bid] = bs.get(qn("w:name"))
                name = bs.get(qn("w:name"))
                if name:
                    names_seen[name] = names_seen.get(name, 0) + 1
            for be in root.iter(qn("w:bookmarkEnd")):
                ends.add(be.get(qn("w:id")))
            for bid, name in starts.items():
                if bid not in ends:
                    add(
                        "bookmarks",
                        "warning",
                        f"bookmarkStart id {bid} ({name!r}) has no bookmarkEnd",
                        part,
                    )
            for bid in sorted(ends - set(starts), key=lambda x: (len(x or ""), x)):
                add(
                    "bookmarks",
                    "warning",
                    f"bookmarkEnd id {bid} has no bookmarkStart",
                    part,
                )
        for name, count in names_seen.items():
            if count > 1:
                add(
                    "bookmarks",
                    "warning",
                    f"bookmark name {name!r} is defined {count} times",
                )

    # ------------------------------------- vertical merge chain integrity
    def check_table_merges():
        # A vMerge continuation must sit directly below a cell of the same
        # grid extent that is part of the chain (restart or continue).
        # Orphaned continuations render unpredictably in Word and are the
        # footprint of structural edits that broke a merge (field test,
        # 2026-09-03).
        for part in _story_parts(pkg):
            for t_i, tbl in enumerate(pkg.root(part).iter(qn("w:tbl"))):
                prev_row: dict[tuple[int, int], str | None] = {}
                for r_i, tr in enumerate(tbl.findall(qn("w:tr"))):
                    cur: dict[tuple[int, int], str | None] = {}
                    pos = 0
                    trpr = tr.find(qn("w:trPr"))
                    if trpr is not None:
                        gb = trpr.find(qn("w:gridBefore"))
                        if gb is not None:
                            try:
                                pos = int(gb.get(qn("w:val"), "0") or 0)
                            except ValueError:
                                pos = 0
                    for tc in tr.findall(qn("w:tc")):
                        tcpr = tc.find(qn("w:tcPr"))
                        span = 1
                        vm = None
                        if tcpr is not None:
                            gs = tcpr.find(qn("w:gridSpan"))
                            if gs is not None:
                                try:
                                    span = max(
                                        1, int(gs.get(qn("w:val"), "1") or 1)
                                    )
                                except ValueError:
                                    span = 1
                            vme = tcpr.find(qn("w:vMerge"))
                            if vme is not None:
                                vm = vme.get(qn("w:val"), "continue")
                        key = (pos, pos + span)
                        cur[key] = vm
                        if vm == "continue" and prev_row.get(key) not in (
                            "restart", "continue",
                        ):
                            add(
                                "tables",
                                "warning",
                                f"table {t_i} row {r_i}: vMerge continuation "
                                f"at grid column {pos} has no merged cell "
                                "directly above it (orphaned continuation; "
                                "a structural edit likely broke the merge "
                                "chain)",
                                part,
                            )
                        pos += span
                    prev_row = cur

    # -------------------------------------------- revision id duplicates
    def check_revision_ids():
        seen: dict[str, int] = {}
        for tag in _REVISION_ID_TAGS:
            for el in pkg.root().iter(qn(tag)):
                rid = el.get(qn("w:id"))
                if rid is not None:
                    seen[rid] = seen.get(rid, 0) + 1
        dups = {rid: n for rid, n in seen.items() if n > 1}
        if dups:
            add(
                "revisions",
                "info",
                f"{len(dups)} revision id(s) used more than once "
                f"(e.g. {sorted(dups)[:5]})",
                "word/document.xml",
            )

    # ---------------------------------------------------- image integrity
    def check_images():
        for part in _story_parts(pkg):
            rels_part = _part_rels_name(part)
            rels = {}
            if pkg.has_part(rels_part):
                rels = {
                    r.get("Id"): r
                    for r in pkg.root(rels_part).findall(
                        f"{{{_REL_NS}}}Relationship"
                    )
                }
            for blip in pkg.root(part).iter(qn("a:blip")):
                rid = blip.get(f"{{{_R_NS}}}embed") or blip.get(f"{{{_R_NS}}}link")
                if rid is None:
                    add("images", "warning", "image blip has no relationship id", part)
                    continue
                rel = rels.get(rid)
                if rel is None:
                    add(
                        "images",
                        "error",
                        f"image references relationship {rid} which does not "
                        f"exist in {rels_part}",
                        part,
                    )
                elif rel.get("TargetMode") != "External":
                    target = _resolve_rel_target(rels_part, rel.get("Target") or "")
                    if not pkg.has_part(target):
                        add(
                            "images",
                            "error",
                            f"image relationship {rid} targets missing part "
                            f"{target!r}",
                            part,
                        )

    # -------------------------------------- cross-references (existing)
    def check_cross_references():
        from .integrity import validate_cross_references

        report = validate_cross_references(pkg)
        for b in report.get("broken", []):
            add(
                "cross_references",
                "warning",
                f"{b['field']} field targets missing bookmark "
                f"{b['bookmark']!r} (renders as an error in Word)",
                b.get("part"),
            )
        info["cross_references"] = {
            "ref_fields": report.get("ref_fields"),
            "broken": len(report.get("broken", [])),
            "unreferenced_bookmarks": len(report.get("unreferenced_bookmarks", [])),
        }

    # --------------------------------------------------------- size profile
    def check_sizes():
        sizes = {name: len(pkg.raw_part(name)) for name in pkg.part_names()}
        largest = sorted(sizes.items(), key=lambda kv: -kv[1])[:5]
        info["size_profile"] = {
            "part_count": len(sizes),
            "total_bytes": sum(sizes.values()),
            "largest_parts": [{"part": n, "bytes": b} for n, b in largest],
        }

    run("content_types", check_content_types)
    run("relationships", check_relationships)
    run("orphan_parts", check_orphan_parts)
    run("fields", check_fields)
    run("notes", check_notes)
    run("styles", check_styles)
    run("numbering", check_numbering)
    run("sdt", check_sdt)
    run("table_merges", check_table_merges)
    run("bookmarks", check_bookmarks)
    run("revision_ids", check_revision_ids)
    run("images", check_images)
    run("cross_references", check_cross_references)
    run("size_profile", check_sizes)

    return {
        "ok": not any(p["severity"] == "error" for p in problems),
        "problems": problems,
        "info": info,
    }
