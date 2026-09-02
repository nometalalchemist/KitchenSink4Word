"""kitchensink4word v2: the consolidated FastMCP surface.

The Great Consolidation (V2_DESIGN.md): the 189-tool v1.6 surface rebuilt
as the consolidated v2 set. The ops/, com/, and core/ engine is unchanged;
this module is thin dispatch plus the v2 grammar:

- Positional tools take the Section 6 location object, resolved through
  core.locate (ambiguity is a loud refusal carrying every match).
- The Section 7 envelope applies AT THE MCP BOUNDARY: the registered tool
  wraps the module-level function, converting typed exceptions into
  structured {ok: false, error: {...}} refusals with isError=true, and
  middleware adds the success fields (ok, file) at serialization.
  In-process callers (the test suite convention: srv.tool(...)) get the
  raw function: v1-shaped returns, typed exceptions raise through.
- Every tool carries a pack tag (packs.py bookkeeping; Section 14.8).
  Pack VISIBILITY wiring (enable_tools/disable_tools, startup modes) is
  Phase 4; in Phase 2 the full surface is registered and enabled.

Contract for every mutating tool (v1.6 rule carried forward):
- Before each mutation the current content rotates into stable backup
  slots (prev/anchor) inside .ks4w-backups/ next to the file, unless
  backup=False (manage_backups lists/restores/purges them).
- Mutations of one file are serialized; saves are atomic and validated.
- A file open in Word is edited live by dual-mode tools (live='auto');
  tools with no live route refuse until it is closed.
"""

from __future__ import annotations

import functools
import inspect as _inspect
import json as _json
from typing import Any

from fastmcp import FastMCP
from fastmcp.server.context import Context as _Context
from fastmcp.server.middleware import Middleware as _FmcpMiddleware
from fastmcp.server.transforms.visibility import Visibility as _Visibility
from fastmcp.tools.function_tool import FunctionTool as _FunctionTool
from fastmcp.tools.tool import ToolResult as _FmcpToolResult

from . import envelope as _envelope
from . import packs as _packs
from .core.errors import (
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from .core.locate import is_range_spec, resolve_location, resolve_range
from .core.package import DocxPackage, qn
from .ops import (
    batch as _batch,
    bibliography as _bib,
    citecheck as _cc,
    comments as _cm,
    fields as _fl,
    furniture as _fu,
    lists as _ls,
    media as _md,
    notes as _nt,
    protection as _pr,
    read as _rd,
    structure as _sx,
    template as _tp,
    revisions as _rv,
    stats as _st,
    tables as _tb,
    text as _tx,
    toc as _tc,
    accessibility as _ac,
    chapterheaders as _ch,
    cleanup as _cu,
    compliance as _cp,
    definedterms as _dt,
    diagnostics as _dg,
    forms as _fm,
    frontmatter as _fma,
    integrity as _ig,
    mailmerge as _mm,
    redaction as _rx,
    reffields as _rf,
    anonymize as _an,
    assembly as _asm,
    backups as _bk,
    charts as _charts,
    citesystem as _cs,
    dataio as _dio,
    workflows as _wf,
    equations as _eq,
    journalcount as _jc,
    preview as _pv,
    reviewcycle as _rc,
    styleconvert as _sc2,
    stylefind as _sf,
    textboxes as _tbx,
    view as _view,
    zoterolib as _zl,
)

mcp = FastMCP(
    "kitchensink4word",
    instructions=(
        "Full-featured Word (.docx) editor, consolidated v2 surface: text, "
        "tables, footnotes/endnotes, TOC and generated lists, "
        "headers/footers, images, charts, equations, citations, comments, "
        "tracked changes, document assembly, protection, and the Word "
        "application tier (com_/live_ tools). Positional tools take one "
        "location object (paragraph | after_heading | outline | bookmark | "
        "search | anchor | cursor, plus a position modifier); ambiguous "
        "text matches refuse loudly with every candidate. File-based with "
        "auto-backup before every mutation; dual-mode tools edit documents "
        "open in Word live (live='auto'), tools with no live route refuse "
        "until the file is closed. list_elements enumerates any "
        "collection; validate runs any read-only check battery; "
        "migration/v1_to_v2.json maps every v1 tool name here."
    ),
)


# ------------------------------------------------------- boundary envelope


class _SuccessEnvelope(_FmcpMiddleware):
    """Section 7.1 success fields at the MCP boundary: object results gain
    ok (and file, when the call named one). In-process calls bypass this
    entirely (module attributes are the raw functions), so the test suite
    keeps v1-shaped returns. List-shaped reader payloads ride under
    fastmcp's own 'result' key and gain the same ok/file siblings."""

    _FILE_KEYS = (
        "file_path", "target_path", "template_path", "pdf_path", "old_path",
    )

    async def on_call_tool(self, context, call_next):
        result = await call_next(context)
        sc = getattr(result, "structured_content", None)
        if isinstance(sc, dict) and "ok" not in sc:
            args = getattr(context.message, "arguments", None) or {}
            out: dict[str, Any] = {"ok": True}
            for key in self._FILE_KEYS:
                value = args.get(key)
                if isinstance(value, str):
                    out["file"] = value
                    break
            out.update(sc)
            return _FmcpToolResult(
                content=_json.dumps(out, indent=2, ensure_ascii=False,
                                    default=str),
                structured_content=out,
            )
        return result


mcp.add_middleware(_SuccessEnvelope())
mcp.add_middleware(_envelope.DisabledToolSignpost())


def _tool(pack: str):
    """Register a v2 tool: the MCP-registered callable is a boundary
    wrapper (typed exceptions -> structured RefusalResult with
    isError=true, the Phase 1 refusal architecture); the module attribute
    stays the raw function so in-process returns stay v1-shaped. Every
    tool lands in the packs registry under its Section 14.8 pack tag
    ('lite' = the always-on core)."""

    def deco(fn):
        if _inspect.iscoroutinefunction(fn):
            @functools.wraps(fn)
            async def boundary(*args, **kwargs):
                try:
                    return await fn(*args, **kwargs)
                except _envelope.CATCHABLE as exc:
                    return _envelope.refuse(exc)
        else:
            @functools.wraps(fn)
            def boundary(*args, **kwargs):
                try:
                    return fn(*args, **kwargs)
                except _envelope.CATCHABLE as exc:
                    return _envelope.refuse(exc)

        tool = _FunctionTool.from_function(boundary)
        mcp.add_tool(tool)
        _packs.register(fn.__name__, None if pack == "lite" else pack, tool)
        return fn

    return deco


# ------------------------------------------------------------ edit plumbing


def _edit(file_path: str, fn, *, backup: bool = True) -> dict:
    from .core.safesave import write_lock

    # Serialize the full read-modify-validate-save cycle per file: concurrent
    # mutations of one document otherwise race (stale response metadata, lost
    # updates). Covers threads in this process AND other server processes.
    with write_lock(file_path):
        pkg = DocxPackage(file_path)
        result = fn(pkg)
        saved = pkg.save(do_backup=backup)
    result["saved"] = str(saved)
    return result


def _route_live(live: str, file_call, live_call) -> dict:
    """live='auto': file-based first, live layer when the doc is open in Word.
    'force': straight to the live layer. 'off': locked = refuse."""
    from .core.errors import DocumentLocked

    if live not in ("auto", "force", "off"):
        raise ValueError(f"live must be auto|force|off, got {live!r}")
    if live == "force":
        return live_call()
    try:
        return file_call()
    except DocumentLocked:
        if live == "off":
            raise
        return live_call()


# ------------------------------------------------------- location adapters


def _last_paragraph_index(pkg: DocxPackage) -> int:
    n = len(pkg.body().findall(qn("w:p")))
    if n == 0:
        raise TargetNotFound("document has no paragraphs")
    return n - 1


def _loc_insert_args(pkg: DocxPackage, location) -> tuple[int | None, bool]:
    """location object -> (after_index, at_end) for v1-shaped inserter ops
    (the Wave C/D document-end convention: location omitted = document
    end; a resolved position 'end' also lands at the end)."""
    if location is None:
        return None, True
    r = resolve_location(pkg, location)
    if r.position == "end":
        return None, True
    if r.position == "after":
        return r.paragraph_index, False
    if r.position == "before":
        if r.paragraph_index == 0:
            raise WordMcpError(
                "position 'before' the first paragraph is not supported "
                "by this inserter; use position 'after' or 'end'"
            )
        return r.paragraph_index - 1, False
    raise WordMcpError(
        f"position {r.position!r} is not meaningful for an inserter; "
        "use before | after | end"
    )


def _anchor_kwargs(pkg: DocxPackage, location) -> dict:
    """Section 6.1 location -> {anchor_text, occurrence} for the
    run-anchored ops inserters (notes, comments, citations, fields, index
    entries): the Wave B shared adapter.

    A search selector carrying only text/occurrence passes straight
    through: that is the exact v1 semantics (runmap-aware fragmented-run
    matching inside ops), the parity-critical path. Any other selector
    resolves through locate and anchors on the resolved paragraph's full
    text. These inserters place content right after the matched text, so
    only position 'after' (the default) is meaningful."""
    if not isinstance(location, dict):
        raise WordMcpError(
            "location must be an object with exactly one selector key "
            "(paragraph, after_heading, outline, bookmark, search, anchor, "
            "cursor) plus optional 'position'"
        )
    if location.get("position", "after") != "after":
        raise WordMcpError(
            "this inserter anchors content right after the matched text; "
            "only position 'after' (the default) is supported"
        )
    search = location.get("search")
    if (
        isinstance(search, dict)
        and set(location) <= {"search", "position"}
        and set(search) <= {"text", "occurrence"}
    ):
        if "text" not in search or not isinstance(search["text"], str):
            raise WordMcpError(
                'search selector requires text, like {"search": {"text": '
                '"specific text", "occurrence": 1}}'
            )
        occurrence = search.get("occurrence")
        if occurrence is None:
            occurrence = 1
        return {"anchor_text": search["text"], "occurrence": occurrence}
    r = resolve_location(pkg, location)
    paras = pkg.body().findall(qn("w:p"))
    if r.paragraph_index >= len(paras):
        raise TargetNotFound(
            "the resolved location has no materialized paragraph to anchor "
            "on; insert text first"
        )
    text = _rd.paragraph_text(paras[r.paragraph_index])
    if not text.strip():
        raise WordMcpError(
            "the resolved paragraph is empty; a run anchor needs text to "
            "attach to. Use a search selector instead."
        )
    occurrence = 1
    for el in paras[: r.paragraph_index]:
        if _rd.paragraph_text(el) == text:
            occurrence += 1
    return {"anchor_text": text, "occurrence": occurrence}


def _expand_range(pkg: DocxPackage, range_spec, *, what: str = "range"):
    """Dual-shape range (Wave E convention 2): {start, end} of ints (index
    shorthand) or location objects -> (start_index, end_index)."""
    if not isinstance(range_spec, dict):
        raise WordMcpError(
            f"{what} takes {{'start': ..., 'end': ...}} with ints or "
            "location objects as endpoints"
        )
    spec = dict(range_spec)
    if is_range_spec(spec):
        for key in ("start", "end"):
            value = spec.get(key)
            if isinstance(value, int) and not isinstance(value, bool):
                spec[key] = {"paragraph": value}
    r = resolve_range(pkg, spec)
    return r.start_index, r.end_index


def _live_cursor_reader(file_path: str):
    """cursor selector support on live branches: reads the 0-based body
    paragraph index of Selection.Range.Start once; never writes the
    Selection (the live-layer rule)."""

    def _read() -> int:
        from .com import live as _lv
        from .com import live_ops as _lo

        def probe(s):
            paras = _lo._body_paragraphs(s.doc)
            sel_start = s.app.Selection.Range.Start
            idx = 0
            for i, p in enumerate(paras):
                if p.Range.Start <= sel_start:
                    idx = i
                else:
                    break
            return {"index": idx}

        return _lv.run_live(file_path, "read cursor", probe)["index"]

    return _read


def _disk_snapshot_pkg(file_path: str) -> DocxPackage:
    """Read-only package snapshot of a possibly Word-locked file, for
    live-branch location resolution ONLY (never saved). DocxPackage
    refuses locked files by design, but Word grants shared reads, so a
    temp copy is taken, loaded fully into memory, and deleted."""
    import os
    import shutil
    import tempfile

    fd, tmp = tempfile.mkstemp(suffix=".ks4w-locsnap.docx")
    os.close(fd)
    try:
        shutil.copy2(file_path, tmp)
        return DocxPackage(tmp)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def _resolve_for_live(file_path: str, location):
    """Resolve a location on a live branch: against the DISK copy (stale
    while Word holds unsaved changes, the documented Wave E caveat), with
    a live cursor reader."""
    return resolve_location(
        _disk_snapshot_pkg(file_path), location,
        cursor_reader=_live_cursor_reader(file_path),
    )


# Selectors whose resolution depends on document CONTENT: on a live branch
# they resolve against the disk snapshot, so the resolved index must be
# re-verified at the COM boundary (live_ops._stale_guard) in case Word
# holds unsaved changes. paragraph/cursor addressing carries no snapshot
# dependency and keeps the v1 index-trust contract.
_TEXT_SELECTORS = frozenset({"search", "after_heading", "outline", "anchor"})


def _snapshot_para_texts(pkg: DocxPackage) -> list[str]:
    """Body-paragraph plain texts of a snapshot, index-aligned with the
    live layer's _body_paragraphs."""
    return [
        _rd.paragraph_text(el)
        for kind, _i, el in _rd.body_items(pkg)
        if kind == "paragraph"
    ]


def _snapshot_verify_text(pkg: DocxPackage, r) -> str | None:
    """The snapshot text the staleness guard verifies against, or None
    when the selector did not depend on snapshot content."""
    if r.selector not in _TEXT_SELECTORS:
        return None
    texts = _snapshot_para_texts(pkg)
    if 0 <= r.paragraph_index < len(texts):
        return texts[r.paragraph_index]
    return None


def _resolve_for_live_verified(file_path: str, location):
    """(resolved, verify_text) for a live branch: verify_text is the disk
    snapshot's text at the resolved paragraph when the selector was
    content-dependent, for verify-or-refuse at the COM boundary."""
    pkg = _disk_snapshot_pkg(file_path)
    r = resolve_location(
        pkg, location, cursor_reader=_live_cursor_reader(file_path)
    )
    return r, _snapshot_verify_text(pkg, r)


def _expand_range_verified(pkg: DocxPackage, range_spec, *,
                           what: str = "range"):
    """_expand_range plus per-endpoint verify texts, for live branches."""
    if not isinstance(range_spec, dict):
        raise WordMcpError(
            f"{what} takes {{'start': ..., 'end': ...}} with ints or "
            "location objects as endpoints"
        )
    spec = dict(range_spec)
    if is_range_spec(spec):
        for key in ("start", "end"):
            value = spec.get(key)
            if isinstance(value, int) and not isinstance(value, bool):
                spec[key] = {"paragraph": value}
    r = resolve_range(pkg, spec)
    return (r, _snapshot_verify_text(pkg, r.start),
            _snapshot_verify_text(pkg, r.end))


# ===================================================== 2.1 document lifecycle


@_tool("lite")
def create_document(file_path: str, title: str | None = None) -> dict:
    """Create a new blank .docx file, optionally setting the title core
    property. Refuses to overwrite an existing file (use copy_document with
    overwrite=True for that). Parent directories are created automatically.
    Populate the document afterward with insert_paragraphs, create_table,
    define_style, and other tools.
    Template-driven builds live in the assembly pack.
    """
    from pathlib import Path

    from docx import Document

    from .core.sandbox import check_path

    p = Path(check_path(file_path, "create document"))
    if p.exists():
        raise FileExistsError(f"{file_path} already exists")
    doc = Document()
    if title:
        doc.core_properties.title = title
    p.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(p))
    return {"created": str(p)}


@_tool("lite")
def copy_document(
    file_path: str, dest_path: str, overwrite: bool = False
) -> dict:
    """Copy a document byte-for-byte, e.g. to a new DTG-stamped filename
    before editing (manage_backups action='snapshot' names such a copy for
    you; this tool takes an explicit dest_path). Refuses an existing
    dest_path unless overwrite=True; an overwritten destination's previous
    content rotates into its .ks4w-backups prev slot first, so the
    overwrite is undoable via manage_backups restore.
    Split/merge and multi-document work live in the assembly pack.
    """
    import shutil
    from pathlib import Path

    from .core import safesave
    from .core.sandbox import check_path

    file_path = check_path(file_path, "copy source (read)")
    dest_path = check_path(dest_path, "copy destination")
    dest = Path(dest_path)
    overwrote = False
    if dest.exists():
        if not overwrite:
            raise FileExistsError(
                f"{dest_path} already exists (pass overwrite=True to "
                "replace it; its current content will be kept in the "
                ".ks4w-backups prev slot)"
            )
        with safesave.write_lock(dest_path):
            safesave.rotate_slots(dest_path)
            # prev.docx may be a hardlink to dest: never write into dest
            # in place (it would write through the link and corrupt the
            # backup); copy aside, then atomically replace.
            tmp_copy = dest.with_name(dest.name + ".ks4w-copy-tmp")
            shutil.copy2(file_path, tmp_copy)
            safesave.replace_with_retry(tmp_copy, dest_path)
        overwrote = True
    else:
        shutil.copy2(file_path, dest_path)
    return {"copied_to": dest_path, "overwrote_existing": overwrote}


@_tool("assembly")
def split_document(
    file_path: str,
    output_dir: str,
    level: int = 1,
    filename_from: str = "heading",
) -> dict:
    """Split a document into one standalone .docx per heading section, the
    inverse of com_multi_document merge (no Word needed). Sections start at
    each heading of `level` or higher and run to the next; content before
    the first heading becomes 00_front_matter.docx when non-empty. Outputs
    carry styles, numbering, fonts, and themes; unused note parts are
    stripped; every output is round-trip validated. filename_from: 'heading'
    or 'index'. The source is never modified; existing output files are
    refused.
    """
    return _dio.split_document(
        file_path, output_dir, level=level, filename_from=filename_from
    )


@_tool("assembly")
def insert_document(
    target_path: str,
    source_path: str,
    location: dict | None = None,
    formatting: str = "source",
    backup: bool = True,
) -> dict:
    """Insert the ENTIRE body of source_path into target_path at one position,
    with full resource reconciliation: the document-assembly tool for
    merging chapter files into one manuscript. Position via the standard
    location object (paragraph, after_heading, outline, bookmark, search,
    anchor); omit location to append at the end. Text selectors refuse on
    multiple matches and list every location; prefer outline or paragraph.
    Carried: tables, images, charts, hyperlinks, lists (fresh numbering),
    footnotes and endnotes (new ids), bookmarks (remapped, collisions
    renamed and reported), tracked changes, equations. Styles reconcile BY
    NAME (the target wins on a match; unmatched styles are cloned in with
    dependency chains). The source's section setup is never carried;
    mid-content section breaks and comment references are stripped and
    reported. OLE objects, ActiveX, subdocuments, and altChunks refuse the
    whole insertion; nothing is half-applied. formatting mirrors Word's
    paste modes: 'source' keeps direct formatting; 'merge' keeps emphasis
    but strips font/size/color/spacing/indent overrides; 'destination'
    strips all but structural properties. Returns per-resource counts,
    style remaps, bookmark renames, and the occupied range. The source file
    is never modified. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        if location is None:
            return _asm.insert_document(
                pkg, source_path, at_end=True, formatting=formatting
            )
        r = resolve_location(pkg, location)
        if r.position not in ("before", "after"):
            raise WordMcpError(
                f"insert_document supports position 'before' and 'after' "
                f"(default); got {r.position!r}. Omit location to append "
                "at the document end."
            )
        # The assembly op takes a body ITEM index (paragraphs AND tables
        # counted together in document order); convert the resolved
        # paragraph index (paragraph-only space) to it.
        item_index = None
        for i, (kind, idx, _el) in enumerate(_rd.body_items(pkg)):
            if kind == "paragraph" and idx == r.paragraph_index:
                item_index = i
                break
        if item_index is None:
            # Implicit paragraph of a fresh document: appending is the
            # only meaningful placement.
            return _asm.insert_document(
                pkg, source_path, at_end=True, formatting=formatting
            )
        if r.position == "before":
            if item_index == 0:
                raise WordMcpError(
                    "position 'before' the first body item is not "
                    "supported; use position 'after' or omit location"
                )
            item_index -= 1
        return _asm.insert_document(
            pkg, source_path, after_index=item_index, formatting=formatting
        )

    return _edit(target_path, _do, backup=backup)


@_tool("lite")
def manage_backups(
    action: str,
    file_path: str | None = None,
    directory: str | None = None,
    source: str | None = None,
    scope: str | None = None,
    dry_run: bool = True,
    label: str | None = None,
    dest_dir: str | None = None,
) -> dict:
    """Manage the automatic backups under the hidden .ks4w-backups/ folder next
    to each mutated document: two stable slots per document, prev (state
    before the most recent mutation) and anchor (session start).
    action='list': slot files with sizes and mtimes, legacy *.bak-* files,
    and orphaned slot folders whose source document is gone; give file_path
    for one document or directory for a folder. action='restore': overwrite
    file_path with a backup; source is 'prev', 'anchor', or a legacy *.bak-*
    path. The current content rotates into prev FIRST, so a restore is
    itself undoable; the payload is validated before the atomic replace, and
    documents open in Word are refused. action='purge': delete backups;
    scope: 'legacy', 'orphans', or 'slots'. dry_run defaults to TRUE (report
    only); dry_run=False deletes. Exact paths and sizes are reported either
    way. action='snapshot': save a DTG-stamped permanent copy of file_path,
    YYYYMMDD_HHMM_<name>.docx (an existing leading DTG is replaced, not
    stacked), optional short label suffix, optional dest_dir. Snapshots are
    permanent keepers: the slots rotate on every mutation, snapshots are
    never auto-pruned and no purge scope touches them. Never overwrites
    (collisions get a numeric suffix); the source document is not modified.
    """
    if action == "snapshot":
        if file_path is None:
            raise WordMcpError("action='snapshot' requires file_path")
        return _bk.create_snapshot(file_path, label=label, dest_dir=dest_dir)
    if label is not None or dest_dir is not None:
        raise WordMcpError(
            "label and dest_dir apply to action='snapshot' only"
        )
    return _bk.manage_backups(
        action,
        file_path=file_path,
        directory=directory,
        source=source,
        scope=scope,
        dry_run=dry_run,
    )


@_tool("academic")
def set_document_properties(
    file_path: str,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    comments: str | None = None,
    update_fields_on_open: bool | None = None,
    backup: bool = True,
) -> dict:
    """Set core document metadata (File > Info): title, author, subject,
    keywords, category, comments; only the parameters given change.
    update_fields_on_open=True additionally sets Word's one-shot 'update all
    fields on next open' flag (Word clears it after updating);
    update_fields_on_open=False clears it. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated save.
    Refuses documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        out = _sx.set_document_properties(
            pkg, title=title, author=author, subject=subject,
            keywords=keywords, category=category, comments=comments,
        )
        if update_fields_on_open is not None:
            flag = _tc.set_update_fields_flag(pkg, update_fields_on_open)
            for key, value in flag.items():
                out[key if key not in out else f"update_fields_{key}"] = value
        return out

    return _edit(file_path, _do, backup=backup)


@_tool("lite")
def get_document_info(file_path: str, live: str = "auto") -> dict:
    """Read a one-call document overview: paragraph/table/footnote/comment/
    revision counts, sections, and package parts. Documents open in Word are
    read live (same key names; live adds 'words' from Word's own
    ComputeStatistics counter plus track_revisions, and omits the part
    list). Read-only.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_document_info(DocxPackage(file_path)),
        lambda: _lo.get_document_info(file_path),
    )


@_tool("lite")
def diagnose_document(file_path: str) -> dict:
    """Produce a one-call structural health report, read-only: content-type
    coverage, dangling relationships and orphan parts, field balance per
    story part, footnote/endnote integrity, references to undefined styles
    and numbering, content-control and bookmark sanity, duplicate revision
    ids, missing image targets, broken cross-references, and a per-part size
    profile. Never fails on a weird-but-openable document; every check
    degrades to a reported problem, and healthy=false only for problems
    that render broken or lose content in Word. The deep companion to
    validate(checks=['core']). No live mode BY DESIGN: this reads the saved
    package's XML, stale while Word holds unsaved changes. Close the
    document first, or use com_validate_opens_clean / live
    get_document_info.
    The full validate check battery lives in the academic pack.
    """
    from .core.errors import DocumentLocked

    try:
        out = _dg.diagnose_document(DocxPackage(file_path))
        # Re-key the v1 domain flag: the envelope's top-level ok means the
        # CALL succeeded, so an unhealthy document must not answer
        # ok:false (adversarial 6a). Ops shape unchanged for v1 file mode.
        return {"healthy": out.pop("ok"), **out}
    except DocumentLocked as exc:
        raise DocumentLocked(
            f"{file_path} is open in Word (or locked). diagnose_document "
            "reads the saved file's XML, which is stale while Word holds "
            "unsaved changes, so there is no live route by design. Close "
            "the document (com_save_document with close=True saves and "
            "closes) and rerun; for open-document checks use "
            "com_validate_opens_clean or get_document_info (live)."
        ) from exc


# ==================================================== 2.3 reading and search


@_tool("lite")
def get_text(
    file_path: str,
    start: int = 0,
    end: int | None = None,
    contains: str | None = None,
    include_textboxes: bool = False,
    textbox: bool | dict | None = None,
    live: str = "auto",
) -> list | dict:
    """Read body paragraphs as [{index, text, style, ...}] with each
    paragraph's effective style. start/end slice by paragraph index
    (0-based, end EXCLUSIVE); contains filters. include_textboxes=True
    appends text-box content as labeled extra entries without touching body
    indices; textbox=true (or {"index": n}) returns ONLY text-box content,
    with the box_index set_textbox_text takes; both file-mode only.
    Documents open in Word are read live, same body shape. Equations are
    read via list_elements type='equations'. Read-only.
    """
    from .com import live_ops as _lo

    if textbox is not None and textbox is not False:
        if start != 0 or end is not None or contains is not None \
                or include_textboxes:
            raise WordMcpError(
                "textbox mode is exclusive with start/end/contains/"
                "include_textboxes: it returns ONLY text-box content"
            )
        if live == "force":
            raise WordMcpError(
                "textbox mode is file-mode only; close the document in "
                "Word first"
            )
        if textbox is not True and not isinstance(textbox, dict):
            raise WordMcpError(
                'textbox takes true (all boxes) or {"index": n} (one box)'
            )
        res = _tbx.get_textbox_text(DocxPackage(file_path))
        if isinstance(textbox, dict):
            unknown = sorted(set(textbox) - {"index"})
            if unknown:
                raise WordMcpError(
                    f"textbox got unknown key(s) {unknown}; it takes "
                    '{"index": n}'
                )
            wanted = textbox.get("index")
            boxes = [
                b for b in res.get("boxes", [])
                if b.get("box_index") == wanted
            ]
            if not boxes:
                raise TargetNotFound(
                    f"no text box with index {wanted!r} "
                    f"({res.get('count', 0)} box(es) in the document)"
                )
            res = {**res, "boxes": boxes, "count": len(boxes)}
        return res

    def _live_call():
        if include_textboxes:
            raise WordMcpError(
                "include_textboxes is file-mode only; close the document "
                "in Word, use get_text(textbox=true), or pass "
                "include_textboxes=False"
            )
        return _lo.get_text(file_path, start, end, contains)

    return _route_live(
        live,
        lambda: _rd.get_paragraphs(
            DocxPackage(file_path), start, end, contains=contains,
            include_textboxes=include_textboxes,
        ),
        _live_call,
    )


@_tool("lite")
def find_text(
    file_path: str,
    query: str | None = None,
    regex: bool = False,
    context_chars: int = 60,
    include_textboxes: bool = False,
    formatting: dict | None = None,
    scope: str = "body",
    live: str = "auto",
) -> list | dict:
    """Find text in paragraphs and table cells; returns locations plus
    context. include_textboxes=True also searches text-box content as
    labeled matches; file-mode only. Pass formatting={...} (bold, font,
    size_pt, color, highlight, style) to search by EFFECTIVE formatting:
    resolved through run, character style, style chain, and defaults;
    query=None returns every formatted stretch; scope 'body' or 'all';
    file-mode, no regex. Plain queries search open documents live (same
    shape, 500-match cap). Read-only.
    """
    from .com import live_ops as _lo

    if formatting is not None:
        if regex:
            raise WordMcpError(
                "formatting mode has no regex support (v1 find_formatted "
                "parity); drop regex or drop formatting"
            )
        if include_textboxes:
            raise WordMcpError(
                "formatting mode never searches text boxes; drop "
                "include_textboxes"
            )
        if live == "force":
            raise WordMcpError(
                "formatting mode is file-mode only; close the document in "
                "Word first"
            )
        return _sf.find_formatted(
            DocxPackage(file_path), query, formatting=formatting, scope=scope
        )
    if query is None:
        raise WordMcpError(
            "query is required (it is optional only with formatting={...})"
        )

    def _live_call():
        if include_textboxes:
            raise WordMcpError(
                "include_textboxes is file-mode only; close the document "
                "in Word, use get_text(textbox=true), or pass "
                "include_textboxes=False"
            )
        return _lo.find_text(file_path, query, regex, context_chars)

    return _route_live(
        live,
        lambda: _rd.find_text(
            DocxPackage(file_path), query, regex=regex,
            context_chars=context_chars, include_textboxes=include_textboxes,
        ),
        _live_call,
    )


@_tool("lite")
def get_outline(file_path: str, live: str = "auto") -> list | dict:
    """List every heading with its body paragraph index and level. Detects both
    heading systems: built-in Heading styles AND w:outlineLvl overrides
    (direct or style-inherited, the academic-template pattern on
    Normal-styled paragraphs); detected_via on each entry names which. The
    paragraph indices and outline numbers feed the location object's
    paragraph and outline selectors. Documents open in Word are read live
    (same flat-list shape). Read-only.
    TOC generation and heading surgery live in the academic pack.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_outline(DocxPackage(file_path)),
        lambda: _lo.get_outline(file_path),
    )


@_tool("academic")
def get_paragraph_format(
    file_path: str, start: int, end: int | None = None
) -> dict:
    """Read EFFECTIVE paragraph formatting for body paragraphs start..end
    (inclusive; end defaults to start): alignment, spacing, line_spacing
    with its rule, indents (negative first-line = hanging), keep_with_next,
    widow_control, page_break_before. Resolution follows Word (own
    properties, then the style basedOn chain, then defaults); every property
    reports value AND source. Numbering-contributed indents are flagged,
    not resolved. set_paragraph_format takes the same keys. Read-only.
    """
    return _sf.get_paragraph_format(DocxPackage(file_path), start, end)


@_tool("academic")
def get_styles(file_path: str) -> list:
    """List the styles defined in the document: id, name, type, based_on, plus
    each style's EXPLICITLY defined paragraph_formatting and
    character_formatting in the exact shape define_style accepts (template
    cloning is one read plus one define). Inherited values are not
    synthesized; follow the based_on chain. Exact/atLeast line spacing and
    theme font or color references have no define_style representation and
    are omitted. Read-only.
    """
    return _rd.list_styles(DocxPackage(file_path))


@_tool("academic")
def word_count(
    file_path: str,
    by_section: bool = True,
    exclusions: list[str] | None = None,
    live: str = "auto",
) -> dict:
    """Count words, characters, and paragraphs, total and per heading section.
    Documents open in Word are counted live via Word's own statistics
    engine; file mode counts whitespace tokens, so the two can differ
    slightly. Pass exclusions=[...] (references, captions, footnotes,
    endnotes, block_quotes, tables, headings, front_matter, abstract) for
    the number a journal wants: returns total, per-zone excluded breakdown,
    and the included count; exclusions mode is file-mode only. Read-only.
    """
    from .com import live_ops as _lo

    if exclusions is not None:
        if live == "force":
            raise WordMcpError(
                "exclusions mode is file-mode only; close the document in "
                "Word first"
            )
        return _jc.word_count_with_exclusions(
            DocxPackage(file_path), exclude=tuple(exclusions)
        )
    return _route_live(
        live,
        lambda: _st.word_count(DocxPackage(file_path), by_section=by_section),
        lambda: _lo.word_count(file_path, by_section=by_section),
    )


@_tool("academic")
def get_headers_footers(file_path: str) -> dict:
    """Every header and footer part across all sections: text content, whether
    it contains a PAGE-number field, and the part type (default, first page,
    even page). Use set_header_footer to write them and
    setup_chapter_headers for per-chapter running headers; page-number
    formatting lives in set_page_numbers. Read-only.
    """
    return _fu.get_headers_footers(DocxPackage(file_path))


@_tool("protection-io")
def get_protection(file_path: str) -> dict:
    """Current document protection state: mode (none, readOnly, comments,
    trackedChanges, forms), whether a password is set, and whether
    formatting restrictions are active. Use set_document_protection to
    enable protection or to disable it (setting protection off is the
    removal path in v2). Read-only.
    """
    return _pr.get_protection(DocxPackage(file_path))


# ------------------------------------------------------- 3.1 list_elements

# type -> {call, items (dict key holding the array, None = list return),
#          name (filter fn or None), drop (sibling keys dropped)}
_ELEMENT_TYPES: dict[str, dict] = {
    "tables": {"call": _rd.list_tables, "items": None, "name": None},
    "images": {"call": _md.list_images, "items": None, "name": None},
    "charts": {
        "call": _charts.list_charts, "items": None,
        "name": lambda it, s: s.lower() in (it.get("title") or "").lower(),
    },
    "equations": {
        "call": _eq.list_equations, "items": "equations", "name": None,
    },
    "bookmarks": {
        "call": _fl.list_bookmarks, "items": None,
        "name": lambda it, s: s.lower() in (it.get("name") or "").lower(),
    },
    "sources": {
        "call": _bib.list_sources, "items": None,
        "name": lambda it, s: (
            s.lower() in (it.get("tag") or "").lower()
            or s.lower() in (it.get("title") or "").lower()
        ),
    },
    "sections": {"call": _fu.list_sections, "items": None, "name": None},
    "section_blocks": {
        "call": _sx.list_section_blocks, "items": None,
        "name": lambda it, s: s.lower() in (it.get("heading") or "").lower(),
    },
    "footnotes": {"call": _rd.list_footnotes, "items": None, "name": None},
    "endnotes": {"call": _rd.list_endnotes, "items": None, "name": None},
    "fields": {
        "call": _fl.list_fields, "items": "fields",
        "name": lambda it, s: s.lower() in (it.get("type") or "").lower(),
    },
    "reference_fields": {
        "call": _rf.list_reference_fields, "items": "fields",
        "name": lambda it, s: s.lower() in (it.get("manager") or "").lower(),
    },
    "form_fields": {
        "call": _fm.list_form_fields, "items": "fields",
        "name": lambda it, s: s.lower() in (it.get("name") or "").lower(),
        "drop": ("count",),
    },
    "content_controls": {
        "call": _fm.list_content_controls, "items": "controls",
        "name": lambda it, s: (
            s.lower() in (it.get("tag") or "").lower()
            or s.lower() in (it.get("alias") or "").lower()
        ),
        "drop": ("count",),
    },
    "template_placeholders": {
        "call": _mm.list_template_placeholders, "items": "placeholders",
        "name": lambda it, s: s.lower() in (it.get("name") or "").lower(),
    },
    "index_entries": {
        "call": _fl.list_index_entries, "items": None,
        "name": lambda it, s: s.lower() in (it.get("entry") or "").lower(),
    },
    "lists": {"call": _ls.get_lists, "items": None, "name": None},
    "toc": {"call": _tc.read_toc, "items": "tocs", "name": None},
}


@_tool("lite")
def list_elements(file_path: str, type: str, filter: dict | None = None) -> dict:
    """Enumerate any collection in the document with one call. type: tables |
    images | charts | equations | bookmarks | sources | sections |
    section_blocks | footnotes | endnotes | fields | reference_fields |
    form_fields | content_controls | template_placeholders | index_entries |
    lists | toc. Returns {type, count, items}; each type keeps its v1 item
    shape, and the ids or indices returned are the handles the
    matching set_/delete_/manage_ tools accept (list then act). Highlights per type: tables reports dimensions and
    the table_index other table tools take; images reports display size and
    the media target; charts reports series; equations
    reads math content; bookmarks excludes internal TOC
    bookmarks; fields covers complex and simple fields with cached results;
    reference_fields inventories Zotero, EndNote, and Mendeley fields and
    flags broken pairs; form_fields and content_controls cover legacy fields
    and SDTs; template_placeholders lists {{name}} and MERGEFIELD keys for
    fill_template and mail_merge; lists groups list paragraphs by numbering
    instance; toc returns TOC-family fields with cached entries (refresh via
    com_refresh_fields, then re-read). filter={"range": {start, end},
    "name": "substring"} applies where meaningful; inapplicable filters
    refuse loudly. Read-only, file mode; close documents open in Word
    first.
    Tools acting on these elements live in packs; enable_tools lists
    them.
    """
    spec = _ELEMENT_TYPES.get(type)
    if spec is None:
        raise WordMcpError(
            f"unknown type {type!r}; one of: "
            + " | ".join(_ELEMENT_TYPES)
        )
    name_sub: str | None = None
    if filter is not None:
        if not isinstance(filter, dict):
            raise WordMcpError(
                'filter takes {"range": {"start": n, "end": n}, '
                '"name": "substring"}'
            )
        unknown = sorted(set(filter) - {"range", "name"})
        if unknown:
            raise WordMcpError(
                f"unknown filter key(s) {unknown}; filter takes 'range' "
                "and/or 'name'"
            )
        if "range" in filter:
            raise WordMcpError(
                "the range filter is RESERVED in v2.0.0: the enumeration "
                "item shapes carry no consistent paragraph attribution "
                "yet, so no type supports it. Filter client-side for now."
            )
        if "name" in filter:
            if spec["name"] is None:
                supported = sorted(
                    t for t, s in _ELEMENT_TYPES.items()
                    if s["name"] is not None
                )
                raise WordMcpError(
                    f"type {type!r} does not support the name filter; "
                    f"types that do: {supported}"
                )
            name_sub = str(filter["name"])
    raw = spec["call"](DocxPackage(file_path))
    siblings: dict = {}
    if isinstance(raw, dict):
        items = raw.get(spec["items"]) or []
        drop = set(spec.get("drop", ())) | {spec["items"]}
        siblings = {k: v for k, v in raw.items() if k not in drop}
    else:
        items = raw
    if name_sub is not None:
        items = [it for it in items if spec["name"](it, name_sub)]
    return {"type": type, "count": len(items), "items": items, **siblings}


# ------------------------------------------------ 2.12 validation & workflow


def _notes_ok(notes_report: dict) -> bool:
    """validate_notes carries per-kind ok flags (footnotes/endnotes), not a
    top-level one; both must hold."""
    return bool(
        notes_report["footnotes"]["ok"] and notes_report["endnotes"]["ok"]
    )


def _check_core(pkg: DocxPackage, _opts: dict) -> dict:
    notes_report = _nt.validate_notes(pkg)
    xml = pkg.raw_part("word/document.xml").decode("utf-8", errors="replace")
    begins = xml.count('w:fldCharType="begin"')
    ends = xml.count('w:fldCharType="end"')
    return {
        "package_ok": True,
        "notes": notes_report,
        "fields_balanced": begins == ends,
        "field_begins": begins,
        "field_ends": ends,
    }


def _opts_keys(check: str, opts: dict, allowed: set) -> None:
    unknown = sorted(set(opts) - allowed)
    if unknown:
        raise WordMcpError(
            f"options.{check} got unknown key(s) {unknown}; it takes "
            f"{sorted(allowed)}"
        )


def _check_forms(pkg, opts):
    _opts_keys("forms", opts, {"required"})
    return _fm.validate_form_completeness(pkg, required=opts.get("required"))


def _check_defined_terms(pkg, opts):
    _opts_keys("defined_terms", opts, {"definition_patterns"})
    return _dt.check_defined_terms(
        pkg, definition_patterns=opts.get("definition_patterns")
    )


def _check_brand(pkg, opts):
    if not opts:
        raise WordMcpError(
            "checks=['brand'] requires options.brand: the rules dict "
            '(e.g. {"fonts": {"allowed": [...]}, '
            '"colors": {"allowed_hex": [...]}})'
        )
    return _cp.check_brand_compliance(pkg, opts)


def _check_template(pkg, opts):
    if not opts:
        raise WordMcpError(
            "checks=['template'] requires options.template: the rules "
            "dict (page/fonts/line_spacing/headings/page_numbering/"
            "required_headings_in_order keys, the v1 "
            "check_template_compliance ruleset)"
        )
    return _cp.check_template_compliance(pkg, opts)


def _check_image_resolution(pkg, opts):
    _opts_keys("image_resolution", opts, {"min_dpi"})
    return _ac.check_image_resolution(
        pkg, min_dpi=opts.get("min_dpi", 300)
    )


def _check_redaction(pkg, opts):
    targets = opts.get("targets")
    if not targets:
        raise WordMcpError(
            "checks=['redaction'] requires options.redaction.targets: "
            "[{find | pattern...}] to re-scan for"
        )
    _opts_keys("redaction", opts, {"targets"})
    return _rx.verify_redaction(pkg, targets)


def _dt_passed(f: dict) -> bool:
    return not (
        f.get("defined_never_used")
        or f.get("used_never_defined")
        or f.get("defined_multiple_times")
        or f.get("first_use_before_definition")
    )


# check -> (runner(pkg, opts) -> findings, passed(findings) -> bool)
_VALIDATE_CHECKS: dict[str, tuple] = {
    "core": (_check_core,
             lambda f: bool(f["fields_balanced"]) and _notes_ok(f["notes"])),
    "captions": (lambda pkg, o: _ig.validate_captions(pkg),
                 lambda f: bool(f["ok"])),
    "chapter_headers": (lambda pkg, o: _ch.validate_chapter_headers(pkg),
                        lambda f: bool(f["ok"])),
    "cross_references": (lambda pkg, o: _ig.validate_cross_references(pkg),
                         lambda f: not f["broken"]),
    "notes": (lambda pkg, o: _nt.validate_notes(pkg), _notes_ok),
    "forms": (_check_forms, lambda f: bool(f["complete"])),
    "citation_parity": (lambda pkg, o: _cc.check_citation_parity(pkg),
                        lambda f: not f["missing_references"]),
    "defined_terms": (_check_defined_terms, _dt_passed),
    "brand": (_check_brand, lambda f: bool(f["compliant"])),
    "template": (_check_template, lambda f: bool(f["compliant"])),
    "reference_fields": (
        lambda pkg, o: _rf.check_reference_field_integrity(pkg),
        lambda f: bool(f["ok"]),
    ),
    "image_resolution": (_check_image_resolution,
                         lambda f: bool(f["pass"])),
    "accessibility": (lambda pkg, o: _ac.audit_accessibility(pkg),
                      lambda f: bool(f["summary"]["pass"])),
    "redaction": (_check_redaction, lambda f: bool(f["clean"])),
}


@_tool("academic")
def validate(
    file_path: str,
    checks: list[str] | None = None,
    options: dict | None = None,
) -> dict:
    """Run read-only correctness checks and return one report. checks (default
    ['core']): core (package opens, notes consistent, field markers
    balanced), captions (tables and images with SEQ captions, mixed
    numbering conventions), chapter_headers (STYLEREF coverage per section),
    cross_references (broken REF/PAGEREF targets, unreferenced bookmarks,
    unverified plain-text references), notes (footnote/endnote integrity and
    orphans), forms (unfilled or missing fields; options.forms.required),
    citation_parity (APA in-text vs reference list, both directions,
    heuristic), defined_terms (legal defined-terms audit, heuristic;
    options.defined_terms.definition_patterns), brand (options.brand ruleset
    with fonts and colors), template (options.template ruleset: margins,
    fonts, spacing, heading order, page numbering), reference_fields
    (Zotero/EndNote/Mendeley field integrity), image_resolution (effective
    print DPI; options.image_resolution.min_dpi, default 300), accessibility
    (headings, alt text, table headers, contrast, title, link text),
    redaction (options.redaction.targets re-scanned across every XML part).
    Returns {passed, results: {check: {passed, findings}}}, each check's
    v1 findings shape kept; passed=false means findings, not a failed
    call. Read-only, always: repairs live in
    fix_accessibility, manage_note, and redact_text. Reads the saved file;
    close documents open in Word first.
    """
    wanted: list[str] = []
    for check in (checks if checks is not None else ["core"]):
        if check not in _VALIDATE_CHECKS:
            raise WordMcpError(
                f"unknown check {check!r}; checks: "
                + " | ".join(_VALIDATE_CHECKS)
            )
        if check not in wanted:
            wanted.append(check)
    if not wanted:
        raise WordMcpError("checks must name at least one check")
    opts = options or {}
    if not isinstance(opts, dict):
        raise WordMcpError("options is a dict namespaced per check")
    stray = sorted(set(opts) - set(wanted))
    if stray:
        raise WordMcpError(
            f"options given for check(s) not in this call: {stray}; add "
            "them to checks or drop the options"
        )
    pkg = DocxPackage(file_path)
    results: dict[str, dict] = {}
    for check in wanted:
        run, passed = _VALIDATE_CHECKS[check]
        findings = run(pkg, opts.get(check) or {})
        results[check] = {"passed": bool(passed(findings)),
                          "findings": findings}
    # "passed", not "ok": the envelope's top-level ok means THE CALL
    # SUCCEEDED; a check battery that ran fine but found problems is a
    # successful call (adversarial 6a: a domain ok:false rode out looking
    # like a failed call with no error object).
    return {
        "passed": all(r["passed"] for r in results.values()),
        "results": results,
    }


@_tool("academic")
def fix_accessibility(
    file_path: str,
    alt_text_placeholders: bool = False,
    heading_skips: bool = False,
    heading_strategy: str = "promote",
    table_headers: bool = False,
    doc_title: bool = False,
    dry_run: bool = False,
    backup: bool = True,
) -> dict:
    """Repair validate(checks=['accessibility']) findings, one opted-in
    category at a time; all default OFF, none enabled refuses.
    alt_text_placeholders marks images with a labeled placeholder, never an
    invented description; heading_skips repairs skipped levels, refusing
    risky patterns; table_headers flags safe first rows; doc_title fills an
    EMPTY title. dry_run=True only reports. Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    if dry_run:
        return _ac.fix_accessibility(
            DocxPackage(file_path),
            alt_text_placeholders=alt_text_placeholders,
            heading_skips=heading_skips,
            heading_strategy=heading_strategy,
            table_headers=table_headers,
            doc_title=doc_title,
            dry_run=True,
        )
    return _edit(
        file_path,
        lambda pkg: _ac.fix_accessibility(
            pkg,
            alt_text_placeholders=alt_text_placeholders,
            heading_skips=heading_skips,
            heading_strategy=heading_strategy,
            table_headers=table_headers,
            doc_title=doc_title,
            dry_run=False,
        ),
        backup=backup,
    )


@_tool("academic")
def prepare_for_submission(
    file_path: str,
    accept_revisions: bool = True,
    remove_comments: bool = True,
    scrub_metadata: bool = True,
    keep_title: bool = True,
    backup: bool = True,
) -> dict:
    """Prepare a manuscript for submission in one call: accept all tracked
    changes (every author), delete all comments including the
    comments-family parts, and scrub identifying metadata (author,
    last-modified-by, company; title kept unless keep_title=False). Content
    is never removed: footnotes, fields, and citations all stay; protected
    documents are refused rather than half-cleaned. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _cu.prepare_for_submission(
            pkg,
            accept_revisions=accept_revisions,
            remove_comments=remove_comments,
            scrub_metadata=scrub_metadata,
            keep_title=keep_title,
        ),
        backup=backup,
    )


@_tool("academic")
def assemble_front_matter(
    file_path: str, spec: dict, backup: bool = True
) -> dict:
    """Assemble front matter in one call: the sequence is inserted at the
    START, page breaks separate the pages, one section break lets
    numbering switch (front lowerRoman, body decimal from 1). spec.sections
    kinds: title_page, blank_or_copyright, abstract, toc, list_of_figures,
    list_of_tables; spec.page_numbering overrides. Refuses existing front
    matter unless spec.force is true. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fma.assemble_front_matter(pkg, spec),
        backup=backup,
    )


@_tool("lite")
def get_workflows(task: str | None = None) -> dict:
    """Recommended tool sequences for common multi-step tasks, with a one-line
    why per step. Call with no task to list the available tasks
    ('process-feedback', 'prepare-submission', 'format-citations',
    'build-from-template', 'heavy-editing', 'migrate-from-v1', 'bulk-edit');
    call with task='<name>' for that task's step-by-step sequence and notes.
    Pure guidance: reads nothing, changes nothing.
    """
    return _wf.get_workflows(task)


# ============================================ 2.4 text and paragraph writing


def _outline_based_headings(pkg: DocxPackage) -> bool:
    """The NSU-template rule: the document has headings and EVERY one is
    outlineLvl-based (no Heading styles in play)."""
    heads = _rd.get_outline(pkg)
    return bool(heads) and all(
        h.get("detected_via") == "outline_level" for h in heads
    )


@_tool("lite")
def insert_paragraphs(
    file_path: str,
    paragraphs: list[dict],
    location: dict | None = None,
    inherit_format: bool = False,
    copy_format_from: int | None = None,
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Insert paragraphs (items {text, style?, formatting?, heading_level?})
    at a location object (omitted = document end). heading_level 1-9 makes
    the item a heading. inherit_format/copy_format_from clone neighbor
    formatting (file mode only); track records insertions by author.
    Auto-backup in file mode: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Documents
    open in Word are edited live; a stale text-selector target (unsaved
    changes moved it) refuses: save in Word, retry.
    """
    from .com import live_ops as _lo

    heading_items: list[tuple[int, int]] = []
    cleaned: list = []
    for pos, item in enumerate(paragraphs):
        if isinstance(item, dict) and "heading_level" in item:
            level = item["heading_level"]
            if "style" in item:
                raise WordMcpError(
                    "heading_level and style on the same item are "
                    "mutually exclusive: heading_level PICKS the style"
                )
            if isinstance(level, bool) or not isinstance(level, int) \
                    or not 1 <= level <= 9:
                raise WordMcpError("heading_level must be an int 1-9")
            heading_items.append((pos, level))
            cleaned.append(
                {k: v for k, v in item.items() if k != "heading_level"}
            )
        else:
            cleaned.append(item)

    def _positioning(pkg_or_none, resolver):
        after_index = before_index = None
        at_end = False
        if location is None:
            at_end = True
        else:
            r = resolver(location)
            if r.position == "end":
                at_end = True
            elif r.position == "after":
                after_index = r.paragraph_index
            elif r.position in ("before", "start"):
                before_index = r.paragraph_index
            else:
                raise WordMcpError(
                    "position 'replace' is not an insertion; use "
                    "set_paragraph_text"
                )
        return after_index, before_index, at_end

    def _file_call() -> dict:
        def _do(pkg: DocxPackage) -> dict:
            after_index, before_index, at_end = _positioning(
                pkg, lambda loc: resolve_location(pkg, loc)
            )
            items = list(cleaned)
            outline_mode = bool(heading_items) and _outline_based_headings(pkg)
            if heading_items and not outline_mode:
                for pos, level in heading_items:
                    item = dict(items[pos])
                    item["style"] = _tx.ensure_heading_style(pkg, level)
                    items[pos] = item
            paras_before = len(pkg.body().findall(qn("w:p")))
            res = _tx.insert_paragraphs(
                pkg, items,
                after_index=after_index, before_index=before_index,
                at_end=at_end, inherit_format=inherit_format,
                copy_format_from=copy_format_from, track=track,
                author=author,
            )
            if heading_items and outline_mode:
                if at_end:
                    first = paras_before
                elif after_index is not None:
                    first = after_index + 1
                else:
                    first = before_index
                for pos, level in heading_items:
                    _tx.set_paragraph_format(
                        pkg, [first + pos], {"outline_level": level - 1}
                    )
                res["outline_levels_set"] = len(heading_items)
            return res

        return _edit(file_path, _do, backup=backup)

    def _live_call() -> dict:
        if inherit_format or copy_format_from is not None:
            raise WordMcpError(
                "inherit_format/copy_format_from are file-mode features and "
                "this document is open in Word: close it in Word and retry, "
                "or insert without format cloning"
            )
        verify = None
        if location is None:
            after_index, before_index, at_end = None, None, True
        else:
            r, verify = _resolve_for_live_verified(file_path, location)
            after_index, before_index, at_end = _positioning(
                None, lambda loc: r
            )
        items = list(cleaned)
        if heading_items:
            # live headings: built-in Heading styles by numeric constant
            # (locale-safe), or direct outline levels on outline-based
            # documents, matching the file path's mode detection
            outline_mode = _outline_based_headings(
                _disk_snapshot_pkg(file_path)
            )
            key = "outline_heading" if outline_mode else "heading_level"
            for pos, level in heading_items:
                item = dict(items[pos])
                item[key] = level
                items[pos] = item
        return _lo.insert_paragraphs(
            file_path, items,
            after_index=after_index, before_index=before_index,
            at_end=at_end, track=track, author=author, verify_text=verify,
        )

    return _route_live(live, _file_call, _live_call)


@_tool("lite")
def delete_paragraphs(
    file_path: str,
    start: int | None = None,
    end: int | None = None,
    range: dict | None = None,
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Delete body paragraphs by 0-based inclusive index (start, end; end
    defaults to start) or by range={start, end} location objects. Refuses
    ranges that cut through a field or carry a section break; deleting
    every paragraph leaves one empty one behind. track records tracked
    deletions by author instead of removing. Auto-backup in file mode:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Documents open in Word are edited live, unsaved
    until the user saves.
    """
    from .com import live_ops as _lo

    if (start is None) == (range is None):
        raise WordMcpError(
            "give exactly one addressing form: start (with optional end) "
            "or range={start, end}"
        )
    if range is not None and end is not None:
        raise WordMcpError("end goes with start, not with range")

    def _file_call() -> dict:
        def _do(pkg: DocxPackage) -> dict:
            if range is not None:
                s, e = _expand_range(pkg, range)
            else:
                s, e = start, end
            return _tx.delete_paragraphs(pkg, s, e, track=track,
                                         author=author)

        return _edit(file_path, _do, backup=backup)

    def _live_call() -> dict:
        vs = ve = None
        if range is not None:
            rr, vs, ve = _expand_range_verified(
                _disk_snapshot_pkg(file_path), range
            )
            s, e = rr.start_index, rr.end_index
        else:
            s, e = start, end
        return _lo.delete_paragraphs(file_path, s, e, track=track,
                                     author=author, verify_start_text=vs,
                                     verify_end_text=ve)

    return _route_live(live, _file_call, _live_call)


@_tool("lite")
def set_paragraph_text(
    file_path: str,
    location: dict,
    new_text: str,
    expect: str | None = None,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Replace one paragraph's full text, keeping style and base formatting;
    address it with a location object ({paragraph: N}, {search: ...},
    {outline: ...}). Indices shift after edits, so pass expect (a substring
    the target must contain) to refuse instead of hitting the wrong
    paragraph; verify the returned replaced_text. Auto-backup in file mode:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Edits go live on open documents;
    tracked-revision paragraphs refuse live.
    """
    from .com import live_ops as _lo

    def _file_call() -> dict:
        def _do(pkg: DocxPackage) -> dict:
            idx = resolve_location(pkg, location).paragraph_index
            return _tx.replace_paragraph_text(pkg, idx, new_text,
                                              expect=expect)

        return _edit(file_path, _do, backup=backup)

    def _live_call() -> dict:
        r, verify = _resolve_for_live_verified(file_path, location)
        return _lo.replace_paragraph_text(file_path, r.paragraph_index,
                                          new_text, expect=expect,
                                          verify_text=verify)

    return _route_live(live, _file_call, _live_call)


@_tool("lite")
def search_and_replace(
    file_path: str,
    replacements: list[dict],
    scope: str = "body",
    preview: bool = False,
    find_formatting: dict | None = None,
    max_replacements: int | None = None,
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Batch find/replace, safe across Word's fragmented runs. Each item:
    {find, replace, regex?}; scope: body | footnotes | headers | all.
    preview=true dry-runs the same engine without touching the file:
    per-item counts, each match with paragraph index and before/after
    context, the grand total, and the refusals a real run would hit;
    review, then rerun with max_replacements set to the previewed total so
    drift aborts instead of over-replacing. find_formatting={bold, italic,
    font, size_pt, color, highlight, style, ...} restricts replacement to
    text carrying that effective formatting (file mode only, no regex; an
    item without find replaces each entire matching stretch, and the
    replacement keeps the matched formatting). max_replacements aborts,
    changing nothing, when total matches would exceed it. track records
    each replacement as a tracked change by author. Sibling:
    set_paragraph_text rewrites one whole paragraph when a find string
    would be unwieldy. Live edits appear immediately as one Ctrl+Z step,
    unsaved until the user saves; the live result adds live:true and skip
    counters, and literal finds beyond Word's ~255-char limit are handled
    automatically. Auto-backup in file mode: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Documents open in Word are edited live.
    """
    from .com import live_ops as _lo

    if preview:
        if find_formatting is not None:
            raise WordMcpError(
                "preview has no formatting criteria yet (the preview "
                "engine mirrors the plain replace path); run "
                "find_text(formatting=...) to scout instead"
            )
        return _pv.preview_replace(
            DocxPackage(file_path), replacements, scope=scope
        )
    if find_formatting is not None:
        if track:
            raise WordMcpError(
                "track is not supported with find_formatting "
                "(v1 replace_formatted parity)"
            )
        if any(
            isinstance(item, dict) and item.get("regex")
            for item in replacements
        ):
            raise WordMcpError(
                "regex is not supported with find_formatting "
                "(v1 replace_formatted parity)"
            )
        if live == "force":
            raise WordMcpError(
                "find_formatting mode is file-mode only; close the "
                "document in Word first"
            )

        def _do(pkg: DocxPackage) -> dict:
            replaced: dict[str, int] = {}
            total = 0
            for i, item in enumerate(replacements):
                if not isinstance(item, dict) or "replace" not in item:
                    raise WordMcpError(
                        "each replacements item needs 'replace' (and "
                        "'find' unless replacing whole formatted stretches)"
                    )
                r = _sf.replace_formatted(
                    pkg,
                    formatting=find_formatting,
                    replace=item["replace"],
                    find=item.get("find"),
                    scope=scope,
                    max_replacements=max_replacements,
                )
                item_total = r.get("total", 0)
                total += item_total
                key = item.get("find") or f"(formatted stretch {i})"
                replaced[key] = item_total
            return {
                "replaced": replaced,
                "total": total,
                "find_formatting": find_formatting,
            }

        return _edit(file_path, _do, backup=backup)
    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.search_and_replace(
                pkg,
                replacements,
                scope=scope,
                max_replacements=max_replacements,
                track=track,
                author=author,
            ),
            backup=backup,
        ),
        lambda: _lo.search_and_replace(
            file_path, replacements, scope, max_replacements, track, author
        ),
    )


@_tool("lite")
def get_document_view(
    file_path: str,
    scope: dict | None = None,
    detail: str = "text",
    include: dict | None = None,
    stamp_anchors: bool = False,
) -> dict:
    """Read the document as an anchored markdown projection, the low-token
    alternative to get_text for orientation and bulk editing. One block
    per paragraph, prefixed [hex] with a stable anchor id (from
    w14:paraId, which survives edits elsewhere in the document); headings
    carry # prefixes, tables render as pipe tables under [t:hex] with
    cells addressed t:hex:rNcN (1-based). Anchors work in every location
    object ({"anchor": "hex"}) and in apply_edits ops. scope:
    {"outline": "3.2"} for one heading's section, or {"paragraphs":
    {"start": N, "end": M}} (end exclusive); omit for the whole document.
    detail: "structure" (headings and counts only), "text" (default),
    "full" (adds {++ins++}/{--del--} revision markers and [cN] comment
    refs with an author legend). include: {"tables": false} to skip
    tables, {"notes": "inline"} to append footnote/endnote text.
    Documents without paraIds get VOLATILE anchors (flagged in the
    header) that change with any edit; stamp_anchors=true writes real
    paraIds so anchors become durable. Stamping is the ONE mutation this
    tool can make and runs only when explicitly requested, with the
    normal backup and validated save; plain reads never modify the file.
    A document open in Word is read from its last saved state.
    """
    from .core.errors import DocumentLocked

    if stamp_anchors:
        probe = None
        try:
            probe = DocxPackage(file_path)
        except DocumentLocked:
            raise WordMcpError(
                "stamp_anchors is a mutation and needs the file closed in "
                "Word; close it and retry, or view without stamping"
            ) from None
        needs = any(
            p.get(qn("w14:paraId")) is None
            for p in probe.root().iter(qn("w:p"))
        )
        del probe
        if needs:
            def _do(pkg: DocxPackage) -> dict:
                res = _view.stamp_anchors(pkg)
                out = _view.get_document_view(
                    pkg, scope=scope, detail=detail, include=include
                )
                out["stamped"] = res["stamped"]
                return out

            return _edit(file_path, _do)
        # everything already stamped: fall through to the pure read
    try:
        return _view.get_document_view(
            DocxPackage(file_path), scope=scope, detail=detail,
            include=include,
        )
    except DocumentLocked:
        out = _view.get_document_view(
            _disk_snapshot_pkg(file_path), scope=scope, detail=detail,
            include=include,
        )
        out["live"] = True
        out["note"] = (
            "document open in Word: this view reflects the last SAVED "
            "state; unsaved changes are not shown"
        )
        return out


@_tool("lite")
def apply_edits(
    file_path: str,
    edits: list[dict],
    atomic: bool = True,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Apply a batch of anchor-addressed edits in one call: one lock,
    backup, and validated save for the whole batch. Anchors come from
    get_document_view. Ops (each edit is a dict with "op"): replace
    {anchor, find, text, occurrence?} (occurrence omitted = every
    match); set_text {anchor, text} (whole paragraph);
    insert {location, markdown} (headings, plain paragraphs, lists, and
    pipe tables become real Word structures; location is the standard
    object, e.g. {"anchor": "hex", "position": "after"}); delete {anchor
    or anchors}; set_style {anchor, style}; format {anchor, formatting,
    find?, occurrence?}; set_paragraph_format {anchor, format}; set_cell
    {anchor: "t:hex:rNcN", text}. The whole batch validates BEFORE
    anything mutates, against BATCH-START text (never reference text an
    earlier op creates); one stale anchor refuses it all (STALE_ANCHOR
    lists failed ops: re-view, resend). The changed map carries per-op
    results, with fresh anchors for inserted paragraphs, so follow-up
    batches chain without re-viewing. Ops run in order; keep deletes
    last. Use this for 3+ edits in one section, else the fine-grained
    tools. Auto-backup in file mode (prev/anchor slots
    in .ks4w-backups; backup=False skips rotation); atomic validated
    save. A document open in Word is edited live as ONE undo step
    (markdown lists and tables are file-mode only there; unsaved-change
    stale targets refuse: save in Word, resend).
    """
    if atomic is not True:
        raise WordMcpError(
            "only atomic=true is supported: the whole batch applies in "
            "one save or nothing does. Split into separate apply_edits "
            "calls for independent failure domains."
        )

    def _file_call() -> dict:
        return _edit(
            file_path,
            lambda pkg: _batch.apply_edits(pkg, edits, atomic=atomic),
            backup=backup,
        )

    def _live_call() -> dict:
        from .com import live_batch as _lb

        snap = _disk_snapshot_pkg(file_path)
        plans = _batch.validate_edits(snap, edits)
        snap_texts = _snapshot_para_texts(snap)
        return _lb.apply_edits_live(
            file_path, edits, plans, len(snap_texts), snap_texts
        )

    return _route_live(live, _file_call, _live_call)


@_tool("lite")
def format_text(
    file_path: str,
    formatting: dict | None = None,
    case: str | None = None,
    range: dict | None = None,
    find: str | None = None,
    occurrence: int | None = None,
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Apply character formatting or change case on a text range. formatting:
    bold, italic, underline, strike, font, size_pt, color, highlight,
    small_caps, char_spacing_pt, language, east_asian_language, and more.
    case: upper | lower | title | sentence. Target: range={start,end},
    find, or both; one of formatting or case per call. Auto-backup in file
    mode: prev/anchor slots in .ks4w-backups (backup=False skips rotation
    only); atomic validated save. Formatting goes live on open documents;
    case is file-mode only.
    """
    from .com import live_ops as _lo

    if (formatting is None) == (case is None):
        raise WordMcpError(
            "give exactly one of formatting or case (make two calls for "
            "both)"
        )
    if range is None and find is None:
        raise WordMcpError("give a target: range={start, end}, find, or both")

    if case is not None:
        if occurrence is not None:
            raise WordMcpError(
                "occurrence applies to formatting mode only; case mode "
                "changes every occurrence of find"
            )
        if live == "force":
            raise WordMcpError(
                "case mode is file-mode only; close the document in Word "
                "first"
            )

        def _do(pkg: DocxPackage) -> dict:
            import builtins

            indices = None
            if range is not None:
                s, e = _expand_range(pkg, range)
                indices = list(builtins.range(s, e + 1))
            return _tx.change_case(pkg, case, indices=indices, find=find)

        return _edit(file_path, _do, backup=backup)

    occ = 1 if occurrence is None else occurrence

    def _single_index(resolver) -> int | None:
        if range is None:
            return None
        s, e = resolver()
        if s != e:
            raise WordMcpError(
                "formatting mode takes a SINGLE-paragraph range (start == "
                "end); make one call per paragraph"
            )
        return s

    def _file_call() -> dict:
        def _do(pkg: DocxPackage) -> dict:
            idx = _single_index(lambda: _expand_range(pkg, range))
            return _tx.format_text(
                pkg, paragraph_index=idx, find=find, occurrence=occ,
                formatting=formatting,
            )

        return _edit(file_path, _do, backup=backup)

    def _live_call() -> dict:
        idx = None
        verify = None
        if range is not None:
            rr, verify, _ve = _expand_range_verified(
                _disk_snapshot_pkg(file_path), range
            )
            if rr.start_index != rr.end_index:
                raise WordMcpError(
                    "formatting mode takes a SINGLE-paragraph range (start "
                    "== end); make one call per paragraph"
                )
            idx = rr.start_index
        return _lo.format_text(
            file_path, formatting, paragraph_index=idx, find=find,
            occurrence=occ, verify_text=verify,
        )

    return _route_live(live, _file_call, _live_call)


@_tool("lite")
def set_paragraph_format(
    file_path: str, indices: list[int], formatting: dict, backup: bool = True,
    live: str = "auto",
) -> dict:
    """Set paragraph formatting on a batch of paragraphs (0-based indices).
    Keys: alignment, space_before_pt, space_after_pt, line_spacing,
    indent_left_pt, indent_right_pt, first_line_indent_pt, keep_with_next,
    outline_level. outline_level (0-8; null removes) sets w:outlineLvl
    without changing the look. Auto-backup in file mode: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Documents open in Word are edited live (shading, borders,
    tab_stops refused live).
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tx.set_paragraph_format(pkg, indices, formatting),
            backup=backup,
        ),
        lambda: _lo.set_paragraph_format(file_path, indices, formatting),
    )


@_tool("lite")
def apply_style(
    file_path: str,
    style: str,
    range: dict | None = None,
    target: dict | None = None,
    backup: bool = True,
) -> dict:
    """Apply a named style. range={start,end} (locations or bare indices)
    applies a paragraph style to those paragraphs, changing their full look
    (Heading1-9 auto-created); target={search:{text, occurrence?}} applies
    a character style to the matched text instead. For an outline level without visual change use
    set_paragraph_format. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    Defining new named styles lives in the academic pack.
    """
    if (range is None) == (target is None):
        raise WordMcpError(
            "give exactly one of range (paragraph style) or target "
            "(character style on matched text)"
        )
    if range is not None:
        def _do(pkg: DocxPackage) -> dict:
            s, e = _expand_range(pkg, range)
            import builtins
            return _tx.apply_style(pkg, list(builtins.range(s, e + 1)), style)

        return _edit(file_path, _do, backup=backup)
    if not isinstance(target, dict) or set(target) != {"search"}:
        raise WordMcpError(
            'target takes {"search": {"text": ..., "occurrence"?: n}}'
        )
    search = target["search"]
    if not isinstance(search, dict) or "text" not in search:
        raise WordMcpError(
            'target.search requires text, like {"search": {"text": "bold '
            'me", "occurrence": 1}}'
        )
    unknown = sorted(set(search) - {"text", "occurrence"})
    if unknown:
        raise WordMcpError(
            f"target.search got unknown key(s) {unknown}; it takes text "
            "plus optional occurrence"
        )
    return _edit(
        file_path,
        lambda pkg: _sx.apply_character_style(
            pkg, find=search["text"], style=style,
            occurrence=search.get("occurrence", 1),
        ),
        backup=backup,
    )


@_tool("academic")
def define_style(
    file_path: str,
    style_id: str,
    name: str,
    style_type: str = "paragraph",
    based_on: str | None = "Normal",
    next_style: str | None = None,
    character_formatting: dict | None = None,
    paragraph_formatting: dict | None = None,
    backup: bool = True,
) -> dict:
    """Create or replace a custom style (paragraph or character) with full
    formatting control; character_formatting takes the format_text keys,
    paragraph_formatting the set_paragraph_format keys. get_styles returns
    definitions in this exact input shape, so cloning is one read plus one
    define. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.define_style(
            pkg, style_id=style_id, name=name, style_type=style_type,
            based_on=based_on, next_style=next_style,
            character_formatting=character_formatting,
            paragraph_formatting=paragraph_formatting,
        ),
        backup=backup,
    )


# ================================================================ 2.8 tables


@_tool("lite")
def create_table(
    file_path: str,
    data: list[list[str]],
    location: dict | None = None,
    header_row: bool = True,
    width_pt: float | None = None,
    backup: bool = True,
) -> dict:
    """Create a table from 2D string data with single-line borders and a bold
    repeating header row (header_row=False for none). location is the
    standard location object ({paragraph}, {after_heading}, {outline},
    {search}, {anchor}); omit it to append at the document end; position
    'after' only. To build a table from a CSV or JSON file use
    import_table (protection-io pack). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        if location is None:
            return _tb.create_table(
                pkg, data, at_end=True, header_row=header_row,
                width_pt=width_pt,
            )
        loc = resolve_location(pkg, location)
        if loc.position == "end":
            return _tb.create_table(
                pkg, data, at_end=True, header_row=header_row,
                width_pt=width_pt,
            )
        if loc.position != "after":
            raise WordMcpError(
                "create_table supports position 'after' (default) or "
                f"'end'; got {loc.position!r}. Address the paragraph the "
                "table should follow."
            )
        return _tb.create_table(
            pkg, data, after_index=loc.paragraph_index,
            header_row=header_row, width_pt=width_pt,
        )

    return _edit(file_path, _do, backup=backup)


@_tool("lite")
def delete_table(file_path: str, table_index: int, backup: bool = True) -> dict:
    """Delete a whole table and its contents. table_index is the 0-based
    position among body-level tables in document order, as reported by
    list_elements(type='tables') or get_document_view; nested tables are
    removed with their host table. To clear cell values while keeping the
    grid, use set_cells instead. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _tb.delete_table(pkg, table_index),
        backup=backup,
    )


@_tool("lite")
def get_table(
    file_path: str,
    table_index: int,
    nested: dict | None = None,
) -> dict:
    """Read one table in full: every cell's text, the merge map (gridSpan
    horizontal, vMerge vertical), and column widths. table_index is 0-based
    among body-level tables in document order. For a table nested inside a
    cell, pass nested={row, cell, index} addressing the host cell (index
    picks among several, default 0). Write values with set_cells; reshape
    with modify_table_structure. Read-only; reads the last-saved state of a
    document open in Word.
    Row/column surgery, styling, and sort: media-forms pack.
    """
    pkg = DocxPackage(file_path)
    if nested is None:
        return _rd.get_table(pkg, table_index)
    unknown = sorted(set(nested) - {"row", "cell", "index"})
    if unknown:
        raise WordMcpError(
            f"nested got unknown key(s) {unknown}; it takes row, cell, and "
            "optional index"
        )
    for key in ("row", "cell"):
        if key not in nested:
            raise WordMcpError(f"nested requires {key!r} (the host cell)")
    return _tb.get_nested_table(
        pkg, table_index,
        row=nested["row"], cell=nested["cell"],
        nested_index=nested.get("index", 0),
    )


_MTS_ACTIONS = ("insert", "delete", "merge", "unmerge", "split")


@_tool("media-forms")
def modify_table_structure(
    file_path: str,
    table_index: int,
    action: str,
    target: str | None = None,
    at: int | None = None,
    count: int = 1,
    copy_format_from: int | None = None,
    width_pt: float | None = None,
    start: int | None = None,
    end: int | None = None,
    columns: list[int] | None = None,
    range: dict | None = None,
    row: int | None = None,
    cell: int | None = None,
    at_row: int | None = None,
    backup: bool = True,
) -> dict:
    """All grid-shape changes to one table, discriminated by action.
    action='insert' with target='rows'|'columns': insert count rows or grid
    columns before position at (at = current count appends); rows can copy
    structure and formatting from an existing row (copy_format_from),
    columns take width_pt. action='delete' with target='rows' (start..end
    inclusive; vertical merges are re-rooted, not broken) or
    target='columns' (columns=[0-based grid indices]; merged cells shrink
    and the grid stays consistent). action='merge': merge the rectangle
    range={start_row, end_row, start_col, end_col} (grid coordinates,
    inclusive). action='unmerge': split the merged cell at row/cell back
    into single cells (horizontal and vertical). action='split': split the
    table into two at at_row (that row starts the new table). Each action
    reads only its own parameters; extras are refused. table_index is
    0-based among body-level tables. Cell values are written with
    set_cells, persistent attributes with set_table_properties.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    if action not in _MTS_ACTIONS:
        raise WordMcpError(
            f"action must be one of {list(_MTS_ACTIONS)}, got {action!r}"
        )
    given = {
        "target": target, "at": at,
        "copy_format_from": copy_format_from, "width_pt": width_pt,
        "start": start, "end": end, "columns": columns, "range": range,
        "row": row, "cell": cell, "at_row": at_row,
    }
    if count != 1:
        given["count"] = count
    branch_params = {
        ("insert", "rows"): {"target", "at", "count", "copy_format_from"},
        ("insert", "columns"): {"target", "at", "count", "width_pt"},
        ("delete", "rows"): {"target", "start", "end"},
        ("delete", "columns"): {"target", "columns"},
        ("merge", None): {"range"},
        ("unmerge", None): {"row", "cell"},
        ("split", None): {"at_row"},
    }
    if action in ("insert", "delete"):
        if target not in ("rows", "columns"):
            raise WordMcpError(
                f"action={action!r} requires target='rows' or "
                "target='columns'"
            )
        key = (action, target)
    else:
        if target is not None:
            raise WordMcpError(
                f"target does not apply to action={action!r}"
            )
        key = (action, None)
    allowed = branch_params[key]
    foreign = sorted(
        name for name, value in given.items()
        if value is not None and name not in allowed
    )
    if foreign:
        raise WordMcpError(
            f"parameter(s) {foreign} do not belong to action={action!r}"
            + (f" target={target!r}" if target else "")
            + "; each action reads only its own parameters"
        )
    if action == "insert" and (
        isinstance(count, bool) or not isinstance(count, int) or count < 1
    ):
        # Adversarial 6a: count=0 / count=-3 previously rode through to a
        # silent zero-row no-op reported as success.
        raise WordMcpError(f"count must be an integer >= 1, got {count!r}")
    if key == ("insert", "rows"):
        if at is None:
            raise WordMcpError("insert rows requires at (position)")
        return _edit(
            file_path,
            lambda pkg: _tb.insert_rows(
                pkg, table_index, at=at, count=count,
                copy_format_from=copy_format_from,
            ),
            backup=backup,
        )
    if key == ("insert", "columns"):
        if at is None:
            raise WordMcpError("insert columns requires at (position)")
        return _edit(
            file_path,
            lambda pkg: _tb.insert_columns(
                pkg, table_index, at=at, count=count, width_pt=width_pt
            ),
            backup=backup,
        )
    if key == ("delete", "rows"):
        if start is None:
            raise WordMcpError("delete rows requires start (first row)")
        return _edit(
            file_path,
            lambda pkg: _tb.delete_rows(pkg, table_index, start, end),
            backup=backup,
        )
    if key == ("delete", "columns"):
        if columns is None:
            raise WordMcpError(
                "delete columns requires columns=[0-based grid indices]"
            )
        return _edit(
            file_path,
            lambda pkg: _tb.delete_columns(pkg, table_index, columns),
            backup=backup,
        )
    if action == "merge":
        if range is None:
            raise WordMcpError(
                "merge requires range={start_row, end_row, start_col, "
                "end_col} (grid coordinates, inclusive)"
            )
        unknown = sorted(
            set(range) - {"start_row", "end_row", "start_col", "end_col"}
        )
        if unknown:
            raise WordMcpError(
                f"merge range got unknown key(s) {unknown}; it takes "
                "start_row, end_row, start_col, end_col"
            )
        for k in ("start_row", "end_row", "start_col", "end_col"):
            if k not in range:
                raise WordMcpError(f"merge range requires {k!r}")
        rng = range
        return _edit(
            file_path,
            lambda pkg: _tb.merge_cells(
                pkg, table_index,
                start_row=rng["start_row"], end_row=rng["end_row"],
                start_col=rng["start_col"], end_col=rng["end_col"],
            ),
            backup=backup,
        )
    if action == "unmerge":
        if row is None or cell is None:
            raise WordMcpError(
                "unmerge requires row and cell (the merged cell's host "
                "coordinates)"
            )
        return _edit(
            file_path,
            lambda pkg: _tb.unmerge_cells(pkg, table_index, row=row,
                                          cell=cell),
            backup=backup,
        )
    # split
    if at_row is None:
        raise WordMcpError(
            "split requires at_row (the first row of the new table)"
        )
    return _edit(
        file_path,
        lambda pkg: _tb.split_table(pkg, table_index, at_row=at_row),
        backup=backup,
    )


@_tool("media-forms")
def set_table_properties(
    file_path: str,
    table_index: int,
    style: str | None = None,
    banded_rows: bool = True,
    first_row_header: bool = True,
    column_widths: list[float] | None = None,
    header_row_repeat: bool | int | None = None,
    backup: bool = True,
) -> dict:
    """Set table-wide attributes; pass any combination. style names a table
    style (TableGrid bordered, PlainTable, BandedTable banded shading;
    banded_rows and first_row_header tune it; missing definitions are
    injected). column_widths sets every grid column width in points, one
    value per column. header_row_repeat repeats the first N rows as the
    header on every page (true=1, false=off). Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    if style is None and column_widths is None and header_row_repeat is None:
        raise WordMcpError(
            "nothing to set; pass style, column_widths, and/or "
            "header_row_repeat"
        )

    def _do(pkg: DocxPackage) -> dict:
        out: dict = {}
        if style is not None:
            out.update(_tb.apply_table_style(
                pkg, table_index, style,
                banded_rows=banded_rows, first_row_header=first_row_header,
            ))
        if column_widths is not None:
            out.update(_tb.set_column_widths(pkg, table_index, column_widths))
        if header_row_repeat is not None:
            # bool checks BEFORE int: bool subclasses int in Python.
            if header_row_repeat is False:
                out.update(_tb.set_header_row_repeat(pkg, table_index,
                                                     on=False))
            else:
                n = 1 if header_row_repeat is True else int(header_row_repeat)
                out.update(_tb.set_header_row_repeat(pkg, table_index,
                                                     rows=n, on=True))
        return out

    return _edit(file_path, _do, backup=backup)


@_tool("lite")
def set_cells(
    file_path: str,
    table_index: int,
    edits: list[dict] | None = None,
    block: dict | None = None,
    nested: dict | None = None,
    track: bool = False,
    author: str = "Claude",
    backup: bool = True,
    live: str = "auto",
) -> dict:
    """Write many table cells in one call. Modes: edits=[{row, cell, text}]
    for scattered cells, or block={origin:{row, cell}, values:[[...]]} for
    a 2D block. nested={row, cell, index} targets a table nested in that
    host cell (edits mode only). track records tracked changes by author.
    Live mode (plain edits only) refuses vertical merges; file mode is
    merge-aware. Auto-backup in file mode: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Documents open in Word are edited live.
    """
    if (edits is None) == (block is None):
        raise WordMcpError(
            "give exactly one input mode: edits=[{row, cell, text}] or "
            "block={origin, values}"
        )
    if nested is not None and block is not None:
        raise WordMcpError("nested tables take scatter edits, not block")
    if nested is not None and track:
        raise WordMcpError(
            "tracked changes are not recorded inside nested tables"
        )
    if block is not None and track:
        raise WordMcpError("track applies to scattered edits only")
    if live == "force" and (block is not None or nested is not None):
        raise WordMcpError(
            "only scattered edits on a body-level table route live; block "
            "and nested modes are file-only"
        )
    if block is not None:
        unknown = sorted(set(block) - {"origin", "values"})
        if unknown:
            raise WordMcpError(
                f"block got unknown key(s) {unknown}; it takes origin and "
                "values"
            )
        origin = block.get("origin")
        if not isinstance(origin, dict) or not {"row", "cell"} <= set(origin):
            raise WordMcpError(
                'block.origin takes {"row": r, "cell": c}'
            )
        if "values" not in block:
            raise WordMcpError("block requires values=[[...]]")
        return _edit(
            file_path,
            lambda pkg: _tb.set_cells_grid(
                pkg, table_index,
                origin_row=origin["row"], origin_cell=origin["cell"],
                data=block["values"],
            ),
            backup=backup,
        )
    if nested is not None:
        unknown = sorted(set(nested) - {"row", "cell", "index"})
        if unknown:
            raise WordMcpError(
                f"nested got unknown key(s) {unknown}; it takes row, cell, "
                "and optional index"
            )
        for key in ("row", "cell"):
            if key not in nested:
                raise WordMcpError(f"nested requires {key!r} (the host cell)")
        return _edit(
            file_path,
            lambda pkg: _tb.set_nested_cells(
                pkg, table_index, row=nested["row"], cell=nested["cell"],
                edits=edits, nested_index=nested.get("index", 0),
            ),
            backup=backup,
        )
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _edit(
            file_path,
            lambda pkg: _tb.set_cells(
                pkg, table_index, edits, track=track, author=author
            ),
            backup=backup,
        ),
        lambda: _lo.set_cells(
            file_path, table_index, edits, track=track, author=author
        ),
    )


@_tool("media-forms")
def format_cells(
    file_path: str,
    table_index: int,
    targets: list[dict],
    formatting: dict,
    backup: bool = True,
) -> dict:
    """Format table cells in bulk. targets=[{row, cell?}], where {row} alone
    selects the whole row; formatting applies shading, bold, italic,
    alignment, valign, and padding_pt to every targeted cell. table_index
    is 0-based among body-level tables. For cell values use set_cells; for
    the whole-table style use set_table_properties. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.format_cells(pkg, table_index, targets, formatting),
        backup=backup,
    )


@_tool("media-forms")
def sort_table(
    file_path: str,
    table_index: int,
    column: int,
    numeric: bool = False,
    descending: bool = False,
    has_header: bool = True,
    backup: bool = True,
) -> dict:
    """Sort a table's data rows by one column (CELL index, 0-based). numeric
    compares values as numbers instead of text, descending reverses the
    order, has_header keeps the first row in place. Refused when vertical
    merges make rows interdependent; unmerge first with
    modify_table_structure. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tb.sort_table(
            pkg, table_index, column=column, numeric=numeric,
            descending=descending, has_header=has_header,
        ),
        backup=backup,
    )


@_tool("protection-io")
def export_table(
    file_path: str,
    table_index: int,
    format: str = "csv",
    output_path: str | None = None,
    include_merges: bool = True,
) -> dict:
    """Export one body-level table to CSV or JSON as a rows x grid-columns
    matrix. Merged cells keep their value in the anchor (top-left) position
    with empty strings under the covered span; include_merges reports the
    merge topology as {row, col, rowspan, colspan}. Nested tables are
    flattened into the host cell's text and flagged. output_path=None
    returns the data inline; an existing output file is refused. The
    document itself is never modified.
    """
    return _dio.export_table(
        DocxPackage(file_path),
        table_index,
        format=format,
        output_path=output_path,
        include_merges=include_merges,
    )


@_tool("protection-io")
def import_table(
    file_path: str,
    data: list | str,
    table_index: int | None = None,
    at_end: bool = False,
    after_anchor: str | None = None,
    has_header: bool = True,
    backup: bool = True,
) -> dict:
    """Fill a table from data: a .csv path, a .json path (export_table's form
    or a bare row list), or an inline list of lists. Without table_index a
    NEW table is created (at_end or after_anchor, default end; has_header
    bolds and repeats the first row). With table_index the existing table's
    cells are OVERWRITTEN in place: data must exactly match rows x
    grid-columns, refused otherwise. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _dio.import_table(
            pkg,
            data,
            table_index=table_index,
            at_end=at_end,
            after_anchor=after_anchor,
            has_header=has_header,
        ),
        backup=backup,
    )


@_tool("assembly")
def copy_table(
    target_path: str,
    source_path: str,
    table_index: int,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    backup: bool = True,
) -> dict:
    """Copy ONE table from source_path into target_path, with
    insert_document's resource reconciliation (styles matched by name;
    numbering, images, notes, bookmarks reconciled). table_index counts the
    SOURCE's body tables, 0-based. Position, one of: after_index (a body
    ITEM index counting paragraphs AND tables), after_anchor (full
    paragraph text), or at_end. The source is not modified. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        target_path,
        lambda pkg: _asm.copy_table(
            pkg,
            source_path,
            table_index,
            after_index=after_index,
            after_anchor=after_anchor,
            at_end=at_end,
        ),
        backup=backup,
    )


# ================================================= 2.9 footnotes and endnotes


@_tool("academic")
def manage_note(
    file_path: str,
    action: str,
    note_type: str = "footnote",
    text: str | None = None,
    location: dict | None = None,
    note_id: str | None = None,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Footnote and endnote lifecycle in one tool. action='insert' creates a
    note: note_type ('footnote' or 'endnote'), text, and a location object
    placing the reference mark (a search selector anchors after the matched
    text; other selectors anchor on the resolved paragraph); all note
    infrastructure (parts, styles, superscript) is created if the document
    has none. action='edit' rewrites a note's text: note_type, text, and
    exactly one of note_id or position (1-based display order).
    action='delete' removes the definition AND the body reference mark,
    always both: note_type plus note_id or position.
    action='cleanup_orphans' purges note definitions no body reference
    points to, both types at once (for documents that arrived with orphans;
    this server's own deletes clean up automatically). Use
    list_elements(type='footnotes' or 'endnotes') for ids and text,
    validate(checks=['notes']) for integrity, convert_notes to switch
    kinds. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    actions = ("insert", "edit", "delete", "cleanup_orphans")
    if action not in actions:
        raise WordMcpError(f"action must be one of {list(actions)}")
    if action != "cleanup_orphans" and note_type not in ("footnote",
                                                         "endnote"):
        raise WordMcpError("note_type must be 'footnote' or 'endnote'")
    if action == "insert":
        if text is None:
            raise WordMcpError("action='insert' requires text (the note)")
        if location is None:
            raise WordMcpError(
                "action='insert' requires a location object placing the "
                "reference mark"
            )
        if note_id is not None or position is not None:
            raise WordMcpError(
                "note_id/position address existing notes; they do not "
                "apply to insert"
            )
        return _edit(
            file_path,
            lambda pkg: _nt.add_note(
                pkg, note_type, note_text=text,
                **_anchor_kwargs(pkg, location),
            ),
            backup=backup,
        )
    if location is not None:
        raise WordMcpError(
            f"location applies to action='insert' only, not {action!r}"
        )
    if action in ("edit", "delete") and (
        note_id is not None and position is not None
    ):
        # Adversarial 6a F7: both given rode through and note_id silently
        # won, mis-targeting when they disagree. Exactly one, as promised.
        raise WordMcpError(
            f"action={action!r} takes exactly ONE of note_id or position, "
            "got both"
        )
    if action == "edit":
        if text is None:
            raise WordMcpError(
                "action='edit' requires text (the replacement)"
            )
        return _edit(
            file_path,
            lambda pkg: _nt.edit_note(
                pkg, note_type, note_id=note_id, position=position,
                new_text=text,
            ),
            backup=backup,
        )
    if action == "delete":
        if text is not None:
            raise WordMcpError("text does not apply to action='delete'")
        return _edit(
            file_path,
            lambda pkg: _nt.delete_note(
                pkg, note_type, note_id=note_id, position=position
            ),
            backup=backup,
        )
    return _edit(file_path, lambda pkg: _nt.purge_orphans(pkg), backup=backup)


@_tool("academic")
def convert_notes(
    file_path: str,
    direction: str,
    note_id: str | None = None,
    position: int | None = None,
    backup: bool = True,
) -> dict:
    """Convert footnotes to endnotes or back. direction: footnotes_to_endnotes
    | endnotes_to_footnotes; give note_id or position for ONE note, neither
    for ALL. Word renumbers automatically. Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _nt.convert_notes(
            pkg, direction, note_id=note_id, position=position
        ),
        backup=backup,
    )


# ====================================== 2.10 references and generated lists


@_tool("references")
def manage_source(
    file_path: str,
    action: str,
    tag: str,
    source_type: str | None = None,
    title: str | None = None,
    year: str | None = None,
    authors: list | None = None,
    editors: list | None = None,
    journal_name: str | None = None,
    book_title: str | None = None,
    publisher: str | None = None,
    city: str | None = None,
    pages: str | None = None,
    volume: str | None = None,
    issue: str | None = None,
    edition: str | None = None,
    institution: str | None = None,
    url: str | None = None,
    internet_site_title: str | None = None,
    style: str = "APA",
    extra_fields: dict | None = None,
    force: bool = False,
    backup: bool = True,
) -> dict:
    """Word-native bibliography source lifecycle. action='add' stores a
    source: tag (unique citation key), source_type ('JournalArticle',
    'Book', 'BookSection', 'Report', 'InternetSite', ...), title, and
    optional year, authors/editors ([{last, first, middle?}] or
    [{corporate}]), journal_name, book_title, publisher, city, pages,
    volume, issue, edition, institution, url, internet_site_title, style,
    extra_fields. action='delete' removes a source by tag; refused while
    cited unless force=True. Cite stored sources with insert_citation and
    render the list with insert_reference_list(type='bibliography');
    enumerate them with list_elements(type='sources'). Run
    detect_citation_system first on unfamiliar documents: mixing
    Word-native and Zotero/Mendeley/EndNote citations creates a
    bibliography no single manager maintains. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    if action == "add":
        if source_type is None or title is None:
            raise WordMcpError(
                "action='add' requires source_type and title"
            )
        return _edit(
            file_path,
            lambda pkg: _bib.add_source(
                pkg, tag=tag, source_type=source_type, title=title,
                year=year, authors=authors, editors=editors,
                journal_name=journal_name, book_title=book_title,
                publisher=publisher, city=city, pages=pages, volume=volume,
                issue=issue, edition=edition, institution=institution,
                url=url, internet_site_title=internet_site_title,
                style=style, extra_fields=extra_fields,
            ),
            backup=backup,
        )
    if action == "delete":
        return _edit(
            file_path,
            lambda pkg: _bib.delete_source(pkg, tag, force=force),
            backup=backup,
        )
    raise WordMcpError("action must be 'add' or 'delete'")


@_tool("references")
def insert_citation(
    file_path: str,
    tag: str,
    location: dict,
    pages: str | None = None,
    suppress_author: bool = False,
    suppress_year: bool = False,
    prefix: str | None = None,
    suffix: str | None = None,
    backup: bool = True,
) -> dict:
    """Insert a CITATION field for a stored source (tag from manage_source or
    list_elements(type='sources')) at a location object position; it
    renders in the document citation style on field update, with a
    placeholder until then. pages, prefix/suffix, and the suppress flags
    shape the rendered cite. Run detect_citation_system before citing in an
    unfamiliar document. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _bib.insert_citation(
            pkg, tag=tag, pages=pages, suppress_author=suppress_author,
            suppress_year=suppress_year, prefix=prefix, suffix=suffix,
            **_anchor_kwargs(pkg, location),
        ),
        backup=backup,
    )


@_tool("references")
def set_bibliography_style(file_path: str, style: str, backup: bool = True) -> dict:
    """Set the style Word-native CITATION and BIBLIOGRAPHY fields render in:
    APA | Chicago | MLA | IEEE | Turabian | Harvard - Anglia | GB7714 |
    GOST - Name Sort | GOST - Title Sort | ISO 690 - First Element and Date
    | ISO 690 - Numerical Reference | SIST02. Affects Word-native fields
    only: manager citations restyle in their manager; plain-text citations
    take convert_citation_style. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path, lambda pkg: _bib.set_bibliography_style(pkg, style),
        backup=backup,
    )


_REFLIST_TYPES = (
    "bibliography", "toc", "figure_list", "table_list", "equation_list",
    "index", "glossary",
)
_REFLIST_DEFAULT_TITLES = {
    "bibliography": "Bibliography",
    "toc": "Table of Contents",
    "figure_list": None,
    "table_list": None,
    "equation_list": None,
    "index": "Index",
    "glossary": "Glossary",
}
_REFLIST_OPTION_KEYS = {
    "bibliography": set(),
    "toc": {"levels"},
    "figure_list": set(),
    "table_list": set(),
    "equation_list": set(),
    "index": {"columns", "letter_headings"},
    "glossary": {"heading_level", "definition_patterns"},
}
_CAPTION_LABELS = {
    "figure_list": "Figure", "table_list": "Table",
    "equation_list": "Equation",
}


@_tool("academic")
def insert_reference_list(
    file_path: str,
    type: str,
    title: str | None = "__default__",
    location: dict | None = None,
    options: dict | None = None,
    update_on_open: bool = True,
    backup: bool = True,
) -> dict:
    """Insert any generated reference list in one tool. type: 'bibliography'
    (Word renders the styled reference list from the manage_source store),
    'toc' (options {levels: '1-3'}), 'figure_list' | 'table_list' |
    'equation_list' (built from insert_caption SEQ entries), 'index'
    (compiles mark_index_entry XE entries; options {columns,
    letter_headings}), or 'glossary' (harvested from defined terms; options
    {heading_level, definition_patterns}). title overrides each type's
    default heading. The location object places the list; omit it for the
    type's default spot (bibliography, index, and glossary at the body end;
    toc and caption lists at the document start). Field-based lists fill in
    their entries and page numbers when Word updates fields: automatically
    on next open (update_on_open) or immediately via com_refresh_fields.
    Delete any generated list with delete_element(type='reference_list').
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    if type not in _REFLIST_TYPES:
        raise WordMcpError(
            f"unknown type {type!r}; one of: " + " | ".join(_REFLIST_TYPES)
        )
    opts = options or {}
    allowed = _REFLIST_OPTION_KEYS[type]
    unknown = sorted(set(opts) - allowed)
    if unknown:
        raise WordMcpError(
            f"unknown option key(s) {unknown} for type {type!r}; it takes "
            f"{sorted(allowed) if allowed else 'no options'}"
        )
    heading = _REFLIST_DEFAULT_TITLES[type] if title == "__default__" \
        else title

    def _do(pkg: DocxPackage) -> dict:
        after_index: int | None = None
        at_end = False
        at_start = False
        if location is not None:
            r = resolve_location(pkg, location)
            if r.position == "end":
                at_end_types = ("bibliography", "index", "glossary")
                if type in at_end_types:
                    at_end = True
                # toc/caption lists: fall through with after_index at the
                # last paragraph.
                else:
                    after_index = _last_paragraph_index(pkg)
            elif r.position == "after":
                after_index = r.paragraph_index
            elif r.position in ("before", "start"):
                if r.paragraph_index == 0:
                    if type in ("toc", *_CAPTION_LABELS):
                        at_start = True
                    else:
                        raise WordMcpError(
                            f"type {type!r} has no document-start form; "
                            "use position 'after' or omit location for "
                            "the default placement"
                        )
                else:
                    after_index = r.paragraph_index - 1
            else:
                raise WordMcpError(
                    "position 'replace' is not meaningful for a generated "
                    "list"
                )
        else:
            if type in ("bibliography", "index", "glossary"):
                at_end = True
        if type == "bibliography":
            return _bib.insert_bibliography(
                pkg, title=heading, after_index=after_index, at_end=at_end,
                update_on_open=update_on_open,
            )
        if type == "toc":
            return _tc.insert_toc(
                pkg, levels=opts.get("levels", "1-3"), title=heading,
                after_index=after_index, at_start=at_start,
                update_on_open=update_on_open,
            )
        if type in _CAPTION_LABELS:
            return _tc.insert_caption_list(
                pkg, label=_CAPTION_LABELS[type], title=heading,
                after_index=after_index, at_start=at_start,
                update_on_open=update_on_open,
            )
        if type == "index":
            return _fl.insert_index(
                pkg, title=heading, columns=opts.get("columns", 2),
                letter_headings=opts.get("letter_headings", True),
                after_index=after_index, at_end=at_end,
                update_on_open=update_on_open,
            )
        # glossary (no update_on_open: literal paragraphs, not a field)
        return _dt.insert_glossary(
            pkg, heading=heading if heading is not None else "Glossary",
            heading_level=opts.get("heading_level", 1),
            after_index=after_index, at_end=at_end,
            definition_patterns=opts.get("definition_patterns"),
        )

    return _edit(file_path, _do, backup=backup)


@_tool("academic")
def insert_cross_reference(
    file_path: str,
    to_bookmark: str,
    location: dict,
    kind: str = "page",
    backup: bool = True,
) -> dict:
    """Insert a cross-reference field pointing at a bookmark (create targets
    with insert_bookmark). kind: 'page' inserts the target's page number,
    'text' the bookmarked text; Word computes the value on field update,
    and validate(checks=['cross_references']) audits for broken targets.
    The location object places the field (a search selector puts it right
    after the matched text). Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        kw = _anchor_kwargs(pkg, location)
        return _fl.add_cross_reference(
            pkg, after_anchor=kw["anchor_text"],
            occurrence=kw["occurrence"], to_bookmark=to_bookmark, kind=kind,
        )

    return _edit(file_path, _do, backup=backup)


@_tool("media-forms")
def insert_field(
    file_path: str,
    field_code: str,
    location: dict,
    placeholder: str = "",
    backup: bool = True,
) -> dict:
    """Insert a Word field at a location object position. field_code examples:
    'DATE', 'FILENAME', 'NUMPAGES', 'PAGE', 'SEQ Exhibit \\* Arabic'. Codes
    are validated against an allowlist of known-safe fields; anything that
    links out or executes is refused, naming the allowlist. The field is
    written dirty so Word computes the result on the next refresh;
    placeholder shows until then. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        kw = _anchor_kwargs(pkg, location)
        return _fl.insert_field(
            pkg, field_code=field_code, after_anchor=kw["anchor_text"],
            occurrence=kw["occurrence"], placeholder=placeholder,
        )

    return _edit(file_path, _do, backup=backup)


@_tool("academic")
def mark_index_entry(
    file_path: str,
    entry: str,
    location: dict,
    subentry: str | None = None,
    bold_page: bool = False,
    italic_page: bool = False,
    see: str | None = None,
    backup: bool = True,
) -> dict:
    """Mark a spot for the index: an invisible XE field written at the
    location object position (a search selector marks where the matched
    text sits). entry and optional subentry name the index line;
    bold_page/italic_page style the page number; see='other entry' writes a
    cross-reference instead of a page number. Compile the finished index
    with insert_reference_list(type='index'). Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        kw = _anchor_kwargs(pkg, location)
        return _fl.mark_index_entry(
            pkg, anchor_text=kw["anchor_text"], entry=entry,
            subentry=subentry, occurrence=kw["occurrence"],
            bold_page=bold_page, italic_page=italic_page, see=see,
        )

    return _edit(file_path, _do, backup=backup)


@_tool("references")
def parse_references(file_path: str, style_hint: str | None = None) -> dict:
    """Parse the manuscript's reference list (References / Bibliography /
    Works Cited heading) and in-text citations into a structured model:
    read-only stage 1 of style conversion. Heuristic parsing: every entry
    gets parse_confidence (full/partial/failed); failed entries return
    verbatim and are never converted. Also reports in-text citations with
    positions, whether the document uses citation FIELDS (text conversion
    refuses to touch those), and the detected system. style_hint tightens
    parsing. Read-only.
    """
    return _sc2.parse_references(DocxPackage(file_path), style_hint=style_hint)


@_tool("references")
def convert_citation_style(
    file_path: str,
    target_style: str,
    source_style: str = "auto",
    dry_run: bool = False,
    backup: bool = True,
) -> dict:
    """BETA (heuristic): review the flagged-items list in the result rather
    than trusting silently. Convert a manuscript's PLAIN-TEXT citations and
    reference list to a target publication style: apa7,
    chicago17-author-date, chicago17-notes, mla9, harvard, ieee, vancouver,
    asa. Heuristic text conversion, not a citation processor: only
    fully-parsed reference entries and unambiguously-resolved citations are
    converted; everything else is left verbatim and flagged for human
    review. Supports author-date to numbered ([n] or Vancouver
    superscripts, numbered by first appearance) and back, author-date to
    Chicago notes (each citation becomes a REAL footnote), and notes to
    author-date (only recognizably pure citation footnotes are harvested;
    mixed ones are flagged and left alone). Narrative citations keep the
    author's name; reference-list italics, ordering, heading, and hanging
    indents are handled. Documents using Word-native or
    Zotero/Mendeley/EndNote citation FIELDS are routed away
    (set_bibliography_style, or restyle in the manager); fields are never
    rewritten as text. dry_run=True returns the complete change plan and
    leaves the file byte-identical; parse_references is the read-only stage
    1. Any error during a real run leaves the original untouched.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    if dry_run:
        return _sc2.convert_citation_style(
            DocxPackage(file_path), target_style,
            source_style=source_style, dry_run=True,
        )
    return _edit(
        file_path,
        lambda pkg: _sc2.convert_citation_style(
            pkg, target_style, source_style=source_style, dry_run=False
        ),
        backup=backup,
    )


@_tool("references")
def detect_citation_system(file_path: str) -> dict:
    """Which citation system(s) the document uses: Word native (CITATION
    fields + sources store), Zotero (ADDIN ZOTERO_ITEM), Mendeley (ADDIN
    CSL_CITATION), EndNote (ADDIN EN.CITE), or plain typed text only.
    Counts per system across body, footnotes, and endnotes, plus a
    split_brain flag when more than one managed system is present (a
    split-brain bibliography: each manager only maintains its own fields).
    Run this BEFORE any citation work on an unfamiliar document.
    """
    return _cs.detect_citation_system(DocxPackage(file_path))


@_tool("references")
def search_zotero_library(
    query: str, db_path: str | None = None, limit: int = 20
) -> dict:
    """Search the user's LOCAL Zotero library (read-only; the database is
    never modified). Every query word must match the title, a creator's
    last name, the year, or the publication name. Returns each match's
    Zotero item key (the handle insert_zotero_citation takes) plus type,
    title, creators, year, and publication. The database defaults to
    <home>/Zotero/zotero.sqlite; pass db_path for a nonstandard data
    directory. Reads a point-in-time snapshot, so last-moment edits in a
    running Zotero may not appear yet.
    """
    return _zl.search_zotero_library(query, db_path=db_path, limit=limit)


@_tool("references")
def insert_zotero_citation(
    file_path: str,
    item_keys: list[str],
    location: dict,
    page: str | None = None,
    prefix: str | None = None,
    suffix: str | None = None,
    db_path: str | None = None,
    backup: bool = True,
) -> dict:
    """Insert a REAL Zotero citation field (ADDIN ZOTERO_ITEM CSL_CITATION)
    at a location object position; Zotero recognizes and refreshes it.
    item_keys come from search_zotero_library; page is the locator;
    prefix/suffix attach to single-item citations only. The visible text is
    a plain (Author, Year) placeholder until Refresh runs in Zotero's Word
    plugin; the database is only read. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _zl.insert_zotero_citation(
            pkg, item_keys, page=page, prefix=prefix, suffix=suffix,
            db_path=db_path, **_anchor_kwargs(pkg, location),
        ),
        backup=backup,
    )


# ============================== 2.11 comments, tracked changes, and review


@_tool("review")
def get_comments(
    file_path: str, author: str | None = None, live: str = "auto"
) -> list | dict:
    """Comments with authors, anchored text, threading, and resolved state,
    optionally filtered by author. Documents open in Word are read live
    (same entry shape; live ids are the comment's position, not the XML
    id). Manage threads with manage_comment; comment_report builds the
    whole-document reviewer matrix. Read-only.
    """
    from .com import live_ops as _lo

    return _route_live(
        live,
        lambda: _rd.get_comments(DocxPackage(file_path), author=author),
        lambda: _lo.get_comments(file_path, author=author),
    )


@_tool("review")
def manage_comment(
    file_path: str,
    action: str,
    text: str | None = None,
    author: str = "Claude",
    comment_id: str | None = None,
    location: dict | None = None,
    done: bool = True,
    backup: bool = True,
) -> dict:
    """Comment lifecycle in one tool. action='add' places a new comment:
    text, author, and a location object choosing the anchored range (a
    search selector comments the matched text; other selectors anchor on
    the resolved paragraph); threaded-comment infrastructure is created as
    needed. action='reply' threads a reply under an existing comment:
    comment_id (from get_comments), text, author. action='resolve' marks a
    thread resolved, or reopens it with done=False. action='delete' removes
    a comment, its replies, and all body markers. Every action except add
    addresses by comment_id. Read threads with get_comments; comment_report
    builds the full reviewer matrix. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    actions = ("add", "reply", "resolve", "delete")
    if action not in actions:
        raise WordMcpError(f"action must be one of {list(actions)}")
    if action == "add":
        if text is None:
            raise WordMcpError("action='add' requires text")
        if location is None:
            raise WordMcpError(
                "action='add' requires a location object choosing the "
                "anchored range"
            )
        return _edit(
            file_path,
            lambda pkg: _cm.add_comment(
                pkg, text=text, author=author,
                **_anchor_kwargs(pkg, location),
            ),
            backup=backup,
        )
    if location is not None:
        raise WordMcpError(
            f"location applies to action='add' only, not {action!r}"
        )
    if comment_id is None:
        raise WordMcpError(
            f"action={action!r} requires comment_id (from get_comments)"
        )
    if action == "reply":
        if text is None:
            raise WordMcpError("action='reply' requires text")
        return _edit(
            file_path,
            lambda pkg: _cm.reply_to_comment(
                pkg, comment_id=comment_id, text=text, author=author
            ),
            backup=backup,
        )
    if action == "resolve":
        return _edit(
            file_path,
            lambda pkg: _cm.resolve_comment(pkg, comment_id=comment_id,
                                            done=done),
            backup=backup,
        )
    return _edit(
        file_path,
        lambda pkg: _cm.delete_comment(pkg, comment_id=comment_id),
        backup=backup,
    )


@_tool("review")
def comment_report(files: list[str], include_resolved: bool = True) -> dict:
    """The full reviewer matrix: every comment thread with author, date,
    anchored text, replies nested under parents, resolved flag, and a
    locator (paragraph index plus heading path); summary gives per-author,
    open/resolved, and per-section counts. files is an array: one file
    reports that document (include_resolved=False drops resolved threads);
    several files merge the matrix across copies of one draft, keyed by
    (author, anchored text), with per-file provenance and collision
    detection. Read-only on every file.
    """
    if not isinstance(files, list) or not files:
        raise WordMcpError("files takes a non-empty list of .docx paths")
    if len(files) == 1:
        return _rc.comment_report(
            DocxPackage(files[0]), include_resolved=include_resolved
        )
    if include_resolved is False:
        raise WordMcpError(
            "include_resolved applies to single-file reports; the "
            "multi-file merge always includes resolved threads"
        )
    return _rc.comment_report_multi(files)


@_tool("review")
def get_tracked_changes(file_path: str, author: str | None = None) -> list:
    """Every tracked change: type (insertion, deletion, move, format change),
    author, date, the affected text, and paragraph location. Filter by
    author to see one reviewer's edits. Use resolve_revisions to accept or
    reject, and get_revision_report for the aggregate summary or per-author
    analytics. Read-only.
    """
    return _rd.get_tracked_changes(DocxPackage(file_path), author=author)


@_tool("review")
def resolve_revisions(
    file_path: str,
    action: str,
    author: str | None = None,
    backup: bool = True,
) -> dict:
    """Resolve tracked changes in bulk. action='accept' applies them;
    action='reject' restores the original text. author=None resolves every
    change in the document; naming an author resolves only that reviewer's.
    The result reports how many revisions were processed. Review first with
    get_tracked_changes or get_revision_report; structured_diff compares
    saved drafts instead. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    if action == "accept":
        return _edit(
            file_path, lambda pkg: _rv.accept_revisions(pkg, author=author),
            backup=backup,
        )
    if action == "reject":
        return _edit(
            file_path, lambda pkg: _rv.reject_revisions(pkg, author=author),
            backup=backup,
        )
    raise WordMcpError("action must be 'accept' or 'reject'")


@_tool("review")
def get_revision_report(file_path: str, mode: str = "summary") -> dict:
    """Aggregate views over tracked changes. mode='summary' counts revisions
    by author and type: the quick health check. mode='analytics' adds
    per-author insertion/deletion counts, words added and removed, move and
    format-change counts, the date range of each author's edits, a
    per-section (heading-path) breakdown of where changes concentrate, and
    the 10 heaviest body paragraphs by revision churn; footnote and endnote
    revisions count under their own buckets. Use get_tracked_changes for
    the change-by-change list. Read-only.
    """
    if mode == "summary":
        return _rd.revision_summary(DocxPackage(file_path))
    if mode == "analytics":
        return _rc.revision_analytics(DocxPackage(file_path))
    raise WordMcpError("mode must be 'summary' or 'analytics'")


@_tool("review")
def structured_diff(old_path: str, new_path: str, detail_cap: int = 200) -> dict:
    """Diff two saved drafts agent-readably, computed without Word:
    unchanged/modified/inserted/deleted paragraphs with indices in both
    documents, moved paragraphs, intra-paragraph change opcodes,
    per-section change summary, table changes, footnote/endnote count
    deltas, and heading structure changes. NOT a redline: use
    com_multi_document(action='compare') for a Word-rendered
    tracked-changes comparison. Per-change detail caps at detail_cap
    entries on huge diffs (counts stay complete). Read-only on both files.
    """
    return _rc.structured_diff(old_path, new_path, detail_cap=detail_cap)


@_tool("protection-io")
def redact_text(
    file_path: str,
    targets: list[dict],
    replacement: str = "[REDACTED]",
    scope: str = "all",
    backup: bool = True,
) -> dict:
    """Permanently REMOVE matched text; characters are replaced in the XML,
    not hidden. targets: [{find, regex?}], safe across fragmented runs.
    Scrubs the scope (body | headers | footnotes | all) plus, always:
    comments, properties, hyperlinks, field code/results, tracked-change
    deletions. Reports per-class counts and verified_clean. Irreversible;
    the prev backup slot is the undo. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _rx.redact_text(
            pkg, targets, replacement=replacement, scope=scope
        ),
        backup=backup,
    )


@_tool("review")
def anonymize_for_review(
    file_path: str,
    author_names: list[str],
    replacement: str = "Author",
    mapping_path: str | None = None,
    backup: bool = True,
) -> dict:
    """BETA (heuristic): review the flagged items rather than trusting
    silently. Anonymize a manuscript for double-blind review, reversibly:
    masks the named authors' self-citations, rewrites their reference
    entries, and scrubs identifying metadata. Prose that identifies the
    author is FLAGGED, never auto-edited. Writes a reversal mapping JSON
    for deanonymize_document; KEEP IT PRIVATE. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _an.anonymize_for_review(
            pkg, author_names, replacement=replacement,
            mapping_path=mapping_path,
        ),
        backup=backup,
    )


@_tool("review")
def deanonymize_document(
    file_path: str, mapping_path: str | None = None, backup: bool = True
) -> dict:
    """Reverse anonymize_for_review using its mapping file (default
    <name>.anonymization.json beside the document). Every recorded change
    is verified to still sit where the mapping says before anything is
    restored; if the document drifted, NOTHING is restored and the refusal
    lists every mismatch. The mapping file stays on disk to delete once the
    restore is confirmed. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _an.deanonymize(pkg, mapping_path=mapping_path),
        backup=backup,
    )


# ============================================ 2.5 structure, breaks, sections

_SECTION_BREAKS = {
    "section_next": "nextPage",
    "section_continuous": "continuous",
    "section_even": "evenPage",
    "section_odd": "oddPage",
}


@_tool("lite")
def insert_break(
    file_path: str,
    type: str = "page",
    location: dict | None = None,
    backup: bool = True,
) -> dict:
    """Insert a break after the located paragraph. type: page starts a new
    page; section_next / section_continuous / section_even / section_odd
    start a new SECTION (own headers, margins, numbering; see
    set_section_properties). location picks the paragraph (omit for
    document end). Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    if type != "page" and type not in _SECTION_BREAKS:
        raise WordMcpError(
            f"type must be one of: page | "
            + " | ".join(_SECTION_BREAKS) + f"; got {type!r}"
        )

    def _do(pkg: DocxPackage) -> dict:
        after_index, at_end = _loc_insert_args(pkg, location)
        if at_end:
            after_index = _last_paragraph_index(pkg)
        if type == "page":
            return _tx.add_page_break(pkg, after_index=after_index)
        return _fu.add_section_break(
            pkg, after_index=after_index, break_type=_SECTION_BREAKS[type]
        )

    return _edit(file_path, _do, backup=backup)


@_tool("academic")
def change_heading_level(
    file_path: str,
    delta: int,
    heading_text: str | None = None,
    paragraph_index: int | None = None,
    subtree: bool = False,
    backup: bool = True,
) -> dict:
    """Promote (delta=-1) or demote (delta=1) a heading, addressed by exact
    heading_text or paragraph_index. subtree=True shifts every subordinate
    heading by the same delta, keeping the branch's shape. Refuses,
    changing nothing, when a level would leave 1-9 or comes from an
    outlineLvl override (use set_paragraph_format's outline_level).
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.change_heading_level(
            pkg, delta=delta, heading_text=heading_text,
            paragraph_index=paragraph_index, subtree=subtree,
        ),
        backup=backup,
    )


@_tool("assembly")
def move_section(
    file_path: str,
    heading_text: str,
    before_heading: str | None = None,
    after_heading: str | None = None,
    at_end: bool = False,
    backup: bool = True,
) -> dict:
    """Move a heading and its ENTIRE section (content and tables until the
    next same-or-higher heading) before or after another heading, or to the
    end. Sections are addressed by exact heading text; only headings match,
    so body prose repeating the words does not collide. Refuses moves that
    would cut fields or section breaks. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sx.move_section(
            pkg, heading_text, before_heading=before_heading,
            after_heading=after_heading, at_end=at_end,
        ),
        backup=backup,
    )


@_tool("academic")
def set_section_properties(
    file_path: str,
    section: int = 0,
    orientation: str | None = None,
    page_width_pt: float | None = None,
    page_height_pt: float | None = None,
    margins_pt: dict | None = None,
    columns: dict | None = None,
    line_numbering: dict | str | None = None,
    backup: bool = True,
) -> dict:
    """Set one section's page size, orientation, margins, multi-column layout
    (columns: {count, space_pt, separator, widths_pt}), and manuscript line
    numbering (line_numbering: {count_by, start, restart, distance_pt};
    "none" removes it). Every response carries the section's full state;
    call with no change parameters to read current values. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    if columns is not None:
        unknown = sorted(
            set(columns) - {"count", "space_pt", "separator", "widths_pt"}
        )
        if unknown:
            raise WordMcpError(
                f"columns got unknown key(s) {unknown}; it takes count, "
                "space_pt, separator, widths_pt"
            )
    if isinstance(line_numbering, str) and line_numbering != "none":
        raise WordMcpError(
            'line_numbering takes {count_by, start, restart, distance_pt} '
            'or "none" (removes)'
        )
    if isinstance(line_numbering, dict):
        unknown = sorted(
            set(line_numbering)
            - {"count_by", "start", "restart", "distance_pt"}
        )
        if unknown:
            raise WordMcpError(
                f"line_numbering got unknown key(s) {unknown}; it takes "
                "count_by, start, restart, distance_pt"
            )

    def _do(pkg: DocxPackage) -> dict:
        if columns is not None:
            _fu.set_columns(pkg, section=section, **columns)
        if line_numbering == "none":
            _fu.set_line_numbering(pkg, section=section, remove=True)
        elif isinstance(line_numbering, dict):
            _fu.set_line_numbering(pkg, section=section, **line_numbering)
        result = _fu.set_section_properties(
            pkg, section=section, orientation=orientation,
            page_width_pt=page_width_pt, page_height_pt=page_height_pt,
            margins_pt=margins_pt,
        )
        result["columns_set"] = columns is not None
        result["line_numbering_set"] = line_numbering is not None
        return result

    return _edit(file_path, _do, backup=backup)


@_tool("lite")
def insert_list(
    file_path: str,
    items: list,
    kind: str = "bullet",
    location: dict | None = None,
    backup: bool = True,
) -> dict:
    """Insert a bulleted or numbered list with real bullet/number glyphs
    (numbering.xml infrastructure created as needed). items: strings or
    {text, level} dicts (level 0-8 nests); kind: bullet | number. Each call
    is an independent list, so numbering restarts at 1. location picks the
    insertion point (omit for document end). Auto-backup: prev/anchor slots
    in .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        after_index, at_end = _loc_insert_args(pkg, location)
        return _ls.add_list(
            pkg, items, kind=kind, after_index=after_index, at_end=at_end
        )

    return _edit(file_path, _do, backup=backup)


@_tool("media-forms")
def set_textbox_text(
    file_path: str, box_index: int, text: str, backup: bool = True
) -> dict:
    """Replace the text of one text box (box_index from
    get_text(textbox=true)). Keeps the first paragraph's style and first
    run's formatting; '\\n' splits into multiple paragraphs; the
    mc:Fallback compatibility copy is rewritten to match. Boxes holding
    non-text content (nested tables, images) are refused rather than
    flattened. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tbx.set_textbox_text(pkg, box_index, text),
        backup=backup,
    )


# ======================================== 2.6 headers, footers, page numbers


@_tool("academic")
def set_header_footer(
    file_path: str,
    part: str,
    text: str,
    section: int = 0,
    ref_type: str = "default",
    alignment: str = "center",
    include_page_number: bool = False,
    backup: bool = True,
) -> dict:
    """Set a section's header or footer text. part: header | footer;
    ref_type: default | first | even picks which part variant;
    include_page_number=True appends a PAGE field (the v1 footer option).
    For per-chapter running headers use setup_chapter_headers; for page
    numbers alone use set_page_numbers. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    if part not in ("header", "footer"):
        raise WordMcpError(f"part must be 'header' or 'footer', got {part!r}")
    return _edit(
        file_path,
        lambda pkg: _fu.set_header_footer(
            pkg, part, text, section=section, ref_type=ref_type,
            alignment=alignment, include_page_number=include_page_number,
        ),
        backup=backup,
    )


@_tool("academic")
def set_page_numbers(
    file_path: str,
    section: int = 0,
    position: str | None = None,
    alignment: str = "center",
    prefix: str = "",
    x_of_y: bool = False,
    format: dict | None = None,
    backup: bool = True,
) -> dict:
    """One tool owns page numbering. position: header | footer inserts a PAGE
    field (x_of_y renders 'Page N of M', prefix prepends text); format:
    {number_format, start_at} sets the per-section numbering format
    (lowerRoman front matter, decimal restarting at 1 for the body). Give
    either or both. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    if position is None and format is None:
        raise WordMcpError(
            "give position to insert page numbers, format to set their "
            "format, or both"
        )
    fmt = format or {}
    unknown = sorted(set(fmt) - {"number_format", "start_at"})
    if unknown:
        raise WordMcpError(
            f"format got unknown key(s) {unknown}; it takes number_format "
            "and start_at"
        )
    if position is None and fmt and "number_format" not in fmt:
        raise WordMcpError(
            "start_at alone is ambiguous; give number_format or position "
            "too"
        )

    def _do(pkg: DocxPackage) -> dict:
        out: dict = {}
        if position is not None:
            out.update(_fu.add_page_numbers(
                pkg, section=section, position=position,
                alignment=alignment, prefix=prefix,
                start_at=fmt.get("start_at"), x_of_y=x_of_y,
            ))
        if fmt.get("number_format") is not None:
            out.update(_fu.set_page_number_format(
                pkg, section=section, number_format=fmt["number_format"],
                start_at=fmt.get("start_at"),
            ))
        return out

    return _edit(file_path, _do, backup=backup)


@_tool("academic")
def setup_chapter_headers(
    file_path: str,
    level: int = 1,
    include_number: bool = False,
    alignment: str = "right",
    first_page_blank: bool = True,
    scope: str | list = "auto",
    backup: bool = True,
) -> dict:
    """Put the current chapter title in the running header via a STYLEREF
    field on the Heading style of `level`: Word's standard mechanism,
    evaluated per page, no per-chapter section breaks needed.
    include_number adds the heading number; scope 'auto' targets every
    section with such headings; first_page_blank sets titlePg. Watermarks
    survive; other header content is replaced. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _ch.setup_chapter_headers(
            pkg, level=level, include_number=include_number,
            alignment=alignment, first_page_blank=first_page_blank,
            scope=scope,
        ),
        backup=backup,
    )


# ===================== 2.7 images, charts, equations, and inline objects


@_tool("media-forms")
def insert_image(
    file_path: str,
    image_path: str,
    location: dict | None = None,
    width_pt: float | None = None,
    alignment: str = "center",
    backup: bool = True,
) -> dict:
    """Insert an inline image (PNG/JPEG/GIF/BMP/TIFF) in its own paragraph at
    the located position (omit location for document end); aspect ratio is
    kept, width defaults to the native size capped at 6.5in. The returned
    image is addressable afterwards via list_elements(type='images') ids
    for set_image and delete_element. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        after_index, at_end = _loc_insert_args(pkg, location)
        return _md.add_image(
            pkg, image_path, after_index=after_index, at_end=at_end,
            width_pt=width_pt, alignment=alignment,
        )

    return _edit(file_path, _do, backup=backup)


@_tool("media-forms")
def set_image(
    file_path: str,
    image_id: int,
    source: str | None = None,
    width_pt: float | None = None,
    alt_text: str | None = None,
    alt_title: str | None = None,
    backup: bool = True,
) -> dict:
    """Set properties on one image (image_id from
    list_elements(type='images')): source swaps the image file keeping
    placement and display size (same file type as the original); width_pt
    resizes with aspect ratio kept; alt_text (plus optional alt_title) sets
    accessibility alt text. Parameters are complementary; give any
    combination in one call. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    if source is None and width_pt is None and alt_text is None:
        raise WordMcpError(
            "give source, width_pt, and/or alt_text; nothing to set"
        )
    if alt_title is not None and alt_text is None:
        raise WordMcpError("alt_title goes with alt_text")

    def _do(pkg: DocxPackage) -> dict:
        changed: dict = {}
        if source is not None:
            changed.update(_md.replace_image(pkg, image_id, source))
        if width_pt is not None:
            changed.update(_md.resize_image(pkg, image_id, width_pt=width_pt))
        if alt_text is not None:
            changed.update(_sx.set_image_alt_text(
                pkg, image_id, description=alt_text, title=alt_title
            ))
        return changed

    return _edit(file_path, _do, backup=backup)


@_tool("media-forms")
def export_images(
    file_path: str, output_dir: str, prefix: str | None = None
) -> dict:
    """Export every image in the document to files in output_dir, named by
    list_elements(type='images') index plus the original extension
    (image0.png, ...; prefix replaces 'image'). Media referenced only from
    headers/footers/notes gets the next indices with the referencing part
    reported. Each entry reports the output file, native pixel dimensions,
    and where the image appears. Collisions with existing files are refused
    before anything is written. Read-only against the document.
    """
    return _dio.extract_images(DocxPackage(file_path), output_dir,
                               prefix=prefix)


@_tool("media-forms")
def insert_chart(
    file_path: str,
    chart_type: str,
    data: list | dict | str,
    title: str | None = None,
    width_pt: float | None = None,
    height_pt: float | None = None,
    location: dict | None = None,
    alignment: str = "center",
    legend: bool = True,
    colors: list | None = None,
    backup: bool = True,
) -> dict:
    """Insert a native, theme-following Word chart from data. chart_type: bar
    | column | line | pie | scatter. data: rows (first row series names,
    first column categories), a {categories, series} dict (scatter: x/y per
    series), or a .csv/.json path. Renders everywhere and Edit Data works
    in Word. colors: one hex per series; other chart types and ragged data
    refused. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """

    def _do(pkg: DocxPackage) -> dict:
        after_index, at_end = _loc_insert_args(pkg, location)
        return _charts.add_chart(
            pkg, chart_type, data, title=title, width_pt=width_pt,
            height_pt=height_pt, after_index=after_index, at_end=at_end,
            alignment=alignment, legend=legend, colors=colors,
        )

    return _edit(file_path, _do, backup=backup)


@_tool("media-forms")
def set_chart_data(
    file_path: str,
    chart_id: int,
    data: list | dict | str,
    series_names: list | None = None,
    backup: bool = True,
) -> dict:
    """Replace a bar/column/line/pie/scatter chart's data in place (chart_id
    from list_elements(type='charts')): caches, range formulas, and the
    embedded workbook are rewritten together, formatting preserved. data
    takes insert_chart's shapes and must keep the series COUNT (point count
    may change); series_names renames. Chartex/modern, combo, 3D, and
    series-count changes are refused. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _charts.update_chart_data(
            pkg, chart_id, data, series_names=series_names
        ),
        backup=backup,
    )


@_tool("media-forms")
def insert_equation(
    file_path: str,
    latex: str,
    display: bool = True,
    location: dict | None = None,
    anchor_text: str | None = None,
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Insert a LaTeX equation as NATIVE, editable Word math. display=True: a
    block equation at the located position; display=False: inline right
    after anchor_text. Covers fractions, roots, sums/integrals, matrices,
    cases, align*, Greek, \\text{}. Unconvertible LaTeX refuses, document
    unmodified; read equations with list_elements(type='equations'), never
    get_text. Auto-backup: prev/anchor slots in .ks4w-backups (backup=False
    skips rotation only); atomic validated save. Refuses documents open in
    Word.
    """
    if display:
        def _do(pkg: DocxPackage) -> dict:
            after_index, at_end = _loc_insert_args(pkg, location)
            return _eq.add_equation(
                pkg, latex, display=True, after_index=after_index,
                at_end=at_end,
            )

        return _edit(file_path, _do, backup=backup)
    if anchor_text is None:
        raise WordMcpError(
            "display=False (inline) requires anchor_text: the equation "
            "lands right after it within that paragraph"
        )
    return _edit(
        file_path,
        lambda pkg: _eq.add_equation(
            pkg, latex, display=False, anchor_text=anchor_text,
            occurrence=occurrence,
        ),
        backup=backup,
    )


@_tool("academic")
def insert_bookmark(
    file_path: str, name: str, anchor_text: str, occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Bookmark a text range by name: the target insert_cross_reference
    points at, a stable {"bookmark": name} address for every
    location-taking tool, and the anchor internal hyperlinks jump to.
    Names: letters, digits, underscore, starting with a letter, max 40
    chars; duplicates refused. Remove later with
    delete_element(type='bookmark'). Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.add_bookmark(
            pkg, name, anchor_text=anchor_text, occurrence=occurrence
        ),
        backup=backup,
    )


@_tool("media-forms")
def insert_hyperlink(
    file_path: str, anchor_text: str, url: str, occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Turn existing text (the occurrence-th match of anchor_text) into an
    external hyperlink styled with the Hyperlink character style; text
    already inside a link is refused. To remove a link but keep its text,
    use delete_element(type='hyperlink'). Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.add_hyperlink(
            pkg, anchor_text=anchor_text, url=url, occurrence=occurrence
        ),
        backup=backup,
    )


@_tool("academic")
def insert_caption(
    file_path: str,
    text: str,
    table_index: int | None = None,
    after_anchor: str | None = None,
    label: str = "Table",
    above: bool = True,
    backup: bool = True,
) -> dict:
    """Insert a numbered caption (SEQ field), 'Table N: text', above or below
    a table or at an anchor. label: Table | Figure | Equation; Word
    recomputes the numbers on field update, and insert_reference_list can
    build a List of Tables/Figures from these entries. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fl.add_caption(
            pkg, table_index=table_index, after_anchor=after_anchor,
            label=label, text=text, above=above,
        ),
        backup=backup,
    )


_DELETE_TYPES = (
    "equation", "image", "chart", "bookmark", "hyperlink", "caption",
    "reference_list", "content_control",
)


def _int_id(type_name: str, id) -> int:
    if isinstance(id, bool) or not isinstance(id, int):
        raise WordMcpError(
            f"type {type_name!r} takes an int id: its "
            "list_elements index"
        )
    return id


@_tool("lite")
def delete_element(
    file_path: str,
    type: str,
    id: int | str | None = None,
    location: dict | None = None,
    backup: bool = True,
) -> dict:
    """Delete one document element that has no lifecycle tool of its own.
    type: equation | image | chart | bookmark | hyperlink | caption |
    reference_list | content_control. Address equations, images, charts by
    their list_elements index; reference_list (TOC, List of
    Tables/Figures, index) by its read order; bookmarks by name (markers
    go, the text stays; Word-internal underscore names refused); hyperlinks
    by target url in `id` and/or a location search on the link text (the
    link is unwrapped, display text kept, the relationship dropped once
    unshared; several matches without occurrence refuse loudly listing
    every candidate); captions by location (Caption-styled paragraphs
    only); content controls by tag or index (the whole control including
    content; locked controls refused). Deleting an image or chart also
    removes its media, chart, and embedded-workbook parts once nothing else
    references them. Objects with lifecycle tools stay there: notes in
    manage_note, comments in manage_comment, sources in manage_source,
    tables in delete_table, paragraphs in delete_paragraphs. Auto-backup:
    prev/anchor slots in .ks4w-backups (backup=False skips rotation only);
    atomic validated save. Refuses documents open in Word.
    """
    if type not in _DELETE_TYPES:
        raise WordMcpError(
            f"unknown type {type!r}; one of: " + " | ".join(_DELETE_TYPES)
        )
    if type in ("equation", "image", "chart", "reference_list", "bookmark",
                "content_control") and location is not None:
        raise WordMcpError(
            f"type {type!r} is id-addressed; location applies to "
            "hyperlink and caption only"
        )
    if type == "equation":
        idx = _int_id(type, id)
        return _edit(file_path, lambda pkg: _eq.delete_equation(pkg, idx),
                     backup=backup)
    if type == "image":
        idx = _int_id(type, id)
        return _edit(file_path, lambda pkg: _md.delete_image(pkg, idx),
                     backup=backup)
    if type == "chart":
        idx = _int_id(type, id)
        return _edit(file_path, lambda pkg: _charts.delete_chart(pkg, idx),
                     backup=backup)
    if type == "reference_list":
        idx = _int_id(type, id)
        return _edit(file_path, lambda pkg: _tc.delete_toc(pkg, which=idx),
                     backup=backup)
    if type == "bookmark":
        if not isinstance(id, str) or not id:
            raise WordMcpError("type 'bookmark' takes id: the bookmark name")
        return _edit(file_path, lambda pkg: _fl.delete_bookmark(pkg, id),
                     backup=backup)
    if type == "hyperlink":
        if id is None and location is None:
            raise WordMcpError(
                "type 'hyperlink' takes id (the target url) and/or "
                "location ({'search': {'text': ..., 'occurrence'?: n}} on "
                "the link text)"
            )
        if id is not None and not isinstance(id, str):
            raise WordMcpError(
                "type 'hyperlink' takes a string id: the target url "
                "(internal anchors as '#name')"
            )
        anchor_text = None
        occurrence = None
        if location is not None:
            search = location.get("search") if isinstance(location, dict) \
                else None
            if not isinstance(search, dict) or "text" not in search:
                raise WordMcpError(
                    "hyperlink location takes the search selector: "
                    "{'search': {'text': ..., 'occurrence'?: n}}"
                )
            anchor_text = search["text"]
            occurrence = search.get("occurrence")
        return _edit(
            file_path,
            lambda pkg: _fl.delete_hyperlink(
                pkg, anchor_text=anchor_text, url=id, occurrence=occurrence
            ),
            backup=backup,
        )
    if type == "caption":
        if location is None:
            raise WordMcpError(
                "type 'caption' is addressed by location (a Caption-styled "
                "paragraph)"
            )
        if id is not None:
            raise WordMcpError("type 'caption' has no id scheme; use location")

        def _do(pkg: DocxPackage) -> dict:
            r = resolve_location(pkg, location)
            paras = pkg.body().findall(qn("w:p"))
            if r.paragraph_index >= len(paras):
                raise TargetNotFound(
                    f"paragraph {r.paragraph_index} does not exist"
                )
            p = paras[r.paragraph_index]
            pstyle = p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
            if pstyle is None or pstyle.get(qn("w:val")) != "Caption":
                raise WordMcpError(
                    f"paragraph {r.paragraph_index} is not Caption-styled; "
                    "refusing to delete it as a caption (use "
                    "delete_paragraphs for prose)"
                )
            text = _rd.paragraph_text(p)
            p.getparent().remove(p)
            pkg.mark_dirty()
            return {"deleted_caption": r.paragraph_index, "text": text}

        return _edit(file_path, _do, backup=backup)
    # content_control
    if id is None:
        raise WordMcpError(
            "type 'content_control' takes id: a tag (string) or a "
            "list_elements(type='content_controls') index (int)"
        )
    if isinstance(id, str):
        kw = {"tag": id}
    else:
        kw = {"index": _int_id(type, id)}
    return _edit(
        file_path,
        lambda pkg: _fm.delete_content_control(pkg, **kw),
        backup=backup,
    )


# ============================= 2.13 templates, forms, and content controls


@_tool("assembly")
def apply_template(
    file_path: str,
    reference_path: str,
    include_page_geometry: bool = True,
    backup: bool = True,
) -> dict:
    """Restyle this document to match a reference document: styles, theme,
    fonts, layout settings, optionally page geometry. This is also the
    import-styles-from-another-document tool: style definitions are
    remapped by name, unmatched styles preserved, content untouched.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _tp.apply_template(
            pkg, reference_path, include_page_geometry=include_page_geometry
        ),
        backup=backup,
    )


@_tool("assembly")
def fill_template(
    file_path: str,
    data: dict,
    missing: str = "error",
    backup: bool = True,
) -> dict:
    """Fill a template IN PLACE: replace every {{name}} with data[name] (safe
    across fragmented runs, first run's formatting kept) and set
    MERGEFIELDs as plain text. missing: 'error' refuses when a key is
    absent; 'skip' leaves markers; 'empty' fills blank. Names come from
    list_elements(type='template_placeholders'); for filled COPIES use
    mail_merge. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _mm.fill_template(pkg, data, missing=missing),
        backup=backup,
    )


@_tool("assembly")
def mail_merge(
    template_path: str,
    data_rows: list[dict] | str,
    output_dir: str,
    filename_pattern: str = "{row_index}.docx",
    missing: str = "error",
) -> dict:
    """Mail merge: one filled .docx per data row, saved into output_dir.
    data_rows: a list of dicts, or a path to a .csv (header = placeholder
    names) or .json (array of objects). filename_pattern supports
    {row_index} (1-based) and any {column}. Refuses BEFORE writing anything
    on existing-file collisions, duplicate output names, or
    (missing='error') rows lacking values the template needs. The template
    file is never modified.
    """
    return _mm.mail_merge(
        template_path, data_rows, output_dir,
        filename_pattern=filename_pattern, missing=missing,
    )


@_tool("media-forms")
def set_form_fields(
    file_path: str,
    values: dict,
    missing: str = "error",
    backup: bool = True,
) -> dict:
    """Set form-field values by name (legacy fields) or tag/alias (content
    controls): text fields take strings, checkboxes booleans, dropdowns
    only their options. Duplicate names refused. missing: 'error' refuses,
    changing nothing, when a key matches no field; 'skip' reports them.
    Fields come from list_elements(type='form_fields'); for skipped SDT
    types use set_content_control. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fm.fill_form_fields(pkg, values, missing=missing),
        backup=backup,
    )


@_tool("media-forms")
def insert_content_control(
    file_path: str,
    tag: str,
    after_anchor: str,
    alias: str | None = None,
    text: str = "",
    occurrence: int = 1,
    backup: bool = True,
) -> dict:
    """Insert a new PLAIN-TEXT content control (inline SDT) right after
    after_anchor text (occurrence picks the match), with a unique tag
    (refused if it exists) and optional alias and initial text. Plain text
    is the one control type built safely; checkbox, dropdown, date,
    picture, and repeating controls are refused. Fill it later via
    set_form_fields or set_content_control by tag. Auto-backup: prev/anchor
    slots in .ks4w-backups (backup=False skips rotation only); atomic
    validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fm.insert_content_control(
            pkg, tag=tag, after_anchor=after_anchor, alias=alias, text=text,
            occurrence=occurrence,
        ),
        backup=backup,
    )


@_tool("media-forms")
def set_content_control(
    file_path: str,
    value: str | bool,
    tag: str | None = None,
    index: int | None = None,
    backup: bool = True,
) -> dict:
    """Set one content control's value, addressed by tag or by
    list_elements(type='content_controls') index (exactly one). Text,
    rich-text, combo, and date controls take a string; checkboxes a
    boolean; dropdowns one of their options (refused otherwise, naming
    them). Locked controls and unwritable types (gallery, repeating
    section, citation, picture, group) refuse with nothing changed.
    Auto-backup: prev/anchor slots in .ks4w-backups (backup=False skips
    rotation only); atomic validated save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _fm.set_content_control_value(
            pkg, value, tag=tag, index=index
        ),
        backup=backup,
    )


@_tool("academic")
def apply_manuscript_format(
    file_path: str,
    style: str,
    running_head: str | None = None,
    author_last_name: str | None = None,
    backup: bool = True,
) -> dict:
    """Apply a style's page-level manuscript conventions where publicly
    well-defined: apa7 (= apa7-student), apa7-professional (adds the
    running head), mla9 (surname + page header), chicago17 (= turabian).
    All set a hanging indent on a found reference list. Contested choices
    are NOT applied, itemized in not_applied; IEEE/Vancouver/ASA/Harvard
    are journal-specific and refused. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    return _edit(
        file_path,
        lambda pkg: _sc2.apply_manuscript_format(
            pkg, style, running_head=running_head,
            author_last_name=author_last_name,
        ),
        backup=backup,
    )


# ================================== 2.14 protection and watermarks


@_tool("protection-io")
def set_document_protection(
    file_path: str,
    protection: str = "trackedChanges",
    password: str | None = None,
    restrict_formatting: bool = False,
    backup: bool = True,
) -> dict:
    """Restrict editing, or lift the restriction. protection: readOnly |
    comments | trackedChanges | forms | none (removal, the set-to-none
    idiom). trackedChanges forces every recipient edit to be tracked (the
    send-to-committee mode). Password hashing is Word-compatible SHA-512,
    NOT encryption (com_save_document's password encrypts); read state with
    get_protection. Auto-backup: prev/anchor slots in .ks4w-backups
    (backup=False skips rotation only); atomic validated save. Refuses
    documents open in Word.
    """
    if protection == "none":
        if password is not None or restrict_formatting:
            raise WordMcpError(
                "protection='none' removes protection; drop "
                "password/restrict_formatting"
            )
        return _edit(
            file_path, lambda pkg: _pr.remove_document_protection(pkg),
            backup=backup,
        )
    return _edit(
        file_path,
        lambda pkg: _pr.set_document_protection(
            pkg, edit=protection, password=password,
            restrict_formatting=restrict_formatting,
        ),
        backup=backup,
    )


@_tool("protection-io")
def set_watermark(
    file_path: str,
    watermark: dict | str,
    backup: bool = True,
) -> dict:
    """Set or remove the page watermark. watermark: {text, color, opacity,
    diagonal} draws a text watermark behind the text on every page (all
    header parts; compatible with Word's own Remove Watermark command);
    "none" removes existing watermark shapes from every header. Setting a
    new watermark replaces the old one. Auto-backup: prev/anchor slots in
    .ks4w-backups (backup=False skips rotation only); atomic validated
    save. Refuses documents open in Word.
    """
    if watermark == "none":
        return _edit(file_path, lambda pkg: _fu.remove_watermark(pkg),
                     backup=backup)
    if isinstance(watermark, dict) and "image" in watermark:
        raise UnsupportedStructure(
            "image watermarks are not supported in v2.0 (no ops builder "
            "exists); the dict form is text-only"
        )
    if isinstance(watermark, dict) and "text" in watermark:
        unknown = sorted(
            set(watermark) - {"text", "color", "opacity", "diagonal"}
        )
        if unknown:
            raise WordMcpError(
                f"watermark got unknown key(s) {unknown}; it takes text, "
                "color, opacity, diagonal"
            )
        return _edit(
            file_path,
            lambda pkg: _fu.add_watermark(
                pkg, watermark["text"],
                color=watermark.get("color", "silver"),
                opacity=watermark.get("opacity", 0.5),
                diagonal=watermark.get("diagonal", True),
            ),
            backup=backup,
        )
    raise WordMcpError('watermark takes {"text": ...} or "none"')


# ================================================= 2.15 COM tier (Word app)


@_tool("com-live")
def com_word_status() -> dict:
    """Check whether Word is running and what state it is in:
    interactive_state (ready, busy if a dialog is open, blocked if a long
    operation is running, not_running), and a list of open documents with
    per-document dirty flag and autosave state. Use before live editing to
    confirm Word is responsive, or to discover which files are locked, and
    before com_save_document to see what needs saving. No file_path needed.
    Read-only.
    """
    from .com import bridge, live

    out = bridge.word_status()
    status = live.interactive_status()
    out["interactive_state"] = status["interactive_state"]
    if status["open_documents"]:
        out["open_documents"] = status["open_documents"]
    if status.get("protected_view_documents"):
        out["protected_view_documents"] = status["protected_view_documents"]
    return out


@_tool("com-live")
def com_refresh_fields(file_path: str) -> dict:
    """Update every field in the document (TOC page numbers, PAGEREF,
    NUMPAGES, SEQ, cross-references) via an invisible Word instance, giving
    correct page numbers and computed values immediately. Use after
    inserting a TOC, index, or caption list for real page numbers, or after
    any edit that shifts pages. The source file is modified in place and
    saved by the invisible instance; open-in-Word copies are untouched
    (Option C: COM saves only what it is asked to). Requires Word
    installed.
    """
    from .com import bridge

    return bridge.refresh_fields(file_path)


@_tool("com-live")
def com_export_pdf(file_path: str, pdf_path: str | None = None) -> dict:
    """Export the document to PDF via an invisible Word instance with full
    fidelity (fields resolved, footnotes, TOC, headers/footers, images).
    pdf_path defaults to the source filename with .pdf extension in the
    same directory. The source .docx is never modified. Requires Word
    installed. Read-only.
    """
    from .com import bridge

    return bridge.export_pdf(file_path, pdf_path)


@_tool("com-live")
def com_import_pdf(pdf_path: str, output_path: str | None = None) -> dict:
    """Convert a PDF to .docx via Word's built-in PDF reflow, in a dedicated
    invisible Word instance. Output defaults next to the PDF with a .docx
    extension; an existing output file is refused; the produced .docx is
    validated by a full package round-trip. Text-based PDFs convert well,
    complex layouts may reflow imperfectly, and scanned image PDFs yield
    little or no text (Word does not OCR; a near-zero word count triggers
    an explicit warning). Requires Word installed.
    """
    from .com import convert

    return convert.import_pdf(pdf_path, output_path)


@_tool("com-live")
def com_multi_document(
    action: str,
    files: list[str],
    output_path: str | None = None,
    author: str = "word-mcp compare",
    section_break_between: bool = True,
) -> dict:
    """Word-native multi-document operations through an invisible instance.
    action=compare diffs two versions of a draft (files=[original,
    revised]): a NEW document where every difference is a tracked change,
    plus a revision summary; author names the change author; output
    defaults to <revised>_COMPARE.docx. action=combine merges two
    reviewers' tracked changes into one unified redline (files=[original,
    revised]), both sets of revisions keeping their original attributions;
    use compare to discover differences, combine to pool existing markup;
    output defaults beside the original. action=merge concatenates two or
    more whole documents in order into ONE new file with full fidelity
    (styles, footnotes, numbering); section_break_between keeps per-chapter
    headers and numbering possible; output_path is required. To insert a
    document INTO another at a chosen position use insert_document; for one
    table, copy_table. Input files are never modified; the result is a new
    file written to disk. Requires Word installed; documents need not be
    open.
    """
    from .com import bridge

    if not isinstance(files, list) or not files:
        raise WordMcpError("files takes a non-empty list of .docx paths")
    if action == "compare":
        if len(files) != 2:
            raise WordMcpError(
                "compare takes files=[original, revised] (exactly two)"
            )
        if section_break_between is not True:
            raise WordMcpError("section_break_between applies to merge only")
        return bridge.compare_documents(
            files[0], files[1], output_path, author=author
        )
    if action == "combine":
        if len(files) != 2:
            raise WordMcpError(
                "combine takes files=[original, revised] (exactly two)"
            )
        if author != "word-mcp compare":
            raise WordMcpError("author applies to compare only")
        if section_break_between is not True:
            raise WordMcpError("section_break_between applies to merge only")
        return bridge.combine_documents(files[0], files[1], output_path)
    if action == "merge":
        if len(files) < 2:
            raise WordMcpError("merge takes two or more files")
        if author != "word-mcp compare":
            raise WordMcpError("author applies to compare only")
        if output_path is None:
            raise WordMcpError("merge requires output_path")
        return bridge.merge_documents(
            files, output_path, section_break_between=section_break_between
        )
    raise WordMcpError("action must be compare | combine | merge")


@_tool("com-live")
def com_save_document(
    file_path: str,
    save: bool = True,
    close: bool = False,
    password: str | None = None,
    output_path: str | None = None,
) -> dict:
    """Save, close, or encrypt a document through Word. Default: tells the
    user's running Word to save the open document so file tools read the
    current state (the explicit save step of the Option C model: live and
    COM edits stay unsaved until requested here or by the user). close=True
    also closes it, releasing the file lock (save=False discards unsaved
    changes). password saves a real AES-encrypted copy via an invisible
    instance (output_path saves it elsewhere; not combinable with close).
    Requires Word installed.
    """
    from .com import bridge

    if password is not None:
        if close or not save:
            raise WordMcpError(
                "password mode saves an encrypted copy via an invisible "
                "instance; it does not combine with close or save=False"
            )
        return bridge.save_with_password(file_path, output_path,
                                         password=password)
    if output_path is not None:
        raise WordMcpError("output_path applies to password mode only")
    if close:
        return bridge.close_open_document(file_path, save=save)
    if not save:
        raise WordMcpError(
            "nothing to do: save=False without close and without password"
        )
    return bridge.save_open_document(file_path)


@_tool("com-live")
def com_proofing_errors(file_path: str, limit: int = 100) -> dict:
    """Word's own spelling and grammar error lists with surrounding context:
    each error, its type (spelling or grammar), the sentence containing it,
    and suggested corrections. Useful as a review aid before submission
    (note: proper nouns and technical terms appear as spelling errors).
    Requires Word installed; opens an invisible instance and modifies
    nothing. The source file is untouched. Read-only.
    """
    from .com import bridge

    return bridge.proofing_errors(file_path, limit=limit)


@_tool("com-live")
def com_readability_statistics(file_path: str) -> dict:
    """Word's own readability statistics via COM: Flesch Reading Ease,
    Flesch-Kincaid Grade Level, word, sentence, and paragraph counts, and
    averages. Requires Word installed; opens an invisible instance to
    compute the statistics and modifies nothing; the source file is
    untouched. Read-only.
    """
    from .com import bridge

    return bridge.readability_statistics(file_path)


@_tool("com-live")
def com_validate_opens_clean(file_path: str) -> dict:
    """Run the definitive corruption check: open the file in an invisible
    Word instance and report clean or fail, with Word's own error where one
    is raised. The Word-verdict companion to validate and
    diagnose_document; use it when XML-level checks pass but Word still
    complains. Requires Word installed; modifies nothing. Read-only.
    """
    from .com import bridge

    return bridge.validate_opens_clean(file_path)


# ============================================ 2.15 live tier (visible Word)


@_tool("com-live")
def live_insert_at_cursor(
    file_path: str, text: str, newline: bool = False
) -> dict:
    """Insert text at the user's cursor position in the open document (main
    body text only; headers, footers, and footnotes not supported). The
    cursor position is read once and never moved; the user's selection is
    untouched. newline=True ends the insertion with a paragraph break. Use
    insert_paragraphs for location-addressed insertion instead. Requires
    the document open in a visible Word; the edit is one undo step and
    stays unsaved until the user saves (Option C).
    """
    from .com import live_ops as _lo

    return _lo.insert_at_cursor(file_path, text, newline=newline)


@_tool("com-live")
def live_scroll_to(
    file_path: str,
    find: str | None = None,
    paragraph_index: int | None = None,
) -> dict:
    """Scroll the user's Word window to show a location without selecting
    anything or moving their cursor (useful after a live edit, to show the
    user what changed). Target by find text (first match, at most 255
    chars) or 0-based body paragraph_index; exactly one is required. The
    document must be open in a visible Word instance. Read-only, nothing is
    modified.
    """
    from .com import live_ops as _lo

    return _lo.scroll_to(file_path, find=find, paragraph_index=paragraph_index)


@_tool("com-live")
def live_set_track_changes(file_path: str, enabled: bool) -> dict:
    """Turn track changes on or off on the open document. This is a
    persistent state change (the document remembers the setting), unlike
    the track flag on edit tools like search_and_replace, which
    auto-restores after the call. Returns the previous state so you can
    restore it later. The document must be open in a visible Word instance;
    the setting change stays unsaved until the user saves or
    com_save_document is called (Option C save semantics).
    """
    from .com import live_ops as _lo

    return _lo.set_track_changes(file_path, enabled)


@_tool("com-live")
def live_repair() -> dict:
    """Recovery tool for the live layer: if a crashed live edit left the
    user's Word frozen (ScreenUpdating off), alerts suppressed, or an undo
    record still recording, a fresh attach fixes all three and reports
    exactly what it repaired. Safe to run anytime, including when nothing
    is wrong; touches no document content and saves nothing. Requires Word
    running with a visible instance.
    """
    from .com import live

    return live.live_repair()


# ------------------------------------------ tiered loading (Section 14)
# Registered LAST: enable_tools' docstring carries the pack menu with
# per-pack token bills computed from the live registry, so every other
# tool must already be registered when this block runs.

_PENDING_VISIBILITY: list[tuple[set[str], bool]] = []


def _record_visibility(names: set[str], enabled: bool) -> None:
    """packs.py visibility hook: bookkeeping flips queue here; the async
    enable_tools/disable_tools bodies apply them session-scoped, and
    main() folds startup flips into the global transform instead."""
    _PENDING_VISIBILITY.append((set(names), enabled))


_packs.set_visibility_hook(_record_visibility)


async def _apply_session_visibility(ctx) -> None:
    """Session-scoped toggles (author ruling 2026-09-02): fastmcp 3.x
    session visibility rules override the startup global transform
    (mark-based semantics, later marks win) and send
    ToolListChangedNotification to this session only. ctx=None
    (in-process callers, unit tests) drains the queue without applying;
    packs bookkeeping stays authoritative for surface_report and the
    DisabledToolSignpost either way."""
    pending = list(_PENDING_VISIBILITY)
    _PENDING_VISIBILITY.clear()
    if ctx is None:
        return
    for names, enabled in pending:
        if enabled:
            await ctx.enable_components(names=names)
        else:
            await ctx.disable_components(names=names)


async def enable_tools(
    packs: list[str], ctx: _Context | None = None
) -> dict:
    result = _packs.enable(packs)
    await _apply_session_visibility(ctx)
    return result


async def disable_tools(
    packs: list[str], ctx: _Context | None = None
) -> dict:
    result = _packs.disable(packs)
    await _apply_session_visibility(ctx)
    return result


_MENU_LINES = "\n".join(
    f"- {name} (~{_packs.pack_cost(name) / 1000:.1f}k): "
    f"{_packs.PACK_SUMMARIES[name]}"
    for name in _packs.pack_names()
)

enable_tools.__doc__ = (
    "Enable optional tool packs mid-session (sessions start lite). "
    "Idempotent; result reports packs enabled, approx tokens added, "
    "new total surface. packs = any combination below or "
    "['everything']; disable_tools reverses it. Packs:\n" + _MENU_LINES
)

disable_tools.__doc__ = (
    "Disable previously enabled tool packs for this session and reclaim "
    "their context; the lite core always stays on. Idempotent. The "
    "result reports the packs just disabled, the approximate tokens "
    "removed, and the remaining surface. packs takes the same names as "
    "enable_tools (its description carries the menu) or ['everything']."
)

enable_tools = _tool("lite")(enable_tools)
disable_tools = _tool("lite")(disable_tools)


def _startup_disabled_names() -> set[str]:
    """Tool names hidden at startup under current packs bookkeeping."""
    return {
        name
        for members in _packs.tool_names().values()
        for name in members
        if not _packs.is_tool_enabled(name)
    }


def main() -> None:
    # KS4W_MODE startup surface: bookkeeping first (a typo in the env
    # fails loudly BEFORE serving), then ONE global visibility transform
    # hiding every tool not enabled at startup. Session rules laid down
    # by enable_tools/disable_tools override this transform. Applied
    # here, not at import, so tests and measure_surface always see the
    # full registry.
    _packs.apply_startup_mode()
    _PENDING_VISIBILITY.clear()  # startup flips ride the global transform
    disabled = _startup_disabled_names()
    if disabled:
        mcp.add_transform(_Visibility(False, names=disabled))
    mcp.run()


if __name__ == "__main__":
    main()
