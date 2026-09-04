"""COM bridge: the operations only a running Word can perform.

Every function opens its own dedicated invisible Word instance (DispatchEx —
never attaches to the user's visible Word), disables alerts, and guarantees
Quit in a finally block, EXCEPT the message-the-visible-instance trio
(word_status, save_open_document, close_open_document), which talk to the
user's Word deliberately and suppress alerts around every call.

Serialization (2026-09-03 stress report): every public function here runs
under the process-wide COM lock (com/serial.py) — via @_serial.serialized,
via _run_bounded's worker, or via a bounded try-acquire for status — so
no bridge call can interleave with a live-layer edit or another bridge
call. Long invisible-instance operations additionally run under a bounded
timeout: on expiry the invisible instance this module spawned (and ONLY
that instance, tracked by PID at DispatchEx time) is terminated, turning
the report's 30-minute silent hang into a clean structured error.
"""

from __future__ import annotations

import contextlib
import functools
import subprocess
import threading
import time
from pathlib import Path

from ..core.errors import (
    DocumentNotFound,
    WordBlocked,
    WordBusy,
    WordMcpError,
)
from ..core.sandbox import check_path
from . import serial as _serial

_WD_ALERTS_NONE = 0
_WD_FORMAT_PDF = 17
_WD_DO_NOT_SAVE = 0
_WD_SAVE = -1

# invisible WINWORD.EXE PIDs spawned by _word(), keyed by spawning thread —
# the timeout kill-switch terminates exactly these, never the user's Word
_INVISIBLE_PIDS: dict[int, set] = {}


def _winword_pids() -> set:
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV",
             "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:
        return set()
    pids = set()
    for ln in result.stdout.splitlines():
        if "WINWORD" not in ln.upper():
            continue
        parts = ln.split('","')
        if len(parts) >= 2:
            with contextlib.suppress(ValueError):
                pids.add(int(parts[1].strip('"')))
    return pids


def _kill_invisible_for_thread(tid) -> bool:
    """Terminate the invisible instance(s) the given worker thread spawned.
    PID-precise: only processes recorded by _word() at DispatchEx time."""
    pids = _INVISIBLE_PIDS.get(tid) or set()
    killed = False
    for pid in pids:
        with contextlib.suppress(Exception):
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F"],
                capture_output=True,
                timeout=15,
            )
            killed = True
    return killed


@contextlib.contextmanager
def _word():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise WordMcpError(
            "pywin32 is not available; COM operations need it"
        ) from exc
    pythoncom.CoInitialize()
    app = None
    tid = threading.get_ident()
    before = _winword_pids()
    try:
        app = win32com.client.DispatchEx("Word.Application")
        created = _winword_pids() - before
        if created:
            _INVISIBLE_PIDS[tid] = created
        app.Visible = False
        app.DisplayAlerts = _WD_ALERTS_NONE
        yield app
    finally:
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit(SaveChanges=_WD_DO_NOT_SAVE)
        _INVISIBLE_PIDS.pop(tid, None)
        pythoncom.CoUninitialize()


def _run_bounded(name: str, timeout: float, fn):
    """Run fn on a worker thread under the COM lock with a hard deadline.

    On expiry: if the worker never got the lock, that is queue contention
    (WordBusy names the running operation). If it got the lock and stalled,
    the invisible instance it spawned is terminated so the COM call errors
    out and the lock is released (WordBlocked)."""
    result: dict = {}
    lock_acquired = threading.Event()
    done = threading.Event()
    worker_tid: list = []

    def worker():
        worker_tid.append(threading.get_ident())
        try:
            with _serial.com_operation(name):
                lock_acquired.set()
                result["value"] = fn()
        except BaseException as exc:  # noqa: BLE001 — re-raised in caller
            result["error"] = exc
        finally:
            lock_acquired.set()
            done.set()

    t = threading.Thread(target=worker, daemon=True, name=f"ks4w-{name}")
    t.start()
    if done.wait(timeout):
        if "error" in result:
            raise result["error"]
        return result["value"]
    if not lock_acquired.is_set():
        snap = _serial.lock_snapshot()
        running = (snap.get("current_op") or {}).get(
            "name", "another COM operation"
        )
        raise WordBusy(
            f"{name} waited {timeout:.0f}s for the COM serialization lock "
            f"({running} is still running); retry when it finishes — "
            "com_word_status reports the running operation"
        )
    killed = _kill_invisible_for_thread(worker_tid[0] if worker_tid else None)
    done.wait(10.0)
    raise WordBlocked(
        f"{name} did not finish within {timeout:.0f}s and was aborted"
        + (" (its invisible Word instance was terminated)" if killed else "")
        + ". The document may be very large, or Word may be stuck — check "
        "com_word_status, and pass a larger timeout to raise the bound."
    )


def _bounded_op(name: str, default: float):
    """Public-function wrapper: adds a tunable timeout parameter and runs
    the body via _run_bounded (which serializes). Marks the function for
    the entry-point coverage audit."""

    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, timeout: float = default, **kwargs):
            try:
                timeout = float(timeout)
            except (TypeError, ValueError):
                raise WordMcpError(
                    f"timeout must be a number of seconds, got {timeout!r}"
                ) from None
            if not 5 <= timeout <= 3600:
                raise WordMcpError(
                    "timeout must be between 5 and 3600 seconds"
                )
            return _run_bounded(name, timeout, lambda: fn(*args, **kwargs))

        wrapper._com_serialized = name
        wrapper._com_bounded = default
        return wrapper

    return deco


def _open_in_running_word(path) -> bool:
    """Is this file open in ANY running Word instance? (primary instance
    scan + ROT file-moniker scan; read-only, best-effort). Guards the
    invisible-open name-collision cascade from the stress report: opening
    a second copy of an open document triggers Word's same-name dialog."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    pythoncom.CoInitialize()
    target = str(Path(path).resolve()).lower()
    try:
        with contextlib.suppress(Exception):
            app = win32com.client.GetActiveObject("Word.Application")
            for doc in app.Documents:
                if doc.FullName.lower() == target:
                    return True
        with contextlib.suppress(Exception):
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning():
                ctx = pythoncom.CreateBindCtx(0)
                try:
                    name = moniker.GetDisplayName(ctx, None)
                except Exception:
                    continue
                if name.lower() == target:
                    return True
        return False
    finally:
        pythoncom.CoUninitialize()


def _open(app, path: Path, *, read_only: bool):
    return app.Documents.Open(
        str(path.resolve()),
        ReadOnly=read_only,
        AddToRecentFiles=False,
        OpenAndRepair=False,
    )


def word_status() -> dict:
    """Is Word installed / running, and which documents are open in the USER's
    Word instance (best effort via the running-object table)? Uses a bounded
    lock acquire so a long-running COM operation cannot hang the status
    probe; when the lock is busy the probe is skipped and says so."""
    out = {"word_running": False, "open_documents": []}
    if not _serial.acquire(timeout=2.0):
        out["note"] = (
            "COM serialization lock held by a running operation; process "
            "state not probed"
        )
        return out
    try:
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("Word.Application")
                out["word_running"] = True
                for doc in app.Documents:
                    out["open_documents"].append(doc.FullName)
            except Exception:
                pass
            finally:
                pythoncom.CoUninitialize()
        except ImportError:
            out["error"] = "pywin32 not installed"
        return out
    finally:
        _serial.release()


word_status._com_serialized = "com_word_status"  # bounded try-acquire form


@_bounded_op("com_refresh_fields", default=300.0)
def refresh_fields(path: str) -> dict:
    """Open invisibly, update every field (TOC page numbers, PAGEREF, SEQ),
    update TOC-family tables explicitly, save, close. This is the immediate
    alternative to the update-on-open flag."""
    path = check_path(path, "refresh fields")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    # Word silently refuses to update fields in a protected document.
    from ..core.package import DocxPackage
    from ..ops.protection import get_protection

    state = get_protection(DocxPackage(p))
    if state.get("protected"):
        raise WordMcpError(
            f"document is protected (edit={state.get('edit')}); Word will not "
            "update fields under protection — remove_document_protection "
            "first, refresh, then re-protect (build → refresh → protect)"
        )
    with _word() as app:
        doc = _open(app, p, read_only=False)
        try:
            for story in doc.StoryRanges:
                story.Fields.Update()
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            for i in range(1, doc.TablesOfFigures.Count + 1):
                doc.TablesOfFigures(i).Update()
            toc_count = doc.TablesOfContents.Count
            doc.Save()
        finally:
            doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    return {"fields_refreshed": True, "tocs_updated": toc_count}


@_bounded_op("com_export_pdf", default=300.0)
def export_pdf(path: str, pdf_path: str | None = None) -> dict:
    """Export to PDF via Word (highest fidelity available on this machine)."""
    path = check_path(path, "PDF export source")
    if pdf_path:
        pdf_path = check_path(pdf_path, "PDF export output")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    out = Path(pdf_path) if pdf_path else p.with_suffix(".pdf")
    with _word() as app:
        doc = _open(app, p, read_only=True)
        try:
            doc.SaveAs2(str(out.resolve()), FileFormat=_WD_FORMAT_PDF)
        finally:
            doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word reported success but no PDF was produced")
    return {"pdf": str(out), "bytes": out.stat().st_size}


@_bounded_op("com_multi_document", default=300.0)
def compare_documents(
    original_path: str,
    revised_path: str,
    output_path: str | None = None,
    *,
    author: str = "word-mcp compare",
) -> dict:
    """Word-native compare: produces a NEW document where every difference
    between original and revised appears as a tracked change. Neither input is
    modified. Perfect for diffing two DTG versions of a draft."""
    original_path = check_path(original_path, "compare original")
    revised_path = check_path(revised_path, "compare revised")
    if output_path:
        output_path = check_path(output_path, "compare output")
    orig = Path(original_path)
    rev = Path(revised_path)
    for p in (orig, rev):
        if not p.exists():
            raise DocumentNotFound(f"no file at {p}")
    out = (
        Path(output_path)
        if output_path
        else rev.with_name(f"{rev.stem}_COMPARE{rev.suffix}")
    )
    with _word() as app:
        doc_a = _open(app, orig, read_only=True)
        doc_b = _open(app, rev, read_only=True)
        try:
            result = app.CompareDocuments(
                OriginalDocument=doc_a,
                RevisedDocument=doc_b,
                Destination=2,  # wdCompareDestinationNew
                Granularity=1,  # wdGranularityWordLevel
                CompareFormatting=True,
                CompareTables=True,
                CompareFootnotes=True,
                CompareHeaders=True,
                CompareFields=False,
                RevisedAuthor=author,
            )
            result.SaveAs2(str(out.resolve()))
            result.Close(SaveChanges=_WD_DO_NOT_SAVE)
        finally:
            doc_a.Close(SaveChanges=_WD_DO_NOT_SAVE)
            doc_b.Close(SaveChanges=_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word reported success but produced no output")
    # Summarize what changed, using our own revision reader.
    from ..core.package import DocxPackage
    from ..ops.read import revision_summary

    summary = revision_summary(DocxPackage(out))
    return {"comparison": str(out), "revisions": summary}


@_bounded_op("com_validate_opens_clean", default=60.0)
def validate_opens_clean(path: str) -> dict:
    """Open in invisible Word and confirm no repair/recovery path triggers.

    Name-collision guard (stress report bug 6): when the file is ALREADY
    open in a running Word, no second invisible copy is opened — the check
    routes to the open document (which by definition opened clean) and
    says so, instead of triggering Word's same-name dialog cascade."""
    path = check_path(path, "validate opens clean")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if _open_in_running_word(p):
        pythoncom, app, doc = _find_open_document(str(p))
        try:
            paragraphs = doc.Paragraphs.Count
            words = int(doc.ComputeStatistics(0))  # wdStatisticWords
        finally:
            pythoncom.CoUninitialize()
        return {
            "opens_clean": True,
            "paragraphs": paragraphs,
            "words": words,
            "note": (
                "document is OPEN in Word: the verdict reflects the open "
                "copy (unsaved changes included); no second invisible copy "
                "was opened, avoiding Word's same-name dialog"
            ),
        }
    with _word() as app:
        try:
            doc = _open(app, p, read_only=True)
        except Exception as exc:
            return {"opens_clean": False, "error": str(exc)}
        try:
            paragraphs = doc.Paragraphs.Count
            # NEVER doc.Words.Count: it counts punctuation runs and paragraph
            # marks as "words" (~18% high on real prose — L1, 2026-08-28).
            # ComputeStatistics(wdStatisticWords) is the number Word's own
            # status bar shows.
            words = int(doc.ComputeStatistics(0))  # wdStatisticWords
        finally:
            doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    return {"opens_clean": True, "paragraphs": paragraphs, "words": words}


@_bounded_op("com_multi_document", default=300.0)
def merge_documents(
    paths: list[str],
    output_path: str,
    *,
    section_break_between: bool = True,
) -> dict:
    """Merge documents in order into one file via Word's InsertFile (full
    fidelity: styles, footnotes, numbering survive). Chapters get a next-page
    section break between them so per-chapter headers/numbering stay possible."""
    if len(paths) < 2:
        raise WordMcpError("give at least two documents to merge")
    paths = [check_path(p, "merge input") for p in paths]
    output_path = check_path(output_path, "merge output")
    srcs = [Path(p) for p in paths]
    for p in srcs:
        if not p.exists():
            raise DocumentNotFound(f"no file at {p}")
    out = Path(output_path)
    with _word() as app:
        doc = app.Documents.Add()
        try:
            rng = doc.Range(0, 0)
            for i, src in enumerate(srcs):
                rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                if i > 0:
                    if section_break_between:
                        rng.InsertBreak(2)  # wdSectionBreakNextPage
                    else:
                        rng.InsertBreak(7)  # wdPageBreak
                    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                rng.InsertFile(str(src.resolve()))
            doc.SaveAs2(str(out.resolve()))
            paragraphs = doc.Paragraphs.Count
        finally:
            doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word reported success but produced no output")
    # Word's InsertFile does not carry the customXml bibliography store;
    # union the inputs' sources into the merged output ourselves.
    sources_merged = _merge_bibliography_stores(out, srcs)
    result = {
        "merged": [str(p) for p in srcs],
        "output": str(out),
        "paragraphs": paragraphs,
    }
    if sources_merged:
        result["bibliography_sources_carried"] = sources_merged
    return result


def _merge_bibliography_stores(out_path, src_paths) -> int:
    import copy as _copy

    from ..core.package import DocxPackage
    from ..ops import bibliography as _bib

    collected = []
    seen_tags = set()
    for sp in src_paths:
        try:
            spkg = DocxPackage(sp)
        except Exception:
            continue
        store = _bib._find_store(spkg)
        if not store:
            continue
        for s in spkg.root(store).findall(_bib._bq("Source")):
            tag = s.findtext(_bib._bq("Tag")) or ""
            if tag and tag not in seen_tags:
                seen_tags.add(tag)
                collected.append(_copy.deepcopy(s))
    if not collected:
        return 0
    opkg = DocxPackage(out_path)
    part = _bib._ensure_store(opkg)
    root = opkg.root(part)
    existing = {
        s.findtext(_bib._bq("Tag")) for s in root.findall(_bib._bq("Source"))
    }
    added = 0
    for s in collected:
        if s.findtext(_bib._bq("Tag")) not in existing:
            root.append(s)
            added += 1
    if added:
        opkg.mark_dirty(part)
        opkg.save(do_backup=False)
    return added


@_bounded_op("com_multi_document", default=300.0)
def combine_documents(
    original_path: str,
    revised_path: str,
    output_path: str | None = None,
) -> dict:
    """Combine two documents' TRACKED CHANGES into one (Word's Combine — for
    merging two reviewers' edits of the same draft; both authors' revisions
    survive as separate attributions). Distinct from compare, which diffs
    content."""
    original_path = check_path(original_path, "combine original")
    revised_path = check_path(revised_path, "combine revised")
    if output_path:
        output_path = check_path(output_path, "combine output")
    orig = Path(original_path)
    rev = Path(revised_path)
    for p in (orig, rev):
        if not p.exists():
            raise DocumentNotFound(f"no file at {p}")
    out = (
        Path(output_path)
        if output_path
        else rev.with_name(f"{rev.stem}_COMBINED{rev.suffix}")
    )
    with _word() as app:
        doc_a = _open(app, orig, read_only=True)
        doc_b = _open(app, rev, read_only=True)
        try:
            result = app.MergeDocuments(
                OriginalDocument=doc_a,
                RevisedDocument=doc_b,
                Destination=2,  # wdCompareDestinationNew
            )
            result.SaveAs2(str(out.resolve()))
            result.Close(SaveChanges=_WD_DO_NOT_SAVE)
        finally:
            doc_a.Close(SaveChanges=_WD_DO_NOT_SAVE)
            doc_b.Close(SaveChanges=_WD_DO_NOT_SAVE)
    from ..core.package import DocxPackage
    from ..ops.read import revision_summary

    summary = revision_summary(DocxPackage(out))
    return {"combined": str(out), "revisions": summary}


def _find_open_document(path: str):
    """(pythoncom, app, doc) for a document open in ANY running interactive
    Word instance, else raise.

    Multi-instance aware (WS-L, 2026-08-28): GetActiveObject returns only
    whichever instance registered first — and that instance can be busy
    (modal dialog) or simply not the one holding the document. Open
    documents register their full path as a ROT file moniker, so the
    fallback binds the document directly, exactly like the live layer's
    _find_doc_via_rot."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    path = check_path(path, "find open document")
    target = str(Path(path).resolve()).lower()
    word_seen = False
    try:
        app = win32com.client.GetActiveObject("Word.Application")
        word_seen = True
        for doc in app.Documents:
            if doc.FullName.lower() == target:
                return pythoncom, app, doc
    except Exception:
        # not running, busy, or dying — the ROT scan below still works for
        # documents held by OTHER instances
        pass
    with contextlib.suppress(Exception):
        rot = pythoncom.GetRunningObjectTable()
        for moniker in rot.EnumRunning():
            ctx = pythoncom.CreateBindCtx(0)
            try:
                name = moniker.GetDisplayName(ctx, None)
            except Exception:
                continue
            if name.lower() != target:
                continue
            with contextlib.suppress(Exception):
                obj = rot.GetObject(moniker)
                doc = win32com.client.Dispatch(
                    obj.QueryInterface(pythoncom.IID_IDispatch)
                )
                return pythoncom, doc.Application, doc
            word_seen = True  # moniker exists but binding failed
    pythoncom.CoUninitialize()
    if not word_seen:
        raise WordMcpError("Word is not running")
    raise DocumentNotFound(
        f"{Path(path).name} is not open in any running Word instance"
    )


@contextlib.contextmanager
def _alerts_suppressed(app):
    """DisplayAlerts off around a call into the USER's Word, restored
    after — contention then raises a catchable com_error instead of
    freezing Word behind a modal dialog (stress report bug 3)."""
    prev = None
    with contextlib.suppress(Exception):
        prev = app.DisplayAlerts
        app.DisplayAlerts = _WD_ALERTS_NONE
    try:
        yield
    finally:
        if prev is not None:
            with contextlib.suppress(Exception):
                app.DisplayAlerts = prev


def _retry_word_call(fn, *, attempts: int = 5, first_delay: float = 0.5):
    """Bounded retry with exponential backoff for save/close contention
    (Word's write-to-temp-then-rename save can transiently collide).
    Total worst-case wait: 0.5+1+2+4 = 7.5s."""
    delay = first_delay
    last_exc = None
    for attempt in range(attempts):
        try:
            return fn(), attempt
        except Exception as exc:
            last_exc = exc
            if attempt < attempts - 1:
                time.sleep(delay)
                delay *= 2
    raise WordBusy(
        f"Word refused the operation {attempts} times (file contention or "
        f"a permission error; last error: {last_exc}). Alerts were "
        "suppressed, so no dialog is pending — wait a moment and retry, "
        "or save manually in Word."
    ) from last_exc


@_serial.serialized("com_save_document")
def save_open_document(path: str) -> dict:
    """Tell the USER's running Word to save a document it has open, so
    file-based tools can read the current state. Runs under the COM
    serialization lock with alerts suppressed and a bounded retry —
    concurrent saves were the report's 11/11 dialog-blocked failure."""
    pythoncom, app, doc = _find_open_document(path)
    try:
        with _alerts_suppressed(app):
            _, retries = _retry_word_call(doc.Save)
        out = {"saved": doc.FullName}
        if retries:
            out["retries"] = retries
        return out
    finally:
        pythoncom.CoUninitialize()


@_serial.serialized("com_save_document")
def close_open_document(path: str, *, save: bool = True) -> dict:
    """Tell the USER's running Word to close a document (saving by default),
    releasing the lock so file-based tools can edit it. Serialized, alerts
    suppressed, bounded retry (same contention path as save)."""
    pythoncom, app, doc = _find_open_document(path)
    try:
        with _alerts_suppressed(app):
            _, retries = _retry_word_call(
                lambda: doc.Close(
                    SaveChanges=_WD_SAVE if save else _WD_DO_NOT_SAVE
                ),
                attempts=3,
            )
        out = {"closed": str(Path(path).resolve()), "saved": save}
        if retries:
            out["retries"] = retries
        return out
    finally:
        pythoncom.CoUninitialize()


@_bounded_op("com_proofing_errors", default=60.0)
def proofing_errors(path: str, *, limit: int = 100) -> dict:
    """Word's own spelling and grammar error lists (with context). Slower on
    long documents; capped by limit per category.

    Name-collision guard (stress report bug 6): proofing needs a WRITE
    open in the invisible instance (the proofed flags are reset), and
    opening a second writable copy of an open document triggers Word's
    same-name dialog cascade — refused up front instead."""
    path = check_path(path, "proofing errors")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if _open_in_running_word(p):
        raise WordBusy(
            f"{p.name} is open in Word right now; proofing opens a second "
            "writable copy invisibly, which triggers Word's same-name "
            "dialog cascade. Save and close it first (com_save_document "
            "with close=true), or proof a copy (copy_document)."
        )
    with _word() as app:
        doc = _open(app, p, read_only=False)
        try:
            # Word marks documents as already-proofed; without resetting these
            # flags SpellingErrors is silently empty (verified 2026-08-28).
            doc.SpellingChecked = False
            doc.GrammarChecked = False
            spelling = []
            for err in doc.SpellingErrors:
                if len(spelling) >= limit:
                    break
                spelling.append(err.Text)
            grammar = []
            for err in doc.GrammaticalErrors:
                if len(grammar) >= limit:
                    break
                text = err.Text
                grammar.append(text[:120])
        finally:
            doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    return {
        "spelling_errors": spelling,
        "spelling_truncated": len(spelling) >= limit,
        "grammar_flagged_ranges": grammar,
        "grammar_truncated": len(grammar) >= limit,
        "note": (
            "Word's proofing engine; includes proper nouns and citations it "
            "does not recognize — review, do not auto-fix"
        ),
    }


@_bounded_op("com_readability_statistics", default=60.0)
def readability_statistics(path: str) -> dict:
    """Word's readability statistics (Flesch Reading Ease, grade level, word
    and sentence counts...). Refuses when the file is open in Word (the
    invisible read-only copy would compute stats for the stale saved state
    while risking instance contention); save and close first, or run on a
    copy."""
    path = check_path(path, "readability statistics")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if _open_in_running_word(p):
        raise WordBusy(
            f"{p.name} is open in Word right now; close it first "
            "(com_save_document with close=true) or run the statistics on "
            "a copy (copy_document) — a second invisible copy of an open "
            "document risks Word's same-name dialog cascade."
        )
    with _word() as app:
        doc = _open(app, p, read_only=True)
        try:
            stats = {}
            for st in doc.Content.ReadabilityStatistics:
                stats[st.Name] = st.Value
        finally:
            doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    return {"readability": stats}


@_bounded_op("com_save_document", default=300.0)
def save_with_password(
    path: str, output_path: str | None = None, *, password: str
) -> dict:
    """Save a copy encrypted with an open-password (real encryption, unlike
    document protection). Word will demand the password to open the copy."""
    path = check_path(path, "encrypt source")
    if output_path:
        output_path = check_path(output_path, "encrypt output")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if not password:
        raise WordMcpError("password must be non-empty")
    out = (
        Path(output_path)
        if output_path
        else p.with_name(f"{p.stem}_PROTECTED{p.suffix}")
    )
    with _word() as app:
        doc = _open(app, p, read_only=True)
        try:
            doc.SaveAs2(str(out.resolve()), Password=password)
        finally:
            doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word produced no output")
    return {
        "encrypted_copy": str(out),
        "note": "keep the password safe; there is no recovery",
    }


def zombie_check() -> dict:
    """Count WINWORD.EXE processes (diagnostic for leak detection)."""
    import subprocess

    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        timeout=30,
    )
    lines = [
        ln for ln in result.stdout.splitlines() if "WINWORD.EXE" in ln.upper()
    ]
    return {"winword_processes": len(lines)}
