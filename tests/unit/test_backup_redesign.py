"""v1.6 Workstream 0 items 2 + 7: slot-based backups (.ks4w-backups/) and
per-file write serialization. Covers: slot rotation (exactly 2 slots ever),
anchor idle-gap rotation, hardlink fallback, unicode/Korean and very long
doc names, the never-absent-target invariant (operation-order recording),
concurrent-mutation serialization, restore-rotates-prev, purge dry_run,
lockfile stale-break and live-holder refusal, and backup=False semantics."""

from __future__ import annotations

import os
import shutil
import threading
import time
from pathlib import Path

import pytest

import word_mcp.server as srv
from word_mcp.core import safesave
from word_mcp.core.package import DocxPackage
from word_mcp.core.safesave import (
    ANCHOR_SLOT,
    BACKUP_DIR_NAME,
    PREV_SLOT,
    SLOT_POLICY,
    MutationLockTimeout,
)
from word_mcp.ops import backups as bk


def _make_doc(tmp_path: Path, name: str = "t.docx") -> Path:
    doc = tmp_path / name
    srv.create_document(str(doc))
    return doc


def _texts(doc: Path) -> list[str]:
    return [p["text"] for p in srv.get_text(str(doc), live="off")]


def _slot_files(doc: Path) -> list[str]:
    d = safesave.slot_dir(doc)
    if not d.is_dir():
        return []
    return sorted(
        p.name for p in d.iterdir() if p.is_file() and p.suffix == ".docx"
    )


# ------------------------------------------------------------- slot rotation


def test_two_slots_ever_across_many_mutations(tmp_path):
    doc = _make_doc(tmp_path)
    for i in range(6):
        srv.insert_paragraphs(str(doc), [{"text": f"edit {i}"}], at_end=True)
    assert _slot_files(doc) == sorted(SLOT_POLICY)
    # No legacy .bak files are produced any more.
    assert not list(tmp_path.glob("*.bak-*")), "old .bak scheme resurfaced"
    # anchor = state before the FIRST mutation (empty new doc);
    # prev = state before the LAST mutation (has edits 0..4).
    d = safesave.slot_dir(doc)
    anchor_texts = [
        p["text"] for p in srv.get_text(str(d / ANCHOR_SLOT), live="off")
    ]
    prev_texts = [p["text"] for p in srv.get_text(str(d / PREV_SLOT), live="off")]
    assert "edit 0" not in " ".join(anchor_texts)
    assert "edit 4" in prev_texts and "edit 5" not in prev_texts


def test_backup_root_is_dot_prefixed_and_next_to_doc(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "x"}], at_end=True)
    root = tmp_path / BACKUP_DIR_NAME
    assert root.is_dir()
    assert (root / doc.name).is_dir()


def test_anchor_rotates_after_idle_gap(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "session one"}], at_end=True)
    srv.insert_paragraphs(str(doc), [{"text": "still session one"}], at_end=True)
    d = safesave.slot_dir(doc)
    # Simulate the idle gap by aging the prev slot's mtime (idle is measured
    # from slot mtimes only - no state database).
    old = time.time() - (safesave.ANCHOR_IDLE_SECONDS + 60)
    os.utime(d / PREV_SLOT, (old, old))
    srv.insert_paragraphs(str(doc), [{"text": "session two"}], at_end=True)
    anchor_texts = [
        p["text"] for p in srv.get_text(str(d / ANCHOR_SLOT), live="off")
    ]
    # New-session anchor = content at the start of session two, which
    # includes everything session one wrote.
    assert "still session one" in anchor_texts
    assert "session two" not in anchor_texts


def test_anchor_stable_within_session(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "first"}], at_end=True)
    d = safesave.slot_dir(doc)
    before = (d / ANCHOR_SLOT).read_bytes()
    srv.insert_paragraphs(str(doc), [{"text": "second"}], at_end=True)
    srv.insert_paragraphs(str(doc), [{"text": "third"}], at_end=True)
    assert (d / ANCHOR_SLOT).read_bytes() == before


def test_hardlink_failure_falls_back_to_copy(tmp_path, monkeypatch):
    doc = _make_doc(tmp_path)

    def no_link(*args, **kwargs):
        raise OSError(1, "cross-device / non-NTFS / cloud placeholder")

    monkeypatch.setattr(os, "link", no_link)
    srv.insert_paragraphs(str(doc), [{"text": "copied not linked"}], at_end=True)
    assert _slot_files(doc) == sorted(SLOT_POLICY)
    d = safesave.slot_dir(doc)
    # Slots are real independent copies and load as valid documents.
    DocxPackage(d / PREV_SLOT)
    DocxPackage(d / ANCHOR_SLOT)


def test_korean_unicode_doc_name(tmp_path):
    doc = _make_doc(tmp_path, "제4장 초안 (최종).docx")
    srv.insert_paragraphs(str(doc), [{"text": "한글 내용"}], at_end=True)
    srv.insert_paragraphs(str(doc), [{"text": "더 많은 내용"}], at_end=True)
    d = safesave.slot_dir(doc)
    assert d.name == doc.name  # reverse mapping stays trivial
    assert _slot_files(doc) == sorted(SLOT_POLICY)
    assert bk.list_backups(file_path=str(doc))["slots"]


def test_very_long_doc_name_gets_hash_suffix_and_breadcrumb(tmp_path):
    name = "아주 긴 한글 파일 이름 " * 8 + "final draft version twelve.docx"
    assert len(name) > 100
    doc = _make_doc(tmp_path, name)
    srv.insert_paragraphs(str(doc), [{"text": "long name content"}], at_end=True)
    d = safesave.slot_dir(doc)
    assert d.is_dir() and len(d.name) < len(name)
    # Breadcrumb maps the truncated folder back to the source document.
    assert safesave.source_doc_for(d) == doc
    # Therefore it is NOT reported as an orphan while the doc exists...
    listing = bk.list_backups(directory=str(tmp_path))
    assert not listing["orphaned_folders"]
    # ...and IS once the doc is gone.
    doc.unlink()
    listing = bk.list_backups(directory=str(tmp_path))
    assert [o["folder"] for o in listing["orphaned_folders"]] == [str(d)]


# -------------------------------------------------- never-absent invariant


def test_target_never_absent_during_save(tmp_path, monkeypatch):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "seed"}], at_end=True)  # anchor set

    target = os.path.normcase(str(doc.resolve()))
    events: list[tuple[str, str, str]] = []
    real_link, real_replace = os.link, os.replace
    real_unlink, real_remove = os.unlink, os.remove

    def norm(p):
        return os.path.normcase(os.path.abspath(os.fspath(p)))

    monkeypatch.setattr(
        os, "link",
        lambda s, dst, **kw: (events.append(("link", norm(s), norm(dst))),
                              real_link(s, dst, **kw))[1],
    )
    monkeypatch.setattr(
        os, "replace",
        lambda s, dst: (events.append(("replace", norm(s), norm(dst))),
                        real_replace(s, dst))[1],
    )
    monkeypatch.setattr(
        os, "unlink",
        lambda p, *a, **kw: (events.append(("unlink", norm(p), "")),
                             real_unlink(p, *a, **kw))[1],
    )
    monkeypatch.setattr(
        os, "remove",
        lambda p, *a, **kw: (events.append(("unlink", norm(p), "")),
                             real_remove(p, *a, **kw))[1],
    )

    srv.insert_paragraphs(str(doc), [{"text": "mutation"}], at_end=True)

    # The document itself is never unlinked/removed at any point.
    assert all(not (op == "unlink" and src == target) for op, src, _ in events)
    # Required sequence: capture current target (link), rotate it onto prev,
    # THEN (and only then) replace the target with the validated temp.
    prev_slot = norm(safesave.slot_dir(doc) / PREV_SLOT)
    i_link = next(
        i for i, (op, src, _) in enumerate(events)
        if op == "link" and src == target
    )
    i_prev = next(
        i for i, (op, _, dst) in enumerate(events)
        if op == "replace" and dst == prev_slot
    )
    i_target = next(
        i for i, (op, _, dst) in enumerate(events)
        if op == "replace" and dst == target
    )
    assert i_link < i_prev < i_target
    # The replace onto the target is the ONLY operation that touches its path.
    touches = [
        e for e in events if e[1] == target or e[2] == target
    ]
    assert [e[0] for e in touches].count("replace") == 1


# ------------------------------------------------------------- concurrency


def test_concurrent_mutations_serialize(tmp_path):
    doc = _make_doc(tmp_path)
    n_seed = 8
    srv.insert_paragraphs(
        str(doc),
        [{"text": f"unique marker {i} alpha"} for i in range(n_seed)],
        at_end=True,
    )

    results: dict[int, dict] = {}
    inserted: list[str] = []
    errors: list[BaseException] = []

    def fmt(i):
        try:
            r = srv.format_text(
                str(doc), {"bold": True}, find=f"unique marker {i} alpha",
                live="off",
            )
            results[i] = r["formatted"]
        except BaseException as exc:  # noqa: BLE001
            errors.append(exc)

    def ins(i):
        try:
            text = f"threaded insert {i}"
            srv.insert_paragraphs(
                str(doc), [{"text": text}], at_end=True, live="off"
            )
            inserted.append(text)
        except BaseException as exc:  # noqa: BLE001
            errors.append(exc)

    threads = [threading.Thread(target=fmt, args=(i,)) for i in range(n_seed)]
    threads += [threading.Thread(target=ins, args=(i,)) for i in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert not errors, errors
    # Final document is valid and complete: NO lost updates. Under the old
    # read-modify-save race, parallel calls clobbered each other (last writer
    # wins), silently dropping inserts and formatting.
    final = _texts(doc)
    assert len([t for t in final if t.startswith("threaded insert")]) == 4
    pkg = DocxPackage(doc)
    from word_mcp.core.package import qn

    bolded = set()
    for p in pkg.body().iter(qn("w:p")):
        text = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if text.startswith("unique marker"):
            runs = list(p.iter(qn("w:r")))
            if runs and all(
                r.find(qn("w:rPr")) is not None
                and r.find(qn("w:rPr")).find(qn("w:b")) is not None
                for r in runs
            ):
                bolded.add(text)
    assert bolded == {f"unique marker {i} alpha" for i in range(n_seed)}, (
        "a concurrent format_text was lost (read-modify-save race)"
    )
    # Response metadata is consistent: each call reports the exact character
    # span of its target. (The paragraph-index field of format_text is NOT
    # asserted here: it is wrong even sequentially due to a pre-existing
    # id()-keyed lxml proxy bug in ops/text.py, tracked separately.)
    assert len(results) == n_seed
    for i, loc in results.items():
        assert (loc["start"], loc["end"]) == (
            0, len(f"unique marker {i} alpha")
        )


# ------------------------------------------------------------------ locking


def test_stale_lockfile_is_broken(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "seed"}], at_end=True)
    d = safesave.slot_dir(doc)
    lock = d / safesave.LOCK_FILE_NAME
    lock.write_text(
        '{"pid": 999999999, "time": 1.0}', encoding="utf-8"
    )  # dead pid, ancient timestamp
    srv.insert_paragraphs(str(doc), [{"text": "after stale break"}], at_end=True)
    assert "after stale break" in _texts(doc)
    assert not lock.exists()


def test_live_foreign_lock_refused_with_holder_named(tmp_path, monkeypatch):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "seed"}], at_end=True)
    d = safesave.slot_dir(doc)
    lock = d / safesave.LOCK_FILE_NAME
    foreign_pid = os.getpid() + 12345
    lock.write_text(
        f'{{"pid": {foreign_pid}, "time": {time.time()}}}', encoding="utf-8"
    )
    monkeypatch.setattr(safesave, "_pid_alive", lambda pid: True)
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)
    with pytest.raises(MutationLockTimeout, match=str(foreign_pid)):
        srv.insert_paragraphs(str(doc), [{"text": "blocked"}], at_end=True)
    lock.unlink()  # release for other tests' sake


def test_same_process_lockfile_is_reentrant(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "seed"}], at_end=True)
    with safesave.write_lock(doc):
        # A nested mutation in the same process must not deadlock or refuse.
        srv.insert_paragraphs(str(doc), [{"text": "nested"}], at_end=True)
    assert "nested" in _texts(doc)


# ------------------------------------------------------------ backup=False


def test_backup_false_skips_rotation_never_the_atomic_save(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(
        str(doc), [{"text": "no backup"}], at_end=True, backup=False
    )
    assert "no backup" in _texts(doc)
    assert _slot_files(doc) == []  # no slots rotated
    # Atomic save still applied: no temp litter, document valid.
    assert not list(tmp_path.glob("*.word-mcp-tmp"))
    DocxPackage(doc)


# ------------------------------------------------------------ manage_backups


def test_list_reports_slots_legacy_and_orphans(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "x"}], at_end=True)
    legacy = tmp_path / f"{doc.stem}.bak-20260828_120000{doc.suffix}"
    shutil.copy2(doc, legacy)
    ghost = _make_doc(tmp_path, "ghost.docx")
    srv.insert_paragraphs(str(ghost), [{"text": "y"}], at_end=True)
    ghost.unlink()

    listing = bk.manage_backups("list", file_path=str(doc))
    assert {s["slot"] for s in listing["slots"]} == {"prev", "anchor"}
    for s in listing["slots"]:
        assert s["size_bytes"] > 0 and s["modified"]
    assert [Path(e["path"]).name for e in listing["legacy_backups"]] == [
        legacy.name
    ]
    assert [Path(o["folder"]).name for o in listing["orphaned_folders"]] == [
        "ghost.docx"
    ]


def test_restore_rotates_prev_first_and_is_undoable(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "version A"}], at_end=True)
    srv.insert_paragraphs(str(doc), [{"text": "version B"}], at_end=True)

    r = bk.manage_backups("restore", file_path=str(doc), source="prev")
    assert r["prev_rotated"] is True
    assert _texts(doc) == ["version A"]
    # Undo the restore by restoring prev again (prev = pre-restore content).
    bk.manage_backups("restore", file_path=str(doc), source="prev")
    assert _texts(doc) == ["version A", "version B"]


def test_restore_from_anchor(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "post-anchor edit"}], at_end=True)
    bk.manage_backups("restore", file_path=str(doc), source="anchor")
    assert "post-anchor edit" not in _texts(doc)
    DocxPackage(doc)


def test_restore_from_legacy_file(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "current"}], at_end=True)
    legacy = tmp_path / f"{doc.stem}.bak-20260828_120000{doc.suffix}"
    shutil.copy2(doc, legacy)
    srv.insert_paragraphs(str(doc), [{"text": "newer"}], at_end=True)
    bk.manage_backups("restore", file_path=str(doc), source=str(legacy))
    assert _texts(doc) == ["current"]


def test_restore_missing_source_is_clear_refusal(tmp_path):
    doc = _make_doc(tmp_path)
    from word_mcp.core.errors import DocumentNotFound

    with pytest.raises(DocumentNotFound, match="no backup to restore"):
        bk.manage_backups("restore", file_path=str(doc), source="prev")


def test_purge_dry_run_is_the_default_and_deletes_nothing(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "x"}], at_end=True)
    legacy = tmp_path / f"{doc.stem}.bak-20260828_120000{doc.suffix}"
    shutil.copy2(doc, legacy)

    r = bk.manage_backups("purge", file_path=str(doc), scope="legacy")
    assert r["dry_run"] is True
    assert [Path(e["path"]).name for e in r["would_delete"]] == [legacy.name]
    assert legacy.exists()

    r2 = bk.manage_backups(
        "purge", file_path=str(doc), scope="legacy", dry_run=False
    )
    assert [Path(e["path"]).name for e in r2["deleted"]] == [legacy.name]
    assert not legacy.exists()


def test_purge_slots_and_orphans(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "x"}], at_end=True)
    ghost = _make_doc(tmp_path, "ghost.docx")
    srv.insert_paragraphs(str(ghost), [{"text": "y"}], at_end=True)
    ghost.unlink()

    r = bk.manage_backups(
        "purge", file_path=str(doc), scope="slots", dry_run=False
    )
    assert r["count"] == 2
    assert _slot_files(doc) == []

    r2 = bk.manage_backups(
        "purge", directory=str(tmp_path), scope="orphans", dry_run=False
    )
    assert r2["count"] == 1
    assert not (tmp_path / BACKUP_DIR_NAME / "ghost.docx").exists()
    # The purged document itself is untouched.
    assert "x" in _texts(doc)


def test_anchor_recreated_after_slot_purge(tmp_path):
    doc = _make_doc(tmp_path)
    srv.insert_paragraphs(str(doc), [{"text": "before purge"}], at_end=True)
    bk.manage_backups("purge", file_path=str(doc), scope="slots", dry_run=False)
    srv.insert_paragraphs(str(doc), [{"text": "after purge"}], at_end=True)
    # Session-start semantics: no anchor -> created on first mutation.
    assert _slot_files(doc) == sorted(SLOT_POLICY)


def test_manage_backups_argument_validation(tmp_path):
    from word_mcp.core.errors import WordMcpError

    with pytest.raises(WordMcpError, match="list|restore|purge"):
        bk.manage_backups("compact")
    with pytest.raises(WordMcpError, match="file_path"):
        bk.manage_backups("list")
    with pytest.raises(WordMcpError, match="source"):
        bk.manage_backups("restore", file_path=str(tmp_path / "a.docx"))
    with pytest.raises(WordMcpError, match="scope"):
        bk.manage_backups("purge", file_path=str(tmp_path / "a.docx"))
