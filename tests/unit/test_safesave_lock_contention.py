"""Regression suite for the FILE-route write lock under contention.

Ported from the Excel sibling (xlsx-mcp), whose concurrency gate disproved
the pre-token two-syscall design that this repo also shipped. The original
``_acquire_lockfile`` created the lockfile with ``O_CREAT|O_EXCL`` and wrote
the payload as a SECOND syscall, which leaves a window in which the file
exists but is EMPTY. A concurrent acquirer that read it inside that window
parsed no pid, concluded the holder was dead, DELETED a live lock and took
its own. Two processes then believed they held the same lock, and each
one's unconditional release unlinked whichever file happened to be there,
including the other's live one, so the lock outlived every writer.

xlsx reproduced that 2/2 at four or more concurrent writers. The tests here
pin the four properties of the corrected design:

- atomic publish (temp file + ``os.link``), so no half-written lockfile is
  ever observable;
- a per-process-instance token, so re-entrancy is token equality and a
  recycled PID under a foreign token is stale by definition;
- release only if the lockfile is still ours;
- staleness that checks the holder's process CREATION TIME, so a recycled
  PID is not mistaken for the live holder.

``test_four_concurrent_writers_never_overlap`` is the end-to-end repro and
is the one that fails on the old logic. The rest are the unit-level pins.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import textwrap
import time
from pathlib import Path

import pytest

from word_mcp.core import safesave
from word_mcp.core.safesave import MutationLockTimeout

SRC = Path(__file__).resolve().parents[2] / "src"


# ------------------------------------------------- the half-written window


def test_an_empty_lockfile_is_not_read_as_a_dead_holder(tmp_path, monkeypatch):
    """THE repro, reduced to one deterministic assertion.

    An empty ``write.lock`` is exactly what the old two-syscall create left
    visible between the create and the write. The old code parsed no pid,
    called it stale, unlinked it and took the lock, which is how a LIVE
    holder lost its lockfile. The corrected code gives an unreadable lock a
    grace window and refuses rather than stealing it.
    """
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.touch()  # the half-written state, frozen
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)

    with pytest.raises(MutationLockTimeout):
        safesave._acquire_lockfile(lock, "contended.docx")

    assert lock.exists(), (
        "an empty (half-written) lockfile was deleted and stolen; a live "
        "holder just lost its lock"
    )


def test_an_unreadable_lockfile_is_reclaimed_once_the_grace_expires(tmp_path):
    """The grace must not turn a genuinely abandoned lockfile into a
    permanent one."""
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.write_text("{not json", encoding="utf-8")
    old = time.time() - (safesave._UNREADABLE_GRACE_SECONDS + 60)
    os.utime(lock, (old, old))

    assert safesave._acquire_lockfile(lock, "abandoned.docx") is True
    assert json.loads(lock.read_text())["token"] == safesave._OWNER_TOKEN
    safesave._release_lockfile(lock)


# ------------------------------------------------------- the four properties


def test_the_lockfile_is_published_complete(tmp_path):
    """Atomic publish: the payload is written before the file is visible, so
    every reader sees a complete record."""
    lock = tmp_path / safesave.LOCK_FILE_NAME
    assert safesave._publish_lockfile(lock) is True
    info = json.loads(lock.read_text(encoding="utf-8"))
    assert info["pid"] == os.getpid()
    assert info["token"] == safesave._OWNER_TOKEN
    assert info["time"] > 0
    assert "pid_created" in info
    # A second publish reports the lock is taken rather than clobbering it.
    assert safesave._publish_lockfile(lock) is False
    assert not list(tmp_path.glob(".lock-*.tmp")), "temp file leaked"
    safesave._release_lockfile(lock)


def test_a_second_acquisition_by_this_process_is_re_entrant(tmp_path):
    lock = tmp_path / safesave.LOCK_FILE_NAME
    assert safesave._acquire_lockfile(lock, "t.docx") is True
    assert safesave._acquire_lockfile(lock, "t.docx") is False
    safesave._release_lockfile(lock)


def test_a_recycled_pid_under_a_foreign_token_is_stale(tmp_path):
    """Our own PID carrying someone else's token means the number was
    recycled. The old PID-only re-entrancy amnesty read that as "mine,
    already held" and never cleaned it up, forever."""
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.write_text(
        json.dumps({"pid": os.getpid(), "token": "someone-else",
                    "time": time.time()}),
        encoding="utf-8",
    )
    assert safesave._acquire_lockfile(lock, "t.docx") is True
    assert json.loads(lock.read_text())["token"] == safesave._OWNER_TOKEN
    safesave._release_lockfile(lock)


def test_a_live_pid_with_a_different_creation_time_is_stale(
    tmp_path, monkeypatch
):
    """Windows recycles PIDs. "Is that number alive" is the wrong question:
    the answer can be an unrelated process. Without the creation-time check
    the acquirer waits the full window for a lock nobody holds."""
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.write_text(
        json.dumps({"pid": 4321, "token": "other", "pid_created": 111.0,
                    "time": time.time()}),
        encoding="utf-8",
    )
    monkeypatch.setattr(safesave, "_pid_alive", lambda pid: True)
    monkeypatch.setattr(safesave, "_process_create_time", lambda pid: 999.0)
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 5.0)

    started = time.monotonic()
    assert safesave._acquire_lockfile(lock, "t.docx") is True
    assert time.monotonic() - started < 2.0, "waited on a recycled PID"
    safesave._release_lockfile(lock)


def test_release_never_removes_someone_elses_lock(tmp_path):
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.write_text(
        json.dumps({"pid": 999999, "token": "not-ours", "time": time.time()}),
        encoding="utf-8",
    )
    safesave._release_lockfile(lock)
    assert lock.exists(), "released a lock this process never held"

    lock.unlink()
    safesave._publish_lockfile(lock)
    safesave._release_lockfile(lock)
    assert not lock.exists()


def test_a_matching_creation_time_keeps_a_live_holder_protected(
    tmp_path, monkeypatch
):
    """The creation-time check must only break recycled PIDs, never a real
    live holder."""
    lock = tmp_path / safesave.LOCK_FILE_NAME
    lock.write_text(
        json.dumps({"pid": 4321, "token": "other", "pid_created": 111.0,
                    "time": time.time()}),
        encoding="utf-8",
    )
    monkeypatch.setattr(safesave, "_pid_alive", lambda pid: True)
    monkeypatch.setattr(safesave, "_process_create_time", lambda pid: 111.0)
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)

    with pytest.raises(MutationLockTimeout, match="4321"):
        safesave._acquire_lockfile(lock, "t.docx")
    assert lock.exists()
    lock.unlink()


# ------------------------------------------------ end-to-end contention repro


WRITER = textwrap.dedent("""\
    import sys, os, time, json
    sys.path.insert(0, sys.argv[1])
    from word_mcp.core.safesave import write_lock

    target, record, hold = sys.argv[2], sys.argv[3], float(sys.argv[4])
    with write_lock(target):
        start = time.time()
        time.sleep(hold)
        end = time.time()
    with open(record, "w", encoding="utf-8") as fh:
        json.dump({"pid": os.getpid(), "start": start, "end": end}, fh)
""")


def _make_doc(tmp_path: Path) -> Path:
    from docx import Document

    p = tmp_path / "contended.docx"
    d = Document()
    d.add_paragraph("seed")
    d.save(str(p))
    return p


def test_four_concurrent_writers_never_overlap(tmp_path):
    """The xlsx contention repro: four or more real processes racing for the
    same document's write lock.

    On the old two-syscall logic this reproduced 2/2 - a writer read a
    half-written lockfile, declared the live holder stale, deleted its lock
    and entered the critical section alongside it, producing overlapping
    hold intervals and a leaked lockfile. Both properties are asserted.
    """
    doc = _make_doc(tmp_path)
    writers = 5
    hold = 0.4

    procs = []
    for i in range(writers):
        record = tmp_path / f"writer{i}.json"
        procs.append((
            record,
            subprocess.Popen(
                [sys.executable, "-X", "utf8", "-c", WRITER,
                 str(SRC), str(doc), str(record), str(hold)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            ),
        ))

    intervals = []
    for record, p in procs:
        _, err = p.communicate(timeout=120)
        assert p.returncode == 0, (
            f"writer failed: {err.decode('utf-8', 'replace')}"
        )
        assert record.exists(), "writer never recorded its critical section"
        intervals.append(json.loads(record.read_text(encoding="utf-8")))

    assert len(intervals) == writers
    intervals.sort(key=lambda r: r["start"])
    for earlier, later in zip(intervals, intervals[1:]):
        assert later["start"] >= earlier["end"], (
            "two processes held the write lock at the same time: "
            f"PID {earlier['pid']} held {earlier['start']}-{earlier['end']}, "
            f"PID {later['pid']} started {later['start']}. The lock admitted "
            "a second writer, which is the half-written-lockfile race."
        )

    lock = safesave.slot_dir(doc) / safesave.LOCK_FILE_NAME
    assert not lock.exists(), (
        "the lockfile outlived every writer - a release removed a file this "
        "process did not own, or two holders unlinked each other's"
    )
    assert not list(safesave.slot_dir(doc).glob(".lock-*.tmp")), (
        "publish temp files leaked"
    )


VICTIM = textwrap.dedent("""\
    import sys, os, time, json
    sys.path.insert(0, sys.argv[1])
    from word_mcp.core.safesave import write_lock

    target, record, flag, hold = (
        sys.argv[2], sys.argv[3], sys.argv[4], float(sys.argv[5]))
    with write_lock(target):
        start = time.time()
        open(flag, "w").close()          # tell the parent the lock is held
        time.sleep(hold)
        end = time.time()
    with open(record, "w", encoding="utf-8") as fh:
        json.dump({"pid": os.getpid(), "start": start, "end": end}, fh)
""")

CONTENDER = textwrap.dedent("""\
    import sys, os, time, json
    sys.path.insert(0, sys.argv[1])
    from word_mcp.core import safesave
    from word_mcp.core.safesave import write_lock, MutationLockTimeout

    target, record = sys.argv[2], sys.argv[3]
    safesave.LOCK_WAIT_SECONDS = float(sys.argv[4])
    out = {"pid": os.getpid(), "entered": False, "refused": False}
    try:
        with write_lock(target):
            out["entered"] = True
            out["start"] = time.time()
            time.sleep(0.1)
            out["end"] = time.time()
    except MutationLockTimeout:
        out["refused"] = True
    with open(record, "w", encoding="utf-8") as fh:
        json.dump(out, fh)
""")


def test_four_concurrent_writers_cannot_steal_a_half_written_lock(tmp_path):
    """The xlsx repro made deterministic.

    Hitting the old create/write window by luck takes microseconds of timing,
    so instead the fault condition is staged directly: one process takes a
    real write lock, its lockfile is then truncated to zero bytes (exactly
    the state the old two-syscall create made observable), and four more
    processes race for the same document.

    On the old logic all four parsed no pid from the empty file, declared the
    live holder stale, deleted its lockfile and entered the critical section
    beside it. On the corrected logic an unreadable lockfile is given a grace
    window instead of being stolen, so the holder keeps its lock and the
    contenders refuse.
    """
    doc = _make_doc(tmp_path)
    lock = safesave.slot_dir(doc, create=True) / safesave.LOCK_FILE_NAME
    flag = tmp_path / "held.flag"
    victim_record = tmp_path / "victim.json"
    hold = 3.0
    contender_wait = 1.5

    victim = subprocess.Popen(
        [sys.executable, "-X", "utf8", "-c", VICTIM,
         str(SRC), str(doc), str(victim_record), str(flag), str(hold)],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    try:
        deadline = time.monotonic() + 60
        while not flag.exists():
            assert victim.poll() is None, "victim died before taking the lock"
            assert time.monotonic() < deadline, "victim never took the lock"
            time.sleep(0.01)

        assert lock.exists()
        with open(lock, "w"):  # truncate: the half-written state, frozen
            pass

        contenders = []
        for i in range(4):
            record = tmp_path / f"contender{i}.json"
            contenders.append((
                record,
                subprocess.Popen(
                    [sys.executable, "-X", "utf8", "-c", CONTENDER,
                     str(SRC), str(doc), str(record), str(contender_wait)],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                ),
            ))

        results = []
        for record, p in contenders:
            _, err = p.communicate(timeout=120)
            assert p.returncode == 0, (
                f"contender crashed: {err.decode('utf-8', 'replace')}"
            )
            results.append(json.loads(record.read_text(encoding="utf-8")))
    finally:
        _, verr = victim.communicate(timeout=120)
        assert victim.returncode == 0, (
            f"victim crashed: {verr.decode('utf-8', 'replace')}"
        )

    held = json.loads(victim_record.read_text(encoding="utf-8"))
    stole = [
        r for r in results
        if r["entered"] and r["start"] < held["end"] and r["end"] > held["start"]
    ]
    assert not stole, (
        f"{len(stole)} of 4 contenders deleted a live holder's half-written "
        f"lockfile and entered the critical section beside PID {held['pid']}: "
        f"{[r['pid'] for r in stole]}"
    )
    assert all(r["refused"] for r in results), (
        "an unreadable lockfile held by a live writer must be waited on and "
        "refused, never stolen"
    )
    assert not lock.exists(), "the lockfile outlived its holder"
