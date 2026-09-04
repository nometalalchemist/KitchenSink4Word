"""Process-wide COM serialization (live COM stress report, 2026-09-03).

Word's COM interface is single-threaded (STA). The stress test proved that
concurrent tool calls reaching the Word proxy interleave at the character
level (garbled Range.Text writes), collide on the save temp-file swap
(modal dialogs), and starve invisible instances (30-minute hangs). The fix
is architectural and small: exactly ONE tool call reaches any Word COM
proxy at a time, enforced by this module's process-wide lock.

Coverage contract: every COM entry point acquires the lock —
- the live layer (live.live_session wraps every live tool and every
  dual-mode auto-route, live_repair, interactive_status),
- the bridge layer (every public invisible-instance and message-the-
  visible-instance function; enforced by test_com_serialization's audit
  of the _com_serialized marker),
- the convert layer (import_pdf).

The lock is an RLock: nested acquisitions on one thread are legal (a
bridge helper called under an already-held lock must not deadlock).

Scope honesty: the lock serializes THIS server process. Two separate
word-mcp processes automating the same Word are not serialized (each MCP
client normally spawns its own server, but the report's failure mode —
several agents of ONE session sharing one server — is fully covered).

Individual operations are fast; serialization latency is negligible
(report finding). Status reporting (lock_snapshot) lets com_word_status
tell callers honestly when a call would queue.
"""

from __future__ import annotations

import contextlib
import threading
import time

COM_LOCK = threading.RLock()

_state_lock = threading.Lock()
_current: dict | None = None      # {name, thread, started_wall, started_mono}
_last: dict | None = None         # {name, duration_ms, waited_ms, finished_wall}
_depth = 0                        # re-entrant depth on the owning thread
_serialized_total = 0             # ops that ran under the lock
_waited_total_ms = 0.0            # cumulative wait time across ops


@contextlib.contextmanager
def com_operation(name: str):
    """Hold the process-wide COM lock for the duration of one COM-touching
    operation. Records timing so com_word_status can report contention."""
    global _current, _last, _depth, _serialized_total, _waited_total_ms
    t0 = time.monotonic()
    COM_LOCK.acquire()
    waited_ms = (time.monotonic() - t0) * 1000.0
    with _state_lock:
        _depth += 1
        outermost = _depth == 1
        if outermost:
            _current = {
                "name": name,
                "thread": threading.get_ident(),
                "started_wall": time.time(),
                "started_mono": time.monotonic(),
            }
            _serialized_total += 1
            _waited_total_ms += waited_ms
    started = time.monotonic()
    try:
        yield
    finally:
        with _state_lock:
            _depth -= 1
            if outermost:
                _last = {
                    "name": name,
                    "duration_ms": round(
                        (time.monotonic() - started) * 1000.0, 1
                    ),
                    "waited_ms": round(waited_ms, 1),
                    "finished_wall": time.time(),
                }
                _current = None
        COM_LOCK.release()


def serialized(name: str):
    """Decorator form of com_operation for whole-function COM operations.
    Marks the function so the coverage audit test can verify every COM
    entry point takes the lock."""

    def deco(fn):
        import functools

        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            with com_operation(name):
                return fn(*args, **kwargs)

        wrapper._com_serialized = name
        return wrapper

    return deco


def lock_snapshot() -> dict:
    """Contention report for com_word_status: is a COM operation running
    right now, what is it, and how long has it held the lock. last_op
    carries the previous operation's duration and queue wait."""
    with _state_lock:
        out: dict = {
            "held": _current is not None,
            "ops_serialized": _serialized_total,
        }
        if _current is not None:
            out["current_op"] = {
                "name": _current["name"],
                "running_ms": round(
                    (time.monotonic() - _current["started_mono"]) * 1000.0, 1
                ),
            }
        if _last is not None:
            out["last_op"] = {
                "name": _last["name"],
                "duration_ms": _last["duration_ms"],
                "waited_ms": _last["waited_ms"],
            }
        return out


def acquire(timeout: float) -> bool:
    """Bounded acquisition for callers that must stay responsive
    (com_word_status). Pair with release() only when this returns True."""
    return COM_LOCK.acquire(timeout=timeout)


def release() -> None:
    COM_LOCK.release()
