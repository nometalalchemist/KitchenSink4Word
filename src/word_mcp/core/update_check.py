"""Is a newer KitchenSink4Word published? A quiet, opt-out startup check.

The family spec (identical in KS4W, KS4P, KS4XL, and KS4Web, modulo the
package and variable names):

- The check runs in a BACKGROUND daemon thread started after the server is
  built, so it can never delay serving. Nothing waits on it.
- It asks PyPI's public JSON endpoint for the package it was installed
  from, at most once every 14 days. The answer is cached in a tiny JSON
  file under the server's own state directory, never beside a user's
  documents.
- Failure is silence. Timeouts (3 seconds), network errors, and malformed
  payloads are swallowed; nothing reaches stderr and nothing reaches the
  caller. A failed attempt IS recorded, with a 1-day retry horizon, so an
  offline machine is neither hammered nor stuck forever.
- Comparison uses packaging.version and is prerelease-aware: a prerelease
  is never offered to somebody running a stable build.
- The result surfaces as ONE line in get_workflows() and nowhere else. No
  startup banner, no per-call nagging.
- KS4W_NO_UPDATE_CHECK=1 (or "true") disables the whole thing: no network
  call, no cache read, no cache write.
- The server never downloads, installs, or executes anything. It reports a
  version number and the command a human can run.
"""

from __future__ import annotations

import json
import os
import tempfile
import threading
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

from packaging.version import InvalidVersion, Version

#: PyPI distribution name (what `pip install -U` takes).
PACKAGE = "kitchensink4word"

#: Product name as a human reads it, used in the one surfaced line.
PRODUCT = "KitchenSink4Word"

#: The public JSON endpoint. Read-only, unauthenticated, no payload.
PYPI_URL = f"https://pypi.org/pypi/{PACKAGE}/json"

#: The opt-out. "1" or "true" (case-insensitive) turns everything off.
OPT_OUT_ENV = "KS4W_NO_UPDATE_CHECK"

#: Test/ops seam: point the cache somewhere else. Undocumented on purpose.
CACHE_DIR_ENV = "KS4W_UPDATE_CACHE_DIR"

#: How long a successful answer is trusted before asking again.
CHECK_INTERVAL = timedelta(days=14)

#: How long a FAILED attempt is honored before retrying (shorter, so an
#: offline machine catches up the day it comes back online).
RETRY_INTERVAL = timedelta(days=1)

#: Hard cap on the request. Three seconds or nothing.
TIMEOUT_SECONDS = 3.0

_CACHE_NAME = "update-check.json"
_STATE_DIR_NAME = "word-mcp"


# ----------------------------------------------------------------- opt-out


def disabled() -> bool:
    """True when the operator has turned the update check off.

    Checked BEFORE any file or network I/O in every entry point here, so
    an opt-out machine performs neither.
    """
    return (os.environ.get(OPT_OUT_ENV) or "").strip().lower() in {
        "1", "true", "yes", "on",
    }


# ------------------------------------------------------------- the version


def current_version() -> str | None:
    """The running build's version, or None when it cannot be read.

    Deliberately NOT importlib.metadata: an editable install's recorded
    metadata goes stale the moment the version constant is bumped, and a
    wrong "you are running X" line is worse than no line at all.
    """
    try:
        from .. import __version__

        return str(__version__)
    except Exception:
        return None


# --------------------------------------------------------------- the cache


def cache_path() -> Path:
    """The cache file: a couple of hundred bytes in the server's own state
    directory, alongside the live-lock directory the COM tier already uses.
    Never beside a user's documents."""
    override = os.environ.get(CACHE_DIR_ENV)
    if override:
        return Path(override) / _CACHE_NAME
    base = (
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("XDG_STATE_HOME")
        or tempfile.gettempdir()
    )
    return Path(base) / _STATE_DIR_NAME / _CACHE_NAME


def read_cache(path: Path | None = None) -> dict | None:
    """The cached answer, or None when there is not a readable one."""
    try:
        raw = (path or cache_path()).read_text(encoding="utf-8")
        data = json.loads(raw)
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def write_cache(data: dict, path: Path | None = None) -> None:
    """Replace the cache file. Silent on any failure (a read-only state
    directory must not break the server)."""
    target = path or cache_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix(".tmp")
        tmp.write_text(json.dumps(data), encoding="utf-8")
        os.replace(tmp, target)
    except Exception:
        pass


def _parse_stamp(value) -> datetime | None:
    try:
        stamp = datetime.fromisoformat(str(value))
    except Exception:
        return None
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=timezone.utc)
    return stamp


def is_due(cache: dict | None, now: datetime | None = None) -> bool:
    """Has the cached answer aged out?

    Two horizons: 14 days after a successful check, 1 day after a failed
    one. No cache, or an unreadable timestamp, means due.
    """
    if not cache:
        return True
    stamp = _parse_stamp(cache.get("last_check"))
    if stamp is None:
        return True
    horizon = CHECK_INTERVAL if cache.get("ok") else RETRY_INTERVAL
    return (now or datetime.now(timezone.utc)) - stamp >= horizon


# ------------------------------------------------------------ the question


def _latest_stable(payload: dict) -> str | None:
    """The newest NON-prerelease release in a PyPI JSON payload.

    Prefers the full releases map (so a prerelease published after the
    last stable cannot win), and falls back to info.version when the map
    is absent. Returns None for anything malformed.
    """
    if not isinstance(payload, dict):
        return None
    best: Version | None = None
    releases = payload.get("releases")
    if isinstance(releases, dict):
        for name, files in releases.items():
            try:
                candidate = Version(str(name))
            except InvalidVersion:
                continue
            if candidate.is_prerelease or candidate.is_devrelease:
                continue
            if isinstance(files, list):
                if not files:
                    continue  # no artifacts: not installable
                if all(
                    isinstance(f, dict) and f.get("yanked") for f in files
                ):
                    continue  # every artifact yanked
            if best is None or candidate > best:
                best = candidate
    if best is not None:
        return str(best)
    info = payload.get("info")
    if isinstance(info, dict):
        try:
            candidate = Version(str(info.get("version")))
        except InvalidVersion:
            return None
        if not (candidate.is_prerelease or candidate.is_devrelease):
            return str(candidate)
    return None


def _fetch(url: str = PYPI_URL) -> dict | None:
    """One GET, 3-second cap, no payload beyond a standard request. Returns
    the decoded JSON or None."""
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/json",
                 "User-Agent": f"{PACKAGE}/{current_version() or 'unknown'}"},
    )
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as resp:
        body = resp.read(1_000_000)
    data = json.loads(body.decode("utf-8"))
    return data if isinstance(data, dict) else None


def run_check(force: bool = False, path: Path | None = None) -> str | None:
    """Perform the check if it is due, refresh the cache, and return the
    latest stable version when one was learned.

    Never raises. Never prints. This is the background thread's body, so
    the outer guard matters: an escaping exception would reach the thread
    excepthook and print a traceback into a stdio transport.
    """
    if disabled():
        return None
    try:
        return _run_check(force, path)
    except Exception:
        return None


def _run_check(force: bool, path: Path | None) -> str | None:
    target = path or cache_path()
    cache = read_cache(target)
    if not force and not is_due(cache):
        return (cache or {}).get("latest_version")
    known = (cache or {}).get("latest_version")
    stamp = datetime.now(timezone.utc).isoformat()
    try:
        latest = _latest_stable(_fetch())
    except Exception:
        latest = None
    if latest is None:
        # Failed or unusable answer: record the ATTEMPT (1-day horizon) and
        # keep whatever we already knew. The user sees nothing.
        write_cache({"last_check": stamp, "latest_version": known,
                     "ok": False}, target)
        return None
    write_cache({"last_check": stamp, "latest_version": latest, "ok": True},
                target)
    return latest


def start_background_check() -> threading.Thread | None:
    """Kick the check off on a daemon thread and return immediately.

    Returns None when the check is disabled or the thread cannot start;
    either way the caller carries on to serve.
    """
    if disabled():
        return None
    try:
        thread = threading.Thread(
            target=run_check, name="ks4w-update-check", daemon=True,
        )
        thread.start()
        return thread
    except Exception:
        return None


# ------------------------------------------------------------- the surface


def update_notice() -> str | None:
    """One calm line when a newer stable release exists, else None.

    Reads the cache only: no network, no blocking, safe to call from any
    tool. Returns None when the check is off, when nothing is cached, when
    the versions cannot be parsed, or when the running build is current.
    """
    if disabled():
        return None
    try:
        cache = read_cache()
        if not cache:
            return None
        latest_raw = cache.get("latest_version")
        current_raw = current_version()
        if not latest_raw or not current_raw:
            return None
        latest = Version(str(latest_raw))
        running = Version(str(current_raw))
        if latest.is_prerelease or latest.is_devrelease:
            return None
        if latest <= running:
            return None
        return (
            f"{PRODUCT} {running} is running; {latest} is available. "
            f"Update: pip install -U {PACKAGE} (or download the new "
            f"installer from the releases page)."
        )
    except Exception:
        return None
