"""The startup update check: cache aging, opt-out, prereleases, silence.

Nothing here touches the network. Every test that would reach PyPI patches
the fetch, and one test proves the opt-out short-circuits BEFORE any file
or network I/O by installing landmines on both.

The properties pinned:

1. HORIZONS. A successful answer is trusted for 14 days; a failed attempt
   is retried after 1 day, so an offline machine is neither hammered nor
   stuck on a stale answer forever.
2. OPT-OUT IS TOTAL. KS4W_NO_UPDATE_CHECK does not merely hide the line;
   it prevents the network call, the cache read, and the cache write.
3. NO PRERELEASES. A prerelease is never offered to a stable build, even
   when it is the newest thing on the index.
4. ONE SURFACE, ONLY WHEN BEHIND. get_workflows carries the line when a
   newer stable exists and carries nothing otherwise. No other tool ever
   mentions it.
5. FAILURE IS SILENCE. Timeouts, connection errors, HTTP errors, and
   malformed payloads raise nothing, print nothing, and surface nothing.
"""

from __future__ import annotations

import json
import tomllib
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest
from packaging.version import Version

from word_mcp import __version__, server
from word_mcp.core import update_check as uc


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    """Every test gets its own cache directory and no opt-out set."""
    monkeypatch.delenv(uc.OPT_OUT_ENV, raising=False)
    monkeypatch.setenv(uc.CACHE_DIR_ENV, str(tmp_path / "state"))
    return tmp_path


def _payload(*versions, prereleases=()):
    releases = {v: [{"filename": f"x-{v}.whl", "yanked": False}]
                for v in list(versions) + list(prereleases)}
    latest = list(versions)[-1] if versions else ""
    return {"info": {"version": latest}, "releases": releases}


def _write_cache(latest, *, ok=True, age_days=0.0):
    stamp = datetime.now(timezone.utc) - timedelta(days=age_days)
    uc.write_cache({"last_check": stamp.isoformat(),
                    "latest_version": latest, "ok": ok})


# --------------------------------------------------------- 1. cache aging


def test_fresh_success_cache_is_not_due():
    assert uc.is_due({"last_check": datetime.now(timezone.utc).isoformat(),
                      "latest_version": "9.9.9", "ok": True}) is False


@pytest.mark.parametrize("age_days,ok,due", [
    (13.0, True, False),   # inside the 14-day success horizon
    (14.5, True, True),    # past it
    (0.5, False, False),   # inside the 1-day retry horizon
    (1.5, False, True),    # past it
])
def test_both_horizons(age_days, ok, due):
    stamp = datetime.now(timezone.utc) - timedelta(days=age_days)
    cache = {"last_check": stamp.isoformat(), "latest_version": "1.0",
             "ok": ok}
    assert uc.is_due(cache) is due


def test_missing_or_garbage_cache_is_due():
    assert uc.is_due(None) is True
    assert uc.is_due({}) is True
    assert uc.is_due({"last_check": "not a date", "ok": True}) is True


def test_fresh_cache_skips_the_network(monkeypatch):
    _write_cache("9.9.9", ok=True, age_days=1.0)
    monkeypatch.setattr(uc, "_fetch", _boom)
    assert uc.run_check() == "9.9.9"


def test_aged_cache_refetches(monkeypatch):
    _write_cache("9.9.9", ok=True, age_days=20.0)
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: _payload("9.9.9",
                                                              "10.0.0"))
    assert uc.run_check() == "10.0.0"
    assert uc.read_cache()["latest_version"] == "10.0.0"


# ------------------------------------------------------------- 2. opt-out


def _boom(*args, **kwargs):
    raise AssertionError("the opt-out did not short-circuit: I/O happened")


@pytest.mark.parametrize("value", ["1", "true", "TRUE", "yes", "on"])
def test_opt_out_values(value, monkeypatch):
    monkeypatch.setenv(uc.OPT_OUT_ENV, value)
    assert uc.disabled() is True


@pytest.mark.parametrize("value", ["", "0", "false", "no", "off"])
def test_non_opt_out_values(value, monkeypatch):
    monkeypatch.setenv(uc.OPT_OUT_ENV, value)
    assert uc.disabled() is False


def test_opt_out_does_no_io_at_all(monkeypatch):
    """No network, no cache read, no cache write, no thread."""
    monkeypatch.setenv(uc.OPT_OUT_ENV, "1")
    monkeypatch.setattr(uc, "_fetch", _boom)
    monkeypatch.setattr(uc, "read_cache", _boom)
    monkeypatch.setattr(uc, "write_cache", _boom)
    assert uc.run_check() is None
    assert uc.start_background_check() is None
    assert uc.update_notice() is None


def test_opt_out_hides_a_real_pending_update(monkeypatch):
    _write_cache("99.0.0")
    assert uc.update_notice() is not None
    monkeypatch.setenv(uc.OPT_OUT_ENV, "1")
    assert uc.update_notice() is None
    assert "update" not in server.get_workflows()


# -------------------------------------------------------- 3. prereleases


def test_prerelease_never_wins_over_stable():
    payload = _payload("2.0.0", "2.0.1", prereleases=("3.0.0rc1", "3.0.0b2"))
    assert uc._latest_stable(payload) == "2.0.1"


def test_prerelease_only_index_yields_nothing():
    assert uc._latest_stable(_payload(prereleases=("1.0.0rc1",))) is None


def test_info_version_fallback_rejects_a_prerelease():
    assert uc._latest_stable({"info": {"version": "3.0.0rc1"}}) is None
    assert uc._latest_stable({"info": {"version": "3.0.0"}}) == "3.0.0"


def test_yanked_and_empty_releases_are_skipped():
    payload = {"info": {"version": "1.0.0"}, "releases": {
        "1.0.0": [{"filename": "a", "yanked": False}],
        "1.1.0": [{"filename": "b", "yanked": True}],
        "1.2.0": [],
    }}
    assert uc._latest_stable(payload) == "1.0.0"


def test_a_cached_prerelease_is_never_surfaced():
    _write_cache("99.0.0rc1")
    assert uc.update_notice() is None


# ---------------------------------------------------------- 4. surfacing


def _one_version_older(version):
    """A version string strictly below the running one, or None when there
    is none (a 0.0.0 pre-release build)."""
    v = Version(version)
    if v.micro:
        return f"{v.major}.{v.minor}.{v.micro - 1}"
    if v.minor:
        return f"{v.major}.{v.minor - 1}.99"
    if v.major:
        return f"{v.major - 1}.99.99"
    return None


def test_notice_only_when_behind():
    _write_cache(__version__)
    assert uc.update_notice() is None
    older = _one_version_older(__version__)
    if older is not None:
        _write_cache(older)
        assert uc.update_notice() is None
    _write_cache("999.0.0")
    line = uc.update_notice()
    assert line is not None
    assert f"{uc.PRODUCT} {__version__} is running" in line
    assert "999.0.0 is available" in line
    assert f"pip install -U {uc.PACKAGE}" in line


def test_no_cache_no_notice():
    assert uc.update_notice() is None


def test_get_workflows_carries_the_line_only_when_behind():
    _write_cache(__version__)
    assert "update" not in server.get_workflows()
    _write_cache("999.0.0")
    out = server.get_workflows()
    assert out["update"] == uc.update_notice()
    # and the per-task call carries it too, without disturbing the recipe
    task = server.get_workflows("bulk-edit")
    assert task.get("update") == uc.update_notice()


def test_no_other_tool_mentions_the_update(tmp_path):
    """The line lives in get_workflows and nowhere else."""
    _write_cache("999.0.0")
    doc = str(tmp_path / "d.docx")
    server.create_document(doc)
    for result in (server.get_document_info(doc),
                   server.diagnose_document(doc),
                   server.get_outline(doc)):
        assert "update" not in result


# ------------------------------------------------------ 5. silent failure


class _Timeout(Exception):
    pass


@pytest.mark.parametrize("blow_up", [
    TimeoutError("timed out"),
    OSError("network unreachable"),
    ValueError("not json"),
    _Timeout("whatever the stack raises"),
])
def test_network_failure_is_silent(blow_up, monkeypatch, capsys):
    def raiser(*a, **k):
        raise blow_up

    monkeypatch.setattr(uc, "_fetch", raiser)
    assert uc.run_check() is None
    assert uc.update_notice() is None
    captured = capsys.readouterr()
    assert captured.out == "" and captured.err == ""


def test_failed_attempt_is_recorded_with_the_short_horizon(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: (_ for _ in ()).throw(
        OSError("down")))
    uc.run_check()
    cache = uc.read_cache()
    assert cache["ok"] is False
    assert cache["latest_version"] is None
    assert uc.is_due(cache) is False               # not hammered today
    stamp = datetime.now(timezone.utc) - timedelta(days=2)
    assert uc.is_due({**cache, "last_check": stamp.isoformat()}) is True


def test_a_failure_keeps_the_last_known_answer(monkeypatch):
    _write_cache("999.0.0", ok=True, age_days=30.0)
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: (_ for _ in ()).throw(
        OSError("down")))
    uc.run_check()
    cache = uc.read_cache()
    assert cache["ok"] is False
    assert cache["latest_version"] == "999.0.0"
    assert uc.update_notice() is not None          # still useful offline


@pytest.mark.parametrize("payload", [
    {}, {"info": {}}, {"releases": "nonsense"},
    {"info": {"version": "not.a.version"}},
    {"releases": {"not.a.version": []}},
])
def test_malformed_payloads_are_treated_as_failures(payload, monkeypatch):
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: payload)
    assert uc.run_check() is None
    assert uc.read_cache()["ok"] is False


def test_an_unwritable_cache_dir_does_not_raise(monkeypatch):
    monkeypatch.setattr(uc, "cache_path", _boom)
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: _payload("1.0.0"))
    # run_check resolves the path itself; a raising resolver is the worst
    # case, and the caller (a daemon thread) must still not see it.
    thread = uc.start_background_check()
    if thread is not None:
        thread.join(timeout=5)
    assert uc.update_notice() is None


# ------------------------------------------------------------ plumbing


def test_background_check_never_blocks(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: _payload("1.0", "2.0"))
    thread = uc.start_background_check()
    assert thread is not None and thread.daemon is True
    thread.join(timeout=10)
    assert uc.read_cache()["latest_version"] == "2.0"


def test_cache_lives_in_the_state_dir_not_beside_documents(monkeypatch):
    monkeypatch.delenv(uc.CACHE_DIR_ENV, raising=False)
    monkeypatch.setenv("LOCALAPPDATA", r"C:\Users\someone\AppData\Local")
    path = uc.cache_path()
    assert path.parent.name == "word-mcp"
    assert path.name == "update-check.json"


def test_version_constant_matches_pyproject():
    """A release bump must land in both places, or the update check tells
    users they are running the wrong version."""
    root = Path(__file__).resolve().parents[2]
    data = tomllib.loads((root / "pyproject.toml").read_text(
        encoding="utf-8"))
    assert data["project"]["version"] == __version__


def test_cache_file_is_small_and_boring():
    _write_cache("1.2.3")
    raw = uc.cache_path().read_text(encoding="utf-8")
    assert len(raw) < 200
    assert set(json.loads(raw)) == {"last_check", "latest_version", "ok"}
