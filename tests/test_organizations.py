from dataclasses import replace
from concurrent.futures import ThreadPoolExecutor
import json
import subprocess
from threading import Barrier, Event
from types import SimpleNamespace

import pytest

from xccb_report.organizations import OrganizationCache, load_settings, query_directory, save_settings, user_directory


def candidate(wwid="123", code="04643", description="KuAT AE CAM"):
    return {"wwid": wwid, "organization_code": code, "organization": description}


def test_cache_reuse_restart_and_normalization(tmp_path):
    path = tmp_path / "cache.txt"
    calls = []
    def query(name, wwid):
        calls.append((name, wwid))
        return [candidate()]
    cache = OrganizationCache(path)
    first, issues = cache.resolve(["Doe, Jane", " DOE,  JANE "], query)
    assert not issues and len(calls) == 1
    again, issues = OrganizationCache(path).resolve(["doe, jane"], query)
    assert again == first and len(calls) == 1
    assert again["doe, jane"].organization_code == "04643"


def test_ambiguous_confirmation_and_refresh_failure(tmp_path):
    cache = OrganizationCache(tmp_path / "cache.txt")
    query = lambda *args: [candidate(), candidate("456", "64530")]
    result, issues = cache.resolve(["Jane"], query)
    assert not result and "ambiguous" in issues["jane"]
    result, _ = cache.resolve(["Jane"], query, confirm=lambda name, rows: "456")
    assert result["jane"].identity_confirmed == "yes"
    def fail(*args):
        raise OSError("offline")
    result, issues = cache.resolve(["Jane"], fail, refresh=True)
    assert result["jane"].wwid == "456" and "retained" in issues["jane"]
    calls = []
    result, _ = cache.resolve(["Jane"], lambda name, wwid: calls.append(wwid) or [candidate("456", "001")], refresh=True)
    assert calls == ["456"] and result["jane"].organization_code == "001"


def test_roundtrip_merge_and_conflict(tmp_path):
    path = tmp_path / "cache.txt"
    first, second = OrganizationCache(path), OrganizationCache(path)
    first.resolve(['Doe, "Jane"\tX'], lambda *args: [candidate(description='Org\t"quoted"\nunit')])
    second.resolve(["Other"], lambda *args: [candidate("456")])
    loaded = OrganizationCache(path)
    assert len(loaded.entries) == 2
    assert loaded.entries['doe, "jane" x'].organization == 'Org\t"quoted"\nunit'
    old = loaded.entries["other"]
    second.save(replace(old, organization="Changed"))
    with pytest.raises(ValueError, match="another instance"):
        loaded.save(replace(old, organization="Conflict"))


def test_malformed_and_atomic_failure_preserve_file(tmp_path, monkeypatch):
    path = tmp_path / "cache.txt"
    cache = OrganizationCache(path)
    cache.resolve(["Jane"], lambda *args: [candidate()])
    previous = path.read_bytes()
    monkeypatch.setattr("xccb_report.organizations.os.replace", lambda *args: (_ for _ in ()).throw(OSError("blocked")))
    cache.resolve(["Other"], lambda *args: [candidate()])
    assert path.read_bytes() == previous
    assert not list(tmp_path.glob("*.tmp"))
    path.write_text("bad\theader\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Malformed"):
        OrganizationCache(path)
    with pytest.raises(ValueError, match="Malformed"):
        cache.save(next(iter(cache.entries.values())))
    assert path.read_text() == "bad\theader\n"


def test_private_paths_settings_and_malformed_settings(tmp_path, monkeypatch):
    for user in ("alice", "bob"):
        monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / user))
        assert load_settings()["organization_codes"] == []
        save_settings(["04643", "64530"])
        assert load_settings()["organization_codes"] == ["04643", "64530"]
        assert user_directory().parent == tmp_path / user
    (user_directory() / "settings.json").write_text('{"organization_codes":[4643]}')
    with pytest.raises(ValueError, match="Malformed"):
        save_settings([])


def test_outage_is_bounded_and_incomplete_rows_are_not_cached(tmp_path):
    cache = OrganizationCache(tmp_path / "cache.txt")
    calls = []
    def fail(name, wwid):
        calls.append(name)
        raise OSError("private connection detail")
    results, issues = cache.resolve(["One", "Two", "Three"], fail)
    assert calls == ["One"] and not results and len(issues) == 3
    assert "private connection detail" not in str(issues)
    results, issues = cache.resolve(["One"], lambda *args: [candidate(code=None)])
    assert not results and not cache.entries and "incomplete" in issues["one"]


def test_concurrent_instances_merge_without_losing_names(tmp_path):
    path = tmp_path / "cache.txt"
    barrier = Barrier(2)
    def run(name):
        cache = OrganizationCache(path)
        barrier.wait(timeout=5)
        result, issues = cache.resolve([name], lambda *args: [candidate(name)])
        assert not issues and result
    with ThreadPoolExecutor(max_workers=2) as executor:
        list(executor.map(run, ["One", "Two"]))
    assert set(OrganizationCache(path).entries) == {"one", "two"}


def test_timeout_and_cancellation_preserve_success(tmp_path, monkeypatch):
    cache = OrganizationCache(tmp_path / "cache.txt")
    cache.resolve(["One"], lambda *args: [candidate()])
    calls = []
    def timeout(*args, **kwargs):
        calls.append(kwargs)
        raise subprocess.TimeoutExpired("python", kwargs["timeout"])
    monkeypatch.setattr("xccb_report.organizations.subprocess.run", timeout)
    result, issues = cache.resolve(["One", "Two"], query_directory, refresh=True)
    assert len(calls) == 1 and calls[0]["timeout"] == 45
    assert result["one"].wwid == "123" and "TimeoutExpired" in issues["one"]
    cancelled = Event()
    cancelled.set()
    result, issues = cache.resolve(["One", "Two"], query_directory, refresh=True, cancel=cancelled)
    assert len(calls) == 1 and "stopped" in issues["two"]


def test_worker_uses_selected_python_without_bundling_datasyncx(tmp_path, monkeypatch):
    calls = []
    monkeypatch.setattr("xccb_report.organizations.subprocess.run", lambda *args, **kwargs:
                        calls.append((args, kwargs)) or SimpleNamespace(returncode=0, stdout=json.dumps([candidate()])))
    assert query_directory("Doe, Jane", python="approved-python.exe")[0]["organization_code"] == "04643"
    command = calls[0][0][0]
    assert command[0] == "approved-python.exe" and command[1] == "-c"
    assert "ALL.EMPLOYEE_HRDS" in command[2]


@pytest.mark.parametrize("content", ["", "wrong\theader\n", 'display_name\tlookup_name\n"unterminated',
                                    "\xff\xfe"])
def test_malformed_cache_is_never_rewritten(tmp_path, content):
    path = tmp_path / "cache.txt"
    path.write_text(content, encoding="utf-8")
    with pytest.raises(ValueError, match="Malformed"):
        OrganizationCache(path)
    assert path.read_text(encoding="utf-8") == content