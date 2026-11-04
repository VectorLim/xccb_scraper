from dataclasses import replace
from concurrent.futures import ThreadPoolExecutor
import json
import subprocess
from threading import Barrier, Event
from types import SimpleNamespace

import pytest

from xccb_report.organizations import (
    OrganizationCache,
    load_settings,
    query_directory,
    query_directory_batch,
    save_settings,
    user_directory,
)


def candidate(wwid="123", code="04643", description="KuAT AE CAM"):
    return {"wwid": wwid, "organization_code": code, "organization": description}


def test_batch_exact_names_without_identity_confirmation(tmp_path):
    cache = OrganizationCache(tmp_path / "cache.txt")
    calls = []

    def batch(names):
        calls.append(names)
        return {
            "doe, jane": [candidate() | {"name": "Doe, Jane"}],
            "same name": [candidate("1") | {"name": "Same Name"},
                          candidate("2") | {"name": "Same Name"}],
            "other": [candidate("3") | {"name": "Other Nickname"}],
        }

    result, issues = cache.resolve(
        ["Doe, Jane", " DOE,  JANE ", "Same Name", "Other"],
        batch_query=batch,
        confirm=lambda *args: pytest.fail("Exact-name lookup must not prompt"),
    )
    assert calls == [["Doe, Jane", "Same Name", "Other"]]
    assert set(result) == {"doe, jane"}
    assert "ambiguous" in issues["same name"]
    assert "no active HR match" in issues["other"]
    again, issues = OrganizationCache(cache.path).resolve(
        ["doe, jane"], batch_query=lambda *args: pytest.fail("Exact cache queried")
    )
    assert again == result and not issues


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


def test_legacy_mapping_requires_exact_revalidation(tmp_path):
    import csv
    from dataclasses import asdict
    from xccb_report.organizations import FIELDS

    path = tmp_path / "legacy.txt"
    cache = OrganizationCache(path)
    cache.resolve(["Jane"], lambda *args: [candidate("old-nickname-match")])
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS[:-1], delimiter="\t", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(asdict(entry) for entry in cache.entries.values())
    cache = OrganizationCache(path)
    previous = path.read_bytes()
    def offline(names):
        raise OSError("private details")
    results, issues = cache.resolve(["Jane"], batch_query=offline)
    assert not results and "Directory unavailable" in issues["jane"]
    assert path.read_bytes() == previous
    results, issues = cache.resolve(["Jane"], batch_query=lambda names: {
        "jane": [candidate("exact-identity") | {"name": "Jane"}]
    })
    assert not issues and results["jane"].wwid == "exact-identity"
    assert OrganizationCache(path).entries["jane"].matched_name == "Jane"


def test_default_exact_cache_preserves_older_app_cache(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    user_directory().mkdir()
    legacy = user_directory() / "author_manager_organizations.txt"
    legacy.write_bytes(b"old app data must remain unchanged")
    cache = OrganizationCache()
    assert not cache.entries and cache.path != legacy
    cache.resolve(["Jane"], batch_query=lambda names: {"jane": [candidate() | {"name": "Jane"}]})
    assert legacy.read_bytes() == b"old app data must remain unchanged"


def test_exact_refresh_does_not_reuse_identity_when_name_becomes_ambiguous(tmp_path):
    cache = OrganizationCache(tmp_path / "cache.txt")
    cache.resolve(["Jane"], batch_query=lambda names: {"jane": [candidate() | {"name": "Jane"}]})
    results, issues = cache.resolve(["Jane"], refresh=True, batch_query=lambda names: {
        "jane": [candidate() | {"name": "Jane"}, candidate("other") | {"name": "Jane"}]
    })
    assert not results and "ambiguous" in issues["jane"]
    assert cache.entries["jane"].wwid == "123"


def test_invalid_batch_response_is_reported_for_every_name(tmp_path):
    results, issues = OrganizationCache(tmp_path / "cache.txt").resolve(
        ["One", "Two"], batch_query=lambda names: []
    )
    assert not results and set(issues) == {"one", "two"}


def test_worker_uses_one_exact_full_name_sql_request(tmp_path, monkeypatch, capsys):
    import io
    import sys
    from xccb_report import organization_worker

    helper = tmp_path / "datasyncx" / "utils" / "get_user_email.py"
    helper.parent.mkdir(parents=True)
    helper.write_text("username = 'test_fixture_only'", encoding="utf-8")
    calls = []

    def query(**kwargs):
        calls.append(kwargs["sql"])
        return SimpleNamespace(to_dict=lambda **options: [
            {"WWID": "1", "CCMAILNAME": "Doe, Jane", "SITECODE": "KM", "ORGUNITCODE": "001", "ORGUNITDESCR": "Exact Org", "GL_COST_CENTER_DESCR": "Cost"},
            {"WWID": "2", "CCMAILNAME": "Doe, Jane Jr", "SITECODE": "KM", "ORGUNITCODE": "002", "ORGUNITDESCR": "Wrong Org", "GL_COST_CENTER_DESCR": "Cost"},
        ])

    monkeypatch.setitem(sys.modules, "datasyncx", SimpleNamespace(__file__=str(helper.parent.parent / "__init__.py")))
    monkeypatch.setitem(sys.modules, "datasyncx.utils", SimpleNamespace())
    monkeypatch.setitem(sys.modules, "datasyncx.utils.oracle_db", SimpleNamespace(query_oracle=query))
    monkeypatch.setitem(sys.modules, "pandas", SimpleNamespace(isna=lambda value: value is None))
    names = ["Doe, Jane", " DOE,  JANE ", "O'Brien"] + [f"Manager {index}" for index in range(1000)]
    monkeypatch.setattr(sys, "argv", ["worker"])
    monkeypatch.setattr(sys, "stdin", io.StringIO(json.dumps({"names": names})))
    assert organization_worker.main() == 0
    result = json.loads(capsys.readouterr().out)
    assert len(calls) == 1 and calls[0].count("REGEXP_REPLACE(") == 2
    assert "'O''BRIEN'" in calls[0] and calls[0].count("'DOE, JANE'") == 1
    assert "NICKNAME" not in calls[0] and "FIRSTNAME" not in calls[0] and " LIKE " not in calls[0]
    assert "GENIEMPSTATUS NOT IN ('R', 'T')" in calls[0]
    assert len(result["doe, jane"]) == 1 and result["doe, jane"][0]["organization_code"] == "001"
    assert "doe, jane jr" not in result and result["o'brien"] == []


def test_ambiguous_names_never_prompt_and_refresh_failure(tmp_path):
    cache = OrganizationCache(tmp_path / "cache.txt")
    query = lambda *args: [candidate(), candidate("456", "64530")]
    result, issues = cache.resolve(["Jane"], query)
    assert not result and "ambiguous" in issues["jane"]
    result, _ = cache.resolve(["Jane"], query, confirm=lambda *args: pytest.fail("Unexpected confirmation"))
    assert not result
    result, _ = cache.resolve(["Jane"], lambda *args: [candidate("456", "64530")])
    assert result["jane"].identity_confirmed == "no"

    def fail(*args):
        raise OSError("offline")

    result, issues = cache.resolve(["Jane"], fail, refresh=True)
    assert result["jane"].wwid == "456" and "retained" in issues["jane"]
    calls = []
    result, _ = cache.resolve(
        ["Jane"],
        lambda name, wwid: calls.append(wwid) or [candidate("456", "001")],
        refresh=True,
    )
    assert calls == ["456"] and result["jane"].organization_code == "001"


def test_roundtrip_merge_and_conflict(tmp_path):
    path = tmp_path / "cache.txt"
    first, second = OrganizationCache(path), OrganizationCache(path)
    first.resolve(
        ['Doe, "Jane"\tX'], lambda *args: [candidate(description='Org\t"quoted"\nunit')]
    )
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
    monkeypatch.setattr(
        "xccb_report.organizations.os.replace",
        lambda *args: (_ for _ in ()).throw(OSError("blocked")),
    )
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
    assert 1 <= len(calls) <= 3 and not results and len(issues) == 3
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
    assert 1 <= len(calls) <= 2 and calls[0]["timeout"] == 45
    assert result["one"].wwid == "123" and "TimeoutExpired" in issues["one"]
    cancelled = Event()
    cancelled.set()
    previous_calls = len(calls)
    result, issues = cache.resolve(
        ["One", "Two"], query_directory, refresh=True, cancel=cancelled
    )
    assert len(calls) == previous_calls and "stopped" in issues["two"]


def test_lookups_overlap_with_bounded_workers_and_deduplication(tmp_path):
    from threading import Lock
    cache = OrganizationCache(tmp_path / "cache.txt")
    overlap = Barrier(4)
    lock = Lock()
    active, peak, calls = [0], [0], []
    def query(name, wwid):
        with lock:
            calls.append(name)
            active[0] += 1
            peak[0] = max(peak[0], active[0])
        overlap.wait(timeout=5)
        with lock:
            active[0] -= 1
        return [candidate(name)]
    results, issues = cache.resolve(["One", "Two", "Three", "Four", " ONE "], query)
    assert not issues and len(results) == 4
    assert peak[0] == 4 and len(calls) == 4
    again, issues = OrganizationCache(cache.path).resolve(["One", "Two", "Three", "Four"], lambda *args: pytest.fail("cached entry queried"))
    assert again == results and not issues


def test_concurrent_outage_stops_submitting_after_inflight_batch(tmp_path):
    from xccb_report.organizations import DirectoryUnavailable
    overlap = Barrier(4)
    calls = []
    def query(name, wwid):
        calls.append(name)
        overlap.wait(timeout=5)
        raise DirectoryUnavailable("Directory offline")
    cache = OrganizationCache(tmp_path / "cache.txt")
    results, issues = cache.resolve([f"Manager {index}" for index in range(12)], query)
    assert len(calls) == 4 and not results and len(issues) == 12
    assert not cache.entries


def test_worker_uses_selected_python_without_bundling_datasyncx(tmp_path, monkeypatch):
    calls = []
    monkeypatch.setattr(
        "xccb_report.organizations.subprocess.run",
        lambda *args, **kwargs: calls.append((args, kwargs))
        or SimpleNamespace(returncode=0, stdout=json.dumps([candidate()])),
    )
    assert (
        query_directory("Doe, Jane", python="approved-python.exe")[0][
            "organization_code"
        ]
        == "04643"
    )
    command = calls[0][0][0]
    assert command[0] == "approved-python.exe" and command[1] == "-c"
    assert "ALL.EMPLOYEE_HRDS" in command[2]


def test_frozen_lookup_uses_bundled_executable_not_saved_python(tmp_path, monkeypatch):
    from pathlib import Path
    monkeypatch.setattr("xccb_report.organizations.sys.frozen", True, raising=False)
    calls = []
    def run(command, **kwargs):
        calls.append((command, kwargs))
        assert json.loads(Path(command[2]).read_text(encoding="utf-8")) == ["Doe, Jane", "123"]
        Path(command[3]).write_text(json.dumps([candidate()]), encoding="utf-8")
        return SimpleNamespace(returncode=0)
    monkeypatch.setattr("xccb_report.organizations.subprocess.run", run)
    rows = query_directory("Doe, Jane", "123", python="obsolete-external-python.exe")
    assert rows[0]["organization_code"] == "04643"
    assert calls[0][0][0] != "obsolete-external-python.exe"
    assert calls[0][0][1] == "--organization-worker"
    assert calls[0][1]["timeout"] == 45
    assert not Path(calls[0][0][2]).exists()


def test_frozen_batch_uses_one_worker(tmp_path, monkeypatch):
    from pathlib import Path
    monkeypatch.setattr("xccb_report.organizations.sys.frozen", True, raising=False)
    calls = []

    def run(command, **kwargs):
        calls.append(command)
        assert json.loads(Path(command[2]).read_text(encoding="utf-8")) == {
            "names": ["Doe, Jane", "Other"]
        }
        Path(command[3]).write_text(json.dumps({"doe, jane": [candidate()], "other": []}), encoding="utf-8")
        return SimpleNamespace(returncode=0)

    monkeypatch.setattr("xccb_report.organizations.subprocess.run", run)
    assert set(query_directory_batch(["Doe, Jane", "Other"])) == {"doe, jane", "other"}
    assert len(calls) == 1
    assert query_directory_batch([]) == {} and len(calls) == 1


@pytest.mark.parametrize(
    "content",
    ["", "wrong\theader\n", 'display_name\tlookup_name\n"unterminated', "\xff\xfe"],
)
def test_malformed_cache_is_never_rewritten(tmp_path, content):
    path = tmp_path / "cache.txt"
    path.write_text(content, encoding="utf-8")
    with pytest.raises(ValueError, match="Malformed"):
        OrganizationCache(path)
    assert path.read_text(encoding="utf-8") == content
