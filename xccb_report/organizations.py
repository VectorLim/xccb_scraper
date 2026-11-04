"""Private per-user organization mappings and bounded DataSyncX lookups."""

import csv
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import subprocess
import sys
from tempfile import NamedTemporaryFile
from threading import Lock
import time

from .rules import normalized


def user_directory():
    base = os.environ.get("LOCALAPPDATA")
    return (Path(base) if base else Path.home() / "AppData" / "Local") / "xCCB_Report"


@contextmanager
def locked(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.with_suffix(path.suffix + ".lock").open("a+b") as stream:
        stream.seek(0, 2)
        if not stream.tell():
            stream.write(b"0")
            stream.flush()
        deadline = time.monotonic() + 5
        while True:
            try:
                stream.seek(0)
                if os.name == "nt":
                    import msvcrt

                    msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise ValueError(f"File is busy: {path}") from None
                time.sleep(0.05)
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == "nt":
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream, fcntl.LOCK_UN)


def atomic_write(path, writer):
    temporary = None
    try:
        with NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="",
            dir=path.parent,
            prefix="." + path.name,
            suffix=".tmp",
            delete=False,
        ) as stream:
            temporary = Path(stream.name)
            writer(stream)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


@dataclass(frozen=True)
class Organization:
    display_name: str
    lookup_name: str
    wwid: str
    organization_code: str
    organization: str
    looked_up_utc: str
    identity_confirmed: str = "no"
    matched_name: str = ""


FIELDS = tuple(Organization.__dataclass_fields__)


class OrganizationCache:
    def __init__(self, path=None):
        self.path = (
            Path(path)
            if path is not None
            else user_directory() / "author_manager_organizations_exact.txt"
        )
        self.entries = self.load()

    def load(self):
        if not self.path.exists():
            return {}
        try:
            with self.path.open(encoding="utf-8-sig", newline="") as stream:
                reader = csv.DictReader(stream, delimiter="\t", strict=True)
                if reader.fieldnames not in (list(FIELDS), list(FIELDS[:-1])):
                    raise ValueError("unexpected header")
                entries = {}
                for row in reader:
                    if set(row) != set(reader.fieldnames) or any(
                        value is None for value in row.values()
                    ):
                        raise ValueError("incomplete row")
                    entry = Organization(**row)
                    stamp = datetime.fromisoformat(entry.looked_up_utc)
                    if (
                        not entry.lookup_name
                        or entry.lookup_name != normalized(entry.display_name)
                        or not entry.wwid.strip()
                        or not entry.organization_code.strip()
                        or not entry.organization.strip()
                        or stamp.tzinfo is None
                        or entry.identity_confirmed not in {"yes", "no"}
                        or (entry.matched_name and normalized(entry.matched_name) != entry.lookup_name)
                        or entry.lookup_name in entries
                    ):
                        raise ValueError("invalid or duplicate mapping")
                    entries[entry.lookup_name] = entry
                return entries
        except (ValueError, TypeError, UnicodeError, csv.Error) as exc:
            raise ValueError(
                f"Malformed organization cache {self.path}: {exc}. Repair the file before saving."
            ) from None

    def save(self, entry):
        with locked(self.path):
            entries = self.load()
            existing = entries.get(entry.lookup_name)
            if (
                existing is not None
                and existing != self.entries.get(entry.lookup_name)
                and existing != entry
            ):
                raise ValueError(
                    "Organization mapping changed in another instance; reload before confirming or refreshing"
                )
            entries[entry.lookup_name] = entry

            def write(stream):
                writer = csv.DictWriter(stream, fieldnames=FIELDS, delimiter="\t")
                writer.writeheader()
                writer.writerows(asdict(entries[key]) for key in sorted(entries))

            atomic_write(self.path, write)
            self.entries = entries

    def resolve(
        self, names, query=None, *, batch_query=None, refresh=False, confirm=None, progress=None, cancel=None
    ):
        results, issues = {}, {}
        unique = {}
        for name in names:
            if normalized(name):
                unique.setdefault(normalized(name), str(name))
        unavailable = ""
        work = []
        for index, (key, name) in enumerate(unique.items(), 1):
            cached = self.entries.get(key)
            if batch_query is not None and cached and not cached.matched_name:
                cached = None
            if cached and not refresh:
                results[key] = cached
            else:
                work.append((index, key, name, cached))

        def completed_lookups():
            nonlocal unavailable
            if batch_query is not None:
                if not work:
                    return
                rows_by_name, error = {}, ""
                try:
                    if cancel and cancel.is_set():
                        raise DirectoryUnavailable("Organization lookup stopped by user")
                    if progress:
                        progress(f"Organization lookup: {len(work)} exact names in one batch")
                    rows_by_name = batch_query([name for _, _, name, _ in work])
                    if not isinstance(rows_by_name, dict):
                        raise ValueError("Invalid organization batch response")
                except Exception as exc:
                    rows_by_name = {}
                    error = (str(exc) if isinstance(exc, DirectoryUnavailable) else
                             f"Directory unavailable ({type(exc).__name__}); check DataSyncX, Oracle client/TNS and network access")
                for item in work:
                    yield item, rows_by_name.get(item[1], []), error
                return
            remaining = iter(work)
            with ThreadPoolExecutor(max_workers=4, thread_name_prefix="xccb-org") as executor:
                pending = {}
                def submit_next():
                    nonlocal unavailable
                    if cancel and cancel.is_set():
                        unavailable = "Organization lookup stopped by user"
                    if unavailable:
                        return
                    item = next(remaining, None)
                    if item is not None:
                        index, key, name, cached = item
                        if progress:
                            progress(f"Organization lookup {index}/{len(unique)}: {name}")
                        pending[executor.submit(query, name, cached.wwid if cached else "")] = item
                for _ in range(4):
                    submit_next()
                while pending:
                    future = next(as_completed(pending))
                    item = pending.pop(future)
                    rows, error = None, ""
                    try:
                        rows = future.result()
                    except Exception as exc:
                        error = (str(exc) if isinstance(exc, DirectoryUnavailable) else
                                 f"Directory unavailable ({type(exc).__name__}); check DataSyncX Python, Oracle client/TNS and network access")
                        unavailable = error
                    yield item, rows, error
                    submit_next()
                if cancel and cancel.is_set():
                    unavailable = "Organization lookup stopped by user"
                for item in remaining:
                    yield item, None, unavailable

        for (_, key, name, cached), rows, error in completed_lookups():
            try:
                if error:
                    raise DirectoryUnavailable(error)
                candidates = {}
                for row in rows:
                    if batch_query is not None and normalized(row.get("name")) != key:
                        continue
                    if any(
                        not isinstance(row.get(field), str) or not row[field].strip()
                        for field in ("wwid", "organization_code", "organization")
                    ):
                        raise ValueError(
                            "HR match has incomplete identity/organization fields"
                        )
                    identity = row["wwid"]
                    if identity in candidates and any(
                        candidates[identity][field] != row[field]
                        for field in ("organization_code", "organization")
                    ):
                        raise ValueError("HR identity has conflicting organizations")
                    candidates[identity] = row
                chosen = candidates.get(cached.wwid) if cached and batch_query is None else None
                confirmed = cached.identity_confirmed if cached and batch_query is None else "no"
                if len(candidates) == 1 and (batch_query is not None or not cached):
                    chosen = next(iter(candidates.values()))
                if chosen is None:
                    issues[key] = (
                        "Identity unresolved: ambiguous name"
                        if candidates
                        else "Identity unresolved: no active HR match"
                    )
                else:
                    entry = Organization(
                        name,
                        key,
                        str(chosen["wwid"]),
                        str(chosen["organization_code"]),
                        chosen["organization"],
                        datetime.now(timezone.utc).isoformat(),
                        confirmed,
                        chosen["name"] if batch_query is not None else "",
                    )
                    if (
                        not entry.organization_code.strip()
                        or not entry.organization.strip()
                    ):
                        raise ValueError("HR match has no organization")
                    self.save(entry)
                    results[key] = entry
            except (
                DirectoryUnavailable,
                OSError,
                ValueError,
                TypeError,
                KeyError,
            ) as exc:
                issues[key] = (
                    str(exc)
                    if isinstance(exc, (DirectoryUnavailable, ValueError))
                    else f"Organization cache update failed ({type(exc).__name__}); check cache permissions"
                )
            if key not in results and cached and not issues.get(key, "").startswith("Identity unresolved:"):
                results[key] = cached
                issues[key] = (
                    issues.get(key, "Refresh unresolved")
                    + "; retained cached mapping from "
                    + cached.looked_up_utc
                )
        return results, issues


def load_settings(path=None):
    path = Path(path) if path is not None else user_directory() / "settings.json"
    if not path.exists():
        return {"organization_codes": []}
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
        codes = value["organization_codes"]
        if (
            not isinstance(value, dict)
            or not isinstance(codes, list)
            or any(not isinstance(code, str) or not code.strip() for code in codes)
        ):
            raise ValueError("organization_codes must be a list of nonblank strings")
        if "organization_python" in value and not isinstance(
            value["organization_python"], str
        ):
            raise ValueError("organization_python must be an executable path string")
        return value
    except (ValueError, KeyError, TypeError, UnicodeError) as exc:
        raise ValueError(f"Malformed organization settings {path}: {exc}") from None


def save_settings(codes, path=None):
    if any(not isinstance(code, str) or not code.strip() for code in codes):
        raise ValueError("Organization codes must be nonblank strings")
    path = Path(path) if path is not None else user_directory() / "settings.json"
    with locked(path):
        settings = load_settings(path)
        settings["organization_codes"] = sorted(set(codes))
        atomic_write(path, lambda stream: json.dump(settings, stream, indent=2))


def save_python(python, path=None):
    path = Path(path) if path is not None else user_directory() / "settings.json"
    with locked(path):
        settings = load_settings(path)
        settings["organization_python"] = str(python)
        atomic_write(path, lambda stream: json.dump(settings, stream, indent=2))


class DirectoryUnavailable(ValueError):
    pass


def query_directory(name, wwid="", python=None):
    return _run_directory([name, wwid], python)


def query_directory_batch(names, python=None):
    if not names:
        return {}
    return _run_directory({"names": list(names)}, python)


def directory_query(python=None, on_unavailable=None, cancel=None):
    """Return a batch query that asks once for a working Python runtime if DataSyncX is unavailable."""
    current = python
    setup_attempted = False
    setup_lock = Lock()

    def query(names):
        nonlocal current, setup_attempted
        attempted = current
        try:
            return query_directory_batch(names, attempted)
        except (DirectoryUnavailable, OSError) as exc:
            with setup_lock:
                if attempted != current:
                    retry = current
                else:
                    if setup_attempted or on_unavailable is None or (cancel and cancel.is_set()):
                        raise
                    setup_attempted = True
                    chosen = on_unavailable(str(exc))
                    if not chosen:
                        raise
                    rows = query_directory_batch(names, chosen)
                    save_python(chosen)
                    current = chosen
                    return rows
                return query_directory_batch(names, retry)

    return query


def _run_directory(payload, python):
    if getattr(sys, "frozen", False):
        from tempfile import TemporaryDirectory
        with TemporaryDirectory(prefix="xccb_org_") as directory:
            request = Path(directory) / "request.json"
            response = Path(directory) / "response.json"
            request.write_text(json.dumps(payload), encoding="utf-8")
            completed = subprocess.run(
                [sys.executable, "--organization-worker", str(request), str(response)],
                capture_output=True, timeout=45,
                cwd=directory,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
            if not response.exists():
                raise DirectoryUnavailable("Bundled DataSyncX worker failed to start; check the application build")
            result = json.loads(response.read_text(encoding="utf-8"))
            if completed.returncode:
                raise DirectoryUnavailable(result.get("error", "Bundled DataSyncX lookup failed"))
            return result
    executable = python or sys.executable
    if not executable:
        raise DirectoryUnavailable(
            "Select a Python installation with DataSyncX in Organizations"
        )
    script = (
        Path(__file__).with_name("organization_worker.py").read_text(encoding="utf-8")
    )
    completed = subprocess.run(
        [executable, "-c", script],
        input=json.dumps(payload),
        text=True,
        encoding="utf-8",
        capture_output=True,
        timeout=45,
        env=os.environ | {"PYTHONIOENCODING": "utf-8"},
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
    )
    if completed.returncode:
        try:
            error = json.loads(completed.stdout)["error"]
        except (ValueError, KeyError, TypeError):
            error = "DataSyncX directory query unavailable; check Python, Oracle client/TNS and network access"
        raise DirectoryUnavailable(error)
    return json.loads(completed.stdout)
