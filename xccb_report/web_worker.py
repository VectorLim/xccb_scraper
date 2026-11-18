"""Run live web reports as the signed-in Windows user."""

import argparse
import json
import logging
import os
import shutil
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from time import monotonic, sleep
from urllib.parse import urlsplit

import requests
from openpyxl import DEFUSEDXML

from .evidence import write_json
from .organizations import OrganizationCache
from .pipeline import report
from .site import WindowsReader

POLL_SECONDS = 5
logger = logging.getLogger(__name__)


def generate(directory):
    job = json.loads((directory / "job.json").read_text(encoding="utf-8"))
    config = job["config"] | {
        "reviewer_names": job["names"],
        "assigned_by_names": job["names"],
        "action_scope": "document",
        "reuse_evidence": False,
    }
    progress_file = directory / "progress.json"
    last_update = 0.0

    def progress(counts, _group, document):
        nonlocal last_update
        if document is not None and monotonic() - last_update >= 1:
            try:
                write_json(progress_file, {"processed": counts["Qualified"]})
            except PermissionError:
                # Windows can deny replacement while the controller is reading the prior count.
                return
            last_update = monotonic()

    reader = WindowsReader(
        interval=config["request_interval_seconds"],
        transport=config["http_transport"],
    )
    report(
        directory / "source.xlsx",
        directory / "report.xlsx",
        config,
        directory / "evidence",
        reader=reader,
        all_dates=True,
        on_progress=progress,
        organization_cache=OrganizationCache(directory / "organizations.txt"),
        include_local_paths=False,
    )


def call(client, base, method, route, **kwargs):
    response = client.request(
        method,
        base + route,
        allow_redirects=False,
        timeout=(5, 20),
        **kwargs,
    )
    if not 200 <= response.status_code < 300:
        response.close()
        raise requests.HTTPError("Worker request rejected", response=response)
    return response


def run_job(client, base, job):
    route = "/worker/jobs/" + job["id"]
    deadline = monotonic() + job["timeout"]
    with TemporaryDirectory(prefix="xccb-worker-") as scratch:
        directory = Path(scratch)
        with (
            call(client, base, "GET", route + "/source", stream=True) as response,
            (directory / "source.xlsx").open("wb") as source,
        ):
            shutil.copyfileobj(response.raw, source)
        write_json(directory / "job.json", job)
        with subprocess.Popen(
            [
                sys.executable,
                "-m",
                "xccb_report.web_worker",
                "--generate",
                str(directory),
            ],
        ) as process:
            try:
                while process.poll() is None:
                    remaining = deadline - monotonic()
                    if remaining <= 0:
                        raise TimeoutError("Report timed out")
                    try:
                        process.wait(timeout=min(POLL_SECONDS, remaining))
                    except subprocess.TimeoutExpired:
                        progress = directory / "progress.json"
                        value = (
                            json.loads(progress.read_text(encoding="utf-8"))
                            if progress.exists()
                            else {"processed": 0}
                        )
                        try:
                            call(
                                client, base, "POST", route + "/progress", json=value
                            ).close()
                        except (requests.ConnectionError, requests.Timeout):
                            continue
            finally:
                if process.poll() is None:
                    process.kill()
                    process.wait()
            if process.returncode:
                raise ValueError("Report generation failed")
        with (directory / "report.xlsx").open("rb") as result:
            call(
                client, base, "POST", route + "/result", files={"report": result}
            ).close()


def serve(base, token):
    with requests.Session() as client:
        client.headers["Authorization"] = "Bearer " + token
        while True:
            try:
                with call(client, base, "POST", "/worker/claim") as response:
                    job = response.json() if response.status_code == 200 else None
            except requests.RequestException as exc:
                logger.warning(
                    "Worker could not poll for a report (%s)", type(exc).__name__
                )
            else:
                if job is not None:
                    try:
                        run_job(client, base, job)
                    except (OSError, ValueError, requests.RequestException) as exc:
                        logger.warning("Report job failed (%s)", type(exc).__name__)
                        try:
                            call(
                                client, base, "POST", f"/worker/jobs/{job['id']}/failed"
                            ).close()
                        except requests.RequestException:
                            logger.warning(
                                "Could not return job failure to the web service"
                            )
                    continue
            sleep(POLL_SECONDS)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--generate", type=Path, help=argparse.SUPPRESS)
    options = parser.parse_args()
    if os.name != "nt":
        parser.error("The live report worker requires the signed-in Windows session.")
    if not DEFUSEDXML:
        parser.error("Install defusedxml and enable hardened workbook XML parsing.")
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if options.generate is not None:
        try:
            generate(options.generate)
        except Exception as exc:  # noqa: BLE001 - the child must exit without logging report data.
            logger.error("Report generation failed (%s)", type(exc).__name__)
            return 1
        return 0
    base = os.getenv("XCCB_WEB_URL", "").rstrip("/")
    location = urlsplit(base)
    if (
        location.scheme != "https"
        or not location.hostname
        or location.username
        or location.password
        or location.query
        or location.fragment
        or location.path
    ):
        parser.error("Set XCCB_WEB_URL to the HTTPS origin of the report app.")
    secret_file = os.getenv("XCCB_WORKER_SECRET_FILE")
    if not secret_file:
        parser.error("Set XCCB_WORKER_SECRET_FILE to the worker token file.")
    token = Path(secret_file).read_text(encoding="utf-8-sig").strip()
    if len(token) < 32:
        parser.error("The worker token must be at least 32 characters.")
    try:
        serve(base, token)
    except KeyboardInterrupt:
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
