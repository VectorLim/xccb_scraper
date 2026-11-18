"""Internal browser UI and job handoff to one Windows worker."""

import os
import secrets
from dataclasses import dataclass, field
from hmac import compare_digest
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Lock, Thread
from time import monotonic, sleep
from zipfile import BadZipFile, ZipFile

from flask import Flask, abort, jsonify, render_template, request, send_file, session
from werkzeug.exceptions import HTTPException

from .actions import workspace_identity
from .config import DEFAULTS, load_config
from .results import DOCUMENT_FIELDS, read_results
from .reviewers import parse_names
from .workbook import MANAGER_HEADERS

LEASE_SECONDS = 60
RESULT_SECONDS = 3600
ACTIVE_STATES = {"queued", "running"}
DOCUMENT_HEADERS = tuple(
    "Workspace" if key == "url" else label for key, label in DOCUMENT_FIELDS
)
UPLOAD_ERROR = "Choose an unencrypted .xlsx within the workbook size limit."


@dataclass
class Job:
    owner: str
    directory: TemporaryDirectory
    names: list[str]
    created: float
    state: str = "queued"
    updated: float = field(init=False)
    processed: int = 0
    message: str = ""
    result: dict | None = None

    def __post_init__(self):
        self.updated = self.created

    @property
    def source(self):
        return Path(self.directory.name) / "source.xlsx"

    @property
    def output(self):
        return Path(self.directory.name) / "report.xlsx"

    def finish(self, state, *, message="", result=None):
        self.state, self.updated = state, monotonic()
        self.message, self.result = message, result
        self.names.clear()
        self.source.unlink(missing_ok=True)


def validate_upload(path, expanded_limit):
    with ZipFile(path) as archive:
        members = archive.infolist()
        if (
            "xl/workbook.xml" not in archive.namelist()
            or sum(member.file_size for member in members) > expanded_limit
            or any(member.flag_bits & 1 for member in members)
        ):
            raise ValueError(UPLOAD_ERROR)


def report_view(path):
    status, managers, documents = read_results(path)
    rows = []
    for row in documents:
        try:
            _, url = workspace_identity(row[9])
        except ValueError:
            url = ""
        rows.append((*row[:9], url, row[10]))
    return {
        "status": status,
        "managers": managers,
        "documents": rows,
        "document_headers": DOCUMENT_HEADERS,
        "manager_headers": MANAGER_HEADERS,
    }


def create_app(settings=None):
    app = Flask(__name__)
    secret_file = os.getenv("XCCB_WORKER_SECRET_FILE")
    config_file = os.getenv("XCCB_REPORT_CONFIG")
    app.config.from_mapping(
        SECRET_KEY=secrets.token_bytes(32),
        SESSION_COOKIE_NAME="__Host-xccb",
        SESSION_COOKIE_SECURE=True,
        SESSION_COOKIE_SAMESITE="Strict",
        MAX_CONTENT_LENGTH=int(os.getenv("XCCB_MAX_UPLOAD_MIB", "25")) * 1024**2,
        MAX_REPORT_BYTES=int(os.getenv("XCCB_MAX_REPORT_MIB", "50")) * 1024**2,
        MAX_EXPANDED_BYTES=int(os.getenv("XCCB_MAX_EXPANDED_MIB", "250")) * 1024**2,
        MAX_JOBS=int(os.getenv("XCCB_MAX_JOBS", "12")),
        JOB_SECONDS=int(os.getenv("XCCB_JOB_TIMEOUT_SECONDS", "3600")),
        REPORTS_ENABLED=os.getenv("XCCB_ENABLE_REPORTS") == "1",
        WORKER_SECRET=Path(secret_file).read_text(encoding="utf-8-sig").strip()
        if secret_file
        else "",
        REPORT_CONFIG=load_config(config_file, use_preferences=False)
        if config_file
        else DEFAULTS,
    )
    app.config.update(settings or {})
    if app.config["REPORTS_ENABLED"] and len(app.config["WORKER_SECRET"]) < 32:
        raise ValueError("Configure a worker secret of at least 32 characters.")
    if any(
        app.config[key] < 1
        for key in (
            "MAX_CONTENT_LENGTH",
            "MAX_REPORT_BYTES",
            "MAX_EXPANDED_BYTES",
            "MAX_JOBS",
            "JOB_SECONDS",
        )
    ):
        raise ValueError("Web request and job limits must be positive.")
    jobs = {}
    lock = Lock()
    worker_seen = None

    def worker_online():
        return worker_seen is not None and monotonic() - worker_seen < LEASE_SECONDS

    def reap():
        now = monotonic()
        for key, job in list(jobs.items()):
            if job.state in ACTIVE_STATES:
                if now - job.created >= app.config["JOB_SECONDS"]:
                    job.finish(
                        "failed", message="The report timed out. Submit it again."
                    )
                elif job.state == "running" and now - job.updated >= LEASE_SECONDS:
                    job.finish(
                        "failed",
                        message="The Windows worker disconnected. Submit the workbook again.",
                    )
            elif now - job.updated >= RESULT_SECONDS:
                job.directory.cleanup()
                del jobs[key]

    def sweep():
        while True:
            sleep(15)
            try:
                with lock:
                    reap()
            except OSError:
                app.logger.error("Could not clean up expired report files")

    if not app.testing:
        Thread(target=sweep, name="report-expiry", daemon=True).start()

    @app.before_request
    def authorize():
        nonlocal worker_seen
        if request.path.startswith("/worker/"):
            auth = request.headers.get("Authorization", "")
            secret = app.config["WORKER_SECRET"]
            if (
                not secret
                or not auth.startswith("Bearer ")
                or not compare_digest(
                    auth[7:].encode(),
                    secret.encode(),
                )
            ):
                abort(401)
            with lock:
                worker_seen = monotonic()
        elif (
            request.method == "POST" and request.headers.get("X-Report-Request") != "1"
        ):
            abort(403)

    @app.after_request
    def protect(response):
        response.headers["Cache-Control"] = "no-store"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'"
        )
        return response

    @app.errorhandler(HTTPException)
    def error(exc):
        return jsonify(error=exc.description), exc.code

    @app.get("/")
    def index():
        session.setdefault("owner", secrets.token_urlsafe(32))
        return render_template(
            "report.html",
            defaults=app.config["REPORT_CONFIG"]["reviewer_names"],
        )

    @app.get("/healthz")
    def health():
        return {"status": "ok"}

    @app.get("/api/status")
    def status():
        with lock:
            reap()
            return {
                "enabled": app.config["REPORTS_ENABLED"],
                "worker_online": worker_online(),
                "active_jobs": sum(job.state in ACTIVE_STATES for job in jobs.values()),
            }

    @app.post("/api/jobs")
    def submit():
        owner = session.get("owner")
        if owner is None:
            abort(403)
        if not app.config["REPORTS_ENABLED"]:
            abort(503, "Live reporting is disabled.")
        upload = request.files.get("workbook")
        if upload is None or not (upload.filename or "").lower().endswith(".xlsx"):
            abort(400, "Choose a Forum Reviewers .xlsx workbook.")
        directory = TemporaryDirectory(prefix="xccb-web-")
        job = Job(
            owner, directory, parse_names(request.form.get("names", "")), monotonic()
        )
        retained = False
        try:
            upload.save(job.source)
            try:
                validate_upload(job.source, app.config["MAX_EXPANDED_BYTES"])
            except (BadZipFile, ValueError):
                abort(400, UPLOAD_ERROR)
            with lock:
                reap()
                if not worker_online():
                    abort(503, "The Windows worker is offline.")
                previous = next(
                    (key for key, item in jobs.items() if item.owner == owner), None
                )
                if previous is not None:
                    item = jobs[previous]
                    if item.state in ACTIVE_STATES:
                        abort(
                            409,
                            "A report is already in progress in this browser session.",
                        )
                    item.directory.cleanup()
                    del jobs[previous]
                if len(jobs) >= app.config["MAX_JOBS"]:
                    abort(429, "The report service is busy. Try again later.")
                key = secrets.token_hex(16)
                jobs[key] = job
                retained = True
            return {"id": key}, 202
        finally:
            if not retained:
                directory.cleanup()

    def owned_job(key):
        reap()
        job = jobs.get(key)
        if job is None or job.owner != session.get("owner"):
            abort(404, "This report is unavailable or has expired.")
        return job

    @app.get("/api/jobs/<key>")
    def job_status(key):
        with lock:
            job = owned_job(key)
            return {
                "state": job.state,
                "processed": job.processed,
                "message": job.message,
                "result": job.result,
            }

    @app.get("/api/jobs/<key>/excel")
    def excel(key):
        with lock:
            job = owned_job(key)
            if job.state != "complete":
                abort(409, "The Excel report is not ready.")
            return send_file(
                job.output,
                as_attachment=True,
                download_name="Flawless_Submission.xlsx",
                conditional=False,
                etag=False,
            )

    def running_job(key):
        reap()
        job = jobs.get(key)
        if job is None or job.state != "running":
            abort(404, "The job is no longer running.")
        return job

    @app.post("/worker/claim")
    def claim():
        with lock:
            reap()
            if app.config["REPORTS_ENABLED"] and not any(
                job.state == "running" for job in jobs.values()
            ):
                for key, job in jobs.items():
                    if job.state == "queued":
                        job.state, job.updated = "running", monotonic()
                        return {
                            "id": key,
                            "names": list(job.names),
                            "config": app.config["REPORT_CONFIG"],
                            "timeout": app.config["JOB_SECONDS"]
                            - (job.updated - job.created),
                        }
            return "", 204

    @app.get("/worker/jobs/<key>/source")
    def source(key):
        with lock:
            job = running_job(key)
            return send_file(job.source, conditional=False, etag=False)

    @app.post("/worker/jobs/<key>/progress")
    def progress(key):
        payload = request.get_json()
        processed = payload.get("processed") if isinstance(payload, dict) else None
        if type(processed) is not int or processed < 0:
            abort(400, "Expected a nonnegative processed row count.")
        with lock:
            job = running_job(key)
            job.updated, job.processed = monotonic(), processed
        return "", 204

    @app.post("/worker/jobs/<key>/failed")
    def failed(key):
        with lock:
            running_job(key).finish(
                "failed",
                message="The worker could not finish the report. Check the workbook or contact the report operator.",
            )
        return "", 204

    @app.post("/worker/jobs/<key>/result")
    def complete(key):
        with lock:
            running_job(key)
        request.max_content_length = app.config["MAX_REPORT_BYTES"]
        output = request.files.get("report")
        if output is None:
            abort(400, "Expected the generated report.")
        with TemporaryDirectory(prefix="xccb-result-") as directory:
            staged = Path(directory) / "report.xlsx"
            output.save(staged)
            try:
                validate_upload(staged, app.config["MAX_EXPANDED_BYTES"])
                result = report_view(staged)
            except (BadZipFile, ValueError, KeyError, OSError, SyntaxError):
                with lock:
                    running_job(key).finish(
                        "failed", message="The worker returned an invalid report."
                    )
                abort(422, "The worker returned an invalid report.")
            with lock:
                job = running_job(key)
                staged.replace(job.output)
                job.finish("complete", result=result)
        return "", 204

    return app
