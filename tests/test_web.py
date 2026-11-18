import json
from datetime import datetime
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook

from xccb_report import web, web_worker
from xccb_report.config import DEFAULTS
from xccb_report.organizations import OrganizationCache
from xccb_report.pipeline import report
from xccb_report.results import read_results

TOKEN = "test-token" * 4
WORKER = {"Authorization": f"Bearer {TOKEN}"}
BROWSER = {"X-Report-Request": "1"}
URL = "https://xccb2.intel.com/Documents/DocumentWorkspace.aspx?DocumentID=123"


@pytest.fixture
def workbook(tmp_path):
    path = tmp_path / "source.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append(["Forum Reviewers"])
    sheet.append([])
    sheet.append(
        [
            "Author Manager",
            "Reviewing Forum",
            "Document WS URL",
            "Reviewer(s)",
            "Phase",
            "Document Submitted",
        ]
    )
    sheet.append(
        ["Sample Manager", "Admin", "Open", "Reviewer", "Final", datetime(2025, 5, 1)]
    )
    sheet["C4"].hyperlink = URL
    book.save(path)
    book.close()
    return path


@pytest.fixture
def app():
    return web.create_app(
        {
            "TESTING": True,
            "REPORTS_ENABLED": True,
            "WORKER_SECRET": TOKEN,
            "REPORT_CONFIG": DEFAULTS,
            "MAX_JOBS": 3,
        }
    )


def submit(browser, workbook):
    with workbook.open("rb") as upload:
        return browser.post(
            "/api/jobs",
            headers=BROWSER,
            data={
                "workbook": (upload, "reviewers.xlsx"),
                "names": " Reviewer\nreviewer ",
            },
        )


def test_web_report_matches_pipeline_and_download_is_private(
    app, workbook, tmp_path, monkeypatch
):
    class Reader:
        def __init__(self, **kwargs):
            pass

        def get(self, url):
            if url == URL:
                return (
                    '<div id="View_DocumentTitleCell">'
                    '<a onclick="ViewOnlineDocument()">Paper.docx</a></div>'
                    '<iframe id="iFrameActions" '
                    'src="DocumentComponents.aspx?Component=Actions&amp;DocumentID=123"></iframe>'
                )
            return (
                '<div>Actions</div><table class="action-panel"><tr><td>'
                '<a href="/Actions/CompleteDocumentAction.aspx?ActionID=1&amp;DocumentID=123">Action</a>'
                "<span>Assigned By: <b>Reviewer</b></span></td></tr></table>"
            )

    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "desktop-user"))
    monkeypatch.setattr(web_worker, "WindowsReader", Reader)
    monkeypatch.setattr(
        "xccb_report.organizations.query_directory_batch",
        lambda names, python: {
            "sample manager": [
                {
                    "name": "Sample Manager",
                    "wwid": "12345",
                    "organization_code": "0012",
                    "organization": "Sample Org",
                }
            ],
        },
    )
    browser, other, worker = (app.test_client() for _ in range(3))
    assert browser.get("/").status_code == 200
    assert other.get("/").status_code == 200
    assert worker.post("/worker/claim", headers=WORKER).status_code == 204
    uploaded = submit(browser, workbook)
    assert uploaded.status_code == 202
    key = uploaded.json["id"]
    url, route = f"/api/jobs/{key}", f"/worker/jobs/{key}"
    assert other.get(url).status_code == other.get(url + "/excel").status_code == 404
    assert worker.get(route + "/source").status_code == 401
    assert submit(browser, workbook).status_code == 409
    claim = worker.post("/worker/claim", headers=WORKER).json
    assert claim["names"] == ["Reviewer"]
    assert worker.post("/worker/claim", headers=WORKER).status_code == 204

    directory = tmp_path / "worker"
    directory.mkdir()
    source = worker.get(route + "/source", headers=WORKER)
    (directory / "source.xlsx").write_bytes(source.data)
    source.close()
    (directory / "job.json").write_text(json.dumps(claim), encoding="utf-8")
    web_worker.generate(directory)
    assert (directory / "organizations.txt").is_file()
    assert not (tmp_path / "desktop-user").exists()
    progress = json.loads((directory / "progress.json").read_text(encoding="utf-8"))
    assert (
        worker.post(route + "/progress", headers=WORKER, json=progress).status_code
        == 204
    )
    assert browser.get(url).json["processed"] == 1
    with (directory / "report.xlsx").open("rb") as output:
        result = worker.post(route + "/result", headers=WORKER, data={"report": output})
    assert result.status_code == 204
    view = browser.get(url).json["result"]
    assert view["documents"][0][:7] == [
        "No",
        "123",
        "Sample Manager",
        "Sample Org",
        "Final",
        "Paper.docx",
        "2025-05-01",
    ]
    assert view["documents"][0][9] == URL
    assert view["managers"][0][:2] == ["Sample Manager", "Sample Org"]
    assert other.get(url + "/excel").status_code == 404
    with browser.get(url + "/excel") as download:
        assert download.status_code == 200
        assert download.headers["Cache-Control"] == "no-store"
        contents = download.data
    with BytesIO(contents) as stream:
        excel = load_workbook(stream)
        assert excel["Audit"]["P2"].value == "12345"
        assert excel["Audit"]["I2"].value == "123.json"
        summary = dict(excel["Summary"].values)
        assert summary["Source workbook"] == "source.xlsx"
        excel.close()
    comparison = tmp_path / "comparison.xlsx"
    report(
        workbook,
        comparison,
        DEFAULTS
        | {
            "reviewer_names": ["Reviewer"],
            "assigned_by_names": ["Reviewer"],
            "action_scope": "document",
        },
        tmp_path / "comparison-evidence",
        reader=Reader(),
        all_dates=True,
        organization_cache=OrganizationCache(tmp_path / "comparison-organizations.txt"),
    )
    assert read_results(BytesIO(contents)) == read_results(comparison)


def test_web_validates_upload_and_requires_live_worker(app, workbook):
    browser = app.test_client()
    assert browser.get("/").status_code == 200
    assert browser.post("/api/jobs").status_code == 403
    assert submit(browser, workbook).status_code == 503
    assert browser.post("/worker/claim").status_code == 401
    assert browser.post("/worker/claim", headers=WORKER).status_code == 204
    assert (
        browser.post(
            "/api/jobs",
            headers=BROWSER,
            data={"workbook": (BytesIO(b"bad"), "invalid.xlsx")},
        ).status_code
        == 400
    )
    app.config["MAX_EXPANDED_BYTES"] = 1
    assert submit(browser, workbook).status_code == 400
    app.config["MAX_CONTENT_LENGTH"] = 1024
    assert submit(browser, workbook).status_code == 413
    app.config["REPORTS_ENABLED"] = False
    assert submit(browser, workbook).status_code == 503


def test_worker_loss_fails_job_and_reaps_artifacts(app, workbook, monkeypatch):
    now = 1.0
    monkeypatch.setattr(web, "monotonic", lambda: now)
    browser, worker = app.test_client(), app.test_client()
    browser.get("/")
    worker.post("/worker/claim", headers=WORKER)
    key = submit(browser, workbook).json["id"]
    assert worker.post("/worker/claim", headers=WORKER).json["id"] == key
    now += web.LEASE_SECONDS
    url = f"/api/jobs/{key}"
    failed = browser.get(url).json
    assert failed["state"] == "failed" and "disconnected" in failed["message"]
    assert worker.get(f"/worker/jobs/{key}/source", headers=WORKER).status_code == 404
    now += web.RESULT_SECONDS
    assert browser.get(url).status_code == 404
    assert browser.get("/api/status").json["active_jobs"] == 0
