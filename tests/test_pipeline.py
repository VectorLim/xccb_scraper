from datetime import datetime, timezone
from pathlib import Path
from copy import deepcopy
import json

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from xccb_report.evidence import inspect_saved, load_evidence, write_json
from xccb_report.__main__ import main
from xccb_report.site import capture
from xccb_report.config import DEFAULTS, load_config
from xccb_report.pipeline import report

URL = "https://xccb2.intel.com/Documents/DocumentWorkspace.aspx?DocumentID=123"
FRAME = "https://xccb2.intel.com/Documents/DocumentComponents.aspx?Component=Actions&DocumentID=123"
HTML = '<body onclick="secret"><form><input name="__VIEWSTATE" value="secret"><!--secret--><div>Actions</div><table class="action-panel"><tr><td><a href="/Actions/CompleteDocumentAction.aspx?ActionID=1&amp;DocumentID=123">Action</a><span>Assigned By: <b>Person, One</b></span></td></tr></table></form><script>secret</script></body>'


@pytest.fixture(autouse=True)
def offline_organizations(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "private_user"))
    monkeypatch.setattr(
        "xccb_report.organizations.query_directory_batch", lambda *args: {}
    )


def config():
    return deepcopy(DEFAULTS) | {
        "assigned_by_names": ["Person, One"],
        "reviewer_names": ["Reviewer, One"],
        "action_scope": "document",
    }


def source(path):
    book = Workbook()
    sheet = book.active
    headers = [
        "Author Manager",
        "Reviewing Forum",
        "Reviewer(s)",
        "Is FYI?",
        "Phase",
        "Document Submitted",
        "Document WS URL",
        "Version",
    ] + [f"Field {i}" for i in range(35)]
    sheet.append(["Title"])
    sheet.append([])
    sheet.append(headers)
    sheet.cell(3, 1).fill = PatternFill("solid", fgColor="123456")
    sheet.column_dimensions["A"].fill = PatternFill("solid", fgColor="ABCDEF")
    for phase in ["Final ", "Preliminary"]:
        sheet.append(
            [
                "Manager, One",
                "Admin Review",
                "Reviewer, One",
                "No",
                phase,
                datetime(2026, 10, 1),
                "Browse",
                5,
            ]
        )
        sheet.cell(sheet.max_row, 7).hyperlink = URL
        sheet.cell(
            sheet.max_row,
            9,
            '=HYPERLINK("https://example.invalid", "literal source text")',
        ).data_type = "s"
    book.save(path)


def test_report_preserves_links_and_caches_document(tmp_path, monkeypatch):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    calls = []

    def evidence(*args):
        calls.append(args[1])
        return [{"assigned_by": "Person, One"}], True, "Validated", "local-evidence"

    monkeypatch.setattr("xccb_report.pipeline.load_evidence", evidence)
    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
    )
    assert calls == ["123"]
    assert counts["No"] == counts["Duplicate"] == 1
    book = load_workbook(output)
    tab = book["Flawless submissions"]
    assert tab.max_column == 46 and tab.max_row == 2
    assert tab["A1"].fill.fgColor.rgb == "00123456"
    assert tab.column_dimensions["A"].fill.fgColor.rgb == "00ABCDEF"
    assert tab["I2"].hyperlink.target == URL
    assert tab["I2"].hyperlink.ref == "I2"
    assert tab["AT2"].value == "No"
    assert tab["K2"].data_type == "s" and tab["K2"].value.startswith("=HYPERLINK")
    assert tab["G1"].value == "Document" and tab["G2"].value is None
    assert book["Audit"].max_row == 3
    assert tab.auto_filter.ref == "A1:AT2"
    assert load_workbook(original).active.max_row == 5
    with pytest.raises(ValueError, match="never overwritten"):
        report(original, output, config(), tmp_path)


@pytest.mark.parametrize(
    "workspace_titles, expected",
    [
        (["Whitepaper & document.docm"], "Whitepaper & document.docm"),
        (["=A title, not a formula"], "=A title, not a formula"),
        ([], ""),
        (["First document.docx", "Second document.docx"], ""),
    ],
)
def test_document_title_is_between_phase_and_submitted_in_export_and_dashboard(
    tmp_path, workspace_titles, expected
):
    from html import escape
    from xccb_report.gui import DOCUMENT_COLUMNS
    from xccb_report.results import read_results

    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    book.active["AQ3"] = "CCB Title"
    book.active["AQ4"] = "Incorrect CCB title"
    book.active["AQ5"] = "Another incorrect CCB title"
    book.save(original)
    book.close()

    class Reader:
        def __init__(self):
            self.urls = []

        def get(self, url):
            self.urls.append(url)
            if url == FRAME:
                return HTML
            return (
                '<div id="ctl00_PlaceHolderContentDescription_ctl01_DocumentTitleCell">'
                + "".join(
                    "<a onclick=\"return ViewOnlineDocument('https://example.invalid/paper.docm')\">"
                    + escape(title)
                    + "</a>"
                    for title in workspace_titles
                )
                + '</div><a id="ctl00_PlaceHolderToolbar_ctl00_tb_view" onclick="return ViewOnlineDocument(\'https://example.invalid/paper.docm\')">V iew</a>'
                + '<iframe id="iFrameActions" src="DocumentComponents.aspx?Component=Actions&amp;DocumentID=123"></iframe>'
            )

    reader, progress = Reader(), []
    report(
        original,
        output,
        config(),
        tmp_path,
        reader=reader,
        reporting_month="2026-10",
        on_progress=lambda *args: progress.append(args),
    )
    assert reader.urls == [URL, FRAME]
    assert all(entry[2]["title"] == expected for entry in progress if entry[2] is not None)
    book = load_workbook(output)
    sheet = book["Flawless submissions"]
    headers = [cell.value for cell in sheet[1]]
    phase = headers.index("Phase")
    assert headers[phase : phase + 3] == ["Phase", "Document", "Document Submitted"]
    assert "CCB Title" not in headers
    assert (sheet.cell(2, phase + 2).value or "") == expected
    assert sheet.max_row == 2
    if expected:
        assert sheet.cell(2, phase + 2).data_type == "s"
    else:
        assert "Document title unavailable" in book["Audit"]["H2"].value
    book.close()
    _, _, documents = read_results(output)
    columns = [column[1] for column in DOCUMENT_COLUMNS]
    phase = columns.index("Phase")
    assert columns[phase : phase + 3] == ["Phase", "Document", "Submitted"]
    assert documents[0][phase + 1] == expected
    assert len(documents) == 1
    offline_output = tmp_path / "offline.xlsx"
    report(original, offline_output, config(), tmp_path, reporting_month="2026-10")
    assert read_results(offline_output)[2] == documents
    book = load_workbook(original)
    assert book.active["AQ3"].value == "CCB Title"
    book.close()


def test_empty_reviewer_list_includes_only_blank_reviewers(tmp_path):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    source_book = load_workbook(original)
    source_book.active["C4"] = None
    source_book.save(original)
    source_book.close()
    counts = report(
        original,
        output,
        deepcopy(DEFAULTS),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
    )
    assert counts["Qualified"] == 1 and counts["Excluded"] == 1
    assert counts["Unresolved Actions"] == 1
    book = load_workbook(output)
    assert book["Flawless submissions"].max_row == 2
    assert book["Summary"]["B2"].value.startswith("DRAFT")
    assert book["By Author Manager"].max_row == 2
    book.close()


def test_reporting_month_does_not_change_evidence_time(tmp_path, monkeypatch):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    now = datetime(2026, 11, 2, tzinfo=timezone.utc)
    evidence_times = []

    def evidence(directory, document_id, settings, checked_at, details=None):
        evidence_times.append(checked_at)
        return [], True, "Validated", "local-evidence"

    monkeypatch.setattr("xccb_report.pipeline.load_evidence", evidence)
    counts = report(
        original, output, config(), tmp_path, now=now, reporting_month="2026-10"
    )
    assert counts["Qualified"] == counts["Yes"] == 1
    assert evidence_times == [now]
    book = load_workbook(output)
    try:
        summary = dict(book["Summary"].values)
        assert summary["Month start inclusive"] == "2026-10-01"
        assert summary["Next month exclusive"] == "2026-11-01"
        assert summary["Generated UTC"] == now.isoformat()
    finally:
        book.close()
    current_counts = report(
        original, tmp_path / "current.xlsx", config(), tmp_path, now=now
    )
    assert current_counts.get("Qualified", 0) == 0
    assert evidence_times == [now]


@pytest.mark.parametrize("month", ["", "2026-1", "2026-13", "not-a-month", "0000-01"])
def test_report_rejects_invalid_month_before_opening_source(tmp_path, month):
    with pytest.raises(ValueError):
        report(
            tmp_path / "missing.xlsx",
            tmp_path / "output.xlsx",
            config(),
            tmp_path,
            reporting_month=month,
        )
    assert not (tmp_path / "output.xlsx").exists()


def test_report_preserves_existing_workbook_at_old_temporary_path(tmp_path):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    previous_temp = tmp_path / "report.tmp.xlsx"
    previous_temp.write_bytes(original.read_bytes())
    expected = previous_temp.read_bytes()
    report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
    )
    assert output.exists()
    assert previous_temp.read_bytes() == expected
    assert not list(tmp_path.glob(".report_*.tmp.xlsx"))


def test_live_capture_uses_observed_frame_and_sanitizes(tmp_path):
    class Reader:
        def __init__(self):
            self.urls = []

        def get(self, url):
            self.urls.append(url)
            return (
                '<div id="ctl00_PlaceHolderContentDescription_ctl01_DocumentTitleCell">'
                "Different CCB heading [Class IV]"
                "<a onclick=\"return ViewOnlineDocument('https://example.invalid/paper.docm')\">"
                "Whitepaper &amp; document.docm</a></div>"
                '<a id="ctl00_PlaceHolderToolbar_ctl00_tb_view" onclick="return ViewOnlineDocument(\'https://example.invalid/paper.docm\')">V iew</a>'
                '<iframe id="iFrameActions" src="DocumentComponents.aspx?Component=Actions&amp;DocumentID=123"></iframe>'
                if len(self.urls) == 1
                else HTML
            )

    reader = Reader()
    result = capture(reader, URL, tmp_path)
    assert reader.urls == [URL, FRAME]
    assert result["document_title"] == "Whitepaper & document.docm"
    assert result["actions"] == [{"action_id": "1", "assigned_by": "Person, One"}]
    assert (
        result["complete"] is (not result["parse_errors"])
        and result["scope"] == "document"
    )
    clean = Path(result["source"]).read_text(encoding="utf-8")
    assert "secret" not in clean
    assert all(tag not in clean for tag in ("<form", "<input", "<script"))
    assert "ViewOnlineDocument" not in clean
    from datetime import timedelta
    from xccb_report.evidence import load_document_title

    now = datetime.now(timezone.utc)
    assert (
        load_document_title(tmp_path, "123", config(), now)
        == "Whitepaper & document.docm"
    )
    assert load_document_title(tmp_path, "123", config(), now + timedelta(days=2)) == ""
    Path(result["source"]).write_text(clean + "tampered", encoding="utf-8")
    assert load_document_title(tmp_path, "123", config(), now) == ""


def test_live_capture_reports_deleted_document(tmp_path):
    from xccb_report.results import short_note

    assert (
        short_note(
            "Actions completeness not verified; Live capture failed: Document deleted on xCCB (x)"
        )
        == "Document deleted on xCCB"
    )

    class Reader:
        def get(self, url):
            return (
                "<div>Document not found</div><div>You have attempted to access a document that was deleted by "
                "Doe, Jane - KM (1) on 10/27/2026 1:26 AM.</div>"
            )

    with pytest.raises(
        ValueError,
        match=r"Document deleted on xCCB \(deleted by Doe, Jane - KM \(1\) on 10/27/2026 1:26 AM\)",
    ):
        capture(Reader(), URL, tmp_path)


@pytest.mark.parametrize("live", [False, True])
def test_cli_runs_report_with_conventional_paths(tmp_path, monkeypatch, live):
    (tmp_path / "config.json").write_text("{}", encoding="utf-8")
    monkeypatch.chdir(tmp_path)
    reader, calls = object(), []
    monkeypatch.setattr("xccb_report.__main__.WindowsReader", lambda **kwargs: reader)
    monkeypatch.setattr(
        "xccb_report.__main__.report",
        lambda *args, **kwargs: calls.append((args, kwargs)) or {},
    )
    monkeypatch.setattr(
        "sys.argv",
        ["xccb_report", "source.xlsx", "report.xlsx"] + (["--live"] if live else []),
    )
    assert main() == 0
    assert calls == [
        (
            ("source.xlsx", "report.xlsx", DEFAULTS, "outputs/evidence"),
            {"reader": reader if live else None},
        )
    ]


def test_evidence_needs_attestation_and_hash(tmp_path):
    html = tmp_path / "123.html"
    html.write_text(HTML, encoding="utf-8")
    value = inspect_saved(html, URL)
    manifest = tmp_path / "123.json"
    write_json(manifest, value)
    now = datetime.now(timezone.utc)
    assert load_evidence(tmp_path, "123", config(), now)[0] == []
    value.update(
        complete=True,
        scope="document",
        validated_by="Test verifier",
        validation_notes="Synthetic fixture checked all records",
    )
    write_json(manifest, value)
    assert load_evidence(tmp_path, "123", config(), now)[1] is True
    html.write_text(HTML + "tampered", encoding="utf-8")
    assert "hash mismatch" in load_evidence(tmp_path, "123", config(), now)[2]


def test_verified_empty_state(tmp_path):
    html = tmp_path / "123.html"
    html.write_text("<div>No actions found.</div>", encoding="utf-8")
    value = inspect_saved(html, URL)
    value.update(
        complete=True,
        scope="document",
        validated_by="Test verifier",
        validation_notes="Synthetic live empty state checked",
        empty_state_text="No actions found.",
        empty_state_verified=True,
    )
    write_json(tmp_path / "123.json", value)
    actions, complete, *_ = load_evidence(
        tmp_path, "123", config(), datetime.now(timezone.utc)
    )
    assert actions == [] and complete is True


def test_config_ignores_obsolete_filter_settings(tmp_path, caplog):
    path = tmp_path / "config.json"
    path.write_text(
        json.dumps({"manager_organizations": {"Name": "KuAT", " NAME ": "Elsewhere"}}),
        encoding="utf-8",
    )
    assert load_config(path) == DEFAULTS
    assert "Ignoring obsolete filter setting" in caplog.text


def test_legacy_phase_allowlist_is_ignored(tmp_path, caplog):
    path = tmp_path / "config.json"
    path.write_text(
        json.dumps({"phases": ["Final", "Preliminary", "WLA CCB Final"]}),
        encoding="utf-8",
    )
    assert load_config(path) == DEFAULTS
    assert "Ignoring obsolete filter setting: phases" in caplog.text


@pytest.mark.parametrize("excluded_phase", ["wLa CCB Final", "LSI Modification"])
def test_report_applies_amended_filters(tmp_path, excluded_phase):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    book.active["C4"] = "Reviewers: Reviewer, One (Admin)"
    book.active["D4"] = "Yes"
    book.active["E4"] = "Do Not Implement"
    book.active["E5"] = excluded_phase
    book.save(original)
    book.close()
    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
    )
    assert counts["Qualified"] == 1 and counts["Excluded"] == 1
    book = load_workbook(output)
    assert book["Flawless submissions"]["F2"].value == "Do Not Implement"
    assert book["Flawless submissions"].max_row == 2
    summary = dict(book["Summary"].values)
    assert (
        summary["Phase rule"]
        == "All phases except values containing WLA or LSI modification (case-insensitive)"
    )
    assert (
        summary["Reviewer rule"]
        == "Reviewer(s) is blank or contains a listed reviewer name (case-insensitive)"
    )
    assert "Phase contains" in book["Audit"]["H3"].value
    book.close()


def test_report_groups_managers_without_excluding_missing_or_unmapped_names(tmp_path):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    sheet = book.active
    sheet["A4"] = "Z Manager"
    sheet["A5"] = "A Manager"
    sheet["G5"].hyperlink = URL.replace("123", "456")
    sheet.append([c.value for c in sheet[4]])
    sheet["A6"] = None
    sheet["G6"].hyperlink = URL.replace("123", "789")
    book.save(original)
    book.close()
    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
    )
    assert counts["Qualified"] == 3
    book = load_workbook(output)
    assert [
        row[0]
        for row in book["Flawless submissions"].iter_rows(min_row=2, values_only=True)
    ] == ["A Manager", "Z Manager", None]
    groups = list(book["By Author Manager"].iter_rows(min_row=2, values_only=True))
    assert groups == [
        ("A Manager", "Unresolved", 1, 0, 0, 1),
        ("Z Manager", "Unresolved", 1, 0, 0, 1),
        ("(Not specified)", "Unresolved", 1, 0, 0, 1),
    ]
    assert sum(row[2] for row in groups) == counts["Qualified"]
    book.close()


def test_live_capture_is_not_future_dated_against_report_start(tmp_path, monkeypatch):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    from xccb_report.filtering import Eligibility

    monkeypatch.setattr(
        "xccb_report.pipeline.eligibility", lambda *_: Eligibility("Qualified")
    )

    def fresh_capture(reader, url, directory):
        html = tmp_path / "123.html"
        html.write_text(HTML, encoding="utf-8")
        value = inspect_saved(html, url)
        value.update(
            captured_at=datetime.now(timezone.utc).isoformat(),
            scope="document",
            complete=True,
            validated_by="Test verifier",
            validation_notes="Synthetic full listing",
        )
        write_json(tmp_path / "123.json", value)

    monkeypatch.setattr("xccb_report.site.capture", fresh_capture)
    counts = report(original, output, config(), tmp_path, reader=object())
    assert counts["No"] == 1
    assert "Unresolved Actions" not in counts


def test_live_failure_does_not_reuse_previous_success(tmp_path, monkeypatch):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)

    def fail(*_):
        raise OSError("HTTP 403")

    monkeypatch.setattr("xccb_report.site.capture", fail)
    monkeypatch.setattr(
        "xccb_report.pipeline.load_evidence",
        lambda *_: pytest.fail("Stale success reused"),
    )
    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        reader=object(),
    )
    assert counts["Unresolved Actions"] == 1
    assert load_workbook(output)["Flawless submissions"]["AS2"].value is None


def test_report_streams_completed_documents_concurrently_in_source_order(
    tmp_path, monkeypatch
):
    from threading import Barrier, Event

    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    sheet = book.active
    sheet["A4"] = "A Slow Manager"
    sheet["A5"] = "Z Fast Manager"
    fast_url = URL.replace("123", "456")
    sheet["G5"].hyperlink = fast_url
    sheet.append([cell.value for cell in sheet[5]])
    sheet["G6"].hyperlink = fast_url
    book.save(original)
    book.close()
    overlap, release_slow = Barrier(2), Event()
    captures, updates = [], []

    def capture_document(reader, url, directory):
        captures.append(url)
        overlap.wait(timeout=5)
        if url == URL:
            assert release_slow.wait(
                timeout=5
            ), "Fast progress was blocked by the slow document"

    def progress(counts, group, document):
        assert not output.exists()
        if document is None:
            return
        assert document["result"] == "No" and document["matches"] == "Person, One"
        updates.append((counts, group))
        if group[0] == "Z Fast Manager":
            release_slow.set()

    monkeypatch.setattr("xccb_report.site.capture", capture_document)
    monkeypatch.setattr(
        "xccb_report.pipeline.load_evidence",
        lambda *_: (
            [{"assigned_by": "Person, One"}],
            True,
            "Validated",
            "local-evidence",
        ),
    )
    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        reader=object(),
        on_progress=progress,
    )
    assert sorted(captures) == sorted([URL, fast_url])
    assert updates[0] == (
        {"Duplicate": 1, "No": 1, "Qualified": 1},
        ("Z Fast Manager", "Unresolved", 1, 0, 1, 0),
    )
    assert updates[1][1] == ("A Slow Manager", "Unresolved", 1, 0, 1, 0)
    assert updates[-1][0] == counts
    assert counts["Qualified"] == counts["No"] == 2
    book = load_workbook(output)
    try:
        assert [
            row[0]
            for row in book["Flawless submissions"].iter_rows(
                min_row=2, values_only=True
            )
        ] == ["A Slow Manager", "Z Fast Manager"]
        assert [
            row[0] for row in book["Audit"].iter_rows(min_row=2, values_only=True)
        ] == [4, 5, 6]
        assert list(book["By Author Manager"].values)[1:] == [
            ("A Slow Manager", "Unresolved", 1, 0, 1, 0),
            ("Z Fast Manager", "Unresolved", 1, 0, 1, 0),
        ]
    finally:
        book.close()


@pytest.mark.parametrize(
    "state",
    [
        "fresh",
        "expired",
        "tampered",
        "incomplete",
        "future",
        "scope",
        "identity",
        "missing",
        "force",
    ],
)
def test_live_evidence_cache_only_reuses_fresh_valid_complete_documents(
    tmp_path, state
):
    from datetime import timedelta

    original = tmp_path / "source.xlsx"
    source(original)
    calls = []

    class Reader:
        def get(self, url):
            calls.append(url)
            return (
                f'<iframe id="iFrameActions" src="{FRAME}"></iframe>'
                if url == URL
                else HTML
            )

    report(
        original,
        tmp_path / "first.xlsx",
        config(),
        tmp_path,
        reader=Reader(),
        all_dates=True,
    )
    assert len(calls) == 2
    manifest = tmp_path / "123.json"
    value = json.loads(manifest.read_text(encoding="utf-8"))
    if state == "expired":
        value["captured_at"] = (
            datetime.now(timezone.utc) - timedelta(minutes=16)
        ).isoformat()
    elif state == "future":
        value["captured_at"] = (
            datetime.now(timezone.utc) + timedelta(minutes=1)
        ).isoformat()
    elif state == "tampered":
        (tmp_path / "123.html").write_text("changed", encoding="utf-8")
    elif state == "incomplete":
        value["complete"] = False
    elif state == "scope":
        value["scope"] = "unconfirmed"
    elif state == "identity":
        value["document_id"] = "456"
    write_json(manifest, value)
    if state == "missing":
        manifest.unlink()
    calls.clear()
    settings = config() | {"reuse_evidence": state != "force", "live_cache_minutes": 15}
    counts = report(
        original,
        tmp_path / "second.xlsx",
        settings,
        tmp_path,
        reader=Reader(),
        all_dates=True,
    )
    assert counts["No"] == 1
    assert len(calls) == (0 if state == "fresh" else 2)
    book = load_workbook(tmp_path / "second.xlsx")
    try:
        summary = dict(book["Summary"].values)
        assert summary["Fresh evidence cache hits"] == (1 if state == "fresh" else 0)
        assert ("fresh cache reused" in book["Audit"]["H2"].value) == (state == "fresh")
    finally:
        book.close()


def test_session_reader_reuses_per_thread_sessions_and_closes_them(monkeypatch):
    from concurrent.futures import ThreadPoolExecutor
    from contextlib import nullcontext
    from threading import Barrier
    from types import SimpleNamespace
    from xccb_report.site import WindowsReader

    sessions, options = [], []
    overlap = Barrier(2)

    class Session:
        def __init__(self):
            self.closed = False
            self.calls = 0
            sessions.append(self)

        def mount(self, prefix, adapter):
            assert prefix == "https://"
            assert adapter.context.verify_mode.name == "CERT_REQUIRED"

        def get(self, url, **kwargs):
            options.append(kwargs)
            self.calls += 1
            return nullcontext(SimpleNamespace(status_code=200, url=url, text="body"))

        def close(self):
            self.closed = True

    monkeypatch.setattr("xccb_report.site.requests.Session", Session)
    monkeypatch.setattr("requests_negotiate_sspi.HttpNegotiateAuth", lambda: object())
    monkeypatch.setattr("xccb_report.site.time.sleep", lambda _: None)
    reader = WindowsReader()

    def fetch(_):
        overlap.wait(timeout=5)
        return reader.get(URL), reader.get(FRAME)

    try:
        with ThreadPoolExecutor(max_workers=2) as executor:
            assert list(executor.map(fetch, range(2))) == [("body", "body")] * 2
        assert len(sessions) == 2 and all(session.calls == 2 for session in sessions)
        assert options == [{"allow_redirects": False, "timeout": (10, 25)}] * 4
        assert reader.metrics["requests"] == 4
    finally:
        reader.close()
    assert all(session.closed for session in sessions)


@pytest.mark.parametrize(
    "status, attempts", [(403, 1), (401, 1), (302, 1), (503, 2), (429, 2)]
)
def test_session_reader_preserves_http_error_retry_rules(monkeypatch, status, attempts):
    from xccb_report.site import WindowsReader

    reader, calls = WindowsReader(), []

    def response(url):
        calls.append(url)
        return {"status": status, "url": url, "html": ""}

    monkeypatch.setattr(reader, "_request", response)
    monkeypatch.setattr("xccb_report.site.time.sleep", lambda _: None)
    with pytest.raises(OSError, match=str(status)):
        reader.get(URL)
    assert len(calls) == attempts


@pytest.mark.parametrize(
    "url",
    [
        "http://xccb2.intel.com/",
        "https://example.com/",
        "https://xccb2.intel.com.evil.example/",
    ],
)
def test_session_reader_rejects_unapproved_hosts(url):
    from xccb_report.site import WindowsReader

    with pytest.raises(ValueError, match="approved HTTPS"):
        WindowsReader().get(url)


@pytest.mark.parametrize(
    "updates",
    [
        {"document_workers": 0},
        {"document_workers": 17},
        {"document_workers": True},
        {"request_interval_seconds": 0},
        {"request_interval_seconds": float("nan")},
        {"live_cache_minutes": -1},
        {"reuse_evidence": "yes"},
        {"http_transport": "unknown"},
    ],
)
def test_performance_settings_reject_invalid_values(updates):
    from xccb_report.config import validate_performance

    with pytest.raises(ValueError):
        validate_performance(config() | updates)


def test_windows_reader_overlaps_io_with_shared_request_spacing(monkeypatch):
    from concurrent.futures import ThreadPoolExecutor
    from threading import Barrier
    from types import SimpleNamespace
    from xccb_report.site import WindowsReader

    clock, starts = [100.0], []
    overlap = Barrier(2)

    def sleep(delay):
        clock[0] += delay

    def get_response(*args, **kwargs):
        starts.append(clock[0])
        overlap.wait(timeout=5)
        return SimpleNamespace(
            returncode=0,
            stdout=json.dumps(
                {
                    "status": 200,
                    "url": kwargs["env"]["XCCB_READ_URL"],
                    "html": "response",
                }
            ),
        )

    monkeypatch.setattr("xccb_report.site.time.monotonic", lambda: clock[0])
    monkeypatch.setattr("xccb_report.site.time.sleep", sleep)
    monkeypatch.setattr("xccb_report.site.subprocess.run", get_response)
    reader = WindowsReader(transport="powershell")
    with ThreadPoolExecutor(max_workers=2) as executor:
        assert list(executor.map(reader.get, [URL, URL.replace("123", "456")])) == [
            "response",
            "response",
        ]
    assert starts == [100.0, 101.0]


def test_windows_reader_stops_before_request_when_cancelled(monkeypatch):
    from threading import Event
    from xccb_report.site import WindowsReader

    calls = []
    monkeypatch.setattr(
        "xccb_report.site.subprocess.run", lambda *a, **k: calls.append(a)
    )
    cancel = Event()
    cancel.set()
    with pytest.raises(OSError, match="Stopped by user"):
        WindowsReader(cancel=cancel).get(URL)
    assert not calls


def test_dashboard_displays_progress_before_completion(tmp_path, monkeypatch):
    import time
    import tkinter as tk
    from threading import Event
    from xccb_report.gui import Dashboard

    monkeypatch.chdir(tmp_path)
    original = tmp_path / "source.xlsx"
    source(original)
    release = Event()

    def streaming_report(*args, on_progress, **kwargs):
        on_progress(
            {"Qualified": 1, "No": 1},
            ("Manager, One", "Unresolved", 1, 0, 1, 0),
            {
                "result": "No",
                "document": "999",
                "title": "Streamed",
                "manager": "Manager, One",
                "phase": "P",
                "submitted": None,
                "url": URL,
                "matches": "Person, One",
                "reasons": "",
            },
        )
        assert release.wait(
            timeout=10
        ), "Dashboard did not show progress before completion"
        return report(
            *args,
            **kwargs,
            now=datetime(2026, 10, 30, tzinfo=timezone.utc),
            on_progress=on_progress,
        )

    errors = []
    monkeypatch.setattr("xccb_report.gui.report", streaming_report)
    monkeypatch.setattr(
        "xccb_report.gui.messagebox.showerror",
        lambda *args, **kwargs: errors.append(args),
    )
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        assert dashboard.run_indicator.cget("text") == "Stopped"
        dashboard.source.set(str(original))
        assert not hasattr(dashboard, "month_picker")
        dashboard.reviewers.insert("1.0", "Reviewer, One\n")
        dashboard.start()
        assert dashboard.run_indicator.cget("text") == "Running"
        deadline = time.monotonic() + 5
        while not dashboard.table.get_children() and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        assert dashboard.running and not errors
        assert dashboard.card_values["Qualified"].get() == "1"
        assert dashboard.table.item(dashboard.table.get_children()[0], "values") == (
            "Manager, One",
            "Unresolved",
            "1",
            "0",
            "1",
            "0",
        )
        streamed = dashboard.documents.item(
            dashboard.documents.get_children()[0], "values"
        )
        assert (
            streamed[0] == "No"
            and streamed[1] == "999"
            and streamed[-1] == "Assigned by: Person, One"
        )
        dashboard.stop()
        assert dashboard.cancel.is_set() and dashboard.running
        assert dashboard.run_indicator.cget("text") == "Running"
        release.set()
        deadline = time.monotonic() + 10
        while dashboard.running and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        assert not dashboard.running and not errors
        assert dashboard.run_indicator.cget("text") == "Stopped"
        assert len(dashboard.table.get_children()) == 1
        assert dashboard.table.item(dashboard.table.get_children()[0], "values") == (
            "Manager, One",
            "Unresolved",
            "1",
            "0",
            "0",
            "1",
        )
    finally:
        release.set()
        root.destroy()


def test_dashboard_generates_grouped_report_in_background(tmp_path, monkeypatch):
    import time
    import tkinter as tk
    from xccb_report.gui import Dashboard

    monkeypatch.chdir(tmp_path)
    original = tmp_path / "source.xlsx"
    source(original)
    temporary_outputs = []

    def dashboard_report(source_path, output_path, *args, **kwargs):
        temporary_outputs.append(Path(output_path))
        return report(
            source_path,
            output_path,
            *args,
            **kwargs,
            now=datetime(2026, 10, 30, tzinfo=timezone.utc),
        )

    monkeypatch.setattr("xccb_report.gui.report", dashboard_report)

    class OfflineReader:
        def __init__(self, **kwargs):
            pass

        def get(self, url):
            raise OSError("offline")

    monkeypatch.setattr("xccb_report.gui.WindowsReader", OfflineReader)
    errors = []
    monkeypatch.setattr(
        "xccb_report.gui.messagebox.showerror",
        lambda *args, **kwargs: errors.append(args),
    )
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        assert not hasattr(dashboard, "settings")
        assert not hasattr(dashboard, "output")
        dashboard.source.set(str(original))
        assert dashboard.run_button.cget("text") == "Process"
        dashboard.evidence = tmp_path / "evidence"
        dashboard.reviewers.insert("1.0", "Reviewer, One\n")
        dashboard.start()
        deadline = time.monotonic() + 15
        while dashboard.running and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        assert not dashboard.running and not errors
        assert len(temporary_outputs) == 1 and not temporary_outputs[0].exists()
        assert dashboard.card_values["Qualified"].get() == "1"
        assert dashboard.status.get().startswith("DRAFT")
        assert len(dashboard.table.get_children()) == 1
        assert dashboard.table.item(dashboard.table.get_children()[0], "values") == (
            "Manager, One",
            "Unresolved",
            "1",
            "0",
            "0",
            "1",
        )
        from xccb_report.config import load_preferences

        assert load_preferences()["reviewer_names"] == ["Reviewer, One"]
        assert not (tmp_path / "reviewer_names.txt").exists()
        rows = [
            dashboard.documents.item(item, "values")
            for item in dashboard.documents.get_children()
        ]
        assert len(rows) == 1 and all(
            row[0] == "Unresolved" and row[1] == "123" and row[-2] == URL and row[-1]
            for row in rows
        )
        assert (
            dashboard.report_bytes
            and str(dashboard.export_button.cget("state")) == "normal"
        )
        dashboard.documents.selection_set(dashboard.documents.get_children()[0])
        assert dashboard.selected_url() == URL
        dashboard.result_filter.set("Yes")
        dashboard.render_documents()
        assert not dashboard.documents.get_children()
        dashboard.result_filter.set("All")
        dashboard.render_documents()
        dashboard.reviewers.delete("1.0", "end")
        dashboard.reviewers.insert("1.0", "Other Reviewer\n")
        dashboard.start()
        deadline = time.monotonic() + 15
        while dashboard.running and time.monotonic() < deadline:
            root.update()
            time.sleep(0.01)
        assert not dashboard.running and not errors
        assert (
            len(temporary_outputs) == 2 and temporary_outputs[0] != temporary_outputs[1]
        )
        assert dashboard.card_values["Qualified"].get() == "0"
        assert dashboard.status.get().startswith("No matching submissions")
        assert not dashboard.table.get_children()
        assert not dashboard.documents.get_children()
        assert all(not path.exists() for path in temporary_outputs)
        assert list(tmp_path.glob("*.xlsx")) == [original]
        assert load_preferences()["reviewer_names"] == ["Other Reviewer"]
        dashboard.source.set(str(tmp_path / "missing.xlsx"))
        dashboard.start()
        assert errors and "existing source workbook" in errors[0][1]
    finally:
        root.destroy()


def test_dashboard_column_filters(tmp_path, monkeypatch):
    import tkinter as tk
    from tkinter import ttk
    from xccb_report.gui import Dashboard
    from xccb_report.results import document_row

    monkeypatch.chdir(tmp_path)
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        rows = [
            document_row(result, str(index), "Title", manager, phase, None, URL, "", "")
            for index, (result, manager, phase) in enumerate(
                (
                    ("Yes", "Manager A", "P"),
                    ("No", "Manager A", "Q"),
                    ("", "Manager B", "P"),
                    ("Yes", "", "P"),
                )
            )
        ]
        dashboard.documents_data = rows
        dashboard.render_documents()

        def visible():
            return [
                dashboard.documents.item(item, "values")[1]
                for item in dashboard.documents.get_children()
            ]

        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)

        root.deiconify()
        root.update()
        popup = dashboard.open_column_filter(2)
        widgets = list(descendants(popup))
        next(
            widget
            for widget in widgets
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Select none"
        ).invoke()
        next(
            widget
            for widget in widgets
            if isinstance(widget, ttk.Checkbutton)
            and widget.cget("text") == "Manager A"
        ).invoke()
        next(
            widget
            for widget in widgets
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Apply"
        ).invoke()
        assert visible() == ["0", "1"]
        assert dashboard.documents.heading("manager", "text").endswith("●")
        popup = dashboard.open_column_filter(2)
        widgets = list(descendants(popup))
        next(widget for widget in widgets if isinstance(widget, ttk.Entry)).insert(
            0, "MANAGER B"
        )
        root.update()
        checks = [widget for widget in widgets if isinstance(widget, ttk.Checkbutton)]
        assert [widget.cget("text") for widget in checks if widget.winfo_manager()] == [
            "Manager B"
        ]
        next(
            widget
            for widget in widgets
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Select all"
        ).invoke()
        next(
            widget
            for widget in widgets
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Cancel"
        ).invoke()
        assert dashboard.column_filters == {2: {"Manager A"}}
        assert visible() == ["0", "1"]
        popup = dashboard.open_column_filter(2)
        widgets = list(descendants(popup))
        next(
            widget
            for widget in widgets
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Select all"
        ).invoke()
        next(
            widget
            for widget in widgets
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Apply"
        ).invoke()
        assert not dashboard.column_filters
        assert visible() == ["0", "1", "2", "3"]
        dashboard.column_filters[2] = {"Manager A"}
        dashboard.column_filters[4] = {"P"}
        dashboard.render_documents()
        assert visible() == ["0"]
        dashboard.result_filter.set("No")
        dashboard.render_documents()
        assert visible() == []
        dashboard.result_filter.set("All")
        dashboard.search.set("missing")
        assert visible() == []
        dashboard.search.set("")
        assert visible() == ["0"]
        dashboard.add_document(
            document_row("Yes", "4", "New", "Manager B", "P", None, URL, "", "")
        )
        dashboard.add_document(
            document_row("Yes", "5", "New", "Manager A", "P", None, URL, "", "")
        )
        assert visible() == ["0", "5"]
        dashboard.clear_document_filters()
        assert visible() == ["0", "1", "2", "3", "4", "5"]
        dashboard.column_filters[0] = {"Unresolved"}
        dashboard.render_documents()
        assert visible() == ["2"]
        dashboard.column_filters = {2: {""}}
        dashboard.render_documents()
        assert visible() == ["3"]
        dashboard.column_filters = {2: set()}
        dashboard.render_documents()
        assert visible() == []
        dashboard.clear_document_filters()
        assert not dashboard.column_filters
        assert dashboard.documents.heading("manager", "text").endswith("▾")
    finally:
        root.destroy()


def test_json_preferences_import_names_once_and_preserve_other_settings(tmp_path):
    from xccb_report.config import load_preferences, save_preferences
    from xccb_report.organizations import load_settings, save_settings

    legacy = tmp_path / "reviewer_names.txt"
    legacy.write_text("Doe, Jane\n", encoding="utf-8")
    save_settings(["04643"])
    assert load_preferences(legacy)["reviewer_names"] == ["Doe, Jane"]
    legacy.write_text("Other Person\n", encoding="utf-8")
    save_preferences(
        {"column_filters": {"Phase": {"mode": "value", "values": ["Final"]}}}
    )
    preferences = load_preferences(legacy)
    assert preferences["reviewer_names"] == ["Doe, Jane"]
    assert preferences["column_filters"]["Phase"]["values"] == ["Final"]
    assert load_settings()["organization_codes"] == ["04643"]


def test_json_preferences_need_no_organization_fields(tmp_path):
    from xccb_report.config import load_preferences
    from xccb_report.organizations import user_directory

    path = user_directory() / "settings.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"reviewer_names": ["Doe, Jane"], "column_filters": {}}),
        encoding="utf-8",
    )
    assert load_preferences()["reviewer_names"] == ["Doe, Jane"]


def test_reviewer_txt_persists_names_and_overrides_config(tmp_path):
    from xccb_report.reviewers import load_names, parse_names

    path = tmp_path / "reviewer_names.txt"
    assert parse_names("Doe, Jane\n \n doe,  JANE\nSmith, Alex\n") == [
        "Doe, Jane",
        "Smith, Alex",
    ]
    path.write_text("Doe, Jane\nSmith, Alex\n", encoding="utf-8")
    assert load_names(path) == ["Doe, Jane", "Smith, Alex"]
    path.write_text("\ufeffNew Reviewer\n", encoding="utf-8")
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps({"reviewer_names": ["Old Reviewer"]}), encoding="utf-8"
    )
    assert load_config(config_path)["reviewer_names"] == ["New Reviewer"]
    assert load_config(config_path)["assigned_by_names"] == ["New Reviewer"]
    assert parse_names("\n \n") == []


def test_dashboard_loads_legacy_names_once_then_reloads_json(tmp_path, monkeypatch):
    import tkinter as tk
    from xccb_report.config import save_preferences
    from xccb_report.gui import Dashboard

    monkeypatch.chdir(tmp_path)
    path = tmp_path / "reviewer_names.txt"
    path.write_text("Doe, Jane\nSmith, Alex\n", encoding="utf-8")
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        assert dashboard.reviewers.get("1.0", "end").strip() == "Doe, Jane\nSmith, Alex"
        save_preferences({"reviewer_names": ["New Reviewer"]})
        dashboard.reload_reviewers()
        assert dashboard.reviewers.get("1.0", "end").strip() == "New Reviewer"
        root.deiconify()
        for width, height in ((1000, 680), (800, 600)):
            root.geometry(f"{width}x{height}")
            root.update()
            table_height = dashboard.documents.winfo_height()
            assert table_height > height / 2 - 60
            for widget in (
                dashboard.documents,
                dashboard.reviewers,
                dashboard.run_button,
                dashboard.export_button,
            ):
                assert widget.winfo_rootx() >= root.winfo_rootx()
                assert widget.winfo_rooty() >= root.winfo_rooty()
                assert (
                    widget.winfo_rootx() + widget.winfo_width()
                    <= root.winfo_rootx() + root.winfo_width()
                )
                assert (
                    widget.winfo_rooty() + widget.winfo_height()
                    <= root.winfo_rooty() + root.winfo_height()
                )
    finally:
        root.destroy()


def test_dashboard_date_filters_persist_by_name_and_ignore_missing_columns(
    tmp_path, monkeypatch
):
    import tkinter as tk
    from tkinter import ttk
    from xccb_report.config import load_preferences, save_preferences
    from xccb_report.gui import Dashboard
    from xccb_report.results import document_row

    monkeypatch.chdir(tmp_path)
    save_preferences(
        {
            "reviewer_names": ["Doe, Jane"],
            "column_filters": {
                "Submitted": {"mode": "month", "values": ["2026-10"]},
                "Missing workbook column": {"mode": "value", "values": ["irrelevant"]},
            },
        }
    )
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        dashboard.documents_data = [
            document_row(
                "Yes",
                "1",
                "Title",
                "Manager",
                "Final",
                datetime(2026, 10, 1),
                URL,
                "",
                "",
            ),
            document_row(
                "Yes",
                "2",
                "Title",
                "Manager",
                "Final",
                datetime(2026, 10, 12),
                URL,
                "",
                "",
            ),
            document_row(
                "Yes",
                "3",
                "Title",
                "Manager",
                "Final",
                datetime(2026, 9, 1),
                URL,
                "",
                "",
            ),
        ]
        dashboard.render_documents()
        assert len(dashboard.documents.get_children()) == 2
        assert set(dashboard.column_filters) == {6}
        dashboard.date_modes[6] = "day"
        dashboard.column_filters[6] = {"2026-10-12"}
        dashboard.save_table_filters()
        dashboard.render_documents()
        assert len(dashboard.documents.get_children()) == 1
        assert load_preferences()["column_filters"]["Submitted"] == {
            "mode": "day",
            "values": ["2026-10-12"],
        }
        root.deiconify()
        root.update()
        popup = dashboard.open_column_filter(6)
        picker = next(
            widget
            for widget in popup.winfo_children()
            if isinstance(widget, ttk.Combobox)
        )
        selection_buttons = next(
            widget
            for widget in popup.winfo_children()
            if any(
                isinstance(child, ttk.Button) and child.cget("text") == "Select all"
                for child in widget.winfo_children()
            )
        )
        assert picker.grid_info()["row"] == selection_buttons.grid_info()["row"]
        root.update_idletasks()
        assert (
            picker.winfo_rootx()
            >= selection_buttons.winfo_rootx() + selection_buttons.winfo_width()
        )
        visible_columns = (
            "result",
            "submitted",
            "manager",
            "organization",
            "reviewers",
            "activity",
            "phase",
            "title",
            "url",
            "notes",
        )
        assert tuple(dashboard.documents.cget("displaycolumns")) == visible_columns
        picker.set("Month")
        picker.event_generate("<<ComboboxSelected>>")
        root.update()
        assert dashboard.date_modes[6] == "month"
        assert 6 not in dashboard.column_filters
        dashboard.close_column_filter()
        original = tmp_path / "source.xlsx"
        source(original)
        dashboard.configure_workbook_columns(original)
        assert tuple(dashboard.documents.cget("displaycolumns")) == tuple(
            column for column in visible_columns if column != "activity"
        )
        restarted = tk.Toplevel(root)
        second = Dashboard(restarted)
        assert second.date_modes[6] == "month"
        assert 6 not in second.column_filters
        assert second.reviewers.get("1.0", "end").strip() == "Doe, Jane"
        restarted.destroy()
    finally:
        root.destroy()


def test_report_retains_existing_exe_features_with_new_phase_position(tmp_path, monkeypatch):
    from xccb_report.results import document_row, read_results

    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    book.active["AQ3"] = "Activity Status"
    book.active["AQ4"] = "Approved"
    book.active["AQ5"] = "Cancelled"
    book.save(original)
    book.close()
    monkeypatch.setattr(
        "xccb_report.pipeline.load_evidence",
        lambda *args: ([{"assigned_by": "Person, One"}], True, "Validated", "local-evidence"),
    )
    progress = []
    counts = report(original, output, config(), tmp_path, all_dates=True, on_progress=lambda *args: progress.append(args))
    assert counts["Qualified"] == counts["No"] == counts["Duplicate"] == 1
    status, groups, rows = read_results(output)
    assert len(rows) == 1
    assert rows == [document_row(**entry[2]) for entry in progress if entry[2] is not None]
    assert rows[0][4] == "Final "
    assert rows[0][7:9] == ("Approved", "Reviewer, One")
    assert sum(group[2] for group in groups) == 1
    book = load_workbook(output)
    assert book["Flawless submissions"].max_row == 2
    assert len([row for row in book["Audit"].values if row[3] == "Duplicate"]) == 1
    assert dict(book["Summary"].values)["Duplicate rows skipped"] == 1
    book.close()


@pytest.mark.parametrize("with_activity", [False, True])
def test_dashboard_phase_follows_activity_without_losing_features(tmp_path, monkeypatch, with_activity):
    import tkinter as tk
    from xccb_report.gui import Dashboard, DOCUMENT_COLUMNS
    from xccb_report.results import document_row

    monkeypatch.chdir(tmp_path)
    original = tmp_path / "source.xlsx"
    source(original)
    if with_activity:
        book = load_workbook(original)
        book.active["AQ3"] = "Activity Status"
        book.active["AQ4"] = "Approved"
        book.save(original)
        book.close()
    root = tk.Tk()
    try:
        dashboard = Dashboard(root)
        dashboard.configure_workbook_columns(original)
        visible = tuple(dashboard.documents.cget("displaycolumns"))
        assert visible[:2] == ("result", "submitted")
        assert "reviewers" in visible
        if with_activity:
            assert visible[visible.index("activity") + 1] == "phase"
        else:
            assert "activity" not in visible and "phase" in visible
        phase = "ATM PCS Post TFC - Control Limit/RFC "
        dashboard.documents_data = [
            document_row("Yes", "123", "Zulu", "Manager", phase, None, URL, "", "", activity="Approved", reviewers="Reviewer, One"),
            document_row("No", "456", "Alpha", "Manager", "Preliminary ", None, URL.replace("123", "456"), "", "", activity="Cancelled", reviewers="Reviewer, Two"),
        ]
        dashboard.render_documents()
        title = next(index for index, column in enumerate(DOCUMENT_COLUMNS) if column[0] == "title")
        dashboard.toggle_document_sort(title)
        items = dashboard.documents.get_children()
        assert [dashboard.documents.set(item, "title") for item in items] == ["Alpha", "Zulu"]
        assert dashboard.documents.set(items[1], "phase") == phase
        assert dashboard.documents.set(items[1], "reviewers") == "Reviewer, One"
        dashboard.toggle_document_sort(title)
        assert [dashboard.documents.set(item, "title") for item in dashboard.documents.get_children()] == ["Zulu", "Alpha"]
        root.update()
        dashboard.documents.xview_moveto(1)
        root.update()
        if with_activity:
            assert dashboard.activity_labels
            for item, label in dashboard.activity_labels.items():
                result = dashboard.documents.set(item, "result")
                tag = "yes" if result == "Yes" else "no"
                assert label.cget("background") == str(dashboard.documents.tag_configure(tag, "background"))
                assert label.cget("font") == str(dashboard.pending_font)
                assert label.cget("foreground") == ("#1E7B34" if result == "Yes" else "#B3261E")
        phase_index = next(index for index, column in enumerate(DOCUMENT_COLUMNS) if column[0] == "phase")
        dashboard.open_column_filter(phase_index)
        assert dashboard.filter_popup is not None
    finally:
        root.destroy()


def test_all_dates_processing_does_not_change_default_cli_month(tmp_path):
    original = tmp_path / "source.xlsx"
    source(original)
    now = datetime(2026, 11, 4, tzinfo=timezone.utc)
    current = report(original, tmp_path / "current.xlsx", config(), tmp_path, now=now)
    all_dates = report(
        original, tmp_path / "all.xlsx", config(), tmp_path, now=now, all_dates=True
    )
    assert current.get("Qualified", 0) == 0
    assert all_dates["Qualified"] == 1


@pytest.mark.parametrize(
    "missing, column", [("Phase", "phase"), ("Document Submitted", "submitted")]
)
def test_missing_optional_workbook_columns_are_hidden_and_not_filtered(
    tmp_path, monkeypatch, missing, column
):
    import tkinter as tk
    from xccb_report.config import save_preferences
    from xccb_report.gui import Dashboard, DOCUMENT_COLUMNS
    from xccb_report.results import read_results

    monkeypatch.chdir(tmp_path)
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    index = next(cell.column for cell in book.active[3] if cell.value == missing)
    book.active.delete_cols(index)
    book.save(original)
    book.close()
    title = next(title for name, title, _ in DOCUMENT_COLUMNS if name == column)
    save_preferences(
        {"column_filters": {title: {"mode": "value", "values": ["not present"]}}}
    )
    counts = report(original, output, config(), tmp_path, all_dates=True)
    assert counts["Qualified"] == 2
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        dashboard.configure_workbook_columns(original)
        dashboard.documents_data = read_results(output)[2]
        dashboard.render_documents()
        assert column not in dashboard.documents.cget("displaycolumns")
        assert len(dashboard.documents.get_children()) == 2
    finally:
        root.destroy()
