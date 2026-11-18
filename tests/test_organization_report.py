from datetime import datetime, timezone

from openpyxl import load_workbook
import pytest

from test_pipeline import config, source
from xccb_report.gui import DOCUMENT_COLUMNS
from xccb_report.organizations import OrganizationCache, load_settings, save_settings
from xccb_report.pipeline import report
from xccb_report.results import read_results


@pytest.fixture(autouse=True)
def private_user(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))


def test_organization_columns_report_selection_and_restart(tmp_path):
    original = tmp_path / "source.xlsx"
    source(original)
    book = load_workbook(original)
    book.active["A5"] = "Other Manager"
    book.save(original)
    book.close()
    calls = []

    def query(name, wwid):
        calls.append(name)
        return [
            {
                "wwid": name,
                "organization_code": "04643" if name == "Manager, One" else "64530",
                "organization": (
                    "KuAT AE CAM" if name == "Manager, One" else "KuAT Test Engineering"
                ),
            }
        ]

    save_settings(["04643"])
    assert load_settings()["organization_codes"] == ["04643"]
    output = tmp_path / "report.xlsx"
    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        organization_query=query,
        organization_codes=["04643"],
    )
    assert counts["Qualified"] == 1 and counts["Excluded"] == 1
    assert len(calls) == 2
    book = load_workbook(output)
    try:
        sheet = book["Flawless submissions"]
        assert [cell.value for cell in sheet[1]][:3] == [
            "Author Manager",
            "Organization",
            "Reviewing Forum",
        ]
        assert (
            sheet["A2"].value == "Manager, One" and sheet["B2"].value == "KuAT AE CAM"
        )
        audit = list(book["Audit"].values)
        assert audit[0][12:16] == (
            "Author Manager",
            "Organization",
            "Organization code",
            "WWID",
        )
        assert audit[1][14] == "04643" and audit[2][14] == "64530"
        assert audit[2][3] == "Excluded"
        assert list(book["By Author Manager"].values)[1][:3] == (
            "Manager, One",
            "KuAT AE CAM",
            1,
        )
        assert dict(book["Summary"].values)["Selected organization codes"] == "04643"
    finally:
        book.close()
    status, groups, documents = read_results(output)
    assert DOCUMENT_COLUMNS[3][1] == "Organization" and documents[0][3] == "KuAT AE CAM"
    again = report(
        original,
        tmp_path / "again.xlsx",
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        organization_query=lambda *args: pytest.fail("cached entry queried"),
        organization_codes=["04643"],
    )
    assert again == counts


@pytest.mark.parametrize(
    "column, value",
    [
        ("B", "Other Forum"),
        ("C", "Other Reviewer"),
        ("E", "WLA"),
        ("E", "LSI Modification"),
        ("F", datetime(2025, 1, 1)),
        ("F", None),
    ],
)
def test_only_dashboard_eligible_managers_are_queried(tmp_path, column, value):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    book.active["A5"] = "Excluded Manager"
    book.active[f"{column}5"] = value
    book.save(original)
    book.close()
    save_settings(["99999"])
    calls = []

    def query(name, wwid):
        calls.append(name)
        return [
            {"wwid": "1", "organization_code": "04643", "organization": "KuAT AE CAM"}
        ]

    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        organization_query=query,
    )
    assert calls == ["Manager, One"]
    assert counts["Qualified"] == 1
    assert set(OrganizationCache().entries) == {"manager, one"}
    book = load_workbook(output)
    try:
        excluded = next(
            row
            for row in book["Audit"].iter_rows(min_row=2, values_only=True)
            if row[0] == 5
        )
        assert excluded[13] is None
        assert excluded[17] == "Not looked up: submission not eligible"
    finally:
        book.close()


def test_final_filter_completes_before_single_batch_and_single_setup(
    tmp_path, monkeypatch
):
    from xccb_report.organizations import DirectoryUnavailable
    from xccb_report.filtering import eligibility

    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    book = load_workbook(original)
    sheet = book.active
    sheet["A5"] = "Other Selected Manager"
    sheet["G5"].hyperlink = sheet["G4"].hyperlink.target.replace("123", "456")
    for name in ("Excluded One", "Excluded Two"):
        sheet.append([cell.value for cell in sheet[4]])
        sheet.cell(sheet.max_row, 1, name)
        sheet.cell(sheet.max_row, 3, "Other Reviewer")
    book.save(original)
    book.close()
    checked, prompts, calls = [], [], []

    def filter_row(row, *args):
        checked.append(row["Author Manager"])
        return eligibility(row, *args)

    monkeypatch.setattr("xccb_report.pipeline.eligibility", filter_row)

    def query(names, python):
        assert len(checked) == 4
        assert names == ["Manager, One", "Other Selected Manager"]
        calls.append((names, python))
        if not python:
            raise DirectoryUnavailable("Python lacks DataSyncX")
        return {
            name.casefold(): [
                {
                    "name": name,
                    "wwid": name,
                    "organization_code": "04643",
                    "organization": "KuAT AE CAM",
                }
            ]
            for name in names
        }

    monkeypatch.setattr("xccb_report.organizations.query_directory_batch", query)
    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        on_directory_unavailable=lambda reason: prompts.append(reason)
        or "approved-python.exe",
    )
    assert counts["Qualified"] == 2 and counts["Excluded"] == 2
    assert len(prompts) == 1 and len(calls) == 2
    assert len(checked) == 4 and len(OrganizationCache().entries) == 2


def test_missing_datasyncx_runtime_can_be_selected_and_retried(tmp_path, monkeypatch):
    from xccb_report.organizations import DirectoryUnavailable

    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    calls, prompts = [], []

    def query(names, python):
        calls.append(python)
        if not python:
            raise DirectoryUnavailable("Selected Python has no DataSyncX")
        return {
            name.casefold(): [
                {
                    "name": name,
                    "wwid": "1",
                    "organization_code": "04643",
                    "organization": "KuAT AE CAM",
                }
            ]
            for name in names
        }

    monkeypatch.setattr("xccb_report.organizations.query_directory_batch", query)

    def setup(reason):
        prompts.append(reason)
        return "approved-python.exe"

    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        on_directory_unavailable=setup,
    )
    assert counts["Qualified"] == 1 and calls == [None, "approved-python.exe"]
    assert len(prompts) == 1 and "no DataSyncX" in prompts[0]
    assert load_settings()["organization_python"] == "approved-python.exe"
    assert len(OrganizationCache().entries) == 1


def test_invalid_runtime_retry_does_not_replace_saved_settings(tmp_path, monkeypatch):
    from xccb_report.organizations import DirectoryUnavailable, save_python

    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    save_python("previous-python.exe")
    monkeypatch.setattr(
        "xccb_report.organizations.query_directory_batch",
        lambda *args: (_ for _ in ()).throw(
            DirectoryUnavailable("DataSyncX unavailable")
        ),
    )
    prompts = []
    report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        on_directory_unavailable=lambda reason: prompts.append(reason)
        or "invalid-python.exe",
    )
    assert len(prompts) == 1
    assert load_settings()["organization_python"] == "previous-python.exe"
    assert not OrganizationCache().entries


def test_ambiguous_names_stay_in_audit_not_selected_report(tmp_path):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)

    def query(*args):
        return [
            {
                "wwid": str(index),
                "organization_code": "04643",
                "organization": "KuAT AE CAM",
            }
            for index in (1, 2)
        ]

    counts = report(
        original,
        output,
        config(),
        tmp_path,
        datetime(2026, 10, 30, tzinfo=timezone.utc),
        organization_codes=["04643"],
        organization_query=query,
    )
    assert counts["Unverified"] == 2 and counts.get("Qualified", 0) == 0
    assert not OrganizationCache().entries
    book = load_workbook(output)
    try:
        assert book["Audit"]["N2"].value == "Unresolved"
        assert "ambiguous" in book["Audit"]["R2"].value
        assert book["Flawless submissions"].max_row == 1
    finally:
        book.close()


def test_dashboard_defaults_to_all_organizations_without_sidebar_button(
    tmp_path, monkeypatch
):
    import tkinter as tk
    from xccb_report.gui import Dashboard
    from xccb_report.organizations import save_python
    from xccb_report.results import document_row

    monkeypatch.chdir(tmp_path)
    cache = OrganizationCache()
    cache.resolve(
        ["Manager"],
        lambda *args: [
            {"wwid": "1", "organization_code": "04643", "organization": "KuAT AE CAM"}
        ],
    )
    save_python("approved-python.exe")
    save_settings(["99999"])
    root = tk.Tk()
    root.withdraw()

    try:
        dashboard = Dashboard(root)
        assert not hasattr(dashboard, "organization_button")
        assert not hasattr(dashboard, "organization_codes")
        assert load_settings()["organization_python"] == "approved-python.exe"
        dashboard.documents_data = [
            document_row(
                "Yes", "1", "Title", "Manager", "P", None, "", "", "", "KuAT AE CAM"
            )
        ]
        dashboard.search.set("KuAT AE CAM")
        assert len(dashboard.documents.get_children()) == 1
        dashboard.column_filters[3] = {"Other"}
        dashboard.render_documents()
        assert not dashboard.documents.get_children()
        assert load_settings()["organization_codes"] == ["99999"]
        root.destroy()
        root = tk.Tk()
        root.withdraw()
        assert not hasattr(Dashboard(root), "organization_button")
    finally:
        root.destroy()


def test_column_filters_are_inline_and_dismissible(tmp_path, monkeypatch):
    import tkinter as tk
    from xccb_report.gui import Dashboard
    from xccb_report.results import document_row

    monkeypatch.chdir(tmp_path)
    root = tk.Tk()
    try:
        dashboard = Dashboard(root)
        root.geometry("800x600")
        dashboard.documents_data = [
            document_row(
                "Yes", "1", "Title", "Manager", "P", None, "", "", "", "KuAT AE CAM"
            )
        ]
        dashboard.render_documents()
        root.update()
        popup = dashboard.open_column_filter(4)
        root.update()
        assert not isinstance(popup, tk.Toplevel) and popup.master == root
        assert not any(
            isinstance(widget, tk.Toplevel) for widget in root.winfo_children()
        )
        assert (
            popup.winfo_x() >= 0
            and popup.winfo_x() + popup.winfo_width() <= root.winfo_width()
        )
        assert popup.winfo_y() + popup.winfo_height() <= root.winfo_height()
        replacement = dashboard.open_column_filter(3)
        assert not popup.winfo_exists() and dashboard.filter_popup == replacement
        root.event_generate("<Escape>")
        root.update()
        assert dashboard.filter_popup is None and not dashboard.filter_bindings
        popup = dashboard.open_column_filter(4)
        dashboard.reviewers.event_generate("<Button-1>")
        root.update()
        assert dashboard.filter_popup is None and not popup.winfo_exists()
        assert not dashboard.filter_bindings
    finally:
        root.destroy()


@pytest.mark.parametrize("choose", [True, False])
def test_dashboard_runtime_prompt_answers_worker_without_settings_side_effects(
    tmp_path, monkeypatch, choose
):
    import tkinter as tk
    from queue import Queue
    from xccb_report.gui import Dashboard

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(
        "xccb_report.gui.messagebox.askyesno", lambda *args, **kwargs: choose
    )
    paths = []
    monkeypatch.setattr(
        "xccb_report.gui.filedialog.askopenfilename",
        lambda **kwargs: paths.append(kwargs) or "approved-python.exe",
    )
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        answer = Queue()
        dashboard.events.put(
            ("organization_setup", ("No DataSyncX in selected Python", answer))
        )
        dashboard.poll()
        assert answer.get_nowait() == ("approved-python.exe" if choose else None)
        assert bool(paths) == choose
        assert "organization_python" not in load_settings()
    finally:
        root.destroy()


def test_dashboard_has_no_identity_confirmation(tmp_path, monkeypatch):
    import tkinter as tk
    from xccb_report.gui import Dashboard

    monkeypatch.chdir(tmp_path)
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        assert not hasattr(dashboard, "confirm_identity")
        assert not hasattr(dashboard, "identity_dialog")
        assert not any(
            isinstance(widget, tk.Toplevel) for widget in root.winfo_children()
        )
    finally:
        root.destroy()
