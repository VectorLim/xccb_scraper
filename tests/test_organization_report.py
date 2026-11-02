from datetime import datetime, timezone

from openpyxl import load_workbook
import pytest

from test_pipeline import config, source
from xccb_report.gui import DOCUMENT_COLUMNS, read_results
from xccb_report.organizations import OrganizationCache, load_settings, save_settings
from xccb_report.workbook import report


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
        return [{"wwid": name, "organization_code": "04643" if name == "Manager, One" else "64530",
                 "organization": "KuAT AE CAM" if name == "Manager, One" else "KuAT Test Engineering"}]
    save_settings(["04643"])
    assert load_settings()["organization_codes"] == ["04643"]
    output = tmp_path / "report.xlsx"
    counts = report(original, output, config(), tmp_path, datetime(2026, 10, 30, tzinfo=timezone.utc), organization_query=query)
    assert counts["Qualified"] == 1 and counts["Excluded"] == 1
    assert len(calls) == 2
    book = load_workbook(output)
    try:
        sheet = book["Flawless submissions"]
        assert [cell.value for cell in sheet[1]][:3] == ["Author Manager", "Organization", "Reviewing Forum"]
        assert sheet["A2"].value == "Manager, One" and sheet["B2"].value == "KuAT AE CAM"
        audit = list(book["Audit"].values)
        assert audit[0][12:16] == ("Author Manager", "Organization", "Organization code", "WWID")
        assert audit[1][14] == "04643" and audit[2][14] == "64530"
        assert audit[2][3] == "Excluded"
        assert list(book["By Author Manager"].values)[1][:3] == ("Manager, One", "KuAT AE CAM", 1)
        assert dict(book["Summary"].values)["Selected organization codes"] == "04643"
    finally:
        book.close()
    status, groups, documents = read_results(output)
    assert DOCUMENT_COLUMNS[4][1] == "Organization" and documents[0][4] == "KuAT AE CAM"
    again = report(original, tmp_path / "again.xlsx", config(), tmp_path, datetime(2026, 10, 30, tzinfo=timezone.utc),
                   organization_query=lambda *args: pytest.fail("cached entry queried"))
    assert again == counts


def test_ambiguous_names_stay_in_audit_not_selected_report(tmp_path):
    original, output = tmp_path / "source.xlsx", tmp_path / "report.xlsx"
    source(original)
    def query(*args):
        return [{"wwid": str(index), "organization_code": "04643", "organization": "KuAT AE CAM"} for index in (1, 2)]
    counts = report(original, output, config(), tmp_path, datetime(2026, 10, 30, tzinfo=timezone.utc),
                    organization_codes=["04643"], organization_query=query)
    assert counts["Unverified"] == 2 and counts.get("Qualified", 0) == 0
    assert not OrganizationCache().entries
    book = load_workbook(output)
    try:
        assert book["Audit"]["N2"].value == "Unresolved"
        assert "ambiguous" in book["Audit"]["R2"].value
        assert book["Flawless submissions"].max_row == 1
    finally:
        book.close()


def test_dashboard_organization_picker_persistence_and_filters(tmp_path, monkeypatch):
    import tkinter as tk
    from tkinter import ttk
    from xccb_report.gui import Dashboard, document_row
    from xccb_report.organizations import save_python
    monkeypatch.chdir(tmp_path)
    cache = OrganizationCache()
    cache.resolve(["Manager"], lambda *args: [{"wwid": "1", "organization_code": "04643", "organization": "KuAT AE CAM"}])
    save_python("approved-python.exe")
    root = tk.Tk()
    root.withdraw()
    def descendants(widget):
        for child in widget.winfo_children():
            yield child
            yield from descendants(child)
    try:
        dashboard = Dashboard(root)
        dashboard.report_bytes = b"old report"
        dashboard.choose_organizations()
        popup = next(widget for widget in root.winfo_children() if isinstance(widget, tk.Toplevel))
        widgets = list(descendants(popup))
        next(widget for widget in widgets if isinstance(widget, tk.Listbox)).selection_set(0)
        next(widget for widget in widgets if isinstance(widget, ttk.Button) and widget.cget("text") == "Apply").invoke()
        assert dashboard.organization_codes == ["04643"] and dashboard.report_bytes is None
        assert str(dashboard.export_button.cget("state")) == "disabled"
        assert load_settings()["organization_python"] == "approved-python.exe"
        dashboard.documents_data = [document_row("Yes", "1", "Title", "Manager", "P", None, "", "", "", "KuAT AE CAM")]
        dashboard.search.set("KuAT AE CAM")
        assert len(dashboard.documents.get_children()) == 1
        dashboard.column_filters[4] = {"Other"}
        dashboard.render_documents()
        assert not dashboard.documents.get_children()
        assert dashboard.organization_codes == ["04643"]
        root.destroy()
        root = tk.Tk()
        root.withdraw()
        assert Dashboard(root).organization_codes == ["04643"]
    finally:
        root.destroy()


def test_dashboard_identity_confirmation_is_explicit(tmp_path, monkeypatch):
    import tkinter as tk
    from tkinter import ttk
    from queue import Queue
    from xccb_report.gui import Dashboard
    monkeypatch.chdir(tmp_path)
    root = tk.Tk()
    root.withdraw()
    try:
        dashboard = Dashboard(root)
        answer = Queue()
        candidates = [{"wwid": "1", "name": "Same Name", "site": "KM", "organization_code": "04643",
                       "organization": "KuAT AE CAM", "cost_center": "Assembly"},
                      {"wwid": "2", "name": "Same Name", "site": "PG", "organization_code": "64530",
                       "organization": "KuAT Test Engineering", "cost_center": "Test"}]
        dashboard.identity_dialog("Same Name", candidates, answer)
        popup = next(widget for widget in root.winfo_children() if isinstance(widget, tk.Toplevel))
        choices = next(widget for widget in popup.winfo_children() if isinstance(widget, tk.Listbox))
        buttons = [child for widget in popup.winfo_children() for child in widget.winfo_children() if isinstance(child, ttk.Button)]
        confirm = next(widget for widget in buttons if widget.cget("text") == "Confirm identity")
        confirm.invoke()
        assert answer.empty()
        choices.selection_set(1)
        confirm.invoke()
        assert answer.get_nowait() == "2"
    finally:
        root.destroy()