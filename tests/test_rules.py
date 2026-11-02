from datetime import date, datetime, timezone

import pytest
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, to_excel

from xccb_report.actions import parse_actions, workspace_identity
from xccb_report.rules import (
    classify,
    eligibility,
    exact_name,
    month_bounds,
    submitted_date,
)


def test_exact_name_ignores_case_and_whitespace_but_rejects_similar_names():
    assert exact_name("  Jane\t Doe ", ["jane doe"])
    assert not exact_name("Jane Doe Jr.", ["Jane Doe"])


def test_one_matching_assigned_by_among_multiple_actions_classifies_no():
    result = classify(
        [
            {"assigned_by": "Someone Else"},
            {"assigned_by": "  DOE,\t Jane "},
            {"assigned_by": "Doe, Jane"},
        ],
        ["Doe, Jane"],
        complete=True,
        scope="document",
    )
    assert result == ("No", ["  DOE,\t Jane ", "Doe, Jane"], "Matching Assigned by identity found")


def test_incomplete_nonmatch_and_unconfirmed_scope_remain_blank():
    assert classify([], ["Doe, Jane"], complete=False, scope="document")[0] == ""
    assert classify([], ["Doe, Jane"], complete=True, scope="unconfirmed")[0] == ""


def test_complete_document_scope_with_zero_actions_classifies_yes():
    assert classify([], ["Doe, Jane"], complete=True, scope="document")[0] == "Yes"


@pytest.mark.parametrize("reviewers,expected", [(None, "Qualified"), (" \t\n", "Qualified"),
                                               (0, "Excluded"), (False, "Excluded"),
                                               ("Any Other Person", "Excluded"),
                                               ("Doe, Jane; Smith, Alex", "Qualified"),
                                               ("Smith, Alex\n DOE,  Jane ", "Qualified"),
                                               ("Doe, Jane Jr.", "Excluded")])
def test_reviewer_accepts_listed_names_or_blanks(reviewers, expected):
    row, config, epoch, now = _eligibility_inputs()
    row["Reviewer(s)"] = reviewers
    assert eligibility(row, config, epoch, now).status == expected


@pytest.mark.parametrize("epoch", [WINDOWS_EPOCH, MAC_EPOCH])
def test_excel_serial_is_interpreted_using_the_workbook_epoch(epoch):
    value = datetime(2026, 3, 15, 12, 30)
    serial = to_excel(value, epoch)
    assert submitted_date(serial, epoch, "Asia/Kuala_Lumpur") == date(2026, 3, 15)


def test_reporting_month_and_submitted_dates_use_malaysia_midnight_boundaries():
    malaysia_midnight = datetime(2026, 1, 31, 16, 0, tzinfo=timezone.utc)
    assert month_bounds(malaysia_midnight, "Asia/Kuala_Lumpur") == (
        date(2026, 2, 1),
        date(2026, 3, 1),
    )
    assert submitted_date(malaysia_midnight, WINDOWS_EPOCH, "Asia/Kuala_Lumpur") == date(2026, 2, 1)
    assert submitted_date(
        datetime(2026, 1, 31, 15, 59, tzinfo=timezone.utc),
        WINDOWS_EPOCH,
        "Asia/Kuala_Lumpur",
    ) == date(2026, 1, 31)


def _eligibility_inputs(phase="Final", author_manager="Alice Manager", submitted=None):
    row = {
        "Reviewing Forum": "Admin Forum",
        "Is FYI?": "No",
        "Phase": phase,
        "Document Submitted": submitted or datetime(2026, 6, 15, 12),
        "Reviewer(s)": "Doe, Jane",
        "Author Manager": author_manager,
    }
    config = {
        "timezone": "Asia/Kuala_Lumpur",
        "reviewer_names": ["Doe, Jane"],
    }
    now = datetime(2026, 6, 15, 12, tzinfo=timezone.utc)
    return row, config, WINDOWS_EPOCH, now


@pytest.mark.parametrize("phase", [" \tFINAL\n ", "Preliminary", "Final Approval", "Do Not Implement",
                                  "ATM Class III-N", "LSI Modification", "", None])
def test_any_phase_without_wla_qualifies(phase):
    row, config, epoch, now = _eligibility_inputs(phase=phase)
    assert eligibility(row, config, epoch, now).status == "Qualified"


@pytest.mark.parametrize("manager", [None, "", "Unknown Manager"])
def test_author_manager_does_not_filter_eligibility(manager):
    row, config, epoch, now = _eligibility_inputs(author_manager=manager)
    assert eligibility(row, config, epoch, now).status == "Qualified"


@pytest.mark.parametrize("phase", ["WLA CCB Final", "WLA CCB Preliminary and PrePilot", " wLa ", "Final WLA"])
def test_any_phase_containing_wla_is_excluded(phase):
    row, config, epoch, now = _eligibility_inputs(phase=phase)
    assert eligibility(row, config, epoch, now).status == "Excluded"
    assert any(reason.startswith("Phase contains WLA:") for reason in eligibility(row, config, epoch, now).reasons)


def test_eligibility_includes_malaysia_month_start_and_excludes_next_month_start():
    row, config, epoch, _ = _eligibility_inputs(
        submitted=datetime(2026, 1, 31, 16, 0, tzinfo=timezone.utc)
    )
    now = datetime(2026, 1, 31, 16, 0, tzinfo=timezone.utc)
    assert eligibility(row, config, epoch, now).status == "Qualified"

    row["Document Submitted"] = datetime(2026, 2, 28, 16, 0, tzinfo=timezone.utc)
    result = eligibility(row, config, epoch, now)
    assert result.status == "Excluded"
    assert "Document Submitted outside reporting month" in result.reasons


def test_workspace_identity_accepts_document_workspace_and_rejects_malformed_or_foreign_urls():
    document_id, canonical = workspace_identity(
        "https://xccb2.intel.com/Documents/DocumentWorkspace.aspx?DocumentID=12345"
    )
    assert document_id == "12345"
    assert canonical.endswith("?DocumentID=12345")

    for url in (
        "https://example.com/Documents/DocumentWorkspace.aspx?DocumentID=12345",
        "https://xccb2.intel.com.evil.example/Documents/DocumentWorkspace.aspx?DocumentID=12345",
        "https://xccb2.intel.com/Documents/Other.aspx?DocumentID=12345",
        "https://xccb2.intel.com/Documents/DocumentWorkspace.aspx",
        "https://xccb2.intel.com/Documents/DocumentWorkspace.aspx?DocumentID=1&documentid=2",
    ):
        with pytest.raises(ValueError):
            workspace_identity(url)


def _action_row(action_id="7", document_id="123", assigned_by="Doe, Jane"):
    assigner = f"<span>Assigned By: {assigned_by}</span>" if assigned_by is not None else "<span>Jane Doe</span>"
    return (
        "<tr><td><a href='/Actions/CompleteDocumentAction.aspx?ActionID="
        f"{action_id}&amp;DocumentID={document_id}'>Open</a></td>"
        f"<td>{assigner}</td><td><span>Assignee: Alex Worker</span></td></tr>"
    )


def test_parser_uses_assigned_by_label_not_adjacent_assignee():
    html = "<table class='action-panel'>" + _action_row() + "</table>"
    actions, errors = parse_actions(html, "123")
    assert actions == [{"action_id": "7", "assigned_by": "Doe, Jane"}]
    assert errors == []


def test_parser_reports_pagination_and_malformed_missing_assigner():
    paginated = (
        "<table class='action-panel'>" + _action_row()
        + "</table><a href='Document.aspx?Page$Next=2'>Next</a>"
    )
    _, errors = parse_actions(paginated, "123")
    assert any("pagination or incomplete listing" in error for error in errors)

    malformed = "<table class='action-panel'>" + _action_row(assigned_by=None) + "</table>"
    actions, errors = parse_actions(malformed, "123")
    assert actions == []
    assert any("Malformed action identity" in error for error in errors)


def test_parser_reports_absent_action_panel_table():
    actions, errors = parse_actions("<div>No action table</div>", "123")
    assert actions == []
    assert errors == ["Expected exactly one action-panel table; unloaded/empty layout is not verified"]
