"""Submission filtering and classification, independent of browser and workbook I/O."""

from dataclasses import dataclass
from datetime import date, datetime
from zoneinfo import ZoneInfo
import re

from openpyxl.utils.datetime import from_excel

PHASE_RULE = (
    "All phases except values containing WLA or LSI modification (case-insensitive)"
)
REVIEWER_RULE = (
    "Reviewer(s) is blank or contains a listed reviewer name (case-insensitive)"
)


def normalized(value):
    return " ".join(str("" if value is None else value).split()).casefold()


def exact_name(value, names):
    value = normalized(value)
    return bool(value) and value in {normalized(n) for n in names}


def reviewer_match(value, names):
    value = normalized(value)
    return not value or any(name and name in value for name in map(normalized, names))


def submitted_date(value, epoch, timezone):
    if isinstance(value, bool):
        raise ValueError("Invalid submitted date")
    if isinstance(value, (int, float)):
        value = from_excel(value, epoch)
    if isinstance(value, datetime):
        return (value.astimezone(ZoneInfo(timezone)) if value.tzinfo else value).date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return submitted_date(
                datetime.fromisoformat(value.strip()), epoch, timezone
            )
        except ValueError:
            pass
    raise ValueError("Missing or invalid Document Submitted date")


def month_bounds(now, timezone, reporting_month=None):
    if reporting_month is None:
        local = now.astimezone(ZoneInfo(timezone))
    else:
        if not isinstance(reporting_month, str) or not re.fullmatch(
            r"[0-9]{4}-(0[1-9]|1[0-2])", reporting_month
        ):
            raise ValueError("Reporting month must be YYYY-MM")
        local = datetime.strptime(reporting_month, "%Y-%m")
    start = date(local.year, local.month, 1)
    end = date(local.year + (local.month == 12), local.month % 12 + 1, 1)
    return start, end


@dataclass(frozen=True)
class Eligibility:
    status: str
    reasons: tuple[str, ...] = ()


def eligibility(row, config, epoch, now, reporting_month=None, *, all_dates=False):
    excluded, unresolved = [], []
    if "admin" not in normalized(row.get("Reviewing Forum")):
        excluded.append("Reviewing Forum does not contain admin")
    phase = normalized(row.get("Phase"))
    for text in ("WLA", "LSI modification"):
        if normalized(text) in phase:
            excluded.append(f"Phase contains {text}: {row.get('Phase')}")
    try:
        submitted = (
            None
            if all_dates and "Document Submitted" not in row
            else submitted_date(
                row.get("Document Submitted"), epoch, config["timezone"]
            )
        )
        if not all_dates:
            start, end = month_bounds(now, config["timezone"], reporting_month)
            if not start <= submitted < end:
                excluded.append("Document Submitted outside reporting month")
    except ValueError as exc:
        unresolved.append(str(exc))
    if not reviewer_match(row.get("Reviewer(s)"), config.get("reviewer_names", [])):
        excluded.append("Reviewer(s) does not contain a listed reviewer name")
    if excluded:
        return Eligibility("Excluded", tuple(excluded + unresolved))
    return Eligibility("Unverified" if unresolved else "Qualified", tuple(unresolved))


def organization_eligibility(decision, organization, codes):
    if not codes or decision.status == "Excluded":
        return decision
    if organization is None:
        return Eligibility(
            "Unverified",
            decision.reasons + ("Organization selection: identity unresolved",),
        )
    if organization.organization_code not in codes:
        return Eligibility(
            "Excluded", decision.reasons + ("Organization code not selected",)
        )
    return decision


def classify(actions, names, *, complete, scope):
    if scope != "document":
        return "", [], "Actions scope unconfirmed or unsupported"
    if not names:
        return "", [], "Assigned by name list not configured"
    matches = sorted(
        {a["assigned_by"] for a in actions if exact_name(a.get("assigned_by"), names)}
    )
    if matches:
        return "No", matches, "Matching Assigned by identity found"
    if not complete:
        return "", [], "Actions completeness not verified"
    return "Yes", [], "All document Actions inspected; no identity matched"
