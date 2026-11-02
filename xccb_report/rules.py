"""Business rules, independent of the browser and workbook writer."""
from dataclasses import dataclass
from datetime import date, datetime
from zoneinfo import ZoneInfo
import re

from openpyxl.utils.datetime import from_excel

PHASE_RULE = "All phases except values containing WLA (case-insensitive)"
REVIEWER_RULE = "Listed reviewer names or blank Reviewer(s) cells"


def normalized(value):
    return " ".join(str("" if value is None else value).split()).casefold()


def exact_name(value, names):
    return bool(normalized(value)) and normalized(value) in {normalized(n) for n in names}


def reviewer_match(value, names):
    if not normalized(value):
        return True
    # Commas belong to identities like 'Surname, Given name'.
    return any(exact_name(part, names) for part in re.split(r"[;\r\n]+", str(value)))


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
            return submitted_date(datetime.fromisoformat(value.strip()), epoch, timezone)
        except ValueError:
            pass
    raise ValueError("Missing or invalid Document Submitted date")


def month_bounds(now, timezone, reporting_month=None):
    if reporting_month is None:
        local = now.astimezone(ZoneInfo(timezone))
    else:
        if not isinstance(reporting_month, str) or not re.fullmatch(r"[0-9]{4}-(0[1-9]|1[0-2])", reporting_month):
            raise ValueError("Reporting month must be YYYY-MM")
        local = datetime.strptime(reporting_month, "%Y-%m")
    start = date(local.year, local.month, 1)
    end = date(local.year + (local.month == 12), local.month % 12 + 1, 1)
    return start, end


@dataclass(frozen=True)
class Eligibility:
    status: str
    reasons: tuple[str, ...] = ()


def eligibility(row, config, epoch, now, reporting_month=None):
    excluded, unresolved = [], []
    if "admin" not in normalized(row.get("Reviewing Forum")):
        excluded.append("Reviewing Forum does not contain admin")
    if normalized(row.get("Is FYI?")) != "no":
        excluded.append("Is FYI? is not No")
    if "wla" in normalized(row.get("Phase")):
        excluded.append(f"Phase contains WLA: {row.get('Phase')}")
    try:
        start, end = month_bounds(now, config["timezone"], reporting_month)
        if not start <= submitted_date(row.get("Document Submitted"), epoch, config["timezone"]) < end:
            excluded.append("Document Submitted outside reporting month")
    except ValueError as exc:
        unresolved.append(str(exc))
    if not reviewer_match(row.get("Reviewer(s)"), config.get("reviewer_names", [])):
        excluded.append("Reviewer(s) does not match the allowed name list")
    if excluded:
        return Eligibility("Excluded", tuple(excluded + unresolved))
    return Eligibility("Unverified" if unresolved else "Qualified", tuple(unresolved))


def classify(actions, names, *, complete, scope):
    if scope != "document":
        return "", [], "Actions scope unconfirmed or unsupported"
    if not names:
        return "", [], "Assigned by name list not configured"
    matches = sorted({a["assigned_by"] for a in actions if exact_name(a.get("assigned_by"), names)})
    if matches:
        return "No", matches, "Matching Assigned by identity found"
    if not complete:
        return "", [], "Actions completeness not verified"
    return "Yes", [], "All document Actions inspected; no identity matched"
