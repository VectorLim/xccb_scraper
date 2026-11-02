"""Workbook loading, filtering and format-preserving report output."""
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import copy
from datetime import datetime, timezone
import json
import logging
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .actions import workspace_identity
from .evidence import load_evidence
from .rules import PHASE_RULE, REVIEWER_RULE, classify, eligibility, month_bounds, normalized
from .reviewers import load_names

REQUIRED = {"Author Manager", "Reviewing Forum", "Reviewer(s)", "Is FYI?", "Phase", "Document Submitted", "Document WS URL"}
DEFAULTS = {"timezone": "Asia/Kuala_Lumpur", "assigned_by_names": [], "reviewer_names": [],
            "action_scope": "unconfirmed", "evidence_max_age_hours": 24}
LEGACY_FILTER_KEYS = {"reviewer_separator", "manager_organizations", "allowed_organizations", "phases"}


def load_config(path):
    supplied = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(supplied, dict) or set(supplied) - set(DEFAULTS) - LEGACY_FILTER_KEYS:
        raise ValueError("Configuration must be an object with supported keys")
    for key in LEGACY_FILTER_KEYS & supplied.keys():
        supplied.pop(key)
        logging.warning("Ignoring obsolete filter setting: %s", key)
    config = DEFAULTS | supplied
    for key in ("assigned_by_names", "reviewer_names"):
        if not isinstance(config[key], list) or any(not isinstance(v, str) or not normalized(v) for v in config[key]):
            raise ValueError(f"{key} must be a list of nonblank names")
    if config["action_scope"] not in {"unconfirmed", "document"}:
        raise ValueError("Only whole-document scope is supported; submission/phase scope requires verified metadata")
    if not isinstance(config["evidence_max_age_hours"], (int, float)) or config["evidence_max_age_hours"] <= 0:
        raise ValueError("Evidence age must be positive")
    month_bounds(datetime.now(timezone.utc), config["timezone"])
    reviewer_path = Path(path).resolve().parent / "reviewer_names.txt"
    if reviewer_path.exists():
        # One listed-names file drives both the Reviewer(s) filter and the Assigned by check.
        config["reviewer_names"] = config["assigned_by_names"] = load_names(reviewer_path)
    return config


def read_source(path):
    workbook = load_workbook(path)
    sheets = [s for s in workbook if REQUIRED <= {c.value for c in s[3]}]
    if len(sheets) != 1:
        raise ValueError("Expected exactly one report sheet with headers on row 3")
    sheet = sheets[0]
    headers = [c.value for c in sheet[3]]
    if len(headers) != 43 or len(set(headers)) != 43 or any(not h for h in headers):
        raise ValueError("Expected 43 unique original columns A:AQ")
    return workbook, sheet, headers


def copy_style(source, target):
    # Style indices belong to a workbook; register their components in the new one.
    for attribute in ("font", "fill", "border", "alignment", "protection"):
        setattr(target, attribute, copy(getattr(source, attribute)))
    target.number_format = source.number_format


def report(source, output, config, evidence_dir, now=None, reader=None, *, on_progress=None, reporting_month=None):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.exists():
        raise ValueError("Choose a new output file; existing workbooks are never overwritten")
    fixed_time = now is not None
    now = now or datetime.now(timezone.utc)
    start, end = month_bounds(now, config["timezone"], reporting_month)
    book, sheet, headers = read_source(source)
    out = Workbook()
    out.epoch = book.epoch
    result_sheet = out.active
    result_sheet.title = "Flawless submissions"
    summary = out.create_sheet("Summary")
    audit = out.create_sheet("Audit")
    managers = out.create_sheet("By Author Manager")
    managers.append(["Author Manager", "Qualified rows", "Yes", "No", "Unresolved Actions"])
    for key, dimension in sheet.column_dimensions.items():
        target_dimension = result_sheet.column_dimensions[key]
        for attribute in ("width", "hidden", "bestFit", "min", "max", "outlineLevel", "collapsed"):
            setattr(target_dimension, attribute, getattr(dimension, attribute))
        copy_style(dimension, target_dimension)
    for column, cell in enumerate(sheet[3], 1):
        target = result_sheet.cell(1, column, cell.value)
        copy_style(cell, target)
        target.alignment = Alignment(wrap_text=True, vertical="top")
    result_sheet.cell(1, 44, "Flawless submission")
    result_sheet.cell(1, 44)._style = copy(result_sheet.cell(1, 43)._style)
    result_sheet.column_dimensions["AR"].width = 24
    result_sheet.row_dimensions[1].height = 75
    result_sheet.freeze_panes = "A2"
    audit.append(["Source row", "Document ID", "Workspace URL", "Eligibility", "Check status", "Flawless submission", "Matching Assigned by", "Reason", "Evidence", "Phase", "Version", "Document Submitted"])
    counts, groups, group_labels = Counter(), {}, {}
    url_column = headers.index("Document WS URL")
    manager_column = headers.index("Author Manager")
    rows = sorted(sheet.iter_rows(min_row=4), key=lambda cells: (
        not bool(normalized(cells[manager_column].value)), normalized(cells[manager_column].value)))
    pending, ready = {}, []
    target_row, audit_row = 1, 1
    for cells in rows:
        if not any(c.value is not None for c in cells):
            continue
        row = dict(zip(headers, [c.value for c in cells]))
        decision = eligibility(row, config, book.epoch, now, reporting_month)
        ws_cell = cells[url_column]
        url = ws_cell.hyperlink.target if ws_cell.hyperlink else ""
        doc_id, canonical, link_error = "", "", ""
        try:
            doc_id, canonical = workspace_identity(url)
        except ValueError as exc:
            link_error = str(exc)
        audit_row += 1
        if decision.status == "Qualified":
            target_row += 1
            manager = row.get("Author Manager")
            key = normalized(manager)
            group_labels.setdefault(key, " ".join(str("" if manager is None else manager).split()) or "(Not specified)")
        entry = (cells, row, decision, doc_id, canonical, link_error, target_row, audit_row)
        if decision.status == "Qualified" and not link_error:
            pending.setdefault(doc_id, []).append(entry)
        else:
            ready.append(entry)

    def resolve_document(entries):
        doc_id, canonical = entries[0][3:5]
        if reader is not None:
            from .site import capture
            try:
                capture(reader, canonical, evidence_dir)
            except (OSError, ValueError) as exc:
                return [], False, f"Live capture failed: {exc}", ""
        evidence_now = now if fixed_time else datetime.now(timezone.utc)
        return load_evidence(evidence_dir, doc_id, config, evidence_now)

    def completed_rows():
        with ThreadPoolExecutor(max_workers=4, thread_name_prefix="xccb-io") as executor:
            futures = {executor.submit(resolve_document, entries): entries for entries in pending.values()}
            for entry in ready:
                yield entry, None
            for future in as_completed(futures):
                evidence = future.result()
                for entry in futures[future]:
                    yield entry, evidence

    for entry, evidence in completed_rows():
        cells, row, decision, doc_id, canonical, link_error, target_row, audit_row = entry
        ws_cell = cells[url_column]
        url = ws_cell.hyperlink.target if ws_cell.hyperlink else ""
        value, matches, reference = "", [], ""
        reasons = list(decision.reasons)
        check = "Not checked"
        progress_group = None
        if decision.status == "Qualified":
            if link_error:
                check = "Unverified"
            else:
                actions, complete, evidence_reason, reference = evidence
                value, matches, reason = classify(actions, config["assigned_by_names"], complete=complete, scope=config["action_scope"])
                check = "Verified" if value else "Unverified"
                reasons.extend([reason, evidence_reason])
            for column, cell in enumerate(cells, 1):
                target = result_sheet.cell(target_row, column, cell.value)
                target.data_type = cell.data_type
                copy_style(cell, target)
                if cell.hyperlink:
                    target.hyperlink = copy(cell.hyperlink)
                    target.hyperlink.ref = target.coordinate
                if cell.comment:
                    target.comment = copy(cell.comment)
            result_sheet.row_dimensions[target_row].height = sheet.row_dimensions[cells[0].row].height
            result_sheet.cell(target_row, 44, value or None)
            counts[value or "Unresolved Actions"] += 1
            key = normalized(row.get("Author Manager"))
            group = groups.setdefault(key, Counter())
            group["Qualified"] += 1
            group[value or "Unresolved Actions"] += 1
            progress_group = (group_labels[key], group["Qualified"], group["Yes"], group["No"], group["Unresolved Actions"])
        counts[decision.status] += 1
        if link_error:
            reasons.append(link_error)
        values = [cells[0].row, doc_id, canonical or url, decision.status, check, value, "; ".join(matches), "; ".join(dict.fromkeys(reasons)), reference, row["Phase"], row.get("Version"), row["Document Submitted"]]
        for column, val in enumerate(values, 1):
            cell = audit.cell(audit_row, column, val)
            if isinstance(val, str):
                cell.data_type = "s"  # Audit text must never become executable Excel formulas.
        if canonical:
            audit.cell(audit_row, 3).hyperlink = canonical
        audit.row_dimensions[audit_row].height = 64
        logging.debug("Source row %s: %s / %s", cells[0].row, decision.status, check)
        if on_progress is not None:
            document = None
            if decision.status == "Qualified":
                document = {"result": value, "document": doc_id, "title": row.get("CCB Title"),
                            "manager": row.get("Author Manager"), "phase": row.get("Phase"),
                            "submitted": row.get("Document Submitted"), "url": url or canonical,
                            "matches": values[6], "reasons": values[7]}
            on_progress(dict(counts), progress_group, document)
    for key in sorted(groups, key=lambda value: (not bool(value), value)):
        group = groups[key]
        managers.append([group_labels[key], group["Qualified"], group["Yes"], group["No"], group["Unresolved Actions"]])
        managers.cell(managers.max_row, 1).data_type = "s"
    pending = (counts["Unverified"] or counts["Unresolved Actions"]
               or not config["assigned_by_names"] or config["action_scope"] != "document")
    summary_rows = [["Metric", "Value"], ["Report status", "DRAFT - unresolved settings, eligibility or Actions" if pending else "Complete for configured rules"],
                    ["Business timezone", config["timezone"]], ["Month start inclusive", str(start)], ["Next month exclusive", str(end)],
                    ["Generated UTC", now.isoformat()], ["Source workbook", str(source)], ["Yes", counts["Yes"]], ["No", counts["No"]],
                    ["Unresolved Actions", counts["Unresolved Actions"]], ["Unresolved eligibility", counts["Unverified"]],
                    ["Qualified rows", counts["Qualified"]], ["Excluded rows", counts["Excluded"]],
                    ["Actions scope", config["action_scope"]], ["Phase rule", PHASE_RULE],
                    ["Reviewer rule", REVIEWER_RULE], ["Allowed reviewer names", "; ".join(config.get("reviewer_names", []))],
                    ["Grouping", "Author Manager; blank values retained"],
                    ["Assigned by list configured", "Yes" if config["assigned_by_names"] else "No"]]
    for values in summary_rows:
        summary.append(values)
        summary.row_dimensions[summary.max_row].height = 32
    for tab in (audit, summary, managers):
        tab.freeze_panes = "A2"
        tab.auto_filter.ref = tab.dimensions
        tab.row_dimensions[1].height = 32
        for c in tab[1]:
            c.fill = PatternFill("solid", fgColor="17365D")
            c.font = Font(color="FFFFFF", bold=True)
        for row_cells in tab:
            for c in row_cells:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    for column in range(1, 13):
        audit.column_dimensions[get_column_letter(column)].width = {1: 12, 2: 16, 3: 65, 7: 34, 8: 95, 9: 65}.get(column, 24)
    summary.column_dimensions["A"].width = 30
    summary.column_dimensions["B"].width = 100
    managers.column_dimensions["A"].width = 40
    for column in "BCDE":
        managers.column_dimensions[column].width = 24
    result_sheet.auto_filter.ref = f"A1:AR{result_sheet.max_row}"
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = None
    try:
        with NamedTemporaryFile(dir=output.parent, prefix=f".{output.stem}_", suffix=".tmp.xlsx", delete=False) as stream:
            temp = Path(stream.name)
            out.save(stream)
        # Windows rename fails if another job created the destination meanwhile.
        temp.rename(output)
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)
        out.close()
        book.close()
    return dict(counts)
