"""Report pipeline: eligibility, organization lookup, Actions evidence and classification."""

from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
import logging
from pathlib import Path

from .actions import workspace_identity
from .evidence import load_document_title, load_evidence
from .organizations import OrganizationCache, directory_query, load_settings
from .filtering import (
    PHASE_RULE,
    REVIEWER_RULE,
    classify,
    eligibility,
    month_bounds,
    normalized,
    organization_eligibility,
)
from .workbook import ReportWriter, read_source


def _organization_status(organization, issue):
    if issue is not None:
        return issue
    if organization and organization.matched_name:
        return "Unique exact HR full-name match"
    if organization:
        return "Unique active HR name candidate; not manually confirmed"
    return "Identity unresolved: blank manager"


def report(
    source,
    output,
    config,
    evidence_dir,
    now=None,
    reader=None,
    *,
    on_progress=None,
    reporting_month=None,
    organization_cache=None,
    organization_query=None,
    organization_codes=None,
    refresh_organizations=False,
    on_organization_progress=None,
    on_directory_unavailable=None,
    cancel=None,
    all_dates=False,
):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.exists():
        raise ValueError(
            "Choose a new output file; existing workbooks are never overwritten"
        )
    fixed_time = now is not None
    now = now or datetime.now(timezone.utc)
    start, end = (None, None) if all_dates else month_bounds(now, config["timezone"], reporting_month)
    settings = load_settings()
    cache = (
        organization_cache if organization_cache is not None else OrganizationCache()
    )
    codes = [] if organization_codes is None else organization_codes
    batch_query = directory_query(
        settings.get("organization_python"), on_directory_unavailable, cancel
    )
    book, sheet, headers = read_source(source)
    try:
        manager_index = headers.index("Author Manager")
        url_column = headers.index("Document WS URL")
        source_rows = list(sheet.iter_rows(min_row=4))
        decisions = {
            cells[0].row: eligibility(
                dict(zip(headers, [cell.value for cell in cells])),
                config,
                book.epoch,
                now,
                reporting_month,
                **({"all_dates": True} if all_dates else {}),
            )
            for cells in source_rows
            if any(cell.value is not None for cell in cells)
        }
        organizations, organization_issues = cache.resolve(
            [
                cells[manager_index].value
                for cells in source_rows
                if cells[0].row in decisions
                and decisions[cells[0].row].status == "Qualified"
            ],
            organization_query,
            batch_query=None if organization_query is not None else batch_query,
            refresh=refresh_organizations,
            progress=on_organization_progress,
            cancel=cancel,
        )
        if organization_issues and on_organization_progress:
            on_organization_progress(
                "Organization lookup warning: "
                + next(iter(organization_issues.values()))
            )

        writer = ReportWriter(sheet, book.epoch)
        counts, groups, group_labels = Counter(), {}, {}
        rows = sorted(
            source_rows,
            key=lambda cells: (
                not bool(normalized(cells[manager_index].value)),
                normalized(cells[manager_index].value),
            ),
        )
        # Output row numbers are fixed in source order before documents complete out of order.
        pending, ready = {}, []
        target_row, audit_row = 1, 1
        for cells in rows:
            if not any(c.value is not None for c in cells):
                continue
            row = dict(zip(headers, [c.value for c in cells]))
            decision = decisions[cells[0].row]
            organization = organizations.get(normalized(row.get("Author Manager")))
            decision = organization_eligibility(decision, organization, codes)
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
                group_labels.setdefault(
                    normalized(manager),
                    " ".join(str("" if manager is None else manager).split())
                    or "(Not specified)",
                )
            entry = (
                cells,
                row,
                decision,
                url,
                doc_id,
                canonical,
                link_error,
                target_row,
                audit_row,
            )
            if decision.status == "Qualified" and not link_error:
                pending.setdefault(doc_id, []).append(entry)
            else:
                ready.append(entry)

        def resolve_document(entries):
            doc_id, canonical = entries[0][4:6]
            if reader is not None:
                from .site import capture

                try:
                    capture(reader, canonical, evidence_dir)
                except (OSError, ValueError) as exc:
                    return ([], False, f"Live capture failed: {exc}", ""), ""
            evidence_now = now if fixed_time else datetime.now(timezone.utc)
            return (
                load_evidence(evidence_dir, doc_id, config, evidence_now),
                load_document_title(evidence_dir, doc_id, config, evidence_now),
            )

        def completed_rows():
            with ThreadPoolExecutor(
                max_workers=4, thread_name_prefix="xccb-io"
            ) as executor:
                futures = {
                    executor.submit(resolve_document, entries): entries
                    for entries in pending.values()
                }
                for entry in ready:
                    yield entry, None
                for future in as_completed(futures):
                    evidence = future.result()
                    for entry in futures[future]:
                        yield entry, evidence

        for entry, evidence in completed_rows():
            (
                cells,
                row,
                decision,
                url,
                doc_id,
                canonical,
                link_error,
                target_row,
                audit_row,
            ) = entry
            value, matches, reference = "", [], ""
            document_title = ""
            reasons = list(decision.reasons)
            check = "Not checked"
            progress_group = None
            key = normalized(row.get("Author Manager"))
            organization = organizations.get(key)
            if decision.status == "Qualified":
                if link_error:
                    check = "Unverified"
                else:
                    action_evidence, document_title = evidence
                    actions, complete, evidence_reason, reference = action_evidence
                    value, matches, reason = classify(
                        actions,
                        config["assigned_by_names"],
                        complete=complete,
                        scope=config["action_scope"],
                    )
                    check = "Verified" if value else "Unverified"
                    reasons.extend([reason, evidence_reason])
                if not document_title:
                    reasons.append("Document title unavailable from xCCB workspace")
                writer.add_submission(target_row, cells, value, document_title)
                counts[value or "Unresolved Actions"] += 1
                group = groups.setdefault(key, Counter())
                group["Qualified"] += 1
                group[value or "Unresolved Actions"] += 1
                progress_group = (
                    group_labels[key],
                    organization.organization if organization else "Unresolved",
                    group["Qualified"],
                    group["Yes"],
                    group["No"],
                    group["Unresolved Actions"],
                )
            counts[decision.status] += 1
            if link_error:
                reasons.append(link_error)
            organization_status = _organization_status(
                organization, organization_issues.get(key)
            )
            not_looked_up = (
                organization is None
                and key not in organization_issues
                and decision.status != "Qualified"
            )
            if not_looked_up:
                organization_status = "Not looked up: submission not eligible"
            values = [
                cells[0].row,
                doc_id,
                canonical or url,
                decision.status,
                check,
                value,
                "; ".join(matches),
                "; ".join(dict.fromkeys(reasons)),
                reference,
                row.get("Phase"),
                row.get("Version"),
                row.get("Document Submitted"),
                row.get("Author Manager"),
                (
                    organization.organization
                    if organization
                    else "" if not_looked_up else "Unresolved"
                ),
                organization.organization_code if organization else "",
                organization.wwid if organization else "",
                organization.looked_up_utc if organization else "",
                organization_status,
            ]
            writer.add_audit(audit_row, values, canonical)
            logging.debug(
                "Source row %s: %s / %s", cells[0].row, decision.status, check
            )
            if on_progress is not None:
                document = None
                if decision.status == "Qualified":
                    document = {
                        "result": value,
                        "document": doc_id,
                        "title": document_title,
                        "manager": row.get("Author Manager"),
                        "phase": row.get("Phase"),
                        "organization": (
                            organization.organization if organization else "Unresolved"
                        ),
                        "submitted": row.get("Document Submitted"),
                        "url": url or canonical,
                        "matches": values[6],
                        "reasons": values[7],
                    }
                on_progress(dict(counts), progress_group, document)

        for key in sorted(groups, key=lambda value: (not bool(value), value)):
            group = groups[key]
            organization = organizations.get(key)
            writer.add_manager(
                [
                    group_labels[key],
                    organization.organization if organization else "Unresolved",
                    group["Qualified"],
                    group["Yes"],
                    group["No"],
                    group["Unresolved Actions"],
                ]
            )
        draft = (
            counts["Unverified"]
            or counts["Unresolved Actions"]
            or organization_issues
            or not config["assigned_by_names"]
            or config["action_scope"] != "document"
        )
        writer.add_summary(
            [
                ["Metric", "Value"],
                [
                    "Report status",
                    (
                        "DRAFT - unresolved settings, eligibility or Actions"
                        if draft
                        else "Complete for configured rules"
                    ),
                ],
                ["Business timezone", config["timezone"]],
                ["Month start inclusive", str(start) if start else "All dates"],
                ["Next month exclusive", str(end) if end else "All dates"],
                ["Generated UTC", now.isoformat()],
                ["Source workbook", str(source)],
                ["Yes", counts["Yes"]],
                ["No", counts["No"]],
                ["Unresolved Actions", counts["Unresolved Actions"]],
                ["Unresolved eligibility", counts["Unverified"]],
                ["Qualified rows", counts["Qualified"]],
                ["Excluded rows", counts["Excluded"]],
                ["Actions scope", config["action_scope"]],
                ["Phase rule", PHASE_RULE],
                ["Reviewer rule", REVIEWER_RULE],
                ["Allowed reviewer names", "; ".join(config.get("reviewer_names", []))],
                ["Grouping", "Author Manager; blank values retained"],
                [
                    "Selected organization codes",
                    "; ".join(codes) or "All organizations",
                ],
                [
                    "Organization basis",
                    "Current active HR name candidates, not historical membership; no inferred KuAT hierarchy",
                ],
                ["Organization lookup issues", len(organization_issues)],
                [
                    "Assigned by list configured",
                    "Yes" if config["assigned_by_names"] else "No",
                ],
            ]
        )
        writer.finish(manager_index, organizations)
        writer.save(output)
    finally:
        book.close()
    return dict(counts)
