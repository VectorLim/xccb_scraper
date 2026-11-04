"""Display rows for generated reports, independent of the Tk dashboard."""

from datetime import date, datetime

from openpyxl import load_workbook

SHORT_NOTES = (
    ("Live capture failed: Stopped by user", "Stopped"),
    ("Live capture failed: Document deleted", "Document deleted on xCCB"),
    ("Live capture failed: Document not found", "Document not found on xCCB"),
    ("Live capture failed", "Capture failed"),
    ("Invalid xCCB", "Invalid link"),
    ("Missing or ambiguous DocumentID", "Invalid link"),
    ("Actions scope", "Scope not confirmed"),
    ("Assigned by name list", "No listed names"),
    ("Actions completeness", "Actions not verified complete"),
    ("All document Actions", "No listed assigner"),
)


def short_note(reasons):
    parts = (reasons or "").split("; ")
    # Earlier SHORT_NOTES entries are root causes, so they win over generic follow-on reasons.
    for prefix, note in SHORT_NOTES:
        if any(part.startswith(prefix) for part in parts):
            return note
    first = parts[0]
    return first if len(first) <= 40 else first[:39] + "…"


def document_row(
    result,
    document,
    title,
    manager,
    phase,
    submitted,
    url,
    matches,
    reasons,
    organization="Unresolved",
):
    if isinstance(submitted, (date, datetime)):
        submitted = submitted.strftime("%Y-%m-%d")
    result = result or ""
    notes = (
        f"Assigned by: {matches}" if result == "No" and matches else short_note(reasons)
    )
    return tuple(
        "" if v is None else " ".join(str(v).split())
        for v in (
            result,
            document,
            manager,
            organization,
            phase,
            title,
            submitted,
            url,
            notes,
        )
    )


def read_results(path):
    """Return summary status, manager groups and per-submission rows from a generated report."""
    book = load_workbook(path)
    try:
        status = book["Summary"]["B2"].value
        groups = list(book["By Author Manager"].iter_rows(min_row=2, values_only=True))
        # Audit and result rows are written in the same order, so qualified audit rows align 1:1.
        audit = [
            row
            for row in book["Audit"].iter_rows(min_row=2, values_only=True)
            if row[3] == "Qualified"
        ]
        sheet = book["Flawless submissions"]
        index = {cell.value: position for position, cell in enumerate(sheet[1])}
        documents = []
        for cells, checked in zip(sheet.iter_rows(min_row=2), audit):
            get = lambda name: cells[index[name]].value if name in index else None
            link = cells[index["Document WS URL"]]
            url = (
                link.hyperlink.target
                if link.hyperlink and link.hyperlink.target
                else (checked[2] or "")
            )
            documents.append(
                document_row(
                    checked[5],
                    checked[1],
                    get("Document") if "Document" in index else get("CCB Title"),
                    get("Author Manager"),
                    get("Phase"),
                    get("Document Submitted"),
                    url,
                    checked[6],
                    checked[7],
                    get("Organization"),
                )
            )
        return status, groups, documents
    finally:
        book.close()
