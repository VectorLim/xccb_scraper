"""Source workbook reading and format-preserving report output."""

from copy import copy
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter

from .filtering import normalized

REQUIRED = {
    "Author Manager",
    "Reviewing Forum",
    "Document WS URL",
}
AUDIT_HEADERS = [
    "Source row",
    "Document ID",
    "Workspace URL",
    "Eligibility",
    "Check status",
    "Flawless submission",
    "Matching Assigned by",
    "Reason",
    "Evidence",
    "Phase",
    "Version",
    "Document Submitted",
    "Author Manager",
    "Organization",
    "Organization code",
    "WWID",
    "Organization lookup UTC",
    "Organization status",
]
MANAGER_HEADERS = [
    "Author Manager",
    "Organization",
    "Qualified rows",
    "Yes",
    "No",
    "Unresolved Actions",
]


def read_source(path):
    workbook = load_workbook(path)
    sheets = [s for s in workbook if REQUIRED <= {c.value for c in s[3]}]
    if len(sheets) != 1:
        raise ValueError("Expected exactly one report sheet with headers on row 3")
    sheet = sheets[0]
    headers = [c.value for c in sheet[3]]
    if len(set(headers)) != len(headers) or any(not h for h in headers):
        raise ValueError("Expected unique nonblank source column headers")
    return workbook, sheet, headers


def copy_style(source, target):
    # Style indices belong to a workbook; register their components in the new one.
    for attribute in ("font", "fill", "border", "alignment", "protection"):
        setattr(target, attribute, copy(getattr(source, attribute)))
    target.number_format = source.number_format


class ReportWriter:
    """Builds the output workbook: result copy of the source sheet plus Summary, Audit and manager tabs."""

    def __init__(self, sheet, epoch):
        self.source = sheet
        self.book = Workbook()
        self.book.epoch = epoch
        self.results = self.book.active
        self.results.title = "Flawless submissions"
        self.summary = self.book.create_sheet("Summary")
        self.audit = self.book.create_sheet("Audit")
        self.managers = self.book.create_sheet("By Author Manager")
        self.managers.append(MANAGER_HEADERS)
        title_column = next(
            (cell.column for cell in sheet[3] if cell.value == "CCB Title"),
            None,
        )
        self.columns = [
            (cell.column, cell.value) for cell in sheet[3]
            if cell.column != title_column
        ]
        phase = next(
            (index for index, (_, name) in enumerate(self.columns) if name == "Phase"),
            next((index - 1 for index, (_, name) in enumerate(self.columns) if name == "Document Submitted"), len(self.columns) - 1),
        )
        self.columns.insert(phase + 1, (title_column, "Document"))
        for column, (source_column, name) in enumerate(self.columns, 1):
            cell = sheet.cell(3, source_column) if source_column else sheet.cell(3, 1)
            target = self.results.cell(1, column, name)
            copy_style(cell, target)
            target.alignment = Alignment(wrap_text=True, vertical="top")
        result_column = len(self.columns) + 1
        self.results.cell(1, result_column, "Flawless submission")
        self.results.cell(1, result_column)._style = copy(self.results.cell(1, result_column - 1)._style)
        self.results.row_dimensions[1].height = 75
        self.results.freeze_panes = "A2"
        self.audit.append(AUDIT_HEADERS)

    def add_submission(self, target_row, cells, value, document_title=""):
        for column, (source_column, name) in enumerate(self.columns, 1):
            if name == "Document":
                target = self.results.cell(target_row, column, document_title or None)
                target.data_type = "s"
                if source_column is not None:
                    copy_style(cells[source_column - 1], target)
                continue
            if source_column is None:
                continue
            cell = cells[source_column - 1]
            target = self.results.cell(target_row, column, cell.value)
            target.data_type = cell.data_type
            copy_style(cell, target)
            if cell.hyperlink:
                target.hyperlink = copy(cell.hyperlink)
                target.hyperlink.ref = target.coordinate
            if cell.comment:
                target.comment = copy(cell.comment)
        self.results.row_dimensions[target_row].height = self.source.row_dimensions[
            cells[0].row
        ].height
        self.results.cell(target_row, len(self.columns) + 1, value or None)

    def add_audit(self, audit_row, values, canonical):
        for column, val in enumerate(values, 1):
            cell = self.audit.cell(audit_row, column, val)
            if isinstance(val, str):
                cell.data_type = (
                    "s"  # Audit text must never become executable Excel formulas.
                )
        if canonical:
            self.audit.cell(audit_row, 3).hyperlink = canonical
        self.audit.row_dimensions[audit_row].height = 64

    def add_manager(self, values):
        self.managers.append(values)
        self.managers.cell(self.managers.max_row, 1).data_type = "s"
        self.managers.cell(self.managers.max_row, 2).data_type = "s"

    def add_summary(self, rows):
        for values in rows:
            self.summary.append(values)
            self.summary.row_dimensions[self.summary.max_row].height = 32

    def finish(self, manager_index, organizations):
        """Apply tab styling and insert the Organization column after Author Manager."""
        for tab in (self.audit, self.summary, self.managers):
            tab.freeze_panes = "A2"
            tab.auto_filter.ref = tab.dimensions
            tab.row_dimensions[1].height = 32
            for c in tab[1]:
                c.fill = PatternFill("solid", fgColor="17365D")
                c.font = Font(color="FFFFFF", bold=True)
            for row_cells in tab:
                for c in row_cells:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        for column in range(1, 19):
            self.audit.column_dimensions[get_column_letter(column)].width = {
                1: 12,
                2: 16,
                3: 65,
                7: 34,
                8: 95,
                9: 65,
            }.get(column, 24)
        self.summary.column_dimensions["A"].width = 30
        self.summary.column_dimensions["B"].width = 100
        self.managers.column_dimensions["A"].width = 40
        for column in "BCDEF":
            self.managers.column_dimensions[column].width = 24
        results = self.results
        organization_column = next(
            column + 1 for column, (source_column, _) in enumerate(self.columns, 1)
            if source_column == manager_index + 1
        )
        results.insert_cols(organization_column)
        column_positions = {
            source_column: column + (column >= organization_column)
            for column, (source_column, _) in enumerate(self.columns, 1)
            if source_column is not None
        }
        results.column_dimensions.clear()
        for key, dimension in self.source.column_dimensions.items():
            index = column_index_from_string(key)
            target_index = column_positions.get(index)
            if target_index is None:
                continue
            target = results.column_dimensions[get_column_letter(target_index)]
            for attribute in ("width", "hidden", "bestFit", "outlineLevel", "collapsed"):
                setattr(target, attribute, getattr(dimension, attribute))
            target.min = (
                column_positions.get(dimension.min, target_index)
                if dimension.min
                else target_index
            )
            target.max = (
                column_positions.get(dimension.max, target_index)
                if dimension.max
                else target_index
            )
            copy_style(dimension, target)
        results.cell(1, organization_column, "Organization")
        copy_style(
            results.cell(1, organization_column - 1),
            results.cell(1, organization_column),
        )
        results.column_dimensions[get_column_letter(organization_column)].width = 42
        for row_number in range(2, results.max_row + 1):
            manager = results.cell(row_number, organization_column - 1).value
            organization = organizations.get(normalized(manager))
            cell = results.cell(
                row_number,
                organization_column,
                organization.organization if organization else "Unresolved",
            )
            cell.data_type = "s"
        for row_cells in results:
            for cell in row_cells:
                if cell.hyperlink:
                    cell.hyperlink.ref = cell.coordinate
        document_column = next(cell.column for cell in results[1] if cell.value == "Document")
        results.column_dimensions[get_column_letter(document_column)].width = 60
        last_column = get_column_letter(results.max_column)
        results.column_dimensions[last_column].width = 24
        results.auto_filter.ref = f"A1:{last_column}{results.max_row}"

    def save(self, output):
        output = Path(output)
        output.parent.mkdir(parents=True, exist_ok=True)
        temp = None
        try:
            with NamedTemporaryFile(
                dir=output.parent,
                prefix=f".{output.stem}_",
                suffix=".tmp.xlsx",
                delete=False,
            ) as stream:
                temp = Path(stream.name)
                self.book.save(stream)
            # Windows rename fails if another job created the destination meanwhile.
            temp.rename(output)
        finally:
            if temp is not None:
                temp.unlink(missing_ok=True)
            self.book.close()
