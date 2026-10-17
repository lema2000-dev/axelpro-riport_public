"""Write import results and document identities into annual reports."""

from datetime import date, datetime

from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.worksheet import Worksheet

from .excel_constants import (
    DAILY_TARGET_ROWS,
    MONTHLY_TARGET_ROWS,
    MONTH_NAMES,
    PROCESSED_DOCUMENTS_SHEET,
)
from .models import DailyTotals, MonthlyTotals, ImportResult

def _find_day_column(
    sheet: Worksheet,
    target_date: date,
) -> int:
    """Find day column."""
    for cell in sheet[2]:
        value = cell.value

        if isinstance(value, datetime):
            value = value.date()

        if value == target_date:
            return cell.column

def _add_cell_amount(
    sheet: Worksheet,
    row: int,
    column: int,
    amount: int,
) -> None:
    """Add cell amount."""
    cell = sheet.cell(row=row, column=column)

    current_value = cell.value

    if current_value is None:
        current_value = 0

    cell.value = current_value + amount

def _write_daily_totals(
    sheet: Worksheet,
    totals: DailyTotals,
) -> None:
    """Write daily totals."""
    column = _find_day_column(
        sheet=sheet,
        target_date=totals.movement_date,
    )

    for field_name, row in DAILY_TARGET_ROWS.items():
        amount = getattr(totals, field_name)

        _add_cell_amount(
            sheet=sheet,
            row=row,
            column=column,
            amount=amount,
        )

def _write_monthly_totals(
    sheet: Worksheet,
    totals: MonthlyTotals,
) -> None:
    """Write monthly totals."""
    for field_name, row in MONTHLY_TARGET_ROWS.items():
        amount = getattr(totals, field_name)

        _add_cell_amount(
            sheet=sheet,
            row=row,
            column=2,
            amount=amount,
        )

def _write_processed_documents(
    workbook: WorkbookType,
    year: int,
    new_documents: set[tuple[int, str]],
) -> None:
    """Write processed documents."""
    sheet = workbook[PROCESSED_DOCUMENTS_SHEET]

    for document_year, document_number in sorted(new_documents):
        if document_year == year:
            sheet.append([document_year, document_number])

def write_import_result(
    workbook: WorkbookType,
    year: int,
    result: ImportResult,
) -> WorkbookType:
    """Apply annual totals and new document identities, returning the workbook."""
    for totals in result.daily_totals:
        if totals.movement_date.year == year:
            month = totals.movement_date.month
            sheet = workbook[MONTH_NAMES[month - 1]]

            _write_daily_totals(
                sheet=sheet,
                totals=totals,
            )

    for totals in result.monthly_totals:
        if totals.year == year:
            sheet = workbook[MONTH_NAMES[totals.month - 1]]

            _write_monthly_totals(
                sheet=sheet,
                totals=totals,
            )

    _write_processed_documents(
        workbook=workbook,
        year=year,
        new_documents=result.new_documents,
    )

    return workbook