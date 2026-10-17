from datetime import date
from sales_report.excel_create_workbook import create_year_workbook
from sales_report.excel_writer import write_import_result
from sales_report.models import DailyTotals, MonthlyTotals, ImportResult
from sales_report.excel_constants import PROCESSED_DOCUMENTS_SHEET


def test_adds_only_selected_year_and_preserves_manual_data_and_formulas(
    template_sheet, tmp_path
):
    from openpyxl import load_workbook

    wb = create_year_workbook(template_sheet, 2026)
    path = tmp_path / "report.xlsx"
    wb.save(path)
    wb = load_workbook(path)
    sheet = wb["January"]
    sheet["J8"] = 50
    sheet["J7"] = 1234
    formula = sheet["J12"].value
    result = ImportResult(
        [
            DailyTotals(date(2026, 1, 5), *range(1, 9)),
            DailyTotals(date(2027, 1, 5), *([900] * 8)),
        ],
        [MonthlyTotals(2026, 1, 10, 20, 30), MonthlyTotals(2027, 1, 900, 900, 900)],
        {(2026, "NEW"), (2027, "OTHER")},
    )
    assert write_import_result(wb, 2026, result) is wb
    assert sheet["J8"].value == 51 and sheet["J9"].value == 2
    assert sheet["J40"].value == 8 and sheet["B82"].value == 10
    assert sheet["B83"].value == 20 and sheet["B84"].value == 30
    assert sheet["J7"].value == 1234 and sheet["J12"].value == formula
    assert list(wb[PROCESSED_DOCUMENTS_SHEET].values)[1:] == [(2026, "NEW")]
