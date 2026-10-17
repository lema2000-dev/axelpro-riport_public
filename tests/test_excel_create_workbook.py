from calendar import monthrange
from datetime import date
import pytest
from sales_report.excel_create_workbook import create_year_workbook
from sales_report.excel_constants import (
    MONTH_NAMES,
    DAILY_EXCHANGE_RATES_SHEET,
    PROCESSED_DOCUMENTS_SHEET,
)


@pytest.mark.parametrize("year", [2024, 2026, 2027])
def test_full_year_dates_hidden_registers_and_summary(template_sheet, year):
    wb = create_year_workbook(template_sheet, year)
    assert set(wb.sheetnames) == set(MONTH_NAMES) | {
        "Exchange_Rates",
        "Summary",
        DAILY_EXCHANGE_RATES_SHEET,
        PROCESSED_DOCUMENTS_SHEET,
    }
    for month, name in enumerate(MONTH_NAMES, 1):
        sheet = wb[name]
        assert sheet.freeze_panes == "D4"
        days = [c.value for c in sheet[2] if isinstance(c.value, date)]
        expected = [
            date(year, month, n)
            for n in range(1, monthrange(year, month)[1] + 1)
            if date(year, month, n).weekday() < 6
        ]
        assert days == expected
        for cell in sheet[2]:
            if isinstance(cell.value, date):
                assert sheet.cell(8, cell.column).value is None
        assert sheet["B82"].value is None
        assert sheet["B49"].value == "=B46+B47+B48"
        assert f"'Exchange_Rates'!$B${month + 1}" in sheet["B60"].value
        assert wb["Summary"].cell(2, month + 1).value == name
        assert f"'{name}'!B82" in wb["Summary"].cell(40, month + 1).value
        assert wb["Exchange_Rates"].cell(month + 1, 2).value == 0
        assert wb["Exchange_Rates"].cell(month + 1, 3).value is None
    for name in (DAILY_EXCHANGE_RATES_SHEET, PROCESSED_DOCUMENTS_SHEET):
        assert wb[name].sheet_state == "veryHidden"
        assert wb[name].max_row == 1
    assert template_sheet["B8"].value == 999


def test_january_fifth_is_in_second_week_not_left_padding(template_sheet):
    wb = create_year_workbook(template_sheet, 2026)
    jan = wb["January"]
    assert jan["B2"].value == date(2026, 1, 1)
    assert jan["J2"].value == date(2026, 1, 5)
    assert jan["H2"].value == 1
    assert jan["H12"].value == "=B12+D12+F12"


def test_generated_months_leave_removed_private_formula_blank(year_report):
    for name in MONTH_NAMES:
        sheet = year_report[name]
        assert sheet["B85"].value is None
        assert sheet.freeze_panes == "D4"


def test_freeze_panes_survive_save_and_reload(year_report):
    # The year_report fixture provides a workbook saved and loaded again.
    for name in MONTH_NAMES:
        assert year_report[name].freeze_panes == "D4"
    for name in ("Exchange_Rates", "Summary", DAILY_EXCHANGE_RATES_SHEET, PROCESSED_DOCUMENTS_SHEET):
        assert year_report[name].freeze_panes is None
