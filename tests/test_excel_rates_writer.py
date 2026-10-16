from datetime import date, datetime
from decimal import Decimal
from axelpro_riport import excel_rates_writer as writer
from axelpro_riport.excel_create_workbook import create_year_workbook
from axelpro_riport.excel_constants import DAILY_EXCHANGE_RATES_SHEET
from axelpro_riport.models import DailyExchangeRate


class FixedDate(date):
    @classmethod
    def today(cls):
        return date(2026, 10, 11)


def test_rates_once_per_day_current_month_only(monkeypatch, template_sheet):
    monkeypatch.setattr(writer, "date", FixedDate)
    wb = create_year_workbook(template_sheet, 2026)
    sheet = wb[DAILY_EXCHANGE_RATES_SHEET]
    sheet.append([datetime(2026, 9, 1), datetime(2026, 9, 1), 390])
    wb["Euro"]["B10"] = 390
    rates = {
        date(2026, 9, 1): DailyExchangeRate(
            date(2026, 9, 1), date(2026, 9, 1), Decimal(390)
        ),
        date(2026, 10, 11): DailyExchangeRate(
            date(2026, 10, 11), date(2026, 10, 9), Decimal("401.25")
        ),
        date(2027, 1, 1): DailyExchangeRate(
            date(2027, 1, 1), date(2026, 10, 9), Decimal("401.25")
        ),
    }
    assert writer.write_exchange_rates(wb, 2026, rates) is wb
    writer.write_exchange_rates(wb, 2026, rates)
    assert sheet.max_row == 3
    assert sheet["C3"].value == 401.25
    assert wb["Euro"]["B10"].value == 390
    assert wb["Euro"]["B11"].value == 401.25
    assert wb["Euro"]["C11"].value == date(2026, 10, 9)
    assert sheet["A3"].number_format == "yyyy-mm-dd"


def test_historical_year_does_not_require_today(monkeypatch, template_sheet):
    monkeypatch.setattr(writer, "date", FixedDate)
    wb = create_year_workbook(template_sheet, 2025)
    writer.write_exchange_rates(wb, 2025, {})
    assert all(wb["Euro"].cell(row, 2).value == 0 for row in range(2, 14))
