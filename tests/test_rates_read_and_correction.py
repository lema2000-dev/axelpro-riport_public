from datetime import date
from decimal import Decimal
from urllib.error import URLError
import pytest
from openpyxl import Workbook
from axelpro_riport import rates_read_and_correction as rates
from axelpro_riport.models import DailyExchangeRate
from axelpro_riport.excel_constants import DAILY_EXCHANGE_RATES_SHEET


class FixedDate(date):
    @classmethod
    def today(cls):
        return date(2026, 10, 11)


@pytest.fixture(autouse=True)
def fixed_today(monkeypatch):
    monkeypatch.setattr(rates, "date", FixedDate)


def store(path, records):
    wb = Workbook()
    wb.active.title = DAILY_EXCHANGE_RATES_SHEET
    wb.active.append(["Application", "Source", "Rate"])
    for record in records:
        wb.active.append([record.application_date, record.source_date, float(record.rate)])
    return wb, path


def test_reads_years_preserves_existing_today_and_uses_previous_source(
    monkeypatch, tmp_path, movement
):
    today = date(2026, 10, 11)
    old = DailyExchangeRate(today, date(2026, 10, 9), Decimal(400))
    workbooks = {2026: store(tmp_path / "report.xlsx", [old]),
                 2027: store(tmp_path / "report2027.xlsx", [])}
    monkeypatch.setattr(
        rates,
        "fetch_current_rate",
        lambda **kwargs: DailyExchangeRate(today, date(2026, 10, 9), Decimal(410)),
    )
    rows = [
        movement(day=today, currency="EUR"),
        movement(day=date(2026, 10, 10), currency="EUR"),
        movement(day=date(2027, 1, 4), currency="EUR"),
        movement(day=date(2026, 1, 1)),
    ]
    result = rates.get_exchange_rates(rows, workbooks)
    assert result[today] == old
    assert result[date(2026, 10, 10)].rate == Decimal(400)
    assert result[date(2027, 1, 4)].source_date == date(2026, 10, 9)
    assert date(2026, 1, 1) not in result
    assert not (tmp_path / "Napi_értékesítési_riport_2027.xlsx").exists()


@pytest.mark.parametrize("exception", [URLError("offline"), TimeoutError()])
def test_network_manual_fallback_and_old_day_manual_then_reuse(
    monkeypatch, tmp_path, movement, exception
):
    def offline(**kwargs):
        raise exception

    monkeypatch.setattr(rates, "fetch_current_rate", offline)
    answers = iter(["405,5", "2026-10-09", "390,25", "2026-01-02"])
    monkeypatch.setattr("builtins.input", lambda prompt: next(answers))
    rows = [
        movement(day=date(2026, 1, 3), currency="EUR"),
        movement(day=date(2026, 1, 4), currency="EUR"),
    ]
    result = rates.get_exchange_rates(rows, {})
    assert result[date(2026, 1, 3)].rate == Decimal("390.25")
    assert result[date(2026, 1, 4)].source_date == date(2026, 1, 2)


def test_header_only_daily_register_is_valid(
    monkeypatch, tmp_path, movement
):
    workbooks = {2026: store(tmp_path / "report.xlsx", [])}
    day = date(2026, 10, 11)
    monkeypatch.setattr(
        rates,
        "fetch_current_rate",
        lambda **kwargs: DailyExchangeRate(day, day, Decimal(400)),
    )
    assert rates.get_exchange_rates([movement()], workbooks) == {
        day: DailyExchangeRate(day, day, Decimal(400))
    }


def test_automatic_fill_excludes_later_source_date(monkeypatch):
    day = date(2026, 9, 18)
    later = DailyExchangeRate(date(2026, 10, 12), date(2026, 10, 12), Decimal(400))
    earlier = DailyExchangeRate(date(2026, 9, 17), date(2026, 9, 17), Decimal(390))
    records = {later.application_date: later, earlier.application_date: earlier}
    def unexpected_manual(*args):
        raise AssertionError("A suitable earlier rate exists")
    monkeypatch.setattr(rates, "_request_manual_exchange_rate", unexpected_manual)
    rates._fill_missing_exchange_rates([day], records)
    assert records[day].rate == Decimal(390)
    assert records[day].source_date == date(2026, 9, 17)


def test_automatic_fill_requests_manual_if_only_later_source_exists(monkeypatch):
    day = date(2026, 9, 18)
    later = DailyExchangeRate(date(2026, 10, 12), date(2026, 10, 12), Decimal(400))
    manual = DailyExchangeRate(day, date(2026, 10, 12), Decimal(405))
    calls = []
    def request(application_date):
        calls.append(application_date)
        return manual
    monkeypatch.setattr(rates, "_request_manual_exchange_rate", request)
    records = {later.application_date: later}
    rates._fill_missing_exchange_rates([day], records)
    assert calls == [day]
    assert records[day] == manual
