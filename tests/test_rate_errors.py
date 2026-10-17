from datetime import date
from http.client import HTTPException, IncompleteRead
from unittest.mock import Mock
import pytest
from openpyxl import Workbook
from sales_report import rates_read_and_correction as rates


@pytest.mark.parametrize(
    "failure",
    [
        ConnectionResetError("reset"),
        HTTPException("bad HTTP"),
        IncompleteRead(b"partial", 20),
    ],
)
def test_extra_network_failures_use_manual_input(monkeypatch, failure):
    def fail(**kwargs):
        raise failure

    monkeypatch.setattr(rates, "fetch_current_rate", fail)
    values = iter(["400,5", date.today().isoformat()])
    monkeypatch.setattr("builtins.input", lambda prompt: next(values))
    result = rates._get_current_exchange_rate()
    assert str(result.rate) == "400.5"
    assert result.application_date == date.today()


def test_rate_reader_does_not_close_shared_workbook_on_failure(tmp_path):
    wb = Workbook()
    wb.close = Mock(wraps=wb.close)
    with pytest.raises(KeyError):
        rates._read_stored_exchange_rates({2026: (wb, tmp_path / "report.xlsx")})
    wb.close.assert_not_called()


@pytest.mark.parametrize("invalid", ["", "abc", "0", "-1", "NaN", "Infinity", "1E999", "1E-999"])
def test_manual_rate_retries_invalid_number(monkeypatch, capsys, invalid):
    values = iter([invalid, "400,25", "2026-10-12"])
    monkeypatch.setattr("builtins.input", lambda prompt: next(values))
    result = rates._request_manual_exchange_rate(date(2026, 9, 18))
    assert str(result.rate) == "400.25"
    assert result.source_date == date(2026, 10, 12)
    assert "positive, finite" in capsys.readouterr().out


@pytest.mark.parametrize("invalid", ["", "2026-02-30", "20261012", "2026/10/12", "abc"])
def test_manual_date_retries_without_requesting_rate_again(monkeypatch, capsys, invalid):
    prompts = []
    values = iter(["400", invalid, "2026-10-12"])
    def answer(prompt):
        prompts.append(prompt)
        return next(values)
    monkeypatch.setattr("builtins.input", answer)
    result = rates._request_manual_exchange_rate(date(2026, 9, 18))
    assert result.source_date == date(2026, 10, 12)
    assert sum("EUR exchange rate" in prompt for prompt in prompts) == 1
    assert "valid date" in capsys.readouterr().out
