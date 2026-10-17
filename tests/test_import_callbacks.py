"""Test import callbacks."""
from datetime import date
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import pytest
from sales_report import import_service as service, rates_read_and_correction as rates
from sales_report.csv_reader import CsvFormatError
from sales_report.models import DailyExchangeRate, ImportResult


@pytest.fixture
def prepared(monkeypatch, movement, tmp_path):
    wb = Mock()
    rows = [movement()]
    reports = {2026: (wb, tmp_path / 'report.xlsx')}
    result = ImportResult([], [], {(2026, 'SZ-1')})
    monkeypatch.setattr(service, 'read_movements', lambda **kwargs: rows)
    monkeypatch.setattr(service, 'load_settings', lambda: SimpleNamespace(output_directory=str(tmp_path), additional_new_part_groups=[]))
    monkeypatch.setattr(service, '_loads_workbooks', lambda **kwargs: reports)
    get_rates = Mock(return_value={})
    monkeypatch.setattr(service, 'get_exchange_rates', get_rates)
    monkeypatch.setattr(service, '_read_processed_documents', lambda **kwargs: set())
    process = Mock(return_value=result)
    monkeypatch.setattr(service, 'process_movements', process)
    write = Mock()
    monkeypatch.setattr(service, 'write_import_result', write)
    monkeypatch.setattr(service, 'write_exchange_rates', Mock())
    save = Mock()
    monkeypatch.setattr(service, 'save_workbook', save)
    return SimpleNamespace(wb=wb, rows=rows, result=result, get_rates=get_rates, process=process, write=write, save=save)


def test_custom_preview_works_with_default_cancellation(prepared):
    preview = Mock(return_value=True)
    service.import_csv('input.csv', confirm_preview=preview)
    preview.assert_called_once()
    prepared.save.assert_called_once()
    prepared.wb.close.assert_called_once()


def test_callbacks_are_forwarded_and_keyword_preview_is_supported(prepared):
    request = Mock()
    message = Mock()
    check = Mock()
    preview = Mock(return_value=True)
    service.import_csv('input.csv', request_rate=request, show_message=message, check_cancelled=check, confirm_preview=preview)
    assert prepared.get_rates.call_args.kwargs['request_rate'] is request
    assert prepared.get_rates.call_args.kwargs['show_message'] is message
    assert preview.call_args.kwargs['import_result'] is prepared.result
    assert check.call_count == 3
    prepared.save.assert_called_once()
    prepared.wb.close.assert_called_once()


@pytest.mark.parametrize('point', [1, 2, 3])
def test_cancellation_before_processing_or_saving_closes_reports(prepared, point):
    failure = RuntimeError('cancelled')
    count = 0
    def check():
        nonlocal count
        count += 1
        if count == point:
            raise failure
    with pytest.raises(RuntimeError) as error:
        service.import_csv('input.csv', check_cancelled=check, confirm_preview=Mock(return_value=True))
    assert error.value is failure
    prepared.save.assert_not_called()
    prepared.wb.close.assert_called_once()
    if point < 3:
        assert prepared.write.call_count == 0


def test_no_new_items_use_message_callback_without_preview(prepared):
    prepared.process.return_value = None
    message = Mock()
    preview = Mock()
    service.import_csv('input.csv', show_message=message, confirm_preview=preview, check_cancelled=lambda: None)
    message.assert_called_once_with('There are no new items to import.')
    preview.assert_not_called()
    prepared.save.assert_not_called()
    prepared.wb.close.assert_called_once()


def test_csv_format_error_preserves_message_and_cause(monkeypatch):
    failure = ValueError('Error on line 4: Invalid date')
    monkeypatch.setattr(service, 'read_movements', Mock(side_effect=failure))
    with pytest.raises(CsvFormatError) as error:
        service.import_csv('input.csv')
    assert str(error.value) == str(failure)
    assert error.value.__cause__ is failure


def test_mnb_failure_calls_message_and_manual_callback(monkeypatch):
    monkeypatch.setattr(rates, 'fetch_current_rate', Mock(side_effect=OSError('offline')))
    day = date.today()
    record = DailyExchangeRate(day, day, Decimal(400))
    request = Mock(return_value=record)
    message = Mock()
    assert rates._get_current_exchange_rate(request_rate=request, show_message=message) == record
    request.assert_called_once_with(day)
    assert 'offline' in message.call_args.args[0]


def test_missing_rates_call_custom_input(monkeypatch):
    day = date(2026, 1, 5)
    record = DailyExchangeRate(day, date(2026, 10, 13), Decimal(400))
    request = Mock(return_value=record)
    records = {}
    rates._fill_missing_exchange_rates([day], records, request_rate=request)
    assert records == {day: record}
    request.assert_called_once_with(day)


def test_ungrouped_decision_callback_is_forwarded_to_csv_reader(prepared, monkeypatch):
    reader = Mock(return_value=prepared.rows)
    decide = Mock(return_value="New parts")
    monkeypatch.setattr(service, "read_movements", reader)
    service.import_csv("input.csv", request_missing_group=decide, confirm_preview=Mock(return_value=False))
    assert reader.call_args.kwargs["request_missing_group"] is decide
    prepared.save.assert_not_called()
    prepared.wb.close.assert_called_once()
