"""Check the public English schema and bundled synthetic examples."""
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from sales_report.csv_reader import read_movements
from sales_report.processor import process_movements
from sales_report.excel_constants import MONTH_NAMES

SAMPLES = Path(__file__).resolve().parents[1] / "samples"


@pytest.mark.parametrize("name", [
    "across_days.csv", "across_days_extended.csv",
    "across_months.csv", "across_years.csv",
])
def test_main_samples_use_the_english_schema(name):
    movements = read_movements(SAMPLES / name)
    assert movements
    assert all(m.transaction_type.casefold() != "inventory" for m in movements)
    assert any(m.group.casefold() == "new parts" for m in movements)
    assert any(m.group.casefold() == "labor" for m in movements)
    assert any(m.currency == "EUR" for m in movements)


def test_extended_sample_preserves_original_document_identities_and_adds_new_ones():
    original = read_movements(SAMPLES / "across_days.csv")
    extended = read_movements(SAMPLES / "across_days_extended.csv")
    original_keys = {m.document_key for m in original}
    extended_keys = {m.document_key for m in extended}
    assert original_keys < extended_keys
    rates = {m.movement_date: Decimal("400") for m in extended}
    result = process_movements(extended, original_keys, rates, {"Filters"})
    assert result is not None
    assert result.new_documents == extended_keys - original_keys
    assert {m.item_code.casefold() for m in extended if m.document_key in result.new_documents} >= {
        "consumables", "diagnostics", "air_conditioning",
    }


def test_generated_report_uses_english_sheets_and_cross_sheet_formulas(year_report):
    assert set(MONTH_NAMES) <= set(year_report.sheetnames)
    assert {"Summary", "Exchange_Rates", "_daily_exchange_rates", "_imported_documents"} <= set(year_report.sheetnames)
    january = year_report["January"]
    assert january["B1"].value == "Thursday"
    assert january.freeze_panes == "D4"
    assert "'Exchange_Rates'!" in january["B60"].value
    assert "'January'!" in year_report["Summary"]["B40"].value
    assert year_report["_imported_documents"]["B1"].value == "Document number"
