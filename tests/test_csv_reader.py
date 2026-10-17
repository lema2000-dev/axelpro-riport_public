from decimal import Decimal
import pytest
from sales_report.csv_reader import read_movements


def test_reads_utf16_tabs_negative_decimal_and_trims_fields(
    tmp_path, write_csv, movement
):
    rows = [
        movement(
            document=" SZ-1 ",
            amount="-12.50",
            currency=" eur ",
            group=" Used parts ",
        )
    ]
    result = read_movements(write_csv(tmp_path / "input.csv", rows))
    assert len(result) == 1
    assert result[0].net_value == Decimal("-12.50")
    assert result[0].currency == "EUR"
    assert result[0].document_number == "SZ-1"
    assert result[0].group == "Used parts"
    assert result[0].movement_date == rows[0].movement_date


def test_inventory_is_ignored(tmp_path, write_csv, movement):
    path = write_csv(
        tmp_path / "input.csv",
        [movement(transaction="INVENTORY", document=""), movement()],
    )
    assert read_movements(path) == [movement()]


@pytest.mark.parametrize(
    "changes", [{"currency": "USD"}, {"document": ""}, {"amount": "NaN"}]
)
def test_existing_validation_reports_invalid_row(
    tmp_path, write_csv, movement, changes
):
    path = write_csv(tmp_path / "input.csv", [movement(**changes)])
    with pytest.raises(ValueError, match="3|2"):
        read_movements(path)


def test_missing_columns_are_reported(tmp_path):
    path = tmp_path / "input.csv"
    path.write_text("Document number\nSZ-1\n", encoding="utf-16")
    with pytest.raises(ValueError, match="Missing CSV columns"):
        read_movements(path)


@pytest.mark.parametrize("expected", ["New parts", "Used parts", "Labor", "Shipping", "Trailer", ""])
def test_ungrouped_invoice_uses_user_decision(tmp_path, write_csv, movement, expected):
    from unittest.mock import Mock
    decide = Mock(return_value=expected)
    path = write_csv(tmp_path / "input.csv", [movement(group="  ", document="DOC", item="ITEM", transaction="INVOICE")])
    result = read_movements(path, request_missing_group=decide)
    decide.assert_called_once_with("DOC", "ITEM")
    assert len(result) == 1
    assert result[0].group == expected
    assert result[0].document_number == "DOC"
    assert result[0].item_code == "ITEM"


@pytest.mark.parametrize("transaction,group", [("Inventory", ""), ("Purchase", ""), ("Invoice", "New parts")])
def test_group_decision_only_for_ungrouped_invoices(tmp_path, write_csv, movement, transaction, group):
    from unittest.mock import Mock
    decide = Mock()
    path = write_csv(tmp_path / "input.csv", [movement(transaction=transaction, group=group)])
    result = read_movements(path, request_missing_group=decide)
    decide.assert_not_called()
    assert len(result) == (0 if transaction == "Inventory" else 1)


def test_ungrouped_invoice_without_callback_remains_unchanged(tmp_path, write_csv, movement):
    row = movement(group="")
    path = write_csv(tmp_path / "input.csv", [row])
    assert read_movements(path) == [row]


@pytest.mark.parametrize("group,category", [
    ("New parts", "new_parts"),
    ("Used parts", "used_parts"),
    ("Labor", "labor"),
    ("Shipping", "shipping"),
    ("Trailer", "trailer"),
    ("", None),
])
def test_selected_missing_group_is_used_by_processor(tmp_path, write_csv, movement, group, category):
    from sales_report.processor import process_movements
    from dataclasses import fields
    path = write_csv(tmp_path / "input.csv", [movement(group="", amount="123", item="PART")])
    rows = read_movements(path, request_missing_group=lambda document, item: group)
    result = process_movements(
        movements=rows,
        processed_documents=set(),
        exchange_rates={},
        additional_new_part_groups=set(),
    )
    assert result is not None
    assert result.new_documents == {(2026, "SZ-1")}
    if category is not None:
        assert len(result.daily_totals) == 1
        daily = result.daily_totals[0]
        for field in fields(daily):
            if field.name != "movement_date":
                assert getattr(daily, field.name) == (123 if field.name == category else 0)
    else:
        assert all(
            getattr(daily, field.name) == 0
            for daily in result.daily_totals
            for field in fields(daily)
            if field.name != "movement_date"
        )


def test_group_callback_is_called_separately_for_each_ungrouped_invoice(tmp_path, write_csv, movement):
    from unittest.mock import Mock, call
    decide = Mock(side_effect=["Labor", "Used parts"])
    path = write_csv(tmp_path / "input.csv", [
        movement(group="", document="DOC", item="LABOR"),
        movement(group="", document="DOC", item="PART"),
    ])
    rows = read_movements(path, request_missing_group=decide)
    assert decide.call_args_list == [call("DOC", "LABOR"), call("DOC", "PART")]
    assert [row.group for row in rows] == ["Labor", "Used parts"]


def test_user_group_choices_preserve_workshop_invoice_classification(tmp_path, write_csv, movement):
    from unittest.mock import Mock
    from sales_report.processor import process_movements
    path = write_csv(tmp_path / "input.csv", [
        movement(group="", document="DOC", item="LABOR", amount="100"),
        movement(group="", document="DOC", item="USED", amount="200"),
        movement(group="", document="DOC", item="NEW", amount="300"),
    ])
    rows = read_movements(path, request_missing_group=Mock(side_effect=[
        "Labor", "Used parts", "New parts",
    ]))
    result = process_movements(rows, set(), {}, set())
    daily = result.daily_totals[0]
    assert (daily.labor, daily.used_parts, daily.new_parts) == (100, 200, 300)
    assert (daily.workshop_used_parts, daily.workshop_new_parts) == (200, 300)
    assert result.new_documents == {(2026, "DOC")}


def test_group_callback_cancellation_is_not_converted_to_csv_format_error(tmp_path, write_csv, movement):
    from unittest.mock import Mock
    failure = RuntimeError("user cancelled")
    path = write_csv(tmp_path / "input.csv", [movement(group="")])
    with pytest.raises(RuntimeError) as error:
        read_movements(path, request_missing_group=Mock(side_effect=failure))
    assert error.value is failure


@pytest.mark.parametrize("partner", [
    "Company 1",
    "Company 2",
    "  cOmPaNy 1  ",
    "  cOmPaNy 2  ",
])
@pytest.mark.parametrize("transaction", ["Invoice", "Purchase", "Delivery note"])
def test_excluded_partners_are_skipped_for_every_transaction_type(
    tmp_path, write_csv, movement, partner, transaction
):
    from unittest.mock import Mock

    decide = Mock(side_effect=AssertionError("Do not request a group for an excluded partner"))
    excluded = movement(document="EXCLUDED", transaction=transaction, group="")
    kept = movement(document="KEPT", amount="17")
    path = write_csv(
        tmp_path / "input.csv", [excluded, kept],
        partners=[partner, "Other partner"],
    )
    assert read_movements(path, request_missing_group=decide) == [kept]
    decide.assert_not_called()


@pytest.mark.parametrize("partner", [
    "Other partner", "", "Company 1 branch",
    "Company 2 copy",
])
def test_other_partners_are_not_excluded(tmp_path, write_csv, movement, partner):
    row = movement()
    path = write_csv(tmp_path / "input.csv", [row], partners=[partner])
    assert read_movements(path) == [row]


def test_partner_column_is_required(tmp_path, write_csv, movement):
    import csv

    path = write_csv(tmp_path / "input.csv", [movement()])
    with path.open(encoding="utf-16", newline="") as stream:
        rows = list(csv.reader(stream, delimiter="\t"))
    index = rows[0].index("Partner")
    with path.open("w", encoding="utf-16", newline="") as stream:
        csv.writer(stream, delimiter="\t").writerows([
            row[:index] + row[index + 1:] for row in rows
        ])
    with pytest.raises(ValueError, match="Missing CSV columns: Partner"):
        read_movements(path)


def test_excluded_partner_rows_do_not_reach_totals_or_document_register(
    tmp_path, write_csv, movement
):
    from sales_report.processor import process_movements

    rows = [
        movement(document="BLOCKED-INVOICE", amount="900", item="CONSUMABLES"),
        movement(document="BLOCKED-BUY", transaction="Purchase", amount="800"),
        movement(document="KEPT", amount="100"),
    ]
    path = write_csv(tmp_path / "input.csv", rows, partners=[
        "Company 1", "Company 2", "Example customer",
    ])
    result = process_movements(read_movements(path), set(), {}, set())
    assert result.new_documents == {(2026, "KEPT")}
    assert result.daily_totals[0].new_parts == 100
    assert result.daily_totals[0].purchases == 0
    assert result.monthly_totals[0].consumables == 0
