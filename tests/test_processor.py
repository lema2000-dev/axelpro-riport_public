from datetime import date
from decimal import Decimal
from axelpro_riport.processor import process_movements


def test_business_categories_currency_and_workshop_are_aggregated(movement):
    rows = [
        movement(group="BONTOTT ALKATRÉSZ", amount="10"),
        movement(group="Munkadíj", amount="20"),
        movement(group="munkadíj", amount="30"),
        movement(group="Szűrők", amount="2", currency="EUR"),
        movement(document="OTHER", amount="7"),
        movement(document="BUY", group="Bármi", transaction="Beszerzés", amount="8"),
        movement(
            document="SHIP",
            group="Szállítási költség",
            transaction="Szállítólevél",
            amount="9",
        ),
        movement(
            document="TRAIL",
            group="Trailer",
            transaction="Szállítólevél",
            amount="1",
            currency="EUR",
        ),
    ]
    result = process_movements(
        rows, set(), {date(2026, 1, 5): Decimal("400")}, {" SZŰRŐK "}
    )
    daily = result.daily_totals[0]
    assert (daily.used_parts, daily.new_parts, daily.labor) == (10, 807, 50)
    assert (daily.workshop_used_parts, daily.workshop_new_parts) == (10, 800)
    assert (daily.purchases, daily.shipping, daily.trailer) == (8, 9, 400)
    assert result.new_documents == {m.document_key for m in rows}


def test_monthly_codes_only_invoices_and_document_key_includes_year(movement):
    rows = [
        movement(item="sany", amount="12.5"),
        movement(document="D", item="DIAG", amount="2"),
        movement(document="K", item="kj", amount="3"),
        movement(document="X", item="SANY", transaction="Beszerzés", amount="100"),
        movement(day=date(2027, 1, 5), item="SANY", amount="8"),
    ]
    result = process_movements(rows, {(2026, "SZ-1")}, {}, set())
    assert [(t.year, t.month, t.sany, t.diag, t.kj) for t in result.monthly_totals] == [
        (2026, 1, 0, 2, 3),
        (2027, 1, 8, 0, 0),
    ]
    assert (2026, "SZ-1") not in result.new_documents
    assert (2027, "SZ-1") in result.new_documents


def test_rounds_after_sum_and_negative_half_away_from_zero(movement):
    rows = [
        movement(amount="0.25"),
        movement(amount="0.25"),
        movement(day=date(2026, 2, 2), document="NEG", amount="-12.5"),
    ]
    result = process_movements(rows, set(), {}, set())
    assert [t.new_parts for t in result.daily_totals] == [1, -13]


def test_reimport_has_no_increments(movement):
    first = process_movements([movement()], set(), {}, set())
    second = process_movements([movement()], first.new_documents, {}, set())
    assert second is None


def test_monthly_eur_rates_vary_by_day_and_month(movement):
    rows = [
        movement(day=date(2026, 1, 5), item="SANY", currency="EUR", amount="1.5"),
        movement(
            day=date(2026, 1, 6), document="B", item="SANY", currency="EUR", amount="2"
        ),
        movement(
            day=date(2026, 2, 2), document="C", item="SANY", currency="EUR", amount="3"
        ),
    ]
    result = process_movements(
        rows,
        set(),
        {
            date(2026, 1, 5): Decimal(400),
            date(2026, 1, 6): Decimal(410),
            date(2026, 2, 2): Decimal(420),
        },
        set(),
    )
    assert [(t.month, t.sany) for t in result.monthly_totals] == [(1, 1420), (2, 1260)]
    assert [t.new_parts for t in result.daily_totals] == [600, 820, 1260]


def test_ungrouped_purchases_are_excluded_but_grouped_purchases_are_counted(movement):
    rows = [
        movement(document="EMPTY-HUF", transaction="Beszerzés", group="", amount="1000"),
        movement(document="EMPTY-EUR", transaction="Beszerzés", group="", amount="20", currency="EUR"),
        movement(document="GROUPED-HUF", transaction="Beszerzés", group="Bármi", amount="8"),
        movement(document="GROUPED-EUR", transaction="Beszerzés", group="Másik csoport", amount="2", currency="EUR"),
    ]
    result = process_movements(rows, set(), {date(2026, 1, 5): Decimal("400")}, set())
    assert result.daily_totals[0].purchases == 808
    # Csak a beszerzés összegéből maradnak ki, nem a feldolgozott bizonylatokból.
    assert result.new_documents == {row.document_key for row in rows}
    assert process_movements(rows, result.new_documents, {}, set()) is None


def test_only_ungrouped_eur_purchase_does_not_require_rate_in_processor(movement):
    row = movement(transaction="Beszerzés", group="", currency="EUR", amount="10")
    result = process_movements([row], set(), {}, set())
    assert result.daily_totals[0].purchases == 0
    assert result.new_documents == {row.document_key}
