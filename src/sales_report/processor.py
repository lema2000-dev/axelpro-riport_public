from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from .models import Movement, DailyTotals, MonthlyTotals, ImportResult


def _filter_new_movements(
    movements: list[Movement], processed_documents: set[tuple[int, str]]
) -> list[Movement]:
    return [
        movement
        for movement in movements
        if movement.document_key not in processed_documents
    ]


def _find_workshop_documents(movements: list[Movement]) -> set[tuple[int, str]]:
    return {
        movement.document_key
        for movement in movements
        if movement.transaction_type.casefold() == "invoice"
        and movement.group.casefold() == "labor"
    }


def _net_value_in_huf(
    movement: Movement, exchange_rates: dict[date, Decimal]
) -> Decimal:
    if movement.currency == "HUF":
        return movement.net_value

    rate = exchange_rates.get(movement.movement_date)

    if rate is None:
        raise ValueError(f"Missing EUR/HUF rate: {movement.movement_date}")

    if not rate.is_finite() or rate <= 0:
        raise ValueError(f"Invalid EUR/HUF rate: {movement.movement_date} -> {rate}")

    return movement.net_value * rate


def _round_huf(value: Decimal) -> int:
    return int(value.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _is_new_part(movement: Movement, new_part_groups: set[str]) -> bool:
    return movement.group.casefold() in new_part_groups


def _prepare_new_part_groups(additional_groups: set[str]) -> set[str]:
    return {"new parts"} | {
        group.strip().casefold() for group in additional_groups if group.strip()
    }


def _calculate_daily_totals(
    movement_date: date,
    movements: list[Movement],
    exchange_rates: dict[date, Decimal],
    new_part_groups: set[str],
    workshop_documents: set[tuple[int, str]],
) -> DailyTotals:
    totals = {
        "used_parts": Decimal(0),
        "new_parts": Decimal(0),
        "labor": Decimal(0),
        "workshop_used_parts": Decimal(0),
        "workshop_new_parts": Decimal(0),
        "purchases": Decimal(0),
        "shipping": Decimal(0),
        "trailer": Decimal(0),
    }

    for movement in movements:
        if movement.movement_date != movement_date:
            continue

        transaction_type = movement.transaction_type.casefold()
        group = movement.group.casefold()

        if transaction_type == "invoice" and (
            group == "used parts"
            or _is_new_part(movement, new_part_groups)
            or group == "labor"
        ):
            value = _net_value_in_huf(movement, exchange_rates)

            if group == "used parts":
                totals["used_parts"] += value

                if movement.document_key in workshop_documents:
                    totals["workshop_used_parts"] += value

            elif _is_new_part(movement, new_part_groups):
                totals["new_parts"] += value

                if movement.document_key in workshop_documents:
                    totals["workshop_new_parts"] += value

            elif group == "labor":
                totals["labor"] += value

        if transaction_type == "purchase" and group:
            value = _net_value_in_huf(movement, exchange_rates)
            totals["purchases"] += value

        if group == "shipping":
            value = _net_value_in_huf(movement, exchange_rates)
            totals["shipping"] += value

        if group == "trailer":
            value = _net_value_in_huf(movement, exchange_rates)
            totals["trailer"] += value

    return DailyTotals(
        movement_date=movement_date,
        used_parts=_round_huf(totals["used_parts"]),
        new_parts=_round_huf(totals["new_parts"]),
        labor=_round_huf(totals["labor"]),
        workshop_used_parts=_round_huf(totals["workshop_used_parts"]),
        workshop_new_parts=_round_huf(totals["workshop_new_parts"]),
        purchases=_round_huf(totals["purchases"]),
        shipping=_round_huf(totals["shipping"]),
        trailer=_round_huf(totals["trailer"]),
    )


def _calculate_monthly_totals(
    year: int,
    month: int,
    movements: list[Movement],
    exchange_rates: dict[date, Decimal],
) -> MonthlyTotals:
    totals = {
        "consumables": Decimal(0),
        "diagnostics": Decimal(0),
        "air_conditioning": Decimal(0),
    }

    for movement in movements:
        if movement.movement_date.year != year or movement.movement_date.month != month:
            continue

        if movement.transaction_type.casefold() != "invoice":
            continue

        item_code = movement.item_code.casefold()

        if item_code not in totals:
            continue

        value = _net_value_in_huf(movement, exchange_rates)
        totals[item_code] += value

    return MonthlyTotals(
        year=year,
        month=month,
        consumables=_round_huf(totals["consumables"]),
        diagnostics=_round_huf(totals["diagnostics"]),
        air_conditioning=_round_huf(totals["air_conditioning"]),
    )


def _aggregate_movements(
    movements: list[Movement],
    processed_documents: set[tuple[int, str]],
    exchange_rates: dict[date, Decimal],
    additional_new_part_groups: set[str],
) -> ImportResult | None:
    """Aggregate movements."""
    new_movements = _filter_new_movements(movements, processed_documents)

    if not new_movements:
        return None

    new_part_groups = _prepare_new_part_groups(additional_new_part_groups)

    workshop_documents = _find_workshop_documents(new_movements)

    dates = sorted({movement.movement_date for movement in new_movements})

    months = sorted({(day.year, day.month) for day in dates})

    daily_totals = [
        _calculate_daily_totals(
            movement_date=day,
            movements=new_movements,
            exchange_rates=exchange_rates,
            new_part_groups=new_part_groups,
            workshop_documents=workshop_documents,
        )
        for day in dates
    ]

    monthly_totals = [
        _calculate_monthly_totals(
            year=year,
            month=month,
            movements=new_movements,
            exchange_rates=exchange_rates,
        )
        for year, month in months
    ]

    new_documents = {movement.document_key for movement in new_movements}

    return ImportResult(
        daily_totals=daily_totals,
        monthly_totals=monthly_totals,
        new_documents=new_documents,
    )


def process_movements(
    movements: list[Movement],
    processed_documents: set[tuple[int, str]],
    exchange_rates: dict[date, Decimal],
    additional_new_part_groups: set[str],
) -> ImportResult | None:
    """Skip recorded documents and calculate daily and monthly import totals."""
    return _aggregate_movements(
        movements,
        processed_documents,
        exchange_rates,
        additional_new_part_groups,
    )