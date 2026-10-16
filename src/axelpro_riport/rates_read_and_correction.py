"""Napi árfolyamok beolvasása, lekérése és a hiányzó napok pótlása."""

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from collections.abc import Callable

from http.client import HTTPException

from openpyxl import Workbook

from .mnb_rates import fetch_current_rate

from .excel_constants import DAILY_EXCHANGE_RATES_SHEET
from .models import DailyExchangeRate, Movement

def _request_manual_exchange_rate(
    application_date: date,
) -> DailyExchangeRate:
    """Árfolyam bekérése, hibás adat esetén újrakéréssel."""
    while True:
        try:
            rate = Decimal(
                input(
                    f"EUR-árfolyam ({application_date:%Y-%m-%d}, Ft/EUR): "
                ).strip().replace(",", ".")
            )

            # Az árfolyamot a meglévő modell ellenőrzi.
            DailyExchangeRate(
                application_date=application_date,
                source_date=application_date,
                rate=rate,
            )
        except (InvalidOperation, ValueError, OverflowError):
            print(
                "Pozitív, véges, Excelben tárolható számot adj meg "
                "(például: 400 vagy 400,25)."
            )
            continue

        break

    while True:
        text = input(
            "Az árfolyam forrásdátuma (ÉÉÉÉ-HH-NN): "
        ).strip()

        try:
            source_date = date.fromisoformat(text)

            if source_date.isoformat() != text:
                raise ValueError
        except ValueError:
            print(
                "Érvényes dátumot adj meg ÉÉÉÉ-HH-NN formátumban "
                "(például: 2026-10-12)."
            )
            continue

        break

    return DailyExchangeRate(
        application_date=application_date,
        source_date=source_date,
        rate=rate,
    )

def _find_missing_exchange_rate_dates(
    movements: list[Movement],
    exchange_rates: dict[date, DailyExchangeRate],
) -> list[date]:
    """Az EUR-tételekhez hiányzó árfolyamnapok kigyűjtése."""
    required_dates = {
        movement.movement_date
        for movement in movements
        if movement.currency == "EUR"
    }

    return sorted(required_dates - exchange_rates.keys())

def _fill_missing_exchange_rates(
    missing_dates: list[date],
    exchange_rates: dict[date, DailyExchangeRate],
    *,
    request_rate: Callable[[date], DailyExchangeRate] | None = None,
) -> None:
    """A hiányzó napok pótlása ismert vagy kézzel megadott árfolyammal."""
    if request_rate is None:
        request_rate = _request_manual_exchange_rate

    for day in missing_dates:
        available_records = [
            record
            for record in exchange_rates.values()
            if record.source_date <= day
        ]

        if available_records:
            latest_record = max(
                available_records,
                key=lambda record: record.source_date,
            )

            record = DailyExchangeRate(
                application_date=day,
                source_date=latest_record.source_date,
                rate=latest_record.rate,
            )
        else:
            record = request_rate(day)

        exchange_rates[day] = record

def _get_current_exchange_rate(
        *,
        request_rate: Callable[[date], DailyExchangeRate] | None = None,
        show_message: Callable[[str], None] = print,
) -> DailyExchangeRate:
    """Aktuális MNB-árfolyam lekérése, kézi beviteli lehetőséggel."""
    if request_rate is None:
        request_rate = _request_manual_exchange_rate

    today = date.today()

    try:
        return fetch_current_rate(today=today)
    except (OSError, HTTPException) as exc:
        show_message(
            "Az MNB-árfolyam lekérése hálózati hiba miatt nem sikerült.\n"
            f"Részletek: {exc}\n"
            "Add meg kézzel az árfolyamot."
        )

        return request_rate(today)

def _read_stored_exchange_rates(
    workbooks: dict[int, tuple[Workbook, Path]],
) -> dict[date, DailyExchangeRate]:
    """Az érintett évek riportjaiban tárolt napi árfolyamok beolvasása."""
    exchange_rates: dict[date, DailyExchangeRate] = {}

    for year in sorted(workbooks.keys()):
        workbook = workbooks[year][0]

        sheet = workbook[DAILY_EXCHANGE_RATES_SHEET]

        for application_date, source_date, rate in sheet.iter_rows(
            min_row=2,
            max_col=3,
            values_only=True,
        ):
            if isinstance(application_date, datetime):
                application_date = application_date.date()

            if isinstance(source_date, datetime):
                source_date = source_date.date()

            exchange_rates[application_date] = DailyExchangeRate(
                application_date=application_date,
                source_date=source_date,
                rate=Decimal(str(rate)),
            )

    return exchange_rates

def get_exchange_rates(
    movements: list[Movement],
    workbooks: dict[int, tuple[Workbook, Path]],
    *,
    request_rate: Callable[[date], DailyExchangeRate] | None = None,
    show_message: Callable[[str], None] = print,
) -> dict[date, DailyExchangeRate]:
    """Az importhoz szükséges napi árfolyamok összeállítása."""
    exchange_rates = _read_stored_exchange_rates(
        workbooks=workbooks,
    )

    current_rate = _get_current_exchange_rate(
        request_rate=request_rate,
        show_message=show_message,
    )
    exchange_rates.setdefault(
        current_rate.application_date,
        current_rate,
    )

    missing_dates = _find_missing_exchange_rate_dates(
        movements=movements,
        exchange_rates=exchange_rates,
    )

    _fill_missing_exchange_rates(
        missing_dates=missing_dates,
        exchange_rates=exchange_rates,
        request_rate=request_rate,
    )

    return exchange_rates