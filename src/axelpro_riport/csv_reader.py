import csv
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from collections.abc import Callable

from .models import Movement

class CsvFileAccessError(OSError):
    """A CSV-fájl olvasása sikertelen."""

class CsvFormatError(ValueError):
    """A CSV szekezete vagy valamelyik adata hibás."""

REQUIRED_COLUMNS = {
    "Árumozgás időpontja",
    "Ügylet típus",
    "Bizonylatszám",
    "Partner",
    "Csoport",
    "Cikkszám",
    "Ne. érték",
    "Pénznem",
}

def _parse_net_value(value: str) -> Decimal:
    normalized = value.strip().replace(',', '.')

    try:
        result = Decimal(normalized)
    except InvalidOperation as exc:
        raise ValueError(f"Hibás nettó érték: {value!r}") from exc

    if not result.is_finite():
        raise ValueError(f"Nem véges nettó érték: {value!r}")

    return result

def _parse_movement_date(value: str) -> date:
    normalized = value.strip()

    try:
        timestamp = datetime.strptime(normalized, "%Y. %m. %d. %H:%M:%S")
    except ValueError as exc:
        raise ValueError(f"Hibás árumozgási időpont: {value!r}") from exc

    return timestamp.date()
    
def _parse_movement(
    row: dict[str, str],
    request_missing_group: Callable[[str, str], str] | None = None,
) -> Movement:
    currency = row["Pénznem"].strip().upper()

    if currency not in ("HUF", "EUR"):
        raise ValueError(f"Nem támogatott pénznem: {currency!r}")

    document_number = row["Bizonylatszám"].strip()

    if not document_number:
        raise ValueError("Hiányzó bizonylatszám")

    group = row["Csoport"].strip()
    item_code = row["Cikkszám"].strip()
    transaction_type = row["Ügylet típus"].strip()

    if not group and transaction_type.casefold() == "számla" and request_missing_group is not None:
        group = request_missing_group(document_number, item_code)

    return Movement(
        movement_date=_parse_movement_date(row["Árumozgás időpontja"]),
        transaction_type=transaction_type,
        document_number=document_number,
        group=group,
        item_code=item_code,
        net_value=_parse_net_value(row["Ne. érték"]),
        currency=currency,
    )

def _read_movements(
        path: str | Path,
        request_missing_group: Callable[[str, str], str] | None = None,
) -> list[Movement]:
    movements: list[Movement] = []

    with Path(path).open(encoding="utf-16", newline="") as file:
        reader = csv.DictReader(file, delimiter="\t")

        columns = set(reader.fieldnames or [])
        missing_columns = REQUIRED_COLUMNS - columns

        if missing_columns:
            raise ValueError("Hiányzó CSV-oszlopok: " + ", ".join(sorted(missing_columns)))

        for row in reader:
            if None in row or any(value is None for value in row.values()):
                raise ValueError(f"Hibás CSV-szerkezet a(z) {reader.line_num}. sorban.")

            if row["Ügylet típus"].strip().casefold() == "leltár" or row["Partner"].strip().casefold() in ("cég1", "cég2"):
                continue

            try:
                movement = _parse_movement(row, request_missing_group)
            except ValueError as exc:
                raise ValueError(f"Hiba a(z) {reader.line_num}. sorban: {exc}") from exc

            movements.append(movement)

    return movements

def read_movements(
        path: str | Path,
        request_missing_group: Callable[[str, str], str] | None = None) -> list[Movement]:
    """CSV-beolvasás a fájlhozzáférési hibák továbbadásával."""
    try:
        return _read_movements(path, request_missing_group)
    except OSError as exc:
        raise CsvFileAccessError(
            f"A CSV-fájl megnyitása vagy olvasása nem sikerült.\n"
            f"Fájl: {path}\n"
            f"Részletek: {exc}"
        ) from exc
