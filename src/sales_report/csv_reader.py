import csv
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from collections.abc import Callable

from .models import Movement

class CsvFileAccessError(OSError):
    """Csv file access error."""

class CsvFormatError(ValueError):
    """Csv format error."""

REQUIRED_COLUMNS = {
    "Movement timestamp",
    "Transaction type",
    "Document number",
    "Partner",
    "Group",
    "Item code",
    "Net value",
    "Currency",
}

def _parse_net_value(value: str) -> Decimal:
    normalized = value.strip().replace(',', '.')

    try:
        result = Decimal(normalized)
    except InvalidOperation as exc:
        raise ValueError(f"Invalid net value: {value!r}") from exc

    if not result.is_finite():
        raise ValueError(f"Non-finite net value: {value!r}")

    return result

def _parse_movement_date(value: str) -> date:
    normalized = value.strip()

    try:
        timestamp = datetime.strptime(normalized, "%Y. %m. %d. %H:%M:%S")
    except ValueError as exc:
        raise ValueError(f"Invalid movement timestamp: {value!r}") from exc

    return timestamp.date()
    
def _parse_movement(
    row: dict[str, str],
    request_missing_group: Callable[[str, str], str] | None = None,
) -> Movement:
    currency = row["Currency"].strip().upper()

    if currency not in ("HUF", "EUR"):
        raise ValueError(f"Unsupported currency: {currency!r}")

    document_number = row["Document number"].strip()

    if not document_number:
        raise ValueError("Missing document number")

    group = row["Group"].strip()
    item_code = row["Item code"].strip()
    transaction_type = row["Transaction type"].strip()

    if not group and transaction_type.casefold() == "invoice" and request_missing_group is not None:
        group = request_missing_group(document_number, item_code)

    return Movement(
        movement_date=_parse_movement_date(row["Movement timestamp"]),
        transaction_type=transaction_type,
        document_number=document_number,
        group=group,
        item_code=item_code,
        net_value=_parse_net_value(row["Net value"]),
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
            raise ValueError("Missing CSV columns: " + ", ".join(sorted(missing_columns)))

        for row in reader:
            if None in row or any(value is None for value in row.values()):
                raise ValueError(f"Invalid CSV structure on line {reader.line_num}..")

            if row["Transaction type"].strip().casefold() == "inventory" or row["Partner"].strip().casefold() in ("company 1", "company 2"):
                continue

            try:
                movement = _parse_movement(row, request_missing_group)
            except ValueError as exc:
                raise ValueError(f"Error on line {reader.line_num}: {exc}") from exc

            movements.append(movement)

    return movements

def read_movements(
        path: str | Path,
        request_missing_group: Callable[[str, str], str] | None = None) -> list[Movement]:
    """Read a UTF-16 tab-separated export and return parsed movement records."""
    try:
        return _read_movements(path, request_missing_group)
    except OSError as exc:
        raise CsvFileAccessError(
            f"Could not open or read the CSV file.\n"
            f"File: {path}\n"
            f"Details: {exc}"
        ) from exc
