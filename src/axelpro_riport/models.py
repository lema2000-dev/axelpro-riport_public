from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from math import isfinite


@dataclass(frozen=True)
class Movement:
    movement_date: date
    transaction_type: str
    document_number: str
    group: str
    item_code: str
    net_value: Decimal
    currency: str

    @property
    def document_key(self) -> tuple[int, str]:
        return self.movement_date.year, self.document_number


@dataclass(frozen=True)
class DailyTotals:
    movement_date: date
    used_parts: int
    new_parts: int
    labor: int
    workshop_used_parts: int
    workshop_new_parts: int
    purchases: int
    shipping: int
    trailer: int


@dataclass(frozen=True)
class MonthlyTotals:
    year: int
    month: int
    sany: int
    diag: int
    kj: int


@dataclass(frozen=True)
class ImportResult:
    daily_totals: list[DailyTotals]
    monthly_totals: list[MonthlyTotals]
    new_documents: set[tuple[int, str]]


@dataclass(frozen=True)
class DailyExchangeRate:
    application_date: date
    source_date: date
    rate: Decimal

    def __post_init__(self) -> None:
        """A külső vagy Excelből olvasott rekord egyszeri ellenőrzése."""
        if (
            not isinstance(self.application_date, date)
            or isinstance(self.application_date, datetime)
            or not isinstance(self.source_date, date)
            or isinstance(self.source_date, datetime)
            or not isinstance(self.rate, Decimal)
            or not self.rate.is_finite()
            or self.rate <= 0
        ):
            raise ValueError("Hibás napi árfolyamrekord.")
        numeric_rate = float(self.rate)
        if not isfinite(numeric_rate) or numeric_rate <= 0:
            raise ValueError("Excelben nem tárolható árfolyam.")
