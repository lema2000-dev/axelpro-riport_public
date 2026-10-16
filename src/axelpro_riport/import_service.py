from pathlib import Path
from datetime import date
from decimal import Decimal
from collections.abc import Callable

from openpyxl import Workbook

from .excel_rates_writer import write_exchange_rates
from .excel_writer import write_import_result
from .processor import process_movements
from .models import DailyExchangeRate, ImportResult, Movement
from .excel_constants import PROCESSED_DOCUMENTS_SHEET
from .excel_files import (
    load_year_workbook,
    save_workbook,
    WorkbookSaveError,
)
from .csv_reader import read_movements, CsvFormatError
from .settings import load_settings
from .rates_read_and_correction import get_exchange_rates
from .template_creator import create_template_workbook
from .excel_constants import (
    DAILY_TARGET_ROWS,
    MONTHLY_TARGET_ROWS,
)

def _read_processed_documents(
    workbooks: dict[int, tuple[Workbook, Path]],
) -> set[tuple[int, str]]:
    """Az érintett évek már importált bizonylatainak kigyűjtése."""
    processed_documents: set[tuple[int, str]] = set()

    for year in sorted(workbooks.keys()):
        workbook = workbooks[year][0]

        sheet = workbook[PROCESSED_DOCUMENTS_SHEET]

        processed_documents.update(
            sheet.iter_rows(
                min_row=2,
                max_col=2,
                values_only=True,
            )
        )

    return processed_documents

def _get_processor_exchange_rates(
    exchange_rates: dict[date, DailyExchangeRate],
) -> dict[date, Decimal]:
    """Az árfolyamrekordokból a processor számára árfolyamszótár készítése."""
    return {
        day: record.rate
        for day, record in exchange_rates.items()
    }

def _confirm_import_preview(
    year: int,
    path: Path,
    movements: list[Movement],
    import_result: ImportResult,
) -> bool:
    """Az adott év importelőnézetének megjelenítése és jóváhagyása."""
    print(f"\n=== {year}. évi import előnézete ===")
    print(f"Célfájl: {path}")

    print("\nHozzáadandó napi összegek (Ft):")

    for totals in sorted(
        import_result.daily_totals,
        key=lambda totals: totals.movement_date,
    ):
        if totals.movement_date.year != year:
            continue

        print(f"\n{totals.movement_date:%Y-%m-%d}")

        for field_name in DAILY_TARGET_ROWS:
            print(f"  {field_name}: {getattr(totals, field_name):,}")

    print("\nHozzáadandó havi összegek (Ft):")

    for totals in sorted(
        import_result.monthly_totals,
        key=lambda totals: (totals.year, totals.month),
    ):
        if totals.year != year:
            continue

        print(f"\n{totals.year}-{totals.month:02d}")

        for field_name in MONTHLY_TARGET_ROWS:
            print(f"  {field_name}: {getattr(totals, field_name):,}")

    new_documents = {
        document_number
        for document_year, document_number in import_result.new_documents
        if document_year == year
    }

    skipped_movements = [
        movement
        for movement in movements
        if movement.movement_date.year == year
        and movement.document_key not in import_result.new_documents
    ]

    print(f"\nÚj bizonylatok száma: {len(new_documents)}")
    print(f"Már feldolgozott, kihagyott tételek száma: {len(skipped_movements)}")

    for movement in skipped_movements:
        print(
            f"  {movement.movement_date:%Y-%m-%d} | "
            f"{movement.document_number} | "
            f"{movement.item_code} | "
            f"{movement.net_value} {movement.currency}"
        )

    answer = input("\nElfogadod és mented ezt az évet? (i/n): ")
    return answer.strip().lower() == "i"

def _loads_workbooks(
    movements: list[Movement],
    directory: str | Path,
) -> dict[int, tuple[Workbook, Path]]:
    """Az érintett évek munkafüzeteinek betöltése vagy létrehozása."""
    years = {movement.movement_date.year for movement in movements}
    workbooks: dict[int, tuple[Workbook, Path]] = {}
    template_workbook = create_template_workbook()

    try:
        for year in sorted(years):
            months = {movement.movement_date.month for movement in movements if movement.movement_date.year == year}
            workbook, path = load_year_workbook(
                directory=directory,
                year=year,
                months=months,
                template_sheet=template_workbook["Szeptember"],
            )
            workbooks[year] = (workbook, path)

        return workbooks

    except Exception:
        # Ha a betöltés félbeszakad, a korábban betöltötteket bezárjuk.
        for workbook, _ in workbooks.values():
            workbook.close()

        raise

    finally:
        template_workbook.close()

def import_csv(
    csv_path: str | Path,
    *,
    request_rate: Callable[[date], DailyExchangeRate] | None = None,
    show_message: Callable[[str], None] = print,
    confirm_preview: Callable[
        [int, Path, list[Movement], ImportResult], bool
    ] | None = None,
    check_cancelled: Callable[[], None] | None = None,
    request_missing_group: Callable[[str, str], str] | None = None,
) -> None:
    """CSV importálása az érintett éves riportokba."""
    if confirm_preview is None:
        confirm_preview = _confirm_import_preview

    if check_cancelled is None:
        check_cancelled = lambda: None

    try:
        movements = read_movements(path=csv_path, request_missing_group=request_missing_group)
    except ValueError as exc:
        raise CsvFormatError(str(exc)) from exc
    
    settings = load_settings()

    workbooks = _loads_workbooks(
        movements=movements,
        directory=settings.output_directory,
    )

    try:
        check_cancelled()

        exchange_rates = get_exchange_rates(
            movements=movements,
            workbooks=workbooks,
            request_rate=request_rate,
            show_message=show_message,
        )

        processed_documents = _read_processed_documents(
            workbooks=workbooks,
        )

        import_result = process_movements(
            movements=movements,
            processed_documents=processed_documents,
            exchange_rates=_get_processor_exchange_rates(
                exchange_rates=exchange_rates,
            ),
            additional_new_part_groups=set(
                settings.additional_new_part_groups
            ),
        )

        if import_result is None:
            show_message("Nincs új tétel az importáláshoz.")
            return

        saved_years: list[int] = []

        for year in sorted(workbooks):
            check_cancelled()

            workbook, path = workbooks[year]

            write_import_result(
                workbook=workbook,
                year=year,
                result=import_result,
            )

            write_exchange_rates(
                workbook=workbook,
                year=year,
                exchange_rates=exchange_rates,
            )

            if not confirm_preview(
                year=year,
                path=path,
                movements=movements,
                import_result=import_result,
            ):
                continue

            check_cancelled()

            try:
                save_workbook(
                    path=path,
                    workbook=workbook,
                )
            except WorkbookSaveError as exc:
                saved_description = (
                    ", ".join(str(value) for value in saved_years)
                    if saved_years
                    else "nincs"
                )

                raise WorkbookSaveError(
                    f"{exc}\n"
                    f"Sikertelen mentés éve: {year}\n"
                    f"Már sikeresen elmentett évek: {saved_description}\n"
                    "Az import további éveinek feldolgozása leállt."
                ) from exc

            saved_years.append(year)

    finally:
        for workbook, _ in workbooks.values():
            workbook.close()