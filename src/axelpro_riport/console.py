"""A program terminálon keresztüli kezelése."""

from time import perf_counter
from pathlib import Path
import warnings

from .import_service import import_csv
from .settings import (
    add_new_part_groups,
    change_output_directory,
    create_settings,
    load_settings,
    reset,
    settings_path,
    remove_new_part_groups,
    SettingsStructureError,
    SettingsDirectoryError,
    SettingsGroupsError,
    replace_new_part_groups,
    SettingsFileMissingError,
    SettingsFileAccessError,
    SettingsSaveError,
    SettingsResetError,
)
from .excel_files import (
    WorkbookValidationWarning,
    WorkbookFileAccessError,
    WorkbookSaveError
)
from .csv_reader import CsvFileAccessError

def _request_missing_group(
    document_number: str,
    item_code: str,
) -> str:
    """Csoport kiválasztása egy csoport nélküli számlatételhez."""
    groups = {
        "1": "Új alkatrész",
        "2": "Bontott alkatrész",
        "3": "Munkadíj",
        "4": "Szállítási költség",
        "5": "Trailer",
        "0": ""
    }

    print(
        "\nEnnek a számlatételnek nincs csoportja:\n"
        f"Bizonylatszám: {document_number}\n"
        f"Cikkszám: {item_code}\n"
    )
    print("1 – Új alkatrész")
    print("2 – Bontott alkatrész")
    print("3 – Munkadíj")
    print("4 – Szállítási költség")
    print("5 – Trailer")
    print("0 – Egyik sem")

    while True:
        choice = input("Válassz egy csoportot: ").strip()

        if choice in groups:
            return groups[choice]
        
        print("Érvénytelen választás. Adj meg egy számot 0 és 5 között.")


def _request_directory(
    *,
    first_open: bool = False,
) -> str | None:
    """Létező riportkönyvtár bekérése."""
    while True:
        if not first_open:
            print("A könyvtár módosításának megszakítása: 0")

        value = input("A riportok könyvtárának útvonala: ").strip()

        if not first_open and value == "0":
            print("A könyvtár módosítása megszakítva.")
            return None

        if not value:
            print("Adj meg egy könyvtárútvonalat.")
            continue

        try:
            path = Path(value).expanduser()

            if not path.is_dir():
                print("A megadott útvonal nem létező könyvtár. Próbáld újra.")
                continue

            resolved_path = path.resolve()

        except (OSError, RuntimeError, ValueError) as exc:
            print(
                f"A könyvtár útvonala nem ellenőrizhető: {value!r}\n"
                f"Részletek: {exc}\n"
                "Próbáld újra."
            )
            continue

        return str(resolved_path)

def _request_new_part_groups(
    *,
    allow_empty: bool = False,
    first_open: bool = False,
) -> list[str] | None:
    """Újalkatrész-csoportok bekérése pontosvesszővel elválasztva."""
    while True:
        if not first_open:
            print("Az újalkatrész-csoportok megadásának megszakítása: 0")
        
        value = input(
            "Újalkatrész-csoportok pontosvesszővel elválasztva: "
        ).strip()

        if value == "0":
            if first_open:
                print("Adj meg csoportneveket, vagy hagyd üresen a listát.")
                continue

            print("Az újalkatrész-csoportok megadásának folyamata megszakítva.")
            return None

        if not value:
            if allow_empty:
                return []

            print("Adj meg legalább egy csoportnevet.")
            continue

        groups = [group.strip() for group in value.split(";")]

        if any(not group for group in groups):
            print("Az elválasztók között nem lehet üres csoportnév.")
            continue

        if len(groups) != len(set(groups)):
            print("Egy csoportnevet csak egyszer adj meg.")
            continue

        return groups

def _request_removed_part_groups(
) -> list[str] | None:
    """Az eltávolítandó újalkatrész-csoportok bekérése pontosvesszővel elválasztva."""
    while True:
        print("Az eltávolítás megszakítása: 0")
        value = input(
            "Eltávolítandó újalkatrész-csoportok pontosvesszővel elválasztva: "
        ).strip()

        if value == "0":
            print("Az eltávolítás megszakítva.")
            return None

        if not value:
            print("Adj meg legalább egy csoportnevet.")
            continue

        groups = [group.strip() for group in value.split(";")]

        if any(not group for group in groups):
            print("Az elválasztók között nem lehet üres csoportnév.")
            continue

        if len(groups) != len(set(groups)):
            print("Egy csoportnevet csak egyszer adj meg.")
            continue

        return groups
        

def _request_csv_path() -> str | None:
    """Létező CSV-fájl útvonalának bekérése."""
    while True:
        print("Az importálás megszakítása: 0")
        value = input("Az importálandó CSV-fájl útvonala: ").strip()

        if value == "0":
            print("Az importálás megszakítva.")
            return None

        if not value:
            print("Adj meg egy fájlútvonalat.")
            continue

        try:
            path = Path(value).expanduser()

            if not path.is_file():
                print("A megadott útvonal nem létező fájl. Próbáld újra.")
                continue

            if path.suffix.lower() != ".csv":
                print("Az importálandó fájlnak .csv fájlnak kell lennie.")
                continue

            resolved_path = path.resolve()

        except (OSError, RuntimeError, ValueError) as exc:
            print(
                f"A CSV-fájl útvonala nem ellenőrizhető: {value!r}\n"
                f"Részletek: {exc}\n"
                "Próbáld újra."
            )
            continue

        return str(resolved_path)

def _ensure_settings() -> None:
    """A beállításfájl létrehozása, ha még nem létezik."""
    path = settings_path()

    try:
        file_exists = path.exists()
    except OSError as exc:
        raise SettingsFileAccessError(
            f"A beállításfájl elérhetősége nem ellenőrizhető.\n"
            f"Fájl: {path}\n"
            f"Részletek: {exc}"
        ) from exc

    if file_exists:
        return

    print("\nElső indítás: add meg a program beállításait.")

    output_directory = _request_directory(first_open=True)
    groups = _request_new_part_groups(allow_empty=True, first_open=True)

    create_settings(
        output_directory=output_directory,
        additional_new_part_groups=groups,
    )

    print("A beállításokat elmentettük.")

def _handle_settings_error(
    error: (
        SettingsStructureError
        | SettingsDirectoryError
        | SettingsGroupsError
        | SettingsFileMissingError
        | SettingsFileAccessError
        | SettingsSaveError
    ),
) -> None:
    """A hibás beállítások javításához szükséges adatok bekérése."""
    print(f"Beállítási hiba: {error}")

    try:
        if isinstance(error, SettingsFileMissingError):
            _ensure_settings()

        elif isinstance(error, SettingsFileAccessError):
            print(
                "Ellenőrizd a fájl és a beállításmappa "
                "hozzáférési jogosultságait, majd próbáld újra."
            )

        elif isinstance(error, SettingsSaveError):
            print(
                "Ellenőrizd a beállításmappa írási jogosultságát "
                "és a rendelkezésre álló lemezterületet, "
                "majd próbáld újra."
            )

        elif isinstance(error, SettingsStructureError):
            reset()
            _ensure_settings()

        elif isinstance(error, SettingsDirectoryError):
            directory = _request_directory()

            if directory is not None:
                change_output_directory(
                    output_directory=directory,
                    validate_settings=False,
                )
                print("Az új riportkönyvtárat elmentettük.")

        elif isinstance(error, SettingsGroupsError):
            if isinstance(error.groups, list):
                print("A jelenleg tárolt elemek:")

                for group in error.groups:
                    print(f"  {group!r}")

            groups = _request_new_part_groups(allow_empty=True)

            if groups is not None:
                replace_new_part_groups(groups=groups)
                print("Az új csoportlistát elmentettük.")

    except (
        SettingsStructureError,
        SettingsDirectoryError,
        SettingsGroupsError,
        SettingsFileMissingError,
        SettingsFileAccessError,
        SettingsSaveError,
        SettingsResetError,
    ) as exc:
        print(f"A beállítások javítása nem sikerült: {exc}")

def run_console() -> None:
    """A program indítása és a terminálmenü kezelése."""
    try:
        _ensure_settings()
    except (SettingsFileAccessError, SettingsSaveError) as exc:
        _handle_settings_error(exc)
        return

    try:
        while True:
            print("\n=== AxelPro riport ===")
            print("1 – Importálás")
            print("2 – Újalkatrész-csoport hozzáadása")
            print("3 – Újalkatrész-csoport eltávolítása")
            print("4 – Riport mappa módosítása")
            print("5 – Jelenlegi beállítások megtekintése")
            print("6 – Alaphelyzetbe állítás és bezárás")
            print("0 – Kilépés")

            choice = input("Válassz egy műveletet: ").strip()

            if choice == "1":
                csv_path = _request_csv_path()

                if csv_path is None:
                    continue

                try:
                    with warnings.catch_warnings(record=True) as collected_warnings:
                        warnings.simplefilter("once", WorkbookValidationWarning)

                        start = perf_counter()

                        try:
                            import_csv(csv_path=csv_path, request_missing_group=_request_missing_group)
                        finally:
                            elapsed = perf_counter() - start
                            print(f"Az importálás futási ideje: {elapsed:.2f} másodperc")
                except (
                    SettingsStructureError,
                    SettingsDirectoryError,
                    SettingsGroupsError,
                    SettingsFileMissingError,
                    SettingsFileAccessError,
                    SettingsSaveError,
                ) as exc:
                    _handle_settings_error(exc)
                    continue

                except ValueError as exc:
                    print(f"Importálási hiba: {exc}")
                    continue

                except CsvFileAccessError as exc:
                    print(f"Importálási hiba: {exc}")
                    continue

                except WorkbookFileAccessError as exc:
                    print(f"Importálási hiba: {exc}")
                    continue

                except WorkbookSaveError as exc:
                    print(f"Importálási hiba: {exc}")
                    continue

                print("Az importálási folyamat befejeződött.")

                first_warning = next(
                    (
                        warning
                        for warning in collected_warnings
                        if issubclass(
                            warning.category,
                            WorkbookValidationWarning
                        )
                    ),
                    None
                )

                if first_warning is not None:
                    print(f"\nFigyelmeztetés: {first_warning.message}")

            elif choice == "2":
                groups = _request_new_part_groups()

                if groups is None:
                    continue

                try:
                    add_new_part_groups(groups=groups)
                except (
                    SettingsStructureError,
                    SettingsDirectoryError,
                    SettingsGroupsError,
                    SettingsFileMissingError,
                    SettingsFileAccessError,
                    SettingsSaveError,
                ) as exc:
                    _handle_settings_error(exc)
                    continue

                print("A csoportlistát elmentettük.")

            elif choice == "3":
                groups = _request_removed_part_groups()

                if groups is None:
                    continue

                try:
                    remove_new_part_groups(groups=groups)
                except (
                    SettingsStructureError,
                    SettingsDirectoryError,
                    SettingsGroupsError,
                    SettingsFileMissingError,
                    SettingsFileAccessError,
                    SettingsSaveError,
                ) as exc:
                    _handle_settings_error(exc)
                    continue

                print("A csoportlistát elmentettük.")

            elif choice == "4":
                directory = _request_directory()

                if directory is None:
                    continue

                try:
                    change_output_directory(output_directory=directory)
                except (
                    SettingsStructureError,
                    SettingsDirectoryError,
                    SettingsGroupsError,
                    SettingsFileMissingError,
                    SettingsFileAccessError,
                    SettingsSaveError,
                ) as exc:
                    _handle_settings_error(exc)
                    continue

                print("Az új riportkönyvtárat elmentettük.")

            elif choice == "5":
                try:
                    settings = load_settings()
                except (
                    SettingsStructureError,
                    SettingsDirectoryError,
                    SettingsGroupsError,
                    SettingsFileMissingError,
                    SettingsFileAccessError,
                    SettingsSaveError,
                ) as exc:
                    _handle_settings_error(exc)
                    continue

                print(f"\nRiportkönyvtár: {settings.output_directory}")
                print("További újalkatrész-csoportok:")

                if settings.additional_new_part_groups:
                    for group in settings.additional_new_part_groups:
                        print(f"  {group}")
                else:
                    print("  Nincsenek további csoportok.")

            elif choice == "6":
                try:
                    reset()
                except SettingsResetError as exc:
                    print(f"Az alaphelyzetbe állítás nem sikerült: {exc}")
                    continue

                return

            elif choice == "0":
                return

            else:
                print("Érvénytelen választás. Próbáld újra.")
    except KeyboardInterrupt:
        return



if __name__ == "__main__":
    run_console()