"""Console interface for sales report imports and application settings."""

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
    """Ask the user to classify an invoice item with an empty group."""
    groups = {
        "1": "New parts",
        "2": "Used parts",
        "3": "Labor",
        "4": "Shipping",
        "5": "Trailer",
        "0": ""
    }

    print(
        "\nThis invoice item has no group:\n"
        f"Document number: {document_number}\n"
        f"Item code: {item_code}\n"
    )
    print("1 – New parts")
    print("2 – Used parts")
    print("3 – Labor")
    print("4 – Shipping")
    print("5 – Trailer")
    print("0 – None of these")

    while True:
        choice = input("Select a group: ").strip()

        if choice in groups:
            return groups[choice]
        
        print("Invalid selection. Enter a number from 0 to 5.")


def _request_directory(
    *,
    first_open: bool = False,
) -> str | None:
    """Request directory."""
    while True:
        if not first_open:
            print("Cancel folder change: 0")

        value = input("Report folder path: ").strip()

        if not first_open and value == "0":
            print("Folder change cancelled.")
            return None

        if not value:
            print("Enter a folder path.")
            continue

        try:
            path = Path(value).expanduser()

            if not path.is_dir():
                print("The path is not an existing folder. Try again.")
                continue

            resolved_path = path.resolve()

        except (OSError, RuntimeError, ValueError) as exc:
            print(
                f"Cannot check the folder path: {value!r}\n"
                f"Details: {exc}\n"
                "Try again."
            )
            continue

        return str(resolved_path)

def _request_new_part_groups(
    *,
    allow_empty: bool = False,
    first_open: bool = False,
) -> list[str] | None:
    """Request new part groups."""
    while True:
        if not first_open:
            print("Cancel category entry: 0")
        
        value = input(
            "New-part categories separated by semicolons: "
        ).strip()

        if value == "0":
            if first_open:
                print("Enter category names or leave the list empty.")
                continue

            print("Category entry cancelled.")
            return None

        if not value:
            if allow_empty:
                return []

            print("Enter at least one category name.")
            continue

        groups = [group.strip() for group in value.split(";")]

        if any(not group for group in groups):
            print("Category names between separators must not be empty.")
            continue

        if len(groups) != len(set(groups)):
            print("Enter each category name only once.")
            continue

        return groups

def _request_removed_part_groups(
) -> list[str] | None:
    """Request removed part groups."""
    while True:
        print("Cancel removal: 0")
        value = input(
            "Categories to remove, separated by semicolons: "
        ).strip()

        if value == "0":
            print("Removal cancelled.")
            return None

        if not value:
            print("Enter at least one category name.")
            continue

        groups = [group.strip() for group in value.split(";")]

        if any(not group for group in groups):
            print("Category names between separators must not be empty.")
            continue

        if len(groups) != len(set(groups)):
            print("Enter each category name only once.")
            continue

        return groups
        

def _request_csv_path() -> str | None:
    """Request csv path."""
    while True:
        print("Cancel import: 0")
        value = input("Path to the CSV file to import: ").strip()

        if value == "0":
            print("Import cancelled.")
            return None

        if not value:
            print("Enter a file path.")
            continue

        try:
            path = Path(value).expanduser()

            if not path.is_file():
                print("The path is not an existing file. Try again.")
                continue

            if path.suffix.lower() != ".csv":
                print("The import file must have the .csv extension.")
                continue

            resolved_path = path.resolve()

        except (OSError, RuntimeError, ValueError) as exc:
            print(
                f"Cannot check the CSV file path: {value!r}\n"
                f"Details: {exc}\n"
                "Try again."
            )
            continue

        return str(resolved_path)

def _ensure_settings() -> None:
    """Ensure settings."""
    path = settings_path()

    try:
        file_exists = path.exists()
    except OSError as exc:
        raise SettingsFileAccessError(
            f"Cannot check settings file availability.\n"
            f"File: {path}\n"
            f"Details: {exc}"
        ) from exc

    if file_exists:
        return

    print("\nFirst launch: enter the application settings.")

    output_directory = _request_directory(first_open=True)
    groups = _request_new_part_groups(allow_empty=True, first_open=True)

    create_settings(
        output_directory=output_directory,
        additional_new_part_groups=groups,
    )

    print("Settings saved.")

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
    """Handle settings error."""
    print(f"Settings error: {error}")

    try:
        if isinstance(error, SettingsFileMissingError):
            _ensure_settings()

        elif isinstance(error, SettingsFileAccessError):
            print(
                "Check the file and settings folder "
                "access permissions, then try again."
            )

        elif isinstance(error, SettingsSaveError):
            print(
                "Check write permissions for the settings folder "
                "and the available disk space, "
                "then try again."
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
                print("New report folder saved.")

        elif isinstance(error, SettingsGroupsError):
            if isinstance(error.groups, list):
                print("Currently stored items:")

                for group in error.groups:
                    print(f"  {group!r}")

            groups = _request_new_part_groups(allow_empty=True)

            if groups is not None:
                replace_new_part_groups(groups=groups)
                print("New category list saved.")

    except (
        SettingsStructureError,
        SettingsDirectoryError,
        SettingsGroupsError,
        SettingsFileMissingError,
        SettingsFileAccessError,
        SettingsSaveError,
        SettingsResetError,
    ) as exc:
        print(f"Could not repair settings: {exc}")

def run_console() -> None:
    """Ensure settings exist and process console menu actions until exit."""
    try:
        _ensure_settings()
    except (SettingsFileAccessError, SettingsSaveError) as exc:
        _handle_settings_error(exc)
        return

    try:
        while True:
            print("\n=== Sales Report ===")
            print("1 – Import")
            print("2 – Add new-part categories")
            print("3 – Remove new-part categories")
            print("4 – Change report folder")
            print("5 – View current settings")
            print("6 – Reset and close")
            print("0 – Exit")

            choice = input("Select an action: ").strip()

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
                            print(f"Import duration: {elapsed:.2f} seconds")
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
                    print(f"Import error: {exc}")
                    continue

                except CsvFileAccessError as exc:
                    print(f"Import error: {exc}")
                    continue

                except WorkbookFileAccessError as exc:
                    print(f"Import error: {exc}")
                    continue

                except WorkbookSaveError as exc:
                    print(f"Import error: {exc}")
                    continue

                print("Import process completed.")

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
                    print(f"\nWarning: {first_warning.message}")

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

                print("Category list saved.")

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

                print("Category list saved.")

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

                print("New report folder saved.")

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

                print(f"\nReport folder: {settings.output_directory}")
                print("Additional new-part categories:")

                if settings.additional_new_part_groups:
                    for group in settings.additional_new_part_groups:
                        print(f"  {group}")
                else:
                    print("  No additional categories.")

            elif choice == "6":
                try:
                    reset()
                except SettingsResetError as exc:
                    print(f"Could not reset the application: {exc}")
                    continue

                return

            elif choice == "0":
                return

            else:
                print("Invalid selection. Try again.")
    except KeyboardInterrupt:
        return



if __name__ == "__main__":
    run_console()