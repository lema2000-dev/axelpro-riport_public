import os
import json

from dataclasses import dataclass, field, asdict
from pathlib import Path
from tempfile import NamedTemporaryFile

class SettingsStructureError(ValueError):
    """A beállításfájl szerkezete hibás."""
class SettingsDirectoryError(ValueError):
    """A beállított riportkönyvtár hibás."""
class SettingsGroupsError(ValueError):
    """Az újalkatrész-csoportok listája hibás."""
    def __init__(self, message: str, groups: object) -> None:
        super().__init__(message)
        self.groups = groups

class SettingsFileMissingError(FileNotFoundError):
    """A beállításfájl nem található."""
    
class SettingsFileAccessError(OSError):
    """A beállításfájl olvasása sikertelen."""

class SettingsSaveError(OSError):
    """A beállításfájl mentése sikertelen."""

class SettingsResetError(OSError):
    """A beállítások alaphelyzetbe állytása nem fejeződött be."""

@dataclass
class Settings:
    output_directory: str = ""
    additional_new_part_groups: list[str] = field(default_factory=list)

def settings_path() -> Path:
    """A helyi beállításfájl elérési útja."""
    appdata = os.environ.get("APPDATA")

    if appdata:
        root = Path(appdata)
    else:
        root = Path.home() / ".config"

    return root / "axelpro_riport" / "settings.json"

def _save_settings(settings: Settings) -> None:
    """A beállítások mentése ideiglenes fájlon keresztül."""
    path = settings_path()
    temporary_path = None

    try:
        path.parent.mkdir(parents=True, exist_ok=True)

        with NamedTemporaryFile(
            dir=path.parent,
            mode="w",
            encoding="utf-8",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

            json.dump(
                asdict(settings),
                temporary_file,
                ensure_ascii=False,
                indent=2,
            )

        temporary_path.replace(path)

    except OSError as exc:
        raise SettingsSaveError(
            f"A beállítások mentése nem sikerült.\n"
            f"Fájl: {path}\n"
            f"Részletek: {exc}"
        ) from exc

    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass

def _read_settings_data() -> dict:
    """A beállítások beolvasása a mezők tartalmi ellenőrzése nélkül."""
    path = settings_path()

    try:
        with path.open(mode="r", encoding="utf-8") as settings_file:
            data = json.load(settings_file)

    except FileNotFoundError as exc:
        raise SettingsFileMissingError(
            f"A beállításfájl nem található.\nFájl: {path}"
        ) from exc

    except OSError as exc:
        raise SettingsFileAccessError(
            f"A beállításfájl nem olvasható.\n"
            f"Fájl: {path}\n"
            f"Részletek: {exc}"
        ) from exc

    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise SettingsStructureError(
            f"A beállításfájl nem megfelelő UTF-8 JSON.\n"
            f"Fájl: {path}"
        ) from exc

    if not isinstance(data, dict):
        raise SettingsStructureError(
            "A beállításfájlnak JSON-objektumot kell tartalmaznia."
        )

    data.pop("template_sheet", None)

    return data

def load_settings() -> Settings:
    """A mentett beállítások betöltése és ellenőrzése."""
    data = _read_settings_data()

    expected_keys = {
        "output_directory",
        "additional_new_part_groups",
    }

    if set(data) != expected_keys:
        raise SettingsStructureError(
            "A beállításfájl mezői nem megfelelőek."
        )

    directory = data["output_directory"]

    if not isinstance(directory, str) or not directory.strip():
        raise SettingsDirectoryError(
            "A riportkönyvtár útvonala nem üres szöveg kell legyen."
        )

    try:
        directory_exists = Path(directory).expanduser().is_dir()
    except (OSError, ValueError) as exc:
        raise SettingsDirectoryError(
            "A riportkönyvtár útvonala nem ellenőrizhető."
        ) from exc

    if not directory_exists:
        raise SettingsDirectoryError(
            "A megadott riportkönyvtár nem létező könyvtár."
        )

    groups = data["additional_new_part_groups"]

    if not isinstance(groups, list):
        raise SettingsGroupsError(
            "Az újalkatrész-csoportokat listában kell tárolni.",
            groups=groups
        )

    if any(
        not isinstance(group, str) or not group.strip()
        for group in groups
    ):
        raise SettingsGroupsError(
            "Minden csoportnévnek nem üres szövegnek kell lennie.",
            groups=groups
        )

    normalized_groups = [group.strip().casefold() for group in groups]

    if len(normalized_groups) != len(set(normalized_groups)):
        raise SettingsGroupsError(
            "Az újalkatrész-csoportok között ismétlődő nevek vannak.",
            groups=groups
        )

    return Settings(**data)

def change_output_directory(
    output_directory: str,
    *,
    validate_settings: bool = True,
) -> None:
    """A riportkönyvtár módosítása, opcionális előzetes ellenőrzéssel."""
    if validate_settings:
        settings = load_settings()
    else:
        settings = Settings(**_read_settings_data())

    settings.output_directory = output_directory
    _save_settings(settings)

def add_new_part_groups(groups: list[str]) -> None:
    """Az újalkatrész-csoportok listájának bővítése."""
    settings = load_settings()

    for group in groups:
        if group not in settings.additional_new_part_groups:
            settings.additional_new_part_groups.append(group)

    _save_settings(settings)

def remove_new_part_groups(groups: list[str]) -> None:
    """Az újalkatrész-csoportok listájának csökkentése."""
    settings = load_settings()

    settings.additional_new_part_groups = [
        group for group in settings.additional_new_part_groups if group not in groups
    ]

    _save_settings(settings)

def create_settings(
    output_directory: str,
    additional_new_part_groups: list[str],
) -> Settings:
    """A megadott beállítások létrehozása és mentése."""
    settings = Settings(
        output_directory=output_directory,
        additional_new_part_groups=additional_new_part_groups,
    )

    _save_settings(settings)

    return settings

def reset() -> None:
    """A beállításfájl és az üres beállításmappa törlése."""
    path = settings_path()

    try:
        path.unlink(missing_ok=True)
    except OSError as exc:
        raise SettingsResetError(
            f"A beállításfájl törlése nem sikerült.\n"
            f"Fájl: {path}\n"
            f"Részletek: {exc}"
        ) from exc

    try:
        path.parent.rmdir()
    except FileNotFoundError:
        pass
    except OSError as exc:
        raise SettingsResetError(
            "A beállításfájl már nem létezik, "
            "de a beállításmappa eltávolítása nem sikerült.\n"
            f"Mappa: {path.parent}\n"
            f"Részletek: {exc}"
        ) from exc

def replace_new_part_groups(groups: list[str]) -> None:
    """A további újalkatrész-csoportok teljes listájának cseréje."""
    data = _read_settings_data()
    data["additional_new_part_groups"] = groups

    _save_settings(Settings(**data))