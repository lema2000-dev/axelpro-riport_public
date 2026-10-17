import os
import json

from dataclasses import dataclass, field, asdict
from pathlib import Path
from tempfile import NamedTemporaryFile

class SettingsStructureError(ValueError):
    """Settings structure error."""
class SettingsDirectoryError(ValueError):
    """Settings directory error."""
class SettingsGroupsError(ValueError):
    """Settings groups error."""
    def __init__(self, message: str, groups: object) -> None:
        super().__init__(message)
        self.groups = groups

class SettingsFileMissingError(FileNotFoundError):
    """Settings file missing error."""
    
class SettingsFileAccessError(OSError):
    """Settings file access error."""

class SettingsSaveError(OSError):
    """Settings save error."""

class SettingsResetError(OSError):
    """Settings reset error."""

@dataclass
class Settings:
    output_directory: str = ""
    additional_new_part_groups: list[str] = field(default_factory=list)

def settings_path() -> Path:
    """Return the per-user settings path for the current operating system."""
    appdata = os.environ.get("APPDATA")

    if appdata:
        root = Path(appdata)
    else:
        root = Path.home() / ".config"

    return root / "sales_report" / "settings.json"

def _save_settings(settings: Settings) -> None:
    """Save settings."""
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
            f"Could not save settings.\n"
            f"File: {path}\n"
            f"Details: {exc}"
        ) from exc

    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass

def _read_settings_data() -> dict:
    """Read settings JSON without validating individual field contents."""
    path = settings_path()

    try:
        with path.open(mode="r", encoding="utf-8") as settings_file:
            data = json.load(settings_file)

    except FileNotFoundError as exc:
        raise SettingsFileMissingError(
            f"Settings file not found.\nFile: {path}"
        ) from exc

    except OSError as exc:
        raise SettingsFileAccessError(
            f"Cannot read the settings file.\n"
            f"File: {path}\n"
            f"Details: {exc}"
        ) from exc

    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise SettingsStructureError(
            f"The settings file is not valid UTF-8 JSON.\n"
            f"File: {path}"
        ) from exc

    if not isinstance(data, dict):
        raise SettingsStructureError(
            "The settings file must contain a JSON object."
        )

    data.pop("template_sheet", None)

    return data

def load_settings() -> Settings:
    """Read settings and validate their structure, output folder and categories."""
    data = _read_settings_data()

    expected_keys = {
        "output_directory",
        "additional_new_part_groups",
    }

    if set(data) != expected_keys:
        raise SettingsStructureError(
            "The settings file has invalid fields."
        )

    directory = data["output_directory"]

    if not isinstance(directory, str) or not directory.strip():
        raise SettingsDirectoryError(
            "The report folder path must be a nonempty string."
        )

    try:
        directory_exists = Path(directory).expanduser().is_dir()
    except (OSError, ValueError) as exc:
        raise SettingsDirectoryError(
            "Cannot check the report folder path."
        ) from exc

    if not directory_exists:
        raise SettingsDirectoryError(
            "The report folder does not exist."
        )

    groups = data["additional_new_part_groups"]

    if not isinstance(groups, list):
        raise SettingsGroupsError(
            "New-part categories must be stored in a list.",
            groups=groups
        )

    if any(
        not isinstance(group, str) or not group.strip()
        for group in groups
    ):
        raise SettingsGroupsError(
            "Each category name must be a nonempty string.",
            groups=groups
        )

    normalized_groups = [group.strip().casefold() for group in groups]

    if len(normalized_groups) != len(set(normalized_groups)):
        raise SettingsGroupsError(
            "The new-part category list contains duplicate names.",
            groups=groups
        )

    return Settings(**data)

def change_output_directory(
    output_directory: str,
    *,
    validate_settings: bool = True,
) -> None:
    """Change output directory."""
    if validate_settings:
        settings = load_settings()
    else:
        settings = Settings(**_read_settings_data())

    settings.output_directory = output_directory
    _save_settings(settings)

def add_new_part_groups(groups: list[str]) -> None:
    """Add new part groups."""
    settings = load_settings()

    for group in groups:
        if group not in settings.additional_new_part_groups:
            settings.additional_new_part_groups.append(group)

    _save_settings(settings)

def remove_new_part_groups(groups: list[str]) -> None:
    """Remove new part groups."""
    settings = load_settings()

    settings.additional_new_part_groups = [
        group for group in settings.additional_new_part_groups if group not in groups
    ]

    _save_settings(settings)

def create_settings(
    output_directory: str,
    additional_new_part_groups: list[str],
) -> Settings:
    """Create settings."""
    settings = Settings(
        output_directory=output_directory,
        additional_new_part_groups=additional_new_part_groups,
    )

    _save_settings(settings)

    return settings

def reset() -> None:
    """Reset."""
    path = settings_path()

    try:
        path.unlink(missing_ok=True)
    except OSError as exc:
        raise SettingsResetError(
            f"Could not delete the settings file.\n"
            f"File: {path}\n"
            f"Details: {exc}"
        ) from exc

    try:
        path.parent.rmdir()
    except FileNotFoundError:
        pass
    except OSError as exc:
        raise SettingsResetError(
            "The settings file no longer exists, "
            "but the settings folder could not be removed.\n"
            f"Folder: {path.parent}\n"
            f"Details: {exc}"
        ) from exc

def replace_new_part_groups(groups: list[str]) -> None:
    """Replace new part groups."""
    data = _read_settings_data()
    data["additional_new_part_groups"] = groups

    _save_settings(Settings(**data))