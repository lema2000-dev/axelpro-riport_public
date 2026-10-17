"""Test settings errors."""

import json
from pathlib import Path
import pytest
from sales_report import settings as s


def write_data(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data), encoding="utf-8")


@pytest.mark.parametrize(
    "data",
    [
        [],
        None,
        42,
        {},
        {"output_directory": "x"},
        {"output_directory": "x", "additional_new_part_groups": [], "extra": 1},
    ],
)
def test_invalid_structure(settings_location, data):
    write_data(settings_location, data)
    with pytest.raises(s.SettingsStructureError):
        s.load_settings()


@pytest.mark.parametrize("content", [b"{broken", b"\xff\xfe"])
def test_invalid_json_or_encoding(settings_location, content):
    settings_location.parent.mkdir()
    settings_location.write_bytes(content)
    with pytest.raises(s.SettingsStructureError) as error:
        s.load_settings()
    assert error.value.__cause__ is not None


@pytest.mark.parametrize("directory", [None, 12, "", " ", "nonexistent/reports"])
def test_invalid_directory(settings_location, directory):
    write_data(
        settings_location,
        {"output_directory": directory, "additional_new_part_groups": []},
    )
    with pytest.raises(s.SettingsDirectoryError):
        s.load_settings()


@pytest.mark.parametrize(
    "groups", [None, "Group", [None], [12], [""], [" "], ["Group", " group "]]
)
def test_invalid_groups_are_available_on_exception(settings_location, tmp_path, groups):
    write_data(
        settings_location,
        {"output_directory": str(tmp_path), "additional_new_part_groups": groups},
    )
    with pytest.raises(s.SettingsGroupsError) as error:
        s.load_settings()
    assert error.value.groups == groups


def test_missing_file(settings_location):
    with pytest.raises(s.SettingsFileMissingError) as error:
        s.load_settings()
    assert str(settings_location) in str(error.value)
    assert isinstance(error.value.__cause__, FileNotFoundError)


def test_read_permission_error_is_not_structure_error(settings_location, monkeypatch):
    original = Path.open

    def deny(path, *args, **kwargs):
        if path == settings_location:
            raise PermissionError("denied")
        return original(path, *args, **kwargs)

    monkeypatch.setattr(Path, "open", deny)
    with pytest.raises(s.SettingsFileAccessError) as error:
        s.load_settings()
    assert isinstance(error.value.__cause__, PermissionError)


def test_directory_and_group_repairs_bypass_invalid_old_values(
    settings_location, tmp_path
):
    write_data(
        settings_location,
        {"output_directory": "missing", "additional_new_part_groups": []},
    )
    s.change_output_directory(str(tmp_path), validate_settings=False)
    assert s.load_settings().output_directory == str(tmp_path)
    write_data(
        settings_location,
        {"output_directory": str(tmp_path), "additional_new_part_groups": "bad"},
    )
    s.replace_new_part_groups(["New"])
    assert s.load_settings().additional_new_part_groups == ["New"]


@pytest.mark.parametrize("stage", ["mkdir", "temporary", "write", "replace"])
def test_save_failure_keeps_old_file_and_original_cause(
    settings_location, tmp_path, monkeypatch, stage
):
    s.create_settings(str(tmp_path), ["Old"])
    before = settings_location.read_bytes()
    failure = PermissionError("save denied")

    def fail(*args, **kwargs):
        raise failure

    if stage == "mkdir":
        monkeypatch.setattr(Path, "mkdir", fail)
    elif stage == "temporary":
        monkeypatch.setattr(s, "NamedTemporaryFile", fail)
    elif stage == "write":
        monkeypatch.setattr(s.json, "dump", fail)
    else:
        monkeypatch.setattr(Path, "replace", fail)
    with pytest.raises(s.SettingsSaveError) as error:
        s.create_settings(str(tmp_path), ["New"])
    assert error.value.__cause__ is failure
    assert settings_location.read_bytes() == before
    assert list(settings_location.parent.iterdir()) == [settings_location]


def test_failed_cleanup_does_not_mask_save_failure(
    settings_location, tmp_path, monkeypatch
):
    original_failure = PermissionError("replace denied")

    def replace(*args, **kwargs):
        raise original_failure

    def unlink(*args, **kwargs):
        raise OSError("cleanup denied")

    monkeypatch.setattr(Path, "replace", replace)
    monkeypatch.setattr(Path, "unlink", unlink)
    with pytest.raises(s.SettingsSaveError) as error:
        s.create_settings(str(tmp_path), [])
    assert error.value.__cause__ is original_failure


def test_reset_missing_file_and_directory_are_allowed(settings_location):
    s.reset()
    assert not settings_location.exists()


def test_reset_file_failure_preserves_settings(
    settings_location, tmp_path, monkeypatch
):
    s.create_settings(str(tmp_path), [])

    def deny(*args, **kwargs):
        raise PermissionError("delete denied")

    monkeypatch.setattr(Path, "unlink", deny)
    with pytest.raises(s.SettingsResetError, match="Could not delete the settings file"):
        s.reset()
    assert settings_location.exists()


def test_reset_nonempty_directory_reports_partial_result(settings_location, tmp_path):
    s.create_settings(str(tmp_path), [])
    other = settings_location.parent / "other.txt"
    other.write_text("keep")
    with pytest.raises(s.SettingsResetError, match="no longer exists"):
        s.reset()
    assert not settings_location.exists()
    assert other.read_text() == "keep"
