"""Hibák konzolos megjelenítése és visszatérés a menübe."""

from pathlib import Path
import warnings
import pytest
from axelpro_riport import console as c, settings as s
from axelpro_riport.csv_reader import CsvFileAccessError
from axelpro_riport.excel_files import (
    WorkbookFileAccessError,
    WorkbookSaveError,
    WorkbookValidationError,
    WorkbookValidationWarning,
)


def answers(monkeypatch, values):
    values = iter(values)
    monkeypatch.setattr("builtins.input", lambda prompt: next(values))


@pytest.mark.parametrize(
    "error",
    [
        ValueError("bad CSV row 3"),
        CsvFileAccessError("CSV denied"),
        WorkbookFileAccessError("report denied"),
        WorkbookSaveError("save failed"),
        WorkbookValidationError("bad workbook"),
    ],
)
def test_import_errors_return_to_menu_without_success(
    tmp_path, settings_location, monkeypatch, capsys, error
):
    s.create_settings(str(tmp_path), [])
    path = tmp_path / "input.csv"
    path.touch()

    def fail(**kwargs):
        raise error

    monkeypatch.setattr(c, "import_csv", fail)
    answers(monkeypatch, ["1", str(path), "0"])
    c.run_console()
    output = capsys.readouterr().out
    assert str(error) in output
    assert "befejeződött" not in output
    assert output.count("=== AxelPro riport ===") == 2


def test_import_warnings_print_once_after_import_returns(
    tmp_path, settings_location, monkeypatch, capsys
):
    s.create_settings(str(tmp_path), [])
    path = tmp_path / "input.csv"
    path.touch()

    def import_file(**kwargs):
        assert "FORMATTING TEST" not in capsys.readouterr().out
        warnings.warn("FORMATTING TEST", WorkbookValidationWarning)
        warnings.warn("SECOND FORMATTING TEST", WorkbookValidationWarning)
        print("SAVE FINISHED")

    monkeypatch.setattr(c, "import_csv", import_file)
    answers(monkeypatch, ["1", str(path), "0"])
    c.run_console()
    output = capsys.readouterr().out
    assert output.count("FORMATTING TEST") == 1
    assert "SECOND FORMATTING TEST" not in output
    assert output.index("SAVE FINISHED") < output.index(
        "Figyelmeztetés: FORMATTING TEST"
    )


@pytest.mark.parametrize(
    "error",
    [s.SettingsFileAccessError("access failed"), s.SettingsSaveError("save failed")],
)
def test_startup_failure_exits_before_menu(monkeypatch, capsys, error):
    def fail():
        raise error

    monkeypatch.setattr(c, "_ensure_settings", fail)
    c.run_console()
    output = capsys.readouterr().out
    assert str(error) in output and "=== AxelPro riport ===" not in output


def test_missing_settings_during_operation_are_recreated(
    tmp_path, settings_location, monkeypatch, capsys
):
    s.create_settings(str(tmp_path), [])

    def disappear(**kwargs):
        settings_location.unlink()
        raise s.SettingsFileMissingError("vanished")

    monkeypatch.setattr(c, "add_new_part_groups", disappear)
    answers(monkeypatch, ["2", "New", str(tmp_path), "", "0"])
    c.run_console()
    assert s.load_settings() == s.Settings(str(tmp_path), [])
    assert "vanished" in capsys.readouterr().out


@pytest.mark.parametrize("stored_groups", ["not a list", ["Good", ""]])
def test_invalid_groups_display_and_replace(
    tmp_path, settings_location, monkeypatch, capsys, stored_groups
):
    s.create_settings(str(tmp_path), [])
    import json

    settings_location.write_text(
        json.dumps(
            {
                "output_directory": str(tmp_path),
                "additional_new_part_groups": stored_groups,
            }
        )
    )
    answers(monkeypatch, ["5", "Correct", "0"])
    c.run_console()
    assert s.load_settings().additional_new_part_groups == ["Correct"]
    output = capsys.readouterr().out
    assert ("A jelenleg tárolt elemek:" in output) == isinstance(stored_groups, list)


def test_invalid_directory_repair(tmp_path, settings_location, monkeypatch):
    s.create_settings(str(tmp_path / "missing"), [])
    answers(monkeypatch, ["5", str(tmp_path), "0"])
    c.run_console()
    assert s.load_settings().output_directory == str(tmp_path)


def test_structure_error_resets_and_recreates(tmp_path, settings_location, monkeypatch):
    s.create_settings(str(tmp_path), [])
    settings_location.write_text("{broken")
    answers(monkeypatch, ["5", str(tmp_path), "New", "0"])
    c.run_console()
    assert s.load_settings().additional_new_part_groups == ["New"]


@pytest.mark.parametrize(
    "error",
    [
        s.SettingsSaveError("repair save failed"),
        s.SettingsResetError("repair delete failed"),
    ],
)
def test_failure_during_repair_is_printed_without_recursive_retry(
    monkeypatch, capsys, error
):
    def fail():
        raise error

    monkeypatch.setattr(c, "reset", fail)
    c._handle_settings_error(s.SettingsStructureError("bad structure"))
    output = capsys.readouterr().out
    assert "A beállítások javítása nem sikerült" in output
    assert str(error) in output


def test_reset_failure_returns_to_menu(
    tmp_path, settings_location, monkeypatch, capsys
):
    s.create_settings(str(tmp_path), [])

    def fail():
        raise s.SettingsResetError("reset denied")

    monkeypatch.setattr(c, "reset", fail)
    answers(monkeypatch, ["6", "0"])
    c.run_console()
    assert settings_location.exists()
    assert capsys.readouterr().out.count("=== AxelPro riport ===") == 2


def test_view_settings_does_not_write_file(
    tmp_path, settings_location, monkeypatch, capsys
):
    s.create_settings(str(tmp_path), ["Visible group"])
    before = settings_location.read_bytes()
    answers(monkeypatch, ["5", "0"])
    c.run_console()
    output = capsys.readouterr().out
    assert str(tmp_path) in output and "Visible group" in output
    assert settings_location.read_bytes() == before


@pytest.mark.parametrize("kind,method", [("directory", "is_dir"), ("csv", "resolve")])
def test_path_failure_reprompts_and_can_cancel(monkeypatch, capsys, kind, method):
    if kind == "csv":
        monkeypatch.setattr(Path, "is_file", lambda self: True)

    def fail(*args):
        raise PermissionError("path denied")

    monkeypatch.setattr(Path, method, fail)
    answers(monkeypatch, ["bad.csv", "0"])
    request = c._request_directory if kind == "directory" else c._request_csv_path
    assert request() is None
    assert "path denied" in capsys.readouterr().out


@pytest.mark.parametrize("answer,expected", [
    ("1", "Új alkatrész"), ("2", "Bontott alkatrész"),
    ("3", "Munkadíj"), ("4", "Szállítási költség"),
    ("5", "Trailer"), ("0", ""), (" 1 ", "Új alkatrész"),
])
def test_ungrouped_invoice_console_decision(monkeypatch, capsys, answer, expected):
    answers(monkeypatch, ["wrong", answer])
    assert c._request_missing_group("DOC-1", "ITEM-1") == expected
    output = capsys.readouterr().out
    assert "DOC-1" in output and "ITEM-1" in output


def test_console_passes_ungrouped_callback(tmp_path, settings_location, monkeypatch):
    from unittest.mock import Mock
    s.create_settings(str(tmp_path), [])
    path = tmp_path / "input.csv"
    path.touch()
    importer = Mock()
    monkeypatch.setattr(c, "import_csv", importer)
    answers(monkeypatch, ["1", str(path), "0"])
    c.run_console()
    assert importer.call_args.kwargs["request_missing_group"] is c._request_missing_group
