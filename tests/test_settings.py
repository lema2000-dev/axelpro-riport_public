import json
from pathlib import Path
from sales_report import settings


def test_create_load_add_remove_and_change_directory(settings_location, tmp_path):
    created = settings.create_settings(str(tmp_path), ["Filters"])
    assert settings.load_settings() == created
    settings.add_new_part_groups(["Filters", "Electronics", "Electronics"])
    settings.remove_new_part_groups(["Filters", "Ismeretlen"])
    new_directory = tmp_path / "new_reports"
    new_directory.mkdir()
    settings.change_output_directory(str(new_directory))
    loaded = settings.load_settings()
    assert loaded.output_directory == str(new_directory)
    assert loaded.additional_new_part_groups == ["Electronics"]
    assert list(settings_location.parent.iterdir()) == [settings_location]
    assert set(json.loads(settings_location.read_text())) == {
        "output_directory",
        "additional_new_part_groups",
    }


def test_reset_deletes_settings_and_next_start_can_create_new(
    settings_location, tmp_path
):
    settings.create_settings(str(tmp_path), ["Filters"])
    settings.reset()
    assert not settings_location.exists()
    assert not settings_location.parent.exists()
    settings.create_settings(str(tmp_path), [])
    assert settings.load_settings().output_directory == str(tmp_path)
    assert settings.load_settings().additional_new_part_groups == []


def test_legacy_template_field_is_ignored_and_removed_on_next_save(
    settings_location, tmp_path
):
    settings_location.parent.mkdir()
    settings_location.write_text(
        json.dumps(
            {
                "output_directory": str(tmp_path),
                "template_sheet": "old_template.xlsx",
                "additional_new_part_groups": ["Filters"],
            }
        ),
        encoding="utf-8",
    )
    assert settings.load_settings() == settings.Settings(str(tmp_path), ["Filters"])
    settings.change_output_directory(str(tmp_path))
    assert "template_sheet" not in json.loads(settings_location.read_text())


def test_settings_location_uses_appdata(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert settings.settings_path() == tmp_path / "sales_report" / "settings.json"


def test_settings_location_falls_back_to_home(monkeypatch, tmp_path):
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.setattr(Path, "home", lambda: tmp_path)
    assert (
        settings.settings_path()
        == tmp_path / ".config" / "sales_report" / "settings.json"
    )
