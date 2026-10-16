import pytest
from axelpro_riport import console, settings


def answers(monkeypatch, values):
    responses = iter(values)
    monkeypatch.setattr("builtins.input", lambda prompt: next(responses))


def test_first_start_reprompts_directory_and_groups(
    monkeypatch, tmp_path, settings_location, capsys
):
    wrong = tmp_path / "wrong.txt"
    wrong.touch()
    answers(
        monkeypatch,
        ["", str(wrong), str(tmp_path), "A;;B", "A;A", " Szűrők ; Elektronika ", "0"],
    )
    console.run_console()
    loaded = settings.load_settings()
    assert loaded == settings.Settings(
        str(tmp_path.resolve()), ["Szűrők", "Elektronika"]
    )
    assert "Próbáld újra" in capsys.readouterr().out


def test_menu_import_add_remove_change_directory(
    monkeypatch, tmp_path, settings_location, capsys
):
    settings.create_settings(str(tmp_path), [])
    csv = tmp_path / "input.CSV"
    csv.touch()
    wrong = tmp_path / "wrong.txt"
    wrong.touch()
    calls = []
    monkeypatch.setattr(console, "import_csv", lambda csv_path, request_missing_group: calls.append(csv_path))
    answers(
        monkeypatch,
        [
            "wrong",
            "1",
            "",
            str(wrong),
            str(csv),
            "2",
            "",
            "New;Other",
            "3",
            "",
            "New",
            "4",
            str(tmp_path),
            "0",
        ],
    )
    console.run_console()
    assert calls == [str(csv.resolve())]
    assert settings.load_settings() == settings.Settings(
        str(tmp_path.resolve()), ["Other"]
    )
    assert "Érvénytelen választás" in capsys.readouterr().out


def test_first_start_allows_empty_groups(monkeypatch, tmp_path, settings_location):
    answers(monkeypatch, [str(tmp_path), "", "0"])
    console.run_console()
    assert settings.load_settings().additional_new_part_groups == []


@pytest.mark.parametrize(
    "values",
    [["0", "directory", "", "0"], ["directory", "0", "", "0"]],
)
def test_first_start_zero_reprompts_instead_of_cancelling(
    monkeypatch, tmp_path, settings_location, capsys, values
):
    monkeypatch.chdir(tmp_path)
    answers(monkeypatch, [str(tmp_path) if x == "directory" else x for x in values])
    console.run_console()
    assert settings.load_settings() == settings.Settings(str(tmp_path.resolve()), [])
    output = capsys.readouterr().out
    assert "megszakítva" not in output
    assert "megszakítása: 0" not in output
    assert "A beállításokat elmentettük." in output


@pytest.mark.parametrize(
    "choice,operation",
    [
        ("1", "import_csv"),
        ("2", "add_new_part_groups"),
        ("3", "remove_new_part_groups"),
        ("4", "change_output_directory"),
    ],
)
def test_cancellation_returns_to_menu_without_calling_operation(
    monkeypatch, tmp_path, settings_location, capsys, choice, operation
):
    settings.create_settings(str(tmp_path), ["Existing"])
    before = settings_location.read_bytes()

    def unexpected(**kwargs):
        pytest.fail("Megszakított műveletet nem szabad elindítani.")

    monkeypatch.setattr(console, operation, unexpected)
    answers(monkeypatch, [choice, "0", "0"])
    console.run_console()
    assert settings_location.read_bytes() == before
    output = capsys.readouterr().out
    assert output.count("=== AxelPro riport ===") == 2
    assert "elmentettük" not in output and "befejeződött" not in output


def test_reset_removes_settings_exits_and_keeps_reports(
    monkeypatch, tmp_path, settings_location
):
    report = tmp_path / "report.xlsx"
    report.write_bytes(b"existing report")
    settings.create_settings(str(tmp_path), ["Existing"])
    answers(monkeypatch, ["6"])
    console.run_console()
    assert not settings_location.exists()
    assert report.read_bytes() == b"existing report"
