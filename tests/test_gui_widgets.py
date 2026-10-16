"""Valódi Tkinter-elemek ellenőrzése; grafikus kijelző szükséges."""
from datetime import date
from pathlib import Path
from queue import Queue
from unittest.mock import Mock
import tkinter as tk
from tkinter import ttk
import pytest
from axelpro_riport import gui
from axelpro_riport.models import ImportResult


@pytest.fixture
def widget_app(monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        pytest.skip(f'Grafikus kijelző nem érhető el: {exc}')
    root.geometry('420x700')
    app = gui._Application.__new__(gui._Application)
    app.root = root
    app.main_frame = ttk.Frame(root, padding=20)
    app.main_frame.pack(fill='both', expand=True)
    app._closing = False
    app._import_running = False
    app._import_thread = None
    app._ui_requests = Queue()
    for name in ('showerror', 'showinfo', 'showwarning'):
        monkeypatch.setattr(gui.messagebox, name, Mock())
    # A párbeszédablakok valódi elemek, a felhasználói műveleteket vezéreljük.
    monkeypatch.setattr(tk.Toplevel, 'wait_visibility', lambda self: self.update_idletasks())
    monkeypatch.setattr(tk.Toplevel, 'grab_set', lambda self: None)
    try:
        yield app
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass


def descendants(widget):
    result = []
    for child in widget.winfo_children():
        result.append(child)
        result.extend(descendants(child))
    return result


def button(window, label):
    return next(w for w in descendants(window) if isinstance(w, ttk.Button) and w.cget('text') == label)


def test_menu_has_seven_sibling_buttons_with_expected_size(widget_app):
    app = widget_app
    app._build_menu()
    app.root.update_idletasks()
    buttons = [w for w in descendants(app.main_frame) if isinstance(w, ttk.Button)]
    assert len(buttons) == 7
    assert [w.cget('text').replace('\n', ' ') for w in buttons] == [
        'Importálás', 'Újalkatrész- csoport hozzáadása', 'Újalkatrész- csoport eltávolítása',
        'Riportmappa módosítása', 'Jelenlegi beállítások megtekintése', 'Kilépés',
        'Alaphelyzetbe állítás és bezárás']
    assert len({w.master.master for w in buttons}) == 1
    for w in buttons:
        assert w.winfo_width() == 240
        assert w.winfo_height() == 80
    assert app.import_button is buttons[0]


@pytest.mark.parametrize('text,first_open,expected', [('Szűrők; Olaj', False, ['Szűrők', 'Olaj']), ('', True, [])])
def test_group_dialog_waits_for_submission(widget_app, monkeypatch, text, first_open, expected):
    def drive(window):
        entries = [w for w in descendants(window) if isinstance(w, ttk.Entry)]
        entries[0].insert(0, text)
        if first_open:
            assert all(w.cget('text') != 'Megszakítás' for w in descendants(window) if isinstance(w, ttk.Button))
        button(window, 'Kész').invoke()
        assert not window.winfo_exists()
    monkeypatch.setattr(widget_app.root, 'wait_window', drive)
    assert widget_app._request_groups(title='Groups', button_text='Kész', first_open=first_open) == expected


@pytest.mark.parametrize('invalid', ['', 'A;;B', 'Szűrők;szűrők'])
def test_group_error_clears_entry_and_allows_retry(widget_app, monkeypatch, invalid):
    def drive(window):
        entry = next(w for w in descendants(window) if isinstance(w, ttk.Entry))
        entry.insert(0, invalid)
        button(window, 'Hozzáadás').invoke()
        gui.messagebox.showerror.assert_called_once()
        assert window.winfo_exists() and entry.get() == ''
        entry.insert(0, 'Szűrők')
        button(window, 'Hozzáadás').invoke()
    monkeypatch.setattr(widget_app.root, 'wait_window', drive)
    assert widget_app._request_groups(title='Groups', button_text='Hozzáadás') == ['Szűrők']


def test_group_cancel_returns_none(widget_app, monkeypatch):
    monkeypatch.setattr(widget_app.root, 'wait_window', lambda window: button(window, 'Megszakítás').invoke())
    assert widget_app._request_groups(title='Groups', button_text='Hozzáadás') is None


def test_rate_dialog_validates_number_and_date_independently(widget_app, monkeypatch):
    def drive(window):
        rate, source = [w for w in descendants(window) if isinstance(w, ttk.Entry)]
        rate.insert(0, 'abc')
        source.insert(0, '2026-02-30')
        button(window, 'Kész').invoke()
        assert rate.get() == '' and source.get() == '2026-02-30'
        rate.insert(0, '400,25')
        button(window, 'Kész').invoke()
        assert rate.get() == '400,25' and source.get() == ''
        source.insert(0, '2026-10-13')
        button(window, 'Kész').invoke()
    monkeypatch.setattr(widget_app.root, 'wait_window', drive)
    record = widget_app._request_rate(date(2026, 9, 18))
    assert str(record.rate) == '400.25'
    assert record.source_date == date(2026, 10, 13)
    assert gui.messagebox.showerror.call_count == 2


def test_rate_close_cancels_import(widget_app, monkeypatch):
    monkeypatch.setattr(widget_app.root, 'wait_window', lambda window: window.destroy())
    with pytest.raises(gui._ImportCancelled):
        widget_app._request_rate(date(2026, 9, 18))


@pytest.mark.parametrize('decision,expected', [('Elfogadás', True), ('Elutasítás', False), ('close', False)])
def test_preview_buttons_and_close(widget_app, monkeypatch, decision, expected):
    def drive(window):
        assert len([w for w in descendants(window) if isinstance(w, ttk.Treeview)]) == 1
        if decision == 'close':
            window.destroy()
        else:
            button(window, decision).invoke()
    monkeypatch.setattr(widget_app.root, 'wait_window', drive)
    assert widget_app._confirm_preview(2026, Path('report.xlsx'), [], ImportResult([], [], set())) is expected


def test_directory_selection_returns_selected_path(widget_app, monkeypatch, tmp_path):
    monkeypatch.setattr(gui.filedialog, 'askdirectory', lambda **kwargs: str(tmp_path))
    monkeypatch.setattr(widget_app.root, 'wait_window', lambda window: button(window, 'Tallózás').invoke())
    assert widget_app._request_directory() == str(tmp_path.resolve())


@pytest.mark.parametrize("label,expected", [
    ("Új alkatrész", "Új alkatrész"),
    ("Bontott alkatrész", "Bontott alkatrész"),
    ("Munkadíj", "Munkadíj"),
    ("Szállítási költség", "Szállítási költség"),
    ("Trailer", "Trailer"),
    ("Egyik sem", ""),
])
def test_missing_group_dialog_has_six_choices_and_returns_selected_group(widget_app, monkeypatch, label, expected):
    def drive(window):
        widgets = descendants(window)
        buttons = [w for w in widgets if isinstance(w, ttk.Button)]
        assert [w.cget("text") for w in buttons] == [
            "Új alkatrész", "Bontott alkatrész", "Munkadíj",
            "Szállítási költség", "Trailer", "Egyik sem",
        ]
        text = "\n".join(w.cget("text") for w in widgets if isinstance(w, ttk.Label))
        assert "DOC-123" in text and "ITEM-456" in text
        button(window, label).invoke()
        assert not window.winfo_exists()
    monkeypatch.setattr(widget_app.root, "wait_window", drive)
    assert widget_app._request_missing_group("DOC-123", "ITEM-456") == expected


def test_missing_group_window_close_cancels_instead_of_returning_function(widget_app, monkeypatch):
    # A destroy az ablakkezelő alapértelmezett X-es bezárását modellezi.
    monkeypatch.setattr(widget_app.root, "wait_window", lambda window: window.destroy())
    with pytest.raises(gui._ImportCancelled):
        widget_app._request_missing_group("DOC", "ITEM")


def test_missing_group_dialog_does_not_open_when_application_is_closing(widget_app, monkeypatch):
    widget_app._closing = True
    create_window = Mock()
    monkeypatch.setattr(gui.tk, "Toplevel", create_window)
    with pytest.raises(gui._ImportCancelled):
        widget_app._request_missing_group("DOC", "ITEM")
    create_window.assert_not_called()


def test_application_close_during_group_dialog_cancels(widget_app, monkeypatch):
    def drive(window):
        widget_app._closing = True
        window.destroy()
    monkeypatch.setattr(widget_app.root, "wait_window", drive)
    with pytest.raises(gui._ImportCancelled):
        widget_app._request_missing_group("DOC", "ITEM")
