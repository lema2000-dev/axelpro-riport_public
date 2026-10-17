"""Test gui."""
from concurrent.futures import Future
from datetime import date
from decimal import Decimal
from pathlib import Path
from queue import Queue
from threading import Thread, get_ident
from types import SimpleNamespace
from unittest.mock import Mock
import pytest
from sales_report import gui
from sales_report.models import DailyExchangeRate, ImportResult


@pytest.fixture
def app(monkeypatch):
    application = gui._Application.__new__(gui._Application)
    application.root = Mock()
    application.root.winfo_children.return_value = []
    application._closing = False
    application._import_running = False
    application._import_thread = None
    application._ui_requests = Queue()
    application.import_button = Mock()
    for name in ('showerror', 'showinfo', 'showwarning', 'askyesno'):
        monkeypatch.setattr(gui.messagebox, name, Mock())
    return application


def test_menu_initialization_waits_for_settings(app):
    order = []
    app._ensure_settings = lambda: order.append('settings')
    app._build_menu = lambda: order.append('menu')
    app._initialize()
    assert order == ['settings', 'menu']


def test_initialization_does_not_build_menu_after_close(app):
    app._ensure_settings = lambda: setattr(app, '_closing', True)
    app._build_menu = Mock()
    app._initialize()
    app._build_menu.assert_not_called()


@pytest.mark.parametrize('cancelled', [False, True])
def test_start_import_cancel_or_start(app, monkeypatch, tmp_path, cancelled):
    path = tmp_path / 'test.csv'
    app._request_csv_path = Mock(return_value=None if cancelled else path)
    thread = Mock()
    constructor = Mock(return_value=thread)
    monkeypatch.setattr(gui, 'Thread', constructor)
    app._start_import()
    assert app._import_running is (not cancelled)
    if cancelled:
        constructor.assert_not_called()
        app.import_button.configure.assert_not_called()
    else:
        assert constructor.call_args.kwargs['args'] == (path,)
        assert constructor.call_args.kwargs['daemon'] is False
        assert constructor.call_args.kwargs['target'] == app._import_worker
        thread.start.assert_called_once()
        app.import_button.configure.assert_called_with(state='disabled')


@pytest.mark.parametrize('state', ['closing', 'running'])
def test_repeated_import_is_not_started(app, state):
    setattr(app, '_closing' if state == 'closing' else '_import_running', True)
    app._request_csv_path = Mock()
    app._start_import()
    app._request_csv_path.assert_not_called()


def test_thread_start_failure_restores_button(app, monkeypatch):
    app._request_csv_path = Mock(return_value=Path('test.csv'))
    thread = Mock()
    thread.start.side_effect = RuntimeError('thread failed')
    monkeypatch.setattr(gui, 'Thread', Mock(return_value=thread))
    app._start_import()
    assert not app._import_running
    app.import_button.configure.assert_called_with(state='normal')
    assert 'thread failed' in gui.messagebox.showerror.call_args.kwargs['message']


@pytest.mark.parametrize('error,retry', [(gui.CsvFormatError('bad csv'), True), (gui.CsvFileAccessError('denied'), True), (ValueError('workbook'), False)])
def test_finish_error_only_reopens_csv_for_csv_errors(app, error, retry):
    app._import_running = True
    app._finish_import(error, False, [])
    assert not app._import_running
    gui.messagebox.showerror.assert_called_once()
    gui.messagebox.showinfo.assert_not_called()
    gui.messagebox.showwarning.assert_not_called()
    if retry:
        app.root.after.assert_called_once_with(0, app._start_import)
    else:
        app.root.after.assert_not_called()


def test_finish_success_and_warnings(app):
    app._finish_import(None, False, ['format'])
    gui.messagebox.showinfo.assert_called_once()
    assert gui.messagebox.showwarning.call_args.kwargs['message'] == 'format'


def test_finish_cancel_does_not_report_success(app):
    app._finish_import(None, True, [])
    assert gui.messagebox.showinfo.call_args.kwargs['title'] == 'Import cancelled'
    app.root.after.assert_not_called()


def test_finish_after_close_does_not_touch_widgets(app):
    app._closing = True
    app._finish_import(None, False, ['warning'])
    app.import_button.configure.assert_not_called()
    gui.messagebox.showinfo.assert_not_called()


def test_ui_queue_preserves_result_and_runs_on_main_thread(app):
    main_thread = get_ident()
    result = Future()
    app._ui_requests.put((lambda x: (get_ident(), x + 1), (4,), result))
    app._process_ui_requests()
    assert result.result() == (main_thread, 5)
    app.root.after.assert_called_once_with(100, app._process_ui_requests)


def test_ui_queue_returns_exception(app):
    failure = ValueError('dialog failed')
    def fail():
        raise failure
    result = Future()
    app._ui_requests.put((fail, (), result))
    app._process_ui_requests()
    assert result.exception() is failure


def test_empty_queue_reschedules_without_blocking(app):
    app._process_ui_requests()
    app.root.after.assert_called_once_with(100, app._process_ui_requests)


def test_call_ui_waits_for_main_thread_result(app):
    output = []
    thread = Thread(target=lambda: output.append(app._call_ui(lambda x: x * 2, 3)))
    thread.start()
    try:
        function, args, result = app._ui_requests.get(timeout=2)
        result.set_result(function(*args))
        thread.join(timeout=2)
        assert not thread.is_alive()
        assert output == [6]
    finally:
        app._closing = True
        thread.join(timeout=2)


def test_call_ui_propagates_function_timeout_not_wait_timeout(app):
    errors = []
    failure = TimeoutError('dialog timeout')
    def worker():
        try:
            app._call_ui(lambda: None)
        except Exception as exc:
            errors.append(exc)
    thread = Thread(target=worker)
    thread.start()
    try:
        _, _, result = app._ui_requests.get(timeout=2)
        result.set_exception(failure)
        thread.join(timeout=2)
        assert not thread.is_alive()
        assert errors == [failure]
    finally:
        app._closing = True
        thread.join(timeout=2)


def test_cancelled_call_never_enqueues(app):
    app._closing = True
    with pytest.raises(gui._ImportCancelled):
        app._call_ui(lambda: None)
    assert app._ui_requests.empty()


@pytest.mark.parametrize('stage', ['success', 'error', 'cancel'])
def test_worker_enqueues_completion(app, monkeypatch, stage):
    failure = ValueError('bad workbook')
    def run(**kwargs):
        if stage == 'error':
            raise failure
        if stage == 'cancel':
            raise gui._ImportCancelled()
    monkeypatch.setattr(gui, 'import_csv', run)
    app._import_worker(Path('test.csv'))
    function, args, future = app._ui_requests.get_nowait()
    assert function == app._finish_import
    assert args == (failure if stage == 'error' else None, stage == 'cancel', [])
    assert isinstance(future, Future)
    assert app._ui_requests.empty()


def test_worker_callback_names_match_keyword_import_contract(app, monkeypatch):
    day = date(2026, 9, 18)
    record = DailyExchangeRate(day, day, Decimal(400))
    app._request_rate = Mock(return_value=record)
    app._show_message = Mock()
    app._confirm_preview = Mock(return_value=True)
    app._request_missing_group = Mock(return_value="Used parts")
    app._call_ui = lambda function, *args: function(*args)
    result = ImportResult([], [], set())
    def run(**kwargs):
        assert kwargs['request_rate'](day) == record
        assert kwargs['request_missing_group']('DOC', 'ITEM') == 'Used parts'
        kwargs['show_message']('message')
        assert kwargs['confirm_preview'](year=2026, path=Path('report.xlsx'), movements=[], import_result=result)
        kwargs['check_cancelled']()
    monkeypatch.setattr(gui, 'import_csv', run)
    app._import_worker(Path('test.csv'))
    _, (error, cancelled, _), _ = app._ui_requests.get_nowait()
    assert error is None and not cancelled
    app._show_message.assert_called_once_with('message')
    app._request_missing_group.assert_called_once_with('DOC', 'ITEM')
    app._confirm_preview.assert_called_once_with(2026, Path('report.xlsx'), [], result)


def test_worker_keeps_only_first_workbook_warning(app, monkeypatch):
    import warnings
    def run(**kwargs):
        warnings.warn('unrelated before', UserWarning)
        warnings.warn('same', gui.WorkbookValidationWarning)
        warnings.warn('same', gui.WorkbookValidationWarning)
        warnings.warn('second workbook warning', gui.WorkbookValidationWarning)
        warnings.warn('other', UserWarning)
    monkeypatch.setattr(gui, 'import_csv', run)
    app._import_worker(Path('test.csv'))
    _, args, _ = app._ui_requests.get_nowait()
    assert args == (None, False, ['same'])


def test_worker_does_not_enqueue_after_close(app, monkeypatch):
    monkeypatch.setattr(gui, 'import_csv', lambda **kwargs: setattr(app, '_closing', True))
    app._import_worker(Path('test.csv'))
    assert app._ui_requests.empty()


def test_exit_waits_for_active_thread_without_destroying_root(app):
    app._import_thread = Mock()
    app._import_thread.is_alive.return_value = True
    app._exit()
    assert app._closing
    app.root.withdraw.assert_called_once()
    app.root.destroy.assert_not_called()
    app.root.after.assert_called_once_with(100, app._wait_for_import_exit)
    app._import_thread.is_alive.return_value = False
    app._wait_for_import_exit()
    app.root.destroy.assert_called_once()


@pytest.mark.parametrize('confirmed', [True, False])
def test_reset_requires_confirmation(app, monkeypatch, confirmed):
    gui.messagebox.askyesno.return_value = confirmed
    reset = Mock()
    monkeypatch.setattr(gui, 'reset', reset)
    app._exit = Mock()
    app._reset_and_exit()
    assert reset.call_count == int(confirmed)
    assert app._exit.call_count == int(confirmed)


def test_reset_failure_keeps_application_open(app, monkeypatch):
    gui.messagebox.askyesno.return_value = True
    monkeypatch.setattr(gui, 'reset', Mock(side_effect=gui.SettingsResetError('denied')))
    app._exit = Mock()
    app._reset_and_exit()
    gui.messagebox.showerror.assert_called_once()
    app._exit.assert_not_called()


def test_csv_picker_retries_bad_extension_then_returns_real_path(app, monkeypatch, tmp_path):
    path = tmp_path / 'input.csv'
    path.write_text('test')
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', Mock(side_effect=['wrong.xlsx', str(path)]))
    assert app._request_csv_path() == path.resolve()
    gui.messagebox.showerror.assert_called_once()


def test_csv_picker_retries_os_error_then_can_cancel(app, monkeypatch):
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', Mock(side_effect=[PermissionError('denied'), '']))
    assert app._request_csv_path() is None
    gui.messagebox.showerror.assert_called_once()


def test_group_save_failure_can_retry_or_cancel(app):
    app._request_groups = Mock(side_effect=[['Filters'], None])
    save = Mock(side_effect=gui.SettingsSaveError('denied'))
    app._change_groups(title='Groups', button_text='Save', save_groups=save)
    save.assert_called_once_with(['Filters'])
    assert app._request_groups.call_count == 2
    gui.messagebox.showerror.assert_called_once()


def test_settings_display_includes_directory_and_groups(app, monkeypatch):
    monkeypatch.setattr(gui, 'load_settings', lambda: SimpleNamespace(output_directory='/reports', additional_new_part_groups=['Filters']))
    app._show_settings()
    text = gui.messagebox.showinfo.call_args.kwargs['message']
    assert '/reports' in text and 'Filters' in text


def test_first_start_empty_groups_are_saved(app, monkeypatch, tmp_path):
    monkeypatch.setattr(gui, 'settings_path', lambda: tmp_path / 'settings.json')
    app._request_directory = Mock(return_value=str(tmp_path))
    app._request_groups = Mock(return_value=[])
    create = Mock()
    monkeypatch.setattr(gui, 'create_settings', create)
    app._ensure_settings()
    create.assert_called_once_with(output_directory=str(tmp_path), additional_new_part_groups=[])
    app._request_directory.assert_called_once_with(first_open=True)


def test_existing_settings_skip_initial_dialogs(app, monkeypatch, tmp_path):
    path = tmp_path / 'settings.json'
    path.write_text('{}')
    monkeypatch.setattr(gui, 'settings_path', lambda: path)
    app._request_directory = Mock()
    app._request_groups = Mock()
    app._ensure_settings()
    app._request_directory.assert_not_called()
    app._request_groups.assert_not_called()


@pytest.mark.parametrize("group", ["New parts", "Used parts", "Labor", "Shipping", "Trailer", ""])
def test_worker_group_choice_goes_through_ui_bridge(app, monkeypatch, group):
    app._request_missing_group = Mock(return_value=group)
    app._call_ui = Mock(side_effect=lambda function, *args: function(*args))
    def run(**kwargs):
        assert kwargs["request_missing_group"]("DOC-1", "ITEM-1") == group
    monkeypatch.setattr(gui, "import_csv", run)
    app._import_worker(Path("test.csv"))
    app._call_ui.assert_called_once_with(app._request_missing_group, "DOC-1", "ITEM-1")
    _, args, _ = app._ui_requests.get_nowait()
    assert args == (None, False, [])


def test_group_dialog_cancellation_reaches_import_completion(app, monkeypatch):
    app._request_missing_group = Mock(side_effect=gui._ImportCancelled())
    app._call_ui = lambda function, *args: function(*args)
    def run(**kwargs):
        kwargs["request_missing_group"]("DOC", "ITEM")
    monkeypatch.setattr(gui, "import_csv", run)
    app._import_worker(Path("test.csv"))
    finish, args, _ = app._ui_requests.get_nowait()
    assert args == (None, True, [])
    finish(*args)
    gui.messagebox.showerror.assert_not_called()
    gui.messagebox.showwarning.assert_not_called()
    assert "cancelled" in gui.messagebox.showinfo.call_args.kwargs["title"].casefold()
    app.import_button.configure.assert_called_once_with(state="normal")


def test_worker_first_warning_can_be_displayed_after_import(app, monkeypatch):
    import warnings
    def run(**kwargs):
        warnings.warn("FIRST", gui.WorkbookValidationWarning)
        warnings.warn("SECOND", gui.WorkbookValidationWarning)
        gui.messagebox.showwarning.assert_not_called()
    monkeypatch.setattr(gui, "import_csv", run)
    app._import_worker(Path("test.csv"))
    finish, args, _ = app._ui_requests.get_nowait()
    assert args == (None, False, ["FIRST"])
    finish(*args)
    gui.messagebox.showwarning.assert_called_once_with(
        title="Workbook warnings", message="FIRST", parent=app.root,
    )


def test_unrelated_warning_does_not_show_workbook_warning(app, monkeypatch):
    import warnings
    def run(**kwargs):
        warnings.warn("unrelated", UserWarning)
    monkeypatch.setattr(gui, "import_csv", run)
    app._import_worker(Path("test.csv"))
    finish, args, _ = app._ui_requests.get_nowait()
    assert args == (None, False, [])
    finish(*args)
    gui.messagebox.showwarning.assert_not_called()
