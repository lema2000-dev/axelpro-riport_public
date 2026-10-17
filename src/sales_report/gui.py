"""Tkinter interface with background imports and interactive approval dialogs."""

import warnings
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread

from pathlib import Path
from collections.abc import Callable
from concurrent.futures import Future, TimeoutError
from queue import Queue, Empty
from datetime import date
from decimal import Decimal, InvalidOperation

from .models import DailyExchangeRate, ImportResult, Movement
from .import_service import import_csv
from .excel_files import WorkbookValidationWarning
from .settings import (
    create_settings,
    settings_path,
    SettingsSaveError,
    load_settings,
    SettingsStructureError,
    SettingsDirectoryError,
    SettingsGroupsError,
    SettingsFileMissingError,
    SettingsFileAccessError,
    SettingsResetError,
    add_new_part_groups,
    remove_new_part_groups,
    change_output_directory,
    reset,
)
from .csv_reader import CsvFileAccessError, CsvFormatError

class _ImportCancelled(Exception):
    """Import cancelled."""

class _Application:
    """Application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Sales Report")
        self.root.geometry("420x750")
        self.root.minsize(360,550)

        self.main_frame = ttk.Frame(
            self.root,
            padding=20
        )
        self.main_frame.pack(
            fill="both",
            expand=True,
        )

        self.root.protocol("WM_DELETE_WINDOW", self._exit)

        self._import_running = False
        self._import_thread: Thread | None = None

        self._closing = False
        self._ui_requests = Queue()

        self.root.after(100, self._process_ui_requests)
        self.root.after(0, self._initialize)

    def _build_menu(self) -> None:
        """Build menu."""
        ttk.Label(
            self.main_frame,
            text="Sales Report",
            font=("", 16, "bold"),
        ).pack(pady=(0, 15))

        menu_frame = ttk.Frame(self.main_frame)
        menu_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(menu_frame, highlightthickness=0)

        scrollbar = ttk.Scrollbar(
            menu_frame,
            orient="vertical",
            command=canvas.yview,
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        buttons_frame = ttk.Frame(canvas)

        window_id = canvas.create_window((0, 0), window=buttons_frame, anchor="nw")

        buttons_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(
                scrollregion=canvas.bbox("all"),
            )
        )

        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfig(window_id, width=event.width)
        )

        # Focus the canvas on click.
        canvas.bind("<Button-1>", lambda event: canvas.focus_set())
        # Focus the canvas when clicking between buttons.
        buttons_frame.bind("<Button-1>", lambda event: canvas.focus_set())
        # Focus the canvas when clicking outside it.
        self.root.bind("<Button-1>", lambda event: canvas.focus_set())
        self.main_frame.bind("<Button-1>", lambda event: canvas.focus_set())
        menu_frame.bind("<Button-1>", lambda event: canvas.focus_set())
        # Keyboard scrolling while the canvas has focus.
        canvas.bind("<Up>", lambda event: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Down>", lambda event: canvas.yview_scroll(1, "units"))

        buttons = [
            ("Import", self._start_import),
            ("Add new-part\ncategories", self._add_groups),
            ("Remove new-part\ncategories", self._remove_groups),
            ("Change report\nfolder", self._change_directory),
            ("View current\nsettings", self._show_settings),
            ("Exit", self._exit),
            ("Reset and\nclose", self._reset_and_exit),
        ]

        for index, (text, command) in enumerate(buttons):
            button_frame = ttk.Frame(buttons_frame, width=240, height=80)
            button_frame.pack(pady=5)
            button_frame.pack_propagate(False)

            button = ttk.Button(
                button_frame,
                text=text,
                command=command,
            )
            button.pack(fill="both", expand=True)
    
            button.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))
            button.bind("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))
            button.bind("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))
            
            button_frame.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))
            button_frame.bind("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))
            button_frame.bind("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))

            if index == 0:
                self.import_button = button

        canvas.focus_set()

    def _request_directory(
            self,
            *,
            first_open: bool = False,
        ) -> str | None:
            """Request directory."""
            selected_directory: str | None = None
    
            window = tk.Toplevel(self.root)
            window.title("Select report folder")
            window.resizable(False, False)
            window.transient(self.root)
    
            frame = ttk.Frame(window, padding=20)
            frame.pack(fill="both", expand=True)
    
            ttk.Label(
                frame,
                text="Select the folder for report files.",
            ).pack(pady=(0, 15))
    
            def browse() -> None:
                nonlocal selected_directory
    
                try:
                    directory = filedialog.askdirectory(
                        parent=window,
                        title="Select report folder",
                        mustexist=True,
                    )
    
                    if not directory:
                        return
    
                    path = Path(directory)
    
                    if not path.is_dir():
                        raise ValueError(
                            "The selected folder is no longer available."
                        )
    
                    selected_directory = str(path.resolve())
    
                except (OSError, RuntimeError, ValueError) as exc:
                    messagebox.showerror(
                        title="Folder selection error",
                        message=str(exc),
                        parent=window,
                    )
                    return
    
                window.destroy()
    
            ttk.Button(
                frame,
                text="Browse",
                command=browse,
            ).pack(fill="x")
    
            if first_open:
                window.protocol("WM_DELETE_WINDOW", lambda: None)
            else:
                ttk.Button(
                    frame,
                    text="Cancel",
                    command=window.destroy,
                ).pack(fill="x", pady=(10, 0))
    
            window.wait_visibility()
            window.grab_set()
            self.root.wait_window(window)
    
            return selected_directory
    
    def _request_groups(
            self,
            *,
            title: str,
            button_text: str,
            first_open: bool = False,
        ) -> list[str] | None:
            """Request groups."""
            selected_groups: list[str] | None = None
    
            window = tk.Toplevel(self.root)
            window.title(title)
            window.resizable(False, False)
            window.transient(self.root)
    
            frame = ttk.Frame(window, padding=20)
            frame.pack(fill="both", expand=True)
    
            ttk.Label(
                frame,
                text="Category names separated by semicolons:",
            ).pack(anchor="w", pady=(0, 5))
    
            entry = ttk.Entry(frame, width=45)
            entry.pack(fill="x", pady=(0, 15))
    
            def submit() -> None:
                nonlocal selected_groups
    
                value = entry.get().strip()
    
                if first_open and not value:
                    selected_groups = []
                    window.destroy()
                    return
    
                groups = [group.strip() for group in value.split(";")]
    
                if not value:
                    error = "Enter at least one category name."
                elif any(not group for group in groups):
                    error = "Category names between separators must not be empty."
                elif len(groups) != len({group.casefold() for group in groups}):
                    error = "Enter each category name only once."
                else:
                    selected_groups = groups
                    window.destroy()
                    return
    
                messagebox.showerror(
                    title="Invalid category list",
                    message=error,
                    parent=window,
                )
    
                entry.delete(0, tk.END)
                entry.focus_set()
    
            ttk.Button(
                frame,
                text=button_text,
                command=submit,
            ).pack(fill="x")
    
            if first_open:
                window.protocol("WM_DELETE_WINDOW", lambda: None)
            else:
                ttk.Button(
                    frame,
                    text="Cancel",
                    command=window.destroy,
                ).pack(fill="x", pady=(10, 0))
    
            window.wait_visibility()
            window.grab_set()
            entry.focus_set()
            self.root.wait_window(window)
    
            return selected_groups

    def _change_groups(
        self,
        *,
        title: str,
        button_text: str,
        save_groups: Callable[[list[str]], None],
    ) -> None:
        """Change groups."""
        while True:
            groups = self._request_groups(
                title=title,
                button_text=button_text,
            )

            if groups is None:
                return

            try:
                save_groups(groups)
            except (
                SettingsStructureError,
                SettingsDirectoryError,
                SettingsGroupsError,
                SettingsFileMissingError,
                SettingsFileAccessError,
                SettingsSaveError,
            ) as exc:
                messagebox.showerror(
                    title="Settings error",
                    message=str(exc),
                    parent=self.root,
                )
                continue

            return

    def _ensure_settings(self) -> None:
        """Ensure settings."""
        while not self._closing:
            try:
                settings_exists = settings_path().exists()
            except OSError as exc:
                messagebox.showerror(
                    title="Settings error",
                    message="Cannot check settings file availability.\n"
                    f"Details: {exc}",
                    parent=self.root,
                )
                continue

            break

        if self._closing:
            return

        if settings_exists:
            return

        directory = self._request_directory(first_open=True)

        if self._closing:
            return

        groups = self._request_groups(
            title="New-part categories",
            button_text="Done",
            first_open=True,
        )

        if self._closing:
            return

        while not self._closing:
            try:
                create_settings(
                    output_directory=directory,
                    additional_new_part_groups=groups,
                )
            except SettingsSaveError as exc:
                messagebox.showerror(
                    title="Save error",
                    message=str(exc),
                    parent=self.root,
                )
                continue

            return

    def _initialize(self) -> None:
        """Initialize."""
        self._ensure_settings()

        if not self._closing:
            self._build_menu()

    def _add_groups(self) -> None:
        """Add groups."""
        self._change_groups(
            title="Add new-part categories",
            button_text="Add",
            save_groups=add_new_part_groups,
        )

    def _remove_groups(self) -> None:
        """Remove groups."""
        self._change_groups(
            title="Remove new-part categories",
            button_text="Remove",
            save_groups=remove_new_part_groups,
        )

    def _change_directory(self) -> None:
        """Change directory."""
        while True:
            directory = self._request_directory()

            if directory is None:
                return

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
                messagebox.showerror(
                    title="Settings error",
                    message=str(exc),
                    parent=self.root,
                )
                continue

            return

    def _show_settings(self) -> None:
            """Show settings."""
            try:
                settings = load_settings()
            except (
                SettingsStructureError,
                SettingsDirectoryError,
                SettingsGroupsError,
                SettingsFileMissingError,
                SettingsFileAccessError,
            ) as exc:
                messagebox.showerror(
                    title="Settings error",
                    message=str(exc),
                    parent=self.root,
                )
                return
    
            groups_text = "\n".join(
                f"• {group}"
                for group in settings.additional_new_part_groups
            ) 
    
            if not groups_text:
                groups_text = "No new-part categories have been configured."
    
            messagebox.showinfo(
                title="Current settings",
                message=f"Report folder:\n{settings.output_directory}\n\n"
                f"Additional new-part categories:\n{groups_text}"
            , parent=self.root)
    
    def _exit(self) -> None:
            """Exit."""
            if self._closing:
                return
    
            self._closing = True
    
            for window in self.root.winfo_children():
                if isinstance(window, tk.Toplevel):
                    window.destroy()
    
            self.root.withdraw()
            self._wait_for_import_exit()

    def _wait_for_import_exit(self) -> None:
        """Wait for import exit."""
        if self._import_thread is not None and self._import_thread.is_alive():
            self.root.after(100, self._wait_for_import_exit)
            return

        self.root.destroy()
        
    def _reset_and_exit(self) -> None:
        """Reset and exit."""
        confirmed = messagebox.askyesno(
            title="Reset and close",
            message=(
                "Warning! This deletes the saved report folder path "
                "and the configured part categories.\n"
                "Existing Excel reports will be kept.\n\n"
                "Are you sure you want to reset the application?"
            ),
            parent=self.root,
        )

        if not confirmed:
            return

        try:
            reset()
        except SettingsResetError as exc:
            messagebox.showerror(
                title="Reset failed",
                message=str(exc),
                parent=self.root,
            )
            return

        self._exit()

    def _request_csv_path(self) -> Path | None:
        """Request csv path."""
        while True:
            try:
                filename = filedialog.askopenfilename(
                    parent=self.root,
                    title="Select CSV to import",
                    filetypes=[
                        ("CSV files", "*.csv"),
                        ("All files", "*"),
                    ],
                )

                if not filename:
                    return None

                path = Path(filename)

                if path.suffix.lower() != ".csv":
                    raise ValueError("Select a CSV file.")

                if not path.is_file():
                    raise ValueError("The selected file is no longer available.")

                return path.resolve()
            except (OSError, RuntimeError, ValueError) as exc:
                messagebox.showerror(
                    title="CSV selection error",
                    message=str(exc),
                    parent=self.root,
                )

    def _start_import(self) -> None:
        """Start import."""
        if self._closing or self._import_running:
            return

        csv_path = self._request_csv_path()

        if csv_path is None or self._closing:
            return

        self._import_running = True
        self.import_button.configure(state="disabled")

        self._import_thread = Thread(
            target=self._import_worker,
            args=(csv_path,),
            daemon=False,
        )

        try:
            self._import_thread.start()
        except RuntimeError as exc:
            self._import_running = False
            self.import_button.configure(state="normal")
            messagebox.showerror(
                title="Import error",
                message=f"Could not start the import.\nDetails: {exc}",
                parent=self.root,
            )

    def _request_rate(
        self,
        application_date: date,
    ) -> DailyExchangeRate:
        """Request rate."""
        selected_rate: DailyExchangeRate | None = None

        window = tk.Toplevel(self.root)
        window.title("Enter exchange rate manually")
        window.resizable(False, False)
        window.transient(self.root)

        frame = ttk.Frame(window, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=f"Application date: {application_date:%Y-%m-%d}",
        ).pack(anchor="w", pady=(0, 15))

        ttk.Label(
            frame,
            text="EUR exchange rate (HUF/EUR):",
        ).pack(anchor="w")

        rate_entry = ttk.Entry(frame, width=35)
        rate_entry.pack(fill="x", pady=(5, 15))

        ttk.Label(
            frame,
            text="Source date (YYYY-MM-DD):",
        ).pack(anchor="w")

        date_entry = ttk.Entry(frame, width=35)
        date_entry.pack(fill="x", pady=(5, 15))

        def show_input_error(message: str, entry: ttk.Entry) -> None:
            messagebox.showerror(
                title="Invalid exchange rate",
                message=message,
                parent=window,
            )
            entry.delete(0, tk.END)
            entry.focus_set()

        def submit() -> None:
            nonlocal selected_rate

            try:
                rate = Decimal(rate_entry.get().strip().replace(",", "."))

                DailyExchangeRate(
                    application_date=application_date,
                    source_date=application_date,
                    rate=rate,
                )

            except (InvalidOperation, ValueError, OverflowError):
                show_input_error(
                    "Enter a positive, finite number that Excel can store "
                    "(for example: 400 or 400.25).",
                    rate_entry,
                )
                return

            text = date_entry.get().strip()
            try:
                source_date = date.fromisoformat(text)

                if source_date.isoformat() != text:
                    raise ValueError()
            except ValueError:
                show_input_error(
                    "Enter a valid date in YYYY-MM-DD format "
                    "(for example: 2026-10-12).",
                    date_entry,
                )
                return

            selected_rate = DailyExchangeRate(
                application_date=application_date,
                source_date=source_date,
                rate=rate,
            )
            window.destroy()

        ttk.Button(
            frame,
            text="Done",
            command=submit,
        ).pack(fill="x")

        window.wait_visibility()
        window.grab_set()
        rate_entry.focus_set()
        self.root.wait_window(window)

        if selected_rate is None:
            raise _ImportCancelled()

        return selected_rate

    def _show_message(self, message: str) -> None:
        """Show message."""
        messagebox.showinfo(
            title="Import",
            message=message,
            parent=self.root,
        )

    def _confirm_preview(
        self,
        year: int,
        path: Path,
        movements: list[Movement],
        import_result: ImportResult,
    ) -> bool:
        """Confirm preview."""
        accepted = False

        window = tk.Toplevel(self.root)
        window.title(f"{year} import preview")
        window.geometry("850x600")
        window.transient(self.root)

        frame = ttk.Frame(window, padding=15)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=f"Target file: {path}",
            wraplength=780,
        ).pack(anchor="w", pady=(0, 10))

        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True)

        table = ttk.Treeview(
            table_frame,
            columns=("period", "category", "amount"),
            show="headings",
        )

        table.heading("period", text="Day / month")
        table.heading("category", text="Description")
        table.heading("amount", text="Amount to add (HUF)")

        table.column("period", width=130, anchor="w")
        table.column("category", width=380, anchor="w")
        table.column("amount", width=180, anchor="e")

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=table.yview,
        )
        table.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        table.pack(side="left", fill="both", expand=True)

        daily_labels = {
            "used_parts": "Used parts",
            "new_parts": "New parts",
            "labor": "Labor",
            "workshop_used_parts": "Workshop used parts",
            "workshop_new_parts": "Workshop new parts",
            "purchases": "Purchase",
            "shipping": "Shipping",
            "trailer": "Trailer",
        }

        monthly_labels = {
            "consumables": "Consumables",
            "diagnostics": "Diagnostics",
            "air_conditioning": "Air conditioning repair",
        }

        for totals in sorted(
            import_result.daily_totals,
            key=lambda totals: totals.movement_date,
        ):
            if totals.movement_date.year != year:
                continue

            for field, label in daily_labels.items():
                amount = getattr(totals, field)

                table.insert(
                    "",
                    "end",
                    values=(
                        totals.movement_date.isoformat(),
                        label,
                        f"{amount:,}".replace(",", " "),
                    ),
                )

        for totals in sorted(
            import_result.monthly_totals,
            key=lambda totals: (totals.year, totals.month),
        ):
            if totals.year != year:
                continue

            for field, label in monthly_labels.items():
                amount = getattr(totals, field)

                table.insert(
                    "",
                    "end",
                    values=(
                        f"{year}-{totals.month:02d}",
                        label,
                        f"{amount:,}".replace(",", " "),
                    ),
                )

        new_documents = {
            document_number
            for document_year, document_number in import_result.new_documents
            if document_year == year
        }

        skipped_movements = [
            movement
            for movement in movements
            if movement.movement_date.year == year
            and movement.document_key not in import_result.new_documents
        ]

        ttk.Label(
            frame,
            text=(
                f"New documents: {len(new_documents)}    "
                f"Skipped items: {len(skipped_movements)}"
            ),
        ).pack(anchor="w", pady=(10, 0))

        for movement in skipped_movements:
            table.insert(
                "",
                "end",
                values=(
                    movement.movement_date.isoformat(),
                    (
                        f"Skipped: {movement.document_number} / "
                        f"{movement.item_code}"
                    ),
                    f"{movement.net_value} {movement.currency}",
                ),
            )

        def accept() -> None:
            nonlocal accepted
            accepted = True
            window.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(15, 0))

        ttk.Button(
            buttons,
            text="Reject",
            command=window.destroy,
        ).pack(side="left")

        ttk.Button(
            buttons,
            text="Accept",
            command=accept,
        ).pack(side="right")

        window.wait_visibility()
        window.grab_set()
        self.root.wait_window(window)

        if self._closing:
            raise _ImportCancelled()

        return accepted

    def _check_cancelled(self) -> None:
        """Check cancelled."""
        if self._closing:
            raise _ImportCancelled()

    def _request_missing_group(
        self,
        document_number: str,
        item_code: str,
    ) -> str:
        """Ask the user to classify an invoice item with an empty group."""
        self._check_cancelled()
        selected_group: str | None = None

        window = tk.Toplevel(self.root)
        window.title("Missing item group")
        window.transient(self.root)
        window.resizable(False, False)

        frame = ttk.Frame(window, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=(
                "This invoice item has no group.\n\n"
                f"Document number: {document_number}\n"
                f"Item code: {item_code}\n\n"
                "Select the item group:"
            ),
            wraplength=400,
            justify="left",
        ).pack(fill="x", pady=(0, 15))

        def select_group(group: str) -> None:
            nonlocal selected_group
            selected_group = group
            window.destroy()

        choices = (
            ("New parts", "New parts"),
            ("Used parts", "Used parts"),
            ("Labor", "Labor"),
            ("Shipping", "Shipping"),
            ("Trailer", "Trailer"),
            ("None of these", ""),
        )

        for label, group in choices:
            ttk.Button(
                frame,
                text=label,
                command=lambda value=group: select_group(value),
            ).pack(fill="x", pady=(0, 5))

        window.wait_visibility()
        window.grab_set()
        self.root.wait_window(window)

        self._check_cancelled()

        if selected_group is None:
            raise _ImportCancelled()
        
        return selected_group

    def _import_worker(self, csv_path: Path) -> None:
        """Import worker."""
        error: Exception | None = None
        cancelled = False
        warning_messages: list[str] = []
        collected_warnings = []

        try:
            with warnings.catch_warnings(record=True) as collected_warnings:
                warnings.simplefilter("always", WorkbookValidationWarning)

                import_csv(
                    csv_path=csv_path,
                    request_rate=lambda day: self._call_ui(self._request_rate, day),
                    show_message=lambda message: self._call_ui(self._show_message, message),
                    confirm_preview=lambda year, path, movements, import_result: (
                        self._call_ui(
                            self._confirm_preview,
                            year,
                            path,
                            movements,
                            import_result,
                        )
                    ),
                    check_cancelled=self._check_cancelled,
                    request_missing_group=lambda document_number, item_code: (
                        self._call_ui(
                            self._request_missing_group,
                            document_number,
                            item_code,
                        )
                    ),
                )

        except _ImportCancelled:
            cancelled = True
        except Exception as exc:
            error = exc

        finally:
            first_warning = next(
                (
                    warning for warning in collected_warnings
                    if isinstance(warning.message, WorkbookValidationWarning)
                ),
                None
            )

            warning_messages = (
                [str(first_warning.message)] if first_warning is not None else []
            )

            if not self._closing:
                self._ui_requests.put(
                    (
                        self._finish_import,
                        (error, cancelled, warning_messages),
                        Future(),
                    )
                )

    def _finish_import(
        self,
        error: Exception | None,
        cancelled: bool,
        warning_messages: list[str],
    ) -> None:
        """Finish import."""
        self._import_running = False

        if self._closing:
            return

        self.import_button.configure(state="normal")

        if error is not None:
            messagebox.showerror(
                title="Import error",
                message=str(error),
                parent=self.root,
            )

            if isinstance(error, (CsvFileAccessError, CsvFormatError)):
                self.root.after(0, self._start_import)

            return

        if cancelled:
            messagebox.showinfo(
                title="Import cancelled",
                message=(
                    "You cancelled the import.\n"
                    "Years already saved during this import are retained."
                ),
                parent=self.root,
            )
            return

        messagebox.showinfo(
            title="Import",
            message="Import process completed.",
            parent=self.root,
        )

        if warning_messages:
            messagebox.showwarning(
                title="Workbook warnings",
                message="\n\n".join(warning_messages),
                parent=self.root,
            )

    def _call_ui(self, function: Callable, *args):
        """Call ui."""
        if self._closing:
            raise _ImportCancelled()

        result = Future()
        self._ui_requests.put((function, args, result))

        while True:
            if self._closing:
                raise _ImportCancelled()

            try:
                return result.result(timeout=0.1)
            except TimeoutError:
                if result.done():
                    raise
                continue

    def _process_ui_requests(self) -> None:
        """Process ui requests."""
        if self._closing:
            return

        try:
            function, args, result = self._ui_requests.get_nowait()
        except Empty:
            pass
        else:
            try:
                value = function(*args)
            except Exception as exc:
                result.set_exception(exc)
            else:
                result.set_result(value)

        if not self._closing:
            self.root.after(100, self._process_ui_requests)

def run_gui() -> None:
    """Create the application root and run the Tkinter event loop."""
    root = tk.Tk()
    app = _Application(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()