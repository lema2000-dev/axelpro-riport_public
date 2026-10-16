"""Az alkalmazás grafikus felülete."""

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
    """Az alkalmazás bezárása miatt megszakított importálás."""

class _Application:
    """A főablak és a hozzá tartozó felületi állapot."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("AxelPro Riport")
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
        """A főmenü gombjainak létrehozása."""
        ttk.Label(
            self.main_frame,
            text="AxelPro Riport",
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

        # Kattintásra fókuszba hozzuk a Canvas-t
        canvas.bind("<Button-1>", lambda event: canvas.focus_set())
        # Gombok közötti üres térre kattintva is fókuszba hozzuk a Canvas-t
        buttons_frame.bind("<Button-1>", lambda event: canvas.focus_set())
        # A canvason kívülre kattintva is fókuszba hozzuk a Canvas-t
        self.root.bind("<Button-1>", lambda event: canvas.focus_set())
        self.main_frame.bind("<Button-1>", lambda event: canvas.focus_set())
        menu_frame.bind("<Button-1>", lambda event: canvas.focus_set())
        # Billentyűzetes görgetés a Canvas-ra fókuszálva
        canvas.bind("<Up>", lambda event: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Down>", lambda event: canvas.yview_scroll(1, "units"))

        buttons = [
            ("Importálás", self._start_import),
            ("Újalkatrész-\ncsoport\nhozzáadása", self._add_groups),
            ("Újalkatrész-\ncsoport\neltávolítása", self._remove_groups),
            ("Riportmappa\nmódosítása", self._change_directory),
            ("Jelenlegi\nbeállítások\nmegtekintése", self._show_settings),
            ("Kilépés", self._exit),
            ("Alaphelyzetbe\nállítás és\nbezárás", self._reset_and_exit),
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
            """Riportmappa kiválasztása tallózással."""
            selected_directory: str | None = None
    
            window = tk.Toplevel(self.root)
            window.title("Riportmappa kiválasztása")
            window.resizable(False, False)
            window.transient(self.root)
    
            frame = ttk.Frame(window, padding=20)
            frame.pack(fill="both", expand=True)
    
            ttk.Label(
                frame,
                text="Válaszd ki a riportok mentési mappáját.",
            ).pack(pady=(0, 15))
    
            def browse() -> None:
                nonlocal selected_directory
    
                try:
                    directory = filedialog.askdirectory(
                        parent=window,
                        title="Riportmappa kiválasztása",
                        mustexist=True,
                    )
    
                    if not directory:
                        return
    
                    path = Path(directory)
    
                    if not path.is_dir():
                        raise ValueError(
                            "A kiválasztott mappa már nem érhető el."
                        )
    
                    selected_directory = str(path.resolve())
    
                except (OSError, RuntimeError, ValueError) as exc:
                    messagebox.showerror(
                        title="Mappaválasztási hiba",
                        message=str(exc),
                        parent=window,
                    )
                    return
    
                window.destroy()
    
            ttk.Button(
                frame,
                text="Tallózás",
                command=browse,
            ).pack(fill="x")
    
            if first_open:
                window.protocol("WM_DELETE_WINDOW", lambda: None)
            else:
                ttk.Button(
                    frame,
                    text="Megszakítás",
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
            """Újalkatrész-csoportok megadása."""
            selected_groups: list[str] | None = None
    
            window = tk.Toplevel(self.root)
            window.title(title)
            window.resizable(False, False)
            window.transient(self.root)
    
            frame = ttk.Frame(window, padding=20)
            frame.pack(fill="both", expand=True)
    
            ttk.Label(
                frame,
                text="Csoportnevek pontosvesszővel elválasztva:",
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
                    error = "Adj meg legalább egy csoportnevet."
                elif any(not group for group in groups):
                    error = "Az elválasztók között nem lehet üres csoportnév."
                elif len(groups) != len({group.casefold() for group in groups}):
                    error = "Egy csoportnevet csak egyszer adj meg."
                else:
                    selected_groups = groups
                    window.destroy()
                    return
    
                messagebox.showerror(
                    title="Hibás csoportlista",
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
                    text="Megszakítás",
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
        """Csoportok bekérése és hozzáadása vagy eltávolítása."""
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
                    title="Beállítási hiba",
                    message=str(exc),
                    parent=self.root,
                )
                continue

            return

    def _ensure_settings(self) -> None:
        """Első indításkor a beállítások bekérése és mentése."""
        while not self._closing:
            try:
                settings_exists = settings_path().exists()
            except OSError as exc:
                messagebox.showerror(
                    title="Beállítási hiba",
                    message="A beállításfájl elérhetősége nem ellenőrizhető.\n"
                    f"Részletek: {exc}",
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
            title="Újalkatrész-csoportok",
            button_text="Kész",
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
                    title="Mentési hiba",
                    message=str(exc),
                    parent=self.root,
                )
                continue

            return

    def _initialize(self) -> None:
        """Beállítások biztosítása, majd a főmenü megnyitása."""
        self._ensure_settings()

        if not self._closing:
            self._build_menu()

    def _add_groups(self) -> None:
        """További újalkatrész-csoportok hozzáadása."""
        self._change_groups(
            title="Újalkatrész-csoportok hozzáadása",
            button_text="Hozzáadás",
            save_groups=add_new_part_groups,
        )

    def _remove_groups(self) -> None:
        """Újalkatrész-csoportok eltávolítása."""
        self._change_groups(
            title="Újalkatrész-csoportok eltávolítása",
            button_text="Eltávolítás",
            save_groups=remove_new_part_groups,
        )

    def _change_directory(self) -> None:
        """A riportmappa módosítása."""
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
                    title="Beállítási hiba",
                    message=str(exc),
                    parent=self.root,
                )
                continue

            return

    def _show_settings(self) -> None:
            """A tárolt riportmappa és csoportnevek megjelenítése."""
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
                    title="Beállítási hiba",
                    message=str(exc),
                    parent=self.root,
                )
                return
    
            groups_text = "\n".join(
                f"• {group}"
                for group in settings.additional_new_part_groups
            ) 
    
            if not groups_text:
                groups_text = "Nincsenek megadva újalkatrész-csoportok."
    
            messagebox.showinfo(
                title="Jelenlegi beállítások",
                message=f"Riportmappa:\n{settings.output_directory}\n\n"
                f"További újalkatrész-csoportok:\n{groups_text}"
            , parent=self.root)
    
    def _exit(self) -> None:
            """Megszakítás jelzése, majd biztonságos kilépés."""
            if self._closing:
                return
    
            self._closing = True
    
            for window in self.root.winfo_children():
                if isinstance(window, tk.Toplevel):
                    window.destroy()
    
            self.root.withdraw()
            self._wait_for_import_exit()

    def _wait_for_import_exit(self) -> None:
        """A háttérszál befejezésének megvárása a főszál blokkolása nélkül."""
        if self._import_thread is not None and self._import_thread.is_alive():
            self.root.after(100, self._wait_for_import_exit)
            return

        self.root.destroy()
        
    def _reset_and_exit(self) -> None:
        """A beállítások törlése megerősítés után, majd kilépés."""
        confirmed = messagebox.askyesno(
            title="Alaphelyzetbe állítás és bezárás",
            message=(
                "Figyelem! Ez törli a mentett riportmappa-útvonalat "
                "és a felvett alkatrészcsoportokat.\n"
                "Az Excel-riportok megmaradnak.\n\n"
                "Biztosan alaphelyzetbe szeretnéd állítani a programot?"
            ),
            parent=self.root,
        )

        if not confirmed:
            return

        try:
            reset()
        except SettingsResetError as exc:
            messagebox.showerror(
                title="Sikertelen alaphelyzetbe állítás",
                message=str(exc),
                parent=self.root,
            )
            return

        self._exit()

    def _request_csv_path(self) -> Path | None:
        """CSV fájl kiválasztása tallózással."""
        while True:
            try:
                filename = filedialog.askopenfilename(
                    parent=self.root,
                    title="Importálandó CSV kiválasztása",
                    filetypes=[
                        ("CSV fájlok", "*.csv"),
                        ("Minden fájl", "*"),
                    ],
                )

                if not filename:
                    return None

                path = Path(filename)

                if path.suffix.lower() != ".csv":
                    raise ValueError("CSV-fájlt válassz ki.")

                if not path.is_file():
                    raise ValueError("A kiválasztott fájl már nem érhető el.")

                return path.resolve()
            except (OSError, RuntimeError, ValueError) as exc:
                messagebox.showerror(
                    title="CSV-kiválasztási hiba",
                    message=str(exc),
                    parent=self.root,
                )

    def _start_import(self) -> None:
        """CSV kiválasztása és az import háttérszálának indítása."""
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
                title="Importálási hiba",
                message=f"Az importálás nem indítható el.\nRészletek: {exc}",
                parent=self.root,
            )

    def _request_rate(
        self,
        application_date: date,
    ) -> DailyExchangeRate:
        """Napi árfolyam bekérése és ellenőrzése."""
        selected_rate: DailyExchangeRate | None = None

        window = tk.Toplevel(self.root)
        window.title("Kézi árfolyam bevitele")
        window.resizable(False, False)
        window.transient(self.root)

        frame = ttk.Frame(window, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=f"Alkalmazás napja: {application_date:%Y-%m-%d}",
        ).pack(anchor="w", pady=(0, 15))

        ttk.Label(
            frame,
            text="EUR-árfolyam (Ft/EUR):",
        ).pack(anchor="w")

        rate_entry = ttk.Entry(frame, width=35)
        rate_entry.pack(fill="x", pady=(5, 15))

        ttk.Label(
            frame,
            text="Forrásdátum (ÉÉÉÉ-HH-NN):",
        ).pack(anchor="w")

        date_entry = ttk.Entry(frame, width=35)
        date_entry.pack(fill="x", pady=(5, 15))

        def show_input_error(message: str, entry: ttk.Entry) -> None:
            messagebox.showerror(
                title="Hibás árfolyamadat",
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
                    "Pozitív, véges, Excelben tárolható számot adj meg "
                    "(például: 400 vagy 400,25).",
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
                    "Érvényes dátumot adj meg ÉÉÉÉ-HH-NN formátumban "
                    "(például: 2026-10-12).",
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
            text="Kész",
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
        """Import közbeni tájékoztató üzenet megjelenítése."""
        messagebox.showinfo(
            title="Importálás",
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
        """Az adott év importelőnézete és mentésének jóváhagyása."""
        accepted = False

        window = tk.Toplevel(self.root)
        window.title(f"{year}. évi import előnézete")
        window.geometry("850x600")
        window.transient(self.root)

        frame = ttk.Frame(window, padding=15)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=f"Célfájl: {path}",
            wraplength=780,
        ).pack(anchor="w", pady=(0, 10))

        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True)

        table = ttk.Treeview(
            table_frame,
            columns=("period", "category", "amount"),
            show="headings",
        )

        table.heading("period", text="Nap / hónap")
        table.heading("category", text="Megnevezés")
        table.heading("amount", text="Hozzáadandó összeg (Ft)")

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
            "used_parts": "Bontott alkatrész",
            "new_parts": "Új alkatrész",
            "labor": "Munkadíj",
            "workshop_used_parts": "Bontott alkatrész – műhely",
            "workshop_new_parts": "Új alkatrész – műhely",
            "purchases": "Beszerzés",
            "shipping": "Szállítási költség",
            "trailer": "Trailer",
        }

        monthly_labels = {
            "sany": "Segédanyag",
            "diag": "Diagnosztika",
            "kj": "Klímajavítás",
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
                f"Új bizonylatok: {len(new_documents)}    "
                f"Kihagyott tételek: {len(skipped_movements)}"
            ),
        ).pack(anchor="w", pady=(10, 0))

        for movement in skipped_movements:
            table.insert(
                "",
                "end",
                values=(
                    movement.movement_date.isoformat(),
                    (
                        f"Kihagyva: {movement.document_number} / "
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
            text="Elutasítás",
            command=window.destroy,
        ).pack(side="left")

        ttk.Button(
            buttons,
            text="Elfogadás",
            command=accept,
        ).pack(side="right")

        window.wait_visibility()
        window.grab_set()
        self.root.wait_window(window)

        if self._closing:
            raise _ImportCancelled()

        return accepted

    def _check_cancelled(self) -> None:
        """Az import megszakítása ha elkezdődött a bezárás."""
        if self._closing:
            raise _ImportCancelled()

    def _request_missing_group(
        self,
        document_number: str,
        item_code: str,
    ) -> str:
        """Csoport kiválasztása egy csoport nélküli számlatételhez."""
        self._check_cancelled()
        selected_group: str | None = None

        window = tk.Toplevel(self.root)
        window.title("Hiányzó alkatrész csoport")
        window.transient(self.root)
        window.resizable(False, False)

        frame = ttk.Frame(window, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=(
                "Ennek a számlatételnek nincs csoportja.\n\n"
                f"Bizonylatszám: {document_number}\n"
                f"Cikkszám: {item_code}\n\n"
                "Válaszd ki a tétel csoportját:"
            ),
            wraplength=400,
            justify="left",
        ).pack(fill="x", pady=(0, 15))

        def select_group(group: str) -> None:
            nonlocal selected_group
            selected_group = group
            window.destroy()

        choices = (
            ("Új alkatrész", "Új alkatrész"),
            ("Bontott alkatrész", "Bontott alkatrész"),
            ("Munkadíj", "Munkadíj"),
            ("Szállítási költség", "Szállítási költség"),
            ("Trailer", "Trailer"),
            ("Egyik sem", ""),
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
        """Az import végrehajtása a háttérszálon."""
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
        """Az import eredményének megjelenítése a főszálon."""
        self._import_running = False

        if self._closing:
            return

        self.import_button.configure(state="normal")

        if error is not None:
            messagebox.showerror(
                title="Importálási hiba",
                message=str(error),
                parent=self.root,
            )

            if isinstance(error, (CsvFileAccessError, CsvFormatError)):
                self.root.after(0, self._start_import)

            return

        if cancelled:
            messagebox.showinfo(
                title="Importálás megszakítva",
                message=(
                    "Az importálást megszakítottad.\n"
                    "Az import közben már elmentett évek megmaradnak."
                ),
                parent=self.root,
            )
            return

        messagebox.showinfo(
            title="Importálás",
            message="Az importálási folyamat befejeződött.",
            parent=self.root,
        )

        if warning_messages:
            messagebox.showwarning(
                title="Munkafüzet-figyelmeztetések",
                message="\n\n".join(warning_messages),
                parent=self.root,
            )

    def _call_ui(self, function: Callable, *args):
        """Felületi művelet kérése a háttérszálból."""
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
        """A háttérszál felületi kéréseinek végrehajtása a főszálon."""
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
    """A grafikus alkalmazás elindítása."""
    root = tk.Tk()
    app = _Application(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()