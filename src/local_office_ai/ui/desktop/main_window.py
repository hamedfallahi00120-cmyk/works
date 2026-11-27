"""The main window: navigation, engine workers, theme and status wiring."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QMainWindow,
    QMessageBox,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.core.settings import Settings
from local_office_ai.ui.desktop import fonts as font_utils
from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.chrome import Banner, HardwarePanel, HeaderBar, Sidebar
from local_office_ai.ui.desktop.formatting import format_number
from local_office_ai.ui.desktop.pages.chat import ChatPage
from local_office_ai.ui.desktop.pages.file_analysis import FileAnalysisPage
from local_office_ai.ui.desktop.pages.folder_analysis import FolderAnalysisPage
from local_office_ai.ui.desktop.pages.home import HomePage
from local_office_ai.ui.desktop.pages.reports import ReportsPage
from local_office_ai.ui.desktop.pages.settings import SettingsPage
from local_office_ai.ui.desktop.state import AppState, validate_input_file, validate_workbook
from local_office_ai.ui.desktop.theme import apply_palette, build_stylesheet, palette_for
from local_office_ai.ui.desktop.ui_settings import UiSettings
from local_office_ai.ui.desktop.workers import (
    AnalysisWorker,
    ChatWorker,
    DocumentSetWorker,
    FolderScanWorker,
    FolderWorker,
    OllamaWorker,
    StageEvent,
)

PAGE_INFO = {
    "home": ("خانه", "نمای کلی، انتخاب منبع و شروع سریع"),
    "file": ("تحلیل فایل‌ها", "منابع، درخواست فارسی، مراحل واقعی موتور و نتیجه کامل"),
    "folder": ("تحلیل پوشه و منابع", "انتخاب منابع، زیرپوشه‌ها و تحلیل ترکیبی فایل‌ها"),
    "chat": ("گفت‌وگوی هوشمند", "پرسش تکمیلی روی همان داده با پاسخ‌های مستند"),
    "reports": ("گزارش‌های قبلی", "گزارش‌های ساخته‌شده در پوشه خروجی"),
    "settings": ("تنظیمات", "فونت فارسی، تم، مدل محلی، مسیرها و پایش"),
}


class MainWindow(QMainWindow):
    def __init__(self, settings: Settings, ui_settings: UiSettings) -> None:
        super().__init__()
        self.setWindowTitle("Local Office AI | دستیار هوشمند اداری")
        self.setMinimumSize(1024, 660)
        self.resize(1280, 840)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.state = AppState(settings, ui_settings, self)
        if self.state.ui.font_family:
            resolved_family = font_utils.resolve_family(self.state.ui.font_family)
            if resolved_family != self.state.ui.font_family:
                self.state.ui.font_family = resolved_family
                self.state.save()
        self._palette = palette_for(self.state.ui.theme)
        self._active_worker: Any = None
        self._ollama_worker: OllamaWorker | None = None
        self._scan_worker: FolderScanWorker | None = None
        self._scan_queue: list[Path] = []
        self._scan_base: list[Path] = []
        self._scan_found: list[Path] = []
        self._scan_rejected: list[str] = []
        self._manual_probe = False
        self._loaded_font_file = ""
        self._theme_signature = ""
        self._closing = False
        self._compact: bool | None = None
        self._floating_monitor: Any = None
        self._build()
        self._connect()
        self.apply_theme(force=True)
        self._restore_session()
        if self.state.ui.widget_visible:
            QTimer.singleShot(0, self._open_floating_monitor)
        QTimer.singleShot(180, self.probe_ollama)

    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        central = QWidget()
        central.setObjectName("Canvas")
        outer = QHBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.sidebar = Sidebar(self._palette)
        outer.addWidget(self.sidebar)

        content = QWidget()
        column = QVBoxLayout(content)
        column.setContentsMargins(24, 20, 20, 0)
        column.setSpacing(12)

        self.header = HeaderBar(self._palette)
        self.banner = Banner()
        self.stack = QStackedWidget()
        self.stack.setObjectName("PageStack")

        self.home_page = HomePage(self.state, self._palette)
        self.file_page = FileAnalysisPage(self.state, self._palette)
        self.folder_page = FolderAnalysisPage(self.state, self._palette)
        self.chat_page = ChatPage(self.state, self._palette)
        self.reports_page = ReportsPage(self.state, self._palette)
        self.settings_page = SettingsPage(self.state, self._palette)
        self.pages: dict[str, QWidget] = {
            "home": self.home_page,
            "file": self.file_page,
            "folder": self.folder_page,
            "chat": self.chat_page,
            "reports": self.reports_page,
            "settings": self.settings_page,
        }
        for page in self.pages.values():
            self.stack.addWidget(page)

        column.addWidget(self.header)
        column.addWidget(self.banner)
        column.addWidget(self.stack, 1)
        self.hardware_panel = HardwarePanel(self.state, self._palette)
        column.addWidget(self.hardware_panel)
        outer.addWidget(content, 1)
        self.setCentralWidget(central)
        self.setWindowIcon(QIcon(icons.pixmap("sparkles", self._palette.accent, 32)))
        self.show_page("home")

    # ------------------------------------------------------------------ #
    def resizeEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        super().resizeEvent(event)
        self._sync_compact()

    def _sync_compact(self) -> None:
        home = getattr(self, "home_page", None)
        if home is None:
            return
        compact = self.height() < 820
        if compact == self._compact:
            return
        self._compact = compact
        home.set_compact(compact)

    # ------------------------------------------------------------------ #
    def _connect(self) -> None:
        self.sidebar.navigate.connect(self.show_page)
        self.header.refresh_ollama.connect(lambda: self.probe_ollama(manual=True))
        self.header.open_settings.connect(lambda: self.show_page("settings"))
        self.header.toggle_monitor.connect(self._toggle_floating_monitor)
        self.settings_page.widget_toggled.connect(self._set_floating_monitor_visible)
        self.settings_page.widget_opacity_changed.connect(self._apply_widget_opacity)
        self.settings_page.widget_autostart_changed.connect(self._apply_widget_autostart)

        home = self.home_page
        home.browse_file.connect(self._choose_file)
        home.browse_folder.connect(self._choose_folder)
        home.analyze_requested.connect(self._home_start)
        home.navigate.connect(self.show_page)
        home.open_last.connect(self._open_last)
        home.remove_source.connect(self._remove_file)
        home.file_dropped.connect(self._file_dropped)
        home.files_dropped.connect(self._files_dropped)

        file_page = self.file_page
        file_page.browse_requested.connect(self._choose_file)
        file_page.add_folder_requested.connect(self._add_folder_to_sources)
        file_page.remove_requested.connect(self._remove_file)
        file_page.file_dropped.connect(self._file_dropped)
        file_page.files_dropped.connect(self._files_dropped)
        file_page.analyze_requested.connect(self._start_file)
        file_page.stop_requested.connect(self._cancel_active)
        file_page.followup_requested.connect(lambda: self.show_page("chat"))

        folder_page = self.folder_page
        folder_page.browse_requested.connect(self._choose_folder)
        folder_page.refresh_requested.connect(folder_page.scan)
        folder_page.folder_dropped.connect(self._folder_dropped)
        folder_page.analyze_requested.connect(self._start_folder)
        folder_page.stop_requested.connect(self._cancel_active)
        folder_page.followup_requested.connect(self._folder_followup)

        self.chat_page.ask_requested.connect(self._start_chat)
        self.chat_page.attach_requested.connect(self._chat_attach)

        self.settings_page.refresh_ollama.connect(lambda: self.probe_ollama(manual=True))
        self.settings_page.saved.connect(lambda message: self.banner.show_message(message, "success", timeout_ms=4000))

        self.state.settings_changed.connect(self._on_settings_changed)
        self.state.ollama_changed.connect(self._refresh_ollama_header)
        self.state.busy_changed.connect(self._busy_changed)
        self.state.analysis_changed.connect(self._on_new_analysis)
        self.state.folder_result_changed.connect(self._on_new_folder_result)

    # ------------------------------------------------------------------ #
    def show_page(self, key: str) -> None:
        page = self.pages.get(key)
        if page is None:
            return
        if key == "reports":
            self.reports_page.refresh()
        self.stack.setCurrentWidget(page)
        title, subtitle = PAGE_INFO.get(key, ("", ""))
        self.header.set_page(title, subtitle)
        self.sidebar.set_current(key)
        if key == "chat":
            self.chat_page._refresh_chip()

    # ------------------------------------------------------------------ #
    def _restore_session(self) -> None:
        ui = self.state.ui
        if ui.last_request:
            self.file_page.set_request(ui.last_request)
            self.home_page.composer.set_request(ui.last_request)
        restored = False
        if ui.last_source:
            path = Path(ui.last_source)
            valid, _ = validate_workbook(path)
            if valid:
                self.state.set_source(path)
                restored = True
        if not restored and ui.last_folder and Path(ui.last_folder).is_dir():
            self.state.set_folder(Path(ui.last_folder))
        if restored:
            self.banner.show_message(
                "فایل انتخاب‌شده قبلی بازیابی شد. برای تحلیل، درخواست را بنویسید و شروع کنید.",
                "info",
                timeout_ms=6000,
            )

    # ------------------------------------------------------------------ #
    def _ensure_floating_monitor(self) -> Any:
        if self._floating_monitor is None:
            from local_office_ai.ui.desktop.floating_monitor import FloatingMonitor

            widget = FloatingMonitor(
                self.state,
                self._palette,
                on_open_main=self._raise_main_window,
            )
            widget.quit_requested.connect(self.close)
            widget.visibility_changed.connect(self._on_floating_visibility)
            self._floating_monitor = widget
        return self._floating_monitor

    def show_floating_monitor(self) -> None:
        widget = self._ensure_floating_monitor()
        if widget is not None and not widget.isVisible():
            widget.show_widget()

    def _open_floating_monitor(self) -> None:
        if self._closing:
            return
        self.show_floating_monitor()

    def _toggle_floating_monitor(self) -> None:
        widget = self._ensure_floating_monitor()
        if widget is None:
            return
        if widget.isVisible():
            widget.hide_to_tray()
        else:
            widget.show_widget()

    def _set_floating_monitor_visible(self, visible: bool) -> None:
        if visible:
            self._open_floating_monitor()
            return
        widget = self._floating_monitor
        if widget is not None:
            widget.deactivate()
        else:
            self.state.set_widget_options(visible=False)

    def _apply_widget_opacity(self, value: float) -> None:
        self._ensure_floating_monitor().apply_opacity(value)

    def _apply_widget_autostart(self, enabled: bool) -> None:
        from local_office_ai.ui.desktop.floating_monitor import set_autostart

        applied = set_autostart(enabled, self.state.project_root)
        if applied:
            self.state.set_widget_options(autostart=enabled)
        self.settings_page.finish_autostart_change(enabled, applied)

    def _on_floating_visibility(self, visible: bool) -> None:
        self.settings_page.sync_widget_visibility(visible)

    def _raise_main_window(self) -> None:
        self.showNormal()
        self.raise_()
        self.activateWindow()

    # ------------------------------------------------------------------ #
    def _choose_file(self) -> bool:
        if self.state.busy:
            self.banner.show_message("در حین پردازش، منبع را تغییر ندهید.", "warning")
            return False
        current = (self.state.sources[0] if self.state.sources else self.state.source)
        start_dir = str(current.parent) if current else str(self.state.settings.input_dir)
        file_filter = (
            "فایل‌های پشتیبانی‌شده (*.xlsx *.xlsm *.csv *.docx *.pptx *.pdf *.txt *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp);;"
            "Excel (*.xlsx *.xlsm);;Word (*.docx);;PowerPoint (*.pptx);;PDF (*.pdf);;CSV (*.csv);;متن (*.txt);;تصویر (*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp)"
        )
        paths, _ = QFileDialog.getOpenFileNames(
            self, "انتخاب یک یا چند فایل", start_dir, file_filter,
        )
        if not paths:
            return False
        accepted: list[Path] = []
        rejected: list[str] = []
        for item in paths:
            candidate = Path(item)
            valid, error = validate_input_file(candidate)
            if valid:
                accepted.append(candidate)
            else:
                rejected.append(f"{candidate.name}: {error}")
        if not accepted:
            self.banner.show_message("هیچ‌یک از فایل‌های انتخاب‌شده معتبر نبودند: " + "؛ ".join(rejected[:3]), "danger", timeout_ms=9000)
            return False
        existing = list(getattr(self.state, "sources", []) or [])
        self._set_source_basket([*existing, *accepted], rejected)
        return True

    def _accept_file(self, path: Path) -> bool:
        valid, error = validate_input_file(path)
        if not valid:
            self.banner.show_message(f"این فایل قابل پردازش نیست: {error}", "danger")
            return False
        self.state.set_sources([path])
        self.file_page.set_sources([path])
        self.banner.show_message(f"فایل انتخاب شد: {path.name}", "success", timeout_ms=4000)
        return True

    def _file_dropped(self, path: str) -> None:
        if self.state.busy:
            self.banner.show_message("در حین پردازش، منبع را تغییر ندهید.", "warning")
            return
        candidate = Path(path)
        if candidate.is_dir():
            self.state.set_folder(candidate)
            self.show_page("folder")
            self.folder_page.scan()
            return
        existing = list(getattr(self.state, "sources", []) or [])
        self._set_source_basket([*existing, candidate], [])

    def _files_dropped(self, paths: list[str]) -> None:
        """Accept a mixed drag-drop set of files and folders into one source basket."""
        if self.state.busy:
            self.banner.show_message("در حین پردازش، منبع را تغییر ندهید.", "warning")
            return
        candidates: list[Path] = []
        folders: list[Path] = []
        rejected: list[str] = []
        for raw in paths:
            item = Path(raw).expanduser()
            if item.is_dir():
                folders.append(item)
            elif item.is_file():
                valid, error = validate_input_file(item)
                if valid:
                    candidates.append(item)
                else:
                    rejected.append(f"{item.name}: {error}")
        existing = list(getattr(self.state, "sources", []) or [])
        if not folders:
            self._set_source_basket([*existing, *candidates], rejected)
            return
        self._scan_folders(folders, [*existing, *candidates], rejected)

    def _add_folder_to_sources(self) -> bool:
        """Append every supported file in a selected folder to the current file basket."""
        if self.state.busy:
            self.banner.show_message("در حین پردازش، منبع را تغییر ندهید.", "warning")
            return False
        current = list(getattr(self.state, "sources", []) or [])
        start_dir = str(current[0].parent) if current else str(self.state.settings.input_dir)
        folder = QFileDialog.getExistingDirectory(self, "افزودن پوشه و زیرپوشه‌ها به فهرست منابع", start_dir)
        if not folder:
            return False
        self._scan_folders([Path(folder)], current, [])
        return True

    # -- folder discovery runs off the interface thread ------------------ #
    def _scan_folders(self, folders: list[Path], base: list[Path], rejected: list[str]) -> None:
        if self._scan_worker is not None:
            self.banner.show_message("فهرست پوشه قبلی در حال خواندن است؛ کمی بعد تلاش کنید.", "warning")
            return
        self._scan_queue = [Path(item) for item in folders]
        self._scan_base = list(base)
        self._scan_rejected = list(rejected)
        self._scan_found = []
        self._scan_next()

    def _scan_next(self) -> None:
        if not self._scan_queue:
            self._finish_scan()
            return
        folder = self._scan_queue.pop(0)
        self.banner.show_message(f"در حال خواندن فایل‌های پوشه: {folder.name}…", "info", timeout_ms=0)
        worker = FolderScanWorker(self.state.settings, folder, recursive=True)
        worker.completed.connect(self._scan_collected)
        worker.failed.connect(self._scan_failed)
        worker.finished.connect(lambda w=worker: self._scan_worker_finished(w))
        self._scan_worker = worker
        worker.start()

    def _scan_collected(self, files: Any) -> None:
        self._scan_found.extend(Path(str(item.get("path", ""))) for item in list(files or []))
        self._scan_next()

    def _scan_failed(self, message: str) -> None:
        self._scan_rejected.append(message)
        self._scan_next()

    def _scan_worker_finished(self, worker: Any) -> None:
        if self._scan_worker is worker:
            self._scan_worker = None
        worker.deleteLater()

    def _finish_scan(self) -> None:
        found = list(self._scan_found)
        base = list(self._scan_base)
        rejected = list(self._scan_rejected)
        self._scan_found = []
        self._scan_base = []
        self._scan_rejected = []
        if not found:
            self.banner.show_message(
                "در پوشه انتخاب‌شده فایل پشتیبانی‌شده‌ای پیدا نشد."
                + ((" " + "؛ ".join(rejected[:3])) if rejected else ""),
                "warning",
                timeout_ms=8000,
            )
            if not base:
                return
        self._set_source_basket([*base, *found], rejected)

    def _set_source_basket(self, candidates: list[Path], rejected: list[str]) -> None:
        unique: list[Path] = []
        seen: set[str] = set()
        for candidate in candidates:
            valid, error = validate_input_file(candidate)
            if not valid:
                rejected.append(f"{candidate.name}: {error}")
                continue
            resolved = candidate.expanduser().resolve()
            key = str(resolved).casefold()
            if key not in seen:
                unique.append(resolved)
                seen.add(key)
        if not unique:
            message = "هیچ فایل پشتیبانی‌شده‌ای پیدا نشد."
            if rejected:
                message += " " + "؛ ".join(rejected[:3])
            self.banner.show_message(message, "warning", timeout_ms=9000)
            return
        self.state.set_sources(unique)
        self.file_page.set_sources(unique)
        message = f"{len(unique)} منبع برای تحلیل آماده شد."
        if rejected:
            message += f" {len(rejected)} مورد نامعتبر کنار گذاشته شد."
        self.banner.show_message(message, "success" if not rejected else "warning", timeout_ms=6500)
        self.show_page("file")

    def _choose_folder(self) -> bool:
        if self.state.busy:
            self.banner.show_message("در حین پردازش، منبع را تغییر ندهید.", "warning")
            return False
        start_dir = str(self.state.folder) if self.state.folder else str(self.state.settings.input_dir)
        path = QFileDialog.getExistingDirectory(self, "انتخاب پوشه حاوی فایل‌های اداری و آماری", start_dir)
        if not path:
            return False
        self.state.set_folder(Path(path))
        self.show_page("folder")
        self.banner.show_message(f"پوشه انتخاب شد: {Path(path).name}", "success", timeout_ms=4000)
        return True

    def _folder_dropped(self, path: str) -> None:
        if self.state.busy:
            self.banner.show_message("در حین پردازش، منبع را تغییر ندهید.", "warning")
            return
        candidate = Path(path)
        if not candidate.is_dir():
            self.banner.show_message("برای تحلیل پوشه، یک پوشه انتخاب کنید.", "warning")
            return
        self.state.set_folder(candidate)
        self.folder_page.scan()

    def _remove_file(self, path: str = "") -> None:
        if self.state.busy:
            self.banner.show_message("در حین پردازش، منبع را تغییر ندهید.", "warning")
            return
        if path and Path(path).is_dir():
            self.state.set_folder(None)
            self.banner.show_message("پوشه از فهرست حذف شد.", "info", timeout_ms=3000)
            return
        remaining = (
            []
            if not str(path)
            else [
                item
                for item in list(getattr(self.state, "sources", []) or [])
                if str(item) != str(path)
            ]
        )
        self.state.set_sources(remaining)
        self.file_page.set_sources(remaining)
        self.banner.show_message(
            "منبع از فهرست حذف شد." if remaining else "فهرست منابع خالی شد.",
            "info",
            timeout_ms=3000,
        )

    # ------------------------------------------------------------------ #
    def _home_start(self, request: str) -> None:
        if getattr(self.state, "sources", None):
            self._start_file(request)
        elif self.state.source is not None:
            self._start_file(request)
        elif self.state.folder is not None:
            self._start_folder(request)
        else:
            self.banner.show_message("ابتدا یک فایل یا پوشه انتخاب کنید.", "warning")
            self.show_page("file")

    def _start_file(self, request: str) -> None:
        if self.state.busy:
            self.banner.show_message("یک پردازش در جریان است؛ پایان آن را منتظر بمانید.", "warning")
            return
        from local_office_ai.ingestion.reader import EXCEL_SUFFIXES
        sources = list(getattr(self.state, "sources", []) or [])
        if not sources and self.state.source is not None:
            sources = [self.state.source]
        if not sources:
            self.show_page("file")
            self.file_page.composer.set_hint("ابتدا یک یا چند فایل انتخاب کنید.", "warning")
            return
        invalid = []
        for source in sources:
            valid, error = validate_input_file(source)
            if not valid:
                invalid.append(f"{source.name}: {error}")
        if invalid:
            message = "فایل‌های نامعتبر: " + "؛ ".join(invalid[:4])
            self.file_page.finish_error(message)
            self.banner.show_message(message, "danger", timeout_ms=10000)
            return
        self.state.set_last_request(request)
        # Keep the highly specialized Excel workflow for one workbook.
        if len(sources) == 1 and sources[0].suffix.lower() in EXCEL_SUFFIXES:
            source = sources[0]
            self.show_page("file")
            self.file_page.set_sources(sources)
            self.file_page.begin_run()
            self._set_busy(True)
            QTimer.singleShot(0, lambda: self.file_page.reveal(self.file_page.progress_card))
            worker = AnalysisWorker(
                self.state.settings, source, request=request, model=self.state.analysis_model,
                use_ai=self.state.use_ai,
            )
            worker.stage.connect(self.file_page.on_stage)
            worker.progress.connect(self.file_page.on_progress)
            worker.completed.connect(self._file_done)
            worker.failed.connect(self._file_failed)
            worker.confirm_needed.connect(self._confirm_structure)
            self._run_worker(worker)
            return
        # Multiple files or non-Excel documents use the shared mixed-format pipeline.
        self.show_page("folder")
        self.folder_page.folder_card.set_folder(None)
        self.folder_page.folder_card.path_label.setText(f"{len(sources)} منبع انتخاب‌شده")
        self.folder_page.begin_run()
        self._set_busy(True)
        QTimer.singleShot(0, lambda: self.folder_page.reveal(self.folder_page.progress_card))
        worker = DocumentSetWorker(
            self.state.settings, paths=sources, request=request, model=self.state.analysis_model,
            use_ai=self.state.use_ai,
        )
        worker.progress.connect(self.folder_page.on_progress)
        worker.completed.connect(self._folder_done)
        worker.failed.connect(self._folder_failed)
        self._run_worker(worker)

    def _start_folder(self, request: str) -> None:
        if self.state.busy:
            self.banner.show_message("یک پردازش در جریان است؛ پایان آن را منتظر بمانید.", "warning")
            return
        folder = self.state.folder
        if folder is None or not folder.is_dir():
            self.show_page("folder")
            self.folder_page.composer.set_hint("ابتدا یک پوشه انتخاب کنید.", "warning")
            return
        selected = self.folder_page.selected_paths()
        if self.folder_page.folder_card.file_table.rowCount() > 0 and not selected:
            self.folder_page.composer.set_hint("حداقل یک فایل را برای تحلیل تیک بزنید.", "warning")
            self.banner.show_message("هیچ فایلی برای تحلیل انتخاب نشده است.", "warning")
            return
        self.state.set_last_request(request)
        self.show_page("folder")
        self.folder_page.begin_run()
        self._set_busy(True)
        QTimer.singleShot(0, lambda: self.folder_page.reveal(self.folder_page.progress_card))
        worker = FolderWorker(
            self.state.settings,
            folder,
            request=request,
            model=self.state.analysis_model,
            use_ai=self.state.use_ai,
            selected_files=selected or None,
            recursive=self.folder_page.folder_card.recursive_check.isChecked(),
        )
        worker.progress.connect(self.folder_page.on_progress)
        worker.completed.connect(self._folder_done)
        worker.failed.connect(self._folder_failed)
        self._run_worker(worker)

    # ------------------------------------------------------------------ #
    def _start_chat(self, question: str) -> None:
        if self.state.busy:
            self.chat_page.show_error("یک پردازش دیگر در جریان است؛ پس از پایان آن دوباره بپرسید.")
            return
        session = self.chat_page.session_analysis()
        document_set_session = self.chat_page.session_document_set()
        source = self.state.source if session is None and document_set_session is None else None
        selected_sources = list(getattr(self.state, "sources", []) or []) if session is None and document_set_session is None else []
        if session is None and document_set_session is None and source is None and not selected_sources:
            self.chat_page.show_error(
                "برای گفت‌وگو ابتدا یک فایل Excel انتخاب کنید. "
                "پس از تحلیل فایل، پاسخ‌ها روی همان داده ساخته می‌شوند."
            )
            return
        if session is None and document_set_session is None and selected_sources and (len(selected_sources) != 1 or selected_sources[0].suffix.lower() not in {".xlsx", ".xlsm"}):
            self.chat_page.show_error("برای گفت‌وگوی تکمیلی روی چند سند، ابتدا تحلیل ترکیبی را اجرا کنید.")
            return
        if source is not None:
            valid, error = validate_workbook(source)
            if not valid:
                self.chat_page.show_error(f"فایل قابل پردازش نیست: {error}")
                return
        self.show_page("chat")
        self.chat_page.set_busy(True)
        self._set_busy(True)
        worker = ChatWorker(
            self.state.settings,
            question,
            analysis=session,
            source=source,
            document_set_result=document_set_session,
            model=self.state.analysis_model,
            use_ai=self.state.use_ai,
        )
        worker.stage.connect(self.chat_page.show_stage)
        worker.completed.connect(self._chat_done)
        worker.failed.connect(self._chat_failed)
        self._run_worker(worker)

    def _folder_followup(self, analysis: Any) -> None:
        if self.state.busy:
            self.banner.show_message("یک پردازش در جریان است؛ پس از پایان آن دوباره تلاش کنید.", "warning")
            return
        if hasattr(analysis, "ai_mode") and hasattr(analysis, "statistics") and hasattr(analysis, "files"):
            self.chat_page.attach_document_set(analysis, reset=True)
            self.show_page("chat")
            self.banner.show_message("همه منابع تحلیل‌شده به گفت‌وگو متصل شدند.", "info", timeout_ms=5000)
            return
        path = Path(str(getattr(getattr(analysis, "meta", None), "path", "") or ""))
        if not path.is_file():
            self.banner.show_message("فایل این نتیجه پیدا نشد؛ پرسش تکمیلی ممکن نیست.", "warning")
            return
        self.chat_page.attach_dataset(analysis, path, reset=True)
        self.show_page("chat")
        self.banner.show_message("داده فایل انتخاب‌شده به گفت‌وگو متصل شد.", "info", timeout_ms=5000)

    def _chat_attach(self) -> None:
        if self.state.busy:
            self.banner.show_message("پردازش در جریان است؛ کمی بعد تلاش کنید.", "warning")
            return
        if self.chat_page.session_analysis() is not None:
            self.banner.show_message("داده‌ای به گفت‌وگو متصل است؛ پرسش خود را بنویسید.", "info", timeout_ms=4000)
            return
        if not getattr(self.state, "sources", None):
            if not self._choose_file():
                return
        self.chat_page.ask("این اطلاعات را بررسی کن، تغییرات مهم را پیدا کن و گزارش مدیریتی تهیه کن.")

    # ------------------------------------------------------------------ #
    def _run_worker(self, worker: Any) -> None:
        self._active_worker = worker
        worker.finished.connect(self._worker_finished)
        worker.start()

    def _worker_finished(self) -> None:
        worker = self._active_worker
        self._active_worker = None
        if worker is not None:
            worker.deleteLater()
        if self.state.busy:
            self._set_busy(False)

    def _cancel_active(self) -> None:
        worker = self._active_worker
        if worker is None:
            return
        worker.cancel()
        self.banner.show_message(
            "درخواست توقف ثبت شد؛ پردازش در پایان مرحله جاری متوقف می‌شود.",
            "warning",
            timeout_ms=6000,
        )

    def _set_busy(self, busy: bool) -> None:
        self.state.set_busy(busy)
        self.file_page.set_running(busy)
        self.folder_page.set_running(busy)
        self.home_page.composer.set_running(busy)

    def _busy_changed(self) -> None:
        busy = self.state.busy
        self.file_page.set_running(busy)
        self.folder_page.set_running(busy)
        self.home_page.composer.set_running(busy)

    # ------------------------------------------------------------------ #
    def _file_done(self, result: Any) -> None:
        self._set_busy(False)
        self.state.set_result(result)
        self.file_page.finish_success(result)
        QTimer.singleShot(0, lambda: self.file_page.reveal(self.file_page.results))
        analysis = getattr(result, "analysis", None)
        if analysis is not None:
            self.chat_page.attach_dataset(analysis, Path(str(analysis.meta.path)))
        status = getattr(result, "status", "completed")
        if status == "completed":
            self.banner.show_message("تحلیل کامل شد و گزارش Excel ساخته شد.", "success")
        else:
            self.banner.show_message("تحلیل با هشدار کامل شد؛ بخش هشدار‌ها را ببینید.", "warning")
        self.reports_page.refresh()

    def _file_failed(self, message: str) -> None:
        self._set_busy(False)
        self.file_page.finish_error(message)
        self.banner.show_message(message, "danger", timeout_ms=12000)

    def _folder_done(self, result: Any) -> None:
        self._set_busy(False)
        self.state.set_folder_result(result)
        if hasattr(result, "ai_mode") and hasattr(result, "statistics"):
            self.chat_page.attach_document_set(result, reset=True)
        self.folder_page.finish_success(result)
        QTimer.singleShot(0, lambda: self.folder_page.reveal(self.folder_page.results))
        if str(getattr(result, "status", "")) == "completed":
            self.banner.show_message("تحلیل منابع کامل شد.", "success")
        else:
            self.banner.show_message("تحلیل منابع با هشدار کامل شد؛ جزئیات هر فایل را ببینید.", "warning")
        self.reports_page.refresh()

    def _folder_failed(self, message: str) -> None:
        self._set_busy(False)
        self.folder_page.finish_error(message)
        self.banner.show_message(message, "danger", timeout_ms=12000)

    def _chat_done(self, reply: Any) -> None:
        self._set_busy(False)
        self.chat_page.show_reply(reply)
        result = getattr(reply, "result", None)
        if result is not None:
            self.state.set_result(result)
            self.file_page.show_result(result)
            self.reports_page.refresh()
        self.chat_page.set_busy(False)

    def _chat_failed(self, message: str) -> None:
        self._set_busy(False)
        self.chat_page.set_busy(False)
        self.chat_page.show_error(message)

    def _on_new_analysis(self) -> None:
        self.home_page._refresh_last()

    def _on_new_folder_result(self) -> None:
        self.home_page._refresh_last()

    def _open_last(self) -> None:
        if self.state.last_result_kind == "folder" and self.state.last_folder_result is not None:
            self.folder_page.show_result(self.state.last_folder_result)
            self.show_page("folder")
            QTimer.singleShot(0, lambda: self.folder_page.reveal(self.folder_page.results))
        elif self.state.last_result is not None:
            self.file_page.show_result(self.state.last_result)
            self.show_page("file")
            QTimer.singleShot(0, lambda: self.file_page.reveal(self.file_page.results))
        elif self.state.last_folder_result is not None:
            self.folder_page.show_result(self.state.last_folder_result)
            self.show_page("folder")
            QTimer.singleShot(0, lambda: self.folder_page.reveal(self.folder_page.results))

    # ------------------------------------------------------------------ #
    def _confirm_structure(self, analysis: Any) -> None:
        worker = self._active_worker
        tables = list(getattr(analysis, "tables", []) or [])[:6]
        lines = [
            f"{table.sheet}!{table.address} — {table.kind.value} — اطمینان {table.confidence:.0%}"
            for table in tables
        ]
        text = "ساختار فایل با اطمینان پایین شناسایی شد:\n\n" + "\n".join(lines)
        primary = getattr(analysis, "primary_table", None)
        if primary is not None:
            text += (
                f"\n\nجدول انتخاب‌شده: {primary.sheet}!{primary.address}"
                "\nادامه تحلیل را تأیید می‌کنید؟"
            )
        answer = QMessageBox.question(
            self,
            "تأیید ساختار",
            text,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if worker is not None:
            worker.resolve_confirm(answer == QMessageBox.StandardButton.Yes)

    # ------------------------------------------------------------------ #
    def probe_ollama(self, manual: bool = False) -> None:
        if self._ollama_worker is not None and self._ollama_worker.isRunning():
            return
        self._manual_probe = manual
        if manual:
            self.header.ollama_chip.set_status("در حال بررسی…", "info")
        worker = OllamaWorker(self.state.settings)
        worker.completed.connect(self._ollama_done)
        worker.finished.connect(self._ollama_cleanup)
        self._ollama_worker = worker
        worker.start()

    def _ollama_done(self, online: bool, models: Any, error: str) -> None:
        self.state.set_models(bool(online), list(models or []), error)
        self.header.update_ollama(bool(online), self.state.model, error)
        if self._manual_probe:
            if online:
                self.banner.show_message(
                    f"Ollama متصل است؛ {len(models or [])} مدل پیدا شد.", "success", timeout_ms=4000
                )
            else:
                self.banner.show_message(
                    "Ollama در دسترس نیست؛ تحلیل با قواعد قطعی Python انجام می‌شود. " + error,
                    "warning",
                    timeout_ms=10000,
                )
        self._manual_probe = False

    def _ollama_cleanup(self) -> None:
        worker = self._ollama_worker
        self._ollama_worker = None
        if worker is not None:
            worker.deleteLater()

    def _refresh_ollama_header(self) -> None:
        self.header.update_ollama(self.state.ollama_online, self.state.model, self.state.ollama_error)

    # ------------------------------------------------------------------ #
    def apply_theme(self, force: bool = False) -> None:
        ui = self.state.ui
        signature = (ui.theme, ui.font_family, ui.font_base_size, ui.font_heading_size, ui.font_file)
        if not force and signature == self._theme_signature:
            return
        self._theme_signature = signature
        app = QApplication.instance()
        if app is None:
            return
        self._palette = palette_for(ui.theme)
        if ui.font_file and Path(ui.font_file).is_file() and ui.font_file != self._loaded_font_file:
            font_utils.load_font_file(ui.font_file)
            self._loaded_font_file = ui.font_file
        family = font_utils.apply_font(app, ui.font_family, ui.font_base_size)
        icons.clear_cache()
        app.setStyleSheet(
            build_stylesheet(
                self._palette,
                family=family,
                base_size=ui.font_base_size,
                heading_size=ui.font_heading_size,
            )
        )
        apply_palette(app, self._palette)
        app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.sidebar.apply_palette(self._palette)
        self.header.apply_palette(self._palette)
        self.banner.apply_palette(self._palette)
        self.hardware_panel.apply_palette(self._palette)
        if self._floating_monitor is not None:
            self._floating_monitor.apply_palette(self._palette)
        for page in self.pages.values():
            apply = getattr(page, "apply_palette", None)
            if callable(apply):
                apply(self._palette)
        self.setWindowIcon(QIcon(icons.pixmap("sparkles", self._palette.accent, 32)))

    def _on_settings_changed(self) -> None:
        self.apply_theme()
        self.header.update_ollama(
            self.state.ollama_online,
            self.state.model,
            self.state.ollama_error,
        )
        self._refresh_ollama_header()

    # ------------------------------------------------------------------ #
    def closeEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        self._closing = True
        self.settings_page.flush_font()
        worker = self._active_worker
        if worker is not None and worker.isRunning():
            answer = QMessageBox.question(
                self,
                "پردازش در جریان است",
                "یک پردازش در حال اجراست. بستن برنامه آن را متوقف می‌کند. بسته شود؟",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                self._closing = False
                event.ignore()
                return
            worker.cancel()
            if not worker.wait(6000):
                worker.terminate()
                worker.wait(1000)
        self.hardware_panel.shutdown()
        scan = self._scan_worker
        if scan is not None and scan.isRunning():
            if not scan.wait(3000):
                scan.terminate()
                scan.wait(500)
        if self._floating_monitor is not None:
            self._floating_monitor.shutdown()
            self._floating_monitor = None
        self.folder_page.shutdown()
        self.reports_page.shutdown()
        probe = self._ollama_worker
        if probe is not None and probe.isRunning():
            if not probe.wait(1200):
                probe.terminate()
                probe.wait(500)
        event.accept()
