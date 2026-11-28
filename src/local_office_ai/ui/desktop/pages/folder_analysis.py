"""Folder analysis page: one request, every workbook of a folder."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
    QCheckBox,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    BusyBar,
    CallbackButton,
    Card,
    SectionHeader,
    make_label,
)
from local_office_ai.ui.desktop.folder_results import FolderResultsView
from local_office_ai.ui.desktop.formatting import align_table_header, elide, format_number, format_bytes, ltr
from local_office_ai.ui.desktop.pages.base import Page
from local_office_ai.ui.desktop.pages.file_analysis import ErrorCard
from local_office_ai.ui.desktop.source import SUPPORTED_FORMATS, RequestComposer
from local_office_ai.ui.desktop.theme import Palette
from local_office_ai.ui.desktop.workers import FolderScanWorker, StageEvent


class FolderCard(Card):
    browse_requested = Signal()
    refresh_requested = Signal()
    folder_dropped = Signal(str)
    selection_changed = Signal()

    def __init__(self, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._palette = palette
        self.setAcceptDrops(True)
        self.add(
            SectionHeader(
                "پوشه داده",
                f"فایل‌های پشتیبانی‌شده در پوشه و زیرپوشه‌ها پیدا می‌شوند. فرمت‌های پذیرفته‌شده: {SUPPORTED_FORMATS}",
            )
        )
        self.path_box = Card(variant="sunken", margins=(14, 12, 14, 12), spacing=6)
        path_row = QHBoxLayout()
        self.folder_icon = QLabel()
        self.folder_icon.setPixmap(icons.pixmap("folder", palette.accent, 26))
        texts = QHBoxLayout()
        self.path_label = QLabel("پوشه‌ای انتخاب نشده است")
        self.path_label.setObjectName("CardTitle")
        self.path_label.setWordWrap(True)
        self.path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        texts.addWidget(self.path_label, 1)
        self.choose_button = CallbackButton(
            "انتخاب پوشه", self.browse_requested.emit, object_name="Primary", icon_name="folder-open"
        )
        path_row.addWidget(self.folder_icon)
        path_row.addLayout(texts, 1)
        path_row.addWidget(self.choose_button)
        self.path_box.add_layout(path_row)
        self.add(self.path_box)

        list_header = QHBoxLayout()
        self.recursive_check = QCheckBox("بررسی زیرپوشه‌ها")
        self.recursive_check.setChecked(True)
        self.recursive_check.setToolTip("در صورت فعال بودن، پوشه‌های داخلی نیز بررسی می‌شوند.")
        self.recursive_check.toggled.connect(self.refresh_requested.emit)
        self.count_label = QLabel("فایل‌های قابل پردازش: —")
        self.count_label.setObjectName("Muted")
        self.refresh_button = CallbackButton(
            "بازخوانی فهرست", self.refresh_requested.emit, object_name="Quiet", icon_name="refresh"
        )
        list_header.addWidget(self.count_label, 1)
        list_header.addWidget(self.recursive_check)
        list_header.addWidget(self.refresh_button)
        self.add_layout(list_header)

        self.file_table = QTableWidget()
        self.file_table.setColumnCount(3)
        self.file_table.setHorizontalHeaderLabels(["شامل", "فایل", "حجم"])
        align_table_header(self.file_table)
        self.file_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.file_table.setAlternatingRowColors(True)
        self.file_table.horizontalHeader().setStretchLastSection(False)
        self.file_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.ResizeToContents
        )
        self.file_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.file_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.ResizeToContents
        )
        self.file_table.verticalHeader().setVisible(False)
        self.file_table.verticalHeader().setDefaultSectionSize(36)
        self.file_table.setShowGrid(False)
        self.file_table.setMinimumHeight(120)
        self.file_table.setMaximumHeight(260)
        self.file_table.itemChanged.connect(self._selection_changed)
        self.add(self.file_table)

        self.scan_status = make_label("برای دیدن فایل‌ها یک پوشه انتخاب کنید.", "Caption", wrap=True)
        self.add(self.scan_status)
        self.add(
            make_label(
                "فقط فایل‌های تیک‌خورده در یک تحلیل ترکیبی پردازش می‌شوند؛ محاسبات با Python و "
                "تفسیر متنی با مدل محلی انجام می‌شود.",
                "Caption",
                wrap=True,
            )
        )

    def set_folder(self, path: Path | None) -> None:
        if path is None:
            self.path_label.setText("پوشه‌ای انتخاب نشده است")
            self.path_label.setToolTip("")
            self.count_label.setText("فایل‌های قابل پردازش: —")
            self.file_table.setRowCount(0)
            self.file_table.setVisible(False)
            self.scan_status.setText("برای دیدن فایل‌ها یک پوشه انتخاب کنید.")
            self.scan_status.setObjectName("Caption")
            self.scan_status.style().unpolish(self.scan_status)
            self.scan_status.style().polish(self.scan_status)
            return
        self.path_label.setText(ltr(str(path)))
        self.path_label.setToolTip(str(path))

    def set_files(self, files: list[dict[str, Any]]) -> None:
        self.file_table.blockSignals(True)
        self.file_table.clearContents()
        self.file_table.setRowCount(len(files))
        for row, item in enumerate(files):
            check = QTableWidgetItem("")
            check.setFlags(
                Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
            )
            check.setCheckState(Qt.CheckState.Checked)
            check.setData(Qt.ItemDataRole.UserRole, str(item.get("path", "")))
            check.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            name = QTableWidgetItem(ltr(item.get("name", "")))
            size = QTableWidgetItem(ltr(format_bytes(item.get("size"))))
            for cell in (name, size):
                cell.setTextAlignment(
                    Qt.AlignmentFlag.AlignRight
                    | Qt.AlignmentFlag.AlignAbsolute
                    | Qt.AlignmentFlag.AlignVCenter
                )
            self.file_table.setItem(row, 0, check)
            self.file_table.setItem(row, 1, name)
            self.file_table.setItem(row, 2, size)
        self.file_table.blockSignals(False)
        self.file_table.resizeColumnsToContents()
        self.file_table.setColumnWidth(0, 58)
        self.file_table.setVisible(bool(files))
        self.file_table.setFixedHeight(max(120, min(260, 40 + 36 * len(files))))
        if files:
            self.scan_status.setText(
                "فقط فایل‌های تیک‌خورده تحلیل می‌شوند؛ نتیجه در یک گزارش مقایسه‌ای جمع می‌شود."
            )
            self.scan_status.setObjectName("Caption")
        else:
            self.scan_status.setText("در این پوشه فایل پشتیبانی‌شده‌ای پیدا نشد.")
            self.scan_status.setObjectName("WarningText")
        self._update_selection()

    def selected_paths(self) -> list[Path]:
        paths: list[Path] = []
        for row in range(self.file_table.rowCount()):
            check = self.file_table.item(row, 0)
            if check is None or check.checkState() != Qt.CheckState.Checked:
                continue
            value = check.data(Qt.ItemDataRole.UserRole)
            if value:
                paths.append(Path(str(value)))
        return paths

    def _selection_changed(self, *_: Any) -> None:
        self._update_selection()

    def _update_selection(self) -> None:
        total = self.file_table.rowCount()
        selected = len(self.selected_paths())
        self.count_label.setText(
            f"فایل‌های قابل پردازش: {format_number(total)}"
            f" • انتخاب‌شده: {format_number(selected)}"
        )
        self.selection_changed.emit()
        self.scan_status.style().unpolish(self.scan_status)
        self.scan_status.style().polish(self.scan_status)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.folder_icon.setPixmap(icons.pixmap("folder", palette.accent, 26))
        self.choose_button.apply_palette(palette)
        self.refresh_button.apply_palette(palette)


class FolderProgressCard(Card):
    def __init__(self, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.add(
            SectionHeader(
                "پیشرفت تحلیل پوشه",
                "فایل‌ها و اسناد منتخب یکی‌یکی استخراج و تحلیل می‌شوند؛ گزارش ترکیبی در پایان ساخته می‌شود.",
            )
        )
        self.busy = BusyBar()
        self.current = QLabel("در انتظار شروع…")
        self.current.setObjectName("CardTitle")
        self.detail = make_label("", "Caption", wrap=True)
        self.add(self.busy)
        self.add(self.current)
        self.add(self.detail)
        self.setVisible(False)

    def begin(self) -> None:
        self.busy.start()
        self.current.setText("در حال آماده‌سازی…")
        self.detail.setText("")
        self.setVisible(True)

    def on_progress(self, message: str) -> None:
        text = str(message).strip()
        if not text or "\n" in text:
            return
        if text.startswith("تحلیل فایل:"):
            name = text.split(":", 1)[1].strip()
            self.current.setText(f"در حال تحلیل: {name}")
            self.detail.setText("ساختار، جدول‌ها و شاخص‌های این فایل بررسی می‌شود.")
            return
        self.detail.setText(elide(text, 180))

    def finish(self, message: str = "تحلیل پوشه کامل شد.") -> None:
        self.busy.stop()
        self.current.setText(message)
        self.detail.setText("گزارش مقایسه‌ای ساخته شد و برای مشاهده در بخش نتیجه آماده است.")


class FolderAnalysisPage(Page):
    browse_requested = Signal()
    refresh_requested = Signal()
    analyze_requested = Signal(str)
    stop_requested = Signal()
    folder_dropped = Signal(str)
    followup_requested = Signal(object)

    def __init__(self, state: Any, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self._palette = palette
        self._scan_worker: FolderScanWorker | None = None
        self._scan_folder: Path | None = None
        self._scan_pending = False
        self.folder_card = FolderCard(palette)
        self.folder_card.browse_requested.connect(self.browse_requested.emit)
        self.folder_card.refresh_requested.connect(self.refresh_requested.emit)
        self.folder_card.folder_dropped.connect(self.folder_dropped.emit)
        self.folder_card.selection_changed.connect(self._update_folder_hint)
        self.add(self.folder_card)

        self.composer = RequestComposer(palette)
        self.composer.start_requested.connect(self.analyze_requested.emit)
        self.composer.stop_requested.connect(self.stop_requested.emit)
        self.add(self.composer)

        self.progress_card = FolderProgressCard(palette)
        self.add(self.progress_card)
        self.error_card = ErrorCard(palette)
        self.add(self.error_card)
        self.results = FolderResultsView()
        self.results.apply_palette(palette)
        self.results.followup_requested.connect(self.followup_requested.emit)
        self.add(self.results)

        state.folder_changed.connect(self._refresh_folder)
        state.source_changed.connect(self._refresh_folder)
        self._refresh_folder()

    def _output_note(self) -> str:
        settings = getattr(self.state, "settings", None)
        output = getattr(settings, "output_dir", None)
        root = getattr(settings, "project_root", None)
        if output is None:
            return ""
        try:
            shown = Path(output).relative_to(root) if root is not None else Path(output)
        except ValueError:
            shown = Path(output)
        return f"خروجی: {ltr(str(shown).replace(chr(92), '/'))} • فایل‌های ورودی تغییر نمی‌کنند."

    def _apply_output_note(self) -> None:
        text = self._output_note()
        if not text:
            return
        settings = self.state.settings
        self.composer.set_output_note(text, str(settings.output_dir))

    # ------------------------------------------------------------------ #
    def _refresh_folder(self) -> None:
        self.folder_card.set_folder(self.state.folder)
        self._apply_output_note()
        result = self.state.last_folder_result
        folder = self.state.folder
        matches = bool(
            folder is not None
            and result is not None
            and Path(str(getattr(result, "folder", ""))).resolve() == folder
        )
        if not matches:
            self.results.setVisible(False)
        if folder is not None:
            self.scan()
        else:
            self.composer.set_hint("ابتدا یک پوشه انتخاب کنید.", "warning")

    def set_folder(self, path: Path | None) -> None:
        self.folder_card.set_folder(path)
        if path is not None:
            self.scan()

    def selected_paths(self) -> list[Path]:
        return self.folder_card.selected_paths()

    def _update_folder_hint(self) -> None:
        total = self.folder_card.file_table.rowCount()
        selected = len(self.selected_paths())
        if total and not selected:
            self.composer.set_hint("حداقل یک فایل را برای تحلیل تیک بزنید.", "warning")
        elif total:
            self.composer.set_hint("فایل‌های انتخاب‌شده آماده تحلیل‌اند.", "muted")

    def scan(self) -> None:
        folder = self.state.folder
        if folder is None:
            return
        if self._scan_worker is not None:
            if self._scan_folder == folder:
                return
            self._scan_pending = True
            return
        self._scan_pending = False
        self._scan_folder = folder
        self.folder_card.scan_status.setText("در حال خواندن فهرست فایل‌ها…")
        worker = FolderScanWorker(self.state.settings, folder, recursive=self.folder_card.recursive_check.isChecked())
        worker.completed.connect(lambda files, f=folder: self._on_scanned(f, files))
        worker.failed.connect(lambda message, f=folder: self._on_scan_failed(f, message))
        worker.finished.connect(lambda w=worker: self._scan_finished(w))
        self._scan_worker = worker
        worker.start()

    def _scan_finished(self, worker: FolderScanWorker) -> None:
        if self._scan_worker is worker:
            self._scan_worker = None
            self._scan_folder = None
        worker.deleteLater()
        if self._scan_pending:
            self._scan_pending = False
            self.scan()

    def _on_scanned(self, folder: Path, files: Any) -> None:
        if self.state.folder != folder:
            return
        self.folder_card.set_files(list(files))
        if not files:
            self.composer.set_hint("در این پوشه فایل قابل پردازشی نیست.", "warning")
        else:
            self._update_folder_hint()

    def _on_scan_failed(self, folder: Path, message: str) -> None:
        if self.state.folder != folder:
            return
        self.folder_card.scan_status.setText(message)

    # ------------------------------------------------------------------ #
    def begin_run(self) -> None:
        self.error_card.setVisible(False)
        self.results.setVisible(False)
        self.composer.set_running(True)
        self.progress_card.begin()

    def on_progress(self, message: str) -> None:
        self.progress_card.on_progress(message)

    def on_stage(self, event: StageEvent) -> None:
        self.progress_card.on_progress(f"تحلیل فایل: {event.name}" if event.state == "running" else "")

    def show_result(self, result: Any) -> None:
        self.results.apply_palette(self._palette)
        self.results.set_result(result, self._palette)

    def finish_success(self, result: Any) -> None:
        self.composer.input.clearFocus()
        self.composer.start_button.clearFocus()
        self.set_running(False)
        self.setFocus()
        self.progress_card.finish()
        self.progress_card.setVisible(False)
        self.show_result(result)
        status = str(getattr(result, "status", ""))
        if status == "failed":
            self.composer.set_hint(
                "تحلیل پوشه انجام شد اما دروازه کیفیت گزارش را معتبر نمی‌داند.",
                "danger",
            )
        elif status == "completed":
            self.composer.set_hint("تحلیل پوشه کامل شد.", "success")
        else:
            self.composer.set_hint("تحلیل پوشه با هشدار کامل شد؛ جزئیات را بررسی کنید.", "warning")

    def finish_error(self, message: str) -> None:
        self.composer.set_running(False)
        self.progress_card.finish()
        self.error_card.show_error(message)
        self.composer.set_hint("تحلیل پوشه انجام نشد.", "danger")

    def set_running(self, running: bool) -> None:
        self.composer.set_running(running)

    def shutdown(self) -> None:
        worker = self._scan_worker
        if worker is not None and worker.isRunning():
            worker.wait(2000)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.folder_card.apply_palette(palette)
        self.composer.apply_palette(palette)
        self.error_card.apply_palette(palette)
        self.results.apply_palette(palette)
