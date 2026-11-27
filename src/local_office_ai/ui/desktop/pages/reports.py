"""Previous reports: whatever really exists in the configured output folder."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    CallbackButton,
    Card,
    EmptyState,
    SectionHeader,
    make_label,
)
from local_office_ai.ui.desktop.formatting import (
    align_table_header,
    elide,
    format_bytes,
    format_timestamp,
    ltr,
)
from local_office_ai.ui.desktop.pages.base import Page
from local_office_ai.ui.desktop.preview import ReportPreviewDialog, open_path
from local_office_ai.ui.desktop.theme import Palette
from local_office_ai.ui.desktop.workers import ReportsScanWorker


class ReportsPage(Page):
    def __init__(self, state: Any, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self._palette = palette
        self._entries: list[dict[str, Any]] = []
        self._worker: ReportsScanWorker | None = None
        self._build()
        state.settings_changed.connect(self._refresh_path_label)

    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        card = Card()
        card.add(
            SectionHeader(
                "گزارش‌های قبلی",
                "فقط فایل‌های xlsx پوشه خروجی نمایش داده می‌شوند؛ چیزی حذف یا بازنویسی نمی‌شود.",
            )
        )
        toolbar = QHBoxLayout()
        self.path_label = make_label("", "Caption", wrap=True)
        self.refresh_button = CallbackButton("بازخوانی", self.refresh, object_name="Ghost", icon_name="refresh")
        self.folder_button = CallbackButton(
            "بازکردن پوشه خروجی", self._open_folder, object_name="Ghost", icon_name="folder-open"
        )
        toolbar.addWidget(self.path_label, 1)
        toolbar.addWidget(self.refresh_button)
        toolbar.addWidget(self.folder_button)
        card.add_layout(toolbar)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["نام گزارش", "تاریخ تغییر", "حجم"])
        align_table_header(self.table)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setMinimumHeight(320)
        self.table.doubleClicked.connect(lambda _index: self._preview())
        self.table.itemSelectionChanged.connect(self._selection_changed)
        card.add(self.table)

        self.empty = EmptyState(
            "history",
            "گزارشی پیدا نشد",
            "پس از اولین تحلیل، گزارش‌های ساخته‌شده در این پوشه فهرست می‌شوند.",
            self._palette,
        )
        self.empty.setVisible(False)
        card.add(self.empty)

        actions = QHBoxLayout()
        self.preview_button = CallbackButton("پیش‌نمایش", self._preview, object_name="Primary", icon_name="eye")
        self.open_button = CallbackButton("بازکردن در Excel", self._open_selected, object_name="Ghost", icon_name="external")
        self.status = make_label("", "Caption")
        actions.addWidget(self.preview_button)
        actions.addWidget(self.open_button)
        actions.addStretch(1)
        actions.addWidget(self.status)
        card.add_layout(actions)
        self.add(card)
        self.preview_button.setEnabled(False)
        self.open_button.setEnabled(False)
        self._refresh_path_label()

    def _refresh_path_label(self, *_: Any) -> None:
        path = str(self.state.settings.output_dir)
        self.path_label.setText("پوشه خروجی: " + elide(ltr(path), 86))
        self.path_label.setToolTip(path)

    # ------------------------------------------------------------------ #
    def refresh(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            return
        self.status.setText("در حال خواندن پوشه…")
        worker = ReportsScanWorker(self.state.settings.output_dir)
        worker.completed.connect(self._on_scanned)
        worker.failed.connect(lambda message: self.status.setText(message))
        worker.finished.connect(self._clear_worker)
        self._worker = worker
        worker.start()

    def _clear_worker(self) -> None:
        self._worker = None

    def _on_scanned(self, entries: Any) -> None:
        self._entries = list(entries)
        self.table.clearContents()
        self.table.setRowCount(len(self._entries))
        for row, item in enumerate(self._entries):
            name = QTableWidgetItem(ltr(item.get("name", "")))
            modified = QTableWidgetItem(ltr(format_timestamp(item.get("modified"))))
            size = QTableWidgetItem(ltr(format_bytes(item.get("size"))))
            for cell in (name, modified, size):
                cell.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row, 0, name)
            self.table.setItem(row, 1, modified)
            self.table.setItem(row, 2, size)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setVisible(bool(self._entries))
        self.empty.setVisible(not self._entries)
        self.status.setText(f"{len(self._entries)} گزارش در پوشه خروجی." if self._entries else "")
        self.preview_button.setEnabled(bool(self._entries))
        self.open_button.setEnabled(bool(self._entries))
        if self._entries:
            self.table.selectRow(0)

    def _selection_changed(self) -> None:
        enabled = bool(self.selected_entry())
        self.preview_button.setEnabled(enabled)
        self.open_button.setEnabled(enabled)

    def selected_entry(self) -> dict[str, Any] | None:
        row = self.table.currentRow()
        if 0 <= row < len(self._entries):
            return self._entries[row]
        return None

    def _preview(self) -> None:
        entry = self.selected_entry()
        if entry is None:
            return
        ReportPreviewDialog(str(entry["path"]), self).exec()

    def _open_selected(self) -> None:
        entry = self.selected_entry()
        if entry is not None:
            open_path(str(entry["path"]))

    def _open_folder(self) -> None:
        open_path(self.state.settings.output_dir)

    def shutdown(self) -> None:
        worker = self._worker
        if worker is not None and worker.isRunning():
            worker.wait(2000)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
