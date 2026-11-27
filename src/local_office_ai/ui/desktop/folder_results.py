"""Result panel for a folder analysis (many workbooks, one comparison)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QHeaderView,
    QAbstractItemView,
    QApplication,
    QHBoxLayout,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    Badge,
    CallbackButton,
    Card,
    IconButton,
    MetricTile,
    SectionHeader,
    horizontal_line,
    make_label,
)
from local_office_ai.ui.desktop.formatting import align_table_header, format_number, ltr
from local_office_ai.ui.desktop.preview import ReportPreviewDialog, open_path
from local_office_ai.ui.desktop.theme import Palette


class FolderResultsView(Card):
    followup_requested = Signal(object)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent, margins=(20, 18, 20, 18), spacing=14)
        self._result: Any = None
        self._path = ""
        self._palette: Palette | None = None
        self._buttons: list[IconButton] = []
        self._outcomes: list[Any] = []
        self._build()
        self.setVisible(False)

    def _build(self) -> None:
        header = QHBoxLayout()
        self.status_badge = Badge("تکمیل شد", "success")
        self.title = QLabel("نتیجه تحلیل منابع")
        self.title.setObjectName("CardTitle")
        header.addWidget(self.status_badge)
        header.addWidget(self.title, 1)
        self.preview_button = self._icon_button("eye", "پیش‌نمایش گزارش")
        self.open_button = self._icon_button("external", "بازکردن فایل Excel")
        self.folder_button = self._icon_button("folder-open", "بازکردن پوشه خروجی")
        self.copy_button = self._icon_button("copy", "کپی خلاصه")
        self.followup_button = self._icon_button("sparkles", "پرسش تکمیلی روی فایل انتخاب‌شده")
        self.followup_button.setEnabled(False)
        self.followup_button.setToolTip("پرسش تکمیلی روی منابع تحلیل‌شده")
        for button in (
            self.preview_button,
            self.open_button,
            self.folder_button,
            self.copy_button,
            self.followup_button,
        ):
            header.addWidget(button)
        self.body.addLayout(header)
        self.path_label = QLabel("")
        self.path_label.setObjectName("Caption")
        self.path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.body.addWidget(self.path_label)
        self.body.addWidget(horizontal_line())

        self.tiles_row = QHBoxLayout()
        self.tiles_row.setSpacing(10)
        self.tile_files = MetricTile("منابع بررسی‌شده", "—")
        self.tile_analyzed = MetricTile("تحلیل‌شده", "—")
        self.tile_metrics = MetricTile("شاخص عددی", "—")
        self.tile_period = MetricTile("دوره تحلیل", "—")
        self.tile_issues = MetricTile("یادداشت‌ها", "—")
        for tile in (self.tile_files, self.tile_analyzed, self.tile_metrics, self.tile_period, self.tile_issues):
            self.tiles_row.addWidget(tile)
        self.body.addLayout(self.tiles_row)

        self.body.addWidget(SectionHeader("منابع بررسی‌شده", "وضعیت و خلاصه هر فایل ورودی"))
        self.table = QTableWidget()
        self.table.setMinimumHeight(240)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.itemSelectionChanged.connect(self._sync_followup)
        self.body.addWidget(self.table)
        self.body.addWidget(
            make_label(
                "پس از تحلیل ترکیبی، می‌توانید در گفت‌وگوی هوشمند درباره همه منابع سؤال تکمیلی بپرسید؛ برای نتیجه Excel مستقل نیز یک ردیف را انتخاب کنید.",
                "Caption",
                wrap=True,
            )
        )

        self.warnings_title = SectionHeader("هشدارها")
        self.body.addWidget(self.warnings_title)
        self.warnings_box = QWidget()
        self.warnings_layout = QVBoxLayout(self.warnings_box)
        self.warnings_layout.setContentsMargins(0, 0, 0, 0)
        self.warnings_layout.setSpacing(4)
        self.body.addWidget(self.warnings_box)

        self.preview_button.clicked.connect(self._preview)
        self.open_button.clicked.connect(lambda: open_path(self._path) if self._path else None)
        self.folder_button.clicked.connect(
            lambda: open_path(Path(self._path).parent) if self._path else None
        )
        self.copy_button.clicked.connect(self._copy)
        self.followup_button.clicked.connect(self._followup)

    def _icon_button(self, name: str, tooltip: str) -> IconButton:
        button = IconButton(name, tooltip)
        self._buttons.append(button)
        return button

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        for button in self._buttons:
            button.apply_palette(palette)

    def set_result(self, result: Any, palette: Palette | None = None) -> None:
        self._result = result
        if palette is not None:
            self.apply_palette(palette)
        self._path = str(getattr(result, "output_path", "") or "")
        files = list(getattr(result, "files", []) or [])
        self._outcomes = files
        analyzed = sum(1 for item in files if getattr(item, "status", "") == "تحلیل شد")
        status = str(getattr(result, "status", ""))
        self.status_badge.set_status(
            "تکمیل شد" if status == "completed" else ("فایلی پیدا نشد" if status == "no_files" else "با هشدار"),
            "success" if status == "completed" else ("warning" if status == "completed_with_warnings" else "danger"),
        )
        self.title.setText(f"تحلیل منابع — {Path(str(getattr(result, 'folder', ''))).name or 'منابع انتخاب‌شده'}")
        self.path_label.setText(ltr(self._path) if self._path else "خروجی ساخته نشده است.")
        has_path = bool(self._path)
        self.preview_button.setEnabled(has_path)
        self.open_button.setEnabled(has_path)
        self.folder_button.setEnabled(has_path)
        self.tile_files.set_value(format_number(len(files)))
        self.tile_analyzed.set_value(format_number(analyzed), f"{format_number(len(files) - analyzed)} ناموفق")
        self.tile_metrics.set_value(format_number(len(getattr(result, "metrics", []) or [])))
        self.tile_period.set_value(str(getattr(result, "period", "") or "—"))
        self.tile_issues.set_value(format_number(len(getattr(result, "warnings", []) or [])))

        headers = ["فایل", "نوع", "وضعیت", "بخش", "ردیف", "ستون", "شاخص", "سلول خالی", "موارد", "خطا"]
        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        align_table_header(self.table)
        self.table.setRowCount(len(files))
        for row_index, item in enumerate(files):
            values = [
                ltr(item.name),
                str(getattr(item, "file_type", "") or "—"),
                item.status,
                format_number(item.sheets),
                format_number(item.rows),
                format_number(item.columns),
                format_number(item.metrics),
                item.missing_ratio or "—",
                format_number(item.issues),
                item.error,
            ]
            for column_index, value in enumerate(values):
                text = str(value).replace("\n", " ")
                cell = QTableWidgetItem(text)
                cell.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter)
                if len(text) > 60:
                    cell.setToolTip(text)
                self.table.setItem(row_index, column_index, cell)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column_index in range(len(headers)):
            if self.table.columnWidth(column_index) > 320:
                self.table.setColumnWidth(column_index, 320)

        while self.warnings_layout.count():
            item = self.warnings_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        warnings = list(getattr(result, "warnings", []) or [])
        self.warnings_title.setVisible(bool(warnings))
        for warning in warnings:
            label = QLabel("• " + str(warning))
            label.setWordWrap(True)
            label.setObjectName("WarningText")
            self.warnings_layout.addWidget(label)
        if not warnings:
            self.warnings_layout.addWidget(make_label("هشدار ثبت‌شده‌ای وجود ندارد.", "Muted"))
        self.table.clearSelection()
        self._sync_followup()
        self.setVisible(True)

    def _selected_row(self) -> int:
        model = self.table.selectionModel()
        rows = model.selectedRows() if model is not None else []
        return rows[0].row() if rows else -1

    def _sync_followup(self) -> None:
        row = self._selected_row()
        outcome = self._outcomes[row] if 0 <= row < len(self._outcomes) else None
        has_document_set = hasattr(self._result, "ai_mode") and hasattr(self._result, "statistics")
        self.followup_button.setEnabled(has_document_set or getattr(outcome, "analysis", None) is not None)

    def _followup(self) -> None:
        row = self._selected_row()
        outcome = self._outcomes[row] if 0 <= row < len(self._outcomes) else None
        analysis = getattr(outcome, "analysis", None)
        if analysis is not None:
            self.followup_requested.emit(analysis)
        elif hasattr(self._result, "ai_mode") and hasattr(self._result, "statistics"):
            self.followup_requested.emit(self._result)

    def _preview(self) -> None:
        if self._path and Path(self._path).exists():
            ReportPreviewDialog(self._path, self).exec()

    def _copy(self) -> None:
        result = self._result
        if result is None:
            return
        lines = [
            "تحلیل منابع: " + str(getattr(result, "folder", "")),
            "درخواست: " + str(getattr(result, "request", "") or "—"),
            "خروجی: " + self._path,
            "فایل‌ها:",
        ]
        for item in list(getattr(result, "files", []) or []):
            lines.append(f"- {item.name}: {item.status} | شاخص {item.metrics} | خطا: {item.error or '—'}")
        warnings = list(getattr(result, "warnings", []) or [])
        if warnings:
            lines.append("هشدارها:")
            lines.extend(f"- {item}" for item in warnings)
        QApplication.clipboard().setText("\n".join(lines))
