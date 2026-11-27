"""Read-only preview of a generated report workbook.

The preview never opens the source workbook and never edits anything: it reads
at most a few dozen rows per sheet with ``openpyxl`` in read-only mode.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.formatting import format_number


@dataclass
class SheetPreview:
    name: str
    rows: list[list[str]]
    max_row: int
    max_column: int
    truncated: bool


def read_workbook_preview(
    path: Path,
    *,
    max_rows: int = 60,
    max_cols: int = 14,
    max_sheets: int = 12,
) -> list[SheetPreview]:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        previews: list[SheetPreview] = []
        for worksheet in workbook.worksheets[:max_sheets]:
            rows: list[list[str]] = []
            for row_index, row in enumerate(
                worksheet.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols, values_only=True),
                start=1,
            ):
                rows.append([_cell_text(value) for value in row])
                if row_index >= max_rows:
                    break
            while rows and not any(cell.strip() for cell in rows[-1]):
                rows.pop()
            total_rows = int(worksheet.max_row or len(rows) or 0)
            total_cols = int(worksheet.max_column or 0)
            previews.append(
                SheetPreview(
                    name=worksheet.title,
                    rows=rows,
                    max_row=total_rows,
                    max_column=total_cols,
                    truncated=total_rows > len(rows) or total_cols > max_cols,
                )
            )
        return previews
    finally:
        workbook.close()


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "بله" if value else "خیر"
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, Decimal):
        return format_number(value)
    return str(value)


def open_path(path: str | Path) -> bool:
    target = Path(path)
    if not target.exists():
        return False
    return QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))


class ReportPreviewDialog(QDialog):
    def __init__(self, path: str | Path, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.path = Path(path)
        self.setWindowTitle(f"پیش‌نمایش گزارش — {self.path.name}")
        self.setMinimumSize(900, 620)
        self.resize(1040, 700)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        title = QLabel(self.path.name)
        title.setObjectName("SectionTitle")
        layout.addWidget(title)
        self.summary = QLabel("در حال خواندن فایل…")
        self.summary.setObjectName("Caption")
        layout.addWidget(self.summary)

        body = QHBoxLayout()
        body.setSpacing(12)
        self.sheet_list = QListWidget()
        self.sheet_list.setFixedWidth(220)
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        body.addWidget(self.sheet_list, 0)
        body.addWidget(self.table, 1)
        layout.addLayout(body, 1)

        footer = QHBoxLayout()
        self.note = QLabel("")
        self.note.setObjectName("Caption")
        footer.addWidget(self.note, 1)
        open_button = QPushButton("بازکردن در Excel")
        open_button.setObjectName("Ghost")
        open_button.setIcon(icons.icon("external", "#5C7288", 16))
        open_button.clicked.connect(self._open_excel)
        close_button = QPushButton("بستن")
        close_button.setObjectName("Primary")
        close_button.clicked.connect(self.accept)
        footer.addWidget(open_button)
        footer.addWidget(close_button)
        layout.addLayout(footer)

        self._previews: list[SheetPreview] = []
        self.sheet_list.currentRowChanged.connect(self._show_sheet)
        self._load()

    # ------------------------------------------------------------------ #
    def _load(self) -> None:
        try:
            self._previews = read_workbook_preview(self.path)
        except Exception as exc:  # noqa: BLE001 - the message is the point here
            self.summary.setText(f"خواندن گزارش ممکن نشد: {type(exc).__name__}: {exc}")
            self.summary.setObjectName("ErrorText")
            return
        if not self._previews:
            self.summary.setText("این فایل برگه‌ای ندارد.")
            return
        self.summary.setText(
            f"{len(self._previews)} برگه | نمایش حداکثر ۶۰ ردیف و ۱۴ ستون از هر برگه"
        )
        for preview in self._previews:
            item = QListWidgetItem(preview.name)
            item.setToolTip(f"{preview.max_row} ردیف × {preview.max_column} ستون")
            self.sheet_list.addItem(item)
        self.sheet_list.setCurrentRow(0)

    def _show_sheet(self, index: int) -> None:
        if not (0 <= index < len(self._previews)):
            return
        preview = self._previews[index]
        from openpyxl.utils import get_column_letter

        column_count = max((len(row) for row in preview.rows), default=0)
        self.table.clear()
        self.table.setRowCount(len(preview.rows))
        self.table.setColumnCount(max(column_count, 1))
        self.table.setHorizontalHeaderLabels(
            [get_column_letter(index + 1) for index in range(max(column_count, 1))]
        )
        self.table.setVerticalHeaderLabels(
            [str(row_number) for row_number in range(1, len(preview.rows) + 1)]
        )
        for row_index, row in enumerate(preview.rows):
            for column_index in range(column_count):
                text = row[column_index] if column_index < len(row) else ""
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row_index, column_index, item)
        for column_index in range(column_count):
            self.table.setColumnWidth(column_index, 120)
        self.note.setText(
            f"اندازه واقعی برگه: {preview.max_row} ردیف × {preview.max_column} ستون"
            + (" — پیش‌نمایش بریده شده است" if preview.truncated else "")
        )

    def _open_excel(self) -> None:
        if not open_path(self.path):
            self.note.setText("بازکردن فایل ممکن نشد؛ مسیر را بررسی کنید.")
