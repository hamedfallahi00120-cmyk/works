"""Source selection and the Persian request composer shared by the pages."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    CallbackButton,
    Card,
    IconButton,
    PromptChips,
    RequestInput,
    SectionHeader,
    make_label,
)
from local_office_ai.ui.desktop.formatting import align_table_header, format_number, ltr
from local_office_ai.ui.desktop.theme import Palette

QUICK_PROMPTS = (
    "این اطلاعات را بررسی کن، تغییرات مهم را پیدا کن و گزارش مدیریتی تهیه کن.",
    "بیشترین و کمترین ماه را با عدد مشخص کن.",
    "کیفیت داده و موارد نیازمند بررسی را تحلیل کن.",
    "دوره‌های موجود را مقایسه کن و تفاوت اصلی را توضیح بده.",
)

SUPPORTED_FORMATS = "XLSX · XLSM · CSV · DOCX · PPTX · PDF · TXT · تصویر (OCR)"

_FILE_TYPES = (
    ((".xlsx", ".xlsm"), "Excel"),
    ((".csv",), "CSV"),
    ((".docx",), "Word"),
    ((".pptx",), "PowerPoint"),
    ((".pdf",), "PDF"),
    ((".txt",), "متن"),
    ((".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"), "تصویر"),
)


def file_type_label(path: Path | str) -> str:
    target = Path(path)
    suffix = target.suffix.lower()
    for suffixes, label in _FILE_TYPES:
        if suffix in suffixes:
            return label
    return "پوشه" if target.is_dir() else "سایر"


class FileDropZone(QFrame):
    browse_requested = Signal()
    file_dropped = Signal(str)
    files_dropped = Signal(list)

    def __init__(self, parent: QWidget | None = None, *, compact: bool = False) -> None:
        super().__init__(parent)
        self.setObjectName("CardSunken")
        self.setMinimumHeight(58 if compact else 78)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAcceptDrops(True)
        self.setToolTip("فایل‌ها یا پوشه‌ها را روی این کادر رها کنید")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 10, 16, 10)
        layout.setSpacing(12)
        self.icon_label = QLabel()
        self.icon_label.setPixmap(icons.pixmap("file", "#0F766E", 24))
        texts = QVBoxLayout()
        texts.setSpacing(1)
        title = QLabel("فایل یا اسناد خود را اینجا رها کنید")
        title.setObjectName("CardTitle")
        subtitle = QLabel(f"فرمت‌های پذیرفته‌شده: {SUPPORTED_FORMATS} • فایل ورودی تغییر نمی‌کند")
        subtitle.setObjectName("Caption")
        subtitle.setWordWrap(True)
        texts.addWidget(title)
        texts.addWidget(subtitle)
        layout.addWidget(self.icon_label)
        layout.addLayout(texts, 1)

    def apply_palette(self, palette: Palette) -> None:
        self.icon_label.setPixmap(icons.pixmap("file", palette.accent, 24))

    def mousePressEvent(self, event: Any) -> None:  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self.browse_requested.emit()
        super().mousePressEvent(event)

    def dragEnterEvent(self, event: Any) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: Any) -> None:  # noqa: N802
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if not paths:
            return
        if len(paths) == 1:
            self.file_dropped.emit(paths[0])
        else:
            self.files_dropped.emit(paths)
        event.acceptProposedAction()


class SourceListCard(Card):
    """The single place where inputs are chosen and reviewed.

    One compact table lists every chosen file with its type and status, so the
    user always knows what will be sent to the engine.
    """

    browse_requested = Signal()
    browse_folder_requested = Signal()
    remove_requested = Signal(str)
    file_dropped = Signal(str)
    files_dropped = Signal(list)

    ROW_HEIGHT = 38
    MAX_ROWS = 5

    def __init__(
        self,
        palette: Palette,
        parent: QWidget | None = None,
        *,
        drop_zone: bool = True,
        subtitle: str = (
            "یک یا چند فایل از مسیرهای مختلف یا یک پوشه را انتخاب کنید؛ "
            "همه در یک درخواست تحلیل می‌شوند."
        ),
    ) -> None:
        super().__init__(parent, margins=(20, 18, 20, 16), spacing=12)
        self._palette = palette
        self._rows: list[tuple[str, str, str, str]] = []
        self.add(SectionHeader("منابع داده", subtitle))

        if drop_zone:
            self.drop_zone: FileDropZone | None = FileDropZone(compact=True)
            self.drop_zone.browse_requested.connect(self.browse_requested.emit)
            self.drop_zone.file_dropped.connect(self.file_dropped.emit)
            self.drop_zone.files_dropped.connect(self.files_dropped.emit)
            self.add(self.drop_zone)
        else:
            self.drop_zone = None

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.browse_button = CallbackButton(
            "افزودن فایل", self.browse_requested.emit, object_name="Primary", icon_name="file"
        )
        self.browse_folder_button = CallbackButton(
            "افزودن پوشه", self.browse_folder_requested.emit, object_name="Ghost", icon_name="folder"
        )
        self.clear_button = CallbackButton(
            "پاک کردن فهرست", self._clear, object_name="Quiet", icon_name="trash"
        )
        self.count_label = make_label("", "Caption")
        actions.addWidget(self.browse_button)
        actions.addWidget(self.browse_folder_button)
        actions.addWidget(self.clear_button)
        actions.addStretch(1)
        actions.addWidget(self.count_label)
        self.add_layout(actions)

        self.table = QTableWidget(0, 4)
        self.table.setObjectName("SourceTable")
        self.table.setHorizontalHeaderLabels(["نام فایل", "نوع", "وضعیت", ""])
        align_table_header(self.table)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(self.ROW_HEIGHT)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        header = self.table.horizontalHeader()
        header.setHighlightSections(False)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column in (1, 2, 3):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setVisible(False)
        self.add(self.table)

        self.empty = make_label(
            "هنوز منبعی انتخاب نشده است. یک فایل یا یک پوشه اضافه کنید.",
            "Muted",
            wrap=True,
        )
        self.add(self.empty)
        self._refresh()

    # ------------------------------------------------------------------ #
    def set_file(self, path: Path | None) -> None:
        self.set_files([path] if path is not None else [])

    def set_files(self, paths: list[Path] | tuple[Path, ...]) -> None:
        self.set_sources(list(paths), None)

    def set_sources(self, files: list[Path] | tuple[Path, ...], folder: Path | None = None) -> None:
        rows: list[tuple[str, str, str, str]] = []
        if folder is not None:
            rows.append((Path(folder).name, "پوشه", "آماده تحلیل پوشه", str(folder)))
        for item in files:
            path = Path(item)
            rows.append((path.name, file_type_label(path), "آماده", str(path)))
        if rows == self._rows:
            return
        self._rows = rows
        self._refresh()

    def _refresh(self) -> None:
        rows = self._rows
        self.table.blockSignals(True)
        self.table.clearContents()
        self.table.setRowCount(len(rows))
        for index, (name, kind, status, full_path) in enumerate(rows):
            name_item = QTableWidgetItem(ltr(name))
            name_item.setToolTip(full_path)
            name_item.setTextAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter
            )
            kind_item = QTableWidgetItem(kind)
            kind_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            status_item = QTableWidgetItem(status)
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            status_item.setForeground(QColor(self._palette.success))
            for column, cell in enumerate((name_item, kind_item, status_item)):
                self.table.setItem(index, column, cell)
            remove = IconButton("trash", "حذف این منبع از فهرست", size=28, icon_size=15)
            remove.apply_palette(self._palette)
            remove.clicked.connect(lambda _checked=False, path=full_path: self.remove_requested.emit(path))
            holder = QWidget()
            holder_layout = QHBoxLayout(holder)
            holder_layout.setContentsMargins(4, 0, 4, 0)
            holder_layout.addWidget(remove)
            self.table.setCellWidget(index, 3, holder)
        self.table.blockSignals(False)

        visible = bool(rows)
        self.table.setVisible(visible)
        self.empty.setVisible(not visible)
        self.clear_button.setVisible(visible)
        if visible:
            header_height = max(30, self.table.horizontalHeader().sizeHint().height())
            shown = min(len(rows), self.MAX_ROWS)
            self.table.setFixedHeight(header_height + shown * self.ROW_HEIGHT + 4)
            kinds = {row[1] for row in rows}
            self.count_label.setText(f"{format_number(len(rows))} منبع • {' • '.join(sorted(kinds))}")
        else:
            self.count_label.setText("بدون منبع")

    def _clear(self) -> None:
        self.remove_requested.emit("")

    # ------------------------------------------------------------------ #
    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        if self.drop_zone is not None:
            self.drop_zone.apply_palette(palette)
        self.browse_button.apply_palette(palette)
        self.browse_folder_button.apply_palette(palette)
        self.clear_button.apply_palette(palette)
        for index in range(self.table.rowCount()):
            holder = self.table.cellWidget(index, 3)
            if holder is not None:
                button = holder.findChild(IconButton)
                if button is not None:
                    button.apply_palette(palette)


class RequestComposer(Card):
    start_requested = Signal(str)
    stop_requested = Signal()

    def __init__(self, palette: Palette, parent: QWidget | None = None, *, compact: bool = False) -> None:
        super().__init__(parent, margins=(20, 18, 20, 16), spacing=12)
        self._palette = palette
        self.compact = compact
        heading = QHBoxLayout()
        heading.setSpacing(10)
        heading.addWidget(
            SectionHeader(
                "درخواست شما",
                "به فارسی بنویسید؛ نوع تحلیل و شاخص‌های هدف از همین متن خوانده می‌شود.",
            ),
            1,
        )
        self.stop_button = CallbackButton("توقف", self.stop_requested.emit, object_name="Ghost", icon_name="stop")
        self.stop_button.setVisible(False)
        self.start_button = CallbackButton(
            "شروع تحلیل", self._emit_start, object_name="Primary", icon_name="play"
        )
        heading.addWidget(self.stop_button, 0, Qt.AlignmentFlag.AlignBottom)
        heading.addWidget(self.start_button, 0, Qt.AlignmentFlag.AlignBottom)
        self.add_layout(heading)

        self.input = RequestInput(
            "مثال: این اطلاعات را بررسی کن، تغییرات مهم را پیدا کن و گزارش مدیریتی تهیه کن."
        )
        self.input.setMinimumHeight(88 if compact else 104)
        self.add(self.input)
        self.chips = PromptChips(QUICK_PROMPTS)
        self.chips.chosen.connect(self._append_prompt)
        self.add(self.chips)

        self.hint = make_label(
            "بدون درخواست، تحلیل عمومی روی ساختار فایل انجام می‌شود.", "Caption", wrap=True
        )
        self.add(self.hint)
        self.output_note = make_label("", "Caption", wrap=True)
        self.output_note.setVisible(False)
        self.add(self.output_note)

    # ------------------------------------------------------------------ #
    def _emit_start(self) -> None:
        self.start_requested.emit(self.request_text())

    def _append_prompt(self, prompt: str) -> None:
        current = self.input.toPlainText().strip()
        self.input.setPlainText(prompt if not current else f"{current} {prompt}")
        self.input.setFocus()

    def request_text(self) -> str:
        return self.input.toPlainText().strip()

    def set_request(self, text: str) -> None:
        if text and not self.input.toPlainText().strip():
            self.input.setPlainText(text)

    def set_running(self, running: bool) -> None:
        self.start_button.setEnabled(not running)
        self.stop_button.setVisible(running)
        self.start_button.setText("در حال پردازش…" if running else "شروع تحلیل")

    def set_output_note(self, text: str, tooltip: str = "") -> None:
        self.output_note.setText(text)
        self.output_note.setToolTip(tooltip)
        self.output_note.setVisible(bool(text))

    def set_hint(self, text: str, tone: str = "muted") -> None:
        self.hint.setText(text)
        self.hint.setObjectName(
            {"danger": "ErrorText", "warning": "WarningText", "success": "SuccessText"}.get(tone, "Caption")
        )
        self.hint.style().unpolish(self.hint)
        self.hint.style().polish(self.hint)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.start_button.apply_palette(palette)
        self.stop_button.apply_palette(palette)
