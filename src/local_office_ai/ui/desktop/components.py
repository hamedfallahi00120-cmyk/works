"""Reusable widgets that give the interface one consistent visual language."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Iterable

from PySide6.QtCore import QPoint, QRect, QRectF, QSize, Qt, Signal
from PySide6.QtGui import (
    QColor,
    QFont,
    QFontMetrics,
    QLinearGradient,
    QMouseEvent,
    QPainter,
    QPaintEvent,
    QPen,
)
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLayout,
    QLayoutItem,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.formatting import format_duration, format_number
from local_office_ai.ui.desktop.theme import Palette, palette_for


def repolish(widget: QWidget) -> None:
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


class Card(QFrame):
    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        variant: str = "card",
        margins: tuple[int, int, int, int] = (20, 18, 20, 18),
        spacing: int = 12,
        shadow: bool = False,
    ) -> None:
        super().__init__(parent)
        name = {"card": "Card", "alt": "CardAlt", "sunken": "CardSunken", "hero": "Hero"}.get(
            variant, "Card"
        )
        self.setObjectName(name)
        self.body = QVBoxLayout(self)
        self.body.setContentsMargins(*margins)
        self.body.setSpacing(spacing)
        # ``shadow`` is kept for call compatibility only: a QGraphicsDropShadowEffect
        # puts the whole card into an offscreen surface and Qt then paints its child
        # widgets blank (dark empty boxes) after a resize or repaint.  Depth is
        # expressed with borders and surface colours instead.
        self.setProperty("depth", "raised" if shadow else "flat")

    def add(self, widget: QWidget, stretch: int = 0) -> QWidget:
        self.body.addWidget(widget, stretch)
        return widget

    def add_layout(self, layout: Any, stretch: int = 0) -> None:
        self.body.addLayout(layout, stretch)

    def add_stretch(self) -> None:
        self.body.addStretch(1)


class Badge(QLabel):
    def __init__(self, text: str = "", tone: str = "muted", parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setObjectName("Badge")
        self.setProperty("tone", tone)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)

    def set_status(self, text: str, tone: str = "muted") -> None:
        self.setText(text)
        self.setProperty("tone", tone)
        repolish(self)
        self.adjustSize()
        self.updateGeometry()


class Chip(QLabel):
    def __init__(self, text: str = "", tone: str = "muted", parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setObjectName("Chip")
        self.setProperty("tone", tone)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)

    def set_status(self, text: str, tone: str = "muted") -> None:
        self.setText(text)
        self.setProperty("tone", tone)
        repolish(self)
        self.adjustSize()
        self.updateGeometry()


class IconButton(QPushButton):
    def __init__(
        self,
        icon_name: str,
        tooltip: str = "",
        *,
        size: int = 36,
        icon_size: int = 18,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._icon_name = icon_name
        self.setObjectName("IconButton")
        self.setFixedSize(size, size)
        self.setIconSize(QSize(icon_size, icon_size))
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._palette = palette_for("light")
        if tooltip:
            self.setToolTip(tooltip)

    def set_icon_name(self, name: str) -> None:
        self._icon_name = name
        self.apply_palette(self._palette)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.setIcon(icons.icon(self._icon_name, palette.muted, self.iconSize().width()))


class NavigationButton(QPushButton):
    """Sidebar row: icon at the reading start, label beside it, full-row hit area.

    The icon and the caption are real child widgets so the gap, the alignment and
    the colour of the checked/hover state stay identical at every window size.
    """

    ICON_SIZE = 18

    def __init__(self, text: str, icon_name: str, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._icon_name = icon_name
        self._text = text
        self._palette = palette
        self._collapsed = False
        self.setObjectName("Nav")
        self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(40)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setToolTip(text)

        self._row = QHBoxLayout(self)
        self._row.setContentsMargins(12, 0, 12, 0)
        self._row.setSpacing(10)
        self.icon_label = QLabel()
        self.icon_label.setFixedSize(self.ICON_SIZE, self.ICON_SIZE)
        self.icon_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.text_label = QLabel(text)
        self.text_label.setObjectName("NavLabel")
        self.text_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self._row.addWidget(self.icon_label, 0, Qt.AlignmentFlag.AlignVCenter)
        self._row.addWidget(self.text_label, 1)
        self._refresh_style()

    # -- state ---------------------------------------------------------- #
    def set_collapsed(self, collapsed: bool) -> None:
        self._collapsed = bool(collapsed)
        self.text_label.setVisible(not self._collapsed)
        if self._collapsed:
            self._row.setContentsMargins(0, 0, 0, 0)
            self._row.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        else:
            self._row.setContentsMargins(12, 0, 12, 0)
            self._row.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.updateGeometry()

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self._refresh_style()

    def setChecked(self, checked: bool) -> None:  # noqa: N802 - Qt naming
        super().setChecked(checked)
        self._refresh_style()

    def enterEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        super().enterEvent(event)
        self._refresh_style(hovered=True)

    def leaveEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        super().leaveEvent(event)
        self._refresh_style()

    def _refresh_style(self, hovered: bool = False) -> None:
        palette = self._palette
        active = self.isChecked()
        color = "#FFFFFF" if active else (palette.sidebar_text if hovered else palette.sidebar_muted)
        self.icon_label.setPixmap(icons.pixmap(self._icon_name, color, self.ICON_SIZE))
        self.text_label.setStyleSheet(
            f"color: {color}; font-weight: {'700' if active else '500'};"
        )


class SectionHeader(QWidget):
    def __init__(self, title: str, subtitle: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(3)
        self.title = QLabel(title)
        self.title.setObjectName("SectionTitle")
        layout.addWidget(self.title)
        self.subtitle = QLabel(subtitle)
        self.subtitle.setObjectName("Muted")
        self.subtitle.setWordWrap(True)
        self.subtitle.setVisible(bool(subtitle))
        layout.addWidget(self.subtitle)


class MetricTile(QFrame):
    def __init__(self, label: str, value: str = "—", hint: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("CardAlt")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(4)
        self.label = QLabel(label)
        self.label.setObjectName("StatLabel")
        self.label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.value = QLabel(value)
        self.value.setObjectName("StatValue")
        self.value.setWordWrap(True)
        self.value.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.hint = QLabel(hint)
        self.hint.setObjectName("Caption")
        self.hint.setWordWrap(True)
        self.hint.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.hint.setVisible(bool(hint))
        layout.addWidget(self.label)
        layout.addWidget(self.value)
        layout.addWidget(self.hint)

    def set_value(self, value: str, hint: str = "") -> None:
        self.value.setText(value)
        self.hint.setText(hint)
        self.hint.setVisible(bool(hint))


class EmptyState(QWidget):
    def __init__(
        self,
        icon_name: str,
        title: str,
        text: str,
        palette: Palette | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._palette = palette or palette_for("light")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 26, 24, 26)
        layout.setSpacing(8)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_label = QLabel()
        self.icon_label.setPixmap(icons.pixmap(icon_name, self._palette.border_strong, 46))
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title = QLabel(title)
        self.title.setObjectName("CardTitle")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text = QLabel(text)
        self.text.setObjectName("Muted")
        self.text.setWordWrap(True)
        self.text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label, 0, Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self.title)
        layout.addWidget(self.text)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.icon_label.setPixmap(icons.pixmap("info", palette.border_strong, 46))


class BusyBar(QProgressBar):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Busy")
        self.setRange(0, 0)
        self.setTextVisible(False)
        self.setFixedHeight(4)
        self.hide()

    def start(self) -> None:
        self.show()
        self.setRange(0, 0)

    def stop(self) -> None:
        self.hide()
        self.setRange(0, 1)


@dataclass
class StepRow:
    dot: QLabel
    name: QLabel
    duration: QLabel
    container: QWidget


class StepList(QWidget):
    """Shows the stages the engine actually reports, never a guessed percent."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows: dict[str, StepRow] = {}
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(8)
        self.empty = QLabel("هنوز مرحله‌ای از موتور گزارش نشده است.")
        self.empty.setObjectName("Caption")
        self._layout.addWidget(self.empty)
        self._layout.addStretch(1)

    def reset(self) -> None:
        for row in self._rows.values():
            row.container.setParent(None)
        self._rows.clear()
        self.empty.show()

    def start(self, name: str) -> None:
        self.empty.hide()
        row = self._rows.get(name)
        if row is None:
            container = QWidget()
            line = QHBoxLayout(container)
            line.setContentsMargins(0, 0, 0, 0)
            line.setSpacing(10)
            dot = QLabel("…")
            dot.setObjectName("StepDot")
            dot.setFixedSize(22, 22)
            dot.setAlignment(Qt.AlignmentFlag.AlignCenter)
            title = QLabel(name)
            title.setObjectName("StepName")
            duration = QLabel("در حال اجرا")
            duration.setObjectName("StepDuration")
            line.addWidget(dot)
            line.addWidget(title, 1)
            line.addWidget(duration)
            row = StepRow(dot, title, duration, container)
            self._rows[name] = row
            self._layout.insertWidget(self._layout.count() - 1, container)
        row.dot.setText("…")
        row.dot.setProperty("state", "running")
        repolish(row.dot)
        row.name.setProperty("state", "running")
        repolish(row.name)
        row.duration.setText("در حال اجرا")
        row.duration.setProperty("state", "")
        repolish(row.duration)

    def finish(self, name: str, seconds: float = 0.0) -> None:
        row = self._rows.get(name)
        if row is None:
            self.start(name)
            row = self._rows[name]
        row.dot.setText("✓")
        row.dot.setProperty("state", "done")
        repolish(row.dot)
        row.name.setProperty("state", "done")
        repolish(row.name)
        row.duration.setText(format_duration(seconds) if seconds else "پایان یافت")
        row.duration.setProperty("state", "")
        repolish(row.duration)

    def fail(self, name: str, error: str = "") -> None:
        row = self._rows.get(name)
        if row is None:
            self.start(name)
            row = self._rows[name]
        row.dot.setText("!")
        row.dot.setProperty("state", "failed")
        repolish(row.dot)
        row.name.setProperty("state", "failed")
        repolish(row.name)
        row.duration.setText(error or "ناموفق")
        row.duration.setToolTip(error)
        row.duration.setProperty("state", "failed")
        repolish(row.duration)


class PromptChips(QWidget):
    chosen = Signal(str)

    def __init__(self, prompts: Iterable[str], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = FlowLayout(self, spacing=6)
        for prompt in prompts:
            button = QPushButton(prompt)
            button.setObjectName("PromptChip")
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.clicked.connect(lambda _checked=False, text=prompt: self.chosen.emit(text))
            layout.addWidget(button)


class FlowLayout(QLayout):
    """Small wrapping layout so prompt chips stay tidy at any window width."""

    def __init__(self, parent: QWidget | None = None, spacing: int = 6) -> None:
        super().__init__(parent)
        self._items: list[QLayoutItem] = []
        self._spacing = spacing
        self.setContentsMargins(0, 0, 0, 0)

    def addItem(self, item: QLayoutItem) -> None:  # noqa: N802
        self._items.append(item)

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, index: int) -> QLayoutItem | None:  # noqa: N802
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index: int) -> QLayoutItem | None:  # noqa: N802
        return self._items.pop(index) if 0 <= index < len(self._items) else None

    def expandingDirections(self) -> Qt.Orientation:  # noqa: N802
        return Qt.Orientation(0x0)

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        return self._do_layout(QRectF(0, 0, width, 0), test_only=True)

    def setGeometry(self, rect: Any) -> None:  # noqa: N802
        super().setGeometry(rect)
        self._do_layout(QRectF(rect), test_only=False)

    def sizeHint(self) -> QSize:  # noqa: N802
        return self.minimumSize()

    def minimumSize(self) -> QSize:  # noqa: N802
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        return size

    def _do_layout(self, rect: QRectF, *, test_only: bool) -> int:
        x = rect.x()
        y = rect.y()
        line_height = 0
        for item in self._items:
            hint = item.sizeHint()
            next_x = x + hint.width() + self._spacing
            if next_x - self._spacing > rect.right() and line_height > 0:
                x = rect.x()
                y += line_height + self._spacing
                next_x = x + hint.width() + self._spacing
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(int(x), int(y), hint.width(), hint.height()))
            x = next_x
            line_height = max(line_height, hint.height())
        return int(y + line_height - rect.y())


class BarChart(QWidget):
    """Lightweight bar chart drawn with Qt; values are the engine's numbers."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._data: list[tuple[str, float]] = []
        self._palette = palette_for("light")
        self._bars: list[tuple[QRectF, str, float]] = []
        self.setMinimumHeight(250)
        self.setMouseTracking(True)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def set_data(self, data: Iterable[tuple[str, Any]], palette: Palette) -> None:
        cleaned: list[tuple[str, float]] = []
        for label, value in data:
            try:
                number = float(value)
            except (TypeError, ValueError):
                continue
            cleaned.append((str(label), number))
        self._data = cleaned
        self._palette = palette
        self.update()

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.update()

    @property
    def palette(self) -> Palette:
        return self._palette

    def paintEvent(self, event: QPaintEvent) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        palette = self._palette
        if not self._data:
            painter.setPen(QColor(palette.muted))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "داده عددی برای نمودار پیدا نشد.")
            return

        plot = self.rect().adjusted(16, 18, -16, -46)
        if plot.width() <= 20 or plot.height() <= 20:
            return
        maximum = max(value for _, value in self._data)
        minimum = min(value for _, value in self._data)
        high = maximum if maximum > 0 else 1.0
        low = minimum if minimum < 0 else 0.0
        span = high - low or 1.0

        grid_pen = QPen(QColor(palette.border))
        grid_pen.setWidth(1)
        painter.setPen(grid_pen)
        for index in range(5):
            y = plot.bottom() - plot.height() * index / 4
            painter.drawLine(int(plot.left()), int(y), int(plot.right()), int(y))

        count = len(self._data)
        slot = plot.width() / count
        bar_width = max(8.0, min(56.0, slot * 0.56))
        self._bars = []
        font = self.font()
        font.setPointSize(max(7, font.pointSize() - 2))
        painter.setFont(font)
        metrics = QFontMetrics(font)

        for index, (label, value) in enumerate(self._data):
            center = plot.right() - slot * (index + 0.5)
            height = (value - low) / span * plot.height()
            bar = QRectF(center - bar_width / 2, plot.bottom() - height, bar_width, max(2.0, height))
            gradient = QLinearGradient(bar.topLeft(), bar.bottomLeft())
            gradient.setColorAt(0.0, QColor(palette.accent_hover))
            gradient.setColorAt(1.0, QColor(palette.accent))
            painter.setBrush(gradient)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(bar, 6, 6)
            self._bars.append((bar, label, value))

            painter.setPen(QColor(palette.muted))
            value_text = format_number(value, decimals=1)
            painter.drawText(
                QRectF(bar.center().x() - slot / 2, bar.top() - 20, slot, 18),
                Qt.AlignmentFlag.AlignCenter,
                value_text,
            )
            label_text = metrics.elidedText(label, Qt.TextElideMode.ElideRight, int(max(40.0, slot - 6)))
            painter.drawText(
                QRectF(center - slot / 2, plot.bottom() + 6, slot, 34),
                Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
                label_text,
            )

    def mouseMoveEvent(self, event: QMouseEvent) -> None:  # noqa: N802
        for rect, label, value in self._bars:
            if rect.adjusted(-6, -10, 6, 10).contains(event.position()):
                self.setToolTip(f"{label}: {format_number(value)}")
                return
        self.setToolTip("")

    def leaveEvent(self, event: Any) -> None:  # noqa: N802
        self.setToolTip("")
        super().leaveEvent(event)


class RequestInput(QPlainTextEdit):
    def __init__(self, placeholder: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("RequestInput")
        self.setPlaceholderText(placeholder)
        self.setMinimumHeight(108)
        self.setMaximumHeight(180)
        self.setTabChangesFocus(True)


class ChatInput(QPlainTextEdit):
    submitted = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("ChatInput")
        self.setPlaceholderText("پرسش تکمیلی خود را بنویسید…")
        self.setFixedHeight(64)
        self.setTabChangesFocus(True)

    def keyPressEvent(self, event: Any) -> None:  # noqa: N802
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and not (
            event.modifiers() & Qt.KeyboardModifier.ShiftModifier
        ):
            self.submitted.emit()
            return
        super().keyPressEvent(event)


def make_label(text: str, object_name: str = "Muted", *, wrap: bool = False) -> QLabel:
    label = QLabel(text)
    label.setObjectName(object_name)
    label.setWordWrap(wrap)
    return label


def form_row(
    caption: str,
    widget: QWidget,
    *,
    label_width: int = 150,
    stretch: bool = False,
) -> QHBoxLayout:
    """One labelled form line: caption at the reading start, control beside it.

    Controls keep their own width instead of being stretched across the card,
    which is what made empty inputs look like large dark boxes.
    """
    row = QHBoxLayout()
    row.setSpacing(12)
    label = make_label(caption, "Muted")
    label.setFixedWidth(label_width)
    label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    row.addWidget(label)
    row.addWidget(widget, 1 if stretch else 0)
    if not stretch:
        row.addStretch(1)
    return row


def horizontal_line(parent: QWidget | None = None) -> QFrame:
    line = QFrame(parent)
    line.setObjectName("Divider")
    line.setFixedHeight(1)
    return line


def add_row(layout: QHBoxLayout, *widgets: QWidget, stretch_index: int | None = None) -> QHBoxLayout:
    for index, widget in enumerate(widgets):
        layout.addWidget(widget, 1 if index == stretch_index else 0)
    return layout


class CallbackButton(QPushButton):
    """A button whose label/tone can be updated after creation."""

    def __init__(
        self,
        text: str,
        callback: Callable[[], None] | None = None,
        *,
        object_name: str = "Ghost",
        icon_name: str = "",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(text, parent)
        self.setObjectName(object_name)
        self._icon_name = icon_name
        self._palette = palette_for("light")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        if callback is not None:
            self.clicked.connect(lambda: callback())
        if icon_name:
            self.apply_palette(self._palette)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        if self._icon_name:
            color = palette.on_accent if self.objectName() == "Primary" else palette.muted
            self.setIcon(icons.icon(self._icon_name, color, 16))
