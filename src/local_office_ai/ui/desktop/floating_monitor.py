"""Optional floating hardware widget that lives on the desktop.

The widget is a thin presentation layer: it reuses ``HardwareMonitor`` through
the existing ``HardwareWorker`` so there is exactly one sensor implementation.
Sampling starts only while the widget is visible and stops the moment it is
hidden, so a closed widget costs nothing.

It never feeds the report writer.  A sensor that cannot be read stays
«در دسترس نیست»; the widget never invents a value.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

from PySide6.QtCore import QPoint, Qt, QTimer, Signal
from PySide6.QtGui import QGuiApplication, QIcon
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import IconButton
from local_office_ai.ui.desktop.formatting import (
    UNAVAILABLE,
    elide,
    format_bytes,
    format_mb,
    format_percent,
    ltr,
)
from local_office_ai.ui.desktop.theme import Palette

AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_NAME = "LocalOfficeAI_FloatingMonitor"
MIN_OPACITY = 0.35


def autostart_supported() -> bool:
    """Windows-only feature; every other host reports it honestly as unsupported."""
    return os.name == "nt"


def autostart_command(project_root: Path) -> str:
    executable = Path(sys.executable)
    windowless = executable.with_name("pythonw.exe")
    python = windowless if windowless.is_file() else executable
    root = Path(project_root).resolve()
    script = root / "run_ui.py"
    return f'"{python}" "{script}" --root "{root}" --floating-only'


def is_autostart_enabled(project_root: Path | None = None) -> bool:
    if not autostart_supported():
        return False
    import winreg  # noqa: PLC0415 - Windows-only module

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_KEY) as key:
            value, _ = winreg.QueryValueEx(key, AUTOSTART_NAME)
    except OSError:
        return False
    if not value:
        return False
    if project_root is None:
        return True
    return str(value).strip().lower() == autostart_command(project_root).strip().lower()


def set_autostart(enabled: bool, project_root: Path) -> bool:
    """Add or remove the per-user Run entry. Returns False when unsupported."""
    if not autostart_supported():
        return False
    import winreg  # noqa: PLC0415 - Windows-only module

    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, AUTOSTART_KEY) as key:
            if enabled:
                winreg.SetValueEx(
                    key,
                    AUTOSTART_NAME,
                    0,
                    winreg.REG_SZ,
                    autostart_command(project_root),
                )
            else:
                try:
                    winreg.DeleteValue(key, AUTOSTART_NAME)
                except FileNotFoundError:
                    pass
        return True
    except OSError:
        return False


class _ReadingRow(QFrame):
    """One compact metric row: name, value and a muted detail line."""

    def __init__(self, icon_name: str, title: str, palette: Palette) -> None:
        super().__init__()
        self.setObjectName("FloatingRow")
        self._icon_name = icon_name
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 7, 10, 7)
        layout.setSpacing(2)

        top = QHBoxLayout()
        top.setSpacing(6)
        self.icon = QLabel()
        self.icon.setPixmap(icons.pixmap(icon_name, palette.muted, 14))
        self.name = QLabel(title)
        self.name.setObjectName("FloatingName")
        self.value = QLabel(UNAVAILABLE)
        self.value.setObjectName("FloatingValue")
        self.value.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        top.addWidget(self.icon)
        top.addWidget(self.name)
        top.addStretch(1)
        top.addWidget(self.value)
        layout.addLayout(top)

        self.detail = QLabel("")
        self.detail.setObjectName("FloatingDetail")
        self.detail.setWordWrap(False)
        layout.addWidget(self.detail)

    def set_reading(self, value: str | None, detail: str) -> None:
        self.value.setText(value if value is not None else UNAVAILABLE)
        self.detail.setText(elide(detail, 46))
        self.detail.setToolTip(detail)

    def apply_palette(self, palette: Palette) -> None:
        self.icon.setPixmap(icons.pixmap(self._icon_name, palette.muted, 14))


class FloatingMonitor(QWidget):
    """Frameless, always-on-top hardware widget with an optional tray icon."""

    visibility_changed = Signal(bool)
    quit_requested = Signal()

    def __init__(
        self,
        state: Any,
        palette: Palette,
        *,
        on_open_main: Any = None,
        standalone: bool = False,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.state = state
        self._palette = palette
        self._on_open_main = on_open_main
        self._standalone = standalone
        self._worker: Any = None
        self._retired: list[Any] = []
        self._drag_offset: QPoint | None = None
        self._placed = False
        self._tray: QSystemTrayIcon | None = None
        self._tray_menu: QMenu | None = None
        self._tray_message_shown = False
        self._quitting = False

        self.setObjectName("FloatingMonitor")
        self.setWindowTitle("پایش سخت‌افزار")
        self.setWindowFlags(
            Qt.WindowType.Tool
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setFixedWidth(306)
        self.setWindowOpacity(self.state.ui.widget_opacity)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        card = QFrame()
        card.setObjectName("FloatingCard")
        outer.addWidget(card)
        body = QVBoxLayout(card)
        body.setContentsMargins(12, 10, 12, 10)
        body.setSpacing(7)

        header = QHBoxLayout()
        header.setSpacing(6)
        self.brand_icon = QLabel()
        self.brand_icon.setPixmap(icons.pixmap("activity", palette.accent, 16))
        self.title = QLabel("پایش سخت‌افزار")
        self.title.setObjectName("FloatingTitle")
        self.provider_label = QLabel("")
        self.provider_label.setObjectName("FloatingProvider")
        self.provider_label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.hide_button = IconButton("eye", "پنهان کردن ویجت", size=26, icon_size=13)
        self.close_button = IconButton("close", "بستن و انتقال به سینی سیستم", size=26, icon_size=13)
        header.addWidget(self.brand_icon)
        header.addWidget(self.title)
        header.addStretch(1)
        header.addWidget(self.provider_label)
        header.addWidget(self.hide_button)
        header.addWidget(self.close_button)
        body.addLayout(header)

        self.cpu_row = _ReadingRow("cpu", "CPU", palette)
        self.ram_row = _ReadingRow("memory", "RAM", palette)
        self.gpu_row = _ReadingRow("gpu", "GPU", palette)
        for row in (self.cpu_row, self.ram_row, self.gpu_row):
            body.addWidget(row)

        self.footnote = QLabel("فقط نمایش — در فایل Excel نمی‌آید")
        self.footnote.setObjectName("FloatingFootnote")
        body.addWidget(self.footnote)

        self.hide_button.clicked.connect(self.hide_to_tray)
        self.close_button.clicked.connect(self.hide_to_tray)
        state.settings_changed.connect(self._on_settings_changed)
        self.setToolTip(
            "ویجت شناور پایش سخت‌افزار\n"
            "برای جابه‌جایی، ویجت را با دکمهٔ چپ ماوس بکشید.\n"
            "هنگام مخفی بودن، هیچ نمونه‌برداری‌ای انجام نمی‌شود."
        )

    # ------------------------------------------------------------------ #
    @property
    def is_sampling(self) -> bool:
        return self._worker is not None

    @property
    def tray_icon(self) -> QSystemTrayIcon | None:
        return self._tray

    # ------------------------------------------------------------------ #
    def show_widget(self) -> None:
        self.show()
        self.raise_()

    def toggle(self) -> None:
        if self.isVisible():
            self.hide_to_tray()
        else:
            self.show_widget()

    def deactivate(self) -> None:
        """Hide the widget and remove its tray icon (settings switch turned off)."""
        self.hide()
        self.state.set_widget_options(visible=False)
        self._remove_tray()

    def hide_to_tray(self) -> None:
        self.hide()
        if self._tray is not None:
            if not self._tray_message_shown:
                self._tray.showMessage(
                    "پایش سخت‌افزار",
                    "ویجت پنهان شد و نمونه‌برداری متوقف شد. از آیکن سینی سیستم دوباره باز کنید.",
                    QSystemTrayIcon.MessageIcon.Information,
                    4000,
                )
                self._tray_message_shown = True
        elif self._standalone:
            self.quit_requested.emit()

    # ------------------------------------------------------------------ #
    def showEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        super().showEvent(event)
        self._quitting = False
        if not self._placed:
            self._placed = True
            self._restore_geometry()
            QTimer.singleShot(0, self._reclamp)
        self.setWindowOpacity(self.state.ui.widget_opacity)
        self._start_sampling()
        self._ensure_tray()
        self.state.set_widget_options(visible=True)
        self.visibility_changed.emit(True)

    def hideEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        super().hideEvent(event)
        self._stop_sampling()
        self._persist_geometry()
        if not self._quitting:
            self.state.set_widget_options(visible=False)
        self.visibility_changed.emit(False)

    # ------------------------------------------------------------------ #
    def _reclamp(self) -> None:
        self._move_clamped(self.pos().x(), self.pos().y())

    def _restore_geometry(self) -> None:
        width, height = self.width(), max(self.sizeHint().height(), self.height(), 240)
        screen = QGuiApplication.primaryScreen()
        if screen is None:
            return
        area = screen.availableGeometry()
        x = int(self.state.ui.widget_x)
        y = int(self.state.ui.widget_y)
        if x < 0 or y < 0 or not area.contains(QPoint(x, y)):
            x = area.right() - width - 24
            y = area.bottom() - height - 24
        self._move_clamped(x, y)

    def _move_clamped(self, x: int, y: int) -> None:
        screen = QGuiApplication.screenAt(QPoint(x, y)) or QGuiApplication.primaryScreen()
        if screen is None:
            self.move(x, y)
            return
        area = screen.availableGeometry()
        width, height = self.width(), max(self.height(), 240)
        x = min(max(x, area.left()), area.right() - width)
        y = min(max(y, area.top()), area.bottom() - height)
        self.move(x, y)

    def _persist_geometry(self) -> None:
        position = self.pos()
        self.state.set_widget_options(position=(position.x(), position.y()))

    # ------------------------------------------------------------------ #
    def mousePressEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_offset = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        if self._drag_offset is not None and event.buttons() & Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: Any) -> None:  # noqa: N802 - Qt naming
        if self._drag_offset is not None and event.button() == Qt.MouseButton.LeftButton:
            self._drag_offset = None
            self._move_clamped(self.pos().x(), self.pos().y())
            self._persist_geometry()
            event.accept()
            return
        super().mouseReleaseEvent(event)

    # ------------------------------------------------------------------ #
    def set_opacity(self, value: float) -> None:
        opacity = max(MIN_OPACITY, min(1.0, float(value)))
        self.setWindowOpacity(opacity)
        self.state.set_widget_options(opacity=opacity)

    def apply_opacity(self, value: float) -> None:
        """Apply a value that the settings page already stored."""
        self.setWindowOpacity(max(MIN_OPACITY, min(1.0, float(value))))

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.brand_icon.setPixmap(icons.pixmap("activity", palette.accent, 16))
        self.hide_button.apply_palette(palette)
        self.close_button.apply_palette(palette)
        for row in (self.cpu_row, self.ram_row, self.gpu_row):
            row.apply_palette(palette)
        if self._tray is not None:
            self._tray.setIcon(QIcon(icons.pixmap("activity", palette.accent, 32)))

    # ------------------------------------------------------------------ #
    def _ensure_tray(self) -> None:
        if self._tray is not None or not QSystemTrayIcon.isSystemTrayAvailable():
            return
        tray = QSystemTrayIcon(QIcon(icons.pixmap("activity", self._palette.accent, 32)), self)
        tray.setToolTip("Local Office AI — پایش سخت‌افزار")
        menu = QMenu()
        toggle_action = menu.addAction("نمایش/پنهان پایش")
        toggle_action.triggered.connect(self.toggle)
        if callable(self._on_open_main):
            open_action = menu.addAction("باز کردن پنجره اصلی")
            open_action.triggered.connect(self._on_open_main)
        menu.addSeparator()
        quit_action = menu.addAction("خروج")
        quit_action.triggered.connect(self._quit)
        tray.setContextMenu(menu)
        tray.activated.connect(self._on_tray_activated)
        tray.show()
        self._tray = tray
        self._tray_menu = menu

    def _on_tray_activated(self, reason: Any) -> None:
        if reason in (
            QSystemTrayIcon.ActivationReason.Trigger,
            QSystemTrayIcon.ActivationReason.DoubleClick,
        ):
            self.toggle()

    def _quit(self) -> None:
        self.shutdown()
        self.quit_requested.emit()

    # ------------------------------------------------------------------ #
    def _start_sampling(self) -> None:
        self._stop_sampling()
        from local_office_ai.ui.desktop.workers import HardwareWorker

        self.provider_label.setText("…")
        worker = HardwareWorker(
            enable_gpu=self.state.ui.hardware_gpu_enabled,
            interval=self.state.ui.hardware_interval,
        )
        worker.snapshot.connect(self._on_snapshot)
        worker.failed.connect(self._on_failure)
        worker.start()
        self._worker = worker

    def _stop_sampling(self, *, wait: bool = False) -> None:
        worker = self._worker
        self._worker = None
        if worker is None:
            return
        worker.stop()
        if wait:
            worker.wait(5000)
            worker.deleteLater()
            return
        self._retired.append(worker)
        worker.finished.connect(lambda: self._retire(worker))

    def _retire(self, worker: Any) -> None:
        if worker in self._retired:
            self._retired.remove(worker)
        worker.deleteLater()

    def _remove_tray(self) -> None:
        if self._tray is not None:
            self._tray.hide()
            self._tray.setContextMenu(None)
            self._tray.setParent(None)
            self._tray = None
        if self._tray_menu is not None:
            self._tray_menu.deleteLater()
            self._tray_menu = None

    def shutdown(self) -> None:
        self._quitting = True
        self._stop_sampling(wait=True)
        for worker in list(self._retired):
            worker.stop()
            if worker.wait(2000):
                self._retire(worker)
        self._remove_tray()
        self.hide()

    # ------------------------------------------------------------------ #
    def _on_settings_changed(self) -> None:
        if self.isVisible():
            self._start_sampling()

    def _on_failure(self, message: str) -> None:
        self.provider_label.setText("خطا")
        self.provider_label.setToolTip(message)

    def _on_snapshot(self, snapshot: Any) -> None:
        if not self.isVisible():
            return
        cpu = snapshot.cpu
        cpu_value = format_percent(cpu.utilization) if cpu.utilization is not None else None
        cpu_parts: list[str] = []
        cpu_parts.append(
            f"دما: {cpu.temperature:.0f}°C" if cpu.temperature is not None else f"دما: {UNAVAILABLE}"
        )
        if cpu.logical_cores:
            cpu_parts.append(f"{cpu.logical_cores} هسته منطقی")
        if cpu.frequency_mhz:
            cpu_parts.append(f"{cpu.frequency_mhz:.0f} MHz")
        self.cpu_row.set_reading(cpu_value, " • ".join(cpu_parts))

        ram = snapshot.ram
        ram_value = format_percent(ram.percent) if ram.percent is not None else None
        if ram.used_bytes is not None and ram.total_bytes is not None:
            ram_detail = f"مصرف {format_bytes(ram.used_bytes)} از {format_bytes(ram.total_bytes)}"
        else:
            ram_detail = UNAVAILABLE
        self.ram_row.set_reading(ram_value, ram_detail)

        gpus = list(getattr(snapshot, "gpus", []) or [])
        if gpus:
            gpu = gpus[0]
            gpu_value = format_percent(gpu.utilization) if gpu.utilization is not None else None
            parts = [gpu.name]
            parts.append(
                f"دما: {gpu.temperature:.0f}°C" if gpu.temperature is not None else f"دما: {UNAVAILABLE}"
            )
            if gpu.vram_used_mb is not None and gpu.vram_total_mb is not None:
                parts.append(f"VRAM: {format_mb(gpu.vram_used_mb)} از {format_mb(gpu.vram_total_mb)}")
            elif gpu.vram_used_mb is not None:
                parts.append(f"VRAM مصرف‌شده: {format_mb(gpu.vram_used_mb)}")
            else:
                parts.append(f"VRAM: {UNAVAILABLE}")
            self.gpu_row.set_reading(gpu_value, " • ".join(parts))
        else:
            self.gpu_row.set_reading(None, "کارت گرافیک قابل شناسایی نبود")

        providers = getattr(snapshot, "providers", {}) or {}
        names: list[str] = []
        for key in ("cpu", "cpu_temp", "gpu"):
            value = providers.get(key)
            if value and value != UNAVAILABLE and value not in names:
                names.append(str(value))
        self.provider_label.setText(ltr(" • ".join(names)) if names else "سیستم")
        notes = getattr(snapshot, "notes", None) or []
        self.provider_label.setToolTip(
            "\n".join([f"به‌روزرسانی هر {self.state.ui.hardware_interval:.1f} ثانیه", *notes])
        )
