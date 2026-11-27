"""Window chrome: right-hand navigation, header status and hardware drawer."""

from __future__ import annotations

from typing import Any

from PySide6.QtCore import QEasingCurve, QPropertyAnimation, Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    Badge,
    Chip,
    FlowLayout,
    IconButton,
    NavigationButton,
    horizontal_line,
    make_label,
)
from local_office_ai.ui.desktop.formatting import (
    UNAVAILABLE,
    elide,
    format_bytes,
    format_mb,
    format_percent,
    ltr,
)
from local_office_ai.ui.desktop.theme import Palette

NAV_SECTIONS = (
    (
        "فضای کار",
        (
            ("home", "خانه", "home"),
            ("file", "تحلیل فایل‌ها", "file"),
            ("folder", "تحلیل پوشه", "folder"),
            ("chat", "گفت‌وگوی هوشمند", "chat"),
            ("reports", "گزارش‌های قبلی", "history"),
        ),
    ),
    ("سیستم", (("settings", "تنظیمات", "settings"),)),
)


class Sidebar(QFrame):
    navigate = Signal(str)
    collapsed_changed = Signal(bool)

    WIDTH = 232
    COLLAPSED_WIDTH = 76

    def __init__(self, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Sidebar")
        self.setFixedWidth(self.WIDTH)
        self._palette = palette
        self._buttons: dict[str, NavigationButton] = {}
        self._section_labels: list[QLabel] = []
        self._current = "home"
        self._collapsed = False

        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(14, 18, 14, 14)
        self._layout.setSpacing(4)

        self._brand = QWidget()
        brand_row = QHBoxLayout(self._brand)
        brand_row.setContentsMargins(2, 0, 2, 0)
        brand_row.setSpacing(10)
        self._mark = QLabel("AI")
        self._mark.setObjectName("BrandMark")
        self._mark.setFixedSize(40, 40)
        self._mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._brand_text = QWidget()
        titles = QVBoxLayout(self._brand_text)
        titles.setContentsMargins(0, 0, 0, 0)
        titles.setSpacing(0)
        title = QLabel("دستیار اداری")
        title.setObjectName("BrandTitle")
        subtitle = QLabel("Local Office AI")
        subtitle.setObjectName("BrandSub")
        titles.addWidget(title)
        titles.addWidget(subtitle)
        brand_row.addWidget(self._mark)
        brand_row.addWidget(self._brand_text, 1)
        self._brand_row = brand_row
        self._layout.addWidget(self._brand)
        self._layout.addSpacing(12)

        divider = QFrame()
        divider.setObjectName("SidebarDivider")
        divider.setFixedHeight(1)
        self._layout.addWidget(divider)
        self._layout.addSpacing(8)

        for section, items in NAV_SECTIONS:
            label = QLabel(section)
            label.setObjectName("NavSection")
            label.setContentsMargins(6, 6, 6, 2)
            self._section_labels.append(label)
            self._layout.addWidget(label)
            for key, caption, icon_name in items:
                self._layout.addWidget(self._nav_button(key, caption, icon_name))
            self._layout.addSpacing(8)

        self._layout.addStretch(1)

        footer = QHBoxLayout()
        footer.setContentsMargins(2, 0, 2, 0)
        footer.setSpacing(6)
        self._version = QLabel("نسخه ۴.۲")
        self._version.setObjectName("SidebarNote")
        self.toggle_button = IconButton("sidebar", "جمع کردن منو", size=32, icon_size=16)
        self.toggle_button.clicked.connect(lambda: self.set_collapsed(not self._collapsed))
        footer.addWidget(self._version, 1)
        footer.addWidget(self.toggle_button)
        self._footer = footer
        self._layout.addLayout(footer)

        self.set_current("home")

    # ------------------------------------------------------------------ #
    def _nav_button(self, key: str, label: str, icon_name: str) -> NavigationButton:
        button = NavigationButton(label, icon_name, self._palette)
        button.setProperty("nav_key", key)
        button.clicked.connect(lambda _checked=False, name=key: self.navigate.emit(name))
        self._buttons[key] = button
        return button

    def set_current(self, key: str) -> None:
        self._current = key
        for name, button in self._buttons.items():
            button.setChecked(name == key)

    @property
    def collapsed(self) -> bool:
        return self._collapsed

    def set_collapsed(self, collapsed: bool) -> None:
        collapsed = bool(collapsed)
        if collapsed == self._collapsed:
            return
        self._collapsed = collapsed
        margin = 10 if collapsed else 14
        self._layout.setContentsMargins(margin, 18, margin, 14)
        self._brand_text.setVisible(not collapsed)
        self._version.setVisible(not collapsed)
        for label in self._section_labels:
            label.setVisible(not collapsed)
        for button in self._buttons.values():
            button.set_collapsed(collapsed)
        self._brand_row.setAlignment(
            Qt.AlignmentFlag.AlignHCenter if collapsed else Qt.AlignmentFlag.AlignLeft
        )
        self.toggle_button.setToolTip("باز کردن منو" if collapsed else "جمع کردن منو")
        self.setFixedWidth(self.COLLAPSED_WIDTH if collapsed else self.WIDTH)
        self.collapsed_changed.emit(collapsed)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        for button in self._buttons.values():
            button.apply_palette(palette)
        self.toggle_button.apply_palette(palette)


class HeaderBar(QFrame):
    refresh_ollama = Signal()
    open_settings = Signal()
    toggle_monitor = Signal()

    def __init__(self, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Header")
        self._palette = palette
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        titles = QVBoxLayout()
        titles.setSpacing(1)
        self.title = QLabel("خانه")
        self.title.setObjectName("PageTitle")
        self.subtitle = QLabel("")
        self.subtitle.setObjectName("PageSubtitle")
        titles.addWidget(self.title)
        titles.addWidget(self.subtitle)
        layout.addLayout(titles)
        layout.addStretch(1)

        self.ollama_chip = Chip("در حال بررسی اتصال…", "info")
        self.model_chip = Chip("مدل: —", "muted")
        self.model_chip.setToolTip("مدل فعال Ollama برای نوشتن متن مدیریتی")
        layout.addWidget(self.ollama_chip)
        layout.addWidget(self.model_chip)
        self.refresh_button = IconButton("refresh", "بررسی دوباره اتصال Ollama")
        self.monitor_button = IconButton("monitor", "نمایش یا پنهان‌کردن ویجت شناور سخت‌افزار")
        self.settings_button = IconButton("settings", "تنظیمات")
        layout.addWidget(self.refresh_button)
        layout.addWidget(self.monitor_button)
        layout.addWidget(self.settings_button)

        self.refresh_button.clicked.connect(self.refresh_ollama.emit)
        self.monitor_button.clicked.connect(self.toggle_monitor.emit)
        self.settings_button.clicked.connect(self.open_settings.emit)
        self.apply_palette(palette)

    def set_page(self, title: str, subtitle: str = "") -> None:
        self.title.setText(title)
        self.subtitle.setText(subtitle)
        self.subtitle.setVisible(bool(subtitle))

    def update_ollama(self, online: bool | None, model: str, error: str = "") -> None:
        if online is None:
            self.ollama_chip.set_status("در حال بررسی اتصال…", "info")
            self.model_chip.set_status("مدل: —", "muted")
            return
        if online:
            self.ollama_chip.set_status("Ollama متصل", "success")
            self.ollama_chip.setToolTip("ارتباط با سرویس محلی Ollama برقرار است.")
            if model:
                self.model_chip.set_status(f"مدل: {ltr(model)}", "accent")
            else:
                self.model_chip.set_status("بدون مدل — قواعد قطعی", "muted")
        else:
            self.ollama_chip.set_status("Ollama در دسترس نیست", "warning")
            self.ollama_chip.setToolTip(error or "سرویس Ollama در حال اجرا نیست.")
            self.model_chip.set_status("بدون مدل — قواعد قطعی", "muted")

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.refresh_button.apply_palette(palette)
        self.monitor_button.apply_palette(palette)
        self.settings_button.apply_palette(palette)


class Banner(QFrame):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Banner")
        self.setProperty("tone", "info")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 10, 10, 10)
        layout.setSpacing(8)
        self.icon_label = QLabel()
        self.text_label = QLabel("")
        self.text_label.setObjectName("BannerText")
        self.text_label.setWordWrap(True)
        self.close_button = IconButton("close", "بستن پیام", size=28, icon_size=14)
        layout.addWidget(self.icon_label, 0, Qt.AlignmentFlag.AlignTop)
        layout.addWidget(self.text_label, 1)
        layout.addWidget(self.close_button, 0, Qt.AlignmentFlag.AlignTop)
        self.close_button.clicked.connect(self.hide)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)
        self.hide()

    def show_message(self, text: str, tone: str = "info", *, timeout_ms: int = 9000) -> None:
        self.text_label.setText(text)
        self.setProperty("tone", tone)
        self.style().unpolish(self)
        self.style().polish(self)
        color = {
            "danger": "#B42318",
            "warning": "#B45309",
            "success": "#15803D",
            "info": "#1D4ED8",
        }.get(tone, "#1D4ED8")
        self.icon_label.setPixmap(icons.pixmap("alert" if tone in {"danger", "warning"} else "info", color, 18))
        self.show()
        if timeout_ms > 0:
            self._timer.start(timeout_ms)

    def apply_palette(self, palette: Palette) -> None:
        self.close_button.apply_palette(palette)


class _HardwareTile(QFrame):
    def __init__(self, icon_name: str, title: str, palette: Palette) -> None:
        super().__init__()
        self.setObjectName("HardwareTile")
        self.setMinimumWidth(210)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(2)
        head = QHBoxLayout()
        head.setSpacing(6)
        self.icon = QLabel()
        self.icon.setPixmap(icons.pixmap(icon_name, palette.muted, 15))
        self.title = QLabel(title)
        self.title.setObjectName("HardwareName")
        head.addWidget(self.icon)
        head.addWidget(self.title, 1)
        layout.addLayout(head)
        self.value = QLabel(UNAVAILABLE)
        self.value.setObjectName("HardwareValue")
        self.value.setAlignment(Qt.AlignmentFlag.AlignRight)
        layout.addWidget(self.value)
        self.detail = QLabel("")
        self.detail.setObjectName("Caption")
        self.detail.setWordWrap(False)
        self.detail.setAlignment(Qt.AlignmentFlag.AlignRight)
        layout.addWidget(self.detail)

    def set_values(self, value: str | None, detail: str) -> None:
        if value is None:
            self.value.setText(UNAVAILABLE)
            self.value.setObjectName("Unavailable")
        else:
            self.value.setText(value)
            self.value.setObjectName("HardwareValue")
        self.value.style().unpolish(self.value)
        self.value.style().polish(self.value)
        self.detail.setText(elide(detail, 56))
        self.detail.setToolTip(detail)


class HardwarePanel(QFrame):
    """Collapsed by default; sampling runs only while the drawer is open."""

    def __init__(self, state: Any, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("HardwarePanel")
        self.state = state
        self._palette = palette
        self._worker: Any = None
        self._retired: list[Any] = []
        self._expanded = False
        self._gpu_tiles: list[_HardwareTile] = []
        self._cpu_temperature: float | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 6, 18, 6)
        layout.setSpacing(6)

        bar = QHBoxLayout()
        bar.setSpacing(8)
        self.toggle_button = IconButton("chevron-up", "نمایش پایش سخت‌افزار")
        self.monitor_icon = QLabel()
        self.monitor_icon.setPixmap(icons.pixmap("activity", palette.muted, 16))
        self.title = QLabel("پایش سخت‌افزار")
        self.title.setObjectName("CardTitle")
        self.summary = QLabel("جمع‌شده — برای مشاهده باز کنید")
        self.summary.setObjectName("Caption")
        bar.addWidget(self.toggle_button)
        bar.addWidget(self.monitor_icon)
        bar.addWidget(self.title)
        bar.addWidget(self.summary, 1)
        self.provider_label = QLabel("")
        self.provider_label.setObjectName("Caption")
        self.provider_label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        bar.addWidget(self.provider_label)
        layout.addLayout(bar)

        self.content = QWidget()
        content_layout = QVBoxLayout(self.content)
        content_layout.setContentsMargins(0, 4, 0, 8)
        content_layout.setSpacing(6)
        self.tiles = QWidget()
        self.tiles_layout = FlowLayout(self.tiles, spacing=8)
        content_layout.addWidget(self.tiles)
        self.detail_label = QLabel(
            "این اعداد فقط برای نمایش هستند و هرگز به فایل Excel خروجی اضافه نمی‌شوند. "
            "سنسور غیرقابل خواندن با «در دسترس نیست» نشان داده می‌شود."
        )
        self.detail_label.setObjectName("Caption")
        self.detail_label.setWordWrap(True)
        content_layout.addWidget(self.detail_label)
        self.content.setVisible(False)
        layout.addWidget(self.content)

        self._cpu_tile = _HardwareTile("cpu", "CPU", palette)
        self._ram_tile = _HardwareTile("memory", "RAM", palette)
        self.tiles_layout.addWidget(self._cpu_tile)
        self.tiles_layout.addWidget(self._ram_tile)
        self._gpu_placeholder = _HardwareTile("gpu", "GPU", palette)
        self.tiles_layout.addWidget(self._gpu_placeholder)

        self.toggle_button.clicked.connect(self.toggle)
        state.settings_changed.connect(self._on_settings_changed)

    # ------------------------------------------------------------------ #
    @property
    def expanded(self) -> bool:
        return self._expanded

    def toggle(self) -> None:
        self.set_expanded(not self._expanded)

    def set_expanded(self, expanded: bool) -> None:
        self._expanded = bool(expanded)
        self.content.setVisible(self._expanded)
        self.toggle_button.set_icon_name("chevron-down" if self._expanded else "chevron-up")
        self.toggle_button.setToolTip("بستن پایش" if self._expanded else "نمایش پایش سخت‌افزار")
        self.summary.setText("" if self._expanded else "جمع‌شده — برای مشاهده باز کنید")
        if self._expanded:
            self._start()
        else:
            self._stop()

    def _start(self) -> None:
        self._stop()
        from local_office_ai.ui.desktop.workers import HardwareWorker

        self.provider_label.setText("در حال خواندن سنسورها…")
        worker = HardwareWorker(
            enable_gpu=self.state.ui.hardware_gpu_enabled,
            interval=self.state.ui.hardware_interval,
        )
        worker.snapshot.connect(self._on_snapshot)
        worker.failed.connect(self._on_failure)
        worker.start()
        self._worker = worker

    def _stop(self, *, wait: bool = False) -> None:
        worker = self._worker
        self._worker = None
        if worker is None:
            return
        worker.stop()
        if wait:
            worker.wait(6000)
            worker.deleteLater()
            return
        self._retired.append(worker)
        worker.finished.connect(lambda: self._retire_done(worker))

    def _retire_done(self, worker: Any) -> None:
        if worker in self._retired:
            self._retired.remove(worker)
        worker.deleteLater()

    def shutdown(self) -> None:
        self._stop(wait=True)
        for worker in list(self._retired):
            worker.stop()
            if worker.wait(2500):
                self._retire_done(worker)

    def _on_settings_changed(self) -> None:
        if self._expanded:
            self._start()

    def _on_failure(self, message: str) -> None:
        self.provider_label.setText(f"خطای پایش: {message}")

    def _on_snapshot(self, snapshot: Any) -> None:
        if not self._expanded:
            return
        cpu = snapshot.cpu
        cpu_value = format_percent(cpu.utilization) if cpu.utilization is not None else None
        cpu_parts = []
        if cpu.temperature is not None:
            cpu_parts.append(f"دما: {cpu.temperature:.0f}°C")
        else:
            cpu_parts.append("دما: " + UNAVAILABLE)
        if cpu.logical_cores:
            cpu_parts.append(f"{cpu.logical_cores} هسته منطقی")
        if cpu.frequency_mhz:
            cpu_parts.append(f"{cpu.frequency_mhz:.0f} MHz")
        self._cpu_tile.set_values(cpu_value, " • ".join(cpu_parts))

        ram = snapshot.ram
        ram_value = format_percent(ram.percent) if ram.percent is not None else None
        if ram.total_bytes is not None and ram.used_bytes is not None:
            ram_detail = (
                f"مصرف {format_bytes(ram.used_bytes)} از {format_bytes(ram.total_bytes)}"
                + (f" • آزاد {format_bytes(ram.available_bytes)}" if ram.available_bytes is not None else "")
            )
        else:
            ram_detail = UNAVAILABLE
        self._ram_tile.set_values(ram_value, ram_detail)

        gpus = list(getattr(snapshot, "gpus", []) or [])
        self._sync_gpu_tiles(len(gpus))
        for tile, gpu in zip(self._gpu_tiles, gpus):
            value = format_percent(gpu.utilization) if gpu.utilization is not None else None
            parts = [str(gpu.name)]
            parts.append(f"دما: {gpu.temperature:.0f}°C" if gpu.temperature is not None else "دما: " + UNAVAILABLE)
            if gpu.vram_used_mb is not None and gpu.vram_total_mb is not None:
                parts.append(f"VRAM: {format_mb(gpu.vram_used_mb)} از {format_mb(gpu.vram_total_mb)}")
            elif gpu.vram_used_mb is not None:
                parts.append(f"VRAM مصرف‌شده: {format_mb(gpu.vram_used_mb)}")
            else:
                parts.append("VRAM: " + UNAVAILABLE)
            tile.set_values(value, " • ".join(parts))
        if not gpus:
            self._gpu_placeholder.set_values(None, "کارت گرافیک قابل شناسایی نبود")

        providers = getattr(snapshot, "providers", {}) or {}
        names: list[str] = []
        for key in ("cpu", "cpu_temp", "gpu"):
            value = providers.get(key)
            if value and value != UNAVAILABLE and value not in names:
                names.append(str(value))
        source = " • ".join(ltr(name) for name in names) if names else "سیستم"
        self.provider_label.setText(source)
        tips = [f"به‌روزرسانی هر {self.state.ui.hardware_interval:.1f} ثانیه"]
        tips.extend(getattr(snapshot, "notes", None) or [])
        self.provider_label.setToolTip("\n".join(tips))

    def _sync_gpu_tiles(self, count: int) -> None:
        count = max(0, count)
        if self._gpu_placeholder is not None:
            self._gpu_placeholder.setVisible(count == 0)
        while len(self._gpu_tiles) < count:
            tile = _HardwareTile("gpu", f"GPU {len(self._gpu_tiles) + 1}", self._palette)
            self.tiles_layout.addWidget(tile)
            self._gpu_tiles.append(tile)
        for index, tile in enumerate(self._gpu_tiles):
            tile.setVisible(index < count)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.toggle_button.apply_palette(palette)
        self.monitor_icon.setPixmap(icons.pixmap("activity", palette.muted, 16))
