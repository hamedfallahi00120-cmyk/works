"""Settings page: Persian font, theme, Ollama, paths and the monitor panel."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSlider,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import fonts as font_utils
from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    Badge,
    CallbackButton,
    Card,
    Chip,
    SectionHeader,
    form_row,
    horizontal_line,
    make_label,
)
from local_office_ai.ui.desktop.formatting import format_number
from local_office_ai.ui.desktop.pages.base import Page
from local_office_ai.ui.desktop.theme import Palette


class SettingsPage(Page):
    refresh_ollama = Signal()
    saved = Signal(str)
    widget_toggled = Signal(bool)
    widget_opacity_changed = Signal(float)
    widget_autostart_changed = Signal(bool)

    def __init__(self, state: Any, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self._palette = palette
        self._font_timer = QTimer(self)
        self._font_timer.setSingleShot(True)
        self._font_timer.setInterval(350)
        self._font_timer.timeout.connect(self._apply_font_choice)
        self._build()
        state.ollama_changed.connect(self._refresh_ollama_ui)
        state.model_changed.connect(self._refresh_model_combo)
        self._refresh_ollama_ui()
        self._refresh_paths()
        self._refresh_hardware()
        self._refresh_widget()
        self._rebuild_font_list()

    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        self.font_card = self._font_card()
        self.appearance_card = self._appearance_card()
        self.ollama_card = self._ollama_card()
        self.paths_card = self._paths_card()
        self.hardware_card = self._hardware_card()
        self.widget_card = self._widget_card()
        self.about_card = self._about_card()
        for card in (
            self.font_card,
            self.appearance_card,
            self.ollama_card,
            self.paths_card,
            self.hardware_card,
            self.widget_card,
            self.about_card,
        ):
            self.add(card)

        footer = QHBoxLayout()
        self.save_status = make_label("تنظیمات به‌صورت خودکار ذخیره می‌شود.", "Caption")
        save_button = CallbackButton("ذخیره تنظیمات", self._save, object_name="Primary", icon_name="check")
        footer.addWidget(self.save_status, 1)
        footer.addWidget(save_button)
        self.add_layout(footer)

    # -- font ----------------------------------------------------------- #
    def _font_card(self) -> Card:
        card = Card()
        card.add(
            SectionHeader(
                "فونت فارسی",
                "فقط فونت‌های نصب‌شده‌ای پیشنهاد می‌شوند که حروف ویژه فارسی (پ، چ، ژ، گ) را دارند.",
            )
        )
        self.font_count = make_label("", "Caption", wrap=True)
        self.family_combo = QComboBox()
        self.family_combo.setFixedWidth(340)
        self.family_combo.currentTextChanged.connect(self._schedule_font)
        self.size_spin = QSpinBox()
        self.size_spin.setRange(font_utils.MIN_SIZE, font_utils.MAX_SIZE)
        self.size_spin.setSuffix(" pt")
        self.size_spin.setFixedWidth(110)
        self.size_spin.valueChanged.connect(self._schedule_font)
        self.heading_spin = QSpinBox()
        self.heading_spin.setRange(font_utils.MIN_HEADING, font_utils.MAX_HEADING)
        self.heading_spin.setSuffix(" pt")
        self.heading_spin.setFixedWidth(110)
        self.heading_spin.valueChanged.connect(self._schedule_font)
        form = QVBoxLayout()
        form.setSpacing(10)
        form.addLayout(form_row("فونت برنامه:", self.family_combo))
        form.addLayout(form_row("اندازه متن:", self.size_spin))
        form.addLayout(form_row("اندازه عنوان‌ها:", self.heading_spin))
        card.add_layout(form)
        card.add(self.font_count)

        preview = Card(variant="alt", margins=(16, 12, 16, 12), spacing=6)
        self.preview_heading = QLabel("گزارش مدیریتی داده‌های اداری")
        self.preview_heading.setWordWrap(True)
        self.preview_body = QLabel(font_utils.PERSIAN_SAMPLE)
        self.preview_body.setWordWrap(True)
        preview.add(self.preview_heading)
        preview.add(self.preview_body)
        card.add(preview)

        custom_row = QHBoxLayout()
        self.custom_font_label = make_label("فونت سفارشی: استفاده نمی‌شود", "Caption", wrap=True)
        load_button = CallbackButton("انتخاب فایل فونت", self._load_custom_font, object_name="Ghost", icon_name="type")
        self.remove_custom_button = CallbackButton(
            "حذف فونت سفارشی", self._remove_custom_font, object_name="Quiet", icon_name="trash"
        )
        custom_row.addWidget(self.custom_font_label, 1)
        custom_row.addWidget(self.remove_custom_button)
        custom_row.addWidget(load_button)
        card.add_layout(custom_row)
        card.add(
            make_label(
                "هیچ فونتی دانلود یا در بسته قرار داده نمی‌شود؛ اگر فونت مناسبی نصب نباشد، "
                "نزدیک‌ترین فونت فارسی سیستم یا فونت پیش‌فرض Qt استفاده می‌شود.",
                "Caption",
                wrap=True,
            )
        )
        reset_row = QHBoxLayout()
        reset_row.addStretch(1)
        self.reset_font_button = CallbackButton("بازنشانی ظاهر پیش‌فرض", self._reset_font, object_name="Quiet", icon_name="refresh")
        reset_row.addWidget(self.reset_font_button)
        card.add_layout(reset_row)
        return card

    def _rebuild_font_list(self) -> None:
        families = font_utils.persian_families(refresh=True)
        installed = set(QFontDatabase.families())
        saved = self.state.ui.font_family
        if not families:
            fallback = font_utils.resolve_family("") or QApplication.font().family()
            if fallback:
                families = [fallback]
        if saved and saved not in families and saved in installed:
            families.insert(0, saved)
        self.family_combo.blockSignals(True)
        self.family_combo.clear()
        self.family_combo.addItems(families)
        if saved:
            self.family_combo.setCurrentText(saved)
        elif families:
            self.family_combo.setCurrentIndex(0)
        self.family_combo.blockSignals(False)
        self.size_spin.blockSignals(True)
        self.heading_spin.blockSignals(True)
        self.size_spin.setValue(self.state.ui.font_base_size)
        self.heading_spin.setValue(self.state.ui.font_heading_size)
        self.size_spin.blockSignals(False)
        self.heading_spin.blockSignals(False)
        if families:
            self.font_count.setText(
                f"{format_number(len(families))} فونت نصب‌شده با پشتیبانی فارسی شناسایی شد."
            )
            self.font_count.setObjectName("Caption")
        else:
            self.font_count.setText(
                "هیچ فونت نصب‌شده‌ای با پشتیبانی کامل فارسی پیدا نشد؛ فونت پیش‌فرض Qt استفاده می‌شود. "
                "می‌توانید یک فایل فونت محلی انتخاب کنید."
            )
            self.font_count.setObjectName("WarningText")
        self.font_count.style().unpolish(self.font_count)
        self.font_count.style().polish(self.font_count)
        self.custom_font_label.setText(
            f"فونت سفارشی: {Path(self.state.ui.font_file).name}"
            if self.state.ui.font_file
            else "فونت سفارشی: استفاده نمی‌شود"
        )
        self.remove_custom_button.setVisible(bool(self.state.ui.font_file))
        self._update_preview()

    def flush_font(self) -> None:
        if self._font_timer.isActive():
            self._font_timer.stop()
            self._apply_font_choice()

    def _schedule_font(self, *_: Any) -> None:
        self._update_preview()
        self._font_timer.start()

    def _apply_font_choice(self) -> None:
        family = self.family_combo.currentText()
        if not family:
            return
        self.state.set_font(family, self.size_spin.value(), self.heading_spin.value())
        self.save_status.setText("فونت ذخیره شد و روی کل برنامه اعمال شد.")

    def _update_preview(self) -> None:
        family = self.family_combo.currentText()
        base = self.size_spin.value()
        heading = self.heading_spin.value()
        family_rule = f'font-family: "{family}";' if family else ""
        self.preview_heading.setStyleSheet(f"{family_rule} font-size: {heading}pt; font-weight: 700;")
        self.preview_body.setStyleSheet(f"{family_rule} font-size: {base}pt;")

    def _load_custom_font(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "انتخاب فایل فونت",
            "",
            "Font files (*.ttf *.otf *.ttc *.otc)",
        )
        if not path:
            return
        family, error = font_utils.load_font_file(path)
        if error:
            self.font_count.setText("بارگذاری فونت ممکن نشد: " + error)
            self.font_count.setObjectName("ErrorText")
            self.font_count.style().unpolish(self.font_count)
            self.font_count.style().polish(self.font_count)
            return
        self.state.set_font_file(path)
        self._rebuild_font_list()
        if family:
            self.family_combo.blockSignals(True)
            self.family_combo.setCurrentText(family)
            self.family_combo.blockSignals(False)
            self.state.set_font(family, self.size_spin.value(), self.heading_spin.value())
        self.save_status.setText("فونت سفارشی بارگذاری و ذخیره شد.")

    def _remove_custom_font(self) -> None:
        self.state.set_font_file("")
        family = font_utils.recommended_default()
        self.state.set_font(family, self.size_spin.value(), self.heading_spin.value())
        self._rebuild_font_list()
        self.save_status.setText("فونت سفارشی حذف شد و فونت پیش‌فرض اعمال شد.")

    def _reset_font(self) -> None:
        family = font_utils.recommended_default()
        self.state.set_font_file("")
        self.state.ui.font_base_size = 12
        self.state.ui.font_heading_size = 19
        if family:
            self.family_combo.blockSignals(True)
            self.family_combo.setCurrentText(family)
            self.family_combo.blockSignals(False)
        self.size_spin.setValue(12)
        self.heading_spin.setValue(19)
        self.state.set_font(family or self.family_combo.currentText(), 12, 19)
        self._rebuild_font_list()
        self.save_status.setText("ظاهر به حالت پیش‌فرض برگشت.")

    # -- appearance ----------------------------------------------------- #
    def _appearance_card(self) -> Card:
        card = Card()
        card.add(SectionHeader("ظاهر برنامه", "تم روشن یا تیره؛ انتخاب روی همه پنجره‌ها اعمال می‌شود."))
        row = QHBoxLayout()
        row.setSpacing(8)
        self.light_button = QPushButton("روشن")
        self.light_button.setObjectName("Segment")
        self.light_button.setCheckable(True)
        self.light_button.setIcon(icons.icon("sun", self._palette.muted, 16))
        self.dark_button = QPushButton("تیره")
        self.dark_button.setObjectName("Segment")
        self.dark_button.setCheckable(True)
        self.dark_button.setIcon(icons.icon("moon", self._palette.muted, 16))
        group = QButtonGroup(self)
        group.setExclusive(True)
        group.addButton(self.light_button)
        group.addButton(self.dark_button)
        self.light_button.setChecked(self.state.ui.theme == "light")
        self.dark_button.setChecked(self.state.ui.theme == "dark")
        self.light_button.clicked.connect(lambda: self._set_theme("light"))
        self.dark_button.clicked.connect(lambda: self._set_theme("dark"))
        row.addWidget(self.light_button)
        row.addWidget(self.dark_button)
        row.addStretch(1)
        card.add_layout(row)
        self._theme_buttons = (self.light_button, self.dark_button)
        return card

    def _set_theme(self, theme: str) -> None:
        self.state.set_theme(theme)
        self.save_status.setText("تم ذخیره شد.")

    # -- Ollama --------------------------------------------------------- #
    def _ollama_card(self) -> Card:
        card = Card()
        card.add(
            SectionHeader(
                "Ollama و مدل محلی",
                "مدل فقط متن مدیریتی را می‌نویسد؛ همه اعداد از محاسبات Python است.",
            )
        )
        status_row = QHBoxLayout()
        self.ollama_status = Chip("در حال بررسی…", "info")
        self.refresh_button = CallbackButton("بررسی اتصال", self.refresh_ollama.emit, object_name="Ghost", icon_name="refresh")
        status_row.addWidget(self.ollama_status)
        status_row.addStretch(1)
        status_row.addWidget(self.refresh_button)
        card.add_layout(status_row)
        self.ollama_error = make_label("", "Caption", wrap=True)
        card.add(self.ollama_error)
        url = make_label(
            "آدرس سرویس: " + str(self.state.settings.ollama_base_url),
            "Caption",
            wrap=True,
        )
        card.add(url)
        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(300)
        self.model_combo.setMaximumWidth(420)
        self.model_combo.currentIndexChanged.connect(self._model_changed)
        card.add_layout(form_row("مدل فعال:", self.model_combo))
        card.add(
            make_label(
                "اگر Ollama در دسترس نباشد، تحلیل با قواعد قطعی Python انجام می‌شود و همین در گزارش ذکر می‌گردد.",
                "Caption",
                wrap=True,
            )
        )
        return card

    def _refresh_ollama_ui(self, *_: Any) -> None:
        if self.state.ollama_online:
            count = len(self.state.models)
            self.ollama_status.set_status(f"متصل — {count} مدل", "success")
            self.ollama_error.setText("")
        else:
            self.ollama_status.set_status("در دسترس نیست", "warning")
            self.ollama_error.setText(
                self.state.ollama_error
                or "سرویس Ollama اجرا نشده است. تحلیل بدون مدل هم کار می‌کند."
            )
        self._refresh_model_combo()

    def _refresh_model_combo(self, *_: Any) -> None:
        self.model_combo.blockSignals(True)
        self.model_combo.clear()
        self.model_combo.addItem("بدون مدل (قواعد قطعی)", "")
        for item in self.state.models:
            name = str(item.get("name", ""))
            label = name
            details = []
            if item.get("parameter_size"):
                details.append(str(item["parameter_size"]))
            if item.get("size_gb"):
                details.append(f"{item['size_gb']} GB")
            if details:
                label += " — " + " / ".join(details)
            self.model_combo.addItem(label, name)
        index = self.model_combo.findData(self.state.model)
        self.model_combo.setCurrentIndex(index if index >= 0 else 0)
        self.model_combo.setEnabled(self.state.ollama_online and bool(self.state.models))
        self.model_combo.blockSignals(False)

    def _model_changed(self, _index: int) -> None:
        name = self.model_combo.currentData() or ""
        self.state.set_model(str(name))
        self.save_status.setText("مدل فعال ذخیره شد.")

    # -- paths ---------------------------------------------------------- #
    def _paths_card(self) -> Card:
        card = Card()
        card.add(SectionHeader("مسیرها", "پوشه ورودی فقط خوانده می‌شود؛ گزارش‌ها در پوشه خروجی ساخته می‌شوند."))
        self.input_edit = QLineEdit()
        self.input_edit.setReadOnly(True)
        self.input_edit.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        input_button = CallbackButton("انتخاب…", self._browse_input, object_name="Ghost")
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        self.output_edit.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        output_button = CallbackButton("انتخاب…", self._browse_output, object_name="Ghost")
        input_row = form_row("پوشه ورودی:", self.input_edit, stretch=True)
        input_row.addWidget(input_button)
        output_row = form_row("پوشه خروجی:", self.output_edit, stretch=True)
        output_row.addWidget(output_button)
        card.add_layout(input_row)
        card.add_layout(output_row)
        row = QHBoxLayout()
        reset_paths = CallbackButton("بازگشت به مسیر پیش‌فرض", self._reset_paths, object_name="Quiet", icon_name="refresh")
        row.addWidget(reset_paths)
        row.addStretch(1)
        card.add_layout(row)
        card.add(
            make_label(
                "فایل ورودی هرگز تغییر نمی‌کند و گزارش با نام تازه در پوشه خروجی ذخیره می‌شود.",
                "Caption",
                wrap=True,
            )
        )
        return card

    def _refresh_paths(self) -> None:
        self.input_edit.setText(str(self.state.settings.input_dir))
        self.output_edit.setText(str(self.state.settings.output_dir))

    def _browse_input(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "انتخاب پوشه ورودی", str(self.state.settings.input_dir))
        if path:
            self.state.update_paths(input_dir=path)
            self._refresh_paths()
            self.save_status.setText("پوشه ورودی ذخیره شد.")

    def _browse_output(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "انتخاب پوشه خروجی", str(self.state.settings.output_dir))
        if path:
            self.state.update_paths(output_dir=path)
            self._refresh_paths()
            self.save_status.setText("پوشه خروجی ذخیره شد.")

    def _reset_paths(self) -> None:
        self.state.update_paths(input_dir="", output_dir="")
        self._refresh_paths()
        self.save_status.setText("مسیرهای پیش‌فرض بازگردانده شد.")

    # -- hardware ------------------------------------------------------- #
    def _hardware_card(self) -> Card:
        card = Card()
        card.add(
            SectionHeader(
                "پایش سخت‌افزار",
                "پنل پایین پنجره به‌صورت پیش‌فرض جمع است و فقط هنگام باز بودن اندازه‌گیری می‌کند.",
            )
        )
        self.gpu_checkbox = QCheckBox("خواندن اطلاعات کارت گرافیک (GPU و VRAM)")
        self.gpu_checkbox.setChecked(self.state.ui.hardware_gpu_enabled)
        self.gpu_checkbox.toggled.connect(self._hardware_changed)
        card.add(self.gpu_checkbox)
        self.interval_spin = QDoubleSpinBox()
        self.interval_spin.setRange(1.0, 10.0)
        self.interval_spin.setSingleStep(0.5)
        self.interval_spin.setDecimals(1)
        self.interval_spin.setSuffix(" ثانیه")
        self.interval_spin.setFixedWidth(130)
        self.interval_spin.setValue(self.state.ui.hardware_interval)
        self.interval_spin.valueChanged.connect(self._hardware_changed)
        card.add_layout(form_row("فاصله به‌روزرسانی:", self.interval_spin))
        card.add(
            make_label(
                "CPU و RAM از psutil خوانده می‌شوند. دمای CPU و اطلاعات GPU تنها اگر سنسور یا "
                "ابزار سیستمی موجود باشد خوانده می‌شود (nvidia-smi، sysfs لینوکس یا "
                "LibreHardwareMonitor در حال اجرا). هیچ ابزاری دانلود نمی‌شود و مقدار نامعلوم "
                "«در دسترس نیست» نمایش داده می‌شود.",
                "Caption",
                wrap=True,
            )
        )
        card.add(
            make_label(
                "این اطلاعات فقط در رابط دیده می‌شود و در فایل Excel خروجی هیچ برگه‌ای برای آن ساخته نمی‌شود.",
                "Caption",
                wrap=True,
            )
        )
        return card

    def _refresh_hardware(self) -> None:
        self.gpu_checkbox.blockSignals(True)
        self.interval_spin.blockSignals(True)
        self.gpu_checkbox.setChecked(self.state.ui.hardware_gpu_enabled)
        self.interval_spin.setValue(self.state.ui.hardware_interval)
        self.gpu_checkbox.blockSignals(False)
        self.interval_spin.blockSignals(False)

    def _hardware_changed(self, *_: Any) -> None:
        self.state.set_hardware_options(
            gpu_enabled=self.gpu_checkbox.isChecked(),
            interval=self.interval_spin.value(),
        )
        self.save_status.setText("تنظیمات پایش ذخیره شد.")

    # -- floating widget ------------------------------------------------ #
    def _widget_card(self) -> Card:
        from local_office_ai.ui.desktop.floating_monitor import autostart_supported

        card = Card()
        card.add(
            SectionHeader(
                "ویجت شناور سخت‌افزار",
                "پنجرهٔ کوچک مستقل روی دسکتاپ، همیشه روی سایر پنجره‌ها. فقط هنگام باز بودن نمونه‌برداری می‌کند.",
            )
        )
        self.widget_checkbox = QCheckBox("نمایش ویجت شناور روی دسکتاپ")
        self.widget_checkbox.setChecked(self.state.ui.widget_visible)
        self.widget_checkbox.toggled.connect(self._widget_toggled)
        card.add(self.widget_checkbox)

        row = QHBoxLayout()
        row.addWidget(make_label("شفافیت ویجت:", "Muted"))
        self.opacity_slider = QSlider(Qt.Orientation.Horizontal)
        self.opacity_slider.setRange(35, 100)
        self.opacity_slider.setSingleStep(5)
        self.opacity_slider.setPageStep(10)
        self.opacity_slider.setValue(int(round(self.state.ui.widget_opacity * 100)))
        self.opacity_slider.valueChanged.connect(self._widget_opacity_changed)
        self.opacity_value = make_label(f"{self.opacity_slider.value()}%", "Caption")
        row.addWidget(self.opacity_slider, 1)
        row.addWidget(self.opacity_value)
        card.add_layout(row)

        self.autostart_checkbox = QCheckBox("اجرای خودکار ویجت هنگام ورود به ویندوز")
        self._autostart_supported = autostart_supported()
        if self._autostart_supported:
            from local_office_ai.ui.desktop.floating_monitor import is_autostart_enabled

            self.autostart_checkbox.setChecked(is_autostart_enabled(self.state.project_root))
        else:
            self.autostart_checkbox.setChecked(False)
            self.autostart_checkbox.setEnabled(False)
            self.autostart_checkbox.setToolTip("این گزینه فقط در ویندوز اعمال می‌شود.")
        self.autostart_checkbox.toggled.connect(self._widget_autostart_toggled)
        card.add(self.autostart_checkbox)
        card.add(
            make_label(
                "ویجت با دکمهٔ نمایشگر در هدر پنجره یا از همین صفحه باز می‌شود. دکمهٔ بستن، آن را "
                "به سینی سیستم می‌فرستد و نمونه‌برداری را کاملاً متوقف می‌کند. محل و شفافیت "
                "ویجت ذخیره می‌شود. مقدار نامعلوم همیشه «در دسترس نیست» نمایش داده می‌شود.",
                "Caption",
                wrap=True,
            )
        )
        card.add(
            make_label(
                "اجرای خودکار فقط یک مقدار در Run ویندوز ثبت می‌کند و برنامه را با حالت "
                "--floating-only اجرا می‌کند؛ پنجرهٔ اصلی باز نمی‌شود.",
                "Caption",
                wrap=True,
            )
        )
        return card

    def _refresh_widget(self) -> None:
        self.widget_checkbox.blockSignals(True)
        self.opacity_slider.blockSignals(True)
        self.autostart_checkbox.blockSignals(True)
        self.widget_checkbox.setChecked(self.state.ui.widget_visible)
        self.opacity_slider.setValue(int(round(self.state.ui.widget_opacity * 100)))
        self.opacity_value.setText(f"{self.opacity_slider.value()}%")
        self.autostart_checkbox.setChecked(
            bool(self.state.ui.widget_autostart) and self._autostart_supported
        )
        self.widget_checkbox.blockSignals(False)
        self.opacity_slider.blockSignals(False)
        self.autostart_checkbox.blockSignals(False)

    def _widget_toggled(self, checked: bool) -> None:
        self.state.set_widget_options(visible=checked)
        self.save_status.setText("وضعیت ویجت شناور ذخیره شد.")
        self.widget_toggled.emit(checked)

    def _widget_opacity_changed(self, value: int) -> None:
        self.opacity_value.setText(f"{value}%")
        opacity = value / 100.0
        self.state.set_widget_options(opacity=opacity)
        self.widget_opacity_changed.emit(opacity)

    def _widget_autostart_toggled(self, checked: bool) -> None:
        self.widget_autostart_changed.emit(checked)

    def sync_widget_visibility(self, visible: bool) -> None:
        self.widget_checkbox.blockSignals(True)
        self.widget_checkbox.setChecked(bool(visible))
        self.widget_checkbox.blockSignals(False)

    def finish_autostart_change(self, enabled: bool, applied: bool) -> None:
        if applied:
            self.save_status.setText(
                "اجرای خودکار ویجت فعال شد." if enabled else "اجرای خودکار ویجت غیرفعال شد."
            )
            return
        self.autostart_checkbox.blockSignals(True)
        self.autostart_checkbox.setChecked(not enabled)
        self.autostart_checkbox.blockSignals(False)
        self.save_status.setText("اجرای خودکار فقط در ویندوز پشتیبانی می‌شود.")

    # -- about ---------------------------------------------------------- #
    def _about_card(self) -> Card:
        card = Card()
        card.add(SectionHeader("درباره این نسخه"))
        card.add(
            make_label(
                "Local Office AI — رابط مدرن PySide6 با چیدمان راست‌به‌چپ. "
                "موتور تحلیل V4.2 بدون تغییر استفاده می‌شود؛ مسیر CLI و رابط Tkinter قدیمی "
                "برای بازگشت حفظ شده‌اند.",
                "Muted",
                wrap=True,
            )
        )
        card.add(horizontal_line())
        card.add(make_label("لانچر جدید: Run_Local_Office_AI_Modern.bat", "Caption"))
        card.add(make_label("لانچر قدیمی Tkinter: Run_Local_Office_AI_GUI.bat", "Caption"))
        card.add(make_label(f"فایل تنظیمات رابط: {self.state.project_root / 'config' / 'ui_settings.json'}", "Caption", wrap=True))
        return card

    # ------------------------------------------------------------------ #
    def _save(self) -> None:
        self.state.save()
        self.save_status.setText("تنظیمات ذخیره شد.")
        self.saved.emit("تنظیمات ذخیره شد.")

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.light_button.setIcon(icons.icon("sun", palette.muted, 16))
        self.dark_button.setIcon(icons.icon("moon", palette.muted, 16))
        self._refresh_paths()
