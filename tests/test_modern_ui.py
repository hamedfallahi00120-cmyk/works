"""Acceptance checks for the modern PySide6 desktop interface.

These tests build the real window, run the real engine through its worker and
read the produced workbook.  No engine module is stubbed.
"""

from __future__ import annotations

import os
import time
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.*=false")

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication
except (ImportError, OSError) as exc:  # pragma: no cover - host without Qt
    pytest.skip(f"PySide6 is not usable here: {exc}", allow_module_level=True)

from openpyxl import load_workbook  # noqa: E402

from local_office_ai.core.settings import load_settings  # noqa: E402
from local_office_ai.ui.desktop import fonts  # noqa: E402
from local_office_ai.ui.desktop.main_window import MainWindow  # noqa: E402
from local_office_ai.ui.desktop.preview import (  # noqa: E402
    ReportPreviewDialog,
    read_workbook_preview,
)
from local_office_ai.ui.desktop.ui_settings import load_ui_settings  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
FORBIDDEN_SHEETS = {
    "Performance Monitoring",
    "GPU Monitoring",
    "CPU Monitoring",
    "RAM Monitoring",
    "Test Results",
    "System Information",
    "Internal Logs",
}


@pytest.fixture(scope="module")
def qt_app() -> QApplication:
    application = QApplication.instance() or QApplication([])
    yield application


def _project_root(tmp_path: Path) -> Path:
    root = tmp_path / "ui_project"
    for name in ("config", "data", "data/input", "data/output", "logs"):
        (root / name).mkdir(parents=True, exist_ok=True)
    return root


@pytest.fixture()
def window(qt_app: QApplication, tmp_path: Path):
    root = _project_root(tmp_path)
    settings = load_settings(project_root=root)
    settings.ensure_directories()
    ui = load_ui_settings(root)
    ui.theme = "light"
    ui.font_family = fonts.resolve_family("")
    ui.font_base_size = 12
    ui.font_heading_size = 19
    window = MainWindow(settings, ui)
    window.resize(1280, 860)
    window.show()
    qt_app.processEvents()
    yield window
    window.close()
    qt_app.processEvents()


def wait_for(qt_app: QApplication, predicate, timeout: float = 60.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        qt_app.processEvents()
        if predicate():
            return True
        time.sleep(0.02)
    qt_app.processEvents()
    return bool(predicate())


def test_window_rtl_font_theme_and_persistence(qt_app, window):
    assert window.stack.count() == 6
    assert qt_app.layoutDirection() == Qt.LayoutDirection.RightToLeft
    assert window.hardware_panel.expanded is False
    assert window.hardware_panel.content.isVisible() is False

    window.show_page("settings")
    qt_app.processEvents()
    assert window.settings_page.family_combo.count() >= 1
    assert window.settings_page.family_combo.currentText()

    window.settings_page._set_theme("dark")
    qt_app.processEvents()
    assert window._palette.name == "dark"

    window.settings_page.family_combo.setCurrentText(window.settings_page.family_combo.currentText())
    window.settings_page.size_spin.setValue(14)
    window.settings_page.heading_spin.setValue(22)
    window.settings_page.flush_font()
    qt_app.processEvents()

    saved = load_ui_settings(window.state.project_root)
    assert saved.theme == "dark"
    assert saved.font_base_size == 14
    assert saved.font_heading_size == 22
    assert saved.font_family == window.settings_page.family_combo.currentText()
    assert "14pt" in qt_app.styleSheet()

    window.settings_page._set_theme("light")
    qt_app.processEvents()
    assert window._palette.name == "light"


def test_real_analysis_chat_preview_and_no_technical_sheets(qt_app, window, sample_path):
    assert window._accept_file(sample_path) is True
    assert window.state.source == sample_path

    window._start_file("این اطلاعات را بررسی کن، تغییرات مهم را پیدا کن و گزارش مدیریتی تهیه کن.")
    assert window.state.busy is True
    assert wait_for(qt_app, lambda: window.state.last_result is not None, timeout=90)

    result = window.state.last_result
    output = Path(result.output_path)
    assert output.is_file()
    assert result.analysis.request_answers

    workbook = load_workbook(output, read_only=True, data_only=True)
    try:
        sheets = set(workbook.sheetnames)
    finally:
        workbook.close()
    assert not (sheets & FORBIDDEN_SHEETS)
    assert "پاسخ درخواست" in sheets
    assert "داده‌های پنل سخت‌افزار" not in sheets

    previews = read_workbook_preview(output)
    assert previews
    assert any(preview.name == "پاسخ درخواست" for preview in previews)
    dialog = ReportPreviewDialog(output, window)
    assert dialog.sheet_list.count() == len(previews)
    dialog.close()

    assert window.chat_page.session_analysis() is not None
    window.show_page("chat")
    before = len(window.chat_page._bubbles)
    window.chat_page.ask("بیشترین و کمترین ماه را با عدد مشخص کن.")
    assert wait_for(qt_app, lambda: not window.state.busy, timeout=60)
    assert len(window.chat_page._bubbles) > before
    assert window.chat_page._pending is None


def test_hardware_panel_reads_real_sensors_and_collapses(qt_app, window):
    window.hardware_panel.set_expanded(True)
    assert wait_for(
        qt_app,
        lambda: window.hardware_panel.provider_label.text()
        not in ("", "در حال خواندن سنسورها…"),
        timeout=8,
    )
    value = window.hardware_panel._cpu_tile.value.text()
    assert "%" in value or value == "در دسترس نیست"
    assert window.hardware_panel.provider_label.text()

    started = time.time()
    window.hardware_panel.set_expanded(False)
    qt_app.processEvents()
    assert time.time() - started < 1.0
    assert window.hardware_panel.content.isVisible() is False


def test_windows_launcher_and_dependency_are_present():
    launcher = (ROOT / "Run_Local_Office_AI_Modern.bat").read_text(encoding="utf-8")
    assert r"venv\Scripts\python.exe" in launcher
    assert "run_ui.py" in launcher
    assert (ROOT / "run_ui.py").is_file()
    assert (ROOT / "Run_Local_Office_AI_GUI.bat").is_file()
    requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")
    assert "PySide6" in requirements


def test_floating_widget_toggle_from_header_and_settings(qt_app, window):
    assert window._floating_monitor is None

    window.header.monitor_button.click()
    qt_app.processEvents()
    widget = window._floating_monitor
    assert widget is not None
    assert widget.isVisible() is True
    assert widget.is_sampling is True
    assert wait_for(qt_app, lambda: widget.cpu_row.value.text() != "در دسترس نیست", timeout=20)
    assert window.settings_page.widget_checkbox.isChecked()

    window.show_page("settings")
    qt_app.processEvents()
    window.settings_page.opacity_slider.setValue(65)
    qt_app.processEvents()
    saved = load_ui_settings(window.state.project_root)
    assert saved.widget_opacity == 0.65

    window.settings_page.widget_checkbox.setChecked(False)
    qt_app.processEvents()
    assert widget.isVisible() is False
    assert widget.is_sampling is False
    assert load_ui_settings(window.state.project_root).widget_visible is False

    window.settings_page.widget_checkbox.setChecked(True)
    qt_app.processEvents()
    assert widget.isVisible() is True
    assert widget.is_sampling is True


def test_ui_analysis_uses_real_ollama_http_transport(qt_app, window, sample_path, fake_ollama):
    window.state.settings.ollama_base_url = fake_ollama.start()
    window.state.settings.ollama_retries = 1
    window.state.set_models(True, fake_ollama.models)
    window.state.set_model("qwen3.5:9b")
    assert window.state.use_ai is True

    assert window._accept_file(sample_path) is True
    window._start_file("این فایل را تحلیل کن و خلاصه مدیریتی بده.")
    assert wait_for(qt_app, lambda: window.state.last_result is not None, timeout=90)

    result = window.state.last_result
    assert result.analysis.ai["mode"] == "model"
    assert result.analysis.ai["model"] == "qwen3.5:9b"
    assert fake_ollama.requests
    assert ("POST", "/api/chat") in fake_ollama.requests
