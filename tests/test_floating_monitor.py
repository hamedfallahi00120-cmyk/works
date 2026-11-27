"""Acceptance checks for the optional floating hardware widget.

The widget reuses the real ``HardwareMonitor`` through ``HardwareWorker``; these
tests start it offscreen and read the real psutil readings.  Nothing is stubbed.
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
    from PySide6.QtWidgets import QApplication, QSystemTrayIcon
except (ImportError, OSError) as exc:  # pragma: no cover - host without Qt
    pytest.skip(f"PySide6 is not usable here: {exc}", allow_module_level=True)

from local_office_ai.core.settings import load_settings  # noqa: E402
from local_office_ai.ui.desktop.floating_monitor import (  # noqa: E402
    FloatingMonitor,
    autostart_command,
    autostart_supported,
    is_autostart_enabled,
    set_autostart,
)
from local_office_ai.ui.desktop.hardware import HardwareSnapshot  # noqa: E402
from local_office_ai.ui.desktop.state import AppState  # noqa: E402
from local_office_ai.ui.desktop.theme import palette_for  # noqa: E402
from local_office_ai.ui.desktop.ui_settings import load_ui_settings  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def qt_app() -> QApplication:
    application = QApplication.instance() or QApplication([])
    yield application


def _project_root(tmp_path: Path) -> Path:
    root = tmp_path / "widget_project"
    for name in ("config", "data/input", "data/output", "logs"):
        (root / name).mkdir(parents=True, exist_ok=True)
    return root


@pytest.fixture()
def state(qt_app: QApplication, tmp_path: Path) -> AppState:
    root = _project_root(tmp_path)
    settings = load_settings(project_root=root)
    settings.ensure_directories()
    return AppState(settings, load_ui_settings(root))


def _wait_for(qt_app: QApplication, predicate, timeout: float = 20.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        qt_app.processEvents()
        if predicate():
            return True
        time.sleep(0.02)
    qt_app.processEvents()
    return bool(predicate())


def test_widget_is_frameless_always_on_top_and_reuses_one_monitor(state):
    widget = FloatingMonitor(state, palette_for("light"))
    flags = widget.windowFlags()
    assert flags & Qt.WindowType.WindowStaysOnTopHint
    assert flags & Qt.WindowType.FramelessWindowHint
    assert flags & Qt.WindowType.Tool

    source = (ROOT / "src/local_office_ai/ui/desktop/floating_monitor.py").read_text(encoding="utf-8")
    assert "nvidia-smi" not in source
    assert "subprocess" not in source
    assert "HardwareWorker" in source
    widget.shutdown()


def test_sampling_runs_only_while_visible(qt_app, state):
    widget = FloatingMonitor(state, palette_for("light"))
    assert widget.is_sampling is False

    widget.show_widget()
    assert _wait_for(qt_app, lambda: widget.cpu_row.value.text() != "در دسترس نیست")
    assert widget.is_sampling is True
    assert widget.ram_row.value.text() != "در دسترس نیست"
    assert load_ui_settings(state.project_root).widget_visible is True

    widget.hide_to_tray()
    qt_app.processEvents()
    assert widget.isVisible() is False
    assert widget.is_sampling is False
    assert load_ui_settings(state.project_root).widget_visible is False

    widget.show_widget()
    assert _wait_for(qt_app, lambda: widget.is_sampling is True)
    widget.shutdown()


def test_position_opacity_and_geometry_are_restored(qt_app, state):
    widget = FloatingMonitor(state, palette_for("light"))
    widget.show_widget()
    qt_app.processEvents()

    widget.move(140, 160)
    widget._persist_geometry()
    widget.set_opacity(0.6)

    saved = load_ui_settings(state.project_root)
    assert (saved.widget_x, saved.widget_y) == (140, 160)
    assert saved.widget_opacity == pytest.approx(0.6, abs=0.01)
    widget.shutdown()

    restored = FloatingMonitor(state, palette_for("light"))
    restored.show_widget()
    qt_app.processEvents()
    assert (restored.pos().x(), restored.pos().y()) == (140, 160)
    restored.shutdown()


def test_missing_gpu_is_reported_as_unavailable_not_zero(qt_app, state):
    widget = FloatingMonitor(state, palette_for("light"))
    widget.show_widget()
    qt_app.processEvents()
    widget._stop_sampling()
    widget._on_snapshot(HardwareSnapshot())
    assert widget.gpu_row.value.text() == "در دسترس نیست"
    assert "0" not in widget.gpu_row.value.text()
    widget.shutdown()


def test_autostart_is_windows_only(tmp_path: Path):
    if os.name == "nt":
        pytest.skip("this assertion covers the honest non-Windows branch")
    assert autostart_supported() is False
    assert set_autostart(True, tmp_path) is False
    assert is_autostart_enabled() is False
    command = autostart_command(tmp_path)
    assert "--floating-only" in command
    assert "--root" in command


def test_widget_visibility_preference_survives_application_shutdown(qt_app, state):
    state.set_widget_options(visible=False)
    widget = FloatingMonitor(state, palette_for("light"))
    widget.show_widget()
    qt_app.processEvents()
    assert state.ui.widget_visible is True

    widget.shutdown()
    qt_app.processEvents()
    assert widget.isVisible() is False
    assert widget.is_sampling is False
    assert state.ui.widget_visible is True


def test_tray_menu_is_owned_and_cleaned_up(qt_app, state, monkeypatch):
    monkeypatch.setattr(
        QSystemTrayIcon, "isSystemTrayAvailable", staticmethod(lambda: True)
    )
    widget = FloatingMonitor(state, palette_for("light"))
    widget.show_widget()
    qt_app.processEvents()
    assert widget.tray_icon is not None
    assert widget._tray_menu is not None
    assert widget._tray_menu.actions()

    widget.shutdown()
    qt_app.processEvents()
    assert widget.tray_icon is None
    assert widget._tray_menu is None


def test_explicit_switch_off_clears_visibility_preference(qt_app, state):
    state.set_widget_options(visible=False)
    widget = FloatingMonitor(state, palette_for("light"))
    widget.show_widget()
    qt_app.processEvents()
    assert state.ui.widget_visible is True

    widget.deactivate()
    qt_app.processEvents()
    assert widget.isVisible() is False
    assert state.ui.widget_visible is False
