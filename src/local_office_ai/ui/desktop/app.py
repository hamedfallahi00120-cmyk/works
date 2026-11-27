"""Entry point for the modern PySide6 desktop interface."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from local_office_ai.core.settings import load_settings
from local_office_ai.ui.desktop.ui_settings import load_ui_settings


_TRANSLATORS: list = []


def _install_translations(application) -> None:
    """Load Qt's own Persian strings so built-in dialogs stop showing English labels."""
    from PySide6.QtCore import QLibraryInfo, QTranslator

    path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    for name in ("qtbase_fa", "qt_fa"):
        translator = QTranslator(application)
        if not translator.load(name, path):
            continue
        application.installTranslator(translator)
        _TRANSLATORS.append(translator)
        return


def _prepare_application(application, ui_settings):
    """Apply the saved theme, font and RTL direction before any window exists."""
    from PySide6.QtCore import Qt

    from local_office_ai.ui.desktop import fonts, icons
    from local_office_ai.ui.desktop.theme import apply_palette, build_stylesheet, palette_for

    _install_translations(application)
    if ui_settings.font_file and Path(ui_settings.font_file).is_file():
        fonts.load_font_file(ui_settings.font_file)
    palette = palette_for(ui_settings.theme)
    family = fonts.apply_font(application, ui_settings.font_family, ui_settings.font_base_size)
    icons.clear_cache()
    application.setStyleSheet(
        build_stylesheet(
            palette,
            family=family,
            base_size=ui_settings.font_base_size,
            heading_size=ui_settings.font_heading_size,
        )
    )
    apply_palette(application, palette)
    application.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    return palette


def main(argv: list[str] | None = None) -> int:
    try:
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QApplication
    except ImportError as exc:  # pragma: no cover - depends on the host install
        print(
            "رابط مدرن به PySide6 نیاز دارد و در این محیط پیدا نشد.\n"
            f"جزئیات: {exc}\n"
            "نصب: venv\\Scripts\\python.exe -m pip install \"PySide6>=6.6,<7\"\n"
            "یا از رابط قدیمی استفاده کنید: Run_Local_Office_AI_GUI.bat",
            file=sys.stderr,
        )
        return 1

    parser = argparse.ArgumentParser(add_help=True, description="Local Office AI — رابط مدرن")
    parser.add_argument("--root", help="ریشه پروژه (پیش‌فرض: محل همین بسته)")
    parser.add_argument(
        "--floating-only",
        action="store_true",
        help="فقط ویجت شناور پایش سخت‌افزار را اجرا کن (بدون پنجرهٔ اصلی)",
    )
    parser.add_argument(
        "--floating",
        action="store_true",
        help="پنجرهٔ اصلی را همراه ویجت شناور باز کن",
    )
    args, qt_args = parser.parse_known_args(argv)

    settings = load_settings(args.root)
    ui_settings = load_ui_settings(settings.project_root)

    application = QApplication([sys.argv[0], *qt_args])
    application.setApplicationName("Local Office AI")
    application.setOrganizationName("Local Office AI")
    application.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    application.setStyle("Fusion")

    palette = _prepare_application(application, ui_settings)

    if args.floating_only:
        from local_office_ai.ui.desktop.floating_monitor import FloatingMonitor
        from local_office_ai.ui.desktop.state import AppState

        state = AppState(settings, ui_settings)
        widget = FloatingMonitor(state, palette, standalone=True)
        widget.quit_requested.connect(application.quit)
        widget.show_widget()
        return application.exec()

    from local_office_ai.ui.desktop.main_window import MainWindow

    window = MainWindow(settings, ui_settings)
    window.show()
    if args.floating:
        window.show_floating_monitor()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
