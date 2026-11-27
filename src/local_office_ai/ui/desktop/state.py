"""Shared application state for the desktop window."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, Signal

from local_office_ai.core.settings import Settings
from local_office_ai.ui.desktop.ui_settings import (
    UiSettings,
    effective_engine_settings,
    save_ui_settings,
)

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
SUPPORTED_INPUT_SUFFIXES = {
    ".xlsx", ".xlsm", ".csv", ".docx", ".pptx", ".pdf", ".txt",
    ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp",
}


def validate_workbook(path: Path) -> tuple[bool, str]:
    """Cheap real check: an xlsx/xlsm is a zip that contains xl/workbook.xml."""
    target = Path(path)
    if not target.exists():
        return False, "فایل پیدا نشد."
    if not target.is_file():
        return False, "مسیر انتخاب‌شده فایل نیست."
    if target.suffix.lower() not in SUPPORTED_SUFFIXES:
        return False, "فقط فایل‌های Excel با پسوند xlsx یا xlsm پشتیبانی می‌شوند."
    if target.stat().st_size == 0:
        return False, "فایل خالی است."
    try:
        with zipfile.ZipFile(target) as archive:
            names = set(archive.namelist())
    except (zipfile.BadZipFile, OSError):
        return False, "فایل، بسته Excel معتبری نیست یا خراب است."
    if "xl/workbook.xml" not in names:
        return False, "ساختار داخلی فایل Excel معتبر نیست."
    return True, ""


def validate_input_file(path: Path) -> tuple[bool, str]:
    """Validate supported administrative inputs without fully parsing them."""
    target = Path(path).expanduser()
    if target.suffix.lower() in SUPPORTED_SUFFIXES:
        return validate_workbook(target)
    if not target.exists():
        return False, "فایل پیدا نشد."
    if not target.is_file():
        return False, "مسیر انتخاب‌شده فایل نیست."
    if target.suffix.lower() not in SUPPORTED_INPUT_SUFFIXES:
        return False, "این قالب فایل پشتیبانی نمی‌شود."
    try:
        if target.stat().st_size <= 0:
            return False, "فایل خالی است."
        if target.suffix.lower() in {".docx", ".pptx"}:
            with zipfile.ZipFile(target) as archive:
                names = set(archive.namelist())
                expected = "word/document.xml" if target.suffix.lower() == ".docx" else "ppt/presentation.xml"
                if expected not in names:
                    return False, "ساختار داخلی سند معتبر نیست."
        elif target.suffix.lower() == ".pdf":
            with target.open("rb") as handle:
                if handle.read(5) != b"%PDF-":
                    return False, "ساختار PDF معتبر نیست."
    except (OSError, zipfile.BadZipFile):
        return False, "فایل خراب است یا قابل خواندن نیست."
    return True, ""


class AppState(QObject):
    source_changed = Signal()
    folder_changed = Signal()
    model_changed = Signal()
    ollama_changed = Signal()
    analysis_changed = Signal()
    folder_result_changed = Signal()
    settings_changed = Signal()
    busy_changed = Signal()

    def __init__(self, settings: Settings, ui: UiSettings, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.base_settings = settings
        self.project_root = settings.project_root
        self.ui = ui
        self.source: Path | None = None
        self.sources: list[Path] = []
        self.folder: Path | None = None
        self.ollama_online = False
        self.ollama_error = ""
        self.models: list[dict[str, Any]] = []
        self.model = ui.model
        self.busy = False
        self.last_result: Any = None
        self.last_folder_result: Any = None
        self.last_result_kind = ""
        self._effective: Settings | None = None

    # ------------------------------------------------------------------ #
    @property
    def settings(self) -> Settings:
        if self._effective is None:
            self._effective = effective_engine_settings(self.base_settings, self.ui)
        return self._effective

    def invalidate_settings(self) -> None:
        self._effective = None

    def save(self) -> None:
        save_ui_settings(self.project_root, self.ui)

    # ------------------------------------------------------------------ #
    def set_source(self, path: Path | None) -> None:
        self.set_sources([path] if path is not None else [])

    def set_sources(self, paths: list[Path] | tuple[Path, ...]) -> None:
        unique: list[Path] = []
        seen: set[str] = set()
        for item in paths:
            resolved = Path(item).expanduser().resolve()
            key = str(resolved).casefold()
            if key not in seen:
                unique.append(resolved)
                seen.add(key)
        self.sources = unique
        self.source = unique[0] if len(unique) == 1 else None
        self.folder = None
        self.ui.last_source = str(unique[0]) if len(unique) == 1 else ""
        self.ui.last_folder = ""
        self.save()
        self.source_changed.emit()

    def set_folder(self, path: Path | None) -> None:
        self.folder = Path(path).expanduser().resolve() if path else None
        if self.folder is not None:
            self.ui.last_folder = str(self.folder)
            self.ui.last_source = ""
            self.source = None
            self.sources = []
        else:
            self.ui.last_folder = ""
        self.save()
        self.folder_changed.emit()
        self.source_changed.emit()

    def set_models(self, online: bool, models: list[dict[str, Any]], error: str = "") -> None:
        self.ollama_online = bool(online)
        self.ollama_error = error
        self.models = list(models or [])
        if self.models:
            names = [item["name"] for item in self.models]
            if self.model not in names:
                self.model = self._preferred_name(names)
                self.ui.model = self.model
                self.save()
        self.ollama_changed.emit()
        self.model_changed.emit()

    def set_model(self, name: str) -> None:
        self.model = name or ""
        self.ui.model = self.model
        self.save()
        self.model_changed.emit()

    def model_names(self) -> list[str]:
        return [str(item.get("name", "")) for item in self.models if item.get("name")]

    @staticmethod
    def _preferred_name(names: list[str]) -> str:
        for prefix in ("qwen3.5", "qwen3", "qwen2.5", "qwen"):
            for name in names:
                if name.lower().startswith(prefix):
                    return name
        return names[0]

    # ------------------------------------------------------------------ #
    @property
    def use_ai(self) -> bool:
        return bool(self.ollama_online and self.model)

    @property
    def analysis_model(self) -> str | None:
        return self.model if self.use_ai else None

    def set_busy(self, busy: bool) -> None:
        self.busy = bool(busy)
        self.busy_changed.emit()

    def set_result(self, result: Any) -> None:
        self.last_result = result
        self.last_result_kind = "file"
        self.analysis_changed.emit()

    def set_folder_result(self, result: Any) -> None:
        self.last_folder_result = result
        self.last_result_kind = "folder"
        self.folder_result_changed.emit()

    # ------------------------------------------------------------------ #
    def update_paths(self, *, input_dir: str | None = None, output_dir: str | None = None) -> None:
        if input_dir is not None:
            self.ui.input_dir = input_dir
        if output_dir is not None:
            self.ui.output_dir = output_dir
        self.invalidate_settings()
        self.settings.ensure_directories()
        self.save()
        self.settings_changed.emit()

    def set_theme(self, theme: str) -> None:
        self.ui.theme = theme
        self.save()
        self.settings_changed.emit()

    def set_font(self, family: str, base_size: int, heading_size: int) -> None:
        self.ui.font_family = family
        self.ui.font_base_size = int(base_size)
        self.ui.font_heading_size = int(heading_size)
        self.ui.normalize()
        self.save()
        self.settings_changed.emit()

    def set_hardware_options(self, *, gpu_enabled: bool, interval: float) -> None:
        self.ui.hardware_gpu_enabled = bool(gpu_enabled)
        self.ui.hardware_interval = float(interval)
        self.ui.normalize()
        self.save()
        self.settings_changed.emit()

    def set_widget_options(
        self,
        *,
        visible: bool | None = None,
        position: tuple[int, int] | None = None,
        opacity: float | None = None,
        autostart: bool | None = None,
    ) -> None:
        if visible is not None:
            self.ui.widget_visible = bool(visible)
        if position is not None:
            self.ui.widget_x, self.ui.widget_y = int(position[0]), int(position[1])
        if opacity is not None:
            self.ui.widget_opacity = float(opacity)
        if autostart is not None:
            self.ui.widget_autostart = bool(autostart)
        self.ui.normalize()
        self.save()

    def set_font_file(self, path: str) -> None:
        self.ui.font_file = path
        self.save()
        self.settings_changed.emit()

    def set_last_request(self, text: str) -> None:
        self.ui.last_request = text
        self.save()
