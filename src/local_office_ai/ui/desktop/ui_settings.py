"""Desktop-only preferences, kept apart from the engine configuration.

``config/settings.json`` belongs to the engine and is never rewritten by the
interface.  This file stores the operator's appearance and window choices; the
selected input/output folders are applied to an in-memory copy of the engine
settings for each run.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, fields, replace
from pathlib import Path
from typing import Any

from local_office_ai.core.settings import Settings

FILE_NAME = "ui_settings.json"
THEMES = ("light", "dark")


@dataclass
class UiSettings:
    theme: str = "light"
    font_family: str = ""
    font_base_size: int = 13
    font_heading_size: int = 19
    font_file: str = ""
    model: str = ""
    input_dir: str = ""
    output_dir: str = ""
    hardware_gpu_enabled: bool = True
    hardware_interval: float = 2.0
    widget_visible: bool = False
    widget_x: int = -1
    widget_y: int = -1
    widget_opacity: float = 0.94
    widget_autostart: bool = False
    last_request: str = ""
    last_source: str = ""
    last_folder: str = ""

    def __post_init__(self) -> None:
        self.normalize()

    def normalize(self) -> None:
        if self.theme not in THEMES:
            self.theme = "light"
        self.font_base_size = max(9, min(22, int(self.font_base_size or 13)))
        self.font_heading_size = max(13, min(32, int(self.font_heading_size or 19)))
        self.hardware_interval = max(1.0, min(10.0, float(self.hardware_interval or 2.0)))
        self.widget_opacity = max(0.35, min(1.0, float(self.widget_opacity or 0.94)))
        self.widget_x = int(self.widget_x) if self.widget_x is not None else -1
        self.widget_y = int(self.widget_y) if self.widget_y is not None else -1
        self.widget_visible = bool(self.widget_visible)
        self.widget_autostart = bool(self.widget_autostart)

    def to_dict(self) -> dict[str, Any]:
        return {item.name: getattr(self, item.name) for item in fields(self)}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "UiSettings":
        known = {item.name for item in fields(cls)}
        values = {key: value for key, value in (data or {}).items() if key in known}
        try:
            return cls(**values)
        except (TypeError, ValueError):
            return cls()


def default_path(project_root: Path) -> Path:
    return Path(project_root) / "config" / FILE_NAME


def load_ui_settings(project_root: Path) -> UiSettings:
    path = default_path(project_root)
    if not path.is_file():
        return UiSettings()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return UiSettings()
    return UiSettings.from_dict(data if isinstance(data, dict) else {})


def save_ui_settings(project_root: Path, settings: UiSettings) -> Path:
    settings.normalize()
    path = default_path(project_root)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp")
    temp.write_text(
        json.dumps(settings.to_dict(), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temp.replace(path)
    return path


def effective_engine_settings(base: Settings, ui: UiSettings) -> Settings:
    """Apply the chosen folders to the engine settings without touching JSON."""
    input_dir = _paths_or_default(ui.input_dir, base.project_root, base.input_dir)
    output_dir = _paths_or_default(ui.output_dir, base.project_root, base.output_dir)
    merged = replace(base, input_dir=input_dir, output_dir=output_dir)
    merged.ensure_directories()
    return merged


def _paths_or_default(value: str, root: Path, default: Path) -> Path:
    if not value:
        return default
    candidate = Path(value).expanduser()
    if not candidate.is_absolute():
        candidate = (root / candidate).resolve()
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        return candidate
    except OSError:
        return default
