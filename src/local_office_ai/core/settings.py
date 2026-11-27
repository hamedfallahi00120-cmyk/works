"""Central configuration.

All paths are derived from the project root, never hard-coded to a drive, so
the package runs unchanged from ``F:\\Local-Office-AI`` or from a USB stick.
A ``config/settings.json`` file may override any field.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, fields
from decimal import Decimal
from pathlib import Path
from typing import Any

DEFAULT_CONFIG_NAME = "settings.json"
ENV_ROOT = "LOCAL_OFFICE_AI_ROOT"


@dataclass
class Settings:
    project_root: Path
    input_dir: Path
    processed_dir: Path
    output_dir: Path
    temp_dir: Path
    log_dir: Path
    synonyms_file: Path | None = None
    synonyms: dict[str, str] = field(default_factory=dict)
    max_file_mb: int = 200
    max_processing_seconds: float = 1800.0
    max_sheets: int = 100
    max_rows_scanned: int = 200_000
    max_columns_scanned: int = 300
    ollama_base_url: str = "http://127.0.0.1:11434"
    ollama_timeout_seconds: float = 300.0
    ollama_connect_timeout_seconds: float = 8.0
    ollama_retries: int = 2
    ollama_num_ctx: int = 8192
    ollama_temperature: float = 0.1
    model_name_prefixes: tuple[str, ...] = ("qwen3.5", "qwen3", "qwen2.5")
    max_ai_context_chars: int = 9_000
    max_ai_prompt_chars: int = 11_000
    max_ai_sheets_per_chunk: int = 3
    ai_chunk_retry_ratio: float = 0.55
    max_ai_rows: int = 40
    max_header_rows: int = 3
    reconciliation_exact_threshold: float = 0.95
    reconciliation_probable_threshold: float = 0.70
    money_tolerance: Decimal = Decimal("0.01")
    outlier_mad_threshold: float = 3.5
    max_stats_rows: int = 20_000
    max_frequency_values: int = 40
    max_outliers: int = 120
    header_scan_rows: int = 25
    min_table_rows: int = 2
    min_table_columns: int = 2
    comfyui_base_url: str = "http://127.0.0.1:8188"
    comfyui_timeout_seconds: float = 60.0
    comfyui_generation_timeout_seconds: float = 1200.0
    comfyui_default_checkpoint: str = ""
    comfyui_steps: int = 20
    comfyui_cfg: float = 7.0
    comfyui_width: int = 512
    comfyui_height: int = 512
    monitor_interval_seconds: float = 1.0
    gpu_interval_seconds: float = 2.5
    enable_gpu_monitoring: bool = True
    # Resource sampling is developer diagnostics only; normal runs skip it so the
    # delivered workbook never carries CPU/RAM/GPU sheets.
    enable_resource_monitoring: bool = False
    # "clean" writes only the sheets the request needs; "audit" adds the
    # internal traceability/performance sheets for development.
    report_profile: str = "clean"
    # Delivery format is resolved from --format, the output suffix or the
    # request text; this value is only the fallback.
    report_format: str = "xlsx"
    docx_template: str = ""
    log_level: str = "INFO"
    keep_intermediate_json: bool = True
    extra: dict[str, Any] = field(default_factory=dict)

    def ensure_directories(self) -> None:
        for directory in (
            self.input_dir,
            self.processed_dir,
            self.output_dir,
            self.temp_dir,
            self.log_dir,
        ):
            directory.mkdir(parents=True, exist_ok=True)

    def output_path_for(
        self,
        source: Path,
        suffix: str = "_AI_Intelligence",
        extension: str = ".xlsx",
    ) -> Path:
        """Return a collision-free output path that never equals the input."""

        ext = extension if extension.startswith(".") else f".{extension}"
        stem = source.stem
        candidate = self.output_dir / f"{stem}{suffix}{ext}"
        counter = 1
        while candidate.exists() or candidate.resolve() == source.resolve():
            candidate = self.output_dir / f"{stem}{suffix}_{counter}{ext}"
            counter += 1
        return candidate

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for item in fields(self):
            value = getattr(self, item.name)
            if isinstance(value, Path):
                data[item.name] = str(value)
            elif isinstance(value, Decimal):
                data[item.name] = str(value)
            elif isinstance(value, tuple):
                data[item.name] = list(value)
            else:
                data[item.name] = value
        return data


_PATH_FIELDS = {
    "project_root",
    "input_dir",
    "processed_dir",
    "output_dir",
    "temp_dir",
    "log_dir",
    "synonyms_file",
}


def default_root() -> Path:
    """Project root: env override, else two levels above this file's package."""
    override = os.environ.get(ENV_ROOT)
    if override:
        return Path(override).expanduser().resolve()
    return Path(__file__).resolve().parents[3]


def load_settings(
    project_root: Path | str | None = None,
    config_path: Path | str | None = None,
) -> Settings:
    root = Path(project_root).expanduser().resolve() if project_root else default_root()
    config_file = Path(config_path) if config_path else root / "config" / DEFAULT_CONFIG_NAME
    raw: dict[str, Any] = {}
    if config_file.is_file():
        with config_file.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)
        if not isinstance(raw, dict):
            raise ValueError(f"config file must contain a JSON object: {config_file}")

    data_dir = Path(raw.pop("data_dir", root / "data"))
    if not data_dir.is_absolute():
        data_dir = (root / data_dir).resolve()

    values: dict[str, Any] = {
        "project_root": root,
        "input_dir": data_dir / "input",
        "processed_dir": data_dir / "processed",
        "output_dir": data_dir / "output",
        "temp_dir": data_dir / "temp",
        "log_dir": root / "logs",
        "synonyms_file": root / "config" / "synonyms.json",
    }
    for key, value in raw.items():
        if key in _PATH_FIELDS:
            path = Path(value)
            values[key] = path if path.is_absolute() else (root / path).resolve()
        elif key == "money_tolerance":
            values[key] = Decimal(str(value))
        elif key == "model_name_prefixes":
            values[key] = tuple(str(item) for item in value)
        else:
            values[key] = value

    # Synonyms are a matching aid only; the source text is never rewritten.
    from local_office_ai.excel.normalization import load_synonyms

    values["synonyms"] = load_synonyms(values.get("synonyms_file"))

    known = {item.name for item in fields(Settings)}
    unknown = {key: value for key, value in values.items() if key not in known}
    settings = Settings(**{key: value for key, value in values.items() if key in known})
    settings.extra.update(unknown)
    settings.ensure_directories()
    return settings
