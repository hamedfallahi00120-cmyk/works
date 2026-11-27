"""Logging setup: UTF-8 rotating file log plus a concise console handler."""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

_FILE_HANDLER: RotatingFileHandler | None = None
_STREAM_HANDLER: logging.StreamHandler | None = None
_FORMAT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"


def configure_logging(
    log_dir: Path,
    level: str = "INFO",
    filename: str = "local_office_ai.log",
    console: bool = True,
) -> Path:
    """Configure root logging and return the active log file path.

    Repeated calls are safe: the file handler is re-pointed when the target log
    file changes (for example when a different project root is used).
    """
    global _FILE_HANDLER, _STREAM_HANDLER
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / filename

    root = logging.getLogger()
    numeric_level = getattr(logging, str(level).upper(), logging.INFO)
    root.setLevel(numeric_level)

    formatter = logging.Formatter(_FORMAT, datefmt="%Y-%m-%d %H:%M:%S")

    if _FILE_HANDLER is not None and Path(_FILE_HANDLER.baseFilename) != log_path:
        root.removeHandler(_FILE_HANDLER)
        _FILE_HANDLER.close()
        _FILE_HANDLER = None
    if _FILE_HANDLER is None:
        _FILE_HANDLER = RotatingFileHandler(
            log_path, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
        )
        _FILE_HANDLER.setFormatter(formatter)
        root.addHandler(_FILE_HANDLER)
    _FILE_HANDLER.setLevel(numeric_level)

    if console and _STREAM_HANDLER is None:
        _STREAM_HANDLER = logging.StreamHandler(stream=sys.stderr)
        _STREAM_HANDLER.setFormatter(formatter)
        root.addHandler(_STREAM_HANDLER)
    if _STREAM_HANDLER is not None:
        _STREAM_HANDLER.setLevel(numeric_level)

    logging.getLogger("openpyxl").setLevel(logging.WARNING)
    return log_path


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)
