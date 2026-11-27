"""Modern PySide6 desktop interface for Local Office AI."""

from __future__ import annotations

__all__ = ["main"]


def main(argv: list[str] | None = None) -> int:
    from local_office_ai.ui.desktop.app import main as _main

    return _main(argv)
