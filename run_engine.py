"""Portable launcher for the Local Office AI engine.

Run with the project virtual environment:
    F:\\Local-Office-AI\\venv\\Scripts\\python.exe run_engine.py --file <path>
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from local_office_ai.ui.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
