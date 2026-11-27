"""Allow ``python -m local_office_ai`` to run the CLI."""

from __future__ import annotations

import sys

from local_office_ai.ui.cli import main

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
