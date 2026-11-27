"""The deployment manifest must ship the modern entry points and no user data."""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REQUIRED = {"run_ui.py", "Run_Local_Office_AI_Modern.bat"}
FORBIDDEN = ("data\\input", "data/input", "reports", "reference", "venv", "logs")


def _bat_text() -> str:
    return (ROOT / "Deploy_Update.bat").read_text(encoding="utf-8")


def _bat_manifest(text: str) -> set[str]:
    """Collect the names from every ``for %%X in (...)`` copy loop."""
    names: set[str] = set()
    for match in re.finditer(r"for %%[A-Z] in \(([^)]*)\) do", text):
        names.update(token.strip() for token in match.group(1).split() if token.strip())
    return names


def test_bat_ships_modern_entry_points():
    manifest = _bat_manifest(_bat_text())
    missing = REQUIRED - manifest
    assert not missing, f"Deploy_Update.bat does not copy: {sorted(missing)}"
    assert "run_engine.py" in manifest
    assert "src" in manifest
    assert "tests" in manifest


def test_bat_never_copies_user_data_or_environment():
    text = _bat_text()
    manifest = _bat_manifest(text)
    for token in ("data\\input", "data/input", "reports", "reference", "venv", "logs"):
        assert token not in manifest, f"Deploy_Update.bat must not copy {token}"
    lowered = text.lower()
    assert "del " not in lowered
    assert "rmdir" not in lowered
    assert "rd /s" not in lowered


def test_bat_does_not_overwrite_itself_while_running():
    assert "Deploy_Update.bat" not in _bat_manifest(_bat_text())


def test_bat_verifies_entry_points_after_copy():
    text = _bat_text()
    assert "Entry points verified" in text
    assert "missing after copy" in text


def test_manual_manifest_in_deploy_doc_is_complete():
    text = (ROOT / "DEPLOY.md").read_text(encoding="utf-8")
    section = text.split("## ۳. استقرار دستی", 1)
    assert len(section) == 2, "the manual deployment section is missing"
    block = section[1].split("```", 2)[1]
    for name in REQUIRED | {"run_engine.py", "src\\", "tests\\", "config\\settings.json"}:
        assert name in block, f"DEPLOY.md manual list is missing {name}"
    for token in ("data\\input", "data\\output", "reports\\", "reference\\"):
        assert token not in block, f"DEPLOY.md manual list must not copy {token}"
