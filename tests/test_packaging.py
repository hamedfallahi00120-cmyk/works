"""Deliverable layout: one main launcher, V4.2 quarantined, sample in data/input."""

from __future__ import annotations

from pathlib import Path

from local_office_ai.core.settings import load_settings

ROOT = Path(__file__).resolve().parents[1]


def test_main_launchers_use_project_venv_only():
    for name in (
        "Run_Local_Office_AI.bat",
        "Run_Local_Office_AI_GUI.bat",
        "Run_Local_Office_AI_Modern.bat",
        "Run_Tests.bat",
    ):
        text = (ROOT / name).read_text(encoding="utf-8")
        assert "venv\\Scripts\\python.exe" in text
        assert "py -" not in text
        assert "py -m" not in text


def test_v42_files_are_in_legacy_not_root():
    assert not (ROOT / "Ollama_Excel_Assistant_V4_2.py").exists()
    assert (ROOT / "legacy" / "Ollama_Excel_Assistant_V4_2.py").is_file()
    assert (ROOT / "legacy" / "Run_Ollama_Excel_Assistant_V4_2.bat").is_file()


def test_sample_workbook_lives_in_data_input():
    assert (ROOT / "data" / "input" / "Book2_AI_V4_2_Monitored.xlsx").is_file()
    assert not (ROOT / "Book2_AI_V4_2_Monitored.xlsx").exists()


def test_settings_root_is_derived_from_project():
    settings = load_settings()
    assert settings.project_root == ROOT
    assert settings.synonyms_file == ROOT / "config" / "synonyms.json"
    assert settings.synonyms
