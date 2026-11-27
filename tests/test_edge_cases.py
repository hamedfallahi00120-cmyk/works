"""Edge cases: protected/empty sheets, inconsistent cells, timeouts, logging."""

from __future__ import annotations

import pytest

from local_office_ai.ai.ollama_client import OllamaClient
from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.exceptions import AIServiceError


def test_protected_sheet_is_read_and_reported(make_workbook, settings):
    path = make_workbook(
        "protected.xlsx",
        {"S": [["نام", "مبلغ"], ["الف", 100], ["ب", 200]]},
        post=lambda workbook: setattr(workbook["S"].protection, "sheet", True),
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    sheet = analysis.meta.sheets[0]
    assert sheet.protected is True
    assert any(item.category == "protected-sheet" for item in analysis.anomalies)
    assert analysis.primary_table is not None


def test_inconsistent_numeric_cell_is_reported_not_guessed(make_workbook, settings):
    path = make_workbook(
        "inconsistent.xlsx",
        {
            "S": [
                ["بانک", "مبلغ وام", "مبلغ هر قسط", "تعداد کل قسط", "تاریخ آخرین قسط"],
                ["بانک الف", 1000, 100, 10, "1404/01/01"],
                ["بانک ب", "نامشخص", 200, 10, "1404/02/01"],
                ["بانک ج", 3000, 300, 10, "1404/03/01"],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    assert len(analysis.loans) == 3
    missing = [loan for loan in analysis.loans if loan.principal is None]
    assert len(missing) == 1
    assert any("principal" in issue for issue in missing[0].issues)
    # The invalid row is excluded from the sum instead of being treated as zero.
    values = {item.name: item.value for item in analysis.calculations}
    assert values["principal_sum"] == 4000
    assert any(item.category == "missing-principal" for item in analysis.anomalies)


def test_empty_sheet_workbook(make_workbook, settings):
    path = make_workbook("empty_sheet.xlsx", {"S": []})
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    assert analysis.primary_table is None
    assert analysis.meta.sheets[0].max_row == 0
    assert any("هیچ داده" in note for note in analysis.meta.sheets[0].suspicious)


def test_empty_cells_do_not_break_extraction(make_workbook, settings):
    path = make_workbook(
        "sparse.xlsx",
        {
            "S": [
                ["بانک", "مبلغ وام", "مبلغ هر قسط", "تعداد کل قسط"],
                ["بانک الف", None, 100, 10],
                ["بانک ب", 2000, None, 10],
                [None, None, None, None],
                ["بانک ج", 3000, 300, None],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    assert len(analysis.loans) == 3
    assert all(loan.issues for loan in analysis.loans)


def test_ollama_timeout_raises_service_error(settings, fake_ollama):
    settings.ollama_base_url = fake_ollama.start()
    settings.ollama_timeout_seconds = 0.05
    settings.ollama_retries = 0
    fake_ollama.chat_delay = 0.4
    with pytest.raises(AIServiceError):
        OllamaClient(settings).chat("qwen3.5:9b", [{"role": "user", "content": "x"}])


def test_log_file_is_written(settings, sample_path, tmp_path):
    coordinator = Coordinator(settings, console=False)
    coordinator.run(sample_path, use_ai=False, output=tmp_path / "logged.xlsx")
    log_path = coordinator.log_path
    assert log_path.exists()
    content = log_path.read_text(encoding="utf-8")
    assert "workbook loaded" in content
    assert "report written" in content


def test_ai_failure_does_not_block_report(settings, sample_path, tmp_path):
    settings.ollama_base_url = "http://127.0.0.1:9"
    settings.ollama_retries = 0
    result = Coordinator(settings, console=False).run(
        sample_path, use_ai=True, output=tmp_path / "no_model.xlsx"
    )
    assert result.status == "completed"
    assert result.analysis is not None
    assert result.analysis.ai["mode"] == "offline"
    assert result.analysis.ai["error"]
    assert result.analysis.ai["response"]["summary"]
