"""Reference compatibility test against the V4.2 sample workbook.

The expected figures come from the sample, not from hard-coded engine logic:
the engine has to derive them from the workbook.
"""

from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.core.models import AuditStatus, MatchStatus, TableKind
from local_office_ai.excel.reader import file_sha256


def test_reference_numbers_from_sample(settings, sample_path):
    from local_office_ai.app.workflow import ExcelIntelligenceWorkflow

    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)
    assert analysis.primary_table is not None
    assert analysis.primary_table.kind is TableKind.LOAN
    assert analysis.primary_table.sheet == "Sheet1"
    assert analysis.primary_table.address == "B2:H8"

    values = {item.name: item.value for item in analysis.calculations}
    assert len(analysis.loans) == 6
    assert values["loan_count"] == 6
    assert values["principal_sum"] == Decimal("325000000")
    assert values["installments_sum"] == Decimal("14749000")
    assert values["obligations_sum"] == Decimal("18999000")
    assert values["contract_repayment_sum"] == Decimal("347576000")
    assert values["obligation_gap"] == Decimal("4250000")


def test_reference_audits_and_reconciliation(settings, sample_path):
    from local_office_ai.app.workflow import ExcelIntelligenceWorkflow

    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)
    audits = {audit.measure: audit for audit in analysis.audits}
    assert audits["principal"].declared == Decimal("325000000")
    assert audits["principal"].computed == Decimal("325000000")
    assert audits["principal"].difference == Decimal("0")
    assert audits["principal"].status is AuditStatus.VERIFIED
    assert audits["installments"].declared == Decimal("14749000")
    assert audits["installments"].difference == Decimal("0")
    assert audits["installments"].status is AuditStatus.VERIFIED

    by_name = {item.loan_name or item.obligation_name: item for item in analysis.reconciliation}
    assert by_name["بانک مهر ایران"].status is MatchStatus.MATCH
    assert by_name["بانک مهر ایران"].loan_total == Decimal("13149000")
    assert by_name["صندق امین"].status is MatchStatus.MATCH
    assert by_name["صندق رفاه کهنز"].status is MatchStatus.OUTSIDE_LOAN_TABLE
    assert by_name["صندق رفاه کهنز"].obligation_total == Decimal("4250000")


def test_full_run_is_reproducible_and_preserves_input(settings, sample_path, tmp_path):
    before = file_sha256(sample_path)
    coordinator = Coordinator(settings, console=False)
    output = tmp_path / "out" / "result.xlsx"
    result = coordinator.run(sample_path, use_ai=False, output=output)
    assert result.status == "completed"
    assert Path(result.output_path).exists()
    assert Path(result.output_path).resolve() != sample_path.resolve()
    assert file_sha256(sample_path) == before

    # Machine-readable intermediate goes to data/processed, not next to the report.
    sidecar = settings.processed_dir / f"{Path(result.output_path).stem}.json"
    assert sidecar.exists()
    assert not Path(result.output_path).with_suffix(".json").exists()
    payload = json.loads(sidecar.read_text(encoding="utf-8"))
    calculations = {item["name"]: item["value"] for item in payload["analysis"]["calculations"]}
    assert calculations["principal_sum"] == "325000000"
    assert calculations["installments_sum"] == "14749000"
    assert calculations["obligations_sum"] == "18999000"
    assert calculations["contract_repayment_sum"] == "347576000"
    assert payload["analysis"]["meta"]["sha256"] == before
    assert payload["raw_cells"], "raw cell values must be preserved in the processed JSON"
    first_cell = payload["raw_cells"][0]["cells"][0]
    assert {"raw", "value", "kind"} <= set(first_cell)


def test_run_with_model_uses_fake_ollama(ollama_settings, sample_path, tmp_path):
    coordinator = Coordinator(ollama_settings, console=False)
    result = coordinator.run(sample_path, use_ai=True, output=tmp_path / "model.xlsx")
    assert result.analysis is not None
    assert result.analysis.ai["mode"] == "model"
    assert result.analysis.ai["model"] == "qwen3.5:9b"
    assert result.analysis.ai["response"]["is_grounded"] is True
    assert result.performance["model"]["output_tokens"] == 120
    assert result.performance["model"]["tokens_per_second"] > 0


def test_output_path_collision_is_avoided(settings, sample_path, tmp_path):
    coordinator = Coordinator(settings, console=False)
    first = coordinator.run(sample_path, use_ai=False, output=tmp_path / "same.xlsx")
    second = coordinator.run(sample_path, use_ai=False)
    assert first.output_path != second.output_path
    assert Path(second.output_path).exists()
