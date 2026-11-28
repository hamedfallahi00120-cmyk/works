"""Prompt 05A: one quality-gate model and a complete provenance ledger.

The tests use real .xlsx workbooks built by openpyxl.  No analyzer, cell reader
or report writer is mocked: the gate evaluates the actual objects the engine
produces, and the workbook assertions reopen the delivered files.
"""

from __future__ import annotations

from dataclasses import replace
from decimal import Decimal

import openpyxl
import pytest

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.folder_workflow import FolderWorkflow
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.models import AnalysisResult
from local_office_ai.quality import (
    STAGE_ORDER,
    GateStage,
    GateStatus,
    ProvenanceLedger,
    ProvenanceRecord,
    QualityGate,
    worst_status,
)
from local_office_ai.reporting import OutputFormat, ReportGenerationResult, ReportSpecification


def _analyze(settings, path):
    return ExcelIntelligenceWorkflow(settings).analyze_only(path)


@pytest.fixture()
def clean_values_workbook(make_workbook):
    """A complete table with a real declared total and no data-quality issue."""

    return make_workbook(
        "quality_clean.xlsx",
        {
            "داده": [
                ["ردیف", "شرح", "مبلغ", "تاریخ", "دسته"],
                [1, "اجاره", 12500000, "1404/01/15", "اجاره"],
                [2, "برق", 1250000, "1404/02/03", "انرژی"],
                [3, "اینترنت", 850000, "1404/02/10", "ارتباطات"],
                ["جمع", None, "=SUM(C2:C4)", None, None],
            ]
        },
    )


@pytest.fixture()
def warning_value_workbook(make_workbook):
    """One negative value, which the quality analyzer keeps and flags."""

    return make_workbook(
        "quality_warning.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", -50],
                ["آب", 300],
            ]
        },
    )


def _assert_all_derived_records_have_sources(analysis) -> None:
    ledger = analysis.provenance
    assert ledger is not None
    for record in ledger.important_records:
        if not record.derived:
            continue
        assert record.sources, f"derived record without source: {record.key}"
        for source in record.sources:
            assert source.filename
            assert source.sheet
            assert source.cell


# --------------------------------------------------------------------------- #
# 1. Complete provenance
# --------------------------------------------------------------------------- #
def test_complete_provenance_keeps_every_required_field(settings, clean_values_workbook):
    analysis = _analyze(settings, clean_values_workbook)

    ledger = analysis.provenance
    assert ledger is not None
    assert ledger.status is GateStatus.PASS
    assert ledger.filename == clean_values_workbook.name
    assert ledger.counts()["important"] > 0
    assert ledger.counts()["incomplete"] == 0

    record = ledger.record("stat.داده.C.total")
    assert record is not None
    assert record.filename == clean_values_workbook.name
    assert record.sheet == "داده"
    assert record.block
    assert record.cell == "C2:C4"
    assert record.original_label == "مبلغ"
    assert "جمع" in record.normalized_label
    assert record.normalized_value == "14600000"
    assert record.operation == "جمع"
    assert record.formula
    assert record.derived is True
    assert record.complete is True
    assert record.missing_fields() == []
    assert record.sources
    assert record.sources[0].cell == "C2:C4"
    assert "12500000" in record.sources[0].original_value

    payload = record.to_dict()
    assert {
        "filename",
        "sheet",
        "block",
        "cell",
        "original_label",
        "normalized_label",
        "original_value",
        "normalized_value",
        "period",
        "unit",
        "operation",
        "formula",
        "sources",
        "validation_status",
    } <= set(payload)
    assert payload["complete"] is True


# --------------------------------------------------------------------------- #
# 2. Missing cell references
# --------------------------------------------------------------------------- #
def test_missing_cell_reference_fails_provenance_and_blocks(
    settings, clean_values_workbook
):
    analysis = _analyze(settings, clean_values_workbook)
    ledger = ProvenanceLedger(
        filename=clean_values_workbook.name,
        records=[
            ProvenanceRecord(
                key="broken.derived",
                filename=clean_values_workbook.name,
                normalized_value="10",
                operation="جمع",
                derived=True,
                important=True,
            )
        ],
    )

    report = QualityGate().evaluate(
        analysis,
        source_path=clean_values_workbook,
        provenance=ledger,
    )

    stage = report.stage(GateStage.PROVENANCE_COMPLETENESS)
    assert stage is not None
    assert stage.status is GateStatus.FAIL
    assert stage.blocking is True
    assert report.blocked is True
    assert report.can_claim_valid_report is False
    assert "sheet" in ledger.missing_fields()
    assert "cell" in ledger.missing_fields()
    assert "sources" in ledger.missing_fields()
    assert [item.key for item in ledger.derived_without_sources] == ["broken.derived"]
    assert any(
        check.code == "provenance-derived-sources" for check in stage.checks
    )


# --------------------------------------------------------------------------- #
# 3. Derived values with source references
# --------------------------------------------------------------------------- #
def test_derived_values_point_back_to_source_cells(settings, clean_values_workbook):
    analysis = _analyze(settings, clean_values_workbook)
    _assert_all_derived_records_have_sources(analysis)

    ledger = analysis.provenance
    assert ledger is not None
    reconciliation = next(
        item
        for item in ledger.important_records
        if item.key.startswith("reconciliation.")
    )
    assert reconciliation.derived is True
    assert reconciliation.normalized_value == "14600000"
    assert {source.cell for source in reconciliation.sources} >= {"C2", "C3", "C4"}
    values = {source.cell: source.original_value for source in reconciliation.sources}
    assert values["C2"] == "12500000"
    assert values["C3"] == "1250000"
    # The formula total has no cached value in a workbook never opened by Excel.
    assert values["C5"] == ""

    period_values = [
        record
        for record in ledger.important_records
        if record.key.startswith("calculation.generic_sum") or record.key.startswith("stat.")
    ]
    assert period_values
    for record in period_values:
        assert record.sources[0].sheet == "داده"
        assert record.cell


# --------------------------------------------------------------------------- #
# 4. Failed extraction
# --------------------------------------------------------------------------- #
def test_failed_extraction_is_a_blocking_failure(settings, clean_values_workbook):
    analysis = _analyze(settings, clean_values_workbook)
    assert isinstance(analysis, AnalysisResult)
    broken = replace(analysis, extraction=None)

    report = QualityGate().evaluate(broken, source_path=clean_values_workbook)

    stage = report.stage(GateStage.EXTRACTION_COMPLETENESS)
    assert stage is not None
    assert stage.status is GateStatus.FAIL
    assert stage.blocking is True
    assert report.blocked is True
    assert report.can_claim_valid_report is False
    assert any(check.code == "extraction-missing" for check in stage.checks)
    assert report.status is GateStatus.FAIL


# --------------------------------------------------------------------------- #
# 5. Statistical validation failure
# --------------------------------------------------------------------------- #
def test_statistical_validation_failure_is_reported(settings, make_workbook):
    path = make_workbook(
        "wrong_declared_total.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
                ["آب", 300],
                ["جمع کل", 700],
            ]
        },
    )
    analysis = _analyze(settings, path)

    report = QualityGate().evaluate(analysis, source_path=path)
    stage = report.stage(GateStage.STATISTICAL_VALIDITY)
    assert stage is not None
    assert stage.status is GateStatus.FAIL
    assert report.status is GateStatus.FAIL
    assert report.is_complete_success is False
    assert any(check.status is GateStatus.FAIL for check in stage.checks)
    assert stage.evidence["verified_reconciliations"] == 1
    # A documented mismatch is a failure of the figure, not a broken deliverable.
    assert stage.blocking is False
    assert report.can_claim_valid_report is True


# --------------------------------------------------------------------------- #
# 6. Warning propagation
# --------------------------------------------------------------------------- #
def test_warning_propagates_and_is_not_complete_success(
    settings, warning_value_workbook, tmp_path
):
    analysis = _analyze(settings, warning_value_workbook)
    report = QualityGate().evaluate(analysis, source_path=warning_value_workbook)

    quality_stage = report.stage(GateStage.DATA_QUALITY)
    assert quality_stage is not None
    assert quality_stage.status is GateStatus.WARNING
    assert report.status is GateStatus.WARNING
    assert report.is_complete_success is False
    assert report.warnings

    result = Coordinator(settings, console=False).run(
        warning_value_workbook,
        use_ai=False,
        output=tmp_path / "warning_report.xlsx",
    )
    assert result.status == "completed_with_warnings"
    assert result.quality_gate.status is GateStatus.WARNING
    assert result.quality_gate.is_complete_success is False


# --------------------------------------------------------------------------- #
# 7. Blocking behavior for critical failures
# --------------------------------------------------------------------------- #
def test_missing_final_document_blocks_a_valid_report_claim(
    settings, clean_values_workbook, tmp_path
):
    analysis = _analyze(settings, clean_values_workbook)
    missing = tmp_path / "never_written.xlsx"
    specification = ReportSpecification(OutputFormat.XLSX, missing)
    generation = ReportGenerationResult(
        output_path=missing,
        output_format=OutputFormat.XLSX,
        writer=OutputFormat.XLSX.writer_name,
        content_validated=True,
    )

    report = QualityGate().evaluate(
        analysis,
        source_path=clean_values_workbook,
        specification=specification,
        generation=generation,
    )

    stage = report.stage(GateStage.DOCUMENT_INTEGRITY)
    assert stage is not None
    assert stage.status is GateStatus.FAIL
    assert stage.blocking is True
    assert report.blocked is True
    assert report.can_claim_valid_report is False
    assert any(check.code == "output-missing" for check in stage.checks)

    # Source removal is a second, independent critical failure.
    source = tmp_path / "source-then-deleted.xlsx"
    source.write_bytes(clean_values_workbook.read_bytes())
    source_analysis = _analyze(settings, source)
    source.unlink()
    source_report = QualityGate().evaluate(source_analysis, source_path=source)
    source_stage = source_report.stage(GateStage.SOURCE_AVAILABILITY)
    assert source_stage is not None
    assert source_stage.status is GateStatus.FAIL
    assert source_stage.blocking is True
    assert source_report.blocked is True


def test_folder_gate_evaluates_cross_source_and_merged_provenance(
    settings, make_workbook, tmp_path
):
    first = make_workbook(
        "منبع-الف.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ", "تاریخ"],
                ["اجاره", 100, "1403/01/15"],
                ["برق", 200, "1403/02/15"],
            ]
        },
    )
    second = make_workbook(
        "منبع-ب.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ", "تاریخ"],
                ["اجاره", 110, "1404/01/15"],
                ["برق", 210, "1404/02/15"],
            ]
        },
    )
    folder = tmp_path / "sources"
    folder.mkdir()
    for path in (first, second):
        (folder / path.name).write_bytes(path.read_bytes())

    result = FolderWorkflow(settings, progress=lambda message: None).run(
        folder,
        request="جمع شاخص‌ها را مقایسه کن",
        use_ai=False,
        output=tmp_path / "folder_gate.xlsx",
    )

    gate = result.quality_gate
    assert gate is not None
    assert gate.finalized is True
    assert [stage.stage for stage in gate.stages] == list(STAGE_ORDER)
    cross = gate.stage(GateStage.CROSS_SOURCE_RECONCILIATION)
    assert cross is not None
    assert cross.evidence["cross_source_decisions"] > 0
    assert cross.status is not GateStatus.PASS
    assert gate.provenance is not None
    assert gate.provenance.counts()["important"] >= 2
    assert result.to_dict()["quality_gate"]["status"] == gate.status.value


# --------------------------------------------------------------------------- #
# 8. Quality status consistency
# --------------------------------------------------------------------------- #
def test_quality_status_consistency_across_all_stages(
    settings, clean_values_workbook, tmp_path
):
    result = Coordinator(settings, console=False).run(
        clean_values_workbook,
        use_ai=False,
        output=tmp_path / "consistent.xlsx",
        profile="audit",
    )
    report = result.quality_gate
    assert report is not None
    assert report.finalized is True
    assert [stage.stage for stage in report.stages] == list(STAGE_ORDER)
    assert report.missing_stages == []
    assert report.status is worst_status(stage.status for stage in report.stages)
    assert report.status is GateStatus.PASS
    assert report.is_complete_success is True
    assert report.blocked is False
    assert report.can_claim_valid_report is True

    counts = report.counts()
    assert counts["PASS"] + counts["WARNING"] + counts["FAIL"] == len(STAGE_ORDER)
    assert counts["stages"] == len(STAGE_ORDER)
    assert counts["blocking"] == 0

    payload = report.to_dict()
    assert payload["status"] == "PASS"
    assert payload["can_claim_valid_report"] is True
    assert len(payload["stages"]) == len(STAGE_ORDER)
    assert payload["provenance"]["status"] == "PASS"
    assert result.to_dict()["quality_gate"]["status"] == "PASS"

    # Evidence reaches the delivered workbook: the audit profile has the gate
    # and provenance sheets; the clean profile only shows the summary.
    workbook = openpyxl.load_workbook(result.output_path)
    assert {"دروازه کیفیت", "ردیابی منبع (شواهد)"} <= set(workbook.sheetnames)
    gate_sheet = workbook["دروازه کیفیت"]
    assert gate_sheet["A1"].value == "دروازه کیفیت یکپارچه — وضعیت هر مرحله"
    workbook.close()

    clean = Coordinator(settings, console=False).run(
        clean_values_workbook,
        use_ai=False,
        output=tmp_path / "consistent_clean.xlsx",
        profile="clean",
    )
    clean_workbook = openpyxl.load_workbook(clean.output_path)
    assert "دروازه کیفیت" not in clean_workbook.sheetnames
    assert "ردیابی منبع (شواهد)" not in clean_workbook.sheetnames
    validation_text = "\n".join(
        str(cell.value)
        for row in clean_workbook["اعتبارسنجی داده"].iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "دروازه کیفیت یکپارچه" in validation_text
    clean_workbook.close()
