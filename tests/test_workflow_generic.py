"""General-purpose behaviour: files that do not look like the V4.2 sample."""

from __future__ import annotations

from decimal import Decimal

import openpyxl

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.models import MatchStatus, TableKind
from local_office_ai.excel.reader import file_sha256
from local_office_ai.quality import GateStage, GateStatus


def test_generic_workbook_is_understood(settings, generic_workbook):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(generic_workbook)
    assert analysis.primary_table is not None
    assert analysis.primary_table.sheet == "هزینه‌ها"
    assert analysis.primary_table.kind in {TableKind.GENERIC, TableKind.OBLIGATION}
    assert analysis.loans == []
    labels = {item.label for item in analysis.calculations}
    assert any(label.startswith("جمع مبلغ") for label in labels)
    categories = {item.category for item in analysis.anomalies}
    assert "formula-error" in categories
    assert "hidden-columns" in categories
    assert any(item.category == "duplicate-row" for item in analysis.anomalies)


def test_generic_workbook_full_run(settings, generic_workbook, tmp_path):
    before = file_sha256(generic_workbook)
    result = Coordinator(settings, console=False).run(
        generic_workbook, use_ai=False, output=tmp_path / "generic_out.xlsx"
    )
    assert result.status == "completed_with_warnings"
    assert result.quality_gate is not None
    assert result.quality_gate.is_complete_success is False
    assert result.quality_gate.status_for(GateStage.FILE_READABILITY) is GateStatus.WARNING
    assert file_sha256(generic_workbook) == before
    workbook = openpyxl.load_workbook(result.output_path)
    assert "داده استانداردشده" in workbook.sheetnames
    assert "ساختار فایل" in workbook.sheetnames
    workbook.close()


def test_english_loans_with_separate_commitment_sheet(settings, loans_english_workbook):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(loans_english_workbook)
    assert analysis.primary_table is not None
    assert analysis.primary_table.sheet == "Loans"
    assert analysis.primary_table.kind is TableKind.LOAN
    assert len(analysis.loans) == 3
    assert len(analysis.obligations) == 4
    values = {item.name: item.value for item in analysis.calculations}
    assert values["principal_sum"] == Decimal("54000")
    assert values["installments_sum"] == Decimal("2450")
    assert values["obligations_sum"] == Decimal("2800")

    by_name = {item.loan_name or item.obligation_name: item for item in analysis.reconciliation}
    assert by_name["Bank A"].status is MatchStatus.MATCH
    assert by_name["Bank B"].status is MatchStatus.PROBABLE
    assert by_name["Bank B"].difference == Decimal("50")
    assert by_name["Fund C"].status is MatchStatus.MATCH
    assert by_name["Charity"].status is MatchStatus.OUTSIDE_LOAN_TABLE


def test_workbook_without_table_does_not_crash(settings, no_table_workbook, tmp_path):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(no_table_workbook)
    assert analysis.primary_table is None
    assert analysis.structure_confidence == 0.0
    assert any("هیچ جدول" in warning for warning in analysis.warnings)
    result = Coordinator(settings, console=False).run(
        no_table_workbook, use_ai=False, output=tmp_path / "empty_out.xlsx"
    )
    assert result.status == "completed_with_warnings"
    assert result.quality_gate.is_complete_success is False
    workbook = openpyxl.load_workbook(result.output_path)
    assert "داشبورد" in workbook.sheetnames
    workbook.close()


def test_multi_sheet_multi_table_detection(settings, make_workbook):
    path = make_workbook(
        "multi.xlsx",
        {
            "فروش": [
                ["فروش فصل اول"],
                [],
                ["کالا", "تعداد", "مبلغ"],
                ["الف", 1, 1000],
                ["ب", 2, 2000],
                [],
                [],
                ["کالا", "تعداد", "مبلغ"],
                ["ج", 3, 3000],
                ["د", 4, 4000],
            ],
            "خرید": [
                ["خرید فصل اول"],
                ["کالا", "مبلغ"],
                ["الف", 500],
                ["ب", 700],
            ],
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    sheets = {table.sheet for table in analysis.tables}
    assert {"فروش", "خرید"}.issubset(sheets)
    sales_tables = [table for table in analysis.tables if table.sheet == "فروش"]
    assert len(sales_tables) == 2
    assert analysis.primary_table is not None


def test_hidden_and_merged_cells_are_reported(settings, generic_workbook):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(generic_workbook)
    sheet = next(item for item in analysis.meta.sheets if item.title == "هزینه‌ها")
    assert sheet.hidden_cols == 1
    assert sheet.merged_ranges == 1
    assert any("ادغام" in note for note in sheet.suspicious)


def test_unknown_structure_requires_confirmation(settings, make_workbook, tmp_path):
    """A low-confidence layout triggers the confirmation callback."""
    path = make_workbook(
        "vague.xlsx",
        {
            "S": [
                ["یادداشت", None, None],
                [None, "متن", 12],
                [None, "متن دوم", 15],
            ]
        },
    )
    calls: list[float] = []

    def confirm(analysis) -> bool:
        calls.append(analysis.structure_confidence)
        return True

    result = Coordinator(settings, console=False).run(
        path, use_ai=False, output=tmp_path / "vague_out.xlsx", confirm_structure=confirm
    )
    assert calls, "confirmation callback must be consulted"
    assert result.status == "completed_with_warnings"
    assert result.quality_gate.is_complete_success is False


def test_confirmation_can_abort(settings, make_workbook, tmp_path):
    path = make_workbook(
        "abort.xlsx",
        {"S": [["یادداشت"], [None, "متن", 12], [None, "متن دوم", 15]]},
    )
    from local_office_ai.core.exceptions import ProcessingAborted
    import pytest

    with pytest.raises(ProcessingAborted):
        Coordinator(settings, console=False).run(
            path,
            use_ai=False,
            output=tmp_path / "abort_out.xlsx",
            confirm_structure=lambda analysis: False,
        )
