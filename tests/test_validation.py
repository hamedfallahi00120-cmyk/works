"""Validation checklist, raw-cell preservation, time budget and aggregation."""

from __future__ import annotations

import threading
from decimal import Decimal

import openpyxl
import pytest

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow, WorkflowOptions
from local_office_ai.core.exceptions import ProcessingAborted
from local_office_ai.core.models import AuditStatus
from local_office_ai.excel.validation import DataValidator, overall_status


def test_sample_workbook_validation_checklist(settings, sample_path):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)
    checks = {check.code: check for check in analysis.validation}
    assert len(checks) == 9
    assert checks["file-sheets"].status == "passed"
    assert checks["structure-confidence"].status == "passed"
    assert checks["declared-totals"].status == "passed"
    assert checks["declared-totals"].affected == 5
    assert checks["loan-fields"].status == "passed"
    assert checks["loan-fields"].affected == 6
    assert checks["reconciliation"].status == "warning"
    assert overall_status(analysis.validation) == "warning"
    assert analysis.to_dict()["validation"]


def test_generic_workbook_skips_loan_checks(settings, generic_workbook):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(generic_workbook)
    checks = {check.code: check for check in analysis.validation}
    assert checks["loan-fields"].status == "skipped"
    assert checks["numeric-integrity"].status in {"passed", "warning"}
    assert checks["structure-confidence"].status in {"passed", "warning", "failed"}


def test_incomplete_loan_row_fails_validation(settings, make_workbook):
    workbook = make_workbook(
        "incomplete.xlsx",
        {
            "وام‌ها": [
                ["ردیف", "نام وام", "مبلغ وام", "مبلغ هر قسط", "تعداد اقساط"],
                [1, "بانک الف", 1000, 100, 10],
                [2, "بانک ب", 2000, None, 8],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(workbook)
    checks = {check.code: check for check in analysis.validation}
    assert checks["loan-fields"].status == "failed"
    assert checks["loan-fields"].affected == 1
    assert "بانک ب" in checks["loan-fields"].detail


def test_report_contains_validation_and_raw_sheets(settings, sample_path, tmp_path):
    result = Coordinator(settings, console=False).run(
        sample_path, use_ai=False, output=tmp_path / "report.xlsx"
    )
    workbook = openpyxl.load_workbook(result.output_path)
    assert "اعتبارسنجی داده" in workbook.sheetnames
    assert "داده خام" in workbook.sheetnames
    assert {
        "شناسنامه داده", "آمار توصیفی", "توزیع فراوانی", "تحلیل گروه‌ها",
        "روند و همبستگی", "روش‌شناسی آماری",
    } <= set(workbook.sheetnames)
    assert len(workbook.sheetnames) >= 22

    validation_ws = workbook["اعتبارسنجی داده"]
    assert "وضعیت کلی" in str(validation_ws["A2"].value)
    labels = [validation_ws.cell(row, 1).value for row in range(5, 14)]
    assert "حسابرسی جمع‌های اعلامی فایل" in labels

    raw_ws = workbook["داده خام"]
    assert [raw_ws.cell(3, column).value for column in range(1, 9)] == [
        "برگه", "سلول", "ستون", "عنوان", "مقدار خام", "مقدار استانداردشده", "نوع", "هشدارها",
    ]
    raw_row = [raw_ws.cell(4, column).value for column in range(1, 7)]
    assert raw_row[0] == "Sheet1"
    assert raw_row[1] and raw_row[4] is not None
    workbook.close()


def test_grouped_totals_and_shares_for_generic_table(settings, generic_workbook):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(generic_workbook)
    names = {item.name: item for item in analysis.calculations}
    assert any(name.startswith("group_count::") for name in names)
    sums = {name: item for name, item in names.items() if name.startswith("group_sum::")}
    assert sums, "a generic table with a category column must produce grouped sums"
    rent = [item for name, item in sums.items() if "اجاره" in name]
    assert rent and rent[0].value == Decimal("25000000")
    shares = {name: item for name, item in names.items() if name.startswith("group_share::")}
    assert shares
    total_share = sum(item.value for item in shares.values())
    assert Decimal("99") <= total_share <= Decimal("101")


def test_remaining_balance_when_paid_column_exists(settings, make_workbook):
    workbook = make_workbook(
        "paid.xlsx",
        {
            "وام‌ها": [
                ["ردیف", "نام وام", "مبلغ وام", "مبلغ هر قسط", "تعداد اقساط", "اقساط پرداخت شده"],
                [1, "بانک الف", 1200000, 100000, 12, 4],
                [2, "بانک ب", 800000, 50000, 16, 6],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(workbook)
    values = {item.name: item.value for item in analysis.calculations}
    assert values["remaining_installments"] == Decimal("18")
    assert values["remaining_amount"] == Decimal("1300000")
    assert analysis.loans[0].paid_installments == Decimal("4")
    assert analysis.loans[0].remaining_amount == Decimal("800000")


def test_processing_time_budget_aborts(settings, sample_path, tmp_path):
    settings.max_processing_seconds = 0.0001
    workflow = ExcelIntelligenceWorkflow(settings)
    with pytest.raises(ProcessingAborted) as error:
        workflow.run(
            WorkflowOptions(
                source=sample_path,
                use_ai=False,
                output_path=tmp_path / "budget.xlsx",
            )
        )
    assert "سقف زمان پردازش" in str(error.value)
    assert not (tmp_path / "budget.xlsx").exists()


def test_abort_event_stops_processing(settings, sample_path, tmp_path):
    workflow = ExcelIntelligenceWorkflow(settings)
    workflow.abort_event = threading.Event()
    workflow.abort_event.set()
    with pytest.raises(ProcessingAborted):
        workflow.run(
            WorkflowOptions(source=sample_path, use_ai=False, output_path=tmp_path / "abort.xlsx")
        )


def test_validator_reports_failed_audit(settings, sample_path):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)
    analysis.audits[0].difference = Decimal("10")
    analysis.audits[0].status = AuditStatus.MISMATCH
    checks = {check.code: check for check in DataValidator(settings).validate(analysis, {})}
    assert checks["declared-totals"].status == "failed"
    assert checks["declared-totals"].affected == 1


def test_raw_values_survive_standardization(settings, generic_workbook):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(generic_workbook)
    rows = analysis.normalized_rows
    assert rows
    converted = [
        cell for row in rows for cell in row["cells"]
        if cell["kind"] == "number" and isinstance(cell["raw"], str)
    ]
    assert converted, "Persian-digit amounts must keep their original text"
    cell = converted[0]
    assert cell["raw"] != str(cell["value"])
    assert isinstance(cell["value"], Decimal)


def test_generic_dashboard_charts_group_totals(settings, generic_workbook, tmp_path):
    result = Coordinator(settings, console=False).run(
        generic_workbook, use_ai=False, output=tmp_path / "generic_report.xlsx"
    )
    workbook = openpyxl.load_workbook(result.output_path)
    assert len(workbook["داشبورد"]._charts) >= 1
    names = workbook.sheetnames
    assert names.index("داده خام") == names.index("داده استانداردشده") + 1
    workbook.close()


def test_validation_scope_uses_file_name(settings, sample_path):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)
    scopes = {check.code: check.scope for check in analysis.validation}
    assert scopes["file-sheets"] == "Book2_AI_V4_2_Monitored.xlsx"
    assert scopes["numeric-integrity"] == "Book2_AI_V4_2_Monitored.xlsx"
