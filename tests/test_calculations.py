from __future__ import annotations

from datetime import date
from decimal import Decimal

from local_office_ai.core.models import (
    AuditStatus,
    DeclaredTotal,
    LoanRecord,
    ObligationRecord,
    Provenance,
)
from local_office_ai.excel.calculations import FinancialCalculator


def _loan(row: int, name: str, principal: str, installment: str, count: str, repayment: str | None = None):
    return LoanRecord(
        institution_raw=name,
        institution_key=name,
        principal=Decimal(principal),
        installment=Decimal(installment),
        installment_count=Decimal(count),
        last_installment_date=date(2027, 1, 1),
        contract_repayment=Decimal(repayment) if repayment else None,
        provenance=Provenance("S", f"C{row}", row, 3),
    )


def _obligation(row: int, name: str, amount: str):
    return ObligationRecord(
        institution_raw=name,
        institution_key=name,
        amount=Decimal(amount),
        provenance=Provenance("S", f"C{row}", row, 3),
    )


def test_decimal_arithmetic_is_exact(settings):
    calculator = FinancialCalculator(settings)
    loans = [
        _loan(3, "A", "0.1", "0.2", "3"),
        _loan(4, "B", "0.2", "0.1", "3"),
    ]
    calculations = calculator.compute(loans, [], [])
    values = {item.name: item.value for item in calculations}
    assert values["principal_sum"] == Decimal("0.3")
    assert values["installments_sum"] == Decimal("0.3")
    assert values["contract_repayment_sum"] == Decimal("0.9")
    assert isinstance(values["principal_sum"], Decimal)


def test_sums_audits_and_verification(settings):
    calculator = FinancialCalculator(settings)
    loans = [
        _loan(3, "بانک الف", "1000", "100", "10", "1000"),
        _loan(4, "بانک ب", "2000", "250", "8", "2000"),
    ]
    obligations = [_obligation(10, "بانک الف", "100"), _obligation(12, "سایر", "50")]
    declared = [
        DeclaredTotal("principal", "جمع کل وام", Decimal("3000"), Provenance("S", "D9", 9, 4)),
        DeclaredTotal("installments", "جمع اقساط", Decimal("400"), Provenance("S", "F9", 9, 6)),
    ]
    calculations = calculator.compute(loans, obligations, declared)
    values = {item.name: item for item in calculations}
    assert values["principal_sum"].value == Decimal("3000")
    assert values["installments_sum"].value == Decimal("350")
    assert values["obligations_sum"].value == Decimal("150")
    assert values["contract_repayment_sum"].value == Decimal("3000")
    assert values["principal_sum"].verified is True
    assert values["contract_repayment_sum"].verified is True
    assert values["principal_sum_check"].verified is True
    assert values["obligation_gap"].value == Decimal("-200")

    audits = calculator.build_audits(calculations, declared)
    statuses = {audit.measure: audit.status for audit in audits}
    assert statuses["principal"] is AuditStatus.VERIFIED
    assert statuses["installments"] is AuditStatus.MISMATCH
    mismatch = next(audit for audit in audits if audit.measure == "installments")
    assert mismatch.difference == Decimal("-50")


def test_contract_repayment_mismatch_is_flagged(settings):
    calculator = FinancialCalculator(settings)
    loans = [_loan(3, "A", "1000", "100", "10", "1200")]
    calculations = calculator.compute(loans, [], [])
    repayment = next(item for item in calculations if item.name == "contract_repayment_sum")
    assert repayment.value == Decimal("1000")
    assert repayment.verification_value == Decimal("1200")
    assert repayment.verified is False


def test_no_declared_value_status(settings):
    calculator = FinancialCalculator(settings)
    calculations = calculator.compute([_loan(3, "A", "100", "10", "10")], [], [])
    audits = calculator.build_audits(calculations, [])
    assert all(audit.status is AuditStatus.NO_DECLARED_VALUE for audit in audits)


def test_generic_column_statistics(generic_workbook, settings):
    from local_office_ai.excel.reader import WorkbookReader
    from local_office_ai.excel.sheet_analyzer import SheetGrid
    from local_office_ai.excel.table_detector import TableDetector
    from local_office_ai.excel.workbook_inspector import inspect_workbook

    bundle = WorkbookReader(settings).load(generic_workbook)
    meta = inspect_workbook(bundle, settings)
    sheet = next(item for item in meta.sheets if item.title == "هزینه‌ها")
    grid = SheetGrid("هزینه‌ها", bundle.values["هزینه‌ها"], bundle.formulas["هزینه‌ها"], sheet.max_row, sheet.max_col)
    table = TableDetector(settings).detect(grid).tables[0]
    calculations = FinancialCalculator(settings).summarize_table(grid, table)
    labels = {item.label: item for item in calculations}
    amount = next(item for label, item in labels.items() if label.startswith("جمع مبلغ"))
    assert amount.value == Decimal("27100000")  # 12.5M + 1.25M + 0.85M + 12.5M
    assert amount.verified is True
    count = next(item for label, item in labels.items() if label.startswith("تعداد مقادیر مبلغ"))
    assert count.value == 4
    bundle.close()
