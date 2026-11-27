from __future__ import annotations

from decimal import Decimal

from local_office_ai.core.models import (
    AuditResult,
    AuditStatus,
    Calculation,
    LoanRecord,
    MatchStatus,
    ObligationRecord,
    Provenance,
    Severity,
    SheetMeta,
    TableCandidate,
    TableKind,
    WorkbookMeta,
)
from local_office_ai.excel.anomaly_detection import AnomalyDetector
from local_office_ai.excel.reconciliation import Reconciler


def _loan(row, name, installment, principal="1000", count="10"):
    return LoanRecord(
        institution_raw=name,
        institution_key=name,
        principal=Decimal(principal),
        installment=Decimal(installment),
        installment_count=Decimal(count),
        last_installment_date=None,
        contract_repayment=None,
        provenance=Provenance("S", f"C{row}", row, 3),
    )


def _obligation(row, name, amount):
    return ObligationRecord(
        institution_raw=name,
        institution_key=name,
        amount=Decimal(amount),
        provenance=Provenance("S", f"C{row}", row, 3),
    )


def test_reconciliation_states(settings):
    loans = [
        _loan(3, "بانک مهر ایران", "1000"),
        _loan(4, "صندق امین", "500"),
        _loan(5, "بانک سامان", "300"),
    ]
    obligations = [
        _obligation(10, "بانک مهر ایران", "1000"),
        _obligation(12, "صندوق امین", "500"),
        _obligation(14, "بانک سامان", "350"),
        _obligation(16, "صندوق رفاه کهنز", "4250"),
    ]
    items = Reconciler(settings).reconcile(loans, obligations)
    by_name = {item.loan_name or item.obligation_name: item for item in items}

    assert by_name["بانک مهر ایران"].status is MatchStatus.MATCH
    assert by_name["صندق امین"].status is MatchStatus.MATCH
    assert by_name["صندق امین"].similarity >= 0.95
    assert by_name["بانک سامان"].status is MatchStatus.PROBABLE
    assert by_name["بانک سامان"].difference == Decimal("50")
    assert by_name["صندوق رفاه کهنز"].status is MatchStatus.OUTSIDE_LOAN_TABLE
    assert by_name["صندوق رفاه کهنز"].obligation_total == Decimal("4250")


def test_probable_name_match_needs_confirmation(settings):
    loans = [_loan(3, "بانک مهر ایران", "1000")]
    obligations = [_obligation(10, "بانک مهر ایرن", "1000")]
    item = Reconciler(settings).reconcile(loans, obligations)[0]
    assert item.status is MatchStatus.PROBABLE
    assert item.similarity >= 0.9
    assert "تأیید" in item.note


def test_no_obligation_for_loan(settings):
    item = Reconciler(settings).reconcile([_loan(3, "بانک الف", "100")], [])[0]
    assert item.status is MatchStatus.NO_OBLIGATION
    assert item.loan_total == Decimal("100")


def test_spelling_variants_are_reported(settings):
    loans = [_loan(3, "بانك مهر ايران", "1000")]
    obligations = [_obligation(10, "بانک مهر ایران", "1000")]
    item = Reconciler(settings).reconcile(loans, obligations)[0]
    assert item.status is MatchStatus.MATCH
    assert "یکسان‌سازی" in item.note


def _meta(**overrides):
    sheet = SheetMeta(
        title="S",
        state="visible",
        declared_dimensions="A1:Z100",
        used_dimensions="A1:D10",
        max_row=10,
        max_col=4,
        **overrides,
    )
    return WorkbookMeta(path="x.xlsx", sha256="abc", size_bytes=1, sheets=[sheet])


def test_anomaly_detection_core(settings):
    loans = [
        _loan(3, "بانک الف", "100", principal="1000"),
        _loan(4, "بانک الف", "100", principal="1000"),
        _loan(5, "بانک ب", "-50", principal="-500"),
        _loan(6, "بانک ج", "100", count="10.5"),
    ]
    loans[0].issues = ["date: تاریخ قابل خواندن نیست"]
    obligations = [_obligation(20, "بانک الف", "100"), _obligation(21, "بانک الف", "100")]
    calculations = [
        Calculation("x", "کنترل", Decimal("1"), verification_value=Decimal("2"), verified=False)
    ]
    audits = [
        AuditResult(
            measure="principal",
            label="مجموع اصل",
            declared=Decimal("10"),
            computed=Decimal("12"),
            difference=Decimal("2"),
            status=AuditStatus.MISMATCH,
            declared_source=Provenance("S", "D9", 9, 4),
        )
    ]
    primary = TableCandidate("S", 1, 2, 10, 1, 4, kind=TableKind.LOAN, confidence=0.9)
    anomalies = AnomalyDetector(settings).detect(
        _meta(hidden_cols=1, error_cells=["F4"]),
        [primary],
        primary,
        loans,
        obligations,
        calculations,
        audits,
        [],
    )
    categories = {item.category for item in anomalies}
    assert "duplicate-loan" in categories
    assert "negative-amount" in categories
    assert "fractional-count" in categories
    assert "duplicate-obligation" in categories
    assert "verification-failed" in categories
    assert "total-mismatch" in categories
    assert "formula-error" in categories
    assert "hidden-columns" in categories
    assert any(item.severity is Severity.CRITICAL for item in anomalies)


def test_no_table_is_critical(settings):
    anomalies = AnomalyDetector(settings).detect(_meta(), [], None, [], [], [], [], [])
    assert any(item.category == "no-financial-table" and item.severity is Severity.CRITICAL for item in anomalies)


def test_outlier_detection(settings):
    loans = [
        _loan(3, "A", "100", principal="1000"),
        _loan(4, "B", "100", principal="1100"),
        _loan(5, "C", "100", principal="1050"),
        _loan(6, "D", "100", principal="100000"),
    ]
    anomalies = AnomalyDetector(settings).detect(_meta(), [], None, loans, [], [], [], [])
    assert any(item.category == "outlier" for item in anomalies)
