"""Prompt 02B: statistical reconciliation and calculation integrity.

Every workbook is a real .xlsx file created by openpyxl.  The tests cover the
nine required cases: correct and incorrect component totals, nested and
overlapping categories, zero-baseline percentages, missing denominators,
weighted versus unweighted averages, rounding tolerances, and increase/decrease
interpretation.  They also guard the workbook integration, the audit-only sheet
and the AI number allow-list.
"""

from __future__ import annotations

import json
from decimal import Decimal

import openpyxl
import pytest

from local_office_ai.ai.context_builder import allowed_numbers, build_facts
from local_office_ai.ai.response_validator import validate_response
from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.excel.statistical_reconciliation import (
    KIND_COMPONENT_TOTAL,
    KIND_GROUP_SUBGROUP,
    KIND_MONTHLY_PERIOD,
    KIND_SUBTOTAL_GRAND_TOTAL,
    ROLE_DATA,
    ROLE_PERCENTAGE,
    ROLE_SUBTOTAL,
    ROLE_TOTAL,
    AggregationComponent,
    HierarchyNode,
    ReconciliationStatus,
    Tolerance,
    compare_categories,
    compare_periods,
    compute_average,
    compute_change,
    compute_ratio,
    reconcile_components,
    reconcile_hierarchy,
)
from local_office_ai.excel.structure_intelligence import expected_days

PASS = ReconciliationStatus.PASS
WARNING = ReconciliationStatus.WARNING
FAIL = ReconciliationStatus.FAIL


def _analyze(settings, path):
    return ExcelIntelligenceWorkflow(settings).analyze_only(path)


def _reconciliations(analysis, kind: str | None = None):
    items = analysis.statistical_reconciliation.reconciliations
    return [item for item in items if kind is None or item.kind == kind]


def _payload(summary: str) -> str:
    return json.dumps(
        {
            "summary": summary,
            "findings": [],
            "missing_data": [],
            "recommendations": [],
            "risk_level": "low",
        },
        ensure_ascii=False,
    )


def _monthly_book(
    path,
    months,
    *,
    include_declared: bool = True,
    skip_last_days: dict[str, int] | None = None,
    empty_months: set[str] | None = None,
):
    """A two-year monthly workbook with numeric (not formula) total rows."""

    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for name, month_no, base in months:
        sheet = workbook.create_sheet(f"{name} 1404")
        sheet["B2"] = f"آمار روزانه {name} 1404"
        sheet["B4"] = "ردیف"
        sheet["C4"] = "تاریخ"
        sheet["D4"] = "تعداد کل"
        sheet["E4"] = "تعداد کل"
        sheet.merge_cells("D4:E4")
        sheet["D5"] = "1403"
        sheet["E5"] = "1404"
        if name in (empty_months or set()):
            sheet["B5"] = "جمع"
            continue
        days = expected_days(month_no)
        written_days = days - (skip_last_days or {}).get(name, 0)
        for day in range(1, written_days + 1):
            sheet.cell(5 + day, 2, day)
            sheet.cell(5 + day, 3, day)
            sheet.cell(5 + day, 4, base)
            sheet.cell(5 + day, 5, base + 100)
        total_row = 5 + written_days + 1
        sheet.cell(total_row, 2, "جمع")
        sheet.cell(total_row, 4, base * written_days if include_declared else None)
        sheet.cell(total_row, 5, (base + 100) * written_days if include_declared else None)
    workbook.save(path)
    return path


# --------------------------------------------------------------------------- #
# 1 and 2: component-to-total reconciliation
# --------------------------------------------------------------------------- #
def test_correct_component_to_total_reconciliation(settings, make_workbook):
    path = make_workbook(
        "correct_total.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
                ["آب", 300],
                ["جمع کل", 600],
            ]
        },
    )
    analysis = _analyze(settings, path)
    items = _reconciliations(analysis, KIND_COMPONENT_TOTAL)
    assert len(items) == 1
    item = items[0]
    assert item.status is PASS
    assert item.expected == Decimal("600")
    assert item.calculated == Decimal("600")
    assert item.difference == Decimal("0")
    assert item.verified
    assert "جمع" in item.label
    assert item.sources


def test_incorrect_component_totals_are_failures(settings, make_workbook):
    path = make_workbook(
        "wrong_total.xlsx",
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
    item = _reconciliations(analysis, KIND_COMPONENT_TOTAL)[0]
    assert item.status is FAIL
    assert item.expected == Decimal("700")
    assert item.calculated == Decimal("600")
    assert item.difference == Decimal("100")
    assert not item.verified
    assert analysis.statistical_reconciliation.status is FAIL
    assert "اختلاف" in item.explanation


def test_subtotal_to_grand_total_uses_subtotals_not_data_rows(settings, make_workbook):
    path = make_workbook(
        "subtotals.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["گروه الف", None],
                ["اجاره", 100],
                ["برق", 200],
                ["جمع گروه", 300],
                ["گروه ب", None],
                ["آب", 400],
                ["گاز", 100],
                ["جمع گروه", 500],
                ["جمع کل", 800],
            ]
        },
    )
    analysis = _analyze(settings, path)
    groups = _reconciliations(analysis, KIND_GROUP_SUBGROUP)
    grand = _reconciliations(analysis, KIND_SUBTOTAL_GRAND_TOTAL)
    assert [item.expected for item in groups] == [Decimal("300"), Decimal("500")]
    assert any("گروه الف" in item.label for item in groups)
    assert any("گروه ب" in item.label for item in groups)
    assert len(grand) == 1
    assert grand[0].status is PASS
    assert grand[0].expected == Decimal("800")
    assert grand[0].calculated == Decimal("800")
    assert {component["role"] for component in grand[0].components} == {ROLE_SUBTOTAL}


def test_grand_total_includes_data_after_the_last_subtotal(settings, make_workbook):
    path = make_workbook(
        "subtotal_and_trailing_data.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["گروه الف", None],
                ["اجاره", 100],
                ["برق", 100],
                ["جمع گروه", 200],
                ["آب", 300],
                ["گاز", 200],
                ["جمع کل", 700],
            ]
        },
    )
    analysis = _analyze(settings, path)
    grand = _reconciliations(analysis, KIND_SUBTOTAL_GRAND_TOTAL)
    assert len(grand) == 1
    assert grand[0].status is PASS
    assert grand[0].expected == Decimal("700")
    assert grand[0].calculated == Decimal("700")
    assert {component["role"] for component in grand[0].components} == {
        ROLE_SUBTOTAL,
        ROLE_DATA,
    }


# --------------------------------------------------------------------------- #
# 3 and 4: nested and overlapping categories
# --------------------------------------------------------------------------- #
def test_nested_category_structures_sum_only_immediate_children():
    root = HierarchyNode(
        "کل",
        Decimal("300"),
        unit="تومان",
        children=[
            HierarchyNode(
                "الف",
                Decimal("100"),
                children=[
                    HierarchyNode("الف۱", Decimal("60")),
                    HierarchyNode("الف۲", Decimal("40")),
                ],
            ),
            HierarchyNode(
                "ب",
                Decimal("200"),
                children=[
                    HierarchyNode("ب۱", Decimal("120")),
                    HierarchyNode("ب۲", Decimal("80")),
                ],
            ),
        ],
    )
    results = reconcile_hierarchy(root)
    assert len(results) == 3
    assert all(item.status is PASS for item in results)
    assert all(item.verified for item in results)
    by_label = {item.label: item for item in results}
    # The root is checked against its two children (100 + 200), never against
    # all four grandchildren.
    assert by_label["کل"].calculated == Decimal("300")
    assert len(by_label["کل"].components) == 2


def test_overlapping_categories_cannot_be_summed():
    components = [
        AggregationComponent(
            "اداری — بخش اول",
            Decimal("100"),
            categories=frozenset({"اداری"}),
        ),
        AggregationComponent(
            "اداری — بخش دوم",
            Decimal("200"),
            categories=frozenset({"اداری"}),
        ),
    ]
    result = reconcile_components(Decimal("300"), components)
    assert result.status is FAIL
    assert result.calculated is None
    assert result.evidence is not None
    assert "اداری" in result.evidence.overlap
    assert "هم‌پوشان" in result.explanation


def test_grand_total_and_percentage_rows_are_never_components():
    with_total = reconcile_components(
        Decimal("10"),
        [
            AggregationComponent("جزء", Decimal("10"), role=ROLE_DATA),
            AggregationComponent("جمع کل", Decimal("10"), role=ROLE_TOTAL),
        ],
    )
    assert with_total.status is FAIL
    assert with_total.calculated is None
    assert "جمع کل" in with_total.explanation

    with_percentage = reconcile_components(
        Decimal("10"),
        [
            AggregationComponent("تعداد", Decimal("10"), role=ROLE_DATA),
            AggregationComponent("درصد", Decimal("50"), role=ROLE_PERCENTAGE),
        ],
    )
    assert with_percentage.status is FAIL
    assert with_percentage.calculated is None
    assert "درصد" in with_percentage.explanation


def test_missing_component_is_not_treated_as_zero():
    result = reconcile_components(
        Decimal("30"),
        [
            AggregationComponent("الف", Decimal("10")),
            AggregationComponent("ب", None),
        ],
    )
    assert result.status is WARNING
    assert result.calculated is None
    assert result.difference is None
    assert "ب" in result.explanation
    assert not result.verified


def test_mixed_units_are_not_summed():
    result = reconcile_components(
        Decimal("30"),
        [
            AggregationComponent("الف", Decimal("10"), unit="تومان"),
            AggregationComponent("ب", Decimal("20"), unit="ریال"),
        ],
    )
    assert result.status is FAIL
    assert result.calculated is None
    assert result.evidence is not None
    assert result.evidence.mixed_units == ["تومان", "ریال"]


def test_component_total_without_expected_is_warning_not_verified():
    result = reconcile_components(
        None,
        [
            AggregationComponent("الف", Decimal("10")),
            AggregationComponent("ب", Decimal("20")),
        ],
    )
    assert result.status is WARNING
    assert result.expected is None
    assert result.calculated == Decimal("30")
    assert result.difference is None
    assert not result.verified
    assert "تأیید نشد" in result.explanation


def test_percentage_total_is_not_reconciled_as_a_count(settings, make_workbook):
    path = make_workbook(
        "percentage_total.xlsx",
        {
            "شاخص‌ها": [
                ["شرح", "درصد"],
                ["الف", 10],
                ["ب", 20],
                ["جمع کل", 30],
            ]
        },
    )
    analysis = _analyze(settings, path)
    assert not _reconciliations(analysis, KIND_COMPONENT_TOTAL)
    averages = analysis.statistical_reconciliation.averages
    assert averages
    assert averages[0].status is WARNING
    assert averages[0].value is None


def test_workbook_unit_column_prevents_mixed_unit_sum(settings, make_workbook):
    path = make_workbook(
        "mixed_units.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "واحد", "مبلغ"],
                ["اجاره", "تومان", 100],
                ["برق", "ریال", 200],
                ["جمع کل", "تومان", 300],
            ]
        },
    )
    analysis = _analyze(settings, path)
    item = _reconciliations(analysis, KIND_COMPONENT_TOTAL)[0]
    assert item.status is FAIL
    assert item.calculated is None
    assert {component["unit"] for component in item.components} == {"تومان", "ریال"}


# --------------------------------------------------------------------------- #
# 5 and 9: zero baseline, increase and decrease
# --------------------------------------------------------------------------- #
def test_zero_baseline_percentage_is_not_fabricated():
    increase = compute_change(Decimal("10"), Decimal("0"))
    assert increase.absolute == Decimal("10")
    assert increase.percent is None
    assert increase.direction == "افزایش"
    assert increase.baseline == "zero"
    assert increase.status is WARNING
    assert "صفر" in increase.explanation

    decrease = compute_change(Decimal("-5"), Decimal("0"))
    assert decrease.direction == "کاهش"
    assert decrease.percent is None
    assert decrease.status is WARNING

    unchanged = compute_change(Decimal("0"), Decimal("0"))
    assert unchanged.direction == "بدون تغییر"
    assert unchanged.percent is None


def test_increase_and_decrease_interpretation():
    up = compute_change(Decimal("120"), Decimal("100"))
    assert up.absolute == Decimal("20")
    assert up.percent == Decimal("20.00")
    assert up.direction == "افزایش"
    assert up.status is PASS

    down = compute_change(Decimal("80"), Decimal("100"))
    assert down.absolute == Decimal("-20")
    assert down.percent == Decimal("-20.00")
    assert down.direction == "کاهش"

    # A negative baseline must not flip an increase into a decrease.
    negative_baseline = compute_change(Decimal("-50"), Decimal("-100"))
    assert negative_baseline.absolute == Decimal("50")
    assert negative_baseline.percent == Decimal("50.00")
    assert negative_baseline.direction == "افزایش"

    missing = compute_change(Decimal("10"), None)
    assert missing.absolute is None
    assert missing.percent is None
    assert missing.status is WARNING


# --------------------------------------------------------------------------- #
# 6 and 7: denominators and weighted averages
# --------------------------------------------------------------------------- #
def test_missing_denominator_refuses_a_ratio_and_a_required_weight():
    ratio = compute_ratio(Decimal("10"), Decimal("0"))
    assert ratio.value is None
    assert ratio.status is WARNING
    assert "صفر" in ratio.explanation

    no_denominator = compute_ratio(Decimal("10"), None)
    assert no_denominator.value is None
    assert no_denominator.status is WARNING

    weighted = compute_average(
        [Decimal("10"), Decimal("20")],
        weights=None,
        weights_required=True,
    )
    assert weighted.value is None
    assert weighted.weighted_used is False
    assert weighted.status is FAIL
    assert "میانگین ساده" in weighted.explanation


def test_weighted_average_differs_from_unweighted_when_required():
    values = [Decimal("10"), Decimal("20")]
    weights = [Decimal("1"), Decimal("3")]

    required = compute_average(values, weights, weights_required=True)
    assert required.status is PASS
    assert required.weighted_used
    assert required.value == Decimal("17.5000")
    assert required.unweighted == Decimal("15.0000")
    assert required.value != required.unweighted

    ordinary = compute_average(values, weights, weights_required=False)
    assert ordinary.status is PASS
    assert ordinary.weighted_used is False
    assert ordinary.value == Decimal("15.0000")
    # The weighted value is still recorded as evidence, never reported instead.
    assert ordinary.weighted == Decimal("17.5000")


def test_successful_ratio_and_zero_weight_failure():
    ratio = compute_ratio(Decimal("10"), Decimal("4"))
    assert ratio.status is PASS
    assert ratio.value == Decimal("2.5000")

    zero_weight = compute_average(
        [Decimal("10"), Decimal("20")],
        [Decimal("0"), Decimal("0")],
        weights_required=True,
    )
    assert zero_weight.status is FAIL
    assert zero_weight.value is None
    assert zero_weight.weighted_used is False


def test_ratio_without_weight_column_is_warning_not_silent(settings, make_workbook):
    path = make_workbook(
        "ratio_without_weight.xlsx",
        {
            "شاخص‌ها": [
                ["شرح", "درصد"],
                ["الف", 10],
                ["ب", 20],
            ]
        },
    )
    analysis = _analyze(settings, path)
    averages = analysis.statistical_reconciliation.averages
    assert averages
    item = averages[0]
    assert item.status is WARNING
    assert item.value is None
    assert item.weights_required
    assert not item.weights_available
    assert "وزن" in item.explanation
    assert "میانگین ساده" in item.explanation
    assert analysis.statistical_reconciliation.status is WARNING


def test_workbook_ratio_with_ambiguous_weight_columns_is_warning(
    settings, make_workbook
):
    path = make_workbook(
        "ambiguous_weight.xlsx",
        {
            "شاخص‌ها": [
                ["شرح", "درصد", "تعداد", "مبلغ"],
                ["الف", 10, 1, 100],
                ["ب", 20, 3, 300],
            ]
        },
    )
    analysis = _analyze(settings, path)
    averages = analysis.statistical_reconciliation.averages
    assert len(averages) == 1
    assert averages[0].status is WARNING
    assert averages[0].value is None
    assert "چند ستون وزن" in averages[0].explanation


def test_workbook_ratio_uses_the_available_weight_column(settings, make_workbook):
    path = make_workbook(
        "weighted.xlsx",
        {
            "شاخص‌ها": [
                ["شرح", "درصد", "تعداد"],
                ["الف", 10, 1],
                ["ب", 20, 3],
            ]
        },
    )
    analysis = _analyze(settings, path)
    averages = analysis.statistical_reconciliation.averages
    assert averages
    item = averages[0]
    assert item.weights_required
    assert item.weighted_used
    assert item.value == Decimal("17.5000")
    assert item.unweighted == Decimal("15.0000")


# --------------------------------------------------------------------------- #
# 8: rounding tolerance
# --------------------------------------------------------------------------- #
def test_rounding_tolerances_pass_warning_and_fail():
    tolerance = Tolerance.for_workbook(components=3)
    assert tolerance.classify(Decimal("0"), Decimal("600")) is PASS
    assert tolerance.classify(Decimal("0.004"), Decimal("600")) is PASS
    assert tolerance.classify(Decimal("0.10"), Decimal("600")) is WARNING
    assert tolerance.classify(Decimal("1"), Decimal("600")) is FAIL

    # The tolerance is evidence, not a silent repair.
    result = reconcile_components(
        Decimal("600.10"),
        [
            AggregationComponent("الف", Decimal("200")),
            AggregationComponent("ب", Decimal("400")),
        ],
        tolerance=Tolerance(absolute=Decimal("0.01"), warning_factor=Decimal("10")),
    )
    assert result.status is WARNING
    assert result.difference == Decimal("0.10")
    assert not result.verified

    large = Tolerance.for_workbook(components=1)
    assert large.classify(Decimal("0.5"), Decimal("1000000")) is PASS
    assert large.classify(Decimal("5"), Decimal("1000000")) is WARNING
    assert large.classify(Decimal("11"), Decimal("1000000")) is FAIL


# --------------------------------------------------------------------------- #
# Category and period comparisons
# --------------------------------------------------------------------------- #
def test_category_level_comparisons():
    current = {"اداری": Decimal("120"), "عمرانی": Decimal("80"), "جدید": Decimal("5")}
    previous = {"اداری": Decimal("100"), "عمرانی": Decimal("100")}
    results = {item.label: item for item in compare_categories(current, previous)}

    assert results["اداری"].change.absolute == Decimal("20")
    assert results["اداری"].change.percent == Decimal("20.00")
    assert results["اداری"].status is PASS
    assert results["عمرانی"].change.direction == "کاهش"
    assert results["عمرانی"].status is PASS
    assert results["جدید"].status is WARNING
    assert results["جدید"].change.percent is None


def test_period_comparison_flags_incomplete_months():
    result = compare_periods(
        current_label="1404",
        previous_label="1403",
        current=Decimal("1200"),
        previous=Decimal("1000"),
        months_used=["فروردین", "اردیبهشت"],
        incomplete_excluded=["خرداد"],
    )
    assert result.change.percent == Decimal("20.00")
    assert result.status is WARNING
    assert "خرداد" in result.explanation
    assert result.months_used == ["فروردین", "اردیبهشت"]

    missing = compare_periods(
        current_label="1404",
        previous_label="1403",
        current=Decimal("1200"),
        previous=Decimal("1000"),
        months_used=["فروردین"],
        months_missing=["تیر"],
    )
    assert missing.status is WARNING
    assert missing.months_missing == ["تیر"]
    assert "بدون داده" in missing.explanation


def test_monthly_to_period_reconciliation_and_cross_period_comparison(tmp_path, settings):
    path = _monthly_book(
        tmp_path / "monthly.xlsx",
        [("فروردین", 1, 10), ("اردیبهشت", 2, 20)],
    )
    analysis = _analyze(settings, path)
    monthly = _reconciliations(analysis, KIND_MONTHLY_PERIOD)
    assert {item.label for item in monthly} == {
        "تعداد کل — 1403",
        "تعداد کل — 1404",
    }
    assert all(item.status is PASS for item in monthly)
    assert {item.expected for item in monthly} == {Decimal("930"), Decimal("7130")}

    comparisons = analysis.statistical_reconciliation.comparisons
    assert len(comparisons) == 1
    comparison = comparisons[0]
    assert comparison.metric == "تعداد کل"
    assert comparison.previous_label == "1403"
    assert comparison.current_label == "1404"
    assert comparison.change.absolute == Decimal("6200")
    assert comparison.change.percent == Decimal("666.67")
    assert comparison.status is PASS
    assert comparison.months_used == ["فروردین", "اردیبهشت"]


def test_overlapping_detector_views_do_not_double_count(tmp_path, settings):
    path = _monthly_book(tmp_path / "monthly.xlsx", [("فروردین", 1, 10)])
    analysis = _analyze(settings, path)
    keys = [
        (item.sheet, item.measure, item.expected_source)
        for item in analysis.statistical_reconciliation.reconciliations
    ]
    assert len(keys) == len(set(keys))


def test_monthly_reconciliation_reports_incomplete_months(tmp_path, settings):
    path = _monthly_book(
        tmp_path / "monthly_incomplete.xlsx",
        [("فروردین", 1, 10), ("اردیبهشت", 2, 20)],
        skip_last_days={"اردیبهشت": 1},
    )
    analysis = _analyze(settings, path)
    monthly = _reconciliations(analysis, KIND_MONTHLY_PERIOD)
    assert monthly
    assert all(item.status is WARNING for item in monthly)
    assert all("ناقص" in item.explanation for item in monthly)
    assert all("اردیبهشت" in item.explanation for item in monthly)
    assert all(not item.verified for item in monthly)


def test_monthly_reconciliation_reports_missing_declared_totals(tmp_path, settings):
    path = _monthly_book(
        tmp_path / "monthly_no_declared.xlsx",
        [("فروردین", 1, 10), ("اردیبهشت", 2, 20)],
        include_declared=False,
    )
    analysis = _analyze(settings, path)
    monthly = _reconciliations(analysis, KIND_MONTHLY_PERIOD)
    assert monthly
    assert all(item.status is WARNING for item in monthly)
    assert all(item.calculated is None for item in monthly)
    assert all("مقدار" in item.explanation for item in monthly)


def test_monthly_reconciliation_reports_a_month_without_values(tmp_path, settings):
    path = _monthly_book(
        tmp_path / "monthly_empty_month.xlsx",
        [("فروردین", 1, 10), ("اردیبهشت", 2, 20)],
        empty_months={"اردیبهشت"},
    )
    analysis = _analyze(settings, path)
    report = analysis.statistical_reconciliation
    assert report.warnings
    assert any("تطبیق ماهانه" in warning for warning in report.warnings)
    assert report.status is WARNING


def test_report_without_a_justified_relationship_is_warning_not_pass(
    settings, make_workbook
):
    path = make_workbook(
        "no_relationship.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
            ]
        },
    )
    analysis = _analyze(settings, path)
    report = analysis.statistical_reconciliation
    assert not report.reconciliations
    assert not report.averages
    assert not report.comparisons
    assert report.status is WARNING
    assert report.warnings


# --------------------------------------------------------------------------- #
# Traceability, output profile and AI restrictions
# --------------------------------------------------------------------------- #
def test_report_serializes_every_required_field(settings, make_workbook):
    path = make_workbook(
        "serialize.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
                ["جمع کل", 300],
            ]
        },
    )
    analysis = _analyze(settings, path)
    payload = analysis.statistical_reconciliation.to_dict()
    record = payload["reconciliations"][0]
    for key in (
        "expected",
        "calculated",
        "difference",
        "tolerance",
        "status",
        "explanation",
        "sources",
    ):
        assert key in record
    assert payload["status"] == "PASS"
    assert analysis.to_dict()["statistical_reconciliation"]["status"] == "PASS"


def test_clean_profile_has_no_audit_sheet_and_audit_profile_does(settings, make_workbook, tmp_path):
    path = make_workbook(
        "delivery.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
                ["جمع کل", 300],
            ]
        },
    )
    clean_path = tmp_path / "clean.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=clean_path, profile="clean"
    )
    clean = openpyxl.load_workbook(clean_path)
    assert "تطبیق آماری و صحت محاسبات" not in clean.sheetnames
    validation = clean["اعتبارسنجی داده"]
    values = [str(validation.cell(row, 1).value) for row in range(1, validation.max_row + 1)]
    assert any("تطبیق آماری" in value for value in values)
    clean.close()

    audit_path = tmp_path / "audit.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=audit_path, profile="audit"
    )
    audit = openpyxl.load_workbook(audit_path)
    assert "تطبیق آماری و صحت محاسبات" in audit.sheetnames
    sheet = audit["تطبیق آماری و صحت محاسبات"]
    assert "وضعیت کلی" in str(sheet["A3"].value)
    audit.close()


def test_clean_profile_renders_period_comparisons_without_a_technical_sheet(
    settings, tmp_path
):
    path = _monthly_book(
        tmp_path / "monthly.xlsx",
        [("فروردین", 1, 10), ("اردیبهشت", 2, 20)],
    )
    output = tmp_path / "clean_monthly.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=output, profile="clean"
    )
    workbook = openpyxl.load_workbook(output)
    assert "تطبیق آماری و صحت محاسبات" not in workbook.sheetnames
    validation = workbook["اعتبارسنجی داده"]
    values = [str(validation.cell(row, 1).value) for row in range(1, validation.max_row + 1)]
    assert any("تطبیق آماری" in value for value in values)
    workbook.close()


def test_clean_profile_shows_the_no_relationship_warning(
    settings, make_workbook, tmp_path
):
    path = make_workbook(
        "clean_no_relationship.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
            ]
        },
    )
    output = tmp_path / "clean_no_relationship_out.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=output, profile="clean"
    )
    workbook = openpyxl.load_workbook(output)
    validation = workbook["اعتبارسنجی داده"]
    values = [
        str(validation.cell(row, column).value)
        for row in range(1, validation.max_row + 1)
        for column in range(1, 6)
    ]
    assert any("رابطه جمع قابل توجیه" in value for value in values)
    workbook.close()

    audit_path = tmp_path / "clean_no_relationship_audit.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=audit_path, profile="audit"
    )
    audit = openpyxl.load_workbook(audit_path)
    sheet = audit["تطبیق آماری و صحت محاسبات"]
    audit_values = [
        str(sheet.cell(row, column).value)
        for row in range(1, sheet.max_row + 1)
        for column in range(1, 12)
    ]
    assert any("هشدارهای موتور تطبیق" in value for value in audit_values)
    assert any("رابطه جمع قابل توجیه" in value for value in audit_values)
    audit.close()


def test_clean_profile_lists_average_warnings(settings, make_workbook, tmp_path):
    path = make_workbook(
        "clean_average_warning.xlsx",
        {
            "شاخص‌ها": [
                ["شرح", "درصد"],
                ["الف", 10],
                ["ب", 20],
            ]
        },
    )
    output = tmp_path / "clean_average_warning_out.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=output, profile="clean"
    )
    workbook = openpyxl.load_workbook(output)
    validation = workbook["اعتبارسنجی داده"]
    values = [
        str(validation.cell(row, column).value)
        for row in range(1, validation.max_row + 1)
        for column in range(1, 6)
    ]
    assert any("میانگین" in value for value in values)
    workbook.close()


def test_ai_may_only_quote_python_approved_reconciliation_numbers(settings, make_workbook):
    path = make_workbook(
        "ai_numbers.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
                ["جمع کل", 300],
            ]
        },
    )
    analysis = _analyze(settings, path)
    allowed = allowed_numbers(analysis)
    assert "300" in allowed

    grounded = validate_response(
        _payload("جمع کل ۳۰۰ است و با جمع اجزا مطابقت دارد."),
        allowed_numbers=allowed,
        mismatch_exists=False,
    )
    assert grounded.is_grounded
    assert not grounded.unsupported_numbers

    invented = validate_response(
        _payload("جمع کل ۳۰۱ است."),
        allowed_numbers=allowed,
        mismatch_exists=False,
    )
    assert not invented.is_grounded
    assert invented.unsupported_numbers

    facts, _meta = build_facts(analysis, settings)
    assert facts["statistical_reconciliation"]["status"] == "PASS"
    assert facts["statistical_reconciliation"]["reconciliations"]


def test_ai_reconciliation_mismatch_is_reported_as_a_real_mismatch(settings, make_workbook):
    path = make_workbook(
        "ai_mismatch.xlsx",
        {
            "هزینه‌ها": [
                ["شرح", "مبلغ"],
                ["اجاره", 100],
                ["برق", 200],
                ["جمع کل", 400],
            ]
        },
    )
    analysis = _analyze(settings, path)
    assert analysis.statistical_reconciliation.status is FAIL
    mismatch = validate_response(
        _payload("مغایرت ۱۰۰ واحدی در جمع کل ثبت شده است."),
        allowed_numbers=allowed_numbers(analysis),
        mismatch_exists=True,
    )
    assert not mismatch.contradictions
