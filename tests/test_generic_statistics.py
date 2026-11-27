"""Focused tests for the general-purpose statistical engine."""

from __future__ import annotations

from decimal import Decimal
from hashlib import sha256

import openpyxl
import pytest
from openpyxl import load_workbook

from local_office_ai.ai.context_builder import allowed_numbers, build_facts
from local_office_ai.ai.prompt_manager import offline_summary
from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.excel.normalization import parse_number


@pytest.fixture()
def office_workbook(make_workbook):
    rows = [
        ["ردیف", "واحد", "جنسیت", "سن", "حقوق", "تاریخ", "رضایت"],
        [1, "A", "زن", 30, 10_000_000, "1403/01/05", 4],
        [2, "A", "مرد", 40, 12_000_000, "1403/01/20", 3],
        [3, "B", "زن", 35, 14_000_000, "1403/02/02", 5],
        [4, "B", "مرد", 50, 15_000_000, "1403/02/10", 2],
        [5, "A", "زن", 28, 9_000_000, "1403/02/18", 4],
        [6, "B", "مرد", 45, "۲ میلیون", "1403/03/01", 3],
        [7, "C", "زن", 32, 11_000_000, "1403/03/12", 5],
        [8, "C", "مرد", 38, 13_000_000, "1403/03/20", None],
        [9, "A", "مرد", 41, 12_000_000, "1403/04/02", 3],
        [9, "A", "مرد", 41, 12_000_000, "1403/04/02", 3],
    ]
    return make_workbook("office.xlsx", {"کارکنان": rows})


@pytest.fixture()
def office_analysis(settings, office_workbook):
    return ExcelIntelligenceWorkflow(settings).analyze_only(
        office_workbook, goal="مقایسه حقوق بر اساس واحد"
    )


def test_scale_units_are_applied_inside_text():
    assert parse_number("حدود ۲ میلیون تومان").value == Decimal("2000000")
    assert parse_number("۳.۵ میلیارد ریال").value == Decimal("3500000000")
    assert parse_number("۱۲ هزار و ۵۰۰").value == Decimal("12500")
    assert parse_number("۲ همت").value == Decimal("2000000000000")


def test_generic_profile_uses_real_column_content(office_analysis):
    generic = office_analysis.generic
    assert generic is not None
    assert generic.row_count == 10
    assert generic.column_count == 7
    columns = {column.header: column for column in generic.columns}
    assert columns["ردیف"].is_identifier
    assert columns["حقوق"].kind.value == "number"
    assert columns["تاریخ"].kind.value == "date"
    assert columns["واحد"].kind.value == "text"
    assert columns["واحد"].distinct == 3
    assert columns["واحد"].count == 10


def test_descriptive_statistics_are_exact_decimals(office_analysis):
    salary = next(column for column in office_analysis.generic.columns if column.header == "حقوق")
    assert salary.total == Decimal("110000000")
    assert salary.mean == Decimal("11000000")
    assert salary.median == Decimal("12000000")
    assert salary.minimum == Decimal("2000000")
    assert salary.maximum == Decimal("15000000")
    assert salary.q1 == Decimal("10250000")
    assert salary.q3 == Decimal("12750000")
    assert salary.iqr == Decimal("2500000")


def test_frequency_groups_crosstab_trend_and_correlation(office_analysis):
    generic = office_analysis.generic
    frequencies = {table.header: table for table in generic.frequencies}
    assert frequencies["واحد"].total == 10
    assert {row.value: row.count for row in frequencies["واحد"].rows} == {"A": 5, "B": 3, "C": 2}

    group = next(item for item in generic.groups if item.group_header == "واحد" and item.measure_header == "حقوق")
    values = {item.label: item for item in group.groups}
    assert values["A"].count == 5
    assert values["A"].total == Decimal("55000000")
    assert values["B"].total == Decimal("31000000")
    assert values["C"].total == Decimal("24000000")

    crosstab = generic.crosstabs[0]
    assert crosstab.total == 10
    assert crosstab.row_header in {"واحد", "جنسیت"}

    trend = generic.trends[0]
    assert trend.date_header == "تاریخ"
    assert trend.measure_header == "حقوق"
    assert len(trend.points) >= 2
    assert trend.calendar == "jalali"
    assert trend.points[0].period_label.startswith("1403/")
    assert "1402/12" not in {point.period_label for point in trend.points}
    assert trend.direction in {"افزایشی", "کاهشی", "بدون تغییر"}

    pair = next(item for item in generic.correlations if {item.header_a, item.header_b} == {"سن", "حقوق"})
    assert pair.pearson is not None
    assert pair.pearson > 0
    assert pair.count == 10


def test_quality_and_outliers_do_not_delete_rows(office_analysis):
    generic = office_analysis.generic
    quality = generic.quality
    assert quality.missing_cells == 1
    assert quality.duplicate_rows == 1
    assert "ردیف" in quality.identifier_columns
    assert any(item.header == "حقوق" and item.value == Decimal("2000000") for item in generic.outliers)


def test_generic_facts_and_offline_summary(settings, office_analysis):
    facts, meta = build_facts(office_analysis, settings)
    assert facts["dataset_profile"]["table"]["rows"] == 10
    assert facts["dataset_profile"]["quality"]["missing_cells"] == 1
    assert "11000000" in allowed_numbers(office_analysis)
    summary = offline_summary(facts)
    assert "10 ردیف" in summary["summary"]
    assert "علت‌ومعلولی" in " ".join(summary["findings"])
    assert meta["chars"] > 0


def test_generic_end_to_end_report_keeps_source_untouched(settings, office_workbook, tmp_path):
    before = sha256(office_workbook.read_bytes()).hexdigest()
    output = tmp_path / "out" / "office_report.xlsx"
    result = Coordinator(settings, console=False).run(
        office_workbook, use_ai=False, output=output, goal="تحلیل کیفیت داده"
    )
    assert str(result.output_path) == str(output)
    assert sha256(office_workbook.read_bytes()).hexdigest() == before

    workbook = openpyxl.load_workbook(output, data_only=True)
    assert {
        "شناسنامه داده", "آمار توصیفی", "توزیع فراوانی", "تحلیل گروه‌ها",
        "روند و همبستگی", "روش‌شناسی آماری",
    } <= set(workbook.sheetnames)
    assert "حسابرسی اعداد" not in workbook.sheetnames
    assert "تطبیق تعهدات" not in workbook.sheetnames
    names = workbook.sheetnames
    assert names.index("شناسنامه داده") < names.index("داده خام")
    workbook.close()



def test_total_row_is_excluded_from_statistics(settings, office_workbook, tmp_path):
    workbook = load_workbook(office_workbook)
    sheet = workbook.active
    sheet.append(["جمع", None, None, None, 110_000_000, None, None])
    target = tmp_path / "with_total.xlsx"
    workbook.save(target)
    workbook.close()

    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(target)
    generic = analysis.generic
    assert generic.row_count == 10
    salary = next(column for column in generic.columns if column.header == "حقوق")
    assert salary.count == 10
    assert salary.total == Decimal("110000000")


def test_gregorian_dates_keep_gregorian_trend_labels(settings, make_workbook):
    rows = [["کد", "شرح", "تاریخ", "مقدار"]]
    for index in range(6):
        rows.append(
            [
                100 + index,
                f"رکورد {index}",
                f"2022-{index + 1:02d}-15",
                1_000_000 + index * 500_000,
            ]
        )
    workbook = make_workbook("gregorian.xlsx", {"داده‌ها": rows})
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(workbook)
    date_column = next(column for column in analysis.generic.columns if column.header == "تاریخ")
    assert date_column.date_calendar == "gregorian"
    assert date_column.date_min == "2022-01-15"
    trend = next(item for item in analysis.generic.trends if item.date_header == "تاریخ")
    assert trend.calendar == "gregorian"
    assert trend.points[0].period_label == "2022-01"
    assert all(not point.period_label.startswith("140") for point in trend.points)
