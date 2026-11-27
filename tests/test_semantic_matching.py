"""Prompt 03A: semantic source matching and period alignment.

Every workbook here is a real .xlsx file built by openpyxl; nothing is mocked.
The tests cover the seven required cases plus the period parser and the
audit-only report sheet.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.folder_workflow import FolderWorkflow
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.excel.period_alignment import (
    compare_period_sets,
    parse_period,
    parse_periods,
)


def _analyse(settings, path):
    return ExcelIntelligenceWorkflow(settings).analyze_only(path)


def _matches_with(report, header: str):
    return [
        match
        for match in report.matches
        if any(member.header == header for member in match.members)
    ]


def _probable_with(report, header: str):
    return [
        match
        for match in report.probable
        if any(member.header == header for member in match.members)
    ]


def _rejection(report, left: str, right: str):
    return [
        item
        for item in report.rejections
        if {item.left.header, item.right.header} == {left, right}
    ]


# --------------------------------------------------------------------------- #
# 1. Different names for the same indicator
# --------------------------------------------------------------------------- #
def test_different_names_for_the_same_indicator_are_matched(settings, make_workbook):
    path = make_workbook(
        "different-names.xlsx",
        {
            "خلاصه": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "جزئیات": [
                ["گزارش تیر ۱۴۰۳"],
                [],
                ["شعبه", "Case Count"],
                ["الف", 12],
                ["ب", 22],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    matches = _matches_with(report, "Case Count")
    assert len(matches) == 1
    match = matches[0]
    assert match.status.value == "matched"
    assert match.comparable is True
    assert {member.header for member in match.members} == {"تعداد پرونده", "Case Count"}
    assert "یکسان" in match.evidence[0] or any("یکسان" in text for text in match.evidence)


# --------------------------------------------------------------------------- #
# 2. Similar names for different indicators
# --------------------------------------------------------------------------- #
def test_similar_names_for_different_indicators_are_rejected(settings, make_workbook):
    path = make_workbook(
        "similar-names.xlsx",
        {
            "ورودی": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده ورودی"],
                ["الف", 10],
                ["ب", 20],
            ],
            "خروجی": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده خروجی"],
                ["الف", 7],
                ["ب", 8],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    assert not _matches_with(report, "تعداد پرونده ورودی")
    assert not _probable_with(report, "تعداد پرونده ورودی")
    rejections = _rejection(report, "تعداد پرونده ورودی", "تعداد پرونده خروجی")
    assert rejections, "the similar-name pair was not explained as a rejection"
    assert "یکسان نیستند" in rejections[0].reason


def test_aggregation_labels_are_not_matched_to_the_base_measure(settings, make_workbook):
    path = make_workbook(
        "aggregation-labels.xlsx",
        {
            "پایه": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "مبلغ"],
                ["الف", 1000],
                ["ب", 2000],
            ],
            "کل": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "مبلغ کل"],
                ["الف", 3000],
                ["ب", 3000],
            ],
            "میانگین": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "میانگین مبلغ"],
                ["الف", 1500],
                ["ب", 2000],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    assert not _matches_with(report, "مبلغ")
    assert not _matches_with(report, "مبلغ کل")
    assert not _matches_with(report, "میانگین مبلغ")
    assert _rejection(report, "مبلغ", "مبلغ کل")
    assert _rejection(report, "مبلغ", "میانگین مبلغ")


# --------------------------------------------------------------------------- #
# 3. Different period formats
# --------------------------------------------------------------------------- #
def test_different_period_formats_align_and_are_comparable(settings, make_workbook):
    path = make_workbook(
        "period-formats.xlsx",
        {
            "اول": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "دوم": [
                ["گزارش 1403/04"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 11],
                ["ب", 21],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    matches = _matches_with(report, "تعداد پرونده")
    assert len(matches) == 1
    match = matches[0]
    assert match.comparable is True
    assert match.period_relation == "equivalent"
    assert match.period_status == "detected"


# --------------------------------------------------------------------------- #
# 4. Persian and Gregorian date references
# --------------------------------------------------------------------------- #
def test_persian_and_gregorian_dates_are_the_same_period(settings, make_workbook):
    path = make_workbook(
        "two-calendars.xlsx",
        {
            "شمسی": [
                ["گزارش 1403/04/15"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "میلادی": [
                ["گزارش 2024-07-05"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 12],
                ["ب", 22],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    matches = _matches_with(report, "تعداد پرونده")
    assert len(matches) == 1
    assert matches[0].comparable is True
    assert matches[0].period_relation == "equivalent"


# --------------------------------------------------------------------------- #
# 5. Different units
# --------------------------------------------------------------------------- #
def test_different_unit_dimensions_are_rejected(settings, make_workbook):
    path = make_workbook(
        "different-units.xlsx",
        {
            "ریالی": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "مبلغ (ریال)"],
                ["الف", 1000],
                ["ب", 2000],
            ],
            "نفری": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "مبلغ (نفر)"],
                ["الف", 3],
                ["ب", 4],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    assert not _matches_with(report, "مبلغ (ریال)")
    rejections = _rejection(report, "مبلغ (ریال)", "مبلغ (نفر)")
    assert rejections, "the incompatible-unit pair was not rejected"
    assert "از یک جنس نیست" in rejections[0].reason


def test_unit_scale_difference_is_matched_but_not_comparable(settings, make_workbook):
    path = make_workbook(
        "unit-scale.xlsx",
        {
            "میلیونی": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "مبلغ (میلیون ریال)"],
                ["الف", 1],
                ["ب", 2],
            ],
            "ریالی": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "مبلغ (ریال)"],
                ["الف", 1_000_000],
                ["ب", 2_000_000],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    matches = _matches_with(report, "مبلغ (میلیون ریال)")
    assert len(matches) == 1
    match = matches[0]
    assert match.status.value == "matched"
    assert match.comparable is False
    assert match.unit_state == "scalable"
    assert "ضریب" in match.reason


# --------------------------------------------------------------------------- #
# 6. Missing period information
# --------------------------------------------------------------------------- #
def test_missing_period_is_probable_not_matched(settings, make_workbook):
    path = make_workbook(
        "missing-period.xlsx",
        {
            "اول": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "دوم": [
                ["شعبه", "تعداد پرونده"],
                ["الف", 11],
                ["ب", 21],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    assert not _matches_with(report, "تعداد پرونده")
    probable = _probable_with(report, "تعداد پرونده")
    assert len(probable) == 1
    assert probable[0].comparable is False
    assert probable[0].period_status == "partial"
    assert "دوره" in probable[0].reason
    second = next(member for member in probable[0].members if member.sheet == "دوم")
    assert second.period_status == "not_detected"


def test_month_title_without_a_year_is_not_invented(settings, make_workbook):
    path = make_workbook(
        "month-no-year.xlsx",
        {
            "داده یک": [
                ["مرداد"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "داده دو": [
                ["شهریور"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 11],
                ["ب", 21],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    assert not _matches_with(report, "تعداد پرونده")
    probable = _probable_with(report, "تعداد پرونده")
    assert len(probable) == 1
    assert probable[0].comparable is False
    assert probable[0].period_status == "not_detected"
    assert all(member.periods == [] for member in probable[0].members)
    assert all(member.period_status == "not_detected" for member in probable[0].members)
    assert "مشخص نیست" in probable[0].reason


# --------------------------------------------------------------------------- #
# 7. Ambiguous semantic matches
# --------------------------------------------------------------------------- #
def test_ambiguous_matches_keep_all_candidates(settings, make_workbook):
    path = make_workbook(
        "ambiguous.xlsx",
        {
            "کل": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 30],
                ["ب", 40],
            ],
            "مختومه": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده مختومه"],
                ["الف", 10],
                ["ب", 15],
            ],
            "جریان": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده در جریان"],
                ["الف", 20],
                ["ب", 25],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    assert not _matches_with(report, "تعداد پرونده")
    anchors = [
        item
        for item in report.ambiguities
        if item.members[0].header == "تعداد پرونده"
    ]
    assert len(anchors) == 1
    candidate_headers = {candidate.field.header for candidate in anchors[0].candidates}
    assert candidate_headers == {"تعداد پرونده مختومه", "تعداد پرونده در جریان"}
    assert anchors[0].comparable is False
    assert any("گزینه" in warning for warning in report.warnings)


# --------------------------------------------------------------------------- #
# Equivalent categories and shared identifiers
# --------------------------------------------------------------------------- #
def test_equivalent_categories_with_different_names_become_probable(settings, make_workbook):
    values = ["سرقت", "سرقت", "ضرب", "ضرب", "جعل", "جعل"]
    path = make_workbook(
        "category-values.xlsx",
        {
            "اول": [
                ["گزارش تیر 1403"],
                [],
                ["نوع جرم", "تعداد"],
                *[[value, 1] for value in values],
            ],
            "دوم": [
                ["گزارش تیر 1403"],
                [],
                ["دسته بزه", "تعداد"],
                *[[value, 1] for value in values],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    pair_headers = {"نوع جرم", "دسته بزه"}
    assert not any(
        {member.header for member in match.members} == pair_headers
        for match in report.matches
    )
    records = [
        item
        for item in report.probable
        if {member.header for member in item.members} == pair_headers
    ]
    assert records, "the value-overlap category pair was not reported as probable"
    assert records[0].comparable is False
    assert "هم‌پوشانی" in " ".join(records[0].evidence)


def test_configured_synonym_maps_equivalent_category_names(settings, make_workbook):
    settings.synonyms = {"دسته بزه": "نوع جرم"}
    path = make_workbook(
        "synonym-categories.xlsx",
        {
            "اول": [
                ["گزارش تیر 1403"],
                [],
                ["نوع جرم", "تعداد"],
                ["سرقت", 10],
                ["ضرب", 20],
            ],
            "دوم": [
                ["گزارش تیر 1403"],
                [],
                ["دسته بزه", "تعداد"],
                ["سرقت", 12],
                ["ضرب", 22],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    matches = [
        match
        for match in report.matches
        if {member.header for member in match.members} == {"نوع جرم", "دسته بزه"}
    ]
    assert matches, "the configured synonym did not map the equivalent categories"
    assert matches[0].comparable is True


def test_reordered_and_extra_columns_still_match_indicators(settings, make_workbook):
    path = make_workbook(
        "different-structures.xlsx",
        {
            "اول": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده", "مبلغ"],
                ["الف", 10, 100],
                ["ب", 20, 200],
            ],
            "دوم": [
                ["گزارش 1403/04"],
                [],
                ["شرح", "شعبه", "مبلغ", "تعداد پرونده"],
                ["x", "الف", 110, 12],
                ["y", "ب", 210, 22],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    for header in ("تعداد پرونده", "مبلغ"):
        matches = _matches_with(report, header)
        assert len(matches) == 1, header
        match = matches[0]
        assert match.comparable is True, header
        assert {member.sheet for member in match.members} == {"اول", "دوم"}, header
        evidence = " ".join(match.evidence)
        assert "هم‌پوشانی برچسب‌های سطر" in evidence, header
        assert "عنوان جدول‌ها هم‌خوان" in evidence, header


def test_shared_identifier_column_is_recorded_as_evidence(settings, make_workbook):
    path = make_workbook(
        "shared-identifier.xlsx",
        {
            "اول": [
                ["گزارش تیر 1403"],
                [],
                ["شماره پرونده", "مبلغ"],
                ["1", 100],
                ["2", 200],
                ["3", 300],
            ],
            "دوم": [
                ["گزارش تیر 1403"],
                [],
                ["شماره پرونده", "مبلغ"],
                ["1", 110],
                ["2", 210],
                ["3", 310],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    amount_matches = [
        match
        for match in report.matches
        if {member.header for member in match.members} == {"مبلغ"}
    ]
    assert len(amount_matches) == 1
    evidence = " ".join(amount_matches[0].evidence)
    assert "ستون شناسه مشترک" in evidence
    assert "هم‌پوشانی مقادیر" in evidence


# --------------------------------------------------------------------------- #
# Period alignment parser
# --------------------------------------------------------------------------- #
def test_period_parser_supports_formats_and_calendars():
    cases = {
        "تیر 1403": ("jalali", "month"),
        "تیرماه ۱۴۰۳": ("jalali", "month"),
        "1403/04": ("jalali", "month"),
        "1403-04": ("jalali", "month"),
        "2024-07": ("gregorian", "month"),
        "July 2024": ("gregorian", "month"),
        "1403/04/15": ("jalali", "day"),
        "2024-07-05": ("gregorian", "day"),
        "1403": ("jalali", "year"),
        "سه‌ماهه اول 1403": ("jalali", "range"),
        "بهار 1403": ("jalali", "range"),
        "نیمه دوم 1403": ("jalali", "range"),
        "نیمسال دوم 1403": ("jalali", "range"),
        "مرداد و شهریور 1403": ("jalali", "range"),
        "مرداد، شهریور 1403": ("jalali", "range"),
        "January and February 2024": ("gregorian", "range"),
        "q2-2024": ("gregorian", "range"),
        "از 1403/01/01 تا 1403/03/31": ("jalali", "range"),
        "1403/01/01-1403/03/31": ("jalali", "range"),
        "2024-01-01-2024-03-31": ("gregorian", "range"),
        "از 1403/01/01 تا 2024-06-20": ("mixed", "range"),
        "1403-1404": ("jalali", "range"),
        "2024-2025": ("gregorian", "range"),
        "2024-04-01 to 2024-06-30": ("gregorian", "range"),
    }
    for text, (calendar, granularity) in cases.items():
        period = parse_period(text)
        assert period is not None, text
        assert period.calendar == calendar, text
        assert period.granularity == granularity, text
        assert period.detected, text

    assert parse_period("ردیف 12") is None
    assert parse_period("کد پرسنلی 1234") is None
    assert parse_period("گزارش 1234") is None
    assert parse_period("گزارش 1403").year == 1403
    assert parse_period("سه ماهه 1403") is None
    assert parse_period("شش ماهه 1403") is None
    assert parse_period("مرداد و مهر 1403") is None
    assert parse_period("quarter 2024") is None
    assert parse_period("تیر 2024") is None
    assert parse_period("July 1403") is None
    assert parse_periods([]) == []
    same = parse_periods(["تیر 1403", "1403/04", "تیرماه ۱۴۰۳"])
    assert len({item.key for item in same}) == 1


def test_period_sets_require_full_coverage_for_equivalence():
    mordad = parse_period("مرداد 1403")
    shahrivar = parse_period("شهریور 1403")
    year = parse_period("1403")
    assert compare_period_sets([mordad, shahrivar], [mordad]).relation == "overlapping"
    assert (
        compare_period_sets([mordad, shahrivar], [mordad, shahrivar]).relation
        == "equivalent"
    )
    assert compare_period_sets([mordad], [shahrivar]).relation == "incompatible"
    assert compare_period_sets([year, mordad], [mordad]).relation == "overlapping"
    two_months = parse_period("مرداد و شهریور 1403")
    assert two_months is not None and two_months.granularity == "range"
    assert (
        compare_period_sets([two_months], [mordad, shahrivar]).relation
        == "equivalent"
    )


# --------------------------------------------------------------------------- #
# Audit-only report sheet and clean delivery
# --------------------------------------------------------------------------- #
def test_matching_report_is_audit_only(settings, make_workbook, tmp_path):
    path = make_workbook(
        "audit-matching.xlsx",
        {
            "خلاصه": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "جزئیات": [
                ["گزارش 1403/04"],
                [],
                ["شعبه", "Case Count"],
                ["الف", 12],
                ["ب", 22],
            ],
        },
    )
    audit_output = tmp_path / "audit.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=audit_output, profile="audit"
    )
    workbook = openpyxl.load_workbook(audit_output)
    assert "تطبیق معنایی منابع" in workbook.sheetnames
    text = " ".join(
        str(cell.value)
        for row in workbook["تطبیق معنایی منابع"].iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "دوره:" in text
    assert "واحد:" in text
    workbook.close()

    clean_output = tmp_path / "clean.xlsx"
    Coordinator(settings, console=False).run(
        path, use_ai=False, output=clean_output, profile="clean"
    )
    workbook = openpyxl.load_workbook(clean_output)
    assert "تطبیق معنایی منابع" not in workbook.sheetnames
    workbook.close()


def test_matching_is_serialized_with_period_and_evidence(settings, make_workbook):
    path = make_workbook(
        "serialized.xlsx",
        {
            "خلاصه": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "جزئیات": [
                ["گزارش 1403/04"],
                [],
                ["شعبه", "Case Count"],
                ["الف", 12],
                ["ب", 22],
            ],
        },
    )
    payload = _analyse(settings, path).to_dict()["semantic_matching"]
    assert payload["counts"]["matched"] >= 1
    match = next(item for item in payload["matches"] if item["comparable"])
    member = match["members"][0]
    assert member["source"] == "serialized.xlsx"
    assert member["sheet"]
    assert member["block"]
    assert member["periods"]
    assert match["evidence"]
    assert match["period_relation"] == "equivalent"


# --------------------------------------------------------------------------- #
# Cross-file matching in folder mode
# --------------------------------------------------------------------------- #
def test_folder_matching_maps_fields_across_files(settings, tmp_path):
    folder = tmp_path / "folder"
    folder.mkdir()
    for name, title, first_value in (
        ("a.xlsx", "گزارش تیر 1403", 10),
        ("b.xlsx", "گزارش 1403/04", 12),
    ):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "داده"
        sheet.append([title])
        sheet.append([])
        sheet.append(["شعبه", "تعداد پرونده"])
        sheet.append(["الف", first_value])
        sheet.append(["ب", first_value + 1])
        workbook.save(folder / name)

    result = FolderWorkflow(settings, progress=lambda message: None).run(
        folder,
        use_ai=False,
        output=tmp_path / "folder-report.xlsx",
        profile="clean",
    )
    assert result.matching is not None
    counts = result.matching.counts()
    assert counts["matched"] >= 2
    cross_file = [
        match
        for match in result.matching.matches
        if {member.source for member in match.members} == {"a.xlsx", "b.xlsx"}
    ]
    assert cross_file, "no field was mapped across the two files"
    assert any(match.comparable for match in cross_file)

    workbook = openpyxl.load_workbook(result.output_path)
    assert "تطبیق شاخص‌ها بین منابع" in workbook.sheetnames
    workbook.close()


def test_conflicting_periods_in_one_field_prevent_a_match(settings, make_workbook):
    path = make_workbook(
        "conflicting-period.xlsx",
        {
            "متناقض": [
                ["گزارش 1403"],
                [],
                ["شعبه", "تعداد پرونده 1404"],
                ["الف", 10],
                ["ب", 20],
            ],
            "روشن": [
                ["گزارش 1404"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 12],
                ["ب", 22],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    assert not _matches_with(report, "تعداد پرونده 1404")
    probable = _probable_with(report, "تعداد پرونده 1404")
    assert len(probable) == 1
    assert probable[0].comparable is False
    assert probable[0].period_status == "conflicting"
    assert "ناسازگار" in probable[0].reason


def test_month_titles_in_a_year_sheet_are_not_comparable(settings, make_workbook):
    path = make_workbook(
        "months.xlsx",
        {
            "1404": [
                ["مرداد"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
            ],
            "1404-b": [
                ["شهریور"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 11],
                ["ب", 21],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    matches = _matches_with(report, "تعداد پرونده")
    assert len(matches) == 1
    assert matches[0].comparable is False
    assert matches[0].period_relation == "incompatible"
    assert "مرداد 1404" in matches[0].period_explanation
    assert "شهریور 1404" in matches[0].period_explanation


def test_two_month_blocks_in_one_year_sheet_keep_their_own_periods(
    settings, make_workbook
):
    path = make_workbook(
        "two-month-blocks.xlsx",
        {
            "1404": [
                ["مرداد"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
                [],
                [],
                ["شهریور"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 11],
                ["ب", 21],
            ],
        },
    )
    report = _analyse(settings, path).semantic_matching
    matches = _matches_with(report, "تعداد پرونده")
    assert len(matches) == 1
    assert matches[0].comparable is False
    assert matches[0].period_relation == "incompatible"
    assert "مرداد 1404" in matches[0].period_explanation
    assert "شهریور 1404" in matches[0].period_explanation


def test_unit_adjectives_are_recognized_without_a_separate_unit_column():
    from decimal import Decimal

    from local_office_ai.excel.semantic_matching import canonical_unit

    assert canonical_unit("مبلغ ریالی").canonical == "IRR"
    assert canonical_unit("مبلغ تومانی").canonical == "TOMAN"
    assert canonical_unit("تعداد نفری").canonical == "PERSON"
    assert canonical_unit("سهم درصدی").canonical == "PERCENT"
    assert canonical_unit("تعداد پرونده").known is False
    hemmat = canonical_unit("مبلغ (همت)")
    assert hemmat.canonical == "TOMAN"
    assert hemmat.scale == Decimal("1000000000000")
