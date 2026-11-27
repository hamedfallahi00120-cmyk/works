"""Prompt 02A: semantic classification and data-quality validation.

Every workbook is a real .xlsx file created by openpyxl.  The tests cover the
eight required checks: numeric identifiers versus measures, date fields versus
measures, missing values, invalid numeric content, negative values, unit
inconsistencies, conflicting period labels and duplicate identifiers.
"""

from __future__ import annotations

import openpyxl
import pytest

from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.models import DataRole
from local_office_ai.excel.data_quality import (
    REQUIRED_CHECK_CODES,
    QualityStatus,
)
from local_office_ai.excel.semantic_classification import (
    SEMANTIC_ROLE_FA,
    SemanticRole,
)


def _analyze(settings, path):
    return ExcelIntelligenceWorkflow(settings).analyze_only(path)


def _fields(analysis, sheet: str | None = None) -> dict:
    fields = {}
    for block in analysis.semantics.blocks:
        if sheet and block.sheet != sheet:
            continue
        for field in block.columns:
            fields.setdefault(field.header, field)
    return fields


def _findings(analysis, code: str):
    return [item for item in analysis.quality.findings if item.code == code]


# --------------------------------------------------------------------------- #
# 1. Numeric identifiers versus numeric measures
# --------------------------------------------------------------------------- #
def test_numeric_identifiers_never_become_measures(settings, make_workbook):
    path = make_workbook(
        "identifiers.xlsx",
        {
            "داده": [
                ["کد پستی", "شناسه پرسنلی", "مشخصه", "مبلغ", "مبلغ کل", "تعداد"],
                [1234567890, 2184, 9000000001, 1000, 1234567890, 5],
                [1234567891, 2079, 9000000002, 2000, 2234567891, 7],
                [1234567892, 2047, 9000000003, 3000, 3234567892, 9],
                [1234567893, 2185, 9000000004, 4000, 4234567893, 4],
            ]
        },
    )
    analysis = _analyze(settings, path)
    fields = _fields(analysis, "داده")

    assert fields["کد پستی"].role is SemanticRole.IDENTIFIER
    assert fields["شناسه پرسنلی"].role is SemanticRole.IDENTIFIER
    # A fixed-width numeric code without an identifier keyword is still a code.
    assert fields["مشخصه"].role is SemanticRole.IDENTIFIER
    assert fields["مبلغ"].role is SemanticRole.MEASURE
    # A wide monetary value is not turned into a code just because it has 10 digits.
    assert fields["مبلغ کل"].role is SemanticRole.MEASURE
    assert fields["تعداد"].role is SemanticRole.MEASURE

    profiles = {column.header: column for column in analysis.column_profiles}
    for header in ("کد پستی", "شناسه پرسنلی", "مشخصه"):
        assert profiles[header].statistical_role is DataRole.IDENTIFIER
    measures = {column.header for column in analysis.intelligence.sheet("داده").measures}
    assert {"مبلغ", "مبلغ کل", "تعداد"} <= measures
    assert not {"کد پستی", "شناسه پرسنلی", "مشخصه"} & measures

    labels = {item.label for item in analysis.calculations}
    for header in ("کد پستی", "شناسه پرسنلی", "مشخصه"):
        for prefix in ("جمع ", "میانگین ", "میانه "):
            assert f"{prefix}{header}" not in labels, (header, labels)


def test_identifier_evidence_is_recorded(settings, make_workbook):
    path = make_workbook(
        "code_evidence.xlsx",
        {
            "داده": [
                ["کد پستی", "مبلغ"],
                [1234567890, 1000],
                [1234567891, 2000],
                [1234567892, 3000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    field = _fields(analysis, "داده")["کد پستی"]
    assert field.evidence, "شاهد طبقه‌بندی شناسه ثبت نشد"
    assert any("طول ثابت" in item or "شناسه" in item for item in field.evidence)
    assert any("برگه" in item for item in field.evidence)


def test_administrative_vehicle_region_and_record_codes_stay_identifiers(
    settings, make_workbook
):
    path = make_workbook(
        "administrative_codes.xlsx",
        {
            "داده": [
                ["کد اداری", "کد خودرو", "کد منطقه", "شماره پرونده", "مبلغ"],
                [1001, 2184, 31, 9001, 1000],
                [1002, 2079, 32, 9002, 2000],
                [1003, 2047, 33, 9003, 3000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    fields = _fields(analysis, "داده")
    identifiers = ("کد اداری", "کد خودرو", "کد منطقه", "شماره پرونده")
    for header in identifiers:
        assert fields[header].role is SemanticRole.IDENTIFIER, header
    assert fields["مبلغ"].role is SemanticRole.MEASURE

    measures = {column.header for column in analysis.intelligence.sheet("داده").measures}
    assert "مبلغ" in measures
    assert not set(identifiers) & measures


def test_english_and_arabic_digit_identifiers_are_protected(settings, make_workbook):
    path = make_workbook(
        "english_codes.xlsx",
        {
            "داده": [
                ["postal code", "amount"],
                ["١٢٣٤٥٦٧٨٩٠", 1000],
                ["١٢٣٤٥٦٧٨٩١", 2000],
                ["١٢٣٤٥٦٧٨٩٢", 3000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    fields = _fields(analysis, "داده")
    assert fields["postal code"].role is SemanticRole.IDENTIFIER
    assert fields["amount"].role is SemanticRole.MEASURE


# --------------------------------------------------------------------------- #
# 2. Date fields versus numeric measures
# --------------------------------------------------------------------------- #
def test_date_and_temporal_fields_never_become_measures(settings, make_workbook):
    path = make_workbook(
        "dates.xlsx",
        {
            "داده": [
                ["تاریخ ثبت", "سال", "ماه", "مبلغ"],
                ["1404/01/15", 1404, 1, 1000],
                ["1404/02/03", 1404, 2, 2000],
                ["1404/02/10", 1405, 3, 3000],
                ["1404/03/01", 1405, 4, 4000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    fields = _fields(analysis, "داده")

    assert fields["تاریخ ثبت"].role is SemanticRole.TEMPORAL_KEY
    assert fields["سال"].role is SemanticRole.TEMPORAL_KEY
    assert fields["ماه"].role is SemanticRole.TEMPORAL_KEY
    assert fields["مبلغ"].role is SemanticRole.MEASURE

    labels = {item.label for item in analysis.calculations}
    for header in ("تاریخ ثبت", "سال", "ماه"):
        for prefix in ("جمع ", "میانگین ", "میانه "):
            assert f"{prefix}{header}" not in labels, (header, labels)


# --------------------------------------------------------------------------- #
# 3. Missing values
# --------------------------------------------------------------------------- #
def test_missing_values_are_recorded_with_their_location(settings, make_workbook):
    path = make_workbook(
        "missing.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ", "تعداد"],
                ["الف", 1000, 5],
                ["ب", None, 7],
                ["ج", 3000, 9],
                ["د", None, 4],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "missing_values")
    assert findings, "مقدار گمشده ثبت نشد"
    finding = findings[0]
    assert finding.status in {QualityStatus.WARNING, QualityStatus.FAIL}
    assert finding.count == 2
    assert {"B3", "B5"} <= set(finding.samples)
    assert finding.sheet == "داده"


# --------------------------------------------------------------------------- #
# 4. Invalid numeric content
# --------------------------------------------------------------------------- #
def test_invalid_numeric_content_is_reported_not_repaired(settings, make_workbook):
    path = make_workbook(
        "invalid.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ", "تعداد"],
                ["الف", 1000, 5],
                ["ب", "#DIV/0!", 7],
                ["ج", "12abc", 9],
                ["د", "نامشخص", 4],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "invalid_numeric")
    assert findings, "مقدار عددی نامعتبر ثبت نشد"
    finding = findings[0]
    assert finding.status in {QualityStatus.WARNING, QualityStatus.FAIL}
    assert finding.count >= 2
    assert "B3" in finding.samples and "B4" in finding.samples

    unexpected = _findings(analysis, "unexpected_text")
    assert unexpected, "متن غیرمنتظره در ستون عددی ثبت نشد"
    assert "B5" in unexpected[0].samples

    # The questionable cells stay exactly as they were in the source file.
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["داده"]
    assert sheet["B3"].value == "#DIV/0!"
    assert sheet["B4"].value == "12abc"
    assert sheet["B5"].value == "نامشخص"
    workbook.close()


def test_all_text_measure_column_is_reported_as_unexpected_text(
    settings, make_workbook
):
    path = make_workbook(
        "all_text_measure.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ", "تعداد", "توضیحات مبلغ"],
                ["الف", "نامشخص", 5, "پرداخت نشد"],
                ["ب", "نامعلوم", 7, "در انتظار"],
                ["ج", "---", 9, "بدون سند"],
            ]
        },
    )
    analysis = _analyze(settings, path)
    unexpected = _findings(analysis, "unexpected_text")
    assert unexpected, "متن غیرمنتظره در ستون مبلغ ثبت نشد"
    assert unexpected[0].column == "B"
    assert unexpected[0].count == 3
    # «توضیحات مبلغ» is descriptive, not a broken numeric column.
    assert all(finding.column != "D" for finding in unexpected)


# --------------------------------------------------------------------------- #
# 5. Negative values
# --------------------------------------------------------------------------- #
def test_negative_values_are_flagged_only_where_they_are_invalid(settings, make_workbook):
    path = make_workbook(
        "negative.xlsx",
        {
            "داده": [
                ["شرح", "تعداد", "مانده"],
                ["الف", 5, 2],
                ["ب", -3, -1],
                ["ج", 7, 4],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "negative_values")
    assert findings, "مقدار منفی نامعتبر ثبت نشد"
    assert any(finding.column == "B" for finding in findings)
    assert all(finding.column != "C" for finding in findings)  # «مانده» may legitimately be negative
    assert all(finding.block for finding in findings)


# --------------------------------------------------------------------------- #
# 6. Inconsistent units
# --------------------------------------------------------------------------- #
def test_inconsistent_units_are_reported(settings, make_workbook):
    path = make_workbook(
        "units.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ"],
                ["الف", "۲ میلیون"],
                ["ب", "۳ میلیارد"],
                ["ج", "۱ میلیون"],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "inconsistent_units")
    assert findings, "ناسازگاری واحد ثبت نشد"
    finding = findings[0]
    assert finding.status is QualityStatus.WARNING
    assert any("مقیاس" in item for item in finding.evidence)


# --------------------------------------------------------------------------- #
# 7. Conflicting period labels
# --------------------------------------------------------------------------- #
def test_conflicting_period_labels_fail_the_check(settings, make_workbook):
    path = make_workbook(
        "گزارش عملکرد 1403.xlsx",
        {
            "فروردین 1404": [
                ["شرح", "مبلغ"],
                ["الف", 1000],
                ["ب", 2000],
                ["ج", 3000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "conflicting_periods")
    assert findings, "تعارض دوره ثبت نشد"
    finding = findings[0]
    assert finding.status is QualityStatus.FAIL
    assert finding.count >= 1
    assert finding.evidence


def test_scalar_label_conflicting_with_a_column_period_fails(settings, make_workbook):
    path = make_workbook(
        "گزارش 1403.xlsx",
        {
            "داده": [
                ["شرح", "فروردین 1404"],
                ["الف", 1000],
                ["ب", 2000],
                ["ج", 3000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "conflicting_periods")
    assert findings, "تعارض دوره فایل با سرستون ثبت نشد"
    finding = findings[0]
    assert finding.status is QualityStatus.FAIL
    assert any("column_header" in item for item in finding.evidence)


# --------------------------------------------------------------------------- #
# 8. Duplicate identifiers
# --------------------------------------------------------------------------- #
def test_duplicate_unique_identifiers_fail_the_check(settings, make_workbook):
    path = make_workbook(
        "duplicates.xlsx",
        {
            "داده": [
                ["شناسه ملی", "مبلغ"],
                [1234567890, 1000],
                [1234567890, 2000],
                [1234567891, 3000],
                [1234567892, 4000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "duplicate_identifiers")
    assert findings, "شناسه تکراری ثبت نشد"
    finding = findings[0]
    assert finding.status is QualityStatus.FAIL
    assert finding.count == 1
    assert any("1234567890" in sample for sample in finding.samples)


def test_duplicate_record_numbers_fail_the_check(settings, make_workbook):
    path = make_workbook(
        "record_duplicates.xlsx",
        {
            "داده": [
                ["شماره پرونده", "مبلغ"],
                [9001, 1000],
                [9001, 2000],
                [9002, 3000],
                [9003, 4000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    findings = _findings(analysis, "duplicate_identifiers")
    assert findings, "شماره پرونده تکراری ثبت نشد"
    finding = findings[0]
    assert finding.status is QualityStatus.FAIL
    assert finding.count == 1
    assert any("9001" in sample for sample in finding.samples)


# --------------------------------------------------------------------------- #
# Quality runs on every analyzable block, not only the primary table
# --------------------------------------------------------------------------- #
def test_quality_runs_on_every_analyzable_block(settings, make_workbook):
    path = make_workbook(
        "two_blocks_quality.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ"],
                ["الف", 1000],
                ["ب", 2000],
                ["ج", 3000],
                [],
                ["شرح", "تعداد"],
                ["الف", 5],
                ["ب", 6],
                ["ج", 7],
            ]
        },
    )
    analysis = _analyze(settings, path)
    blocks = {finding.block for finding in analysis.quality.findings if finding.block}
    assert len(blocks) >= 2, blocks
    for block in blocks:
        codes = {
            finding.code
            for finding in analysis.quality.findings
            if finding.block == block
        }
        assert REQUIRED_CHECK_CODES <= codes, (block, codes)


# --------------------------------------------------------------------------- #
# Contract: every required check has an explicit status
# --------------------------------------------------------------------------- #
def test_every_required_check_has_an_explicit_status(settings, make_workbook):
    path = make_workbook(
        "clean.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ"],
                ["الف", 1000],
                ["ب", 2000],
                ["ج", 3000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    assert REQUIRED_CHECK_CODES <= {item.code for item in analysis.quality.findings}
    for code in REQUIRED_CHECK_CODES:
        assert analysis.quality.check_status(code) in set(QualityStatus)
    assert analysis.quality.status in set(QualityStatus)


def test_required_semantic_roles_are_supported():
    required = {
        "identifier",
        "temporal_key",
        "measure",
        "category",
        "group",
        "subtotal",
        "grand_total",
        "descriptive",
    }
    assert required <= {role.value for role in SemanticRole}
    assert required <= {role.value for role in SEMANTIC_ROLE_FA}


def test_category_group_and_descriptive_roles_are_produced(settings, make_workbook):
    rows = [["شرح", "گروه", "توضیحات", "مبلغ"]]
    for index in range(30):
        rows.append(
            [
                ["الف", "ب", "ج", "د"][index % 4],
                ["گروه ۱", "گروه ۲"][index % 2],
                f"توضیح کامل و یکتای ردیف شماره {index} درباره پرونده اداری",
                (index + 1) * 1000,
            ]
        )
    path = make_workbook("roles.xlsx", {"داده": rows})
    analysis = _analyze(settings, path)
    fields = _fields(analysis, "داده")
    assert fields["شرح"].role is SemanticRole.CATEGORY
    assert fields["گروه"].role is SemanticRole.GROUP
    assert fields["توضیحات"].role is SemanticRole.DESCRIPTIVE
    assert fields["مبلغ"].role is SemanticRole.MEASURE


def test_subtotal_and_grand_total_rows_get_semantic_roles(settings, make_workbook):
    path = make_workbook(
        "totals.xlsx",
        {
            "داده": [
                ["شرح", "مبلغ"],
                ["الف", 10],
                ["ب", 20],
                ["جمع جزء", 30],
                ["ج", 5],
                ["جمع کل", 65],
            ]
        },
    )
    analysis = _analyze(settings, path)
    roles = {row.role for block in analysis.semantics.blocks for row in block.rows}
    assert SemanticRole.SUBTOTAL in roles
    assert SemanticRole.GRAND_TOTAL in roles


def test_semantics_and_quality_are_serialized(settings, make_workbook):
    path = make_workbook(
        "serialize.xlsx",
        {
            "داده": [
                ["کد پستی", "مبلغ"],
                [1234567890, 1000],
                [1234567891, 2000],
                [1234567892, 3000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    payload = analysis.to_dict()
    assert payload["semantics"]["blocks"]
    assert payload["quality"]["findings"]
    assert payload["semantics"]["role_counts"].get("identifier", 0) >= 1
