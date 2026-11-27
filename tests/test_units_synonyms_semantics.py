"""Focused tests for the capabilities ported from V5.

Covers Decimal scale units, config-driven entity synonyms (matching only) and
header/content conflict reporting in semantic role assignment.
"""

from __future__ import annotations

from decimal import Decimal

from local_office_ai.core.models import ColumnRole, TableKind
from local_office_ai.core.settings import load_settings
from local_office_ai.excel.normalization import (
    canonical_number,
    extract_numeric_claims,
    institution_similarity,
    normalize_institution,
    parse_number,
)
from local_office_ai.excel.reader import WorkbookReader
from local_office_ai.excel.table_detector import TableDetector
from local_office_ai.excel.workbook_inspector import inspect_workbook


def test_scale_units_multiply_and_keep_decimal():
    cases = {
        "حدود ۲ میلیون تومان": Decimal("2000000"),
        "2.5 میلیون": Decimal("2500000"),
        "۱۵ هزار تومان": Decimal("15000"),
        "۳ میلیارد ریال": Decimal("3000000000"),
        "۲ همت": Decimal("2000000000000"),
        "1.25 thousand": Decimal("1250"),
    }
    for raw, expected in cases.items():
        parsed = parse_number(raw)
        assert isinstance(parsed.value, Decimal)
        assert parsed.value == expected, raw
        assert parsed.warnings, f"scale conversion must be documented: {raw}"


def test_scale_word_does_not_change_plain_amounts():
    assert parse_number("325,000,000").value == Decimal("325000000")
    assert canonical_number(parse_number("۲ میلیون").value) == "2000000"


def test_compound_scale_and_negative_words():
    cases = {
        "۱۲ هزار و ۵۰۰": Decimal("12500"),
        "۲ میلیون و ۵۰۰ هزار": Decimal("2500000"),
        "۱ میلیارد و ۲۰۰ میلیون و ۳۰۰ هزار": Decimal("1200300000"),
        "منفی ۲ میلیون": Decimal("-2000000"),
        "بدهکار ۳ میلیارد ریال": Decimal("-3000000000"),
    }
    for raw, expected in cases.items():
        parsed = parse_number(raw)
        assert isinstance(parsed.value, Decimal), raw
        assert parsed.value == expected, raw


def test_numeric_claims_capture_scale_word():
    claims = extract_numeric_claims("حدود ۲ میلیون تومان و ۱.۵ میلیارد ریال و ۳۲۵,۰۰۰,۰۰۰")
    assert "2 میلیون" in claims
    assert "1.5 میلیارد" in claims
    assert "325,000,000" in claims


def test_synonyms_are_matching_only(tmp_path):
    config = tmp_path / "synonyms.json"
    config.write_text('{"صندق": "صندوق", "_comment": "ignored"}', encoding="utf-8")
    settings = load_settings(project_root=tmp_path)
    # load_settings reads config/synonyms.json; write it there for this root.
    (tmp_path / "config").mkdir(exist_ok=True)
    (tmp_path / "config" / "synonyms.json").write_text(config.read_text(encoding="utf-8"), encoding="utf-8")
    settings = load_settings(project_root=tmp_path)
    assert settings.synonyms.get("صندق") == "صندوق"
    assert normalize_institution("صندق امین", settings.synonyms) == normalize_institution(
        "صندوق امین", settings.synonyms
    )
    assert institution_similarity("صندق امین", "صندوق امین", settings.synonyms) == 1.0
    # The raw source text is not rewritten by normalization.
    assert "صندق" in "صندق امین"


def test_synonym_matching_reconciles_spelling_variant(make_workbook, settings):
    path = make_workbook(
        "variant.xlsx",
        {
            "S": [
                ["ردیف", "نام مؤسسه", "مبلغ وام", "مبلغ هر قسط", "تعداد کل قسط"],
                [1, "صندق امین", 1000000, 100000, 12],
                [2, "صندوق رفاه", 2000000, 200000, 10],
                [],
                ["صندق امین", 100000],
            ]
        },
    )
    bundle = WorkbookReader(settings).load(path)
    try:
        meta = inspect_workbook(bundle, settings)
        sheet_meta = next(item for item in meta.sheets if item.title == "S")
        from local_office_ai.excel.sheet_analyzer import SheetGrid

        grid = SheetGrid("S", bundle.values["S"], bundle.formulas["S"], sheet_meta.max_row, sheet_meta.max_col)
        structure = TableDetector(settings).detect(grid)
        loan_table = next(table for table in structure.tables if table.kind is TableKind.LOAN)
        from local_office_ai.excel.data_extractor import DataExtractor

        extractor = DataExtractor(settings)
        loans = extractor.extract_loans(grid, loan_table)
        assert loans
        assert loans[0].institution_raw == "صندق امین"
        assert loans[0].institution_key == "امین"
    finally:
        bundle.close()


def test_header_content_conflict_is_reported(make_workbook, settings):
    path = make_workbook(
        "conflict.xlsx",
        {
            "S": [
                ["ردیف", "نام مؤسسه", "مبلغ هر قسط", "تعداد کل قسط"],
                [1, "بانک الف", "زیاد", 12],
                [2, "بانک ب", "کم", 10],
                [3, "بانک ج", "متوسط", 8],
            ]
        },
    )
    bundle = WorkbookReader(settings).load(path)
    try:
        meta = inspect_workbook(bundle, settings)
        sheet_meta = next(item for item in meta.sheets if item.title == "S")
        from local_office_ai.excel.sheet_analyzer import SheetGrid

        grid = SheetGrid("S", bundle.values["S"], bundle.formulas["S"], sheet_meta.max_row, sheet_meta.max_col)
        structure = TableDetector(settings).detect(grid)
        table = structure.tables[0]
        installment_col = next(column for column in table.columns if column.header == "مبلغ هر قسط")
        assert installment_col.role is not ColumnRole.INSTALLMENT
        assert any("نقش نادیده گرفته شد" in reason for reason in table.reasons)
        institution_col = next(column for column in table.columns if column.header == "نام مؤسسه")
        assert institution_col.role is ColumnRole.INSTITUTION
    finally:
        bundle.close()
