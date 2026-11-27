from __future__ import annotations

from datetime import date
from decimal import Decimal

import pytest

from local_office_ai.core.models import ValueKind
from local_office_ai.excel.normalization import (
    canonical_number,
    excel_serial_to_date,
    extract_numeric_tokens,
    format_jalali,
    gregorian_to_jalali,
    institution_similarity,
    is_valid_jalali,
    jalali_to_gregorian,
    looks_like_date,
    looks_like_number,
    normalize_digits,
    normalize_institution,
    normalize_text,
    parse_date,
    parse_number,
    parse_value,
)


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("۱۲۳", Decimal("123")),
        ("٤٥٦", Decimal("456")),
        ("1,234,567", Decimal("1234567")),
        ("۱٬۲۳۴٬۵۶۷", Decimal("1234567")),
        ("1 234 567", Decimal("1234567")),
        ("1.234,50", Decimal("1234.50")),
        ("1,234.50", Decimal("1234.50")),
        ("1٫5", Decimal("1.5")),
    ],
)
def test_parse_number_variants(raw, expected):
    parsed = parse_number(raw)
    assert parsed.kind is ValueKind.NUMBER
    assert parsed.value == expected


def test_negative_forms():
    assert parse_number("(1,234)").value == Decimal("-1234")
    assert parse_number("-1,234").value == Decimal("-1234")
    assert parse_number("−500").value == Decimal("-500")


def test_currency_and_embedded_number():
    parsed = parse_number("مبلغ 2,500,000 تومان")
    assert parsed.value == Decimal("2500000")
    assert any("extracted" in warning for warning in parsed.warnings)
    assert parse_number("۱۲۵ ریال").value == Decimal("125")


def test_leading_zero_is_identifier():
    parsed = parse_number("000123")
    assert parsed.kind is ValueKind.TEXT
    assert parsed.value == "000123"
    # Inside a numeric column the same text is an amount.
    assert parse_number("000123", numeric_context=True).value == Decimal("123")


def test_ambiguous_grouping_is_reported():
    parsed = parse_number("1.234")
    assert parsed.value == Decimal("1.234")
    assert any("ambiguous" in warning for warning in parsed.warnings)


def test_number_rejects_dates_and_errors():
    assert parse_number(date(2024, 1, 1)).kind is ValueKind.DATE
    assert parse_number("#DIV/0!").kind is ValueKind.ERROR
    assert parse_number("هیچ عددی ندارد").kind is ValueKind.TEXT


def test_jalali_known_anchors():
    assert jalali_to_gregorian(1399, 1, 1) == (2020, 3, 20)
    assert jalali_to_gregorian(1403, 1, 1) == (2024, 3, 20)
    assert jalali_to_gregorian(1404, 1, 1) == (2025, 3, 21)
    assert gregorian_to_jalali(2024, 3, 20) == (1403, 1, 1)


def test_jalali_roundtrip_for_a_full_year():
    for month in range(1, 13):
        for day in (1, 15, 29):
            if not is_valid_jalali(1404, month, day):
                continue
            gregorian = jalali_to_gregorian(1404, month, day)
            assert gregorian_to_jalali(*gregorian) == (1404, month, day)


def test_parse_jalali_and_gregorian_dates():
    parsed = parse_date("1406/12/01")
    assert parsed.kind is ValueKind.DATE
    assert parsed.value == date(2028, 2, 20)
    assert parsed.method == "jalali-YMD"

    parsed = parse_date("2025-03-21")
    assert parsed.value == date(2025, 3, 21)
    assert parsed.method == "gregorian-YMD"

    assert parse_date("۱۲ فروردین ۱۴۰۴").value == date(2025, 4, 1)
    assert format_jalali(date(2028, 2, 20)) == "1406/12/01"


def test_ambiguous_day_month_warns():
    parsed = parse_date("01/02/2025")
    assert parsed.value == date(2025, 2, 1)
    assert any("ambiguous" in warning for warning in parsed.warnings)
    assert parse_date("01/02/2025", day_first=False).value == date(2025, 1, 2)


def test_invalid_date_rejected():
    parsed = parse_date("1404/13/40")
    assert parsed.kind is ValueKind.TEXT
    assert any("not valid" in warning for warning in parsed.warnings)


def test_excel_serial():
    assert excel_serial_to_date(45000) == date(2023, 3, 15)
    assert parse_date(45000).value == date(2023, 3, 15)
    assert parse_date(45000).method == "excel-serial"


def test_text_normalization():
    assert normalize_digits("۱۲3") == "123"
    assert normalize_text("بانك  مهر‌ايران") == "بانک مهر ایران"
    assert normalize_text("متن\u200bبا\u200eکاراکتر نامرئی") == "متنباکاراکتر نامرئی"
    assert normalize_institution("بانك مهر ايران") == normalize_institution("مهر ایران")


def test_institution_similarity_levels():
    assert institution_similarity("صندق امین", "صندوق امین") == 1.0
    assert institution_similarity("بانک مهر ایران", "مهر ایران") >= 0.9
    assert institution_similarity("بانک مهر ایران", "صندوق رفاه کهنز") < 0.4
    assert institution_similarity("", "x") == 0.0


def test_canonical_numbers_and_tokens():
    assert canonical_number(Decimal("14749000.00")) == "14749000"
    assert canonical_number(Decimal("0.000001")) == "0.000001"
    assert extract_numeric_tokens("جمع ۳۲۵,۰۰۰,۰۰۰ و 12.5") == ["325,000,000", "12.5"]


def test_heuristics():
    assert looks_like_number("۱۲٬۵۰۰")
    assert not looks_like_number("بانک مهر")
    assert looks_like_date("1404/01/15")
    assert looks_like_date("12 فروردین 1404")
    assert not looks_like_date("1404")


def test_parse_value_dispatcher():
    assert parse_value("1404/01/15").kind is ValueKind.DATE
    assert parse_value("۱۲٬۵۰۰").kind is ValueKind.NUMBER
    assert parse_value("بانک").kind is ValueKind.TEXT
    assert parse_value(None).kind is ValueKind.EMPTY
    assert parse_value("   ").kind is ValueKind.EMPTY
    assert parse_value(True).kind is ValueKind.BOOLEAN


def test_formula_text_is_never_parsed_as_number():
    """A cell holding ``=SUM(...)`` without a cached value must stay text."""
    from local_office_ai.core.models import ValueKind
    from local_office_ai.excel.normalization import parse_value

    parsed = parse_value("=SUM(C4:C7)", kind_hint=ValueKind.NUMBER)
    assert parsed.value is None
    assert parsed.kind is ValueKind.TEXT
    assert parsed.warnings
