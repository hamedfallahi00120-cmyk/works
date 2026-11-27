"""Exact Decimal comparison policy for model-quoted figures.

The V5 validator accepted a 0.1% relative tolerance. The merged engine uses
zero tolerance for financial figures; these tests pin that behaviour.
"""

from __future__ import annotations

import json
from decimal import Decimal

from local_office_ai.ai.response_validator import (
    DEFAULT_NUMBER_POLICY,
    validate_response,
)


def _payload(summary: str) -> str:
    return json.dumps(
        {"summary": summary, "findings": [], "missing_data": [], "recommendations": [], "risk_level": "low"},
        ensure_ascii=False,
    )


def test_default_policy_has_zero_tolerance():
    assert DEFAULT_NUMBER_POLICY.tolerance == Decimal("0")
    assert "اختلاف مجاز صفر" in DEFAULT_NUMBER_POLICY.describe_fa()


def test_exact_match_is_grounded():
    response = validate_response(_payload("جمع اصل وام ۳۲۵,۰۰۰,۰۰۰ تومان است."), allowed_numbers={"325000000"}, mismatch_exists=False)
    assert response.is_grounded
    assert response.unsupported_numbers == []


def test_one_unit_difference_is_not_accepted():
    response = validate_response(_payload("جمع اصل وام ۳۲۵,۰۰۰,۰۰۱ تومان است."), allowed_numbers={"325000000"}, mismatch_exists=False)
    assert not response.is_grounded
    assert response.unsupported_numbers


def test_relative_tolerance_is_not_used():
    # 0.0001% above the allowed value would pass a 0.1% tolerance; it must fail.
    response = validate_response(_payload("مبلغ ۳۲۵,۰۰۰,۳۲۵ تومان است."), allowed_numbers={"325000000"}, mismatch_exists=False)
    assert not response.is_grounded


def test_scale_words_are_resolved_before_comparison():
    response = validate_response(_payload("حدود ۲ میلیون تومان تعهد وجود دارد."), allowed_numbers={"2000000"}, mismatch_exists=False)
    assert response.is_grounded
    assert response.unsupported_numbers == []


def test_decimal_formatting_is_normalized_but_value_is_exact():
    accepted = validate_response(
        _payload("مبلغ ۱۱,۰۰۰,۰۰۰.۰۰ تومان است."),
        allowed_numbers={"11000000"},
        mismatch_exists=False,
    )
    assert accepted.is_grounded
    rejected = validate_response(
        _payload("مبلغ ۱۱,۰۰۰,۰۰۰.۰۱ تومان است."),
        allowed_numbers={"11000000"},
        mismatch_exists=False,
    )
    assert not rejected.is_grounded


def test_scale_word_with_wrong_magnitude_is_rejected():
    response = validate_response(_payload("حدود ۲ میلیون تومان تعهد وجود دارد."), allowed_numbers={"2"}, mismatch_exists=False)
    assert not response.is_grounded
