"""Validation of model output against the verified facts.

The model may only describe the data; it may not invent figures.  Every number
in its answer is checked against the set of numbers Python computed.  A number
that cannot be traced is reported as ``unsupported`` and the answer is marked
as not fully grounded, but it is still shown with that warning instead of being
silently trusted.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable

from local_office_ai.core.logging_config import get_logger
from local_office_ai.excel.normalization import (
    canonical_number,
    extract_numeric_claims,
    parse_number,
)

LOGGER = get_logger(__name__)

MISMATCH_WORDS = ("مغایرت", "مغایر", "اختلاف", "discrepancy", "mismatch")
REQUIRED_KEYS = ("summary", "findings", "missing_data", "recommendations")


@dataclass(frozen=True)
class NumberComparisonPolicy:
    """How a number quoted by the model is compared with verified figures.

    Financial amounts are compared on their standardized ``Decimal`` value with
    zero tolerance: 325,000,000 and 325,000,001 are different figures.  Currency
    words and scale words are normalized first, but never rounded.
    """

    tolerance: Decimal = Decimal("0")
    rounding: str = "none"
    currency_agnostic: bool = True
    scale_units: bool = True

    def describe_fa(self) -> str:
        return (
            "مقایسه Decimal استاندارد؛ اختلاف مجاز صفر؛ بدون گردکردن؛ "
            "واحد پول و واحد مقیاس (هزار/میلیون/میلیارد/همت) پیش از مقایسه نرمال می‌شوند."
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "tolerance": str(self.tolerance),
            "rounding": self.rounding,
            "currency_agnostic": self.currency_agnostic,
            "scale_units": self.scale_units,
            "description": self.describe_fa(),
        }


DEFAULT_NUMBER_POLICY = NumberComparisonPolicy()


@dataclass
class ValidatedResponse:
    summary: str = ""
    findings: list[str] = field(default_factory=list)
    missing_data: list[str] = field(default_factory=list)
    recommendations: list[str] = field(default_factory=list)
    risk_level: str = "نامشخص"
    unsupported_numbers: list[str] = field(default_factory=list)
    contradictions: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    is_grounded: bool = True
    parsed_as_json: bool = True
    policy: NumberComparisonPolicy = DEFAULT_NUMBER_POLICY

    def to_dict(self) -> dict[str, Any]:
        return {
            "summary": self.summary,
            "findings": list(self.findings),
            "missing_data": list(self.missing_data),
            "recommendations": list(self.recommendations),
            "risk_level": self.risk_level,
            "unsupported_numbers": list(self.unsupported_numbers),
            "contradictions": list(self.contradictions),
            "warnings": list(self.warnings),
            "is_grounded": self.is_grounded,
            "parsed_as_json": self.parsed_as_json,
            "policy": self.policy.to_dict(),
        }

    @property
    def all_text(self) -> str:
        return " ".join([self.summary, *self.findings, *self.missing_data, *self.recommendations])


def _strip_code_fence(content: str) -> str:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    return text.strip()


def _as_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value.strip()] if value.strip() else []
    if isinstance(value, (list, tuple)):
        result = []
        for item in value:
            if isinstance(item, str):
                if item.strip():
                    result.append(item.strip())
            elif isinstance(item, (int, float, Decimal)):
                result.append(str(item))
            elif isinstance(item, dict):
                result.append("؛ ".join(f"{key}: {val}" for key, val in item.items()))
        return result
    return [str(value)]


def validate_response(
    content: str,
    *,
    allowed_numbers: Iterable[str],
    mismatch_exists: bool,
    extra_allowed: Iterable[str] = (),
    policy: NumberComparisonPolicy = DEFAULT_NUMBER_POLICY,
) -> ValidatedResponse:
    response = ValidatedResponse(policy=policy)
    text = _strip_code_fence(content)
    if not text:
        response.summary = ""
        response.warnings.append("پاسخ مدل خالی بود.")
        response.is_grounded = False
        return response

    try:
        payload = json.loads(text)
        if not isinstance(payload, dict):
            raise ValueError("top-level JSON is not an object")
    except (json.JSONDecodeError, ValueError):
        response.parsed_as_json = False
        response.warnings.append("پاسخ مدل JSON معتبر نبود؛ متن خام با هشدار نمایش داده می‌شود.")
        response.summary = text[:3500]
        response.is_grounded = False
        payload = {}

    if payload:
        response.summary = str(payload.get("summary") or "").strip()
        response.findings = _as_list(payload.get("findings"))
        response.missing_data = _as_list(payload.get("missing_data"))
        response.recommendations = _as_list(payload.get("recommendations"))
        risk = str(payload.get("risk_level") or "").strip().lower()
        response.risk_level = risk if risk in {"low", "medium", "high"} else "نامشخص"
        missing = [key for key in REQUIRED_KEYS if key not in payload]
        if missing:
            response.warnings.append("کلیدهای غایب در پاسخ مدل: " + "، ".join(missing))
        if not response.summary:
            response.warnings.append("خلاصه مدل خالی بود.")

    allowed = {canonical_number(value) for value in allowed_numbers if str(value).strip()}
    allowed.update(canonical_number(value) for value in extra_allowed if str(value).strip())
    unsupported: list[str] = []
    for token in extract_numeric_claims(response.all_text):
        canonical = _canonical_token(token)
        if canonical is None:
            continue
        if _number_is_allowed(canonical, allowed, policy) or _is_benign_number(canonical):
            continue
        if token not in unsupported:
            unsupported.append(token)
    response.unsupported_numbers = unsupported
    if unsupported:
        response.is_grounded = False
        response.warnings.append(
            "اعداد بدون منبع در پاسخ مدل: " + "، ".join(unsupported[:10])
        )

    if not mismatch_exists:
        lowered = response.all_text
        if any(word in lowered for word in MISMATCH_WORDS):
            response.contradictions.append(
                "مدل از مغایرت/اختلاف سخن گفته اما حسابرسی Python مغایرتی ثبت نکرده است."
            )
            response.is_grounded = False

    if response.parsed_as_json and not response.findings and not response.recommendations:
        response.warnings.append("پاسخ مدل فقط خلاصه دارد و یافته/پیشنهاد جداگانه ندارد.")

    LOGGER.info(
        "ai response validated: grounded=%s unsupported=%d warnings=%d",
        response.is_grounded, len(response.unsupported_numbers), len(response.warnings),
    )
    return response


def _number_is_allowed(
    canonical: str,
    allowed: set[str],
    policy: NumberComparisonPolicy,
) -> bool:
    """Exact canonical match, with no default financial tolerance.

    A non-zero ``policy.tolerance`` is supported for explicitly requested
    non-financial comparisons, but the engine's default policy is zero.
    """
    if canonical in allowed:
        return True
    if policy.tolerance <= 0:
        return False
    try:
        value = Decimal(canonical)
    except (InvalidOperation, ValueError):
        return False
    for candidate in allowed:
        try:
            if abs(value - Decimal(candidate)) <= policy.tolerance:
                return True
        except (InvalidOperation, ValueError):
            continue
    return False


def _canonical_token(token: str) -> str | None:
    parsed = parse_number(token, numeric_context=True)
    if isinstance(parsed.value, Decimal):
        try:
            return canonical_number(parsed.value)
        except (InvalidOperation, ValueError):
            return None
    return None


def _is_benign_number(canonical: str) -> bool:
    """Percentages, small counts and years are not financial claims."""
    try:
        value = Decimal(canonical)
    except (InvalidOperation, ValueError):
        return False
    if value == value.to_integral_value():
        integer = int(value)
        if 0 <= integer <= 100:
            return True
        if 1200 <= integer <= 2200:
            return True
    return False
