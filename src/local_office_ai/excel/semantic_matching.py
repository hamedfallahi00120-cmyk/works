"""Controlled semantic matching of equivalent fields across sources.

Two columns are not the same indicator because their names look alike.  This
module decides whether two extracted fields describe the same information and
records the evidence for the decision:

* the normalized label and its canonical tokens (spelling, ZWNJ, plural forms,
  a small cross-language administrative vocabulary and configured synonyms);
* the semantic role (a measure is never matched to an identifier or a category);
* the unit and its scale (``ریال`` and ``نفر`` are different dimensions;
  ``میلیون ریال`` and ``ریال`` differ only by a scale that must be applied);
* the reporting period, compared as real date ranges in the Jalali or Gregorian
  calendar (see :mod:`period_alignment`);
* shared identifier columns and overlapping category values, which support a
  match but never replace the label evidence.

Only a pair whose label is canonical-equal, whose units agree and whose periods
are equivalent is reported as ``matched``.  A pair that is plausible but cannot
be proven (missing period, missing unit, scale difference) is ``probable``; two
or more plausible candidates for one field are ``ambiguous`` and every candidate
is kept.  Similar names that describe different indicators are ``rejected`` with
the reason.  This stage never aggregates values.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from difflib import SequenceMatcher
from enum import Enum
from typing import Any, Iterable

from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import normalize_digits, normalize_text
from local_office_ai.excel.period_alignment import (
    RELATION_FA,
    NormalizedPeriod,
    PeriodCompatibility,
    compare_period_sets,
    compare_periods,
    parse_period,
    parse_periods,
)
from local_office_ai.excel.semantic_classification import SemanticRole
from local_office_ai.excel.semantic_vocabulary import (
    currency_words_in,
    matches_word,
)

KIND_MEASURE = "measure"
KIND_CATEGORY = "category"
KIND_IDENTIFIER = "identifier"
KIND_OTHER = "other"

KIND_FA = {
    KIND_MEASURE: "شاخص اندازه‌گیری",
    KIND_CATEGORY: "دسته/گروه",
    KIND_IDENTIFIER: "شناسه",
    KIND_OTHER: "سایر",
}


class MatchStatus(str, Enum):
    MATCHED = "matched"
    PROBABLE = "probable"
    AMBIGUOUS = "ambiguous"
    REJECTED = "rejected"


STATUS_FA = {
    MatchStatus.MATCHED: "تطبیق قطعی",
    MatchStatus.PROBABLE: "تطبیق احتمالی",
    MatchStatus.AMBIGUOUS: "مبهم",
    MatchStatus.REJECTED: "رد شد",
}

PERIOD_STATUS_FA = {
    "detected": "شناسایی‌شده",
    "not_detected": "شناسایی‌نشده",
    "conflicting": "ناسازگار",
    "partial": "بخشی",
}

UNIT_STATE_FA = {
    "equal": "یکسان",
    "both_missing": "ثبت‌نشده در هر دو",
    "scalable": "مقیاس متفاوت",
    "same_dimension": "هم‌جنس با واحد متفاوت",
    "unknown": "نامشخص",
    "incompatible": "ناسازگار",
    "mixed": "ترکیبی",
}

_ROLE_KIND = {
    SemanticRole.MEASURE: KIND_MEASURE,
    SemanticRole.CALCULATED: KIND_MEASURE,
    SemanticRole.CATEGORY: KIND_CATEGORY,
    SemanticRole.GROUP: KIND_CATEGORY,
    SemanticRole.IDENTIFIER: KIND_IDENTIFIER,
}

# Canonical tokens that only say *how* a quantity is measured.  They may differ
# between two otherwise identical labels (``تعداد پرونده`` / ``پرونده``), but two
# labels that consist only of such words are not automatically the same measure.
_MEASURE_TOKENS = {
    "count", "amount", "total", "average", "percent", "rate", "share",
    "quantity", "score", "weight", "volume", "area", "distance", "usage",
    "consumption", "production", "sales", "purchase", "stock", "debt",
    "credit", "profit", "loss", "payment", "income", "expense", "budget",
}

# A counting word may be present on one side and absent on the other without
# changing the indicator (``تعداد پرونده`` / ``پرونده``).  Aggregation words
# (``کل``, ``جمع``, ``میانگین``, ``درصد``) change the meaning and must agree.
_INTERCHANGEABLE_MEASURE_TOKENS = {"count", "quantity"}
_AGGREGATION_MEASURE_TOKENS = {"total", "average", "percent", "rate", "share"}

# A deliberately small cross-language administrative vocabulary.  It is an aid,
# not a translation engine: only canonical-token equality can make a match, and
# unit/period evidence must still agree.
_DEFAULT_ALIASES = {
    "تعداد": "count",
    "تعدادها": "count",
    "تعدادهای": "count",
    "count": "count",
    "counts": "count",
    "number": "count",
    "numbers": "count",
    "qty": "count",
    "quantity": "count",
    "مبلغ": "amount",
    "مبالغ": "amount",
    "amount": "amount",
    "amounts": "amount",
    "value": "amount",
    "values": "amount",
    "جمع": "total",
    "مجموع": "total",
    "کل": "total",
    "total": "total",
    "totals": "total",
    "sum": "total",
    "میانگین": "average",
    "average": "average",
    "mean": "average",
    "پرونده": "case",
    "پرونده‌ها": "case",
    "پرونده‌های": "case",
    "case": "case",
    "cases": "case",
    "شعبه": "branch",
    "شعبه‌ها": "branch",
    "branches": "branch",
    "branch": "branch",
    "متهم": "defendant",
    "متهمان": "defendant",
    "defendant": "defendant",
    "defendants": "defendant",
    "accused": "defendant",
    "درآمد": "income",
    "درامد": "income",
    "income": "income",
    "revenue": "income",
    "revenues": "income",
    "هزینه": "cost",
    "هزینه‌ها": "cost",
    "cost": "cost",
    "costs": "cost",
    "expense": "cost",
    "expenses": "cost",
    "موجودی": "stock",
    "stock": "stock",
    "balance": "stock",
    "نفر": "person",
    "نفرات": "person",
    "person": "person",
    "people": "person",
    "persons": "person",
    "بازداشت": "arrest",
    "arrest": "arrest",
    "arrests": "arrest",
    "detention": "arrest",
    "درصد": "percent",
    "percent": "percent",
    "percentage": "percent",
    "نرخ": "rate",
    "rate": "rate",
    "سهم": "share",
    "share": "share",
}

_STOPWORDS = {
    "از", "در", "به", "با", "برای", "و", "یا", "که", "این", "آن",
    "the", "of", "for", "in", "and", "or", "to", "a", "an", "on", "at",
}

_GENERIC_LABEL_WORDS = {
    "گزارش", "آمار", "جدول", "فرم", "لیست", "فهرست", "دوره", "ماه", "سال",
    "ماهه", "سالانه", "report", "table", "form", "list", "period", "month",
    "year", "yearly", "monthly",
}

_GENERIC_IDENTIFIER_LABELS = {
    frozenset({"ردیف"}),
    frozenset({"row"}),
    frozenset({"rows"}),
    frozenset({"index"}),
    frozenset({"no"}),
    frozenset({"number"}),
    frozenset({"شماره"}),
}

_PUNCTUATION = re.compile(r"[.,;:!?()\[\]{}«»\"'`ـ_/\\|+*&^%$#@~=<>،؛]")

_PERIOD_PATTERNS = (
    re.compile(r"(?<!\d)\d{4}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{1,2}(?!\d)"),
    re.compile(r"(?<!\d)\d{4}\s*[/\-.]\s*\d{1,2}(?![\d/.\-])"),
    re.compile(r"(?<![\d/.\-])\d{1,2}\s*[/\-.]\s*\d{4}(?!\d)"),
    re.compile(r"(?<!\d)(?:1[1-6]\d{2}|1[7-9]\d{2}|20\d{2}|21\d{2})(?!\d)"),
)

_JALALI_MONTH_WORDS = (
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
)
_GREGORIAN_MONTH_WORDS = (
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
    "jan", "feb", "mar", "apr", "jun", "jul", "aug", "sep", "sept",
    "oct", "nov", "dec",
)

# --------------------------------------------------------------------------- #
# Units
# --------------------------------------------------------------------------- #
_UNIT_PATTERNS: tuple[tuple[str, str, str], ...] = (
    ("متر مربع", "area", "SQM"),
    ("مترمربع", "area", "SQM"),
    ("کیلومتر", "length", "KM"),
    ("کیلوگرم", "mass", "KG"),
    ("هکتار", "area", "HECTARE"),
    ("لیتر", "volume", "LITER"),
    ("ریال", "currency", "IRR"),
    ("ریالی", "currency", "IRR"),
    ("تومان", "currency", "TOMAN"),
    ("تومانی", "currency", "TOMAN"),
    ("تومن", "currency", "TOMAN"),
    ("همت", "currency", "TOMAN"),
    ("درهم", "currency", "AED"),
    ("دلار", "currency", "USD"),
    ("یورو", "currency", "EUR"),
    ("پوند", "currency", "GBP"),
    ("درصد", "ratio", "PERCENT"),
    ("درصدی", "ratio", "PERCENT"),
    ("نفر", "count", "PERSON"),
    ("نفری", "count", "PERSON"),
    ("عدد", "count", "UNIT"),
    ("عددی", "count", "UNIT"),
    ("متر", "length", "METER"),
    ("گرم", "mass", "GRAM"),
    ("ساعت", "time", "HOUR"),
    ("دقیقه", "time", "MINUTE"),
    ("روز", "time", "DAY"),
    ("ماه", "time", "MONTH"),
    ("سال", "time", "YEAR"),
    ("percent", "ratio", "PERCENT"),
    ("percentage", "ratio", "PERCENT"),
    ("rial", "currency", "IRR"),
    ("irr", "currency", "IRR"),
    ("toman", "currency", "TOMAN"),
    ("usd", "currency", "USD"),
    ("eur", "currency", "EUR"),
    ("person", "count", "PERSON"),
    ("people", "count", "PERSON"),
    ("kg", "mass", "KG"),
    ("km", "length", "KM"),
    ("hour", "time", "HOUR"),
    ("day", "time", "DAY"),
    ("month", "time", "MONTH"),
    ("year", "time", "YEAR"),
    ("unit", "count", "UNIT"),
)

_SCALE_FACTORS: tuple[tuple[str, Decimal], ...] = (
    ("تریلیون", Decimal("1000000000000")),
    ("trillion", Decimal("1000000000000")),
    ("همت", Decimal("1000000000000")),
    ("میلیارد", Decimal("1000000000")),
    ("billion", Decimal("1000000000")),
    ("میلیون", Decimal("1000000")),
    ("million", Decimal("1000000")),
    ("هزار", Decimal("1000")),
    ("thousand", Decimal("1000")),
)

# Unit and scale words are removed from the label because they are compared in
# their own dimension.  A lone unit word (e.g. ``درصد``) is kept as the label.
_UNIT_TOKENS: set[str] = {
    word
    for phrase, _dimension, _code in _UNIT_PATTERNS
    for word in phrase.split()
}
_UNIT_TOKENS.update(word for word, _factor in _SCALE_FACTORS)
_UNIT_TOKENS.update(
    {"ریالی", "تومانی", "نفری", "عددی", "درصدی", "میلیونی", "میلیاردی", "هزاری"}
)


@dataclass(frozen=True)
class UnitInfo:
    dimension: str = ""
    canonical: str = ""
    scale: Decimal = Decimal(1)
    text: str = ""

    @property
    def known(self) -> bool:
        return bool(self.dimension)

    def label(self) -> str:
        if not self.known:
            return self.text or "—"
        if self.scale != Decimal(1):
            return f"{self.canonical}×{self.scale}"
        return self.canonical

    def to_dict(self) -> dict[str, Any]:
        return {
            "dimension": self.dimension,
            "canonical": self.canonical,
            "scale": str(self.scale),
            "text": self.text,
        }


def canonical_unit(*texts: Any) -> UnitInfo:
    """Recognize one unit and its scale from a header or a ``واحد`` column."""
    raw = " ".join(
        normalize_text(normalize_digits(str(text or ""))).lower() for text in texts if text
    ).strip()
    if not raw:
        return UnitInfo()
    dimension = canonical = ""
    for phrase, dim, code in _UNIT_PATTERNS:
        if matches_word(raw, phrase):
            dimension, canonical = dim, code
            break
    if not dimension:
        currencies = currency_words_in(raw)
        if currencies:
            dimension = "currency"
            canonical = sorted(currencies)[0].upper()
    scale = Decimal(1)
    for word, factor in _SCALE_FACTORS:
        if matches_word(raw, word):
            scale = factor
            break
    if not dimension:
        # Scale words alone are not a unit; keep them out of the comparison.
        return UnitInfo()
    return UnitInfo(dimension=dimension, canonical=canonical, scale=scale, text=raw)


@dataclass
class UnitComparison:
    state: str
    note: str
    left: UnitInfo
    right: UnitInfo

    @property
    def compatible(self) -> bool:
        return self.state in {"equal", "both_missing", "scalable", "same_dimension"}

    def score(self) -> float:
        return {
            "equal": 1.0,
            "both_missing": 0.75,
            "scalable": 0.6,
            "same_dimension": 0.5,
            "unknown": 0.4,
            "incompatible": 0.0,
        }.get(self.state, 0.0)

    def to_dict(self) -> dict[str, Any]:
        return {
            "state": self.state,
            "compatible": self.compatible,
            "note": self.note,
            "left": self.left.to_dict(),
            "right": self.right.to_dict(),
        }


def compare_units(left: UnitInfo, right: UnitInfo) -> UnitComparison:
    if not left.known and not right.known:
        return UnitComparison("both_missing", "واحد برای هیچ‌کدام ثبت نشده است.", left, right)
    if not left.known or not right.known:
        return UnitComparison(
            "unknown",
            "واحد فقط برای یکی از دو فیلد مشخص است؛ برابری واحد اثبات نمی‌شود.",
            left,
            right,
        )
    if left.dimension != right.dimension:
        return UnitComparison(
            "incompatible",
            f"واحد «{left.label()}» ({left.dimension}) با «{right.label()}» ({right.dimension}) "
            "از یک جنس نیست؛ مقایسه مجاز نیست.",
            left,
            right,
        )
    if left.canonical != right.canonical:
        return UnitComparison(
            "same_dimension",
            f"هر دو از جنس {left.dimension} هستند اما واحد «{left.label()}» و "
            f"«{right.label()}» یکسان نیستند.",
            left,
            right,
        )
    if left.scale != right.scale:
        factor = left.scale / right.scale
        return UnitComparison(
            "scalable",
            f"واحد یکسان است اما مقیاس متفاوت است؛ برای مقایسه باید مقدار «{left.label()}» "
            f"در ضریب {factor} نسبت به «{right.label()}» تبدیل شود.",
            left,
            right,
        )
    return UnitComparison("equal", f"واحد هر دو «{left.label()}» است.", left, right)


# --------------------------------------------------------------------------- #
# Label normalization
# --------------------------------------------------------------------------- #
def _strip_plural(word: str) -> str:
    if len(word) > 4 and word.endswith("های"):
        return word[:-3]
    if len(word) > 3 and word.endswith("ها"):
        return word[:-2]
    return word


def _strip_periods(text: str) -> str:
    cleaned = text
    for pattern in _PERIOD_PATTERNS:
        cleaned = pattern.sub(" ", cleaned)
    for month in _JALALI_MONTH_WORDS:
        cleaned = re.sub(rf"(?<![\wآ-ی]){month}(?:ماه)?(?![\wآ-ی])", " ", cleaned)
    for month in _GREGORIAN_MONTH_WORDS:
        cleaned = re.sub(rf"(?<![a-z]){month}(?![a-z])", " ", cleaned, flags=re.IGNORECASE)
    return cleaned


def normalize_label(
    header: Any,
    *,
    synonyms: dict[str, str] | None = None,
    aliases: dict[str, str] | None = None,
) -> tuple[str, list[str]]:
    """Return a readable normalized label and its canonical token list."""
    text = normalize_text(normalize_digits(str(header or ""))).lower()
    for wrong, correct in (synonyms or {}).items():
        if wrong:
            text = text.replace(str(wrong).lower(), str(correct).lower())
    text = _strip_periods(text)
    text = _PUNCTUATION.sub(" ", text)
    text = re.sub(r"[\s\-]+", " ", text).strip()

    mapping = dict(_DEFAULT_ALIASES)
    mapping.update({str(key).lower(): str(value).lower() for key, value in (aliases or {}).items()})

    tokens: list[str] = []
    for word in text.split():
        if word in _STOPWORDS or word in _GENERIC_LABEL_WORDS:
            continue
        word = _strip_plural(word)
        if not word:
            continue
        word = mapping.get(word, word)
        if word and word not in tokens:
            tokens.append(word)
    if len(tokens) > 1:
        without_units = [token for token in tokens if token not in _UNIT_TOKENS]
        if without_units:
            tokens = without_units
    readable = " ".join(tokens)
    return readable, tokens


# --------------------------------------------------------------------------- #
# Source fields
# --------------------------------------------------------------------------- #
@dataclass
class SourceField:
    source: str
    sheet: str
    block: str
    header: str
    normalized_label: str
    tokens: list[str]
    role: str
    kind: str
    unit: UnitInfo
    periods: list[NormalizedPeriod]
    table_title: str = ""
    letter: str = ""
    samples: list[str] = field(default_factory=list)
    block_identifiers: list[str] = field(default_factory=list)
    identifier_values: dict[str, list[str]] = field(default_factory=dict)
    row_labels: list[str] = field(default_factory=list)
    evidence: list[str] = field(default_factory=list)
    period_conflict: bool = False

    @property
    def identity(self) -> str:
        return f"{self.source}::{self.sheet}::{self.block}::{self.letter}"

    @property
    def kind_label(self) -> str:
        return KIND_FA.get(self.kind, self.kind)

    @property
    def content_tokens(self) -> set[str]:
        return {token for token in self.tokens if token not in _MEASURE_TOKENS}

    @property
    def measure_tokens(self) -> set[str]:
        return {token for token in self.tokens if token in _MEASURE_TOKENS}

    @property
    def period_labels(self) -> list[str]:
        return [period.label for period in self.periods]

    @property
    def display_label(self) -> str:
        periods = "، ".join(self.period_labels) or "—"
        unit = self.unit.label() if self.unit.known else "—"
        return (
            f"{self.source}/{self.sheet}/{self.block}/«{self.header}» "
            f"[دوره: {periods}; واحد: {unit}]"
        )

    @property
    def period_status(self) -> str:
        if self.period_conflict and self.periods:
            return "conflicting"
        return "detected" if self.periods else "not_detected"

    def to_dict(self) -> dict[str, Any]:
        return {
            "identity": self.identity,
            "source": self.source,
            "sheet": self.sheet,
            "block": self.block,
            "letter": self.letter,
            "header": self.header,
            "normalized_label": self.normalized_label,
            "tokens": list(self.tokens),
            "role": self.role,
            "kind": self.kind,
            "kind_label": self.kind_label,
            "unit": self.unit.to_dict(),
            "periods": [period.to_dict() for period in self.periods],
            "period_status": self.period_status,
            "period_status_label": PERIOD_STATUS_FA.get(self.period_status, self.period_status),
            "table_title": self.table_title,
            "samples": list(self.samples[:8]),
            "block_identifiers": list(self.block_identifiers),
            "identifier_values": {
                header: list(values[:8])
                for header, values in self.identifier_values.items()
            },
            "row_labels": list(self.row_labels[:20]),
            "evidence": list(self.evidence),
            "period_conflict": self.period_conflict,
        }


@dataclass
class MatchCandidate:
    field: SourceField
    status: MatchStatus
    score: float
    evidence: list[str] = field(default_factory=list)
    reason: str = ""
    comparable: bool = False

    @property
    def period_status(self) -> str:
        return self.field.period_status

    def to_dict(self) -> dict[str, Any]:
        return {
            "field": self.field.to_dict(),
            "status": self.status.value,
            "status_label": STATUS_FA.get(self.status, self.status.value),
            "score": round(self.score, 4),
            "period_status": self.period_status,
            "period_status_label": PERIOD_STATUS_FA.get(
                self.period_status, self.period_status
            ),
            "evidence": list(self.evidence),
            "reason": self.reason,
            "comparable": self.comparable,
        }


@dataclass
class FieldMatch:
    match_id: str
    status: MatchStatus
    kind: str
    normalized_label: str
    confidence: float
    members: list[SourceField]
    evidence: list[str] = field(default_factory=list)
    period_relation: str = ""
    period_explanation: str = ""
    period_status: str = ""
    unit_state: str = ""
    candidates: list[MatchCandidate] = field(default_factory=list)
    reason: str = ""
    comparable: bool = False

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    @property
    def kind_label(self) -> str:
        return KIND_FA.get(self.kind, self.kind)

    @property
    def member_labels(self) -> list[str]:
        return [item.display_label for item in self.members]

    def to_dict(self) -> dict[str, Any]:
        return {
            "match_id": self.match_id,
            "status": self.status.value,
            "status_label": self.status_label,
            "kind": self.kind,
            "kind_label": self.kind_label,
            "normalized_label": self.normalized_label,
            "confidence": round(self.confidence, 4),
            "members": [item.to_dict() for item in self.members],
            "evidence": list(self.evidence),
            "period_relation": self.period_relation,
            "period_explanation": self.period_explanation,
            "period_status": self.period_status,
            "period_status_label": PERIOD_STATUS_FA.get(self.period_status, self.period_status),
            "unit_state": self.unit_state,
            "candidates": [item.to_dict() for item in self.candidates],
            "reason": self.reason,
            "comparable": self.comparable,
        }


@dataclass
class RejectedPair:
    left: SourceField
    right: SourceField
    reason: str
    label_score: float
    period_relation: str
    unit_state: str
    period_status: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "left": self.left.to_dict(),
            "right": self.right.to_dict(),
            "reason": self.reason,
            "label_score": round(self.label_score, 4),
            "period_relation": self.period_relation,
            "unit_state": self.unit_state,
            "period_status": self.period_status,
            "period_status_label": PERIOD_STATUS_FA.get(self.period_status, self.period_status),
        }


@dataclass
class SemanticMatchingReport:
    fields: list[SourceField] = field(default_factory=list)
    matches: list[FieldMatch] = field(default_factory=list)
    probable: list[FieldMatch] = field(default_factory=list)
    ambiguities: list[FieldMatch] = field(default_factory=list)
    rejections: list[RejectedPair] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def matched_fields(self) -> list[SourceField]:
        return [item for match in self.matches for item in match.members]

    def counts(self) -> dict[str, int]:
        return {
            "fields": len(self.fields),
            "matched": len(self.matches),
            "probable": len(self.probable),
            "ambiguous": len(self.ambiguities),
            "rejected": len(self.rejections),
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "counts": self.counts(),
            "fields": [item.to_dict() for item in self.fields],
            "matches": [item.to_dict() for item in self.matches],
            "probable": [item.to_dict() for item in self.probable],
            "ambiguities": [item.to_dict() for item in self.ambiguities],
            "rejections": [item.to_dict() for item in self.rejections],
            "warnings": list(self.warnings),
        }


# --------------------------------------------------------------------------- #
# Collection from the semantic model
# --------------------------------------------------------------------------- #
def _period_rank(period: NormalizedPeriod) -> int:
    return {"day": 3, "month": 2, "range": 2, "year": 1}.get(period.granularity, 0)


def _finest_periods(periods: list[NormalizedPeriod]) -> list[NormalizedPeriod]:
    """Keep the most specific periods so a bare year never hides a month."""
    if not periods:
        return periods
    best = max(_period_rank(period) for period in periods)
    return [period for period in periods if _period_rank(period) == best]


def _periods_conflict(periods: list[NormalizedPeriod]) -> bool:
    """True when one field carries two disjoint reporting periods."""
    for index, left in enumerate(periods):
        for right in periods[index + 1 :]:
            if compare_periods(left, right).relation == "incompatible":
                return True
    return False


def _title_tokens(text: str) -> set[str]:
    cleaned = _PUNCTUATION.sub(
        " ", normalize_text(normalize_digits(str(text or ""))).lower()
    )
    return {
        word
        for word in cleaned.split()
        if word
        and word not in _STOPWORDS
        and word not in _GENERIC_LABEL_WORDS
        and not any(char.isdigit() for char in word)
    }


def _select_sheet_periods(labels: list[str], title: str) -> list[str]:
    """Use the sheet period that names the same month as this block title."""
    if len(labels) <= 1 or not title:
        return labels
    wanted = _title_tokens(title)
    if not wanted:
        return labels
    selected = [label for label in labels if _title_tokens(label) & wanted]
    return selected or labels


def _has_month_word(text: str) -> bool:
    from local_office_ai.excel.period_alignment import GREGORIAN_MONTHS, JALALI_MONTHS

    return any(
        matches_word(text, month) for month in (*JALALI_MONTHS, *GREGORIAN_MONTHS)
    )


def _sheet_year(block: Any, sheet_periods: dict[str, list[str]]) -> int | None:
    sheet = getattr(block, "sheet", "")
    for source in [sheet, *(sheet_periods.get(sheet, []) or [])]:
        period = parse_period(source)
        if period is not None and period.detected and period.year:
            return period.year
    return None


def _title_period(
    block: Any,
    sheet_periods: dict[str, list[str]],
) -> NormalizedPeriod | None:
    """Use the block title itself, joining a month with the sheet year."""
    title = str(getattr(block, "title", "") or "")
    period = parse_period(title)
    if period is not None and period.detected:
        return period
    if _has_month_word(title):
        year = _sheet_year(block, sheet_periods)
        if year:
            joined = parse_period(f"{title} {year}")
            if joined is not None and joined.detected:
                return joined
    return None


def _block_periods(
    block: Any,
    field: Any,
    sheet_periods: dict[str, list[str]],
) -> tuple[list[NormalizedPeriod], bool]:
    """Collect period evidence and separate the finest range from conflicts."""
    sources: list[Any] = []
    specific = False
    title_period = _title_period(block, sheet_periods)
    if getattr(field, "period", ""):
        sources.append(field.period)
        specific = True
        if title_period is not None:
            sources.append(title_period.label)
    elif title_period is not None:
        sources.append(title_period.label)
        specific = True
    if not specific:
        labels = sheet_periods.get(getattr(block, "sheet", ""), [])
        sources.extend(_select_sheet_periods(labels, getattr(block, "title", "")))
        sources.extend(getattr(block, "periods", []) or [])
    sources.append(getattr(field, "header", ""))
    sources.append(getattr(block, "title", ""))
    sources.append(getattr(block, "sheet", ""))
    all_periods = parse_periods(sources)
    if _has_month_word(getattr(block, "title", "")) and all_periods and all(
        period.granularity == "year" for period in all_periods
    ):
        # A month title without a parseable year must not fall back to a bare
        # year: that would make two different months look equivalent.
        return [], True
    return _finest_periods(all_periods), _periods_conflict(all_periods)


def collect_source_fields(
    analysis: Any,
    *,
    source_name: str = "",
    synonyms: dict[str, str] | None = None,
    aliases: dict[str, str] | None = None,
) -> list[SourceField]:
    """Turn the semantic blocks of an analysis into matchable source fields."""
    semantics = getattr(analysis, "semantics", None)
    if semantics is None:
        return []
    filename = source_name or getattr(semantics, "source_filename", "") or ""
    sheet_periods: dict[str, list[str]] = {}
    for sheet in getattr(getattr(analysis, "intelligence", None), "sheets", []) or []:
        name = str(getattr(sheet, "sheet", "") or "")
        if name:
            sheet_periods[name] = list(getattr(sheet, "periods", []) or [])
    result: list[SourceField] = []
    for block in getattr(semantics, "blocks", []) or []:
        block_identifiers = [
            field.header
            for field in getattr(block, "columns", []) or []
            if getattr(field, "role", None) is SemanticRole.IDENTIFIER
            and getattr(field, "header", "")
        ]
        block_identifier_values: dict[str, list[str]] = {}
        for identifier in getattr(block, "columns", []) or []:
            if (
                getattr(identifier, "role", None) is not SemanticRole.IDENTIFIER
                or not getattr(identifier, "header", "")
            ):
                continue
            values = [
                str(value)
                for value in (getattr(identifier, "samples", []) or [])
                if str(value).strip()
            ]
            if values:
                block_identifier_values[str(identifier.header)] = values[:20]
        for field in getattr(block, "columns", []) or []:
            kind = _ROLE_KIND.get(getattr(field, "role", None), KIND_OTHER)
            if kind == KIND_OTHER:
                continue
            header = str(getattr(field, "header", "") or "").strip()
            if not header:
                continue
            readable, tokens = normalize_label(header, synonyms=synonyms, aliases=aliases)
            if not tokens:
                continue
            if kind == KIND_IDENTIFIER and frozenset(tokens) in _GENERIC_IDENTIFIER_LABELS:
                continue
            periods, period_conflict = _block_periods(block, field, sheet_periods)
            if not periods and filename and not _has_month_word(getattr(block, "title", "")):
                periods = parse_periods([filename])
            unit = canonical_unit(getattr(field, "unit", ""), header)
            result.append(
                SourceField(
                    source=filename,
                    sheet=str(getattr(block, "sheet", "") or ""),
                    block=str(getattr(block, "address", "") or getattr(block, "block_id", "")),
                    header=header,
                    normalized_label=readable,
                    tokens=tokens,
                    role=getattr(getattr(field, "role", None), "value", ""),
                    kind=kind,
                    unit=unit,
                    periods=periods,
                    table_title=str(getattr(block, "title", "") or ""),
                    letter=str(getattr(field, "letter", "") or ""),
                    samples=[str(sample) for sample in (getattr(field, "samples", []) or [])][:20],
                    block_identifiers=block_identifiers,
                    identifier_values=block_identifier_values,
                    row_labels=list(getattr(block, "row_labels", []) or [])[:50],
                    evidence=list(getattr(field, "evidence", []) or []),
                    period_conflict=period_conflict,
                )
            )
    return result


# --------------------------------------------------------------------------- #
# Evaluation
# --------------------------------------------------------------------------- #
def _label_evidence(left: SourceField, right: SourceField) -> tuple[float, str]:
    left_tokens = set(left.tokens)
    right_tokens = set(right.tokens)
    if not left_tokens or not right_tokens:
        return 0.0, "یکی از برچسب‌ها واژه قابل مقایسه ندارد."
    if left_tokens == right_tokens:
        return 1.0, "مجموعه واژه‌های نرمال‌شده یکسان است."

    left_content, right_content = left.content_tokens, right.content_tokens
    left_measure, right_measure = left.measure_tokens, right.measure_tokens
    if left_content and left_content == right_content:
        if left_measure == right_measure:
            return 1.0, "مفهوم اصلی و واژه اندازه‌گیری یکسان است."
        if (
            not left_measure
            and right_measure <= _INTERCHANGEABLE_MEASURE_TOKENS
        ) or (
            not right_measure
            and left_measure <= _INTERCHANGEABLE_MEASURE_TOKENS
        ):
            return 0.95, (
                "مفهوم اصلی یکسان است و تفاوت فقط در واژه شمارش "
                "(مثل «تعداد») است."
            )
        return 0.6, (
            "مفهوم اصلی یکسان است اما یکی از دو برچسب واژه تجمیع/میانگین "
            "یا واحد اندازه‌گیری متفاوت دارد."
        )
    if not left_content and not right_content:
        if left_measure != right_measure:
            return 0.3, (
                "هر دو برچسب فقط از واژه اندازه‌گیری/تجمیع ساخته شده‌اند و "
                "این واژه‌ها یکسان نیستند."
            )
        return 0.3, "هر دو برچسب فقط از واژه عمومی اندازه‌گیری ساخته شده‌اند."
    if left_content and right_content and (
        left_content < right_content or right_content < left_content
    ):
        return 0.45, "یکی از برچسب‌ها یک قید یا واژه اضافه دارد."
    shared = left_tokens & right_tokens
    union = left_tokens | right_tokens
    jaccard = len(shared) / len(union) if union else 0.0
    if jaccard >= 0.6:
        return 0.5, f"هم‌پوشانی واژگانی {jaccard:.0%} است اما برچسب‌ها یکسان نیستند."
    return 0.0, "واژه‌های اصلی برچسب‌ها یکسان نیستند."


def _string_similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    return SequenceMatcher(None, left, right).ratio()


def _value_overlap(left: SourceField, right: SourceField) -> float | None:
    if left.kind not in {KIND_CATEGORY, KIND_IDENTIFIER}:
        return None
    left_values = {
        normalize_text(normalize_digits(str(value))).lower().strip()
        for value in left.samples
        if str(value).strip()
    }
    right_values = {
        normalize_text(normalize_digits(str(value))).lower().strip()
        for value in right.samples
        if str(value).strip()
    }
    if len(left_values) < 2 or len(right_values) < 2:
        return None
    union = left_values | right_values
    if not union:
        return None
    return len(left_values & right_values) / len(union)


def _set_overlap(left: Iterable[str], right: Iterable[str]) -> float | None:
    left_values = {
        normalize_text(normalize_digits(str(value))).lower().strip()
        for value in left
        if str(value).strip()
    }
    right_values = {
        normalize_text(normalize_digits(str(value))).lower().strip()
        for value in right
        if str(value).strip()
    }
    if len(left_values) < 2 or len(right_values) < 2:
        return None
    union = left_values | right_values
    if not union:
        return None
    return len(left_values & right_values) / len(union)


def _combined_period_status(*statuses: str) -> str:
    if "conflicting" in statuses:
        return "conflicting"
    detected = sum(status == "detected" for status in statuses)
    if detected == len(statuses):
        return "detected"
    if detected:
        return "partial"
    return "not_detected"


@dataclass
class _PairEvaluation:
    left: int
    right: int
    status: MatchStatus
    score: float
    label_score: float
    label_reason: str
    unit: UnitComparison
    period: PeriodCompatibility
    role_compatible: bool
    evidence: list[str] = field(default_factory=list)
    reason: str = ""
    candidate: bool = False
    similar: bool = False
    comparable: bool = False

    def other(self, index: int) -> int:
        return self.right if index == self.left else self.left


class SemanticMatcher:
    """Match extracted fields across worksheets and files without aggregating."""

    def __init__(
        self,
        settings: Settings,
        *,
        aliases: dict[str, str] | None = None,
        matched_threshold: float = 0.9,
        candidate_threshold: float = 0.45,
        ambiguity_margin: float = 0.08,
        max_rejections: int = 200,
    ) -> None:
        self.settings = settings
        self.aliases = dict(aliases or {})
        self.matched_threshold = matched_threshold
        self.candidate_threshold = candidate_threshold
        self.ambiguity_margin = ambiguity_margin
        self.max_rejections = max_rejections

    # ------------------------------------------------------------------ #
    def match_fields(self, fields: Iterable[SourceField]) -> SemanticMatchingReport:
        field_list = list(fields)
        report = SemanticMatchingReport(fields=field_list)
        if len(field_list) < 2:
            report.warnings.append("برای تطبیق، دست‌کم دو فیلد از دو منبع لازم است.")
            return report

        evaluations: list[_PairEvaluation] = []
        for left in range(len(field_list)):
            for right in range(left + 1, len(field_list)):
                evaluation = self._evaluate(field_list, left, right)
                if evaluation is not None:
                    evaluations.append(evaluation)

        candidates_by_field: dict[int, list[_PairEvaluation]] = {
            index: [] for index in range(len(field_list))
        }
        for evaluation in evaluations:
            if evaluation.candidate:
                candidates_by_field[evaluation.left].append(evaluation)
                candidates_by_field[evaluation.right].append(evaluation)

        best: dict[int, _PairEvaluation] = {}
        ambiguous: set[int] = set()
        for index, items in candidates_by_field.items():
            if not items:
                continue
            items.sort(key=lambda item: item.score, reverse=True)
            best[index] = items[0]
            if len(items) >= 2 and items[0].score - items[1].score < self.ambiguity_margin:
                ambiguous.add(index)

        # Union only mutually best, unambiguous, fully proven pairs.
        parent = list(range(len(field_list)))

        def find(node: int) -> int:
            while parent[node] != node:
                parent[node] = parent[parent[node]]
                node = parent[node]
            return node

        def union(a: int, b: int) -> None:
            root_a, root_b = find(a), find(b)
            if root_a != root_b:
                parent[root_b] = root_a

        matched_pairs: list[_PairEvaluation] = []
        for index, evaluation in best.items():
            if index in ambiguous or evaluation.status is not MatchStatus.MATCHED:
                continue
            partner = evaluation.other(index)
            if partner in ambiguous:
                continue
            if best.get(partner) is not evaluation:
                continue
            if index < partner:
                matched_pairs.append(evaluation)
                union(index, partner)

        groups: dict[int, list[int]] = {}
        for index in range(len(field_list)):
            groups.setdefault(find(index), []).append(index)
        matched_indices: set[int] = set()
        for root, members in groups.items():
            if len(members) < 2:
                continue
            pair = next(
                (
                    item
                    for item in matched_pairs
                    if item.left in members and item.right in members
                ),
                None,
            )
            if pair is None:
                continue
            matched_indices.update(members)
            members.sort(key=lambda item: field_list[item].identity)
            member_set = set(members)
            group_evaluations = [
                evaluation
                for evaluation in evaluations
                if evaluation.candidate
                and evaluation.left in member_set
                and evaluation.right in member_set
            ]
            if not group_evaluations:
                continue
            report.matches.append(
                self._group_match(
                    len(report.matches) + 1,
                    [field_list[item] for item in members],
                    group_evaluations,
                )
            )

        # Fields that are not in a proven group: report a probable pair or an
        # explained ambiguity.  Ambiguities are emitted once per candidate set.
        probable_pairs: set[frozenset[int]] = set()
        ambiguity_sets: set[frozenset[int]] = set()
        for index in range(len(field_list)):
            if index in matched_indices:
                continue
            items = sorted(
                candidates_by_field.get(index, []),
                key=lambda item: item.score,
                reverse=True,
            )
            if not items:
                continue
            top = items[0]
            partner = top.other(index)
            if partner in matched_indices:
                continue
            close = len(items) >= 2 and items[0].score - items[1].score < self.ambiguity_margin
            mutual = best.get(partner) is top
            if (
                not close
                and index not in ambiguous
                and partner not in ambiguous
                and mutual
                and top.status is MatchStatus.PROBABLE
            ):
                pair_key = frozenset({index, partner})
                if pair_key in probable_pairs:
                    continue
                probable_pairs.add(pair_key)
                report.probable.append(
                    self._pair_match(
                        len(report.matches) + len(report.probable) + 1,
                        field_list[index],
                        field_list[partner],
                        top,
                    )
                )
                continue
            if len(items) == 1 and partner in ambiguous:
                # The partner's own record already lists this candidate.
                continue
            candidate_indices = [index, *(item.other(index) for item in items[:3])]
            candidate_set = frozenset(candidate_indices)
            if candidate_set in ambiguity_sets:
                continue
            ambiguity_sets.add(candidate_set)
            report.ambiguities.append(
                self._ambiguity_match(
                    len(report.matches) + len(report.probable) + len(report.ambiguities) + 1,
                    field_list[index],
                    [field_list[item.other(index)] for item in items[:3]],
                    items[:3],
                )
            )

        # Similar-looking pairs that were rejected, with the reason.
        seen_rejections: set[frozenset[int]] = set()
        for evaluation in evaluations:
            if evaluation.status is not MatchStatus.REJECTED or not evaluation.similar:
                continue
            key = frozenset({evaluation.left, evaluation.right})
            if key in seen_rejections:
                continue
            seen_rejections.add(key)
            report.rejections.append(
                RejectedPair(
                    left=field_list[evaluation.left],
                    right=field_list[evaluation.right],
                    reason=evaluation.reason,
                    label_score=evaluation.label_score,
                    period_relation=evaluation.period.relation,
                    unit_state=evaluation.unit.state,
                    period_status=_combined_period_status(
                        field_list[evaluation.left].period_status,
                        field_list[evaluation.right].period_status,
                    ),
                )
            )
            if len(report.rejections) >= self.max_rejections:
                report.warnings.append(
                    f"فهرست موارد رد‌شده به سقف {self.max_rejections} مورد محدود شد."
                )
                break

        if not report.matches and not report.probable and not report.ambiguities:
            report.warnings.append(
                "هیچ تطبیق قطعی یا احتمالی بین فیلدهای منابع پیدا نشد؛ "
                "نام‌ها، واحد یا دوره گزارش‌دهی برای برابری کافی نبودند."
            )
        if report.ambiguities:
            report.warnings.append(
                f"{len(report.ambiguities)} فیلد بیش از یک گزینه نزدیک دارد؛ "
                "گزینه‌ها برای بررسی کاربر نگه داشته شدند و هیچ‌کدام قطعی اعلام نشد."
            )
        return report

    # ------------------------------------------------------------------ #
    def match_analysis(self, analysis: Any, *, source_name: str = "") -> SemanticMatchingReport:
        fields = collect_source_fields(
            analysis,
            source_name=source_name,
            synonyms=getattr(self.settings, "synonyms", {}) or {},
            aliases=self.aliases,
        )
        return self.match_fields(fields)

    # ------------------------------------------------------------------ #
    def _evaluate(
        self,
        fields: list[SourceField],
        left: int,
        right: int,
    ) -> _PairEvaluation | None:
        first, second = fields[left], fields[right]
        role_compatible = first.kind == second.kind
        label_score, label_reason = _label_evidence(first, second)
        string_similarity = _string_similarity(first.normalized_label, second.normalized_label)
        unit = compare_units(first.unit, second.unit)
        period = compare_period_sets(first.periods, second.periods)
        period_conflict = first.period_conflict or second.period_conflict

        evidence = [label_reason]
        overlap = _value_overlap(first, second)
        value_supported = False
        if overlap is not None:
            if overlap >= 0.5:
                evidence.append(f"هم‌پوشانی مقادیر دسته‌ای {overlap:.0%}؛ شاهد کمکی.")
                label_score = min(1.0, label_score + 0.1)
            elif first.kind == KIND_CATEGORY:
                evidence.append(f"هم‌پوشانی مقادیر دسته‌ای فقط {overlap:.0%} است.")
            if (
                overlap >= 0.8
                and label_score < self.candidate_threshold
                and first.kind in {KIND_CATEGORY, KIND_IDENTIFIER}
            ):
                # Identical values are real evidence, but names alone cannot prove
                # the two columns mean the same thing.  Keep it probable, never
                # a definitive match.
                label_score = 0.5
                value_supported = True
                evidence.append(
                    "نام‌ها یکسان نیستند اما مقادیر نمونه هم‌پوشانی بالایی دارند؛ "
                    "برای تطبیق قطعی به تأیید نیاز است."
                )
        shared_identifiers = set(first.block_identifiers) & set(second.block_identifiers)
        for identifier in sorted(shared_identifiers)[:3]:
            identifier_overlap = _set_overlap(
                first.identifier_values.get(identifier, []),
                second.identifier_values.get(identifier, []),
            )
            if identifier_overlap is not None:
                evidence.append(
                    f"ستون شناسه مشترک «{identifier}» با هم‌پوشانی مقادیر "
                    f"{identifier_overlap:.0%}."
                )
            else:
                evidence.append(f"هر دو بلوک ستون شناسه مشترک دارند: {identifier}")

        row_overlap = _set_overlap(first.row_labels, second.row_labels)
        if row_overlap is not None:
            if row_overlap >= 0.5:
                evidence.append(
                    f"هم‌پوشانی برچسب‌های سطر {row_overlap:.0%}؛ ساختار سطرها "
                    "هم‌خوان است."
                )
            else:
                evidence.append(
                    f"برچسب‌های سطر فقط {row_overlap:.0%} هم‌پوشانی دارند."
                )

        if first.table_title and second.table_title:
            title_similarity = _string_similarity(
                normalize_text(normalize_digits(first.table_title)).lower(),
                normalize_text(normalize_digits(second.table_title)).lower(),
            )
            if title_similarity >= 0.6:
                evidence.append(
                    f"عنوان جدول‌ها هم‌خوان است (شباهت {title_similarity:.0%})."
                )
        if period_conflict:
            evidence.append(
                "برای دست‌کم یکی از دو فیلد دوره گزارش‌دهی قطعی نیست یا بیش از یک دوره ناسازگار ثبت شده است."
            )

        aggregation_difference = (
            first.measure_tokens != second.measure_tokens
            and bool(
                (first.measure_tokens | second.measure_tokens)
                & _AGGREGATION_MEASURE_TOKENS
            )
        )
        similar = (
            label_score >= self.candidate_threshold
            or string_similarity >= 0.7
            or aggregation_difference
        )
        # Cross-role pairs are only interesting when the names are close; every
        # other cross-role pair would flood the report with structural noise.
        if not role_compatible and not similar:
            return None

        reasons: list[str] = []
        if not role_compatible:
            reasons.append(
                f"نقش معنایی متفاوت است ({first.kind_label} در برابر {second.kind_label})."
            )
        if unit.state == "incompatible":
            reasons.append(unit.note)

        if reasons:
            return _PairEvaluation(
                left=left,
                right=right,
                status=MatchStatus.REJECTED,
                score=0.0,
                label_score=label_score,
                label_reason=label_reason,
                unit=unit,
                period=period,
                role_compatible=role_compatible,
                evidence=evidence,
                reason=" ".join(reasons),
                candidate=False,
                similar=similar,
            )

        if value_supported:
            return _PairEvaluation(
                left=left,
                right=right,
                status=MatchStatus.PROBABLE,
                score=min(0.85, 0.55 + 0.3 * (overlap or 0.0)),
                label_score=label_score,
                label_reason=label_reason,
                unit=unit,
                period=period,
                role_compatible=role_compatible,
                evidence=evidence,
                reason=(
                    "نام‌ها یکسان نیستند اما مقادیر نمونه هم‌پوشانی بالایی دارند؛ "
                    "این تطبیق احتمالی است و مقدارها بدون تأیید مقایسه نمی‌شوند."
                    + (
                        " " + period.explanation
                        if period.relation != "equivalent"
                        else ""
                    )
                ),
                candidate=True,
                similar=True,
                comparable=False,
            )

        if label_score >= self.matched_threshold:
            # Identity and value comparability are separate decisions.  The same
            # indicator in two different periods is still one indicator, but its
            # two values are not equivalent and must not be compared as such.
            unit_identity_ok = unit.state in {"equal", "both_missing", "scalable"}
            period_known = period.relation != "unknown"
            if unit_identity_ok and period_known and not period_conflict:
                status = MatchStatus.MATCHED
            else:
                status = MatchStatus.PROBABLE
            comparable = (
                unit.state in {"equal", "both_missing"}
                and period.equivalent
                and not period_conflict
            )

            notes: list[str] = []
            if period_conflict:
                notes.append(
                    "دوره گزارش‌دهی برای همین فیلد قطعی نیست یا بیش از یک دوره ناسازگار دارد؛ "
                    "تا رفع ابهام، برابری مقدارها اثبات‌شدنی نیست."
                )
            if unit.state in {"unknown", "same_dimension"}:
                notes.append(unit.note)
            if unit.state == "scalable":
                notes.append(unit.note)
            if period.relation == "unknown":
                notes.append(period.explanation)
            elif not period.equivalent:
                notes.append(
                    "همان شاخص است اما دوره گزارش‌دهی یکسان نیست؛ مقدارها را "
                    "نمی‌توان بدون هم‌ترازسازی دوره برابر دانست. " + period.explanation
                )
            if comparable:
                if unit.state == "both_missing":
                    reason = (
                        "برچسب نرمال‌شده و دوره گزارش‌دهی سازگارند؛ واحد در هیچ منبعی "
                        "ثبت نشده و برابری واحد اثبات نشده است."
                    )
                else:
                    reason = "برچسب نرمال‌شده، واحد و دوره گزارش‌دهی با هم سازگارند."
            elif status is MatchStatus.MATCHED:
                reason = " ".join(notes)
            else:
                reason = " ".join(notes) or "برای اثبات برابری شاهد کافی نیست."
            score = (
                label_score * 0.6
                + unit.score() * 0.15
                + self._period_score(period) * 0.25
            )
            return _PairEvaluation(
                left=left,
                right=right,
                status=status,
                score=score,
                label_score=label_score,
                label_reason=label_reason,
                unit=unit,
                period=period,
                role_compatible=role_compatible,
                evidence=evidence,
                reason=reason,
                candidate=True,
                similar=True,
                comparable=comparable,
            )

        if label_score >= self.candidate_threshold:
            return _PairEvaluation(
                left=left,
                right=right,
                status=MatchStatus.AMBIGUOUS,
                score=label_score,
                label_score=label_score,
                label_reason=label_reason,
                unit=unit,
                period=period,
                role_compatible=role_compatible,
                evidence=evidence,
                reason=(
                    "نام‌ها فقط بخشی از واژه‌ها را مشترک دارند؛ برای تطبیق قطعی کافی نیست. "
                    + label_reason
                ),
                candidate=True,
                similar=True,
                comparable=False,
            )

        return _PairEvaluation(
            left=left,
            right=right,
            status=MatchStatus.REJECTED,
            score=0.0,
            label_score=label_score,
            label_reason=label_reason,
            unit=unit,
            period=period,
            role_compatible=role_compatible,
            evidence=evidence,
            reason=f"نام‌ها به اندازه کافی یکسان نیستند. {label_reason}",
            candidate=False,
            similar=similar,
        )

    @staticmethod
    def _period_score(period: PeriodCompatibility) -> float:
        return {
            "equivalent": 1.0,
            "contains": 0.8,
            "within": 0.8,
            "overlapping": 0.6,
            "unknown": 0.2,
            "incompatible": 0.0,
        }.get(period.relation, 0.0)

    # ------------------------------------------------------------------ #
    def _group_match(
        self,
        number: int,
        members: list[SourceField],
        evaluations: list[_PairEvaluation],
    ) -> FieldMatch:
        evidence: list[str] = []
        for field_item in members:
            evidence.append(
                f"{field_item.source} / {field_item.sheet} / {field_item.block} / "
                f"«{field_item.header}»"
            )
        seen_evidence: set[str] = set()
        for evaluation in evaluations:
            for text in evaluation.evidence:
                if text not in seen_evidence:
                    seen_evidence.add(text)
                    evidence.append(text)

        period_relations = {evaluation.period.relation for evaluation in evaluations}
        if period_relations == {"equivalent"}:
            period_relation = "equivalent"
        elif "incompatible" in period_relations:
            period_relation = "incompatible"
        elif len(period_relations) == 1:
            period_relation = next(iter(period_relations))
        else:
            period_relation = "mixed"

        unit_states = {evaluation.unit.state for evaluation in evaluations}
        if unit_states == {"equal"}:
            unit_state = "equal"
        elif "scalable" in unit_states:
            unit_state = "scalable"
        elif unit_states == {"both_missing"}:
            unit_state = "both_missing"
        elif len(unit_states) == 1:
            unit_state = next(iter(unit_states))
        else:
            unit_state = "mixed"

        explanations: list[str] = []
        for evaluation in evaluations:
            text = evaluation.period.explanation
            if text and text not in explanations:
                explanations.append(text)
        comparable = all(evaluation.comparable for evaluation in evaluations)
        pair_reasons: list[str] = []
        for evaluation in evaluations:
            if evaluation.reason and evaluation.reason not in pair_reasons:
                pair_reasons.append(evaluation.reason)
        reason = " ".join(pair_reasons[:2]) or "برچسب نرمال‌شده سازگار است."
        if not comparable:
            reason += (
                " مقدارهای اعضا با هم قابل مقایسه نیستند؛ دوره گزارش‌دهی یا واحد "
                "یکی نیست و هیچ تجمیعی انجام نشد."
            )
            detail = next(
                (
                    evaluation.reason
                    for evaluation in evaluations
                    if evaluation.reason and not evaluation.comparable
                ),
                "",
            )
            if detail and detail not in reason:
                reason += " " + detail
        return FieldMatch(
            match_id=f"M{number}",
            status=MatchStatus.MATCHED,
            kind=members[0].kind,
            normalized_label=members[0].normalized_label,
            confidence=sum(item.score for item in evaluations) / len(evaluations),
            members=members,
            evidence=evidence,
            period_relation=period_relation,
            period_explanation=" ".join(explanations[:3]),
            period_status=_combined_period_status(*(member.period_status for member in members)),
            unit_state=unit_state,
            reason=reason,
            comparable=comparable,
        )

    def _pair_match(
        self,
        number: int,
        left: SourceField,
        right: SourceField,
        evaluation: _PairEvaluation,
    ) -> FieldMatch:
        return FieldMatch(
            match_id=f"P{number}",
            status=MatchStatus.PROBABLE,
            kind=left.kind,
            normalized_label=left.normalized_label,
            confidence=evaluation.score,
            members=[left, right],
            evidence=list(evaluation.evidence),
            period_relation=evaluation.period.relation,
            period_explanation=evaluation.period.explanation,
            period_status=_combined_period_status(left.period_status, right.period_status),
            unit_state=evaluation.unit.state,
            reason=evaluation.reason,
            comparable=evaluation.comparable,
        )

    def _ambiguity_match(
        self,
        number: int,
        anchor: SourceField,
        candidates: list[SourceField],
        evaluations: list[_PairEvaluation],
    ) -> FieldMatch:
        candidate_records: list[MatchCandidate] = []
        for field_item, evaluation in zip(candidates, evaluations):
            candidate_records.append(
                MatchCandidate(
                    field=field_item,
                    status=evaluation.status,
                    score=evaluation.score,
                    evidence=list(evaluation.evidence),
                    reason=evaluation.reason,
                    comparable=evaluation.comparable,
                )
            )
        return FieldMatch(
            match_id=f"A{number}",
            status=MatchStatus.AMBIGUOUS,
            kind=anchor.kind,
            normalized_label=anchor.normalized_label,
            confidence=candidate_records[0].score if candidate_records else 0.0,
            members=[anchor],
            evidence=[
                f"فیلد «{anchor.source} / {anchor.sheet} / {anchor.header}» "
                f"{len(candidate_records)} گزینه نزدیک دارد."
            ],
            period_status=_combined_period_status(
                anchor.period_status, *(candidate.period_status for candidate in candidates)
            ),
            candidates=candidate_records,
            reason=(
                "چند فیلد به یک اندازه نزدیک هستند؛ تطبیق خودکار انجام نشد تا "
                "کاربر یکی را تأیید کند."
            ),
            comparable=False,
        )


# --------------------------------------------------------------------------- #
# Cross-source helper
# --------------------------------------------------------------------------- #
def match_across_sources(
    settings: Settings,
    analyses: Iterable[tuple[str, Any]],
    *,
    aliases: dict[str, str] | None = None,
) -> SemanticMatchingReport:
    """Collect fields from several ``(source_name, analysis)`` pairs and match."""
    fields: list[SourceField] = []
    for source_name, analysis in analyses:
        fields.extend(
            collect_source_fields(
                analysis,
                source_name=source_name,
                synonyms=getattr(settings, "synonyms", {}) or {},
                aliases=aliases,
            )
        )
    return SemanticMatcher(settings, aliases=aliases).match_fields(fields)
