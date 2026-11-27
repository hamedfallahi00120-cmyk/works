"""Deterministic reconciliation and calculation-integrity evidence.

The engine never decides that a set of rows *can* be summed just because it is
numeric.  Every aggregation needs an explicit, checkable relationship:

* components must be distinct and non-overlapping;
* a grand total is never one of its own components;
* a parent category is never added to its children;
* a percentage, rate or difference row is never added to a count or an amount;
* components must share a unit.

When that relationship cannot be established the result is a failure with a
reason, not a silently summed figure.  A missing component is never treated as
zero.

Every arithmetic operation happens here, in Python, using :class:`Decimal`.
The language model may later explain a result, but its numbers are checked
against the values recorded by this module before they reach a report.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from enum import Enum
from typing import Any, Iterable, Mapping, Sequence

from local_office_ai.core.models import AnalysisResult, TableCandidate, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import normalize_digits, normalize_text, parse_value
from local_office_ai.excel.row_classification import (
    ROW_CATEGORY,
    ROW_DATA,
    ROW_DATA_MISSING,
    ROW_GROUP_HEADING,
    ROW_SUBTOTAL,
    ROW_TOTAL,
    classify_block_rows,
)
from local_office_ai.excel.semantic_classification import (
    DATA_ROLE_TO_SEMANTIC,
    MEASURE_ROLES,
    SemanticBlock,
    SemanticField,
    SemanticRole,
)
from local_office_ai.excel.sheet_analyzer import SheetGrid
from local_office_ai.excel.table_detector import is_report_sheet

ZERO = Decimal("0")
ONE = Decimal("1")
HUNDRED = Decimal("100")


class ReconciliationStatus(str, Enum):
    PASS = "PASS"
    WARNING = "WARNING"
    FAIL = "FAIL"


STATUS_FA = {
    ReconciliationStatus.PASS: "قبول",
    ReconciliationStatus.WARNING: "هشدار",
    ReconciliationStatus.FAIL: "رد",
}

KIND_COMPONENT_TOTAL = "component_total"
KIND_SUBTOTAL_GRAND_TOTAL = "subtotal_grand_total"
KIND_GROUP_SUBGROUP = "group_subgroup"
KIND_MONTHLY_PERIOD = "monthly_period"
KIND_CATEGORY_COMPARISON = "category_comparison"
KIND_PERIOD_COMPARISON = "period_comparison"

KIND_FA = {
    KIND_COMPONENT_TOTAL: "جزء به کل",
    KIND_SUBTOTAL_GRAND_TOTAL: "جمع میانی به جمع کل",
    KIND_GROUP_SUBGROUP: "گروه به زیرگروه",
    KIND_MONTHLY_PERIOD: "ماه به دوره",
    KIND_CATEGORY_COMPARISON: "مقایسه دسته‌ای",
    KIND_PERIOD_COMPARISON: "مقایسه بین‌دوره‌ای",
}

ROLE_DATA = "data"
ROLE_CATEGORY = "category"
ROLE_SUBTOTAL = "subtotal"
ROLE_TOTAL = "total"
ROLE_PERCENTAGE = "percentage"
ROLE_DIFFERENCE = "difference"

_FORBIDDEN_COMPONENT_ROLES = {ROLE_TOTAL, ROLE_PERCENTAGE, ROLE_DIFFERENCE}
_SUBTOTAL_KINDS = {KIND_SUBTOTAL_GRAND_TOTAL, KIND_GROUP_SUBGROUP}

_RATIO_WORDS = ("درصد", "نرخ", "نسبت", "ratio", "rate", "percent", "percentage")
_WEIGHT_WORDS = (
    "تعداد", "حجم", "مبلغ", "وزن", "فراوانی", "مقدار",
    "count", "weight", "quantity", "amount", "volume",
)


def _text(value: Decimal | None) -> str | None:
    if value is None:
        return None
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f")


def _jsonable(value: Decimal | None) -> str | None:
    return _text(value)


def _quantize(value: Decimal, precision: int) -> Decimal:
    quantum = Decimal(1).scaleb(-max(0, int(precision)))
    return value.quantize(quantum, rounding=ROUND_HALF_UP)


def _normalized(value: object) -> str:
    if not isinstance(value, str):
        return ""
    return normalize_text(normalize_digits(value)).lower().strip()


# --------------------------------------------------------------------------- #
# Tolerance
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Tolerance:
    """Absolute and relative room for a legitimate rounding difference.

    A difference inside ``absolute``/``relative`` is a pass.  A difference up to
    ``warning_factor`` times that limit is a warning (rounding, display
    precision or a partial source); anything larger is a failure.  The limit is
    never used to hide a missing value or an invalid relationship.
    """

    absolute: Decimal = Decimal("0.01")
    relative: Decimal = ZERO
    warning_factor: Decimal = Decimal("10")

    @classmethod
    def for_decimals(
        cls,
        decimals: int,
        *,
        components: int = 1,
        relative: Decimal = ZERO,
        warning_factor: Decimal = Decimal("10"),
    ) -> "Tolerance":
        """Half a unit in the last displayed place, accumulated over components."""

        decimals = max(0, int(decimals))
        unit = Decimal(1).scaleb(-decimals)
        absolute = unit * Decimal("0.5") * Decimal(max(1, int(components)))
        return cls(
            absolute=absolute.normalize(),
            relative=relative,
            warning_factor=warning_factor,
        )

    @classmethod
    def for_workbook(
        cls,
        *,
        decimals: int = 2,
        components: int = 1,
    ) -> "Tolerance":
        """Rounding tolerance for real workbooks with large monetary totals."""

        return cls.for_decimals(
            decimals,
            components=components,
            relative=Decimal("0.000001"),
        )

    def limit(self, expected: Decimal | None) -> Decimal:
        limit = self.absolute
        if expected is not None and self.relative > ZERO:
            limit = max(limit, abs(expected) * self.relative)
        return limit

    def warning_limit(self, expected: Decimal | None) -> Decimal:
        return self.limit(expected) * self.warning_factor

    def classify(self, difference: Decimal | None, expected: Decimal | None) -> ReconciliationStatus:
        if difference is None:
            return ReconciliationStatus.WARNING
        gap = abs(difference)
        if gap <= self.limit(expected):
            return ReconciliationStatus.PASS
        if gap <= self.warning_limit(expected):
            return ReconciliationStatus.WARNING
        return ReconciliationStatus.FAIL

    def describe_fa(self, expected: Decimal | None = None) -> str:
        limit = self.limit(expected)
        warning = self.warning_limit(expected)
        return (
            f"تحمل تأیید {_text(limit)} و آستانه هشدار {_text(warning)}"
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "absolute": _text(self.absolute),
            "relative": _text(self.relative),
            "warning_factor": _text(self.warning_factor),
        }


# --------------------------------------------------------------------------- #
# Change and average calculations
# --------------------------------------------------------------------------- #
@dataclass
class ChangeResult:
    """Absolute and percentage change with an explicit baseline state."""

    current: Decimal | None
    previous: Decimal | None
    absolute: Decimal | None = None
    percent: Decimal | None = None
    direction: str = "قابل محاسبه نیست"
    baseline: str = "missing"
    status: ReconciliationStatus = ReconciliationStatus.WARNING
    explanation: str = ""
    label: str = ""
    unit: str = ""
    source: str = ""

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    @property
    def verified(self) -> bool:
        return self.status is ReconciliationStatus.PASS

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "unit": self.unit,
            "current": _text(self.current),
            "previous": _text(self.previous),
            "absolute": _text(self.absolute),
            "percent": _text(self.percent),
            "direction": self.direction,
            "baseline": self.baseline,
            "status": self.status.value,
            "status_label": self.status_label,
            "verified": self.verified,
            "explanation": self.explanation,
            "source": self.source,
        }


def compute_change(
    current: Decimal | None,
    previous: Decimal | None,
    *,
    precision: int = 2,
    label: str = "",
    unit: str = "",
    source: str = "",
) -> ChangeResult:
    """Absolute and percentage change; never divides by zero.

    Direction comes from the absolute difference, so a negative baseline cannot
    turn an increase into a reported decrease.  The percentage uses the
    magnitude of the baseline, which is the standard relative-change reading for
    administrative figures.  A zero or missing baseline keeps ``percent`` as
    ``None`` with a warning; no percentage is fabricated.
    """

    if current is None or previous is None:
        missing = "مقدار فعلی" if current is None else "مقدار مبنا"
        return ChangeResult(
            current=current,
            previous=previous,
            direction="قابل محاسبه نیست",
            baseline="missing",
            status=ReconciliationStatus.WARNING,
            explanation=f"{missing} موجود نیست؛ تغییر مطلق و درصدی محاسبه نشد.",
            label=label,
            unit=unit,
            source=source,
        )

    difference = current - previous
    if difference > ZERO:
        direction = "افزایش"
    elif difference < ZERO:
        direction = "کاهش"
    else:
        direction = "بدون تغییر"

    if previous == ZERO:
        return ChangeResult(
            current=current,
            previous=previous,
            absolute=difference,
            percent=None,
            direction=direction,
            baseline="zero",
            status=ReconciliationStatus.WARNING,
            explanation=(
                f"مبنای قبلی صفر است؛ تغییر مطلق {_text(difference)} است اما درصد تغییر "
                "تعریف‌شده نیست و محاسبه نشد."
            ),
            label=label,
            unit=unit,
            source=source,
        )

    percent = _quantize(difference / abs(previous) * HUNDRED, precision)
    return ChangeResult(
        current=current,
        previous=previous,
        absolute=difference,
        percent=percent,
        direction=direction,
        baseline="valid",
        status=ReconciliationStatus.PASS,
        explanation=(
            f"{direction} به مقدار {_text(difference)} واحد و {_text(percent)} درصد "
            f"نسبت به مبنای {_text(previous)}."
        ),
        label=label,
        unit=unit,
        source=source,
    )


@dataclass
class RatioResult:
    numerator: Decimal | None
    denominator: Decimal | None
    value: Decimal | None
    status: ReconciliationStatus
    explanation: str
    label: str = ""
    unit: str = ""

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "unit": self.unit,
            "numerator": _text(self.numerator),
            "denominator": _text(self.denominator),
            "value": _text(self.value),
            "status": self.status.value,
            "status_label": self.status_label,
            "explanation": self.explanation,
        }


def compute_ratio(
    numerator: Decimal | None,
    denominator: Decimal | None,
    *,
    precision: int = 4,
    label: str = "",
    unit: str = "",
) -> RatioResult:
    """A ratio only exists when both parts exist and the denominator is non-zero."""

    if numerator is None or denominator is None:
        return RatioResult(
            numerator=numerator,
            denominator=denominator,
            value=None,
            status=ReconciliationStatus.WARNING,
            explanation="صورت یا مخرج نسبت موجود نیست؛ مقدار نسبت محاسبه نشد.",
            label=label,
            unit=unit,
        )
    if denominator == ZERO:
        return RatioResult(
            numerator=numerator,
            denominator=denominator,
            value=None,
            status=ReconciliationStatus.WARNING,
            explanation="مخرج نسبت صفر است؛ تقسیم انجام نشد و مقداری ساخته نشد.",
            label=label,
            unit=unit,
        )
    value = _quantize(numerator / denominator, precision)
    return RatioResult(
        numerator=numerator,
        denominator=denominator,
        value=value,
        status=ReconciliationStatus.PASS,
        explanation=f"نسبت از تقسیم صورت {_text(numerator)} بر مخرج {_text(denominator)} محاسبه شد.",
        label=label,
        unit=unit,
    )


@dataclass
class AverageResult:
    value: Decimal | None
    unweighted: Decimal | None
    weighted: Decimal | None
    weights_required: bool
    weights_available: bool
    weighted_used: bool
    status: ReconciliationStatus
    explanation: str
    label: str = ""
    unit: str = ""
    observations: int = 0

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "unit": self.unit,
            "value": _text(self.value),
            "unweighted": _text(self.unweighted),
            "weighted": _text(self.weighted),
            "weights_required": self.weights_required,
            "weights_available": self.weights_available,
            "weighted_used": self.weighted_used,
            "observations": self.observations,
            "status": self.status.value,
            "status_label": self.status_label,
            "explanation": self.explanation,
        }


def _simple_mean(values: Sequence[Decimal]) -> Decimal | None:
    if not values:
        return None
    return sum(values, ZERO) / Decimal(len(values))


def _weighted_mean(
    values: Sequence[Decimal],
    weights: Sequence[Decimal | None] | None,
) -> tuple[Decimal | None, bool]:
    if weights is None or len(weights) != len(values):
        return None, False
    usable: list[tuple[Decimal, Decimal]] = []
    for value, weight in zip(values, weights):
        if weight is None or weight < ZERO:
            return None, False
        usable.append((value, weight))
    total_weight = sum((weight for _, weight in usable), ZERO)
    if total_weight <= ZERO:
        return None, False
    return sum((value * weight for value, weight in usable), ZERO) / total_weight, True


def compute_average(
    values: Sequence[Decimal],
    weights: Sequence[Decimal | None] | None = None,
    *,
    weights_required: bool = False,
    precision: int = 4,
    label: str = "",
    unit: str = "",
) -> AverageResult:
    """Simple or weighted mean, refusing a simple mean when weighting is required.

    ``weights_required`` is set by the caller when the figures are rates or
    ratios whose denominator is the weight.  If the weights are missing, the
    result is ``None`` with ``FAIL``; an ordinary average is not substituted.
    """

    observations = len(values)
    if not values:
        return AverageResult(
            value=None,
            unweighted=None,
            weighted=None,
            weights_required=weights_required,
            weights_available=False,
            weighted_used=False,
            status=ReconciliationStatus.FAIL,
            explanation="هیچ مشاهده عددی برای میانگین وجود ندارد.",
            label=label,
            unit=unit,
        )

    unweighted = _quantize(_simple_mean(values) or ZERO, precision)
    weighted, weights_available = _weighted_mean(values, weights)

    if weights_required:
        if not weights_available or weighted is None:
            return AverageResult(
                value=None,
                unweighted=unweighted,
                weighted=None,
                weights_required=True,
                weights_available=weights_available,
                weighted_used=False,
                status=ReconciliationStatus.FAIL,
                explanation=(
                    "برای این نسبت میانگین وزنی لازم است، اما وزن یا مخرج متناظر موجود نیست؛ "
                    "میانگین ساده به‌جای آن محاسبه نشد."
                ),
                label=label,
                unit=unit,
                observations=observations,
            )
        return AverageResult(
            value=_quantize(weighted, precision),
            unweighted=unweighted,
            weighted=_quantize(weighted, precision),
            weights_required=True,
            weights_available=True,
            weighted_used=True,
            status=ReconciliationStatus.PASS,
            explanation="میانگین وزنی با مجموع وزن‌های موجود محاسبه شد.",
            label=label,
            unit=unit,
            observations=observations,
        )

    return AverageResult(
        value=unweighted,
        unweighted=unweighted,
        weighted=_quantize(weighted, precision) if weighted is not None else None,
        weights_required=False,
        weights_available=weights_available,
        weighted_used=False,
        status=ReconciliationStatus.PASS,
        explanation="میانگین ساده محاسبه شد؛ وزنی برای این شاخص لازم نبود.",
        label=label,
        unit=unit,
        observations=observations,
    )


# --------------------------------------------------------------------------- #
# Aggregation safety
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class AggregationComponent:
    label: str
    value: Decimal | None
    role: str = ROLE_DATA
    unit: str = ""
    categories: frozenset[str] = frozenset()
    source: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "value": _text(self.value),
            "role": self.role,
            "unit": self.unit,
            "categories": sorted(self.categories),
            "source": self.source,
        }


@dataclass
class AggregationEvidence:
    justificed: bool
    reasons: list[str] = field(default_factory=list)
    overlap: list[str] = field(default_factory=list)
    missing: list[str] = field(default_factory=list)
    mixed_units: list[str] = field(default_factory=list)
    forbidden_roles: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "justificed": self.justificed,
            "reasons": list(self.reasons),
            "overlap": list(self.overlap),
            "missing": list(self.missing),
            "mixed_units": list(self.mixed_units),
            "forbidden_roles": list(self.forbidden_roles),
        }


def assess_aggregation(
    components: Sequence[AggregationComponent],
    *,
    kind: str = KIND_COMPONENT_TOTAL,
    unit: str = "",
    distinct_labels: bool = False,
) -> AggregationEvidence:
    """Decide whether the components may be summed, and say why not."""

    reasons: list[str] = []
    overlap: list[str] = []
    missing: list[str] = []
    mixed_units: list[str] = []
    forbidden_roles: list[str] = []

    if not components:
        reasons.append("هیچ جزء عددی برای جمع وجود ندارد.")
        return AggregationEvidence(
            justificed=False,
            reasons=reasons,
            missing=missing,
            mixed_units=mixed_units,
            forbidden_roles=forbidden_roles,
        )

    seen_roles: set[str] = set()
    seen_units: set[str] = set()
    category_owners: dict[str, list[str]] = {}
    label_counts: dict[str, int] = {}
    for component in components:
        seen_roles.add(component.role)
        if component.unit:
            seen_units.add(component.unit)
        label_counts[component.label] = label_counts.get(component.label, 0) + 1
        for category in component.categories:
            category_owners.setdefault(category, []).append(component.label)

    for role in sorted(seen_roles & _FORBIDDEN_COMPONENT_ROLES):
        forbidden_roles.append(role)
        if role == ROLE_TOTAL:
            reasons.append("جمع کل یا مقدار کل نمی‌تواند یکی از اجزای جمع باشد.")
        elif role == ROLE_PERCENTAGE:
            reasons.append("درصد، نرخ یا نسبت را نمی‌توان با تعداد یا مبلغ جمع کرد.")
        else:
            reasons.append("ردیف مقایسه/اختلاف را نمی‌توان به‌عنوان جزء جمع به کار برد.")
    if ROLE_SUBTOTAL in seen_roles and kind not in _SUBTOTAL_KINDS:
        forbidden_roles.append(ROLE_SUBTOTAL)
        reasons.append(
            "جمع میانی در کنار اجزای خودش قرار گرفته است؛ جمع دوباره‌شماری می‌شود."
        )

    if len(seen_units) > 1:
        mixed_units.extend(sorted(seen_units))
        reasons.append("اجزا واحد یکسان ندارند: " + "، ".join(sorted(seen_units)) + ".")
    if unit:
        wrong = sorted(item for item in seen_units if item != unit)
        if wrong:
            mixed_units.extend(wrong)
            reasons.append(
                f"واحد اجزا ({'، '.join(wrong)}) با واحد کل ({unit}) یکسان نیست."
            )

    for category, owners in sorted(category_owners.items()):
        if len(owners) > 1:
            overlap.append(category)
            reasons.append(
                f"دسته «{category}» در بیش از یک جزء آمده است؛ این طبقه‌بندی هم‌پوشان است."
            )

    if distinct_labels:
        duplicates = sorted(label for label, count in label_counts.items() if count > 1 and label)
        if duplicates:
            overlap.extend(duplicates)
            reasons.append(
                "برچسب تکراری در اجزا: " + "، ".join(duplicates) + "؛ رابطه یک‌به‌یک نیست."
            )

    missing.extend(component.label for component in components if component.value is None)
    if missing:
        reasons.append(
            "جزء بدون مقدار: " + "، ".join(missing) + "؛ مقدار گمشده صفر فرض نشد."
        )

    return AggregationEvidence(
        justificed=not forbidden_roles and not mixed_units and not overlap,
        reasons=reasons,
        overlap=sorted(set(overlap)),
        missing=missing,
        mixed_units=sorted(set(mixed_units)),
        forbidden_roles=sorted(set(forbidden_roles)),
    )


# --------------------------------------------------------------------------- #
# Reconciliation record
# --------------------------------------------------------------------------- #
@dataclass
class ReconciliationResult:
    kind: str
    status: ReconciliationStatus
    expected: Decimal | None
    calculated: Decimal | None
    difference: Decimal | None
    tolerance: Tolerance
    explanation: str
    sources: list[str] = field(default_factory=list)
    sheet: str = ""
    block: str = ""
    measure: str = ""
    label: str = ""
    unit: str = ""
    expected_source: str = ""
    components: list[dict[str, Any]] = field(default_factory=list)
    evidence: AggregationEvidence | None = None

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    @property
    def kind_label(self) -> str:
        return KIND_FA.get(self.kind, self.kind)

    @property
    def verified(self) -> bool:
        return (
            self.status is ReconciliationStatus.PASS
            and self.evidence is not None
            and self.evidence.justificed
            and self.difference is not None
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "kind_label": self.kind_label,
            "status": self.status.value,
            "status_label": self.status_label,
            "verified": self.verified,
            "expected": _text(self.expected),
            "calculated": _text(self.calculated),
            "difference": _text(self.difference),
            "tolerance": self.tolerance.to_dict(),
            "tolerance_description": self.tolerance.describe_fa(self.expected),
            "explanation": self.explanation,
            "sources": list(self.sources),
            "sheet": self.sheet,
            "block": self.block,
            "measure": self.measure,
            "label": self.label,
            "unit": self.unit,
            "expected_source": self.expected_source,
            "components": list(self.components),
            "evidence": self.evidence.to_dict() if self.evidence else None,
        }


def _sum_components(components: Sequence[AggregationComponent]) -> Decimal:
    return sum((component.value for component in components if component.value is not None), ZERO)


def reconcile_components(
    expected: Decimal | None,
    components: Sequence[AggregationComponent],
    *,
    kind: str = KIND_COMPONENT_TOTAL,
    tolerance: Tolerance | None = None,
    unit: str = "",
    label: str = "",
    measure: str = "",
    sheet: str = "",
    block: str = "",
    expected_source: str = "",
    distinct_labels: bool = False,
) -> ReconciliationResult:
    """Compare a declared total with the sum of its justified components."""

    tolerance = tolerance or Tolerance.for_workbook(components=len(components))
    evidence = assess_aggregation(
        components, kind=kind, unit=unit, distinct_labels=distinct_labels
    )
    sources = [component.source for component in components if component.source]
    if expected_source:
        sources.append(expected_source)
    payload = [component.to_dict() for component in components]

    if not evidence.justificed:
        return ReconciliationResult(
            kind=kind,
            status=ReconciliationStatus.FAIL,
            expected=expected,
            calculated=None,
            difference=None,
            tolerance=tolerance,
            explanation="رابطه جمع توجیه‌پذیر نیست؛ " + " ".join(evidence.reasons),
            sources=sources,
            sheet=sheet,
            block=block,
            measure=measure,
            label=label,
            unit=unit,
            expected_source=expected_source,
            components=payload,
            evidence=evidence,
        )

    if evidence.missing:
        return ReconciliationResult(
            kind=kind,
            status=ReconciliationStatus.WARNING,
            expected=expected,
            calculated=None,
            difference=None,
            tolerance=tolerance,
            explanation=(
                "مقدار برخی اجزا موجود نیست؛ جمع انجام نشد و مقدار گمشده صفر فرض نشد: "
                + "، ".join(evidence.missing)
            ),
            sources=sources,
            sheet=sheet,
            block=block,
            measure=measure,
            label=label,
            unit=unit,
            expected_source=expected_source,
            components=payload,
            evidence=evidence,
        )

    calculated = _sum_components(components)
    if expected is None:
        return ReconciliationResult(
            kind=kind,
            status=ReconciliationStatus.WARNING,
            expected=None,
            calculated=calculated,
            difference=None,
            tolerance=tolerance,
            explanation=(
                f"جمع اجزا {_text(calculated)} است، اما مقدار کل مورد انتظار در فایل موجود نیست؛ "
                "نتیجه تأیید نشد."
            ),
            sources=sources,
            sheet=sheet,
            block=block,
            measure=measure,
            label=label,
            unit=unit,
            expected_source=expected_source,
            components=payload,
            evidence=evidence,
        )

    difference = expected - calculated
    status = tolerance.classify(difference, expected)
    if status is ReconciliationStatus.PASS:
        explanation = (
            f"جمع اجزا {_text(calculated)} با مقدار کل {_text(expected)} مطابقت دارد؛ "
            f"اختلاف {_text(difference)} در {tolerance.describe_fa(expected)} است."
        )
    elif status is ReconciliationStatus.WARNING:
        explanation = (
            f"اختلاف {_text(difference)} بین کل {_text(expected)} و جمع {_text(calculated)} "
            f"از تحمل تأیید بیشتر ولی در {tolerance.describe_fa(expected)} است."
        )
    else:
        explanation = (
            f"اختلاف {_text(difference)} بین کل {_text(expected)} و جمع {_text(calculated)} "
            f"بیشتر از {tolerance.describe_fa(expected)} است."
        )
    return ReconciliationResult(
        kind=kind,
        status=status,
        expected=expected,
        calculated=calculated,
        difference=difference,
        tolerance=tolerance,
        explanation=explanation,
        sources=sources,
        sheet=sheet,
        block=block,
        measure=measure,
        label=label,
        unit=unit,
        expected_source=expected_source,
        components=payload,
        evidence=evidence,
    )


# --------------------------------------------------------------------------- #
# Hierarchy (group -> subgroup)
# --------------------------------------------------------------------------- #
@dataclass
class HierarchyNode:
    label: str
    value: Decimal | None
    children: list["HierarchyNode"] = field(default_factory=list)
    unit: str = ""
    source: str = ""
    categories: frozenset[str] = frozenset()


def reconcile_hierarchy(
    root: HierarchyNode,
    *,
    tolerance: Tolerance | None = None,
    kind: str = KIND_GROUP_SUBGROUP,
) -> list[ReconciliationResult]:
    """Check every parent against its *immediate* children only.

    Nested levels are checked one at a time, so a parent is never added to its
    grandchildren and a subtotal is never counted twice.
    """

    results: list[ReconciliationResult] = []

    def walk(node: HierarchyNode) -> None:
        if node.children:
            components = [
                AggregationComponent(
                    label=child.label,
                    value=child.value,
                    role=ROLE_SUBTOTAL if child.children else ROLE_DATA,
                    unit=child.unit or node.unit,
                    categories=child.categories,
                    source=child.source,
                )
                for child in node.children
            ]
            results.append(
                reconcile_components(
                    node.value,
                    components,
                    kind=kind,
                    tolerance=tolerance,
                    unit=node.unit,
                    label=node.label,
                    measure=node.label,
                    expected_source=node.source,
                    distinct_labels=True,
                )
            )
        for child in node.children:
            walk(child)

    walk(root)
    return results


# --------------------------------------------------------------------------- #
# Category and period comparisons
# --------------------------------------------------------------------------- #
@dataclass
class CategoryComparison:
    label: str
    change: ChangeResult
    status: ReconciliationStatus
    explanation: str
    unit: str = ""

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": KIND_CATEGORY_COMPARISON,
            "kind_label": KIND_FA[KIND_CATEGORY_COMPARISON],
            "label": self.label,
            "unit": self.unit,
            "status": self.status.value,
            "status_label": self.status_label,
            "change": self.change.to_dict(),
            "explanation": self.explanation,
        }


def compare_categories(
    current: Mapping[str, Decimal | None],
    previous: Mapping[str, Decimal | None],
    *,
    unit: str = "",
) -> list[CategoryComparison]:
    """Compare like-named categories across two periods.

    A category that exists on one side only is reported as a warning with no
    percentage; it is never treated as a zero baseline.
    """

    order = list(current) + [label for label in previous if label not in current]
    results: list[CategoryComparison] = []
    for label in order:
        change = compute_change(current.get(label), previous.get(label), unit=unit, label=label)
        if label not in current or label not in previous:
            status = ReconciliationStatus.WARNING
            explanation = (
                f"دسته «{label}» فقط در یک دوره وجود دارد؛ مقایسه هم‌سنگ ممکن نیست و "
                "مقدار غایب صفر فرض نشد."
            )
        else:
            status = change.status
            explanation = change.explanation
        results.append(
            CategoryComparison(
                label=label,
                change=change,
                status=status,
                explanation=explanation,
                unit=unit,
            )
        )
    return results


@dataclass
class PeriodComparisonResult:
    current_label: str
    previous_label: str
    change: ChangeResult
    status: ReconciliationStatus
    explanation: str
    months_used: list[str] = field(default_factory=list)
    months_missing: list[str] = field(default_factory=list)
    incomplete_excluded: list[str] = field(default_factory=list)
    metric: str = ""
    source: str = ""

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    @property
    def kind_label(self) -> str:
        return KIND_FA[KIND_PERIOD_COMPARISON]

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": KIND_PERIOD_COMPARISON,
            "kind_label": KIND_FA[KIND_PERIOD_COMPARISON],
            "metric": self.metric,
            "current_label": self.current_label,
            "previous_label": self.previous_label,
            "status": self.status.value,
            "status_label": self.status_label,
            "change": self.change.to_dict(),
            "months_used": list(self.months_used),
            "months_missing": list(self.months_missing),
            "incomplete_excluded": list(self.incomplete_excluded),
            "explanation": self.explanation,
            "source": self.source,
        }


def compare_periods(
    *,
    current_label: str,
    previous_label: str,
    current: Decimal | None,
    previous: Decimal | None,
    months_used: Iterable[str] = (),
    months_missing: Iterable[str] = (),
    incomplete_excluded: Iterable[str] = (),
    metric: str = "",
    unit: str = "",
    source: str = "",
) -> PeriodComparisonResult:
    """Compare two periods and downgrade the result when coverage is not equal."""

    change = compute_change(
        current,
        previous,
        unit=unit,
        label=metric,
        source=source,
    )
    notes = [change.explanation]
    status = change.status
    used = list(months_used)
    missing = list(months_missing)
    incomplete = list(incomplete_excluded)

    if current is None or previous is None:
        status = ReconciliationStatus.WARNING
    elif not used:
        status = ReconciliationStatus.WARNING
        notes.append("هیچ ماه مشترک و کاملی برای مقایسه دو دوره پیدا نشد.")
    if incomplete:
        if status is ReconciliationStatus.PASS:
            status = ReconciliationStatus.WARNING
        notes.append("ماه ناقص از مقایسه کنار گذاشته شد: " + "، ".join(incomplete) + ".")
    if missing:
        if status is ReconciliationStatus.PASS:
            status = ReconciliationStatus.WARNING
        notes.append("ماه بدون داده در یک دوره: " + "، ".join(missing) + ".")

    return PeriodComparisonResult(
        current_label=current_label,
        previous_label=previous_label,
        change=change,
        status=status,
        explanation=" ".join(note for note in notes if note),
        months_used=used,
        months_missing=missing,
        incomplete_excluded=incomplete,
        metric=metric,
        source=source,
    )


# --------------------------------------------------------------------------- #
# Workbook report
# --------------------------------------------------------------------------- #
@dataclass
class StatisticalReconciliationReport:
    reconciliations: list[ReconciliationResult] = field(default_factory=list)
    changes: list[ChangeResult] = field(default_factory=list)
    averages: list[AverageResult] = field(default_factory=list)
    comparisons: list[PeriodComparisonResult] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def status(self) -> ReconciliationStatus:
        statuses = [item.status for item in self.reconciliations]
        statuses.extend(item.status for item in self.comparisons)
        statuses.extend(item.status for item in self.averages)
        if ReconciliationStatus.FAIL in statuses:
            return ReconciliationStatus.FAIL
        if ReconciliationStatus.WARNING in statuses:
            return ReconciliationStatus.WARNING
        return ReconciliationStatus.PASS

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status.value)

    def counts(self) -> dict[str, int]:
        counts = {status.value: 0 for status in ReconciliationStatus}
        for item in self.reconciliations:
            counts[item.status.value] += 1
        for item in self.comparisons:
            counts[item.status.value] += 1
        for item in self.averages:
            counts[item.status.value] += 1
        return counts

    @property
    def failed(self) -> list[Any]:
        return [
            item
            for item in [*self.reconciliations, *self.comparisons, *self.averages]
            if item.status is ReconciliationStatus.FAIL
        ]

    @property
    def warnings_only(self) -> list[Any]:
        return [
            item
            for item in [*self.reconciliations, *self.comparisons, *self.averages]
            if item.status is ReconciliationStatus.WARNING
        ]

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status.value,
            "status_label": self.status_label,
            "counts": self.counts(),
            "reconciliations": [item.to_dict() for item in self.reconciliations],
            "changes": [item.to_dict() for item in self.changes],
            "averages": [item.to_dict() for item in self.averages],
            "comparisons": [item.to_dict() for item in self.comparisons],
            "warnings": list(self.warnings),
        }


class StatisticalReconciler:
    """Build reconciliation evidence from the extracted workbook structure."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def analyze(
        self,
        analysis: AnalysisResult,
        grids: dict[str, SheetGrid] | None = None,
    ) -> StatisticalReconciliationReport:
        grids = grids or {}
        report = StatisticalReconciliationReport()
        for table in self._tables(analysis):
            grid = grids.get(table.sheet)
            if grid is None:
                continue
            block = self._semantic_block(analysis, table)
            report.reconciliations.extend(self._block_reconciliations(grid, table, block))
            report.averages.extend(self._block_averages(grid, table, block))
        report.reconciliations.extend(self._monthly_reconciliations(analysis))
        report.comparisons.extend(self._period_comparisons(analysis))
        report.changes.extend(item.change for item in report.comparisons)
        if not (report.reconciliations or report.comparisons or report.averages):
            report.warnings.append(
                "رابطه جمع قابل توجیهی برای تطبیق آماری در این فایل پیدا نشد."
            )
        return report

    # ------------------------------------------------------------------ #
    @staticmethod
    def _tables(analysis: AnalysisResult) -> list[TableCandidate]:
        result: list[TableCandidate] = []
        seen: set[int] = set()
        for table in getattr(analysis, "tables", []) or []:
            if is_report_sheet(table.sheet):
                continue
            if table.row_count < 1 or len(table.columns) < 1:
                continue
            if id(table) in seen:
                continue
            seen.add(id(table))
            result.append(table)
        # The detector can report a narrower view of the same block (for
        # example a two-year table plus a single-year slice of it).  Reconciling
        # both would count the same cells twice, so a strictly contained block
        # is dropped from reconciliation only.
        return [
            table
            for table in result
            if not any(
                other is not table and _strictly_contains(other, table)
                for other in result
            )
        ]

    @staticmethod
    def _semantic_block(analysis: AnalysisResult, table: TableCandidate) -> SemanticBlock | None:
        semantics = getattr(analysis, "semantics", None)
        if semantics is not None:
            for block in getattr(semantics, "blocks", []):
                if block.sheet == table.sheet and block.address == table.address:
                    return block
        return None

    @staticmethod
    def _fields(block: SemanticBlock | None, table: TableCandidate) -> dict[int, SemanticField]:
        if block is not None:
            return {field.column: field for field in block.columns}
        fields: dict[int, SemanticField] = {}
        for profile in table.columns:
            role = DATA_ROLE_TO_SEMANTIC.get(profile.statistical_role, SemanticRole.UNKNOWN)
            fields[profile.index] = SemanticField(
                column=profile.index,
                letter=profile.letter,
                header=profile.header,
                role=role,
                statistical_role=profile.statistical_role.value,
                business_role=profile.role.value,
            )
        return fields

    @staticmethod
    def _number(grid: SheetGrid, row: int, column: int) -> tuple[Decimal | None, bool]:
        cell = grid.get(row, column)
        if cell is None or (cell.is_empty and not cell.formula):
            return None, False
        parsed = parse_value(cell.value, kind_hint=ValueKind.NUMBER)
        if isinstance(parsed.value, Decimal):
            return parsed.value, True
        return None, True

    # ------------------------------------------------------------------ #
    def _block_reconciliations(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        block: SemanticBlock | None,
    ) -> list[ReconciliationResult]:
        fields = self._fields(block, table)
        headers = list(table.header_rows or [table.header_row])
        records = classify_block_rows(
            grid, table, headers, block_end_row=table.block_end
        )
        roles = {record.row: record.role for record in records}
        labels = {record.row: record.label for record in records}

        data_roles = {ROW_DATA, ROW_DATA_MISSING, ROW_CATEGORY}
        data_rows = [
            row
            for row in range(table.first_data_row, table.last_data_row + 1)
            if roles.get(row) in data_roles
        ]
        subtotal_rows = sorted(
            {
                *[row for row in table.subtotal_rows if row <= table.block_end],
                *[row for row in roles if roles[row] == ROW_SUBTOTAL],
            }
        )
        grand_rows = sorted(
            {
                *[row for row in table.grand_total_rows if row <= table.block_end],
                *[row for row in roles if roles[row] == ROW_TOTAL],
            }
            - set(subtotal_rows)
        )

        results: list[ReconciliationResult] = []
        for column, field in sorted(fields.items()):
            if field.role not in MEASURE_ROLES:
                continue
            if not any(profile.index == column for profile in table.columns):
                continue
            unit = field.unit
            tolerance = Tolerance.for_workbook(
                components=max(1, len(data_rows))
            )
            data_components = self._components(
                grid, table, data_rows, column, labels, field, role=ROLE_DATA
            )

            section_start = 0
            for subtotal_row in subtotal_rows:
                section = [
                    component
                    for component in data_components
                    if section_start < _row_from_source(component.source) < subtotal_row
                ]
                expected, present = self._number(grid, subtotal_row, column)
                expected_source = f"{grid.sheet_name}!{field.letter}{subtotal_row}"
                label = labels.get(subtotal_row) or "جمع میانی"
                headings = [
                    row
                    for row in range(section_start + 1, subtotal_row)
                    if roles.get(row) == ROW_GROUP_HEADING
                ]
                if headings:
                    kind = KIND_GROUP_SUBGROUP
                    group_label = labels.get(headings[-1]) or ""
                    if group_label:
                        label = f"{group_label} — {label}"
                else:
                    kind = KIND_COMPONENT_TOTAL
                if not present and not section:
                    continue
                if not present:
                    results.append(
                        ReconciliationResult(
                            kind=kind,
                            status=ReconciliationStatus.WARNING,
                            expected=None,
                            calculated=None,
                            difference=None,
                            tolerance=tolerance,
                            explanation=(
                                f"ردیف «{label}» در ستون «{field.header or field.letter}» "
                                "مقدار جمع ندارد؛ جمع اجزا تأیید نشد."
                            ),
                            sources=[component.source for component in section],
                            sheet=grid.sheet_name,
                            block=table.address,
                            measure=field.header or field.letter,
                            label=label,
                            unit=unit,
                            expected_source=expected_source,
                            components=[component.to_dict() for component in section],
                        )
                    )
                    section_start = subtotal_row
                    continue
                results.append(
                    reconcile_components(
                        expected,
                        section,
                        kind=kind,
                        tolerance=tolerance,
                        unit=unit,
                        label=label,
                        measure=field.header or field.letter,
                        sheet=grid.sheet_name,
                        block=table.address,
                        expected_source=expected_source,
                    )
                )
                section_start = subtotal_row

            for grand_row in grand_rows:
                prior_subtotals = [row for row in subtotal_rows if row < grand_row]
                if prior_subtotals:
                    components = self._components(
                        grid, table, prior_subtotals, column, labels, field, role=ROLE_SUBTOTAL
                    )
                    kind = KIND_SUBTOTAL_GRAND_TOTAL
                else:
                    components = data_components
                    kind = KIND_COMPONENT_TOTAL
                expected, present = self._number(grid, grand_row, column)
                expected_source = f"{grid.sheet_name}!{field.letter}{grand_row}"
                label = labels.get(grand_row) or "جمع کل"
                if not present:
                    results.append(
                        ReconciliationResult(
                            kind=kind,
                            status=ReconciliationStatus.WARNING,
                            expected=None,
                            calculated=None,
                            difference=None,
                            tolerance=tolerance,
                            explanation=(
                                f"ردیف «{label}» در ستون «{field.header or field.letter}» "
                                "مقدار جمع ندارد؛ جمع اجزا تأیید نشد."
                            ),
                            sources=[component.source for component in components],
                            sheet=grid.sheet_name,
                            block=table.address,
                            measure=field.header or field.letter,
                            label=label,
                            unit=unit,
                            expected_source=expected_source,
                            components=[component.to_dict() for component in components],
                        )
                    )
                    continue
                results.append(
                    reconcile_components(
                        expected,
                        components,
                        kind=kind,
                        tolerance=tolerance,
                        unit=unit,
                        label=label,
                        measure=field.header or field.letter,
                        sheet=grid.sheet_name,
                        block=table.address,
                        expected_source=expected_source,
                    )
                )
        return results

    def _components(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        rows: Sequence[int],
        column: int,
        labels: Mapping[int, str],
        field: SemanticField,
        *,
        role: str,
    ) -> list[AggregationComponent]:
        components: list[AggregationComponent] = []
        for row in rows:
            value, _present = self._number(grid, row, column)
            label = labels.get(row) or f"ردیف {row}"
            components.append(
                AggregationComponent(
                    label=label,
                    value=value,
                    role=role,
                    unit=self._row_unit(grid, table, column, row, field.unit),
                    source=f"{grid.sheet_name}!{field.letter}{row}",
                )
            )
        return components

    @staticmethod
    def _row_unit(
        grid: SheetGrid,
        table: TableCandidate,
        column: int,
        row: int,
        fallback: str,
    ) -> str:
        """Read a per-row unit from an adjacent «واحد» column when present."""

        for neighbour in (column + 1, column - 1):
            profile = next(
                (item for item in table.columns if item.index == neighbour), None
            )
            if profile is None:
                continue
            header = _normalized(profile.header)
            if "واحد" not in header and "unit" not in header:
                continue
            cell = grid.get(row, neighbour)
            if cell is None or cell.is_empty:
                continue
            text = normalize_text(normalize_digits(str(cell.value))).strip()
            if text:
                return text[:40]
        return fallback

    # ------------------------------------------------------------------ #
    def _block_averages(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        block: SemanticBlock | None,
    ) -> list[AverageResult]:
        fields = self._fields(block, table)
        ratio_fields = [
            field
            for field in fields.values()
            if field.role in MEASURE_ROLES and _is_ratio_field(field)
        ]
        weight_fields = [
            field
            for field in fields.values()
            if field.role in MEASURE_ROLES and _is_weight_field(field)
        ]
        if not ratio_fields or not weight_fields:
            return []

        rows = [
            row
            for row in range(table.first_data_row, table.last_data_row + 1)
        ]
        results: list[AverageResult] = []
        for ratio in ratio_fields:
            weight = next(
                (field for field in weight_fields if field.column != ratio.column),
                None,
            )
            if weight is None:
                continue
            values: list[Decimal] = []
            weights: list[Decimal | None] = []
            skipped = 0
            for row in rows:
                value, value_present = self._number(grid, row, ratio.column)
                mass, mass_present = self._number(grid, row, weight.column)
                if not value_present:
                    skipped += 1
                    continue
                values.append(value if value is not None else ZERO)
                weights.append(mass if mass_present else None)
            if not values:
                continue
            result = compute_average(
                values,
                weights,
                weights_required=True,
                label=ratio.header or ratio.letter,
                unit=ratio.unit,
            )
            note = (
                f" وزن از ستون «{weight.header or weight.letter}» گرفته شد؛ "
                "میانگین ساده برای این نسبت گزارش نمی‌شود."
            )
            if skipped:
                note += f" {skipped} ردیف بدون مقدار در محاسبه نیامد."
            result.explanation += note
            results.append(result)
        return results

    # ------------------------------------------------------------------ #
    def _monthly_reconciliations(self, analysis: AnalysisResult) -> list[ReconciliationResult]:
        integration = getattr(analysis, "integration", None)
        if integration is None or not getattr(integration, "is_usable", False):
            return []
        results: list[ReconciliationResult] = []
        for series in integration.series[:12]:
            for period in series.periods():
                cells = [
                    cell
                    for cell in series.cells
                    if cell.period == period and cell.complete and cell.total is not None
                ]
                if not cells:
                    continue
                if any(cell.declared_total is None for cell in cells):
                    continue
                expected = series.total(period)
                components = [
                    AggregationComponent(
                        label=cell.month,
                        value=cell.declared_total,
                        role=ROLE_DATA,
                        source=(
                            f"{cell.source}!{cell.total_address}"
                            if cell.total_address
                            else cell.source
                        ),
                    )
                    for cell in cells
                ]
                result = reconcile_components(
                    expected,
                    components,
                    kind=KIND_MONTHLY_PERIOD,
                    tolerance=Tolerance.for_workbook(components=len(components)),
                    label=f"{series.metric} — {period}",
                    measure=series.metric,
                    sheet=cells[0].source,
                    expected_source=f"جمع ماهانه {period}",
                )
                result.explanation = (
                    f"دوره {period}: " + result.explanation
                )
                results.append(result)
        return results

    def _period_comparisons(self, analysis: AnalysisResult) -> list[PeriodComparisonResult]:
        integration = getattr(analysis, "integration", None)
        if integration is None or not getattr(integration, "is_usable", False):
            return []
        results: list[PeriodComparisonResult] = []
        for series in integration.series[:12]:
            periods = series.periods()
            for index in range(len(periods) - 1):
                previous_label = periods[index]
                current_label = periods[index + 1]
                change = series.compare(previous_label, current_label)
                result = compare_periods(
                    current_label=current_label,
                    previous_label=previous_label,
                    current=change.total_b,
                    previous=change.total_a,
                    months_used=change.months_used,
                    months_missing=change.months_missing,
                    incomplete_excluded=change.incomplete_excluded,
                    metric=series.metric,
                    source="؛ ".join(change.months_used),
                )
                if change.total_a is None or change.total_b is None:
                    result.explanation = change.note
                results.append(result)
        return results


def _strictly_contains(outer: TableCandidate, inner: TableCandidate) -> bool:
    if outer.sheet != inner.sheet:
        return False
    if outer.first_col > inner.first_col or outer.last_col < inner.last_col:
        return False
    if outer.header_row > inner.header_row or outer.block_end < inner.block_end:
        return False
    return (
        outer.first_col < inner.first_col
        or outer.last_col > inner.last_col
        or outer.header_row < inner.header_row
        or outer.block_end > inner.block_end
    )


def _row_from_source(source: str) -> int:
    match = re.search(r"!([A-Z]+)(\d+)$", source)
    return int(match.group(2)) if match else 0


def _is_ratio_field(field: SemanticField) -> bool:
    unit = _normalized(field.unit)
    header = _normalized(field.header)
    if unit in {"%", "percent", "percentage", "درصد"}:
        return True
    return any(word in header for word in _RATIO_WORDS)


def _is_weight_field(field: SemanticField) -> bool:
    header = _normalized(field.header)
    return any(word in header for word in _WEIGHT_WORDS)
