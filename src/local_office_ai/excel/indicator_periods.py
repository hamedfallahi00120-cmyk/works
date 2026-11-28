"""Compare one indicator across reporting periods inside a row-oriented block.

Administrative workbooks often place indicators in rows and one column per
reporting year (``شاخص | 1404 | 1405``).  The period matrix
(:mod:`period_integration`) reads the opposite shape: metrics in columns and a
month per worksheet, so it cannot compare two year columns without summing
unrelated indicators.

This module adds the missing, deterministic view.  Every observation is one
cell; an indicator comparison pairs the same row label and the same measure
label across periods, and never adds one indicator to another.  Summary,
percentage, difference and note rows are excluded by the shared row classifier.
A missing value stays ``None``; a zero baseline yields ``percent=None``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from typing import Any, Iterable

from local_office_ai.core.models import DataRole, ValueKind
from local_office_ai.excel.normalization import (
    normalize_digits,
    normalize_text,
    parse_value,
)
from local_office_ai.excel.period_alignment import NormalizedPeriod, parse_period
from local_office_ai.excel.row_classification import (
    ROW_CATEGORY,
    ROW_DATA,
    ROW_DATA_MISSING,
    ROW_TOTAL,
    SUMMARY_ROLES,
)

ZERO = Decimal("0")
HUNDRED = Decimal("100")

STATUS_COMPARED = "compared"
STATUS_PREVIOUS_MISSING = "previous_missing"
STATUS_CURRENT_MISSING = "current_missing"
STATUS_BOTH_MISSING = "both_missing"
STATUS_ZERO_BASELINE = "zero_baseline"

STATUS_FA = {
    STATUS_COMPARED: "مقایسه شد",
    STATUS_PREVIOUS_MISSING: "مقدار دوره قبل موجود نیست",
    STATUS_CURRENT_MISSING: "مقدار دوره جاری موجود نیست",
    STATUS_BOTH_MISSING: "هیچ یک از دو دوره مقدار ندارد",
    STATUS_ZERO_BASELINE: "مبنا صفر است؛ درصد تعریف نمی‌شود",
}

_ALLOWED_ROLES = {ROW_DATA, ROW_DATA_MISSING, ROW_CATEGORY}
_STRUCTURAL_ROLES = {
    "identifier",
    "temporal_key",
    "group",
    "text",
}
_PERCENT_WORDS = ("درصد", "percent", "ratio", "سهم", "نرخ")
_DIFFERENCE_WORDS = ("اختلاف", "تفاضل", "difference", "تغییر", "change")
_PERIOD_TOKEN = re.compile(r"(?<!\d)(?:1[34]\d{2}|19\d{2}|20\d{2})(?!\d)")


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return format(value, "f")
    return str(value)


def _decimal(value: Any) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    parsed = parse_value(value)
    if parsed.kind is ValueKind.NUMBER and parsed.value is not None:
        number = parsed.value
        if isinstance(number, Decimal):
            return number
        try:
            return Decimal(str(number))
        except (InvalidOperation, ValueError):
            return None
    return None


def period_label(period: NormalizedPeriod | None) -> str:
    if period is None:
        return ""
    if period.granularity == "year" and period.year:
        return str(period.year)
    if period.granularity == "month" and period.year and period.month:
        return f"{period.year}/{period.month:02d}"
    if period.granularity == "day" and period.start:
        return period.start.isoformat()
    return (period.label or "").strip()


def _period_sort_key(label: str) -> tuple[int, str]:
    match = re.match(r"^(\d{4})(?:/(\d{1,2}))?(?:/(\d{1,2}))?$", label)
    if not match:
        return (9999, label)
    return (int(match.group(1)), label)


def _base_label(header: str, period: NormalizedPeriod | None) -> str:
    text = normalize_text(normalize_digits(header or "")).strip()
    if period is not None and period.year:
        text = _PERIOD_TOKEN.sub(" ", text)
    text = re.sub(r"\s+", " ", text).strip(" -_/،:؛")
    return text


def _is_percent_or_difference(header: str, role: str) -> bool:
    text = normalize_text(normalize_digits(header or "")).lower()
    if any(word in text for word in _PERCENT_WORDS):
        return True
    if role == DataRole.CALCULATED.value and any(
        word in text for word in _DIFFERENCE_WORDS
    ):
        return True
    return False


@dataclass
class IndicatorObservation:
    indicator: str
    measure: str
    period: str
    value: Decimal | None
    sheet: str
    block: str
    address: str
    source: str = ""
    unit: str = ""
    row_label: str = ""
    role: str = "data"

    def to_dict(self) -> dict[str, Any]:
        return {
            "indicator": self.indicator,
            "measure": self.measure,
            "row_label": self.row_label,
            "period": self.period,
            "value": _text(self.value),
            "sheet": self.sheet,
            "block": self.block,
            "address": self.address,
            "source": self.source,
            "unit": self.unit,
            "role": self.role,
        }


@dataclass
class IndicatorComparison:
    indicator: str
    measure: str
    previous_period: str
    current_period: str
    previous: Decimal | None
    current: Decimal | None
    previous_address: str = ""
    current_address: str = ""
    sheet: str = ""
    block: str = ""
    unit: str = ""
    difference: Decimal | None = None
    percent: Decimal | None = None
    status: str = STATUS_COMPARED
    note: str = ""
    source: str = ""
    role: str = "data"

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status)

    @property
    def direction(self) -> str:
        if self.difference is None:
            return "قابل مقایسه نیست"
        if self.difference > 0:
            return "افزایش"
        if self.difference < 0:
            return "کاهش"
        return "بدون تغییر"

    def to_dict(self) -> dict[str, Any]:
        return {
            "indicator": self.indicator,
            "measure": self.measure,
            "row_label": self.row_label,
            "previous_period": self.previous_period,
            "current_period": self.current_period,
            "previous": _text(self.previous),
            "current": _text(self.current),
            "difference": _text(self.difference),
            "percent": _text(self.percent),
            "direction": self.direction,
            "status": self.status,
            "status_label": self.status_label,
            "sheet": self.sheet,
            "block": self.block,
            "unit": self.unit,
            "previous_address": self.previous_address,
            "current_address": self.current_address,
            "note": self.note,
            "source": self.source,
            "role": self.role,
        }


@dataclass
class IndicatorPeriodReport:
    source: str = ""
    observations: list[IndicatorObservation] = field(default_factory=list)
    comparisons: list[IndicatorComparison] = field(default_factory=list)
    periods: list[str] = field(default_factory=list)
    excluded_rows: list[dict[str, Any]] = field(default_factory=list)
    blocked_blocks: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def comparison(self, label: str) -> IndicatorComparison | None:
        """Find one indicator by normalized label, exact first then near match."""

        key = normalize_text(normalize_digits(label))
        exact = [
            item
            for item in self.comparisons
            if normalize_text(normalize_digits(item.indicator)) == key
        ]
        if len(exact) == 1:
            return exact[0]
        best: tuple[float, IndicatorComparison] | None = None
        for item in self.comparisons:
            candidate = normalize_text(normalize_digits(item.indicator))
            score = SequenceMatcher(None, key, candidate).ratio()
            if best is None or score > best[0]:
                best = (score, item)
        if best is not None and best[0] >= 0.85:
            return best[1]
        return None

    def values_for(self, label: str, period: str) -> IndicatorObservation | None:
        key = normalize_text(normalize_digits(label))
        for item in self.observations:
            if (
                normalize_text(normalize_digits(item.indicator)) == key
                and item.period == period
            ):
                return item
        return None

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "periods": list(self.periods),
            "observations": [item.to_dict() for item in self.observations],
            "comparisons": [item.to_dict() for item in self.comparisons],
            "excluded_rows": list(self.excluded_rows),
            "blocked_blocks": list(self.blocked_blocks),
            "warnings": list(self.warnings),
        }


def _cells_by_address(block: Any) -> dict[tuple[int, int], Any]:
    return {(cell.row, cell.column): cell for cell in getattr(block, "cells", [])}


def _row_roles(block: Any) -> dict[int, str]:
    roles: dict[int, str] = {}
    for item in getattr(block, "row_roles", []) or []:
        if isinstance(item, dict) and item.get("row") is not None:
            roles[int(item["row"])] = str(item.get("role") or "")
    return roles


def _row_labels(block: Any, label_columns: list[Any]) -> dict[int, str]:
    labels: dict[int, str] = {}
    cells = _cells_by_address(block)
    for column in label_columns:
        for (row, column_index), cell in cells.items():
            if column_index != column.index or cell.value in (None, ""):
                continue
            text = normalize_text(normalize_digits(str(cell.value))).strip()
            if text:
                labels.setdefault(row, text)
    return labels


def build_indicator_periods(
    analysis: Any,
    *,
    source: str = "",
    blocked_role: str = "",
) -> IndicatorPeriodReport:
    """Compare row-oriented indicators across the periods of their columns."""

    report = IndicatorPeriodReport(source=source)
    extraction = getattr(analysis, "extraction", None)
    if extraction is None:
        report.warnings.append("مدل استخراج در دسترس نیست؛ مقایسه شاخص ممکن نشد.")
        return report

    for sheet in getattr(extraction, "sheets", []) or []:
        for block in getattr(sheet, "blocks", []) or []:
            _collect_block(block, report)

    report.periods = sorted(
        {item.period for item in report.observations}, key=_period_sort_key
    )
    report.comparisons = _comparisons(report.observations)
    if not report.observations and not report.blocked_blocks:
        report.warnings.append(
            "هیچ بلوکی با شاخص‌های ردیفی و ستون‌های دوره‌ای برای مقایسه پیدا نشد."
        )
    return report


def _collect_block(block: Any, report: IndicatorPeriodReport) -> None:
    columns = list(getattr(block, "columns", []) or [])
    period_columns: list[tuple[Any, NormalizedPeriod]] = []
    for column in columns:
        if column.semantic_role in _STRUCTURAL_ROLES:
            continue
        header = str(getattr(column, "header", "") or "")
        if _is_percent_or_difference(header, str(column.semantic_role)):
            continue
        period = parse_period(column.period or header)
        if period is None or period.year is None:
            continue
        period_columns.append((column, period))

    if len(period_columns) < 2:
        return

    period_names = {period_label(period) for _, period in period_columns}
    if len(period_names) < 2:
        return
    bases = {
        _base_label(str(column.header or ""), period)
        for column, period in period_columns
    }
    if len(bases) == 1 and "" not in bases:
        # One metric repeated per period is a monthly series, not row indicators.
        return

    label_columns = [
        column
        for column in columns
        if column.semantic_role in {DataRole.GROUP.value, DataRole.CATEGORICAL.value}
    ]
    if not label_columns:
        report.blocked_blocks.append(
            {
                "sheet": getattr(block, "sheet", ""),
                "block": getattr(block, "block_id", ""),
                "reason": "ستون برچسب ردیف شناسایی نشد؛ مقدار هر ردیف به شاخص نسبت داده نشد.",
            }
        )
        return

    boundaries = getattr(block, "boundaries", {}) or {}
    first_row = int(boundaries.get("first_data_row") or 0)
    last_row = int(
        boundaries.get("block_end_row") or boundaries.get("last_data_row") or 0
    )
    if not first_row or not last_row:
        report.blocked_blocks.append(
            {
                "sheet": getattr(block, "sheet", ""),
                "block": getattr(block, "block_id", ""),
                "reason": "محدوده ردیف‌های داده مشخص نیست.",
            }
        )
        return

    labels = _row_labels(block, label_columns)
    roles = _row_roles(block)
    cells = _cells_by_address(block)
    measures = {column.index: _base_label(str(column.header or ""), period) for column, period in period_columns}
    period_by_column = {column.index: period_label(period) for column, period in period_columns}

    observations = 0
    for row in range(first_row, last_row + 1):
        role = roles.get(row, ROW_DATA)
        is_total = role in SUMMARY_ROLES
        if role not in _ALLOWED_ROLES and not is_total:
            report.excluded_rows.append(
                {
                    "sheet": getattr(block, "sheet", ""),
                    "block": getattr(block, "block_id", ""),
                    "row": row,
                    "role": role,
                    "label": labels.get(row, ""),
                    "reason": "ردیف درصد/توضیح در مقایسه شاخص وارد نمی‌شود.",
                }
            )
            continue
        if is_total and role != ROW_TOTAL:
            report.excluded_rows.append(
                {
                    "sheet": getattr(block, "sheet", ""),
                    "block": getattr(block, "block_id", ""),
                    "row": row,
                    "role": role,
                    "label": labels.get(row, ""),
                    "reason": "ردیف جمع میانی جدا از اجزا نگه داشته می‌شود و به‌عنوان جمع ثبت‌شده مقایسه می‌شود.",
                }
            )
        label = labels.get(row, "")
        if not label:
            continue
        for column_index, period in period_by_column.items():
            cell = cells.get((row, column_index))
            if cell is None:
                continue
            value = _decimal(cell.value)
            measure = measures.get(column_index, "")
            indicator = label if not measure or measure == label else f"{label} - {measure}"
            report.observations.append(
                IndicatorObservation(
                    indicator=indicator,
                    measure=measure,
                    row_label=label,
                    period=period,
                    value=value,
                    sheet=getattr(block, "sheet", ""),
                    block=getattr(block, "block_id", ""),
                    address=str(getattr(cell, "address", "")),
                    source=report.source,
                    role="total" if is_total else "data",
                )
            )
            observations += 1



def _comparisons(observations: list[IndicatorObservation]) -> list[IndicatorComparison]:
    grouped: dict[tuple[str, str], list[IndicatorObservation]] = {}
    for item in observations:
        grouped.setdefault((item.indicator, item.unit), []).append(item)

    comparisons: list[IndicatorComparison] = []
    for (indicator, unit), items in grouped.items():
        periods = sorted({item.period for item in items}, key=_period_sort_key)
        if len(periods) < 2:
            continue
        previous_period, current_period = periods[0], periods[-1]
        previous = _first(items, previous_period)
        current = _first(items, current_period)
        comparison = IndicatorComparison(
            indicator=indicator,
            measure=previous.measure if previous else (current.measure if current else ""),
            previous_period=previous_period,
            current_period=current_period,
            previous=previous.value if previous else None,
            current=current.value if current else None,
            previous_address=previous.address if previous else "",
            current_address=current.address if current else "",
            sheet=previous.sheet if previous else (current.sheet if current else ""),
            block=previous.block if previous else (current.block if current else ""),
            unit=unit,
            source=(previous.source if previous else (current.source if current else "")),
            role=(
                "total"
                if any(item.role == "total" for item in items)
                else "data"
            ),
        )
        if previous is None or previous.value is None:
            comparison.status = (
                STATUS_BOTH_MISSING
                if current is None or current.value is None
                else STATUS_PREVIOUS_MISSING
            )
        elif current is None or current.value is None:
            comparison.status = STATUS_CURRENT_MISSING
        else:
            comparison.difference = current.value - previous.value
            if previous.value == ZERO:
                comparison.status = STATUS_ZERO_BASELINE
                comparison.note = "مبنا صفر است؛ درصد تغییر محاسبه نمی‌شود."
            else:
                comparison.percent = (
                    comparison.difference / abs(previous.value) * HUNDRED
                ).quantize(Decimal("0.01"))
                comparison.status = STATUS_COMPARED
        comparisons.append(comparison)
    comparisons.sort(key=lambda item: (item.sheet, item.block, item.indicator))
    return comparisons


def _first(items: list[IndicatorObservation], period: str) -> IndicatorObservation | None:
    for item in items:
        if item.period == period:
            return item
    return None


def combine_indicator_periods(
    reports: Iterable[IndicatorPeriodReport],
) -> IndicatorPeriodReport:
    combined = IndicatorPeriodReport(source="folder")
    for report in reports:
        combined.observations.extend(report.observations)
        combined.excluded_rows.extend(report.excluded_rows)
        combined.blocked_blocks.extend(report.blocked_blocks)
        combined.warnings.extend(
            warning for warning in report.warnings if warning not in combined.warnings
        )
    combined.periods = sorted(
        {item.period for item in combined.observations}, key=_period_sort_key
    )
    combined.comparisons = _comparisons(combined.observations)
    return combined
