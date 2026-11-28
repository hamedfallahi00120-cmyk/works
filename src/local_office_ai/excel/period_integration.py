"""Bring the per-sheet (or per-file) periods into one comparable matrix.

A workbook that repeats the same table once per month is the normal shape of an
administrative report, yet every sheet on its own only answers "what happened
in this month".  This module merges those sheets into a single matrix — metric
x month x period — and recomputes every cell from the raw cells instead of
trusting the declared totals.

Rules kept from the rest of the engine:

* ``Decimal`` only; no float arithmetic.
* A cell that was not measured stays ``None``; it is never coerced to zero.
* An incomplete month is kept in the matrix but excluded from totals, means
  and rankings, and is named in ``excluded`` so the report can say why.
* Comparing two periods only uses months that exist in both, and lists the
  months that were left out.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Iterable

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import ColumnRole, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import parse_value
from local_office_ai.excel.sheet_analyzer import JALALI_MONTHS, SheetGrid

LOGGER = get_logger(__name__)

ZERO = Decimal("0")
HUNDRED = Decimal("100")

MONTH_INDEX = {name: index for index, name in enumerate(JALALI_MONTHS, start=1)}

VERIFIED = "verified"
MISMATCH = "mismatch"
NOT_CHECKED = "not_checked"

_REF_RE = re.compile(
    r"(?:'([^']+)'|([^\s'!()=+*/\-]+))!\$?([A-Z]{1,3})\$?(\d+)"
)


@dataclass
class PeriodTotal:
    metric: str
    period: str
    month: str
    month_index: int | None
    source: str
    total: Decimal | None
    observations: int = 0
    complete: bool = True
    total_address: str = ""
    declared_total: Decimal | None = None
    verification: str = NOT_CHECKED
    note: str = ""

    @property
    def label(self) -> str:
        return f"{self.month} {self.period}".strip()

    def to_dict(self) -> dict[str, Any]:
        return {
            "metric": self.metric,
            "period": self.period,
            "month": self.month,
            "month_index": self.month_index,
            "source": self.source,
            "total": _text(self.total),
            "observations": self.observations,
            "complete": self.complete,
            "total_address": self.total_address,
            "declared_total": _text(self.declared_total),
            "verification": self.verification,
            "note": self.note,
        }


@dataclass
class ChangeResult:
    period_a: str
    period_b: str
    total_a: Decimal | None
    total_b: Decimal | None
    difference: Decimal | None = None
    percent: Decimal | None = None
    months_used: list[str] = field(default_factory=list)
    months_missing: list[str] = field(default_factory=list)
    incomplete_excluded: list[str] = field(default_factory=list)
    note: str = ""

    @property
    def direction(self) -> str:
        if self.difference is None:
            return "قابل محاسبه نیست"
        if self.difference > 0:
            return "افزایش"
        if self.difference < 0:
            return "کاهش"
        return "بدون تغییر"

    def to_dict(self) -> dict[str, Any]:
        return {
            "period_a": self.period_a,
            "period_b": self.period_b,
            "total_a": _text(self.total_a),
            "total_b": _text(self.total_b),
            "difference": _text(self.difference),
            "percent": _text(self.percent),
            "direction": self.direction,
            "months_used": list(self.months_used),
            "months_missing": list(self.months_missing),
            "incomplete_excluded": list(self.incomplete_excluded),
            "note": self.note,
        }


@dataclass
class MetricSeries:
    metric: str
    cells: list[PeriodTotal] = field(default_factory=list)

    # -- lookups ------------------------------------------------------- #
    def periods(self) -> list[str]:
        return _ordered_periods({cell.period for cell in self.cells})

    def months(self, period: str | None = None) -> list[str]:
        cells = [cell for cell in self.cells if period is None or cell.period == period]
        known = {cell.month for cell in cells}
        return [name for name in JALALI_MONTHS if name in known]

    def cell(self, period: str, month: str) -> PeriodTotal | None:
        for cell in self.cells:
            if cell.period == period and cell.month == month:
                return cell
        return None

    # -- aggregates ---------------------------------------------------- #
    def total(self, period: str, *, include_incomplete: bool = False) -> Decimal | None:
        values = [
            cell.total
            for cell in self.cells
            if cell.period == period
            and cell.total is not None
            and (include_incomplete or cell.complete)
        ]
        return sum(values, ZERO) if values else None

    def observations(self, period: str, *, include_incomplete: bool = False) -> int:
        return sum(
            cell.observations
            for cell in self.cells
            if cell.period == period and (include_incomplete or cell.complete)
        )

    def mean(self, period: str) -> Decimal | None:
        total = self.total(period)
        count = sum(
            1 for cell in self.cells if cell.period == period and cell.complete and cell.total is not None
        )
        if total is None or count == 0:
            return None
        return total / Decimal(count)

    def months_used(self, period: str) -> list[str]:
        return [
            cell.month
            for cell in self.cells
            if cell.period == period and cell.complete and cell.total is not None
        ]

    def excluded_incomplete(self, period: str) -> list[str]:
        return [
            cell.month for cell in self.cells if cell.period == period and not cell.complete
        ]

    def best(self, period: str) -> PeriodTotal | None:
        candidates = [
            cell for cell in self.cells if cell.period == period and cell.complete and cell.total is not None
        ]
        return max(candidates, key=lambda cell: cell.total or ZERO) if candidates else None

    def worst(self, period: str) -> PeriodTotal | None:
        candidates = [
            cell for cell in self.cells if cell.period == period and cell.complete and cell.total is not None
        ]
        return min(candidates, key=lambda cell: cell.total or ZERO) if candidates else None

    def month_changes(self, period: str) -> list[tuple[str, Decimal, Decimal | None, Decimal | None]]:
        """Month-over-month change inside one period, complete months only."""

        ordered = [
            cell
            for month in self.months(period)
            for cell in [self.cell(period, month)]
            if cell is not None and cell.complete and cell.total is not None
        ]
        result: list[tuple[str, Decimal, Decimal | None, Decimal | None]] = []
        previous: Decimal | None = None
        for cell in ordered:
            assert cell.total is not None
            if previous is None:
                result.append((cell.month, cell.total, None, None))
            else:
                difference = cell.total - previous
                percent = (
                    (difference / abs(previous) * HUNDRED).quantize(Decimal("0.01"))
                    if previous != ZERO
                    else None
                )
                result.append((cell.month, cell.total, difference, percent))
            previous = cell.total
        return result

    def compare(self, period_a: str, period_b: str) -> ChangeResult:
        cells_a = {cell.month: cell for cell in self.cells if cell.period == period_a}
        cells_b = {cell.month: cell for cell in self.cells if cell.period == period_b}
        months = [
            month
            for month in JALALI_MONTHS
            if month in cells_a and month in cells_b
        ]
        complete: list[str] = []
        incomplete: list[str] = []
        for month in months:
            left, right = cells_a[month], cells_b[month]
            if left.total is None or right.total is None:
                incomplete.append(month)
                continue
            if left.complete and right.complete:
                complete.append(month)
                continue
            # A partial month on both sides with the same number of observations
            # is still a like-for-like comparison; it is flagged, not dropped.
            if not left.complete and not right.complete and left.observations == right.observations:
                complete.append(month)
                continue
            incomplete.append(month)
        missing = [month for month in cells_a if month not in cells_b] + [
            month for month in cells_b if month not in cells_a
        ]
        if not complete:
            return ChangeResult(
                period_a=period_a,
                period_b=period_b,
                total_a=None,
                total_b=None,
                months_used=[],
                months_missing=_ordered_months(missing),
                incomplete_excluded=_ordered_months(incomplete),
                note="هیچ ماه مشترک و کاملی برای مقایسه دو دوره وجود ندارد.",
            )
        total_a = sum((cells_a[month].total for month in complete), ZERO)
        total_b = sum((cells_b[month].total for month in complete), ZERO)
        difference = total_b - total_a
        percent = (
            (difference / abs(total_a) * HUNDRED).quantize(Decimal("0.01"))
            if total_a != ZERO
            else None
        )
        partial = [
            month
            for month in complete
            if not cells_a[month].complete or not cells_b[month].complete
        ]
        note_parts = ["مقایسه فقط روی ماه‌های مشترک و هم‌پوشش انجام شد."]
        if partial:
            note_parts.append(
                "ماه ناقص با پوشش برابر در دو دوره لحاظ شد: " + "، ".join(partial) + "."
            )
        if incomplete:
            note_parts.append(
                "ماه ناقص کنار گذاشته شد: " + "، ".join(_ordered_months(incomplete)) + "."
            )
        if missing:
            note_parts.append(
                "ماه بدون داده در یک دوره: " + "، ".join(_ordered_months(set(missing))) + "."
            )
        return ChangeResult(
            period_a=period_a,
            period_b=period_b,
            total_a=total_a,
            total_b=total_b,
            difference=difference,
            percent=percent,
            months_used=complete,
            months_missing=_ordered_months(set(missing)),
            incomplete_excluded=_ordered_months(incomplete),
            note=" ".join(note_parts),
        )

    def to_dict(self) -> dict[str, Any]:
        periods = self.periods()
        return {
            "metric": self.metric,
            "periods": periods,
            "cells": [cell.to_dict() for cell in self.cells],
            "totals": {period: _text(self.total(period)) for period in periods},
            "best": (
                {
                    period: (self.best(period).label if self.best(period) else None)
                    for period in periods
                }
            ),
            "worst": (
                {
                    period: (self.worst(period).label if self.worst(period) else None)
                    for period in periods
                }
            ),
            "comparisons": [
                self.compare(periods[index], periods[index + 1]).to_dict()
                for index in range(len(periods) - 1)
            ],
        }


@dataclass
class IntegrationCheck:
    code: str
    label: str
    status: str
    detail: str = ""
    affected: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "label": self.label,
            "status": self.status,
            "detail": self.detail,
            "affected": self.affected,
        }


@dataclass
class PeriodIntegration:
    source_kind: str = "workbook"
    months: list[str] = field(default_factory=list)
    periods: list[str] = field(default_factory=list)
    series: list[MetricSeries] = field(default_factory=list)
    sources: list[str] = field(default_factory=list)
    excluded_sources: list[str] = field(default_factory=list)
    checks: list[IntegrationCheck] = field(default_factory=list)
    alignment: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    verified_cells: int = 0
    mismatched_cells: int = 0
    missing_metric_records: int = 0
    # Periods whose label only covers one month: the label is not a real
    # comparison year in this workbook (a mislabelled or stub column).
    suspicious_periods: list[str] = field(default_factory=list)

    # ------------------------------------------------------------------ #
    @property
    def metric_names(self) -> list[str]:
        return [item.metric for item in self.series]

    def metric(self, name: str) -> MetricSeries | None:
        for item in self.series:
            if item.metric == name:
                return item
        return None

    def select(self, names: Iterable[str] | None = None) -> list[MetricSeries]:
        if not names:
            return list(self.series)
        wanted = list(names)
        chosen = [item for item in self.series if item.metric in wanted]
        return chosen or list(self.series)

    def headline_metric(self) -> MetricSeries | None:
        """Largest cumulative volume in the newest period, by computed totals."""

        if not self.series:
            return None
        period = self.periods[-1] if self.periods else ""
        scored = [
            (item.total(period) if period else None, item) for item in self.series
        ]
        scored = [(value, item) for value, item in scored if value is not None]
        if not scored:
            return self.series[0]
        return max(scored, key=lambda pair: pair[0])[1]

    def period_months(self, period: str) -> list[str]:
        months = {
            cell.month
            for item in self.series
            for cell in item.cells
            if cell.period == period and cell.total is not None and cell.complete
        }
        return [name for name in JALALI_MONTHS if name in months]

    def periods_summary(self) -> str:
        """Human label per period: month count plus a flag for partly covered periods."""

        if not self.periods:
            return "—"
        widest = max(len(self.period_months(period)) for period in self.periods)
        labels = []
        for period in self.periods:
            months = len(self.period_months(period))
            suffix = " — فقط بخشی از برگه‌ها" if 0 < months < widest else ""
            labels.append(f"{period} ({months} ماه{suffix})")
        return "، ".join(labels)

    def period_observations(self, period: str) -> int:
        return sum(
            cell.observations
            for item in self.series
            for cell in item.cells
            if cell.period == period and cell.total is not None
        )

    def main_period(self) -> str:
        """Period with the widest month coverage, then the most observations."""

        if not self.periods:
            return ""
        return max(
            self.periods,
            key=lambda period: (
                len(self.period_months(period)),
                self.period_observations(period),
                period,
            ),
        )

    def comparison_periods(self) -> list[str]:
        """The two periods that are actually comparable in this workbook."""

        substantial = [
            period
            for period in self.periods
            if len(self.period_months(period)) >= 2
        ]
        if len(substantial) < 2:
            return substantial
        return substantial[-2:]

    @property
    def is_usable(self) -> bool:
        return len(self.months) >= 2 and bool(self.series)

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_kind": self.source_kind,
            "months": list(self.months),
            "periods": list(self.periods),
            "sources": list(self.sources),
            "excluded_sources": list(self.excluded_sources),
            "verified_cells": self.verified_cells,
            "mismatched_cells": self.mismatched_cells,
            "missing_metric_records": self.missing_metric_records,
            "suspicious_periods": list(self.suspicious_periods),
            "period_months": {period: self.period_months(period) for period in self.periods},
            "checks": [check.to_dict() for check in self.checks],
            "alignment": list(self.alignment),
            "warnings": list(self.warnings),
            "series": [item.to_dict() for item in self.series],
        }


# --------------------------------------------------------------------------- #
# Builder
# --------------------------------------------------------------------------- #
class PeriodIntegrator:
    """Merge per-sheet period totals and verify every cell against raw data."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def build(
        self,
        sheets: Iterable[Any],
        grids: dict[str, SheetGrid] | None = None,
        *,
        formulas: dict[str, Any] | None = None,
        source_kind: str = "workbook",
    ) -> PeriodIntegration:
        grids = grids or {}
        formulas = formulas or {}
        sheets = list(sheets)
        integration = PeriodIntegration(source_kind=source_kind)
        series: dict[str, MetricSeries] = {}
        month_set: set[str] = set()
        period_set: set[str] = set()
        sources: list[str] = []
        checked = 0
        mismatched = 0

        for sheet in sheets:
            if getattr(sheet, "derived", False):
                integration.excluded_sources.append(sheet.sheet)
                continue
            if getattr(sheet, "row_oriented_periods", False):
                integration.excluded_sources.append(sheet.sheet)
                continue
            month = getattr(sheet, "month", "") or ""
            month_index = getattr(sheet, "month_index", None)
            if not month or month_index is None:
                if sheet.column_totals:
                    integration.excluded_sources.append(sheet.sheet)
                continue
            sources.append(sheet.sheet)
            grid = grids.get(sheet.sheet)
            completeness = not (
                getattr(sheet, "coverage", None) is not None
                and getattr(sheet.coverage, "incomplete_period", False)
            )
            total_row = (getattr(sheet, "summary_rows", {}) or {}).get("total", [])
            total_row_number = total_row[0] if total_row else 0
            for record in getattr(sheet, "column_totals", []) or []:
                period = str(record.get("period") or "")
                metric = str(record.get("metric") or "")
                total = record.get("total")
                letter = str(record.get("letter") or "")
                if not metric or not period or letter == "":
                    continue
                if isinstance(total, (int, float)) and not isinstance(total, bool):
                    total = Decimal(str(total))
                if total is not None and not isinstance(total, Decimal):
                    continue
                cell = PeriodTotal(
                    metric=metric,
                    period=period,
                    month=month,
                    month_index=month_index,
                    source=sheet.sheet,
                    total=total,
                    observations=int(record.get("count") or 0),
                    complete=completeness,
                    total_address=f"{letter}{total_row_number}" if total_row_number else "",
                )
                if grid is not None:
                    verification, recomputed = self._verify(grid, sheet, letter, total)
                    cell.verification = verification
                    checked += 1
                    if verification == MISMATCH:
                        mismatched += 1
                        cell.note = f"جمع بازمحاسبه‌شده {_text(recomputed)} با مجموع ثبت‌شده {_text(total)} یکسان نیست."
                if total_row_number and grid is not None:
                    declared = self._declared_total(grid, total_row_number, letter)
                    cell.declared_total = declared
                bucket = series.setdefault(metric, MetricSeries(metric=metric))
                bucket.cells.append(cell)
                month_set.add(month)
                period_set.add(period)

        for bucket in series.values():
            bucket.cells.sort(key=lambda item: (item.month_index or 99, item.period))

        integration.months = [name for name in JALALI_MONTHS if name in month_set]
        integration.periods = _ordered_periods(period_set)
        newest = integration.periods[-1] if integration.periods else ""
        integration.series = sorted(
            series.values(),
            key=lambda item: (-(item.total(newest) or ZERO), item.metric),
        )
        integration.sources = sources
        integration.verified_cells = checked
        integration.mismatched_cells = mismatched
        integration.missing_metric_records = sum(
            1
            for bucket in integration.series
            for period in integration.periods
            if bucket.total(period) is None
        )
        integration.suspicious_periods = _suspicious_periods(integration)
        integration.alignment = self._summary_alignment(sheets, formulas)
        integration.checks = self._checks(integration)
        integration.warnings = self._warnings(integration)
        return integration

    def merge(self, integrations: Iterable[PeriodIntegration], *, source_kind: str = "folder") -> PeriodIntegration:
        """Combine already-built per-file integrations into one matrix."""

        merged = PeriodIntegration(source_kind=source_kind)
        buckets: dict[str, MetricSeries] = {}
        months: set[str] = set()
        periods: set[str] = set()
        for integration in integrations:
            for sheet in integration.excluded_sources:
                merged.excluded_sources.append(sheet)
            for name in integration.sources:
                merged.sources.append(name)
            for bucket in integration.series:
                target = buckets.setdefault(bucket.metric, MetricSeries(metric=bucket.metric))
                target.cells.extend(bucket.cells)
                months.update(bucket.months())
                periods.update(bucket.periods())
            merged.verified_cells += integration.verified_cells
            merged.mismatched_cells += integration.mismatched_cells
            merged.alignment.extend(integration.alignment)
            for warning in integration.warnings:
                if warning not in merged.warnings:
                    merged.warnings.append(warning)
        merged.series = list(buckets.values())
        merged.months = [name for name in JALALI_MONTHS if name in months]
        merged.periods = _ordered_periods(periods)
        merged.suspicious_periods = _suspicious_periods(merged)
        merged.checks = self._checks(merged)
        return merged

    # ------------------------------------------------------------------ #
    @staticmethod
    def _verify(
        grid: SheetGrid,
        sheet: Any,
        letter: str,
        reported: Decimal | None,
    ) -> tuple[str, Decimal | None]:
        """Recompute one column without using the intelligence layer's totals."""

        start = getattr(sheet, "data_start", 0) or 0
        end = getattr(sheet, "data_end", 0) or 0
        if not start or not end or end < start:
            return NOT_CHECKED, None
        try:
            from openpyxl.utils import column_index_from_string

            column_index = column_index_from_string(letter)
        except (ValueError, KeyError):
            return NOT_CHECKED, None
        values: list[Decimal] = []
        errors = 0
        columns = {profile.letter: profile for profile in getattr(sheet, "columns", [])}
        profile = columns.get(letter)
        role = getattr(profile, "role", None)
        if role in {ColumnRole.ID, ColumnRole.DATE}:
            return NOT_CHECKED, None
        role_map = {
            item["row"]: item["role"]
            for item in (getattr(sheet, "row_roles", []) or [])
        }
        for row in range(start, end + 1):
            if role_map.get(row, "data") != "data":
                continue
            raw = grid.value(row, column_index)
            parsed = parse_value(raw, kind_hint=ValueKind.NUMBER)
            if parsed.kind is ValueKind.NUMBER and isinstance(parsed.value, Decimal):
                values.append(parsed.value)
            elif parsed.kind is ValueKind.ERROR:
                errors += 1
        recomputed = sum(values, ZERO) if values else None
        if reported is None or recomputed is None:
            return NOT_CHECKED, recomputed
        return (VERIFIED if reported == recomputed else MISMATCH), recomputed

    @staticmethod
    def _declared_total(grid: SheetGrid, row: int, letter: str) -> Decimal | None:
        from openpyxl.utils import column_index_from_string

        try:
            column_index = column_index_from_string(letter)
        except (ValueError, KeyError):
            return None
        parsed = parse_value(grid.value(row, column_index), kind_hint=ValueKind.NUMBER)
        return parsed.value if isinstance(parsed.value, Decimal) else None

    # ------------------------------------------------------------------ #
    @staticmethod
    def _summary_alignment(
        sheets: Iterable[Any], formulas: dict[str, Any] | None = None
    ) -> list[dict[str, Any]]:
        """Find aggregate sheets whose formulas read a different year column.

        The aggregate sheet keeps its own period label per column; when the
        formula points at a column of another sheet that carries a different
        period label, the summary is not a like-for-like total.  This is a real
        property of the source file, reported with cell addresses, not a guess.
        """

        sheets = list(sheets)
        formulas = formulas or {}
        by_name = {sheet.sheet: sheet for sheet in sheets}
        counts: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        for sheet in sheets:
            if not getattr(sheet, "derived", False):
                continue
            book = formulas.get(sheet.sheet)
            if book is None:
                continue
            probe_row = getattr(sheet, "data_start", 0) or 0
            if not probe_row:
                continue
            for profile in getattr(sheet, "columns", []):
                expected = (profile.period or "").strip()
                if not expected:
                    continue
                for referenced_sheet, column, _row in _formula_references(
                    book, probe_row, profile.letter
                ):
                    target = by_name.get(referenced_sheet)
                    if target is None:
                        continue
                    target_period = {
                        item.letter: (item.period or "") for item in target.columns
                    }.get(column, "")
                    if not target_period or target_period == expected:
                        continue
                    key = (sheet.sheet, expected, referenced_sheet, target_period)
                    bucket = counts.setdefault(
                        key,
                        {
                            "summary_sheet": sheet.sheet,
                            "summary_column": profile.letter,
                            "summary_period": expected,
                            "referenced_sheet": referenced_sheet,
                            "referenced_column": column,
                            "referenced_period": target_period,
                            "cells": 0,
                        },
                    )
                    bucket["cells"] += 1
        return sorted(counts.values(), key=lambda item: -item["cells"])

    # ------------------------------------------------------------------ #
    def _checks(self, integration: PeriodIntegration) -> list[IntegrationCheck]:
        checks: list[IntegrationCheck] = []
        checks.append(
            IntegrationCheck(
                code="period_sources",
                label="برگه‌های دوره‌ای قابل استفاده",
                status="passed" if integration.sources else "failed",
                affected=len(integration.sources),
                detail="، ".join(integration.sources) or "برگه دوره‌ای شناسایی نشد.",
            )
        )
        if integration.excluded_sources:
            checks.append(
                IntegrationCheck(
                    code="derived_sources",
                    label="برگه‌های مشتق/بدون دوره (کنار گذاشته شد)",
                    status="warning",
                    affected=len(integration.excluded_sources),
                    detail="، ".join(integration.excluded_sources)
                    + "؛ این برگه‌ها از فرمول سایر برگه‌ها ساخته شده‌اند و در ماتریس ماهانه دوباره جمع نمی‌شوند.",
                )
            )
        if integration.verified_cells:
            status = "passed" if integration.mismatched_cells == 0 else "failed"
            checks.append(
                IntegrationCheck(
                    code="recalculation",
                    label="بازمحاسبه مستقل جمع هر ستون",
                    status=status,
                    affected=integration.mismatched_cells,
                    detail=(
                        f"{integration.verified_cells} ستون بازمحاسبه شد؛ "
                        f"{integration.mismatched_cells} مورد با مجموع ثبت‌شده اختلاف داشت."
                    ),
                )
            )
        if integration.alignment:
            affected = sum(item["cells"] for item in integration.alignment)
            checks.append(
                IntegrationCheck(
                    code="summary_alignment",
                    label="هم‌ترازی سال در برگه خلاصه",
                    status="failed",
                    affected=affected,
                    detail=(
                        "برگه خلاصه، ستون سال را از ستون برگه‌های ماهانه می‌خواند که برچسب سال دیگری دارند؛ "
                        "جمع برگه خلاصه با جمع ماهانه هم‌تعریف نیست."
                    ),
                )
            )
        if integration.suspicious_periods:
            checks.append(
                IntegrationCheck(
                    code="period_labels",
                    label="برچسب دوره‌های ناقص/مشکوک",
                    status="warning",
                    affected=len(integration.suspicious_periods),
                    detail=(
                        "این دوره‌ها در کمتر از دو ماه داده دارند و برچسب سال آن‌ها در فایل ناهم‌خوان است: "
                        + "، ".join(integration.suspicious_periods)
                        + "؛ از مقایسه پیش‌فرض کنار گذاشته شدند."
                    ),
                )
            )
        if integration.missing_metric_records:
            checks.append(
                IntegrationCheck(
                    code="metric_coverage",
                    label="پوشش شاخص در همه دوره‌ها",
                    status="warning",
                    affected=integration.missing_metric_records,
                    detail=(
                        "برای بخشی از شاخص‌ها در یک دوره داده‌ای وجود ندارد؛ "
                        "خانه خالی به صفر تبدیل نشد."
                    ),
                )
            )
        return checks

    @staticmethod
    def _warnings(integration: PeriodIntegration) -> list[str]:
        warnings: list[str] = []
        if integration.mismatched_cells:
            warnings.append(
                f"{integration.mismatched_cells} ستون با بازمحاسبه مستقل اختلاف دارد؛ "
                "مقدار محاسبه‌شده در گزارش استفاده شد و اختلاف ثبت شد."
            )
        for item in integration.alignment:
            warnings.append(
                f"«{item['summary_sheet']}» ستون {item['summary_column']} (سال {item['summary_period']}) "
                f"را از ستون {item['referenced_column']} برگه «{item['referenced_sheet']}» می‌خواند "
                f"که سال {item['referenced_period']} است."
            )
        for check in integration.checks:
            if check.code in {"period_sources", "recalculation", "summary_alignment"}:
                continue
            if check.status in {"warning", "failed"}:
                warnings.append(check.detail)
        return warnings


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _formula_references(formulas: Any, row: int, letter: str) -> list[tuple[str, str, int]]:
    try:
        from openpyxl.utils import column_index_from_string

        value = formulas.value(row, column_index_from_string(letter))
    except (ValueError, KeyError, AttributeError):
        return []
    if value is None or not isinstance(value, str) or not value.startswith("="):
        return []
    return [
        (match.group(1) or match.group(2), match.group(3), int(match.group(4)))
        for match in _REF_RE.finditer(value)
    ]


def _suspicious_periods(integration: PeriodIntegration) -> list[str]:
    """Period labels that cover far fewer months than the best-covered period.

    A column labelled with a year that only exists for one or two months is a
    stub or a mislabel, not a comparison period; it stays visible but is not
    used as the basis of a comparison.
    """

    if not integration.periods:
        return []
    main = integration.main_period()
    coverage = {period: len(integration.period_months(period)) for period in integration.periods}
    threshold = max(2, coverage.get(main, 0) // 2)
    return [
        period
        for period in integration.periods
        if period != main and coverage.get(period, 0) < threshold
    ]


def _ordered_periods(values: Iterable[str]) -> list[str]:
    def key(value: str) -> tuple[int, str]:
        digits = re.sub(r"\D", "", value)
        return (int(digits) if digits else 0, value)

    return sorted(set(values), key=key)


def _ordered_months(values: Iterable[str]) -> list[str]:
    known = set(values)
    ordered = [name for name in JALALI_MONTHS if name in known]
    return ordered + sorted(known - set(ordered))


def _text(value: Decimal | None) -> str | None:
    if value is None:
        return None
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f")
