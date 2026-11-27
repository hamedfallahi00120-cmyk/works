"""General-purpose statistical profiling for office and survey workbooks.

Nothing here depends on loans, money or a particular business vocabulary.  The
analyzer reads a detected table, standardizes each cell with the shared
normalizers, and chooses analyses from the observed data types.  It never
deletes rows: outliers and missing values are labelled and counted, and the
method used for every figure is recorded.
"""

from __future__ import annotations

import math
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Any

from local_office_ai.core.models import ColumnRole, DataRole, TableCandidate, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    format_jalali,
    gregorian_to_jalali,
    parse_value,
)
from local_office_ai.excel.sheet_analyzer import SheetGrid

ZERO = Decimal("0")
HUNDRED = Decimal("100")


@dataclass
class ColumnStats:
    key: str
    header: str
    letter: str
    index: int
    kind: ValueKind
    role: ColumnRole
    count: int
    missing: int
    distinct: int
    coverage: float
    is_identifier: bool = False
    is_temporal_key: bool = False
    total: Decimal | None = None
    mean: Decimal | None = None
    median: Decimal | None = None
    mode: list[str] = field(default_factory=list)
    mode_count: int = 0
    minimum: Decimal | None = None
    maximum: Decimal | None = None
    value_range: Decimal | None = None
    variance_population: Decimal | None = None
    variance_sample: Decimal | None = None
    std_population: Decimal | None = None
    std_sample: Decimal | None = None
    cv_percent: Decimal | None = None
    q1: Decimal | None = None
    q3: Decimal | None = None
    iqr: Decimal | None = None
    p10: Decimal | None = None
    p90: Decimal | None = None
    date_min: str = ""
    date_max: str = ""
    date_calendar: str = ""
    range_days: int | None = None
    top_values: list[tuple[str, int]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_numeric(self) -> bool:
        return self.kind is ValueKind.NUMBER

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "header": self.header,
            "letter": self.letter,
            "index": self.index,
            "kind": self.kind.value,
            "role": self.role.value,
            "count": self.count,
            "missing": self.missing,
            "distinct": self.distinct,
            "coverage": round(self.coverage, 4),
            "is_identifier": self.is_identifier,
            "is_temporal_key": self.is_temporal_key,
            "total": _d(self.total),
            "mean": _d(self.mean),
            "median": _d(self.median),
            "mode": list(self.mode),
            "mode_count": self.mode_count,
            "minimum": _d(self.minimum),
            "maximum": _d(self.maximum),
            "range": _d(self.value_range),
            "variance_population": _d(self.variance_population),
            "variance_sample": _d(self.variance_sample),
            "std_population": _d(self.std_population),
            "std_sample": _d(self.std_sample),
            "cv_percent": _d(self.cv_percent),
            "q1": _d(self.q1),
            "q3": _d(self.q3),
            "iqr": _d(self.iqr),
            "p10": _d(self.p10),
            "p90": _d(self.p90),
            "date_min": self.date_min,
            "date_max": self.date_max,
            "date_calendar": self.date_calendar,
            "range_days": self.range_days,
            "top_values": [[value, count] for value, count in self.top_values],
            "warnings": list(self.warnings),
        }


@dataclass
class FrequencyRow:
    value: str
    count: int
    percent: Decimal
    cumulative_percent: Decimal


@dataclass
class HistogramBin:
    lower: Decimal
    upper: Decimal
    count: int
    percent: Decimal


@dataclass
class FrequencyTable:
    key: str
    header: str
    letter: str
    kind: ValueKind
    total: int
    distinct: int
    rows: list[FrequencyRow] = field(default_factory=list)
    bins: list[HistogramBin] = field(default_factory=list)
    other_count: int = 0
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "header": self.header,
            "letter": self.letter,
            "kind": self.kind.value,
            "total": self.total,
            "distinct": self.distinct,
            "rows": [
                {
                    "value": row.value,
                    "count": row.count,
                    "percent": _d(row.percent),
                    "cumulative_percent": _d(row.cumulative_percent),
                }
                for row in self.rows
            ],
            "bins": [
                {
                    "lower": _d(bin.lower),
                    "upper": _d(bin.upper),
                    "count": bin.count,
                    "percent": _d(bin.percent),
                }
                for bin in self.bins
            ],
            "other_count": self.other_count,
            "note": self.note,
        }


@dataclass
class GroupValue:
    label: str
    count: int
    total: Decimal
    mean: Decimal
    median: Decimal
    minimum: Decimal
    maximum: Decimal
    std_sample: Decimal | None


@dataclass
class GroupComparison:
    group_key: str
    group_header: str
    measure_key: str
    measure_header: str
    groups: list[GroupValue]
    missing_pairs: int
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "group_key": self.group_key,
            "group_header": self.group_header,
            "measure_key": self.measure_key,
            "measure_header": self.measure_header,
            "groups": [
                {
                    "label": group.label,
                    "count": group.count,
                    "total": _d(group.total),
                    "mean": _d(group.mean),
                    "median": _d(group.median),
                    "minimum": _d(group.minimum),
                    "maximum": _d(group.maximum),
                    "std_sample": _d(group.std_sample),
                }
                for group in self.groups
            ],
            "missing_pairs": self.missing_pairs,
            "note": self.note,
        }


@dataclass
class CrossTabRow:
    label: str
    counts: dict[str, int]
    total: int


@dataclass
class CrossTab:
    row_key: str
    row_header: str
    column_key: str
    column_header: str
    columns: list[str]
    rows: list[CrossTabRow]
    column_totals: dict[str, int]
    total: int
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "row_key": self.row_key,
            "row_header": self.row_header,
            "column_key": self.column_key,
            "column_header": self.column_header,
            "columns": list(self.columns),
            "rows": [
                {"label": row.label, "counts": dict(row.counts), "total": row.total}
                for row in self.rows
            ],
            "column_totals": dict(self.column_totals),
            "total": self.total,
            "note": self.note,
        }


@dataclass
class TrendPoint:
    period: str
    period_label: str
    count: int
    total: Decimal
    mean: Decimal


@dataclass
class TrendAnalysis:
    date_key: str
    date_header: str
    measure_key: str
    measure_header: str
    granularity: str
    points: list[TrendPoint]
    direction: str
    change_absolute: Decimal | None
    change_percent: Decimal | None
    note: str = ""
    calendar: str = "gregorian"

    def to_dict(self) -> dict[str, Any]:
        return {
            "date_key": self.date_key,
            "date_header": self.date_header,
            "measure_key": self.measure_key,
            "measure_header": self.measure_header,
            "granularity": self.granularity,
            "points": [
                {
                    "period": point.period,
                    "period_label": point.period_label,
                    "count": point.count,
                    "total": _d(point.total),
                    "mean": _d(point.mean),
                }
                for point in self.points
            ],
            "direction": self.direction,
            "change_absolute": _d(self.change_absolute),
            "change_percent": _d(self.change_percent),
            "note": self.note,
            "calendar": self.calendar,
        }


@dataclass
class CorrelationPair:
    column_a: str
    header_a: str
    column_b: str
    header_b: str
    count: int
    pearson: Decimal | None
    spearman: Decimal | None
    strength: str
    direction: str
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "column_a": self.column_a,
            "header_a": self.header_a,
            "column_b": self.column_b,
            "header_b": self.header_b,
            "count": self.count,
            "pearson": _d(self.pearson),
            "spearman": _d(self.spearman),
            "strength": self.strength,
            "direction": self.direction,
            "note": self.note,
        }


@dataclass
class OutlierRecord:
    column_key: str
    header: str
    row: int
    address: str
    raw: Any
    value: Decimal
    method: str
    lower_fence: Decimal
    upper_fence: Decimal
    score: Decimal

    def to_dict(self) -> dict[str, Any]:
        return {
            "column_key": self.column_key,
            "header": self.header,
            "row": self.row,
            "address": self.address,
            "raw": _jsonable(self.raw),
            "value": _d(self.value),
            "method": self.method,
            "lower_fence": _d(self.lower_fence),
            "upper_fence": _d(self.upper_fence),
            "score": _d(self.score),
        }


@dataclass
class DataQuality:
    row_count: int
    column_count: int
    total_cells: int
    missing_cells: int
    missing_ratio: float
    duplicate_rows: int
    empty_rows: int
    constant_columns: list[str] = field(default_factory=list)
    mixed_type_columns: list[str] = field(default_factory=list)
    identifier_columns: list[str] = field(default_factory=list)
    temporal_columns: list[str] = field(default_factory=list)
    invalid_cells: int = 0
    duplicate_examples: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    # A column with no measured value at all is an unmeasured period/key, not
    # scattered missing data; the two ratios are reported separately.
    empty_columns: list[str] = field(default_factory=list)
    measured_cells: int = 0
    measured_missing_cells: int = 0
    measured_missing_ratio: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return {
            "row_count": self.row_count,
            "column_count": self.column_count,
            "total_cells": self.total_cells,
            "missing_cells": self.missing_cells,
            "missing_ratio": round(self.missing_ratio, 4),
            "empty_columns": list(self.empty_columns),
            "measured_cells": self.measured_cells,
            "measured_missing_cells": self.measured_missing_cells,
            "measured_missing_ratio": round(self.measured_missing_ratio, 4),
            "duplicate_rows": self.duplicate_rows,
            "empty_rows": self.empty_rows,
            "constant_columns": list(self.constant_columns),
            "mixed_type_columns": list(self.mixed_type_columns),
            "identifier_columns": list(self.identifier_columns),
            "temporal_columns": list(self.temporal_columns),
            "invalid_cells": self.invalid_cells,
            "duplicate_examples": list(self.duplicate_examples),
            "notes": list(self.notes),
        }


@dataclass
class GenericAnalysis:
    table_sheet: str
    table_address: str
    table_kind: str
    row_count: int
    column_count: int
    columns: list[ColumnStats] = field(default_factory=list)
    frequencies: list[FrequencyTable] = field(default_factory=list)
    groups: list[GroupComparison] = field(default_factory=list)
    crosstabs: list[CrossTab] = field(default_factory=list)
    trends: list[TrendAnalysis] = field(default_factory=list)
    correlations: list[CorrelationPair] = field(default_factory=list)
    outliers: list[OutlierRecord] = field(default_factory=list)
    quality: DataQuality | None = None
    methodology: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    goal: str | None = None

    def has_section(self, name: str) -> bool:
        return bool(getattr(self, name, None))

    def to_dict(self) -> dict[str, Any]:
        return {
            "table_sheet": self.table_sheet,
            "table_address": self.table_address,
            "table_kind": self.table_kind,
            "row_count": self.row_count,
            "column_count": self.column_count,
            "columns": [column.to_dict() for column in self.columns],
            "frequencies": [item.to_dict() for item in self.frequencies],
            "groups": [item.to_dict() for item in self.groups],
            "crosstabs": [item.to_dict() for item in self.crosstabs],
            "trends": [item.to_dict() for item in self.trends],
            "correlations": [item.to_dict() for item in self.correlations],
            "outliers": [item.to_dict() for item in self.outliers],
            "quality": self.quality.to_dict() if self.quality else None,
            "methodology": list(self.methodology),
            "warnings": list(self.warnings),
            "goal": self.goal,
        }


class StatisticalAnalyzer:
    """Type-aware descriptive statistics and exploratory analysis."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.max_rows = max(100, int(getattr(settings, "max_stats_rows", 20_000)))
        self.max_frequency_values = max(5, int(getattr(settings, "max_frequency_values", 40)))
        self.max_outliers = max(10, int(getattr(settings, "max_outliers", 120)))

    # ------------------------------------------------------------------ #
    def analyze(self, grid: SheetGrid, table: TableCandidate, *, goal: str | None = None) -> GenericAnalysis:
        channel = _TableChannel(grid, table, self.max_rows)
        if not channel.rows:
            quality = DataQuality(0, len(table.columns), 0, 0, 0.0, 0, 0)
            return GenericAnalysis(
                table.sheet,
                table.address,
                table.kind.value,
                table.row_count,
                len(table.columns),
                quality=quality,
                methodology=["جدول داده‌ای برای تحلیل آماری پیدا نشد."],
                warnings=["هیچ ردیف داده‌ای در جدول اصلی وجود ندارد."],
                goal=goal,
            )

        columns = [self._profile_column(column, channel) for column in table.columns]
        numeric = [
            column
            for column in columns
            if column.is_numeric and not column.is_identifier and not column.is_temporal_key
        ]
        date_columns = [column for column in columns if column.kind is ValueKind.DATE]
        temporal_columns = [column for column in columns if column.is_temporal_key]
        text_columns = [column for column in columns if column.kind is ValueKind.TEXT]

        frequencies: list[FrequencyTable] = []
        for column in text_columns:
            if column.is_identifier:
                continue
            if column.distinct and column.distinct <= max(self.max_frequency_values, int(len(channel.rows) * 0.4)):
                frequencies.append(self._frequency_text(column, channel))
        for column in numeric:
            if column.distinct > 10:
                frequencies.append(self._frequency_numeric(column, channel))

        groups = self._group_comparisons(channel, text_columns, numeric)
        crosstabs = self._crosstabs(channel, text_columns)
        trends = self._trends(channel, date_columns, numeric, temporal_columns)
        correlations = self._correlations(channel, numeric)
        outliers = self._outliers(channel, numeric)
        quality = self._quality(channel, columns)

        methodology = self._methodology(channel, columns, groups, crosstabs, trends, correlations, outliers)
        warnings: list[str] = []
        if channel.truncated:
            warnings.append(f"تحلیل آماری روی {len(channel.rows)} ردیف اول اجرا شد؛ سقف پیکربندی‌شده اعمال شده است.")
        if quality.missing_ratio >= 0.2:
            warnings.append(f"{quality.missing_ratio:.0%} سلول‌های جدول خالی است؛ تفسیر نتایج با احتیاط انجام شود.")
        if quality.duplicate_rows:
            warnings.append(f"{quality.duplicate_rows} ردیف کاملاً تکراری شناسایی شد؛ حذف خودکار انجام نشد.")
        if not numeric:
            warnings.append("ستون عددی معتبری برای محاسبات آماری پیدا نشد.")
        if len(channel.rows) < 8:
            warnings.append("تعداد ردیف‌ها کم است؛ آماره‌های پرت و همبستگی قدرت پایینی دارند.")

        return GenericAnalysis(
            table_sheet=table.sheet,
            table_address=table.address,
            table_kind=table.kind.value,
            row_count=len(channel.rows),
            column_count=len(table.columns),
            columns=columns,
            frequencies=frequencies,
            groups=groups,
            crosstabs=crosstabs,
            trends=trends,
            correlations=correlations,
            outliers=outliers,
            quality=quality,
            methodology=methodology,
            warnings=warnings,
            goal=goal,
        )

    # ------------------------------------------------------------------ #
    def _profile_column(self, column, channel: "_TableChannel") -> ColumnStats:
        key = f"{column.letter}:{column.header or column.letter}"
        values = channel.values[column.index]
        numbers = channel.numbers[column.index]
        dates = channel.dates[column.index]
        texts = channel.texts[column.index]
        non_empty = len(values)
        kinds = channel.kinds[column.index]
        total = max(1, non_empty)

        numeric_ratio = kinds[ValueKind.NUMBER] / total
        date_ratio = kinds[ValueKind.DATE] / total
        text_ratio = (kinds[ValueKind.TEXT] + kinds[ValueKind.BOOLEAN]) / total

        if numeric_ratio >= 0.8 and numbers:
            kind = ValueKind.NUMBER
        elif date_ratio >= 0.8 and dates:
            kind = ValueKind.DATE
        elif texts:
            kind = ValueKind.TEXT
        else:
            kind = ValueKind.TEXT if non_empty else ValueKind.EMPTY

        identifier = self._is_identifier(column, values, texts, numbers, kind)
        temporal = column.statistical_role is DataRole.TEMPORAL_KEY
        if temporal:
            identifier = False
        stats = ColumnStats(
            key=key,
            header=column.header or column.letter,
            letter=column.letter,
            index=column.index,
            kind=kind,
            role=column.role,
            count=non_empty,
            missing=max(0, len(channel.rows) - non_empty),
            distinct=len({str(value) for value in values}),
            coverage=non_empty / max(1, len(channel.rows)),
            is_identifier=identifier,
            is_temporal_key=temporal,
        )

        if kind is ValueKind.NUMBER and numbers and not temporal:
            self._fill_numeric(stats, numbers)
        elif kind is ValueKind.DATE and dates:
            methods = channel.date_methods.get(column.index, Counter())
            jalali_count = sum(count for method, count in methods.items() if method.startswith("jalali"))
            use_jalali = jalali_count > len(dates) / 2
            stats.date_calendar = "jalali" if use_jalali else "gregorian"
            if use_jalali:
                stats.date_min = format_jalali(min(dates))
                stats.date_max = format_jalali(max(dates))
            else:
                stats.date_min = min(dates).isoformat()
                stats.date_max = max(dates).isoformat()
            stats.range_days = (max(dates) - min(dates)).days
            stats.minimum = Decimal(str((min(dates) - date(1970, 1, 1)).days))
            stats.maximum = Decimal(str((max(dates) - date(1970, 1, 1)).days))

        counts = Counter(str(value) for value in values)
        top = counts.most_common(8)
        stats.top_values = [(value, count) for value, count in top]
        if top and top[0][1] > 1:
            stats.mode = [value for value, count in top if count == top[0][1]][:5]
            stats.mode_count = top[0][1]
        if kind is ValueKind.NUMBER and non_empty > len(numbers):
            stats.warnings.append(
                f"{non_empty - len(numbers)} مقدار غیرعددی در ستون عددی کنار گذاشته شد."
            )
        if kind is ValueKind.DATE and non_empty > len(dates):
            stats.warnings.append(
                f"{non_empty - len(dates)} مقدار غیرتاریخی در ستون تاریخ کنار گذاشته شد."
            )
        if identifier:
            stats.warnings.append("ستون به‌عنوان شناسه تشخیص داده شد؛ در آماره‌های توصیفی و هم‌بستگی استفاده نمی‌شود.")
        if temporal:
            stats.warnings.append(
                "ستون کلید زمانی است (تاریخ یا شماره روز)؛ میانگین و همبستگی آن محاسبه نشد."
            )
        if 0.2 <= numeric_ratio < 0.8 and text_ratio >= 0.2:
            stats.warnings.append("ستون مقادیر عددی و متنی را با هم دارد؛ نوع غالب انتخاب شده است.")
        return stats

    def _fill_numeric(self, stats: ColumnStats, numbers: list[Decimal]) -> None:
        values = sorted(numbers)
        n = len(values)
        stats.count = n
        stats.total = sum(values, ZERO)
        stats.mean = stats.total / Decimal(n)
        stats.median = _percentile(values, Decimal(50))
        stats.minimum = values[0]
        stats.maximum = values[-1]
        stats.value_range = stats.maximum - stats.minimum
        stats.q1 = _percentile(values, Decimal(25))
        stats.q3 = _percentile(values, Decimal(75))
        stats.iqr = stats.q3 - stats.q1
        stats.p10 = _percentile(values, Decimal(10))
        stats.p90 = _percentile(values, Decimal(90))
        mean = stats.mean
        squared = sum(((value - mean) ** 2 for value in values), ZERO)
        stats.variance_population = squared / Decimal(n)
        stats.variance_sample = squared / Decimal(n - 1) if n > 1 else None
        stats.std_population = _sqrt(stats.variance_population)
        stats.std_sample = _sqrt(stats.variance_sample) if stats.variance_sample is not None else None
        if mean != 0 and stats.std_sample is not None:
            stats.cv_percent = (stats.std_sample / abs(mean) * HUNDRED).quantize(Decimal("0.01"))

    def _is_identifier(self, column, values, texts, numbers, kind) -> bool:
        if column.role is ColumnRole.ID:
            return True
        if not values:
            return False
        unique_ratio = len({str(value) for value in values}) / len(values)
        header = (column.header or "").lower()
        id_words = ("کد", "شناسه", "شماره", "ردیف", "id", "code", "number")
        if any(word in header for word in id_words) and unique_ratio >= 0.9:
            return True
        if unique_ratio >= 0.98 and len(values) >= 8 and kind is ValueKind.NUMBER and numbers:
            integers = [int(number) for number in numbers if number == int(number)]
            if len(integers) == len(numbers):
                consecutive = all(
                    integers[index + 1] == integers[index] + 1 for index in range(len(integers) - 1)
                )
                if consecutive and abs(integers[0]) <= 1:
                    return True
        return False

    # ------------------------------------------------------------------ #
    def _frequency_text(self, column: ColumnStats, channel: "_TableChannel") -> FrequencyTable:
        values = [str(value) for value in channel.values[column.index]]
        counts = Counter(values)
        total = len(values)
        rows: list[FrequencyRow] = []
        cumulative = 0
        for value, count in counts.most_common(self.max_frequency_values):
            percent = Decimal(count) / Decimal(total) * HUNDRED
            cumulative += count
            rows.append(
                FrequencyRow(
                    value,
                    count,
                    percent.quantize(Decimal("0.01")),
                    (Decimal(cumulative) / Decimal(total) * HUNDRED).quantize(Decimal("0.01")),
                )
            )
        other = total - sum(row.count for row in rows)
        note = ""
        if other:
            note = f"{other} مقدار دیگر در {len(counts) - len(rows)} گروه یکتا نمایش داده نشده است."
        return FrequencyTable(
            column.key, column.header, column.letter, column.kind, total, len(counts),
            rows=rows, other_count=other, note=note,
        )

    def _frequency_numeric(self, column: ColumnStats, channel: "_TableChannel") -> FrequencyTable:
        values = channel.numbers[column.index]
        if not values:
            return FrequencyTable(column.key, column.header, column.letter, column.kind, 0, 0)
        n = len(values)
        minimum, maximum = min(values), max(values)
        bins_count = min(12, max(5, math.ceil(1 + 3.322 * math.log10(max(2, n)))))
        width = (maximum - minimum) / Decimal(bins_count)
        bins: list[HistogramBin] = []
        if width == 0:
            return FrequencyTable(
                column.key, column.header, column.letter, column.kind, n, 1,
                bins=[HistogramBin(minimum, maximum, n, HUNDRED)],
                note="همه مقادیر یکسان‌اند.",
            )
        counts = [0] * bins_count
        for value in values:
            index = int((value - minimum) / width)
            index = min(bins_count - 1, max(0, index))
            counts[index] += 1
        for index, count in enumerate(counts):
            lower = minimum + width * Decimal(index)
            upper = maximum if index == bins_count - 1 else lower + width
            bins.append(
                HistogramBin(
                    lower.quantize(Decimal("0.01")),
                    upper.quantize(Decimal("0.01")),
                    count,
                    (Decimal(count) / Decimal(n) * HUNDRED).quantize(Decimal("0.01")),
                )
            )
        return FrequencyTable(
            column.key, column.header, column.letter, column.kind, n, len(set(values)),
            bins=bins, note="بازه‌ها با قاعده استرجس و عرض مساوی ساخته شده‌اند.",
        )

    # ------------------------------------------------------------------ #
    def _group_comparisons(
        self,
        channel: "_TableChannel",
        text_columns: list[ColumnStats],
        numeric_columns: list[ColumnStats],
    ) -> list[GroupComparison]:
        group_candidates = [
            column
            for column in text_columns
            if 2 <= column.distinct <= 20 and not column.is_identifier and column.coverage >= 0.5
        ]
        group_candidates.sort(key=lambda column: (-column.coverage, -column.distinct))
        measure_candidates = [
            column
            for column in numeric_columns
            if column.count >= 3 and column.distinct > 1
        ]
        measure_candidates.sort(key=lambda column: (-column.coverage, -column.count))
        results: list[GroupComparison] = []
        for group_column in group_candidates[:2]:
            for measure_column in measure_candidates[:2]:
                result = self._group_stat(channel, group_column, measure_column)
                if result is not None:
                    results.append(result)
        return results

    def _group_stat(
        self,
        channel: "_TableChannel",
        group_column: ColumnStats,
        measure_column: ColumnStats,
    ) -> GroupComparison | None:
        grouped: dict[str, list[Decimal]] = defaultdict(list)
        missing = 0
        for row in channel.rows:
            group = channel.value_map.get((group_column.key, row))
            measure = channel.number_map.get((measure_column.key, row))
            if group is None or measure is None:
                missing += 1
                continue
            label = str(group).strip()
            if not label or _looks_like_total(label):
                continue
            grouped[label].append(measure)
        if len(grouped) < 2:
            return None
        groups: list[GroupValue] = []
        for label, values in grouped.items():
            ordered = sorted(values)
            total = sum(values, ZERO)
            groups.append(
                GroupValue(
                    label=label,
                    count=len(values),
                    total=total,
                    mean=total / Decimal(len(values)),
                    median=_percentile(ordered, Decimal(50)),
                    minimum=ordered[0],
                    maximum=ordered[-1],
                    std_sample=_std_sample(values),
                )
            )
        groups.sort(key=lambda item: item.total, reverse=True)
        return GroupComparison(
            group_column.key,
            group_column.header,
            measure_column.key,
            measure_column.header,
            groups[:40],
            missing,
            "میانگین و میانه هر گروه جداگانه محاسبه شده؛ تعداد مشاهدات هر گروه در ستون count است.",
        )

    # ------------------------------------------------------------------ #
    def _crosstabs(self, channel: "_TableChannel", text_columns: list[ColumnStats]) -> list[CrossTab]:
        candidates = [
            column
            for column in text_columns
            if 2 <= column.distinct <= 10 and not column.is_identifier and column.coverage >= 0.6
        ]
        best = None
        for i, first in enumerate(candidates):
            for second in candidates[i + 1:]:
                if first.distinct * second.distinct > 60:
                    continue
                score = (first.distinct * second.distinct, -(first.distinct + second.distinct))
                if best is None or score > best[0]:
                    best = (score, first, second)
        if best is None:
            return []
        _, first, second = best
        counts: Counter[tuple[str, str]] = Counter()
        first_values: Counter[str] = Counter()
        second_values: Counter[str] = Counter()
        total = 0
        for row in channel.rows:
            a = channel.value_map.get((first.key, row))
            b = channel.value_map.get((second.key, row))
            if a is None or b is None:
                continue
            a, b = str(a).strip(), str(b).strip()
            if not a or not b:
                continue
            counts[(a, b)] += 1
            first_values[a] += 1
            second_values[b] += 1
            total += 1
        if total < 4 or len(first_values) < 2 or len(second_values) < 2:
            return []
        columns = [value for value, _ in second_values.most_common(12)]
        rows = []
        for label, row_total in first_values.most_common(20):
            row_counts = {column: counts.get((label, column), 0) for column in columns}
            rows.append(CrossTabRow(label, row_counts, row_total))
        return [
            CrossTab(
                first.key,
                first.header,
                second.key,
                second.header,
                columns,
                rows,
                {column: second_values[column] for column in columns},
                total,
                "جدول متقاطع توصیفی است؛ برای استنباط آماری به آزمون مناسب و حجم نمونه کافی نیاز است.",
            )
        ]

    # ------------------------------------------------------------------ #
    def _trends(
        self,
        channel: "_TableChannel",
        date_columns: list[ColumnStats],
        numeric_columns: list[ColumnStats],
        temporal_columns: list[ColumnStats] | None = None,
    ) -> list[TrendAnalysis]:
        if not numeric_columns:
            return []
        if not date_columns:
            return self._sequential_trend(channel, temporal_columns or [], numeric_columns)
        date_column = max(date_columns, key=lambda column: (column.count, column.coverage))
        measure = max(
            numeric_columns,
            key=lambda column: (column.count, column.value_range or ZERO, column.distinct),
        )
        dates: list[date] = []
        for row in channel.rows:
            day = channel.date_map.get((date_column.key, row))
            value = channel.number_map.get((measure.key, row))
            if day is not None and value is not None:
                dates.append(day)
        if len(dates) < 2:
            return []
        granularity = self._granularity(dates)
        methods = channel.date_methods.get(date_column.index, Counter())
        jalali_count = sum(count for method, count in methods.items() if method.startswith("jalali"))
        use_jalali = jalali_count > len(dates) / 2
        calendar = "jalali" if use_jalali else "gregorian"
        points: dict[tuple[int, int, int], list[Decimal]] = defaultdict(list)
        for row in channel.rows:
            day = channel.date_map.get((date_column.key, row))
            value = channel.number_map.get((measure.key, row))
            if day is not None and value is not None:
                points[_period_key(day, granularity, jalali=use_jalali)].append(value)
        if len(points) < 2:
            return []
        ordered = sorted(points.items())
        trend_points: list[TrendPoint] = []
        for key, values in ordered:
            total = sum(values, ZERO)
            trend_points.append(
                TrendPoint(
                    period=_period_label(key, granularity, jalali=use_jalali),
                    period_label=_period_label(key, granularity, jalali=use_jalali),
                    count=len(values),
                    total=total,
                    mean=total / Decimal(len(values)),
                )
            )
        first, last = trend_points[0], trend_points[-1]
        change = last.mean - first.mean
        percent = (change / abs(first.mean) * HUNDRED).quantize(Decimal("0.01")) if first.mean != 0 else None
        if change > 0:
            direction = "افزایشی"
        elif change < 0:
            direction = "کاهشی"
        else:
            direction = "بدون تغییر"
        return [
            TrendAnalysis(
                date_column.key,
                date_column.header,
                measure.key,
                measure.header,
                granularity,
                trend_points,
                direction,
                change,
                percent,
                "روند توصیفی بر اساس میانگین هر دوره است و علت‌ومعلول را اثبات نمی‌کند.",
                calendar,
            )
        ]

    def _sequential_trend(
        self,
        channel: "_TableChannel",
        temporal_columns: list[ColumnStats],
        numeric_columns: list[ColumnStats],
    ) -> list[TrendAnalysis]:
        """Describe direction by observation order when dates are day keys, not full dates."""
        if not temporal_columns:
            return []
        temporal = max(temporal_columns, key=lambda column: (column.count, column.coverage))
        measure = max(
            numeric_columns,
            key=lambda column: (column.count, column.value_range or ZERO, column.distinct),
        )
        points: list[TrendPoint] = []
        for index, row in enumerate(channel.rows, start=1):
            value = channel.number_map.get((measure.key, row))
            if value is None:
                continue
            label = channel.display_map.get((temporal.key, row)) or str(index)
            points.append(
                TrendPoint(
                    period=label,
                    period_label=label,
                    count=1,
                    total=value,
                    mean=value,
                )
            )
        if len(points) < 2:
            return []
        first, last = points[0], points[-1]
        change = last.mean - first.mean
        percent = (change / abs(first.mean) * HUNDRED).quantize(Decimal("0.01")) if first.mean != 0 else None
        if change > 0:
            direction = "افزایشی"
        elif change < 0:
            direction = "کاهشی"
        else:
            direction = "بدون تغییر"
        return [
            TrendAnalysis(
                temporal.key,
                temporal.header,
                measure.key,
                measure.header,
                "observation",
                points,
                direction,
                change,
                percent,
                "تاریخ کامل قابل بازسازی نبود؛ روند فقط بر اساس ترتیب مشاهدات/روز ماه است و "
                "فاصله تقویمی واقعی یا رابطه علت‌ومعلولی را اثبات نمی‌کند.",
                "observation",
            )
        ]

    @staticmethod
    def _granularity(dates: list[date]) -> str:
        if not dates:
            return "month"
        months = {(day.year, day.month) for day in dates}
        if len(months) >= 2:
            if len(months) <= 36:
                return "month"
            years = {day.year for day in dates}
            return "quarter" if len(years) <= 10 else "year"
        return "day"

    # ------------------------------------------------------------------ #
    def _correlations(self, channel: "_TableChannel", numeric_columns: list[ColumnStats]) -> list[CorrelationPair]:
        candidates = [column for column in numeric_columns if column.count >= 3 and column.distinct > 1][:8]
        pairs: list[CorrelationPair] = []
        for i, first in enumerate(candidates):
            for second in candidates[i + 1:]:
                xs: list[Decimal] = []
                ys: list[Decimal] = []
                for row in channel.rows:
                    x = channel.number_map.get((first.key, row))
                    y = channel.number_map.get((second.key, row))
                    if x is None or y is None:
                        continue
                    xs.append(x)
                    ys.append(y)
                if len(xs) < 3:
                    continue
                pearson = _pearson(xs, ys)
                spearman = _spearman(xs, ys)
                if pearson is None and spearman is None:
                    continue
                value = pearson if pearson is not None else spearman
                strength = _strength(value)
                direction = "مستقیم" if value and value > 0 else "معکوس" if value and value < 0 else "بدون جهت"
                pairs.append(
                    CorrelationPair(
                        first.key,
                        first.header,
                        second.key,
                        second.header,
                        len(xs),
                        pearson,
                        spearman,
                        strength,
                        direction,
                        "همبستگی رابطه را نشان می‌دهد، نه علت‌ومعلول؛ مقادیر گمشده به‌صورت جفتی کنار گذاشته شده‌اند.",
                    )
                )
        pairs.sort(key=lambda item: abs(item.pearson or item.spearman or ZERO), reverse=True)
        return pairs[:10]

    # ------------------------------------------------------------------ #
    def _outliers(self, channel: "_TableChannel", numeric_columns: list[ColumnStats]) -> list[OutlierRecord]:
        results: list[OutlierRecord] = []
        for column in numeric_columns:
            values = channel.numbers[column.index]
            if len(values) < 8 or column.q1 is None or column.q3 is None or column.iqr is None:
                continue
            if column.iqr == 0:
                continue
            lower = column.q1 - Decimal("1.5") * column.iqr
            upper = column.q3 + Decimal("1.5") * column.iqr
            found = 0
            for row in channel.rows:
                value = channel.number_map.get((column.key, row))
                if value is None or lower <= value <= upper:
                    continue
                cell = channel.grid.get(row, column.index)
                results.append(
                    OutlierRecord(
                        column.key,
                        column.header,
                        row,
                        cell.address if cell is not None else f"{column.letter}{row}",
                        cell.value if cell is not None else str(value),
                        value,
                        "IQR 1.5",
                        lower,
                        upper,
                        ((value - column.median) / column.iqr).quantize(Decimal("0.01")) if column.median is not None else ZERO,
                    )
                )
                found += 1
                if found >= 20 or len(results) >= self.max_outliers:
                    break
            if len(results) >= self.max_outliers:
                break
        return results

    # ------------------------------------------------------------------ #
    def _quality(self, channel: "_TableChannel", columns: list[ColumnStats]) -> DataQuality:
        total_cells = len(channel.rows) * len(columns)
        missing = sum(column.missing for column in columns)
        seen: dict[tuple[str, ...], list[int]] = defaultdict(list)
        empty_rows = 0
        for row in channel.rows:
            row_values = tuple(channel.display_map.get((column.key, row), "") for column in columns)
            if not any(row_values):
                empty_rows += 1
                continue
            seen[row_values].append(row)
        duplicates = {tuple(values): rows for values, rows in seen.items() if len(rows) > 1}
        duplicate_examples = [
            f"ردیف‌های {', '.join(str(row) for row in rows[:5])}" for rows in list(duplicates.values())[:8]
        ]
        measured_columns = [column for column in columns if column.count]
        measured_cells = len(channel.rows) * len(measured_columns)
        measured_missing = sum(column.missing for column in measured_columns)
        return DataQuality(
            row_count=len(channel.rows),
            column_count=len(columns),
            total_cells=total_cells,
            missing_cells=missing,
            missing_ratio=missing / total_cells if total_cells else 0.0,
            empty_columns=[column.header for column in columns if not column.count],
            measured_cells=measured_cells,
            measured_missing_cells=measured_missing,
            measured_missing_ratio=measured_missing / measured_cells if measured_cells else 0.0,
            duplicate_rows=sum(len(rows) - 1 for rows in duplicates.values()),
            empty_rows=empty_rows,
            constant_columns=[column.header for column in columns if column.count and column.distinct == 1],
            mixed_type_columns=[
                column.header for column in columns if any("عدد" in warning and "متن" in warning for warning in column.warnings)
            ],
            identifier_columns=[column.header for column in columns if column.is_identifier],
            temporal_columns=[column.header for column in columns if column.is_temporal_key],
            invalid_cells=channel.errors + channel.numeric_invalid,
            duplicate_examples=duplicate_examples,
            notes=[
                "ردیف تکراری به‌صورت کامل و بر اساس همه ستون‌های جدول سنجیده شده است.",
                "هیچ ردیف یا مقدار پرتی حذف نشده است.",
                "ستون‌های شناسه و کلید زمانی از محاسبات آماری کنار گذاشته شده‌اند.",
            ],
        )

    @staticmethod
    def _methodology(
        channel: "_TableChannel",
        columns: list[ColumnStats],
        groups: list[GroupComparison],
        crosstabs: list[CrossTab],
        trends: list[TrendAnalysis],
        correlations: list[CorrelationPair],
        outliers: list[OutlierRecord],
    ) -> list[str]:
        lines = [
            "آماره‌های توصیفی روی همه مقادیر معتبر جدول محاسبه شده‌اند؛ تعداد مشاهدات هر ستون جداگانه ثبت است.",
            "واریانس و انحراف معیار هم برای جامعه (تقسیم بر n) و هم برای نمونه (تقسیم بر n-1) گزارش شده‌اند.",
            "میانه، چارک‌ها و صدک‌ها با درون‌یابی خطی (روش R-7) محاسبه شده‌اند.",
            "مقادیر گمشده در هر تحلیل به‌صورت جفتی کنار گذاشته شده‌اند و حذف فیزیکی انجام نشده است.",
            "داده‌های پرت با فنس‌های IQR (Q1 - 1.5×IQR و Q3 + 1.5×IQR) علامت‌گذاری شده‌اند و حذف نشده‌اند.",
            "همبستگی پیرسون و اسپیرمن گزارش شده است؛ همبستگی به‌معنای رابطه علت‌ومعلولی نیست.",
        ]
        if groups:
            lines.append("مقایسه گروه‌ها با میانگین، میانه، مجموع و انحراف معیار نمونه هر گروه انجام شده است.")
        if crosstabs:
            lines.append("جدول متقاطع فقط توصیفی است؛ برای استنباط نیاز به آزمون آماری و حجم نمونه کافی است.")
        if trends:
            if any(trend.calendar == "observation" for trend in trends):
                lines.append(
                    "روند زمانی بر اساس ترتیب مشاهدات/روز ماه ساخته شده است؛ "
                    "چون تاریخ کامل قابل بازسازی نبود، فاصله تقویمی واقعی مبنا قرار نگرفت."
                )
            else:
                lines.append("روند بر اساس دوره‌های زمانی ستون تاریخ ساخته شده و میانگین هر دوره مبنا است.")
        if outliers:
            lines.append("فهرست پرت‌ها آدرس سلول را نگه می‌دارد تا کاربر بتواند مقدار اصلی را بررسی کند.")
        if not any(column.is_numeric for column in columns):
            lines.append("ستون عددی معتبری وجود ندارد؛ تحلیل‌های عددی و همبستگی اجرا نشد.")
        return lines


class _TableChannel:
    """Read the table once and expose typed value maps by column."""

    def __init__(self, grid: SheetGrid, table: TableCandidate, max_rows: int) -> None:
        self.grid = grid
        self.values: dict[int, list[Any]] = {column.index: [] for column in table.columns}
        self.numbers: dict[int, list[Decimal]] = {column.index: [] for column in table.columns}
        self.dates: dict[int, list[date]] = {column.index: [] for column in table.columns}
        self.texts: dict[int, list[str]] = {column.index: [] for column in table.columns}
        self.kinds: dict[int, Counter[ValueKind]] = {column.index: Counter() for column in table.columns}
        self.value_map: dict[tuple[str, int], Any] = {}
        self.number_map: dict[tuple[str, int], Decimal] = {}
        self.date_map: dict[tuple[str, int], date] = {}
        self.date_methods: dict[int, Counter[str]] = {column.index: Counter() for column in table.columns}
        self.display_map: dict[tuple[str, int], str] = {}
        self.rows: list[int] = []
        self.errors = 0
        self.numeric_invalid = 0
        self.empty_rows = 0
        self.truncated = table.row_count > max_rows
        last_row = min(table.last_data_row, table.first_data_row + max_rows - 1)
        for row in range(table.first_data_row, last_row + 1):
            parsed_cells: list[tuple[Any, Any, str]] = []
            total_row = False
            row_errors = 0
            row_numeric_invalid = 0
            for column in table.columns:
                cell = grid.get(row, column.index)
                if cell is None or cell.is_empty:
                    continue
                parsed = parse_value(cell.value)
                if parsed.kind is ValueKind.ERROR:
                    row_errors += 1
                elif column.data_type is ValueKind.NUMBER and parsed.kind not in (ValueKind.NUMBER, ValueKind.EMPTY):
                    row_numeric_invalid += 1
                if parsed.kind is ValueKind.EMPTY:
                    continue
                key = f"{column.letter}:{column.header or column.letter}"
                if parsed.kind in (ValueKind.TEXT, ValueKind.BOOLEAN):
                    text = str(parsed.value).strip()
                    if text and _looks_like_total(text):
                        total_row = True
                parsed_cells.append((column, parsed, key))
            if total_row:
                continue
            if not parsed_cells:
                self.empty_rows += 1
                continue
            self.errors += row_errors
            self.numeric_invalid += row_numeric_invalid
            for column, parsed, key in parsed_cells:
                if parsed.kind is ValueKind.NUMBER and isinstance(parsed.value, Decimal):
                    self.numbers[column.index].append(parsed.value)
                    self.number_map[(key, row)] = parsed.value
                elif parsed.kind is ValueKind.DATE and isinstance(parsed.value, date):
                    self.dates[column.index].append(parsed.value)
                    self.date_methods[column.index][parsed.method] += 1
                    self.date_map[(key, row)] = parsed.value
                if parsed.kind in (ValueKind.TEXT, ValueKind.BOOLEAN):
                    text = str(parsed.value).strip()
                    if text:
                        self.texts[column.index].append(text)
                        self.value_map[(key, row)] = text
                self.values[column.index].append(parsed.value)
                self.display_map[(key, row)] = str(parsed.value)
                self.kinds[column.index][parsed.kind] += 1
            self.rows.append(row)

    def cell_address(self, row: int, column_index: int) -> str:
        cell = self.grid.get(row, column_index)
        return cell.address if cell is not None else ""


def _looks_like_total(text: str) -> bool:
    from local_office_ai.excel.normalization import normalize_digits, normalize_text

    value = normalize_text(normalize_digits(text)).lower().strip(" :.-")
    exact = {"جمع", "جمع کل", "مجموع", "مجموع کل", "کل", "total", "grand total", "sum"}
    return value in exact or value.startswith(("جمع ", "مجموع ", "جمع‌کل", "total ", "sum "))


def _percentile(values: list[Decimal], p: Decimal) -> Decimal:
    if not values:
        return ZERO
    if len(values) == 1:
        return values[0]
    rank = (Decimal(len(values) - 1) * p / HUNDRED)
    lower = int(rank)
    upper = min(len(values) - 1, lower + 1)
    fraction = rank - Decimal(lower)
    return values[lower] + (values[upper] - values[lower]) * fraction


def _sqrt(value: Decimal | None) -> Decimal | None:
    if value is None:
        return None
    if value < 0:
        return ZERO
    return value.sqrt()


def _std_sample(values: list[Decimal]) -> Decimal | None:
    if len(values) < 2:
        return None
    mean = sum(values, ZERO) / Decimal(len(values))
    variance = sum(((value - mean) ** 2 for value in values), ZERO) / Decimal(len(values) - 1)
    return _sqrt(variance)


def _pearson(xs: list[Decimal], ys: list[Decimal]) -> Decimal | None:
    n = len(xs)
    if n < 3:
        return None
    sx, sy = sum(xs, ZERO), sum(ys, ZERO)
    sxx, syy = sum((x * x for x in xs), ZERO), sum((y * y for y in ys), ZERO)
    sxy = sum((x * y for x, y in zip(xs, ys)), ZERO)
    numerator = Decimal(n) * sxy - sx * sy
    denominator_sq = (Decimal(n) * sxx - sx * sx) * (Decimal(n) * syy - sy * sy)
    if denominator_sq <= 0:
        return None
    return (numerator / _sqrt(denominator_sq)).quantize(Decimal("0.000001"))


def _spearman(xs: list[Decimal], ys: list[Decimal]) -> Decimal | None:
    if len(xs) < 3:
        return None
    return _pearson(_rank(xs), _rank(ys))


def _rank(values: list[Decimal]) -> list[Decimal]:
    ordered = sorted(range(len(values)), key=lambda index: values[index])
    ranks = [ZERO] * len(values)
    index = 0
    while index < len(ordered):
        end = index + 1
        while end < len(ordered) and values[ordered[end]] == values[ordered[index]]:
            end += 1
        average = Decimal(index + end + 1) / Decimal(2)
        for position in range(index, end):
            ranks[ordered[position]] = average
        index = end
    return ranks


def _strength(value: Decimal | None) -> str:
    if value is None:
        return "نامشخص"
    absolute = abs(value)
    if absolute >= Decimal("0.7"):
        return "قوی"
    if absolute >= Decimal("0.4"):
        return "متوسط"
    if absolute >= Decimal("0.2"):
        return "ضعیف"
    return "بسیار ضعیف"


def _period_key(day: date, granularity: str, *, jalali: bool = False) -> tuple[int, int, int]:
    if jalali:
        year, month, day_number = gregorian_to_jalali(day.year, day.month, day.day)
    else:
        year, month, day_number = day.year, day.month, day.day
    if granularity == "day":
        return (year, month, day_number)
    if granularity == "quarter":
        return (year, (month - 1) // 3 + 1, 0)
    if granularity == "year":
        return (year, 0, 0)
    return (year, month, 0)


def _period_label(key: tuple[int, int, int], granularity: str, *, jalali: bool = False) -> str:
    year, month, day = key
    if granularity == "day":
        source = date(year, month, day)
    elif granularity == "year":
        source = date(year, 1, 1)
    elif granularity == "quarter":
        source = date(year, (month - 1) * 3 + 1, 1)
    else:
        source = date(year, month, 1)
    if jalali:
        if granularity == "year":
            return f"{year}"
        if granularity == "quarter":
            return f"{year}/Q{month}"
        if granularity == "day":
            return f"{year}/{month:02d}/{day:02d}"
        return f"{year}/{month:02d}"
    if granularity == "day":
        return f"{year:04d}-{month:02d}-{day:02d}"
    if granularity == "quarter":
        return f"{year:04d}-Q{month}"
    if granularity == "year":
        return f"{year:04d}"
    return f"{year:04d}-{month:02d}"


def _d(value: Decimal | None) -> str | None:
    return None if value is None else str(value)


def _jsonable(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date,)):
        return value.isoformat()
    return value
