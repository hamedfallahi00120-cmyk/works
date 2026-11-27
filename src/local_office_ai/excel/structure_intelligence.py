"""Workbook-level structure intelligence.

This module answers the questions a generic engine must answer before it can
trust a workbook: where the real table starts and ends, which rows are headers,
data, totals, comparisons, percentages, notes or footers, which column belongs
to which parent header and period, whether metadata sources disagree, how much
of each period is actually observed, and whether declared totals reproduce the
daily data.

Nothing here is specific to one workbook.  Every rule is derived from the
layout, the labels, the periods and the cell contents.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable

from openpyxl.utils import column_index_from_string, get_column_letter

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import (
    AuditStatus,
    ColumnProfile,
    ColumnRole,
    DataRole,
    Severity,
    TableCandidate,
    ValueKind,
)
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    EXCEL_ERRORS,
    jalali_to_gregorian,
    looks_like_date,
    looks_like_number,
    normalize_digits,
    normalize_text,
    parse_value,
)
from local_office_ai.excel.period_detection import (
    KIND_MONTH,
    KIND_YEAR,
    SOURCE_BLOCK_TITLE,
    SOURCE_COLUMN_HEADER,
    SOURCE_FILENAME,
    SOURCE_ROW_LABEL,
    SOURCE_SHEET_NAME,
    STATUS_CONFLICTING,
    PeriodSource,
    detect_periods,
    scalar_conflicts,
)
from local_office_ai.excel.sheet_analyzer import (
    JALALI_MONTHS,
    HeaderBand,
    SheetGrid,
    extract_periods,
    is_total_label,
)

LOGGER = get_logger(__name__)

ZERO = Decimal("0")
HUNDRED = Decimal("100")

ROW_HEADER = "header"
ROW_SUBHEADER = "subheader"
ROW_DATA = "data"
ROW_DATA_MISSING = "data_missing"
ROW_TOTAL = "total"
ROW_DIFFERENCE = "difference"
ROW_PERCENTAGE = "percentage"
ROW_NOTE = "note"
ROW_FOOTER = "footer"
ROW_TITLE = "title"
ROW_BLANK = "blank"

STATUS_FULL = "fully_analyzed"
STATUS_PARTIAL = "partially_analyzed"
STATUS_NONE = "not_analyzed"

STATUS_FA = {
    STATUS_FULL: "کاملاً تحلیل شد",
    STATUS_PARTIAL: "جزئی تحلیل شد",
    STATUS_NONE: "تحلیل نشد",
}

ROW_ROLE_FA = {
    ROW_HEADER: "عنوان",
    ROW_SUBHEADER: "عنوان فرعی",
    ROW_DATA: "داده",
    ROW_DATA_MISSING: "دوره بدون مقدار",
    ROW_TOTAL: "جمع",
    ROW_DIFFERENCE: "مقایسه",
    ROW_PERCENTAGE: "درصد",
    ROW_NOTE: "توضیح",
    ROW_FOOTER: "پانویس",
    ROW_TITLE: "عنوان اصلی",
    ROW_BLANK: "خالی",
}

_DIFFERENCE_WORDS = ("مقایسه", "اختلاف", "تفاضل", "تغییر", "difference", "delta", "change", "variance")
_PERCENT_WORDS = ("درصد", "percentage", "percent", "ratio", "نرخ", "سهم")
_NOTE_WORDS = (
    "لازم به ذکر", "توضیح", "یادداشت", "منبع", "مأخذ", "ماخذ", "پانویس",
    "note", "notes", "source", "footnote", "comment",
)
_SUM_RANGE = re.compile(r"(?<![A-Z0-9])\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)")
_SHEET_REF = re.compile(r"(?:'([^']+)'|([^\s'!()=+*/-]+))!")

DAYS_IN_JALALI_MONTH = (31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29)


def _decimal(raw: Any) -> Decimal | None:
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, str) and raw.strip().upper() in EXCEL_ERRORS:
        return None
    parsed = parse_value(raw)
    if parsed.kind is ValueKind.NUMBER and isinstance(parsed.value, Decimal):
        return parsed.value
    if isinstance(parsed.value, (int, float)) and not isinstance(parsed.value, bool):
        return Decimal(str(parsed.value))
    return None


def _is_error(raw: Any) -> bool:
    return isinstance(raw, str) and raw.strip().upper() in EXCEL_ERRORS


def _text(raw: Any) -> str:
    if raw is None or isinstance(raw, bool):
        return ""
    return normalize_text(normalize_digits(str(raw))).strip()


def _jsonable(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dict):
        return {key: _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    return value


def month_from_text(*texts: Any) -> tuple[str, int | None]:
    for text in texts:
        normalized = _text(text)
        for index, month in enumerate(JALALI_MONTHS, start=1):
            if month in normalized:
                return month, index
    return "", None


def expected_days(month_index: int | None) -> int | None:
    if month_index is None:
        return None
    return DAYS_IN_JALALI_MONTH[month_index - 1]


def row_role_from_label(label: str, *, has_numbers: bool, long_text: bool = False) -> str:
    value = _text(label).lower().strip(" :.-")
    if not value:
        return ROW_BLANK
    if is_total_label(value):
        return ROW_TOTAL
    if any(word in value for word in _DIFFERENCE_WORDS):
        return ROW_DIFFERENCE
    if any(word in value for word in _PERCENT_WORDS):
        return ROW_PERCENTAGE
    if any(word in value for word in _NOTE_WORDS) or long_text:
        return ROW_NOTE
    if not has_numbers and len(value) > 60:
        return ROW_FOOTER
    return ROW_DATA


@dataclass
class MetadataConflict:
    sheet: str
    kind: str
    message: str
    values: dict[str, Any] = field(default_factory=dict)
    severity: str = Severity.WARNING.value
    confidence: float = 0.9

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "kind": self.kind,
            "message": self.message,
            "values": _jsonable(self.values),
            "severity": self.severity,
            "confidence": round(self.confidence, 3),
        }


@dataclass
class DataCoverage:
    sheet: str
    expected_observations: int | None = None
    observed_observations: int = 0
    missing_observations: int | None = None
    first_period: str = ""
    last_period: str = ""
    missing_labels: list[str] = field(default_factory=list)
    explicit_zero_cells: int = 0
    blank_cells: int = 0
    invalid_cells: int = 0
    error_cells: int = 0
    pending_rows: int = 0
    incomplete_period: bool = False
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "expected_observations": self.expected_observations,
            "observed_observations": self.observed_observations,
            "missing_observations": self.missing_observations,
            "first_period": self.first_period,
            "last_period": self.last_period,
            "missing_labels": list(self.missing_labels),
            "explicit_zero_cells": self.explicit_zero_cells,
            "blank_cells": self.blank_cells,
            "invalid_cells": self.invalid_cells,
            "error_cells": self.error_cells,
            "pending_rows": self.pending_rows,
            "incomplete_period": self.incomplete_period,
            "notes": list(self.notes),
        }


@dataclass
class TotalReconciliation:
    sheet: str
    kind: str
    address: str
    row: int
    column: int
    letter: str
    metric: str
    declared: Decimal | None
    computed: Decimal | None
    difference: Decimal | None
    status: AuditStatus
    data_range: str = ""
    formula: str | None = None
    formula_range: str = ""
    observations: int = 0
    explicit_zeros: int = 0
    blanks: int = 0
    cause: str = ""
    confidence: float = 0.9

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "kind": self.kind,
            "address": self.address,
            "row": self.row,
            "column": self.column,
            "letter": self.letter,
            "metric": self.metric,
            "declared": _jsonable(self.declared),
            "computed": _jsonable(self.computed),
            "difference": _jsonable(self.difference),
            "status": self.status.value,
            "data_range": self.data_range,
            "formula": self.formula,
            "formula_range": self.formula_range,
            "observations": self.observations,
            "explicit_zeros": self.explicit_zeros,
            "blanks": self.blanks,
            "cause": self.cause,
            "confidence": round(self.confidence, 3),
        }


@dataclass
class PeriodComparison:
    sheet: str
    metric: str
    period_a: str
    period_b: str
    total_a: Decimal | None
    total_b: Decimal | None
    count_a: int
    count_b: int
    mean_a: Decimal | None = None
    mean_b: Decimal | None = None
    difference: Decimal | None = None
    percent: Decimal | None = None
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "metric": self.metric,
            "period_a": self.period_a,
            "period_b": self.period_b,
            "total_a": _jsonable(self.total_a),
            "total_b": _jsonable(self.total_b),
            "count_a": self.count_a,
            "count_b": self.count_b,
            "mean_a": _jsonable(self.mean_a),
            "mean_b": _jsonable(self.mean_b),
            "difference": _jsonable(self.difference),
            "percent": _jsonable(self.percent),
            "note": self.note,
        }


@dataclass
class AggregationRecord:
    metric: str
    period: str
    total: Decimal | None
    observations: int
    sheets: list[str] = field(default_factory=list)
    months: list[str] = field(default_factory=list)
    note: str = ""
    overlaps: list[str] = field(default_factory=list)
    excluded_incomplete: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "metric": self.metric,
            "period": self.period,
            "total": _jsonable(self.total),
            "observations": self.observations,
            "sheets": list(self.sheets),
            "months": list(self.months),
            "note": self.note,
            "overlaps": list(self.overlaps),
            "excluded_incomplete": list(self.excluded_incomplete),
        }


@dataclass
class SheetIntelligence:
    sheet: str
    title: str = ""
    state: str = "visible"
    status: str = STATUS_NONE
    table_address: str = ""
    table_kind: str = ""
    confidence: float = 0.0
    header_rows: list[int] = field(default_factory=list)
    data_start: int = 0
    data_end: int = 0
    row_count: int = 0
    column_count: int = 0
    columns: list[ColumnProfile] = field(default_factory=list)
    month: str = ""
    month_index: int | None = None
    sheet_years: list[str] = field(default_factory=list)
    title_years: list[str] = field(default_factory=list)
    subheader_years: list[str] = field(default_factory=list)
    filename_years: list[str] = field(default_factory=list)
    periods: list[str] = field(default_factory=list)
    period_evidence: list[dict[str, Any]] = field(default_factory=list)
    period_sources: list[str] = field(default_factory=list)
    period_status: str = "not_detected"
    period_confidence: float = 0.0
    year_pairs: list[dict[str, Any]] = field(default_factory=list)
    derived: bool = False
    referenced_sheets: list[str] = field(default_factory=list)
    column_totals: list[dict[str, Any]] = field(default_factory=list)
    column_references: dict[str, list[str]] = field(default_factory=dict)
    row_roles: list[dict[str, Any]] = field(default_factory=list)
    summary_rows: dict[str, list[int]] = field(default_factory=dict)
    coverage: DataCoverage | None = None
    audits: list[TotalReconciliation] = field(default_factory=list)
    metadata_conflicts: list[MetadataConflict] = field(default_factory=list)
    period_comparisons: list[PeriodComparison] = field(default_factory=list)
    reconstructed_dates: list[dict[str, Any]] = field(default_factory=list)
    generic: Any = None
    warnings: list[str] = field(default_factory=list)

    @property
    def status_label(self) -> str:
        return STATUS_FA.get(self.status, self.status)

    @property
    def measures(self) -> list[ColumnProfile]:
        return [
            column
            for column in self.columns
            if column.statistical_role in {DataRole.NUMERIC_MEASURE, DataRole.CALCULATED}
            or (
                column.statistical_role is DataRole.UNKNOWN
                and column.data_type is ValueKind.NUMBER
                and column.role not in {ColumnRole.ID, ColumnRole.DATE}
            )
        ]

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "title": self.title,
            "state": self.state,
            "status": self.status,
            "status_label": self.status_label,
            "table_address": self.table_address,
            "table_kind": self.table_kind,
            "confidence": round(self.confidence, 4),
            "header_rows": list(self.header_rows),
            "data_start": self.data_start,
            "data_end": self.data_end,
            "row_count": self.row_count,
            "column_count": self.column_count,
            "columns": [column.to_dict() for column in self.columns],
            "month": self.month,
            "month_index": self.month_index,
            "sheet_years": list(self.sheet_years),
            "title_years": list(self.title_years),
            "subheader_years": list(self.subheader_years),
            "filename_years": list(self.filename_years),
            "periods": list(self.periods),
            "period_evidence": list(self.period_evidence),
            "period_sources": list(self.period_sources),
            "period_status": self.period_status,
            "period_confidence": round(self.period_confidence, 3),
            "year_pairs": list(self.year_pairs),
            "derived": self.derived,
            "referenced_sheets": list(self.referenced_sheets),
            "column_totals": _jsonable(self.column_totals),
            "column_references": {key: list(value) for key, value in self.column_references.items()},
            "row_roles": list(self.row_roles),
            "summary_rows": {key: list(value) for key, value in self.summary_rows.items()},
            "coverage": self.coverage.to_dict() if self.coverage else None,
            "audits": [item.to_dict() for item in self.audits],
            "metadata_conflicts": [item.to_dict() for item in self.metadata_conflicts],
            "period_comparisons": [item.to_dict() for item in self.period_comparisons],
            "reconstructed_dates": list(self.reconstructed_dates),
            "generic": self.generic.to_dict() if hasattr(self.generic, "to_dict") else self.generic,
            "warnings": list(self.warnings),
        }


@dataclass
class WorkbookIntelligence:
    path: str
    sheet_count: int = 0
    sheets: list[SheetIntelligence] = field(default_factory=list)
    similar_groups: list[dict[str, Any]] = field(default_factory=list)
    metadata_conflicts: list[MetadataConflict] = field(default_factory=list)
    audits: list[TotalReconciliation] = field(default_factory=list)
    cross_sheet_audits: list[TotalReconciliation] = field(default_factory=list)
    period_comparisons: list[PeriodComparison] = field(default_factory=list)
    aggregations: list[AggregationRecord] = field(default_factory=list)
    status_counts: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    methodology: list[str] = field(default_factory=list)

    def sheet(self, name: str) -> SheetIntelligence | None:
        for item in self.sheets:
            if item.sheet == name:
                return item
        return None

    @property
    def all_audits(self) -> list[TotalReconciliation]:
        return [*self.audits, *self.cross_sheet_audits]

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "sheet_count": self.sheet_count,
            "sheets": [item.to_dict() for item in self.sheets],
            "similar_groups": list(self.similar_groups),
            "metadata_conflicts": [item.to_dict() for item in self.metadata_conflicts],
            "audits": [item.to_dict() for item in self.audits],
            "cross_sheet_audits": [item.to_dict() for item in self.cross_sheet_audits],
            "period_comparisons": [item.to_dict() for item in self.period_comparisons],
            "aggregations": [item.to_dict() for item in self.aggregations],
            "status_counts": dict(self.status_counts),
            "warnings": list(self.warnings),
            "methodology": list(self.methodology),
        }

class StructureIntelligenceAnalyzer:
    """Inspect every sheet and reconcile the workbook as a whole."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def inspect_sheet(
        self,
        grid: SheetGrid,
        sheet_meta: Any,
        tables: list[TableCandidate],
        primary: TableCandidate | None,
        *,
        generic: Any = None,
        known_sheets: set[str] | None = None,
        source_name: str = "",
    ) -> SheetIntelligence:
        intel = SheetIntelligence(
            sheet=grid.sheet_name,
            state=getattr(sheet_meta, "state", "visible") or "visible",
            generic=generic,
        )
        if primary is None:
            intel.status = STATUS_NONE
            intel.warnings.append("هیچ جدول داده‌ای در این برگه شناسایی نشد.")
            intel.row_roles = self._fallback_roles(grid)
            self._apply_periods(
                intel,
                detect_periods(self._scalar_period_sources(grid.sheet_name, "", source_name)),
            )
            return intel

        header_rows = primary.header_rows or [primary.header_row]
        intel.table_address = primary.address
        intel.table_kind = primary.kind.value
        intel.confidence = primary.confidence
        intel.header_rows = list(header_rows)
        intel.data_start = primary.first_data_row
        intel.data_end = primary.last_data_row
        intel.row_count = primary.row_count
        intel.column_count = len(primary.columns)
        intel.columns = list(primary.columns)

        title_rows = [
            row for row in range(1, header_rows[0]) if self._row_has_content(grid, row)
        ]
        intel.title = self._title_text(grid, title_rows)
        intel.month, intel.month_index = month_from_text(grid.sheet_name, intel.title)
        intel.sheet_years = extract_periods([grid.sheet_name])
        intel.title_years = extract_periods([intel.title])
        intel.subheader_years = self._unique_periods(primary.columns)
        intel.filename_years = extract_periods([source_name]) if source_name else []
        intel.year_pairs = self._build_pairs(primary.columns)
        self._apply_periods(
            intel,
            detect_periods(self._period_sources(grid, primary, intel.title, source_name)),
        )
        if not intel.month and source_name:
            filename_month, filename_index = month_from_text(source_name)
            if filename_month:
                intel.month, intel.month_index = filename_month, filename_index
                intel.warnings.append(
                    "ماه از نام فایل استنباط شد؛ نام برگه و عنوان هیچ ماهی نداشتند."
                )

        intel.row_roles, intel.summary_rows = self._row_roles(grid, primary, header_rows)
        intel.derived, intel.referenced_sheets = self._references(
            grid, primary, known_sheets or {grid.sheet_name}
        )
        intel.column_totals, intel.column_references = self._column_totals(
            grid, primary, known_sheets or {grid.sheet_name}
        )
        intel.coverage = self._coverage(grid, primary, intel)
        intel.metadata_conflicts = self._metadata_conflicts(intel)
        intel.period_comparisons = self._period_comparisons(grid, primary, intel)
        intel.reconstructed_dates = self._reconstructed_dates(grid, primary, intel)
        intel.audits = self._audits(grid, primary, intel)
        intel.warnings.extend(self._sheet_warnings(intel))
        if intel.coverage and intel.coverage.incomplete_period:
            intel.status = STATUS_PARTIAL
        elif intel.measures and intel.row_count > 0:
            intel.status = STATUS_FULL
        elif intel.row_count > 0:
            intel.status = STATUS_PARTIAL
        else:
            intel.status = STATUS_NONE
        return intel

    # ------------------------------------------------------------------ #
    @staticmethod
    def _row_has_content(grid: SheetGrid, row: int) -> bool:
        for col in range(1, grid.max_col + 1):
            cell = grid.get(row, col)
            if cell is not None and (not cell.is_empty or cell.formula):
                return True
        return False

    @staticmethod
    def _title_text(grid: SheetGrid, rows: Iterable[int]) -> str:
        parts: list[str] = []
        for row in rows:
            cells = []
            for col in range(1, grid.max_col + 1):
                cell = grid.get(row, col)
                if cell is not None and not cell.is_empty:
                    value = _text(cell.value)
                    if value:
                        cells.append(value)
            if cells:
                parts.append(" | ".join(cells))
        return " | ".join(parts)[:500]

    # ------------------------------------------------------------------ #
    # Period evidence
    # ------------------------------------------------------------------ #
    @staticmethod
    def _scalar_period_sources(
        sheet_name: str,
        title: str,
        source_name: str,
        title_address: str = "",
    ) -> list[PeriodSource]:
        sources = [PeriodSource(SOURCE_SHEET_NAME, sheet_name, address=sheet_name)]
        if title:
            sources.append(PeriodSource(SOURCE_BLOCK_TITLE, title, address=title_address))
        if source_name:
            sources.append(PeriodSource(SOURCE_FILENAME, source_name, address=source_name))
        return sources

    @staticmethod
    def _title_address(grid: SheetGrid, header_row: int, title: str) -> str:
        target = normalize_text(normalize_digits(title))
        if not target:
            return ""
        for row in range(1, header_row):
            for col in range(1, grid.max_col + 1):
                cell = grid.get(row, col)
                if cell is None or not isinstance(cell.value, str):
                    continue
                value = normalize_text(normalize_digits(cell.value))
                if value and value in target:
                    return f"{get_column_letter(col)}{row}"
        return ""

    def _period_sources(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        title: str,
        source_name: str,
    ) -> list[PeriodSource]:
        header_row = min(table.header_rows or [table.header_row])
        sources = self._scalar_period_sources(
            grid.sheet_name,
            title,
            source_name,
            self._title_address(grid, header_row, title),
        )
        for column in table.columns:
            text = " ".join(
                part for part in (column.parent_header, column.sub_header, column.period) if part
            )
            if text:
                sources.append(
                    PeriodSource(SOURCE_COLUMN_HEADER, text, address=column.header_address)
                )
        for row in range(
            table.first_data_row,
            min(table.last_data_row, table.first_data_row + 499) + 1,
        ):
            label = self._row_label(grid, row, table.first_col, table.last_col)
            if label and not looks_like_date(label) and not looks_like_number(label):
                sources.append(
                    PeriodSource(
                        SOURCE_ROW_LABEL,
                        label,
                        address=self._row_label_address(
                            grid, row, table.first_col, table.last_col, label
                        ),
                    )
                )
        return sources

    @staticmethod
    def _row_label_address(
        grid: SheetGrid,
        row: int,
        first_col: int,
        last_col: int,
        label: str,
    ) -> str:
        target = normalize_text(normalize_digits(label))
        for col in range(first_col, last_col + 1):
            cell = grid.get(row, col)
            if (
                cell is not None
                and isinstance(cell.value, str)
                and normalize_text(normalize_digits(cell.value)) == target
            ):
                return f"{get_column_letter(col)}{row}"
        return f"{get_column_letter(first_col)}{row}"

    @staticmethod
    def _apply_periods(intel: SheetIntelligence, detection: Any) -> None:
        intel.periods = [item.token for item in detection.periods]
        intel.period_evidence = [item.to_dict() for item in detection.evidence]
        intel.period_sources = detection.sources
        intel.period_status = detection.status
        intel.period_confidence = detection.confidence
        conflicts = scalar_conflicts(detection, dimension=KIND_YEAR)
        conflicts.extend(scalar_conflicts(detection, dimension=KIND_MONTH))
        if conflicts:
            intel.period_status = STATUS_CONFLICTING

    @staticmethod
    def _unique_periods(columns: list[ColumnProfile]) -> list[str]:
        periods: list[str] = []
        for column in columns:
            for token in extract_periods([column.period, column.sub_header]):
                if token not in periods:
                    periods.append(token)
        return periods

    @staticmethod
    def _build_pairs(columns: list[ColumnProfile]) -> list[dict[str, Any]]:
        grouped: dict[str, list[ColumnProfile]] = defaultdict(list)
        for column in columns:
            parent = column.parent_header or column.header or column.letter
            grouped[parent].append(column)
        pairs: list[dict[str, Any]] = []
        for parent, group in grouped.items():
            ordered = sorted(group, key=lambda column: column.index)
            for left, right in zip(ordered, ordered[1:]):
                if not left.period or not right.period or left.period == right.period:
                    continue
                pairs.append(
                    {
                        "metric": parent,
                        "period_a": left.period,
                        "period_b": right.period,
                        "column_a": left.letter,
                        "column_b": right.letter,
                        "index_a": left.index,
                        "index_b": right.index,
                    }
                )
        return pairs

    def _row_roles(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        header_rows: list[int],
    ) -> tuple[list[dict[str, Any]], dict[str, list[int]]]:
        roles: list[dict[str, Any]] = []
        summary: dict[str, list[int]] = defaultdict(list)
        measure_cols = [
            column
            for column in table.columns
            if column.role not in {ColumnRole.ID, ColumnRole.DATE}
        ]
        last_summary = 0
        provisional: dict[int, str] = {}
        labels: dict[int, str] = {}
        for row in range(1, grid.max_row + 1):
            label = self._row_label(grid, row, table.first_col, table.last_col)
            labels[row] = label
            has_number = self._row_has_numeric(grid, row, measure_cols)
            has_any_number = self._row_has_numeric(grid, row, table.columns)
            content = self._row_has_content(grid, row)
            if not content:
                role = ROW_BLANK
            elif row in header_rows:
                role = ROW_HEADER if row == header_rows[0] else ROW_SUBHEADER
            elif row < table.first_data_row:
                role = ROW_TITLE if label else ROW_BLANK
            elif table.first_data_row <= row <= table.last_data_row:
                role = ROW_DATA if has_number else ROW_DATA_MISSING
            else:
                long_text = len(label) > 60
                role = row_role_from_label(label, has_numbers=has_any_number, long_text=long_text)
                if role == ROW_BLANK and has_any_number:
                    role = ROW_DATA_MISSING
                if role == ROW_DATA:
                    role = ROW_NOTE if label else ROW_BLANK
            provisional[row] = role
            if role in {ROW_TOTAL, ROW_DIFFERENCE, ROW_PERCENTAGE}:
                summary[role].append(row)
                last_summary = max(last_summary, row)

        for row in range(1, grid.max_row + 1):
            role = provisional[row]
            if role == ROW_NOTE and row > last_summary:
                role = ROW_FOOTER
            roles.append(
                {
                    "row": row,
                    "role": role,
                    "role_label": ROW_ROLE_FA.get(role, role),
                    "label": labels.get(row, "")[:160],
                }
            )
        return roles, {key: sorted(value) for key, value in summary.items()}

    @staticmethod
    def _fallback_roles(grid: SheetGrid) -> list[dict[str, Any]]:
        roles = []
        for row in range(1, grid.max_row + 1):
            has = any(
                (cell := grid.get(row, col)) is not None and not cell.is_empty
                for col in range(1, grid.max_col + 1)
            )
            role = ROW_DATA if has else ROW_BLANK
            roles.append(
                {
                    "row": row,
                    "role": role,
                    "role_label": ROW_ROLE_FA[role],
                    "label": "",
                }
            )
        return roles

    @staticmethod
    def _row_label(grid: SheetGrid, row: int, first_col: int, last_col: int) -> str:
        for col in range(first_col, min(last_col, first_col + 5) + 1):
            cell = grid.get(row, col)
            if cell is None or cell.is_empty:
                continue
            value = cell.value
            if isinstance(value, str):
                if parse_value(value).kind is not ValueKind.NUMBER and not _is_error(value):
                    text = _text(value)
                    if text:
                        return text
        for col in range(first_col, last_col + 1):
            cell = grid.get(row, col)
            if cell is not None and isinstance(cell.value, str) and cell.value.strip():
                if parse_value(cell.value).kind is ValueKind.TEXT:
                    return _text(cell.value)
        return ""

    @staticmethod
    def _row_has_numeric(grid: SheetGrid, row: int, columns: Iterable[ColumnProfile]) -> bool:
        for column in columns:
            cell = grid.get(row, column.index)
            if cell is None or cell.is_empty:
                continue
            if parse_value(cell.value).kind is ValueKind.NUMBER:
                return True
        return False

    @staticmethod
    def _column_values(
        grid: SheetGrid,
        column: ColumnProfile,
        first_row: int,
        last_row: int,
    ) -> tuple[list[Decimal], int, int, int, int]:
        values: list[Decimal] = []
        zeros = blanks = invalid = errors = 0
        for row in range(first_row, last_row + 1):
            cell = grid.get(row, column.index)
            if cell is None or cell.is_empty:
                blanks += 1
                continue
            parsed = parse_value(cell.value)
            if parsed.kind is ValueKind.ERROR:
                errors += 1
            elif parsed.kind is ValueKind.NUMBER:
                number = _decimal(parsed.value)
                if number is None:
                    invalid += 1
                else:
                    values.append(number)
                    if number == ZERO:
                        zeros += 1
            else:
                invalid += 1
        return values, zeros, blanks, invalid, errors

    def _coverage(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        intel: SheetIntelligence,
    ) -> DataCoverage:
        coverage = DataCoverage(sheet=grid.sheet_name)
        temporal = next(
            (
                column
                for column in table.columns
                if column.statistical_role is DataRole.TEMPORAL_KEY
            ),
            None,
        )
        measure_cols = intel.measures
        observed_rows: list[int] = []
        for row in range(table.first_data_row, table.last_data_row + 1):
            if self._row_has_numeric(grid, row, measure_cols):
                observed_rows.append(row)
        observed_labels: list[str] = []
        labels_are_days = False
        if temporal is not None:
            for row in observed_rows:
                cell = grid.get(row, temporal.index)
                if cell is None:
                    continue
                value = _decimal(cell.value)
                if value is not None:
                    observed_labels.append(str(int(value)))
            labels_are_days = bool(observed_labels)
        if not observed_labels:
            observed_labels = [str(row) for row in observed_rows]
        coverage.observed_observations = len(observed_rows)
        if observed_labels:
            coverage.first_period = observed_labels[0]
            coverage.last_period = observed_labels[-1]
        # Day-by-day coverage is only meaningful when the date column really
        # holds day numbers; row numbers must never be read as days.
        expected = (
            expected_days(intel.month_index)
            if not intel.derived and labels_are_days
            else None
        )
        if expected is not None and temporal is not None:
            coverage.expected_observations = expected
            observed_set = set(observed_labels)
            coverage.missing_labels = [
                str(day) for day in range(1, expected + 1) if str(day) not in observed_set
            ]
            coverage.missing_observations = len(coverage.missing_labels)
            coverage.incomplete_period = bool(coverage.missing_labels)
        else:
            coverage.missing_observations = None
            coverage.notes.append(
                "دوره کامل تقویمی قابل تعیین نبود (نام ماه یا بازه زمانی یکتا در برگه پیدا نشد)؛ "
                "پوشش فقط بر اساس ردیف‌های مشاهده‌شده گزارش می‌شود."
            )
        pending = 0
        for row in range(table.last_data_row + 1, grid.max_row + 1):
            role = next(
                (item["role"] for item in intel.row_roles if item["row"] == row),
                ROW_BLANK,
            )
            if role == ROW_DATA_MISSING:
                pending += 1
                if expected is not None and temporal is not None and labels_are_days:
                    value = _decimal(grid.value(row, temporal.index))
                    if value is not None and str(int(value)) not in coverage.missing_labels:
                        coverage.missing_labels.append(str(int(value)))
        coverage.pending_rows = pending
        if coverage.missing_labels and coverage.missing_observations is not None:
            coverage.missing_observations = len(coverage.missing_labels)
        for column in measure_cols:
            values, zeros, blanks, invalid, errors = self._column_values(
                grid, column, table.first_data_row, table.last_data_row
            )
            coverage.explicit_zero_cells += zeros
            coverage.blank_cells += blanks
            coverage.invalid_cells += invalid
            coverage.error_cells += errors
        coverage.notes.append(
            "صفر واقعی، سلول خالی، مقدار نامعتبر و خطای فرمول جداگانه شمرده شده‌اند؛ خالی به صفر تبدیل نشد."
        )
        return coverage

    def _metadata_conflicts(self, intel: SheetIntelligence) -> list[MetadataConflict]:
        conflicts: list[MetadataConflict] = []
        sheet_years = set(intel.sheet_years)
        title_years = set(intel.title_years)
        sub_years = set(intel.subheader_years)
        filename_years = set(intel.filename_years)
        if sheet_years and title_years and not (sheet_years & title_years):
            conflicts.append(
                MetadataConflict(
                    sheet=intel.sheet,
                    kind="sheet_name_vs_title_year",
                    message=(
                        f"سال نام برگه {sorted(sheet_years)} با سال عنوان اصلی {sorted(title_years)} "
                        "یکسان نیست؛ هیچ منبعی بر دیگری ترجیح داده نشد."
                    ),
                    values={
                        "sheet_name": intel.sheet,
                        "sheet_years": sorted(sheet_years),
                        "title": intel.title,
                        "title_years": sorted(title_years),
                    },
                )
            )
        if sub_years and sheet_years and not (sheet_years & sub_years):
            conflicts.append(
                MetadataConflict(
                    sheet=intel.sheet,
                    kind="sheet_name_vs_subheader_years",
                    message=(
                        f"سال نام برگه {sorted(sheet_years)} در زیرعنوان‌های ستون‌ها {sorted(sub_years)} نیست."
                    ),
                    values={"sheet_years": sorted(sheet_years), "subheader_years": sorted(sub_years)},
                )
            )
        if sub_years and title_years and not (title_years & sub_years):
            conflicts.append(
                MetadataConflict(
                    sheet=intel.sheet,
                    kind="title_vs_subheader_years",
                    message=(
                        f"سال عنوان اصلی {sorted(title_years)} در زیرعنوان‌های ستون‌ها {sorted(sub_years)} نیست."
                    ),
                    values={
                        "title": intel.title,
                        "title_years": sorted(title_years),
                        "subheader_years": sorted(sub_years),
                    },
                )
            )
        if len(title_years) > 1:
            conflicts.append(
                MetadataConflict(
                    sheet=intel.sheet,
                    kind="multiple_title_years",
                    message=f"عنوان اصلی چند سال مختلف را نام می‌برد: {sorted(title_years)}.",
                    values={"title_years": sorted(title_years)},
                )
            )
        if filename_years and sheet_years and not (filename_years & sheet_years):
            conflicts.append(
                MetadataConflict(
                    sheet=intel.sheet,
                    kind="filename_vs_sheet_name_year",
                    message=(
                        f"سال نام فایل {sorted(filename_years)} با سال نام برگه {sorted(sheet_years)} "
                        "یکسان نیست؛ هیچ منبعی بر دیگری ترجیح داده نشد."
                    ),
                    values={"filename_years": sorted(filename_years), "sheet_years": sorted(sheet_years)},
                )
            )
        if filename_years and title_years and not (filename_years & title_years):
            conflicts.append(
                MetadataConflict(
                    sheet=intel.sheet,
                    kind="filename_vs_title_year",
                    message=(
                        f"سال نام فایل {sorted(filename_years)} با سال عنوان اصلی {sorted(title_years)} "
                        "یکسان نیست."
                    ),
                    values={"filename_years": sorted(filename_years), "title_years": sorted(title_years)},
                )
            )
        return conflicts

    def _period_comparisons(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        intel: SheetIntelligence,
    ) -> list[PeriodComparison]:
        by_index = {column.index: column for column in table.columns}
        comparisons: list[PeriodComparison] = []
        for pair in intel.year_pairs:
            left = by_index.get(pair["index_a"])
            right = by_index.get(pair["index_b"])
            if left is None or right is None:
                continue
            values_a, _, _, _, _ = self._column_values(
                grid, left, table.first_data_row, table.last_data_row
            )
            values_b, _, _, _, _ = self._column_values(
                grid, right, table.first_data_row, table.last_data_row
            )
            if not values_a and not values_b:
                continue
            total_a = sum(values_a, ZERO) if values_a else None
            total_b = sum(values_b, ZERO) if values_b else None
            difference = total_b - total_a if total_a is not None and total_b is not None else None
            percent = None
            if difference is not None and total_a not in (None, ZERO):
                percent = (difference / abs(total_a) * HUNDRED).quantize(Decimal("0.01"))
            comparisons.append(
                PeriodComparison(
                    sheet=grid.sheet_name,
                    metric=pair["metric"],
                    period_a=pair["period_a"],
                    period_b=pair["period_b"],
                    total_a=total_a,
                    total_b=total_b,
                    count_a=len(values_a),
                    count_b=len(values_b),
                    mean_a=(total_a / Decimal(len(values_a))) if values_a and total_a is not None else None,
                    mean_b=(total_b / Decimal(len(values_b))) if values_b and total_b is not None else None,
                    difference=difference,
                    percent=percent,
                    note=(
                        "مقایسه فقط برای ستون‌هایی انجام شد که تعریف شاخص و دوره آن‌ها هم‌تراز بود؛ "
                        "اگر یک طرف مشاهده نداشته باشد، تفاوت محاسبه نمی‌شود."
                    ),
                )
            )
        return comparisons

    def _reconstructed_dates(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        intel: SheetIntelligence,
    ) -> list[dict[str, Any]]:
        if intel.month_index is None:
            return []
        temporal = next(
            (column for column in table.columns if column.statistical_role is DataRole.TEMPORAL_KEY),
            None,
        )
        if temporal is None:
            return []
        days = []
        for row in range(table.first_data_row, table.last_data_row + 1):
            value = _decimal(grid.value(row, temporal.index))
            if value is not None and float(value).is_integer():
                days.append(int(value))
        if not days:
            return []
        results: list[dict[str, Any]] = []
        for period in intel.subheader_years:
            if not period.isdigit():
                continue
            year = int(period)
            try:
                first = min(days)
                last = max(days)
                gy, gm, gd = jalali_to_gregorian(year, intel.month_index, first)
                first_gregorian = date(gy, gm, gd).isoformat()
                gy, gm, gd = jalali_to_gregorian(year, intel.month_index, last)
                last_gregorian = date(gy, gm, gd).isoformat()
            except (ValueError, OverflowError):
                continue
            results.append(
                {
                    "period": period,
                    "month": intel.month,
                    "first_day": first,
                    "last_day": last,
                    "first_jalali": f"{year}/{intel.month_index:02d}/{first:02d}",
                    "last_jalali": f"{year}/{intel.month_index:02d}/{last:02d}",
                    "first_gregorian": first_gregorian,
                    "last_gregorian": last_gregorian,
                    "confidence": 0.75 if intel.metadata_conflicts else 0.9,
                    "source": "month from title/sheet name + year from subheader + day key",
                    "note": (
                        "تاریخ فقط از منابع صریح بازسازی شد. در صورت تعارض سال، تاریخ به‌عنوان "
                        "کاندید ثبت شده و مبنای محاسبه قرار نگرفته است."
                    ),
                }
            )
        return results

    # ------------------------------------------------------------------ #
    def _references(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        known_sheets: set[str],
    ) -> tuple[bool, list[str]]:
        references: list[str] = []
        for column in table.columns:
            for row in range(table.first_data_row, table.last_data_row + 1):
                cell = grid.get(row, column.index)
                if cell is None or not cell.formula:
                    continue
                for match in _SHEET_REF.finditer(cell.formula):
                    name = (match.group(1) or match.group(2) or "").strip()
                    if name in known_sheets and name != grid.sheet_name and name not in references:
                        references.append(name)
        return bool(references), references

    def _column_totals(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        known_sheets: set[str],
    ) -> tuple[list[dict[str, Any]], dict[str, list[str]]]:
        totals: list[dict[str, Any]] = []
        references: dict[str, list[str]] = defaultdict(list)
        for column in table.columns:
            if column.role in {ColumnRole.ID, ColumnRole.DATE}:
                continue
            metric = column.parent_header or column.header or column.letter
            period = column.period
            values, _, _, _, _ = self._column_values(
                grid, column, table.first_data_row, table.last_data_row
            )
            total = sum(values, ZERO) if values else None
            totals.append(
                {
                    "metric": metric,
                    "period": period,
                    "letter": column.letter,
                    "total": total,
                    "count": len(values),
                    "statistical_role": column.statistical_role.value,
                }
            )
            if period:
                key = f"{metric} | {period}"
                for row in range(table.first_data_row, table.last_data_row + 1):
                    cell = grid.get(row, column.index)
                    if cell is None or not cell.formula:
                        continue
                    for match in _SHEET_REF.finditer(cell.formula):
                        name = (match.group(1) or match.group(2) or "").strip()
                        if (
                            name in known_sheets
                            and name != grid.sheet_name
                            and name not in references[key]
                        ):
                            references[key].append(name)
        return totals, dict(references)

    # ------------------------------------------------------------------ #
    def _audits(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        intel: SheetIntelligence,
    ) -> list[TotalReconciliation]:
        audits: list[TotalReconciliation] = []
        total_rows = intel.summary_rows.get(ROW_TOTAL, [])
        diff_rows = intel.summary_rows.get(ROW_DIFFERENCE, [])
        percent_rows = intel.summary_rows.get(ROW_PERCENTAGE, [])
        by_index = {column.index: column for column in table.columns}
        total_map: dict[int, Decimal | None] = {}

        for total_row in total_rows:
            for col in range(table.first_col, table.last_col + 1):
                column = by_index.get(col)
                if column is not None and column.role in {ColumnRole.ID, ColumnRole.DATE}:
                    continue
                cell = grid.get(total_row, col)
                if cell is None or cell.is_empty:
                    continue
                raw = cell.value
                metric = (column.header if column is not None else "") or get_column_letter(col)
                if _is_error(raw):
                    audits.append(
                        TotalReconciliation(
                            sheet=grid.sheet_name,
                            kind="total",
                            address=f"{get_column_letter(col)}{total_row}",
                            row=total_row,
                            column=col,
                            letter=get_column_letter(col),
                            metric=metric,
                            declared=None,
                            computed=None,
                            difference=None,
                            status=AuditStatus.NOT_VERIFIABLE,
                            formula=cell.formula,
                            cause=f"خطای فرمول: {raw}",
                        )
                    )
                    total_map[col] = None
                    continue
                declared = _decimal(raw)
                if declared is None:
                    continue
                values, zeros, blanks, invalid, errors = self._column_values(
                    grid, column, table.first_data_row, table.last_data_row
                )
                computed = sum(values, ZERO) if values else None
                difference = declared - computed if computed is not None else None
                formula_range = ""
                formula = cell.formula
                formula_note = ""
                if formula:
                    match = _SUM_RANGE.search(formula)
                    if match:
                        formula_range = (
                            f"{match.group(1)}{match.group(2)}:{match.group(3)}{match.group(4)}"
                        )
                        if (
                            int(match.group(2)) > table.first_data_row
                            or int(match.group(4)) < table.last_data_row
                        ):
                            formula_note = "محدوده فرمول تمام ردیف‌های داده را پوشش نمی‌دهد"
                cause = ""
                status = AuditStatus.VERIFIED
                if computed is None:
                    status = AuditStatus.NOT_VERIFIABLE
                    cause = "داده‌ای برای جمع وجود ندارد؛ سلول خالی به صفر تبدیل نشد."
                elif declared != computed:
                    status = AuditStatus.MISMATCH
                    if formula_note:
                        cause = formula_note
                    elif not formula and declared == ZERO:
                        cause = "مقدار ثابت صفر بدون فرمول، در حالی که داده روزانه مقدار دارد."
                    elif self._matches_partial_sum(declared, values):
                        cause = "مقدار ثبت‌شده با جمع زیرمجموعه‌ای از ردیف‌ها برابر است (جمع ناقص)."
                    else:
                        cause = "مقدار ثبت‌شده با مجموع داده‌های روزانه برابر نیست."
                audits.append(
                    TotalReconciliation(
                        sheet=grid.sheet_name,
                        kind="total",
                        address=f"{get_column_letter(col)}{total_row}",
                        row=total_row,
                        column=col,
                        letter=get_column_letter(col),
                        metric=metric,
                        declared=declared,
                        computed=computed,
                        difference=difference,
                        status=status,
                        data_range=(
                            f"{get_column_letter(col)}{table.first_data_row}:"
                            f"{get_column_letter(col)}{table.last_data_row}"
                        ),
                        formula=formula,
                        formula_range=formula_range,
                        observations=len(values),
                        explicit_zeros=zeros,
                        blanks=blanks,
                        cause=cause,
                    )
                )
                total_map[col] = declared

        for row_kind, rows in ((ROW_DIFFERENCE, diff_rows), (ROW_PERCENTAGE, percent_rows)):
            for row in rows:
                for col in range(table.first_col, table.last_col + 1):
                    cell = grid.get(row, col)
                    if cell is None or not _is_error(cell.value):
                        continue
                    column = by_index.get(col)
                    audits.append(
                        TotalReconciliation(
                            sheet=grid.sheet_name,
                            kind=row_kind,
                            address=f"{get_column_letter(col)}{row}",
                            row=row,
                            column=col,
                            letter=get_column_letter(col),
                            metric=(column.header if column is not None else get_column_letter(col)),
                            declared=None,
                            computed=None,
                            difference=None,
                            status=AuditStatus.NOT_VERIFIABLE,
                            formula=cell.formula,
                            cause=(
                                "خطای تقسیم بر صفر در ردیف درصد"
                                if row_kind == ROW_PERCENTAGE
                                else f"خطای فرمول: {cell.value}"
                            ),
                        )
                    )

        self._audit_differences(grid, table, intel, total_map, audited=audits)
        self._audit_percentages(grid, table, intel, total_map, audited=audits)
        return audits

    @staticmethod
    def _matches_partial_sum(declared: Decimal, values: list[Decimal]) -> bool:
        running = ZERO
        for value in values[:-1]:
            running += value
            if running == declared:
                return True
        return False

    def _audit_differences(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        intel: SheetIntelligence,
        total_map: dict[int, Decimal | None],
        *,
        audited: list[TotalReconciliation],
    ) -> None:
        rows = intel.summary_rows.get(ROW_DIFFERENCE, [])
        if not rows:
            return
        row = rows[0]
        for pair in intel.year_pairs:
            left = pair["index_a"]
            right = pair["index_b"]
            raw = grid.anchor_value(row, left)
            if raw is None or _is_error(raw):
                continue
            declared = _decimal(raw)
            a = total_map.get(left)
            b = total_map.get(right)
            if declared is None or a is None or b is None:
                continue
            expected = b - a
            if declared == expected:
                continue
            cell = grid.get(row, left)
            audited.append(
                TotalReconciliation(
                    sheet=grid.sheet_name,
                    kind=ROW_DIFFERENCE,
                    address=f"{get_column_letter(left)}{row}",
                    row=row,
                    column=left,
                    letter=get_column_letter(left),
                    metric=pair["metric"],
                    declared=declared,
                    computed=expected,
                    difference=declared - expected,
                    status=AuditStatus.MISMATCH,
                    data_range=f"{pair['period_a']} vs {pair['period_b']} totals",
                    formula=cell.formula if cell else None,
                    cause="ردیف مقایسه با تفاوت جمع‌های اعلامی همان دو دوره برابر نیست.",
                )
            )

    def _audit_percentages(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        intel: SheetIntelligence,
        total_map: dict[int, Decimal | None],
        *,
        audited: list[TotalReconciliation],
    ) -> None:
        rows = intel.summary_rows.get(ROW_PERCENTAGE, [])
        if not rows:
            return
        row = rows[0]
        for pair in intel.year_pairs:
            left = pair["index_a"]
            right = pair["index_b"]
            raw = grid.anchor_value(row, left)
            if raw is None or _is_error(raw):
                continue
            declared = _decimal(raw)
            a = total_map.get(left)
            b = total_map.get(right)
            if declared is None or a is None or b is None:
                continue
            if a == ZERO:
                audited.append(
                    TotalReconciliation(
                        sheet=grid.sheet_name,
                        kind=ROW_PERCENTAGE,
                        address=f"{get_column_letter(left)}{row}",
                        row=row,
                        column=left,
                        letter=get_column_letter(left),
                        metric=pair["metric"],
                        declared=declared,
                        computed=None,
                        difference=None,
                        status=AuditStatus.NOT_VERIFIABLE,
                        cause="مخرج درصد صفر است؛ درصد قابل محاسبه نیست.",
                    )
                )
                continue
            expected = ((b - a) / abs(a)).quantize(Decimal("0.000001"))
            if abs(declared - expected) <= Decimal("0.000001"):
                continue
            audited.append(
                TotalReconciliation(
                    sheet=grid.sheet_name,
                    kind=ROW_PERCENTAGE,
                    address=f"{get_column_letter(left)}{row}",
                    row=row,
                    column=left,
                    letter=get_column_letter(left),
                    metric=pair["metric"],
                    declared=declared,
                    computed=expected,
                    difference=declared - expected,
                    status=AuditStatus.MISMATCH,
                    cause="درصد ثبت‌شده با نسبت تفاوت دو دوره به دوره پایه برابر نیست.",
                )
            )

    # ------------------------------------------------------------------ #
    @staticmethod
    def _sheet_warnings(intel: SheetIntelligence) -> list[str]:
        warnings: list[str] = []
        if intel.coverage and intel.coverage.incomplete_period:
            warnings.append(
                f"دوره برگه کامل نیست: {intel.coverage.observed_observations} مشاهده از "
                f"{intel.coverage.expected_observations} روز؛ روزهای بدون مقدار: "
                f"{'، '.join(intel.coverage.missing_labels[:12])}"
            )
        mismatches = [item for item in intel.audits if item.status is AuditStatus.MISMATCH]
        if mismatches:
            warnings.append(f"{len(mismatches)} مغایرت بین جمع/مقایسه اعلامی و محاسبه مستقل پیدا شد.")
        for conflict in intel.metadata_conflicts:
            warnings.append(conflict.message)
        return warnings

    # ------------------------------------------------------------------ #
    def finalize(self, path: str, sheets: list[SheetIntelligence]) -> WorkbookIntelligence:
        intelligence = WorkbookIntelligence(path=path, sheet_count=len(sheets), sheets=sheets)
        counts = Counter(item.status for item in sheets)
        intelligence.status_counts = {status: counts.get(status, 0) for status in STATUS_FA}
        intelligence.metadata_conflicts = [
            conflict for item in sheets for conflict in item.metadata_conflicts
        ]
        intelligence.audits = [item for sheet in sheets for item in sheet.audits]
        intelligence.period_comparisons = [
            item for sheet in sheets for item in sheet.period_comparisons
        ]
        intelligence.similar_groups = self._similar_groups(sheets)
        intelligence.aggregations = self._aggregate(sheets)
        intelligence.cross_sheet_audits = self._cross_sheet_audits(sheets)
        intelligence.warnings = self._workbook_warnings(intelligence)
        intelligence.methodology = [
            "هر برگه ابتدا مستقل تحلیل شد: محدوده واقعی داده، باند عنوان، نقش ردیف‌ها، نقش ستون‌ها، پوشش و مغایرت‌ها.",
            "عنوان چندسطحی به‌صورت سلسله‌مراتبی نگه داشته شد؛ عنوان والد از سلول ادغام‌شده به ستون‌های پوشیده‌شده تعمیم یافت.",
            "ردیف جمع، مقایسه، درصد، توضیح و پانویس از ردیف داده جدا شدند و وارد آماره‌های توصیفی نشدند.",
            "جمع هر ستون دوباره با Decimal محاسبه و با مقدار اعلامی مقایسه شد؛ اختلاف بدون تأیید اصلاح نشد.",
            "سلول خالی، صفر واقعی، مقدار نامعتبر، خطای فرمول و دوره تکمیل‌نشده جداگانه ثبت شدند.",
            "تطبیق بین برگه‌ها فقط بر اساس نام شاخص و برچسب دوره انجام شد؛ برگه‌های مشتق/تجمیعی از جمع بین‌برگه‌ای کنار گذاشته شدند.",
            "همبستگی و مقایسه دوره‌ها توصیفی‌اند و رابطه علت‌ومعلولی اثبات نمی‌کنند.",
        ]
        return intelligence

    @staticmethod
    def _similar_groups(sheets: list[SheetIntelligence]) -> list[dict[str, Any]]:
        groups: dict[tuple[str, ...], list[SheetIntelligence]] = defaultdict(list)
        for sheet in sheets:
            if not sheet.columns:
                continue
            signature = tuple(
                (column.parent_header or column.header or column.letter)
                for column in sheet.columns
                if column.role not in {ColumnRole.ID, ColumnRole.DATE}
            )
            if signature:
                groups[signature].append(sheet)
        results: list[dict[str, Any]] = []
        for signature, members in groups.items():
            if len(members) < 2:
                continue
            results.append(
                {
                    "columns": list(signature),
                    "column_count": len(signature),
                    "sheets": [item.sheet for item in members],
                    "period_pairs": [
                        {"sheet": item.sheet, "years": list(item.subheader_years)}
                        for item in members
                    ],
                    "note": (
                        "این برگه‌ها ساختار ستونی یکسان دارند؛ ستون‌ها فقط با نام والد و برچسب دوره تطبیق داده شدند."
                    ),
                }
            )
        results.sort(key=lambda item: len(item["sheets"]), reverse=True)
        return results

    @staticmethod
    def _aggregate(sheets: list[SheetIntelligence]) -> list[AggregationRecord]:
        buckets: dict[tuple[str, str], list[tuple[str, str, Decimal, int, bool]]] = defaultdict(list)
        for sheet in sheets:
            if sheet.derived or sheet.month_index is None or not sheet.month:
                continue
            incomplete = bool(sheet.coverage and sheet.coverage.incomplete_period)
            for record in sheet.column_totals:
                total = record.get("total")
                period = record.get("period") or ""
                if total is None or not period:
                    continue
                buckets[(record["metric"], period)].append(
                    (sheet.sheet, sheet.month, total, int(record.get("count") or 0), incomplete)
                )
        results: list[AggregationRecord] = []
        for (metric, period), items in sorted(buckets.items()):
            seen: set[str] = set()
            months: list[str] = []
            overlaps: list[str] = []
            total = ZERO
            observations = 0
            used: list[str] = []
            excluded: list[str] = []
            for name, month, value, count, incomplete in items:
                if incomplete:
                    if name not in excluded:
                        excluded.append(name)
                    continue
                if month in seen:
                    if name not in overlaps:
                        overlaps.append(name)
                    continue
                seen.add(month)
                months.append(month)
                used.append(name)
                total += value
                observations += count
            results.append(
                AggregationRecord(
                    metric=metric,
                    period=period,
                    total=total if used else None,
                    observations=observations,
                    sheets=used,
                    months=months,
                    overlaps=overlaps,
                    excluded_incomplete=excluded,
                    note=(
                        "جمع بین‌برگه‌ای فقط برای ماه‌های متمایز، برگه‌های غیرمشتق و دوره‌های کامل انجام شد. "
                        "این عدد جمع سالانه نیست؛ پوشش ماه‌ها در ستون months آمده است."
                        + (" ماه تکراری حذف شد: " + "، ".join(overlaps) if overlaps else "")
                        + (" دوره ناقص وارد جمع نشد: " + "، ".join(excluded) if excluded else "")
                    ),
                )
            )
        return results

    @staticmethod
    def _cross_sheet_audits(sheets: list[SheetIntelligence]) -> list[TotalReconciliation]:
        by_name = {sheet.sheet: sheet for sheet in sheets}
        audits: list[TotalReconciliation] = []
        for sheet in sheets:
            if not sheet.derived:
                continue
            for record in sheet.column_totals:
                metric = record["metric"]
                period = record.get("period") or ""
                if not period:
                    continue
                key = f"{metric} | {period}"
                references = sheet.column_references.get(key) or []
                if not references:
                    continue
                missing: list[str] = []
                incomplete: list[str] = []
                values: list[Decimal] = []
                observations = 0
                for name in references:
                    other = by_name.get(name)
                    if other is None:
                        missing.append(name)
                        continue
                    if other.coverage is not None and other.coverage.incomplete_period:
                        incomplete.append(name)
                        continue
                    match = next(
                        (
                            item
                            for item in other.column_totals
                            if item["metric"] == metric
                            and (item.get("period") or "") == period
                            and item.get("total") is not None
                        ),
                        None,
                    )
                    if match is None:
                        missing.append(name)
                    else:
                        values.append(match["total"])
                        observations += int(match.get("count") or 0)
                declared = record.get("total")
                computed = sum(values, ZERO) if values else None
                if missing or incomplete:
                    status = AuditStatus.NOT_VERIFIABLE
                    parts: list[str] = []
                    if incomplete:
                        parts.append(
                            "دوره ناقص در برگه‌های ارجاع‌شده: " + "، ".join(incomplete[:8])
                        )
                    if missing:
                        parts.append(
                            "مقدار قابل تطبیق برای همین شاخص و دوره پیدا نشد: "
                            + "، ".join(missing[:8])
                        )
                    cause = "؛ ".join(parts) + "؛ جمع ناهمگون انجام نشد."
                elif declared is None or computed is None:
                    status = AuditStatus.NOT_VERIFIABLE
                    cause = "یک طرف تطبیق مقدار ندارد."
                elif declared == computed:
                    status = AuditStatus.VERIFIED
                    cause = ""
                else:
                    status = AuditStatus.MISMATCH
                    cause = "جمع برگه تجمیعی با جمع برگه‌های ارجاع‌شده برابر نیست."
                audits.append(
                    TotalReconciliation(
                        sheet=sheet.sheet,
                        kind="cross_sheet",
                        address=(
                            f"{sheet.sheet}!{record['letter']}{sheet.data_start}:"
                            f"{record['letter']}{sheet.data_end}"
                        ),
                        row=sheet.data_start,
                        column=0,
                        letter=record["letter"],
                        metric=metric,
                        declared=declared,
                        computed=computed,
                        difference=(
                            declared - computed
                            if declared is not None and computed is not None
                            else None
                        ),
                        status=status,
                        data_range=", ".join(references),
                        observations=observations,
                        cause=cause,
                    )
                )
        return audits

    @staticmethod
    def _workbook_warnings(intelligence: WorkbookIntelligence) -> list[str]:
        warnings: list[str] = []
        counts = intelligence.status_counts
        if counts.get(STATUS_PARTIAL):
            warnings.append(f"{counts[STATUS_PARTIAL]} برگه فقط جزئی تحلیل شد.")
        if counts.get(STATUS_NONE):
            warnings.append(f"{counts[STATUS_NONE]} برگه تحلیل نشد.")
        mismatches = [item for item in intelligence.audits if item.status is AuditStatus.MISMATCH]
        unverifiable_totals = [
            item
            for item in intelligence.audits
            if item.status is AuditStatus.NOT_VERIFIABLE and item.kind == "total"
        ]
        unverifiable_percent = [
            item
            for item in intelligence.audits
            if item.status is AuditStatus.NOT_VERIFIABLE and item.kind == "percentage"
        ]
        if mismatches:
            warnings.append(f"{len(mismatches)} مغایرت جمع/مقایسه در سطح برگه‌ها ثبت شد.")
        if unverifiable_totals:
            warnings.append(f"{len(unverifiable_totals)} جمع اعلامی قابل بازتولید مستقل نبود.")
        if unverifiable_percent:
            warnings.append(
                f"{len(unverifiable_percent)} ردیف درصد/مقایسه با داده مرجع قابل بازتولید نبود."
            )
        pair_groups: dict[tuple[str, ...], list[str]] = defaultdict(list)
        for sheet in intelligence.sheets:
            if len(sheet.subheader_years) >= 2:
                pair_groups[tuple(sheet.subheader_years)].append(sheet.sheet)
        if len(pair_groups) > 1:
            parts = [
                f"{'/'.join(pair)}: {'، '.join(names)}"
                for pair, names in sorted(pair_groups.items())
            ]
            warnings.append(
                "زوج سال‌های زیرعنوان بین برگه‌ها یکسان نیست: "
                + " | ".join(parts)
                + ". هیچ سالی به‌صورت خودکار هم‌تراز نشد."
            )
        if intelligence.metadata_conflicts:
            warnings.append(
                f"{len(intelligence.metadata_conflicts)} تعارض متادیتا (نام برگه/عنوان/زیرعنوان) ثبت شد."
            )
        return warnings
