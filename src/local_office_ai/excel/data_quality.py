"""Explicit data-quality findings recorded before conclusions are drawn.

The analyzer never deletes, repairs or replaces a questionable value.  It states
what the issue is, how often it occurs and where it lives, and gives every check
one of three statuses: ``PASS``, ``WARNING`` or ``FAIL``.  Severity is tied to
the actual reporting requirement: a complete measure passes, a measure with some
blanks is a warning, and a conflicting period or a broken unique identifier is a
failure.

The eight required checks are:

``missing_values``, ``invalid_numeric``, ``unexpected_text``, ``negative_values``,
``duplicate_identifiers``, ``inconsistent_units``, ``conflicting_periods`` and
``suspicious_values``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any

from local_office_ai.core.models import AnalysisResult, ColumnProfile, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    looks_like_date,
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
    PeriodSource,
    cell_text,
    detect_periods,
    scalar_conflicts,
)
from local_office_ai.excel.row_classification import (
    ROW_DATA,
    ROW_DATA_MISSING,
    classify_block_rows,
)
from local_office_ai.excel.semantic_classification import (
    DATA_ROLE_TO_SEMANTIC,
    MEASURE_ROLES,
    SemanticField,
    SemanticRole,
)
from local_office_ai.excel.semantic_vocabulary import (
    allows_negative,
    currency_words_in,
    has_measure_word,
    requires_unique_identifier,
    scale_words_in,
)
from local_office_ai.excel.sheet_analyzer import SheetGrid, summary_label_role


class QualityStatus(str, Enum):
    PASS = "PASS"
    WARNING = "WARNING"
    FAIL = "FAIL"


QUALITY_STATUS_FA = {
    QualityStatus.PASS: "قبول",
    QualityStatus.WARNING: "هشدار",
    QualityStatus.FAIL: "رد",
}

REQUIRED_CHECKS = (
    ("missing_values", "مقادیر گمشده"),
    ("invalid_numeric", "مقدار عددی نامعتبر"),
    ("unexpected_text", "متن غیرمنتظره در ستون عددی"),
    ("negative_values", "مقدار منفی نامعتبر"),
    ("duplicate_identifiers", "شناسه تکراری"),
    ("inconsistent_units", "ناسازگاری واحد"),
    ("conflicting_periods", "دوره متناقض"),
    ("suspicious_values", "مقدار مشکوک/پرت"),
)
REQUIRED_CHECK_CODES = {code for code, _ in REQUIRED_CHECKS}

_EXCEL_ERROR = re.compile(r"^#(?:DIV/0!|N/A|VALUE!|REF!|NAME\?|NUM!|NULL!)$", re.IGNORECASE)
_LETTER = re.compile(r"[A-Za-z\u0600-\u06FF]")
_DIGIT = re.compile(r"\d")


@dataclass
class QualityFinding:
    code: str
    title: str
    status: QualityStatus
    message: str
    scope: str = ""
    sheet: str = ""
    address: str = ""
    column: str = ""
    count: int = 0
    samples: list[str] = field(default_factory=list)
    evidence: list[str] = field(default_factory=list)

    @property
    def status_label(self) -> str:
        return QUALITY_STATUS_FA.get(self.status, self.status.value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "title": self.title,
            "status": self.status.value,
            "status_label": self.status_label,
            "message": self.message,
            "scope": self.scope,
            "sheet": self.sheet,
            "address": self.address,
            "column": self.column,
            "count": self.count,
            "samples": list(self.samples),
            "evidence": list(self.evidence),
        }


@dataclass
class QualityReport:
    findings: list[QualityFinding] = field(default_factory=list)

    @property
    def status(self) -> QualityStatus:
        statuses = {finding.status for finding in self.findings}
        if QualityStatus.FAIL in statuses:
            return QualityStatus.FAIL
        if QualityStatus.WARNING in statuses:
            return QualityStatus.WARNING
        return QualityStatus.PASS

    @property
    def status_label(self) -> str:
        return QUALITY_STATUS_FA.get(self.status, self.status.value)

    @property
    def failed(self) -> list[QualityFinding]:
        return [item for item in self.findings if item.status is QualityStatus.FAIL]

    @property
    def warnings(self) -> list[QualityFinding]:
        return [item for item in self.findings if item.status is QualityStatus.WARNING]

    @property
    def passed(self) -> list[QualityFinding]:
        return [item for item in self.findings if item.status is QualityStatus.PASS]

    def counts(self) -> dict[str, int]:
        counts = {status.value: 0 for status in QualityStatus}
        for finding in self.findings:
            counts[finding.status.value] += 1
        return counts

    def check_status(self, code: str) -> QualityStatus | None:
        relevant = [item.status for item in self.findings if item.code == code]
        if not relevant:
            return None
        if QualityStatus.FAIL in relevant:
            return QualityStatus.FAIL
        if QualityStatus.WARNING in relevant:
            return QualityStatus.WARNING
        return QualityStatus.PASS

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status.value,
            "status_label": self.status_label,
            "counts": self.counts(),
            "findings": [finding.to_dict() for finding in self.findings],
        }


class DataQualityAnalyzer:
    """Run the required quality checks over the primary extracted table."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def analyze(
        self,
        analysis: AnalysisResult,
        grids: dict[str, SheetGrid] | None = None,
    ) -> QualityReport:
        grids = grids or {}
        table = analysis.primary_table
        if table is None or table.sheet not in grids:
            return QualityReport(
                [
                    QualityFinding(
                        code="primary-table",
                        title="جدول اصلی",
                        status=QualityStatus.FAIL,
                        message="جدول داده‌ای برای اعتبارسنجی معنایی و کیفیت پیدا نشد.",
                        scope=Path(analysis.meta.path).name,
                    )
                ]
            )

        grid = grids[table.sheet]
        profiles = {column.index: column for column in table.columns}
        fields = self._fields(analysis, table)
        rows = self._data_rows(grid, table)

        findings: list[QualityFinding] = []
        applicable: set[str] = set()

        for check in (
            self._missing_values,
            self._invalid_numeric,
            self._negative_values,
            self._duplicate_identifiers,
            self._inconsistent_units,
            self._suspicious_values,
        ):
            produced, used = check(grid, table, fields, profiles, rows)
            findings.extend(produced)
            if used:
                applicable.add(check.__name__)

        periods, period_used = self._period_conflicts(analysis, grid, table)
        findings.extend(periods)
        if period_used:
            applicable.add("_period_conflicts")

        findings.extend(self._ensure_pass(findings, applicable, table))
        return QualityReport(findings)

    # ------------------------------------------------------------------ #
    def _fields(self, analysis: AnalysisResult, table: Any) -> list[SemanticField]:
        semantics = getattr(analysis, "semantics", None)
        if semantics is not None:
            for block in getattr(semantics, "blocks", []):
                if block.sheet == table.sheet and block.address == table.address:
                    return list(block.columns)
        fields: list[SemanticField] = []
        for profile in table.columns:
            role = DATA_ROLE_TO_SEMANTIC.get(profile.statistical_role, SemanticRole.UNKNOWN)
            fields.append(
                SemanticField(
                    column=profile.index,
                    letter=profile.letter,
                    header=profile.header,
                    role=role,
                    statistical_role=profile.statistical_role.value,
                    business_role=profile.role.value,
                )
            )
        return fields

    @staticmethod
    def _data_rows(grid: SheetGrid, table: Any) -> list[int]:
        records = classify_block_rows(
            grid, table, list(table.header_rows or [table.header_row]), block_end_row=table.block_end
        )
        rows = [
            record.row
            for record in records
            if record.role in {ROW_DATA, ROW_DATA_MISSING}
            and table.first_data_row <= record.row <= table.last_data_row
        ]
        if rows:
            return rows
        return list(range(table.first_data_row, table.last_data_row + 1))

    @staticmethod
    def _expects_numeric(field: SemanticField, profile: ColumnProfile | None) -> bool:
        """True when a column is a measure, even if invalid text diluted its role."""
        if field.role in MEASURE_ROLES:
            return True
        if profile is None:
            return False
        if profile.data_type in {ValueKind.NUMBER, ValueKind.ERROR}:
            return True
        return has_measure_word(field.header) and profile.numeric_ratio > 0

    # ------------------------------------------------------------------ #
    def _missing_values(
        self,
        grid: SheetGrid,
        table: Any,
        fields: list[SemanticField],
        profiles: dict[int, ColumnProfile],
        rows: list[int],
    ) -> tuple[list[QualityFinding], bool]:
        findings: list[QualityFinding] = []
        used = False
        for field in fields:
            profile = profiles.get(field.column)
            if not self._expects_numeric(field, profile):
                continue
            used = True
            missing: list[str] = []
            for row in rows:
                cell = grid.get(row, field.column)
                if cell is None or (cell.is_empty and not cell.formula):
                    missing.append(f"{field.letter}{row}")
            if not missing:
                continue
            ratio = len(missing) / max(1, len(rows))
            findings.append(
                QualityFinding(
                    code="missing_values",
                    title="مقادیر گمشده",
                    status=QualityStatus.FAIL if ratio >= 0.5 else QualityStatus.WARNING,
                    message=(
                        f"ستون «{field.header or field.letter}» {len(missing)} مقدار از "
                        f"{len(rows)} ردیف داده را ندارد ({ratio:.0%})."
                    ),
                    scope=f"{table.sheet}!{field.letter}",
                    sheet=table.sheet,
                    address=missing[0],
                    column=field.letter,
                    count=len(missing),
                    samples=missing[:8],
                )
            )
        return findings, used

    def _invalid_numeric(
        self,
        grid: SheetGrid,
        table: Any,
        fields: list[SemanticField],
        profiles: dict[int, ColumnProfile],
        rows: list[int],
    ) -> tuple[list[QualityFinding], bool]:
        findings: list[QualityFinding] = []
        used = False
        for field in fields:
            profile = profiles.get(field.column)
            numeric_identifier = (
                field.role is SemanticRole.IDENTIFIER
                and profile is not None
                and profile.data_type in {ValueKind.NUMBER, ValueKind.ERROR}
            )
            if not (self._expects_numeric(field, profile) or numeric_identifier):
                continue
            used = True
            invalid: list[str] = []
            unexpected: list[str] = []
            for row in rows:
                cell = grid.get(row, field.column)
                if cell is None or cell.is_empty:
                    continue
                value = cell.value
                if isinstance(value, str) and _EXCEL_ERROR.match(value.strip()):
                    invalid.append(f"{field.letter}{row}")
                    continue
                parsed = parse_value(value)
                if parsed.kind is ValueKind.ERROR:
                    invalid.append(f"{field.letter}{row}")
                    continue
                if isinstance(value, str):
                    if looks_like_date(value) or parsed.kind is ValueKind.NUMBER:
                        continue
                    if _DIGIT.search(value) and _LETTER.search(value):
                        invalid.append(f"{field.letter}{row}")
                    elif not _DIGIT.search(value):
                        unexpected.append(f"{field.letter}{row}")
            if invalid:
                ratio = len(invalid) / max(1, len(rows))
                findings.append(
                    QualityFinding(
                        code="invalid_numeric",
                        title="مقدار عددی نامعتبر",
                        status=QualityStatus.FAIL if ratio >= 0.5 else QualityStatus.WARNING,
                        message=(
                            f"ستون «{field.header or field.letter}» {len(invalid)} مقدار عددی "
                            "نامعتبر یا خطای فرمول دارد."
                        ),
                        scope=f"{table.sheet}!{field.letter}",
                        sheet=table.sheet,
                        address=invalid[0],
                        column=field.letter,
                        count=len(invalid),
                        samples=invalid[:8],
                    )
                )
            if unexpected:
                findings.append(
                    QualityFinding(
                        code="unexpected_text",
                        title="متن غیرمنتظره در ستون عددی",
                        status=QualityStatus.WARNING,
                        message=(
                            f"ستون «{field.header or field.letter}» {len(unexpected)} مقدار متنی "
                            "بدون رقم دارد؛ مقدار جایگزین یا اصلاح نشد."
                        ),
                        scope=f"{table.sheet}!{field.letter}",
                        sheet=table.sheet,
                        address=unexpected[0],
                        column=field.letter,
                        count=len(unexpected),
                        samples=unexpected[:8],
                    )
                )
        return findings, used

    def _negative_values(
        self,
        grid: SheetGrid,
        table: Any,
        fields: list[SemanticField],
        profiles: dict[int, ColumnProfile],
        rows: list[int],
    ) -> tuple[list[QualityFinding], bool]:
        findings: list[QualityFinding] = []
        used = False
        for field in fields:
            profile = profiles.get(field.column)
            if not self._expects_numeric(field, profile):
                continue
            used = True
            if allows_negative(field.header):
                continue
            negatives: list[str] = []
            for row in rows:
                cell = grid.get(row, field.column)
                if cell is None or cell.is_empty:
                    continue
                parsed = parse_value(cell.value)
                if parsed.kind is not ValueKind.NUMBER:
                    continue
                try:
                    if float(parsed.value) < 0:
                        negatives.append(f"{field.letter}{row}")
                except (TypeError, ValueError):
                    continue
            if negatives:
                findings.append(
                    QualityFinding(
                        code="negative_values",
                        title="مقدار منفی نامعتبر",
                        status=QualityStatus.WARNING,
                        message=(
                            f"ستون «{field.header or field.letter}» {len(negatives)} مقدار منفی دارد "
                            "که با معنای اندازه‌گیری سازگار نیست؛ مقدار حفظ شد."
                        ),
                        scope=f"{table.sheet}!{field.letter}",
                        sheet=table.sheet,
                        address=negatives[0],
                        column=field.letter,
                        count=len(negatives),
                        samples=negatives[:8],
                    )
                )
        return findings, used

    def _duplicate_identifiers(
        self,
        grid: SheetGrid,
        table: Any,
        fields: list[SemanticField],
        profiles: dict[int, ColumnProfile],
        rows: list[int],
    ) -> tuple[list[QualityFinding], bool]:
        findings: list[QualityFinding] = []
        used = False
        for field in fields:
            if field.role is not SemanticRole.IDENTIFIER:
                continue
            used = True
            seen: dict[str, list[str]] = {}
            for row in rows:
                cell = grid.get(row, field.column)
                if cell is None or cell.is_empty:
                    continue
                key = normalize_text(normalize_digits(cell_text(cell.value))).strip()
                if not key:
                    continue
                seen.setdefault(key, []).append(f"{field.letter}{row}")
            duplicates = {key: cells for key, cells in seen.items() if len(cells) > 1}
            if not duplicates:
                continue
            unique_ratio = len(seen) / max(1, sum(len(cells) for cells in seen.values()))
            strict = requires_unique_identifier(field.header)
            if not strict and unique_ratio < 0.8:
                # A repeated region/category code is a grouping key, not a broken
                # unique identifier, so it is not reported as a duplicate.
                continue
            samples = [
                f"{key} ×{len(cells)} ({cells[0]})"
                for key, cells in list(duplicates.items())[:8]
            ]
            findings.append(
                QualityFinding(
                    code="duplicate_identifiers",
                    title="شناسه تکراری",
                    status=QualityStatus.FAIL if strict else QualityStatus.WARNING,
                    message=(
                        f"ستون «{field.header or field.letter}» {len(duplicates)} مقدار تکراری دارد"
                        + (" و یکتایی آن الزامی است." if strict else "؛ یکتایی آن تأیید نشد.")
                    ),
                    scope=f"{table.sheet}!{field.letter}",
                    sheet=table.sheet,
                    address=next(iter(duplicates.values()))[0],
                    column=field.letter,
                    count=len(duplicates),
                    samples=samples,
                )
            )
        return findings, used

    def _inconsistent_units(
        self,
        grid: SheetGrid,
        table: Any,
        fields: list[SemanticField],
        profiles: dict[int, ColumnProfile],
        rows: list[int],
    ) -> tuple[list[QualityFinding], bool]:
        findings: list[QualityFinding] = []
        used = False
        for field in fields:
            profile = profiles.get(field.column)
            if not self._expects_numeric(field, profile):
                continue
            used = True
            cell_scales: set[str] = set()
            cell_currencies: set[str] = set()
            addresses: list[str] = []
            for row in rows:
                cell = grid.get(row, field.column)
                if cell is None or cell.is_empty or not isinstance(cell.value, str):
                    continue
                scales = scale_words_in(cell.value)
                currencies = currency_words_in(cell.value)
                if scales or currencies:
                    addresses.append(f"{field.letter}{row}")
                cell_scales |= scales
                cell_currencies |= currencies
            header_scales = scale_words_in(field.header)
            header_currencies = currency_words_in(field.header)
            problems: list[str] = []
            if len(cell_scales) > 1:
                problems.append("مقیاس‌های مختلط: " + "، ".join(sorted(cell_scales)))
            if len(cell_currencies) > 1:
                problems.append("واحدهای پول مختلط: " + "، ".join(sorted(cell_currencies)))
            if header_scales and cell_scales and header_scales != cell_scales:
                problems.append(
                    "مقیاس عنوان و سلول یکسان نیست: "
                    f"عنوان {sorted(header_scales)} در برابر سلول {sorted(cell_scales)}"
                )
            if (
                header_currencies
                and cell_currencies
                and not (header_currencies & cell_currencies)
            ):
                problems.append(
                    "واحد پول عنوان و سلول یکسان نیست: "
                    f"عنوان {sorted(header_currencies)} در برابر سلول {sorted(cell_currencies)}"
                )
            if problems:
                findings.append(
                    QualityFinding(
                        code="inconsistent_units",
                        title="ناسازگاری واحد",
                        status=QualityStatus.WARNING,
                        message=(
                            f"ستون «{field.header or field.letter}» واحد یکدست ندارد؛ "
                            "مقدارها تبدیل یا اصلاح نشدند."
                        ),
                        scope=f"{table.sheet}!{field.letter}",
                        sheet=table.sheet,
                        address=addresses[0] if addresses else "",
                        column=field.letter,
                        count=len(addresses),
                        samples=addresses[:8],
                        evidence=problems,
                    )
                )
        return findings, used

    def _suspicious_values(
        self,
        grid: SheetGrid,
        table: Any,
        fields: list[SemanticField],
        profiles: dict[int, ColumnProfile],
        rows: list[int],
    ) -> tuple[list[QualityFinding], bool]:
        findings: list[QualityFinding] = []
        used = False
        for field in fields:
            profile = profiles.get(field.column)
            if not self._expects_numeric(field, profile):
                continue
            used = True
            values: list[tuple[str, float]] = []
            for row in rows:
                cell = grid.get(row, field.column)
                if cell is None or cell.is_empty:
                    continue
                parsed = parse_value(cell.value)
                if parsed.kind is not ValueKind.NUMBER:
                    continue
                try:
                    values.append((f"{field.letter}{row}", float(parsed.value)))
                except (TypeError, ValueError):
                    continue
            if len(values) < 3:
                continue
            numbers = [value for _, value in values]
            if len(set(numbers)) == 1:
                findings.append(
                    QualityFinding(
                        code="suspicious_values",
                        title="مقدار مشکوک/پرت",
                        status=QualityStatus.WARNING,
                        message=(
                            f"ستون «{field.header or field.letter}» در همه {len(numbers)} مقدار "
                            "ثابت است؛ واریانس صفر است."
                        ),
                        scope=f"{table.sheet}!{field.letter}",
                        sheet=table.sheet,
                        address=values[0][0],
                        column=field.letter,
                        count=len(numbers),
                        samples=[values[0][0]],
                    )
                )
                continue
            outliers = self._iqr_outliers(values)
            extreme = [address for address, value in values if abs(value) >= 1e15]
            flagged = list(dict.fromkeys(outliers + extreme))
            if flagged:
                findings.append(
                    QualityFinding(
                        code="suspicious_values",
                        title="مقدار مشکوک/پرت",
                        status=QualityStatus.WARNING,
                        message=(
                            f"ستون «{field.header or field.letter}» {len(flagged)} مقدار خارج از "
                            "بازه معمول دارد؛ مقدار حذف نشد و فقط علامت‌گذاری شد."
                        ),
                        scope=f"{table.sheet}!{field.letter}",
                        sheet=table.sheet,
                        address=flagged[0],
                        column=field.letter,
                        count=len(flagged),
                        samples=flagged[:8],
                    )
                )
        return findings, used

    @staticmethod
    def _iqr_outliers(values: list[tuple[str, float]]) -> list[str]:
        ordered = sorted(value for _, value in values)
        count = len(ordered)
        q1 = ordered[count // 4]
        q3 = ordered[(count * 3) // 4]
        iqr = q3 - q1
        if iqr <= 0:
            return []
        low = q1 - 1.5 * iqr
        high = q3 + 1.5 * iqr
        return [address for address, value in values if value < low or value > high]

    # ------------------------------------------------------------------ #
    def _period_conflicts(
        self,
        analysis: AnalysisResult,
        grid: SheetGrid,
        table: Any,
    ) -> tuple[list[QualityFinding], bool]:
        sources = self._period_sources(grid, table, Path(analysis.meta.path).name)
        detection = detect_periods(sources)
        conflicts = scalar_conflicts(detection, dimension=KIND_YEAR)
        conflicts.extend(scalar_conflicts(detection, dimension=KIND_MONTH))

        intelligence = getattr(analysis, "intelligence", None)
        if intelligence is not None:
            for sheet in getattr(intelligence, "sheets", []):
                if sheet.sheet != table.sheet:
                    continue
                for conflict in getattr(sheet, "metadata_conflicts", []):
                    payload = conflict.to_dict() if hasattr(conflict, "to_dict") else dict(conflict)
                    conflicts.append(payload)
        if not conflicts:
            return [], True
        evidence = [
            "؛ ".join(
                f"{key}={value}"
                for key, value in conflict.items()
                if key in {"kind", "values", "sources", "detail", "field", "sheet", "title"}
            )
            for conflict in conflicts
        ]
        return (
            [
                QualityFinding(
                    code="conflicting_periods",
                    title="دوره متناقض",
                    status=QualityStatus.FAIL,
                    message=(
                        f"{len(conflicts)} تعارض دوره در برگه «{table.sheet}» ثبت شد؛ "
                        "هیچ دوره‌ای به‌صورت حدسی انتخاب نشد."
                    ),
                    scope=f"{table.sheet}!{table.address}",
                    sheet=table.sheet,
                    address=table.address,
                    count=len(conflicts),
                    evidence=evidence[:8],
                )
            ],
            True,
        )

    @staticmethod
    def _period_sources(
        grid: SheetGrid,
        table: Any,
        filename: str,
    ) -> list[PeriodSource]:
        sources: list[PeriodSource] = [
            PeriodSource(SOURCE_SHEET_NAME, table.sheet),
            PeriodSource(SOURCE_FILENAME, filename),
        ]
        if table.title:
            sources.append(PeriodSource(SOURCE_BLOCK_TITLE, table.title))
        for column in table.columns:
            text = " ".join(
                part for part in (column.parent_header, column.sub_header, column.period) if part
            )
            if text:
                sources.append(
                    PeriodSource(SOURCE_COLUMN_HEADER, text, address=column.header_address)
                )
        limit = min(table.last_data_row, table.first_data_row + 299)
        for row in range(table.first_data_row, limit + 1):
            for column in table.columns[:2]:
                cell = grid.get(row, column.index)
                if cell is None or cell.is_empty:
                    continue
                text = cell_text(cell.value)
                if not text or summary_label_role(text):
                    continue
                if isinstance(cell.value, str) and (
                    looks_like_date(text) or parse_value(text).kind is ValueKind.NUMBER
                ):
                    continue
                if not DataQualityAnalyzer._row_has_numeric(grid, table, row):
                    continue
                sources.append(
                    PeriodSource(
                        SOURCE_ROW_LABEL,
                        text,
                        address=f"{column.letter}{row}",
                    )
                )
                break
        return sources

    @staticmethod
    def _row_has_numeric(grid: SheetGrid, table: Any, row: int) -> bool:
        for column in table.columns:
            cell = grid.get(row, column.index)
            if cell is None or cell.is_empty:
                continue
            if parse_value(cell.value).kind is ValueKind.NUMBER:
                return True
        return False

    # ------------------------------------------------------------------ #
    @staticmethod
    def _ensure_pass(
        findings: list[QualityFinding],
        applicable: set[str],
        table: Any,
    ) -> list[QualityFinding]:
        present = {finding.code for finding in findings}
        additions: list[QualityFinding] = []
        for code, title in REQUIRED_CHECKS:
            if code in present:
                continue
            used = _check_name(code) in applicable
            additions.append(
                QualityFinding(
                    code=code,
                    title=title,
                    status=QualityStatus.PASS,
                    message=(
                        "بررسی انجام شد؛ موردی یافت نشد."
                        if used
                        else "ستون مرتبطی برای این بررسی وجود نداشت."
                    ),
                    scope=f"{table.sheet}!{table.address}",
                    sheet=table.sheet,
                )
            )
        return additions


def _check_name(code: str) -> str:
    return {
        "missing_values": "_missing_values",
        "invalid_numeric": "_invalid_numeric",
        "unexpected_text": "_invalid_numeric",
        "negative_values": "_negative_values",
        "duplicate_identifiers": "_duplicate_identifiers",
        "inconsistent_units": "_inconsistent_units",
        "conflicting_periods": "_period_conflicts",
        "suspicious_values": "_suspicious_values",
    }.get(code, code)
