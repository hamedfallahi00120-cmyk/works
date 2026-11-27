"""Explicit semantic classification of extracted fields before analysis.

The detector already assigns a statistical role to every column.  This module
turns those roles into one named vocabulary and records the evidence behind
each decision, so later stages can say *why* a field is a measure, a key or a
descriptive label instead of trusting a bare enum value.

Required roles
--------------
``identifier``, ``temporal_key``, ``measure``, ``category``, ``group``,
``subtotal``, ``grand_total`` and ``descriptive``.  ``calculated`` and
``unknown`` are kept for traceability.  Subtotal and grand total are row roles;
they are reported on the block, never folded into a column measure.

Classification uses column labels, row labels, neighbouring fields, value
patterns, workbook structure (parent/sub/period headers, block kind) and the
available metadata (filename, sheet and block titles).  A numeric identifier is
never reclassified as a measure.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Iterable

from local_office_ai.core.models import ColumnProfile, ColumnRole, DataRole, TableCandidate, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    normalize_digits,
    normalize_text,
    parse_value,
)
from local_office_ai.excel.row_classification import (
    ROW_SUBTOTAL,
    ROW_TOTAL,
    classify_block_rows,
    summary_rows_from_roles,
)
from local_office_ai.excel.semantic_vocabulary import (
    has_identifier_word,
    has_measure_word,
    identifier_evidence,
    matches_word,
    unit_from_header,
)
from local_office_ai.excel.sheet_analyzer import SheetGrid


class SemanticRole(str, Enum):
    """The classification vocabulary required by the analysis contract."""

    IDENTIFIER = "identifier"
    TEMPORAL_KEY = "temporal_key"
    MEASURE = "measure"
    CATEGORY = "category"
    GROUP = "group"
    SUBTOTAL = "subtotal"
    GRAND_TOTAL = "grand_total"
    DESCRIPTIVE = "descriptive"
    CALCULATED = "calculated"
    UNKNOWN = "unknown"


SEMANTIC_ROLE_FA = {
    SemanticRole.IDENTIFIER: "شناسه",
    SemanticRole.TEMPORAL_KEY: "کلید زمانی",
    SemanticRole.MEASURE: "متغیر اندازه‌گیری",
    SemanticRole.CATEGORY: "دسته‌بندی",
    SemanticRole.GROUP: "گروه‌بندی",
    SemanticRole.SUBTOTAL: "جمع میانی",
    SemanticRole.GRAND_TOTAL: "جمع کل",
    SemanticRole.DESCRIPTIVE: "فیلد توصیفی",
    SemanticRole.CALCULATED: "فیلد محاسباتی",
    SemanticRole.UNKNOWN: "نامشخص",
}

DATA_ROLE_TO_SEMANTIC = {
    DataRole.IDENTIFIER: SemanticRole.IDENTIFIER,
    DataRole.TEMPORAL_KEY: SemanticRole.TEMPORAL_KEY,
    DataRole.NUMERIC_MEASURE: SemanticRole.MEASURE,
    DataRole.CATEGORICAL: SemanticRole.CATEGORY,
    DataRole.GROUP: SemanticRole.GROUP,
    DataRole.TEXT: SemanticRole.DESCRIPTIVE,
    DataRole.CALCULATED: SemanticRole.CALCULATED,
    DataRole.UNKNOWN: SemanticRole.UNKNOWN,
}

MEASURE_ROLES = {SemanticRole.MEASURE, SemanticRole.CALCULATED}
KEY_ROLES = {SemanticRole.IDENTIFIER, SemanticRole.TEMPORAL_KEY}
UNIT_HEADER_WORDS = ("واحد", "unit", "measure")


@dataclass
class SemanticField:
    column: int
    letter: str
    header: str
    role: SemanticRole
    confidence: float = 0.0
    business_role: str = ColumnRole.UNKNOWN.value
    statistical_role: str = DataRole.UNKNOWN.value
    unit: str = ""
    period: str = ""
    samples: list[str] = field(default_factory=list)
    evidence: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def role_label(self) -> str:
        return SEMANTIC_ROLE_FA.get(self.role, self.role.value)

    @property
    def is_measure(self) -> bool:
        return self.role in MEASURE_ROLES

    @property
    def is_key(self) -> bool:
        return self.role in KEY_ROLES

    def to_dict(self) -> dict[str, Any]:
        return {
            "column": self.column,
            "letter": self.letter,
            "header": self.header,
            "role": self.role.value,
            "role_label": self.role_label,
            "confidence": round(self.confidence, 4),
            "business_role": self.business_role,
            "statistical_role": self.statistical_role,
            "unit": self.unit,
            "period": self.period,
            "samples": list(self.samples),
            "evidence": list(self.evidence),
            "notes": list(self.notes),
        }


@dataclass
class SemanticRow:
    row: int
    role: SemanticRole
    label: str = ""
    source_role: str = ""

    @property
    def role_label(self) -> str:
        return SEMANTIC_ROLE_FA.get(self.role, self.role.value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "row": self.row,
            "role": self.role.value,
            "role_label": self.role_label,
            "label": self.label[:160],
            "source_role": self.source_role,
        }


@dataclass
class SemanticBlock:
    block_id: str
    sheet: str
    kind: str
    title: str
    address: str
    columns: list[SemanticField]
    rows: list[SemanticRow] = field(default_factory=list)
    periods: list[str] = field(default_factory=list)

    @property
    def measures(self) -> list[SemanticField]:
        return [column for column in self.columns if column.is_measure]

    @property
    def keys(self) -> list[SemanticField]:
        return [column for column in self.columns if column.is_key]

    def role_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for column in self.columns:
            counts[column.role.value] = counts.get(column.role.value, 0) + 1
        return counts

    def to_dict(self) -> dict[str, Any]:
        return {
            "block_id": self.block_id,
            "sheet": self.sheet,
            "kind": self.kind,
            "title": self.title,
            "address": self.address,
            "columns": [column.to_dict() for column in self.columns],
            "rows": [row.to_dict() for row in self.rows],
            "periods": list(self.periods),
            "role_counts": self.role_counts(),
        }


@dataclass
class WorkbookSemantics:
    source_filename: str
    blocks: list[SemanticBlock]

    def field_roles(self) -> dict[str, SemanticRole]:
        return {
            field.header: field.role
            for block in self.blocks
            for field in block.columns
            if field.header
        }

    def role_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for block in self.blocks:
            for role, count in block.role_counts().items():
                counts[role] = counts.get(role, 0) + count
        return counts

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_filename": self.source_filename,
            "blocks": [block.to_dict() for block in self.blocks],
            "role_counts": self.role_counts(),
        }


class SemanticClassifier:
    """Classify every detected column into the shared semantic vocabulary."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def classify(
        self,
        *,
        source_path: Path | str,
        sheet_names: Iterable[str],
        grids: dict[str, SheetGrid],
        structures: dict[str, Any],
    ) -> WorkbookSemantics:
        path = Path(source_path)
        blocks: list[SemanticBlock] = []
        for name in sheet_names:
            grid = grids.get(name)
            structure = structures.get(name)
            if grid is None or structure is None:
                continue
            tables = list(getattr(structure, "tables", []) or [])
            tables.extend(getattr(structure, "key_value_tables", []) or [])
            for index, table in enumerate(tables, start=1):
                blocks.append(
                    self.classify_table(
                        grid,
                        table,
                        block_id=f"{name}#B{index}",
                        filename=path.name,
                    )
                )
        return WorkbookSemantics(source_filename=path.name, blocks=blocks)

    # ------------------------------------------------------------------ #
    def classify_table(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        *,
        block_id: str = "",
        filename: str = "",
    ) -> SemanticBlock:
        header_rows = list(table.header_rows or [table.header_row])
        fields = [
            self._classify_column(grid, table, column, filename=filename)
            for column in table.columns
        ]
        row_records = classify_block_rows(
            grid, table, header_rows, block_end_row=table.block_end
        )
        summary = summary_rows_from_roles(row_records)
        labels = {record.row: record.label for record in row_records}
        rows: list[SemanticRow] = []
        for source_role, semantic in (
            (ROW_SUBTOTAL, SemanticRole.SUBTOTAL),
            (ROW_TOTAL, SemanticRole.GRAND_TOTAL),
        ):
            for row in summary.get(source_role, []):
                rows.append(
                    SemanticRow(
                        row=row,
                        role=semantic,
                        label=labels.get(row, ""),
                        source_role=source_role,
                    )
                )
        rows.sort(key=lambda item: (item.row, item.role.value))
        periods = list(
            dict.fromkeys(
                column.period
                for column in table.columns
                if column.period
            )
        )
        return SemanticBlock(
            block_id=block_id or f"{table.sheet}#{table.address}",
            sheet=table.sheet,
            kind=table.kind.value,
            title=table.title,
            address=table.address,
            columns=fields,
            rows=rows,
            periods=periods,
        )

    # ------------------------------------------------------------------ #
    def _classify_column(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        profile: ColumnProfile,
        *,
        filename: str,
    ) -> SemanticField:
        numbers = self._numbers(grid, table, profile)
        samples = [str(sample) for sample in profile.samples]
        base = DATA_ROLE_TO_SEMANTIC.get(profile.statistical_role, SemanticRole.UNKNOWN)
        evidence: list[str] = []
        notes = list(profile.notes)

        id_evidence = identifier_evidence(profile.header, numbers, samples)
        if base is SemanticRole.TEMPORAL_KEY:
            role = SemanticRole.TEMPORAL_KEY
            evidence.append("عنوان و محتوا کلید زمانی سال/ماه/روز را نشان می‌دهد")
        elif id_evidence:
            role = SemanticRole.IDENTIFIER
            evidence.extend(id_evidence)
        elif profile.role is ColumnRole.ID:
            role = SemanticRole.IDENTIFIER
            evidence.append("نقش تجاری شناسه از عنوان ستون")
        else:
            role = base

        if role is SemanticRole.MEASURE:
            evidence.append(f"نسبت مقادیر عددی {profile.numeric_ratio:.0%}")
        elif role in {SemanticRole.CATEGORY, SemanticRole.GROUP}:
            evidence.append(
                f"{profile.distinct} مقدار متمایز در {profile.non_empty} مقدار غیرخالی"
            )
            if role is SemanticRole.GROUP:
                evidence.append("عنوان ستون واژه گروه/دسته دارد")
        elif role is SemanticRole.DESCRIPTIVE:
            evidence.append(f"نسبت مقادیر متنی {profile.text_ratio:.0%}")
        elif role is SemanticRole.CALCULATED:
            evidence.append(f"نسبت سلول‌های فرمول {profile.formula_ratio:.0%}")

        if profile.parent_header or profile.sub_header:
            parts = [part for part in (profile.parent_header, profile.sub_header) if part]
            evidence.append("سرستون والد/فرعی: " + " / ".join(parts))
        if profile.period:
            evidence.append(f"دوره ستون: {profile.period}")
        if table.title:
            evidence.append(f"عنوان جدول/بلوک: {table.title}")
        evidence.append(f"برگه: {table.sheet}")

        unit = unit_from_header(profile.header) or self._neighbour_unit(grid, table, profile)
        if unit:
            evidence.append(f"واحد: {unit}")

        if filename:
            evidence.append(f"منبع: {filename}")
        if not evidence:
            evidence.append("شاهد کافی برای نقش قطعی پیدا نشد")

        metadata = " ".join(part for part in (filename, table.sheet, table.title) if part)
        confidence = self._confidence(profile, role)
        if role is SemanticRole.IDENTIFIER and has_identifier_word(metadata, strong_only=True):
            confidence = min(1.0, confidence + 0.05)
            evidence.append("برگه/عنوان فایل نیز به شناسه اشاره دارد")
        elif role is SemanticRole.MEASURE and has_measure_word(metadata):
            confidence = min(1.0, confidence + 0.05)
            evidence.append("برگه/عنوان فایل نیز به اندازه‌گیری اشاره دارد")

        return SemanticField(
            column=profile.index,
            letter=profile.letter,
            header=profile.header,
            role=role,
            confidence=confidence,
            business_role=profile.role.value,
            statistical_role=profile.statistical_role.value,
            unit=unit,
            period=profile.period or "",
            samples=list(profile.samples),
            evidence=evidence,
            notes=notes,
        )

    # ------------------------------------------------------------------ #
    @staticmethod
    def _confidence(profile: ColumnProfile, role: SemanticRole) -> float:
        if role in KEY_ROLES:
            return max(0.85, profile.confidence)
        if role is SemanticRole.MEASURE:
            return max(profile.numeric_ratio, profile.confidence)
        if role in {SemanticRole.CATEGORY, SemanticRole.GROUP, SemanticRole.DESCRIPTIVE}:
            return max(profile.text_ratio, profile.confidence)
        if role is SemanticRole.CALCULATED:
            return max(profile.formula_ratio, profile.confidence)
        return profile.confidence

    @staticmethod
    def _numbers(
        grid: SheetGrid,
        table: TableCandidate,
        profile: ColumnProfile,
    ) -> list[float]:
        values: list[float] = []
        for row in range(table.first_data_row, table.last_data_row + 1):
            cell = grid.get(row, profile.index)
            if cell is None or cell.is_empty:
                continue
            parsed = parse_value(cell.value)
            if parsed.kind is not ValueKind.NUMBER:
                continue
            try:
                values.append(float(parsed.value))
            except (TypeError, ValueError):
                continue
        return values

    @staticmethod
    def _neighbour_unit(
        grid: SheetGrid,
        table: TableCandidate,
        profile: ColumnProfile,
    ) -> str:
        """Read a unit from an adjacent «واحد» column when one exists."""
        for neighbour in (profile.index + 1, profile.index - 1):
            column = next(
                (item for item in table.columns if item.index == neighbour), None
            )
            if column is None:
                continue
            header = normalize_text(normalize_digits(column.header or ""))
            if not any(matches_word(header, word) for word in UNIT_HEADER_WORDS):
                continue
            for row in range(table.first_data_row, table.last_data_row + 1):
                cell = grid.get(row, neighbour)
                if cell is None or cell.is_empty:
                    continue
                text = normalize_text(normalize_digits(str(cell.value)))
                if text:
                    return text[:40]
        return ""
