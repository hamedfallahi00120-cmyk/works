"""Structured intermediate representation of an extracted workbook.

The detectors already find tables, headers and periods, but their results are
scattered across ``TableCandidate``, ``SheetGrid`` and ``SheetIntelligence``.
Later stages then have to rediscover where a block starts, which cells a merged
header covers, or why a period was inferred.

This module keeps one explicit model per workbook, sheet and block:

* source filename and worksheet name,
* block identifier and boundaries,
* header rows and column names,
* merged-cell relationships and cell addresses,
* original cell values and formulas,
* detected periods with the evidence and confidence behind each one,
* the semantic role of every column.

It is a projection over the existing detectors, not a second detector: the
builder only reads ``TableCandidate``/``ColumnProfile``/``SheetGrid`` results.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from pathlib import Path
from typing import Any, Iterable

from openpyxl.utils import get_column_letter

from local_office_ai.core.models import ColumnProfile, DataRole, TableCandidate, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    looks_like_date,
    looks_like_number,
    normalize_digits,
    normalize_text,
)
from local_office_ai.excel.period_detection import (
    KIND_MONTH,
    KIND_YEAR,
    STATUS_CONFLICTING,
    STATUS_DETECTED,
    SOURCE_BLOCK_TITLE,
    SOURCE_COLUMN_HEADER,
    SOURCE_FILENAME,
    SOURCE_ROW_LABEL,
    SOURCE_SHEET_NAME,
    PeriodDetection,
    PeriodSource,
    detect_periods,
    scalar_conflicts,
)
from local_office_ai.excel.sheet_analyzer import SheetGrid

DEFAULT_MAX_CELLS_PER_BLOCK = 4000
DEFAULT_MAX_ROW_LABEL_SCAN = 300


def _jsonable(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    return value


def _cell_address(row: int, column: int) -> str:
    return f"{get_column_letter(column)}{row}"


def _range_address(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"


def _role_confidence(profile: ColumnProfile) -> float:
    role = profile.statistical_role
    if role in {DataRole.IDENTIFIER, DataRole.TEMPORAL_KEY}:
        return max(0.8, profile.confidence)
    if role is DataRole.NUMERIC_MEASURE:
        return max(profile.numeric_ratio, profile.confidence)
    if role in {DataRole.CATEGORICAL, DataRole.GROUP}:
        return max(profile.text_ratio, profile.confidence)
    if role is DataRole.CALCULATED:
        return max(profile.formula_ratio, profile.confidence)
    return profile.confidence


@dataclass
class CellRecord:
    address: str
    row: int
    column: int
    letter: str
    value: Any
    formula: str | None = None
    data_type: str = ""
    number_format: str = ""
    merged_anchor: str | None = None
    merged_range: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "address": self.address,
            "row": self.row,
            "column": self.column,
            "letter": self.letter,
            "value": _jsonable(self.value),
            "formula": self.formula,
            "data_type": self.data_type,
            "number_format": self.number_format,
            "merged_anchor": self.merged_anchor,
            "merged_range": self.merged_range,
        }


@dataclass
class MergedRangeRecord:
    address: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    anchor_address: str
    anchor_value: Any
    width: int
    height: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "address": self.address,
            "min_row": self.min_row,
            "min_col": self.min_col,
            "max_row": self.max_row,
            "max_col": self.max_col,
            "anchor_address": self.anchor_address,
            "anchor_value": _jsonable(self.anchor_value),
            "width": self.width,
            "height": self.height,
        }


@dataclass
class ColumnModel:
    index: int
    letter: str
    header: str
    header_address: str
    header_rows: list[int]
    original_header: str = ""
    parent_header: str = ""
    sub_header: str = ""
    period: str = ""
    business_role: str = "unknown"
    semantic_role: str = DataRole.UNKNOWN.value
    role_confidence: float = 0.0
    data_type: str = ValueKind.EMPTY.value
    value_count: int = 0
    empty_ratio: float = 0.0
    samples: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "letter": self.letter,
            "header": self.header,
            "header_address": self.header_address,
            "header_rows": list(self.header_rows),
            "original_header": self.original_header,
            "parent_header": self.parent_header,
            "sub_header": self.sub_header,
            "period": self.period,
            "business_role": self.business_role,
            "semantic_role": self.semantic_role,
            "role_confidence": round(self.role_confidence, 4),
            "data_type": self.data_type,
            "value_count": self.value_count,
            "empty_ratio": round(self.empty_ratio, 4),
            "samples": list(self.samples),
            "notes": list(self.notes),
        }


@dataclass
class BlockModel:
    block_id: str
    sheet: str
    kind: str
    title: str
    boundaries: dict[str, Any]
    header_rows: list[int]
    columns: list[ColumnModel]
    merged_ranges: list[MergedRangeRecord]
    cells: list[CellRecord]
    periods: list[dict[str, Any]]
    period_evidence: list[dict[str, Any]]
    period_status: str
    period_confidence: float
    period_conflicts: list[dict[str, Any]]
    confidence: float
    reasons: list[str] = field(default_factory=list)
    cells_truncated: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "block_id": self.block_id,
            "sheet": self.sheet,
            "kind": self.kind,
            "title": self.title,
            "boundaries": dict(self.boundaries),
            "header_rows": list(self.header_rows),
            "columns": [column.to_dict() for column in self.columns],
            "merged_ranges": [item.to_dict() for item in self.merged_ranges],
            "cells": [cell.to_dict() for cell in self.cells],
            "cells_truncated": self.cells_truncated,
            "periods": list(self.periods),
            "period_evidence": list(self.period_evidence),
            "period_status": self.period_status,
            "period_confidence": round(self.period_confidence, 3),
            "period_conflicts": list(self.period_conflicts),
            "confidence": round(self.confidence, 4),
            "reasons": list(self.reasons),
        }


@dataclass
class SheetModel:
    name: str
    title: str
    blocks: list[BlockModel]
    merged_ranges: list[MergedRangeRecord]
    periods: list[dict[str, Any]]
    period_evidence: list[dict[str, Any]]
    period_status: str
    period_confidence: float
    period_conflicts: list[dict[str, Any]]
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "title": self.title,
            "blocks": [block.to_dict() for block in self.blocks],
            "merged_ranges": [item.to_dict() for item in self.merged_ranges],
            "periods": list(self.periods),
            "period_evidence": list(self.period_evidence),
            "period_status": self.period_status,
            "period_confidence": round(self.period_confidence, 3),
            "period_conflicts": list(self.period_conflicts),
            "warnings": list(self.warnings),
        }


@dataclass
class WorkbookModel:
    source_filename: str
    source_path: str
    sheets: list[SheetModel]
    periods: list[dict[str, Any]]
    period_evidence: list[dict[str, Any]]
    period_status: str
    period_confidence: float
    warnings: list[str] = field(default_factory=list)

    def sheet(self, name: str) -> SheetModel | None:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None

    def block(self, block_id: str) -> BlockModel | None:
        for sheet in self.sheets:
            for block in sheet.blocks:
                if block.block_id == block_id:
                    return block
        return None

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_filename": self.source_filename,
            "source_path": self.source_path,
            "sheets": [sheet.to_dict() for sheet in self.sheets],
            "periods": list(self.periods),
            "period_evidence": list(self.period_evidence),
            "period_status": self.period_status,
            "period_confidence": round(self.period_confidence, 3),
            "warnings": list(self.warnings),
        }


class ExtractionModelBuilder:
    """Project detected tables into the structured extraction model."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def build(
        self,
        *,
        source_path: Path | str,
        sheet_names: Iterable[str],
        grids: dict[str, SheetGrid],
        structures: dict[str, Any],
        intelligence: Any = None,
        max_cells_per_block: int = DEFAULT_MAX_CELLS_PER_BLOCK,
    ) -> WorkbookModel:
        path = Path(source_path)
        sheets: list[SheetModel] = []
        workbook_sources: list[PeriodSource] = [
            PeriodSource(SOURCE_FILENAME, path.name, address=path.name)
        ]
        warnings: list[str] = []
        for name in sheet_names:
            grid = grids.get(name)
            if grid is None:
                continue
            structure = structures.get(name)
            tables = list(getattr(structure, "tables", []) or [])
            tables.extend(getattr(structure, "key_value_tables", []) or [])
            sheet_intel = intelligence.sheet(name) if intelligence is not None else None
            sheet_model, sheet_sources = self._build_sheet(
                grid, name, tables, sheet_intel, path.name, max_cells_per_block
            )
            sheets.append(sheet_model)
            workbook_sources.extend(sheet_sources)
        detection = detect_periods(self._unique_sources(workbook_sources))
        if not sheets:
            warnings.append("هیچ برگه‌ای برای ساخت مدل استخراج پیدا نشد.")
        return WorkbookModel(
            source_filename=path.name,
            source_path=str(path),
            sheets=sheets,
            periods=[item.to_dict() for item in detection.periods],
            period_evidence=[item.to_dict() for item in detection.evidence],
            period_status=detection.status,
            period_confidence=detection.confidence,
            warnings=warnings,
        )

    # ------------------------------------------------------------------ #
    @staticmethod
    def _unique_sources(sources: Iterable[PeriodSource]) -> list[PeriodSource]:
        seen: set[tuple[str, str, str]] = set()
        unique: list[PeriodSource] = []
        for source in sources:
            key = (source.source, source.text, source.address)
            if key in seen:
                continue
            seen.add(key)
            unique.append(source)
        return unique

    def _build_sheet(
        self,
        grid: SheetGrid,
        name: str,
        tables: list[TableCandidate],
        sheet_intel: Any,
        filename: str,
        max_cells_per_block: int,
    ) -> tuple[SheetModel, list[PeriodSource]]:
        title = str(getattr(sheet_intel, "title", "") or "")
        blocks: list[BlockModel] = []
        sheet_sources: list[PeriodSource] = [
            PeriodSource(SOURCE_SHEET_NAME, name, address=name),
            PeriodSource(SOURCE_FILENAME, filename, address=filename),
        ]
        if title:
            sheet_sources.append(PeriodSource(SOURCE_BLOCK_TITLE, title))

        for index, table in enumerate(tables, start=1):
            block_id = f"{name}#B{index}"
            block_sources = self._block_sources(grid, table, filename)
            block, own_sources = self._build_block(
                grid, table, block_id, block_sources, max_cells_per_block
            )
            blocks.append(block)
            sheet_sources.extend(own_sources)
        blocks = self._drop_contained_blocks(blocks)

        detection = detect_periods(self._unique_sources(sheet_sources))
        conflicts = self._sheet_conflicts(name, title, filename)
        if conflicts:
            detection.status = STATUS_CONFLICTING
        return (
            SheetModel(
                name=name,
                title=title,
                blocks=blocks,
                merged_ranges=self._merged_ranges(grid, 1, grid.max_row, 1, grid.max_col),
                periods=[item.to_dict() for item in detection.periods],
                period_evidence=[item.to_dict() for item in detection.evidence],
                period_status=detection.status,
                period_confidence=detection.confidence,
                period_conflicts=conflicts,
                warnings=self._sheet_warnings(grid, blocks, detection),
            ),
            sheet_sources,
        )

    @staticmethod
    def _drop_contained_blocks(blocks: list[BlockModel]) -> list[BlockModel]:
        """Keep the wider table when two candidates cover the same cells."""
        kept: list[BlockModel] = []
        for block in blocks:
            bounds = block.boundaries
            contained = any(
                other is not block
                and other.boundaries["header_row"] <= bounds["header_row"]
                and other.boundaries["last_data_row"] >= bounds["last_data_row"]
                and other.boundaries["first_col_index"] <= bounds["first_col_index"]
                and other.boundaries["last_col_index"] >= bounds["last_col_index"]
                and (
                    other.boundaries["first_col_index"],
                    other.boundaries["last_col_index"],
                )
                != (bounds["first_col_index"], bounds["last_col_index"])
                for other in blocks
            )
            if not contained:
                kept.append(block)
        return kept

    def _block_sources(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        filename: str,
    ) -> list[PeriodSource]:
        sources: list[PeriodSource] = [
            PeriodSource(SOURCE_SHEET_NAME, table.sheet, address=table.sheet),
            PeriodSource(SOURCE_FILENAME, filename, address=filename),
        ]
        if table.title:
            sources.append(
                PeriodSource(
                    SOURCE_BLOCK_TITLE,
                    table.title,
                    address=self._title_address(grid, table),
                )
            )
        for column in table.columns:
            text = " ".join(
                part for part in (column.parent_header, column.sub_header, column.period) if part
            )
            if text:
                sources.append(
                    PeriodSource(SOURCE_COLUMN_HEADER, text, address=column.header_address)
                )
        sources.extend(self._row_label_sources(grid, table))
        return sources

    @staticmethod
    def _title_address(grid: SheetGrid, table: TableCandidate) -> str:
        target = normalize_text(normalize_digits(table.title))
        if not target:
            return ""
        for row in range(max(1, table.header_row - 5), table.header_row):
            for col in range(1, grid.max_col + 1):
                cell = grid.get(row, col)
                if cell is None or not isinstance(cell.value, str):
                    continue
                if normalize_text(normalize_digits(cell.value)) == target:
                    return _cell_address(row, col)
        return ""

    @staticmethod
    def _row_label_sources(grid: SheetGrid, table: TableCandidate) -> list[PeriodSource]:
        sources: list[PeriodSource] = []
        limit = min(
            table.last_data_row,
            table.first_data_row + DEFAULT_MAX_ROW_LABEL_SCAN - 1,
        )
        for row in range(table.first_data_row, limit + 1):
            for col in range(table.first_col, min(table.last_col, table.first_col + 2) + 1):
                cell = grid.get(row, col)
                if cell is None or cell.is_empty:
                    continue
                if not isinstance(cell.value, str):
                    continue
                text = cell.value.strip()
                if not text or looks_like_date(text) or looks_like_number(text):
                    continue
                sources.append(PeriodSource(SOURCE_ROW_LABEL, text, address=_cell_address(row, col)))
                break
        return sources

    # ------------------------------------------------------------------ #
    def _build_block(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        block_id: str,
        sources: list[PeriodSource],
        max_cells: int,
    ) -> tuple[BlockModel, list[PeriodSource]]:
        header_rows = list(table.header_rows or [table.header_row])
        detection = detect_periods(sources)
        conflicts = scalar_conflicts(detection, dimension=KIND_YEAR)
        conflicts.extend(scalar_conflicts(detection, dimension=KIND_MONTH))
        if conflicts:
            detection.status = STATUS_CONFLICTING
        own_sources = [
            source
            for source in sources
            if source.source in {SOURCE_COLUMN_HEADER, SOURCE_ROW_LABEL, SOURCE_BLOCK_TITLE}
        ]
        columns = [
            ColumnModel(
                index=column.index,
                letter=column.letter,
                header=column.header,
                header_address=column.header_address,
                header_rows=header_rows,
                original_header=column.original_header,
                parent_header=column.parent_header,
                sub_header=column.sub_header,
                period=column.period,
                business_role=column.role.value,
                semantic_role=column.statistical_role.value,
                role_confidence=_role_confidence(column),
                data_type=column.data_type.value,
                value_count=column.non_empty,
                empty_ratio=column.empty_ratio,
                samples=list(column.samples),
                notes=list(column.notes),
            )
            for column in table.columns
        ]
        cells, truncated = self._cells(grid, table, header_rows, max_cells)
        block = BlockModel(
            block_id=block_id,
            sheet=table.sheet,
            kind=table.kind.value,
            title=table.title,
            boundaries={
                "address": table.address,
                "header_row": table.header_row,
                "first_col": get_column_letter(table.first_col),
                "last_col": get_column_letter(table.last_col),
                "first_col_index": table.first_col,
                "last_col_index": table.last_col,
                "first_data_row": table.first_data_row,
                "last_data_row": table.last_data_row,
                "header_rows": header_rows,
                "row_count": table.row_count,
                "column_count": len(table.columns),
            },
            header_rows=header_rows,
            columns=columns,
            merged_ranges=self._merged_ranges(
                grid,
                min(header_rows),
                table.last_data_row,
                table.first_col,
                table.last_col,
            ),
            cells=cells,
            periods=[item.to_dict() for item in detection.periods],
            period_evidence=[item.to_dict() for item in detection.evidence],
            period_status=detection.status,
            period_confidence=detection.confidence,
            period_conflicts=conflicts,
            confidence=table.confidence,
            reasons=list(table.reasons),
            cells_truncated=truncated,
        )
        return block, own_sources

    @staticmethod
    def _cells(
        grid: SheetGrid,
        table: TableCandidate,
        header_rows: list[int],
        max_cells: int,
    ) -> tuple[list[CellRecord], bool]:
        records: list[CellRecord] = []
        truncated = False
        first_row = min(header_rows)
        for row in range(first_row, table.last_data_row + 1):
            for col in range(table.first_col, table.last_col + 1):
                cell = grid.get(row, col)
                if cell is None:
                    continue
                if len(records) >= max_cells:
                    return records, True
                span = grid.merged_span(row, col)
                anchor = grid.merged_map.get((row, col))
                records.append(
                    CellRecord(
                        address=_cell_address(row, col),
                        row=row,
                        column=col,
                        letter=get_column_letter(col),
                        value=cell.value,
                        formula=cell.formula,
                        data_type=cell.data_type,
                        number_format=cell.number_format,
                        merged_anchor=_cell_address(*anchor) if anchor else None,
                        merged_range=(
                            _range_address(*span) if span is not None else None
                        ),
                    )
                )
        return records, truncated

    @staticmethod
    def _merged_ranges(
        grid: SheetGrid,
        first_row: int,
        last_row: int,
        first_col: int,
        last_col: int,
    ) -> list[MergedRangeRecord]:
        records: list[MergedRangeRecord] = []
        for min_row, min_col, max_row, max_col in grid.merged_ranges:
            if max_row < first_row or min_row > last_row:
                continue
            if max_col < first_col or min_col > last_col:
                continue
            records.append(
                MergedRangeRecord(
                    address=_range_address(min_row, min_col, max_row, max_col),
                    min_row=min_row,
                    min_col=min_col,
                    max_row=max_row,
                    max_col=max_col,
                    anchor_address=_cell_address(min_row, min_col),
                    anchor_value=grid.value(min_row, min_col),
                    width=max_col - min_col + 1,
                    height=max_row - min_row + 1,
                )
            )
        return records

    @staticmethod
    def _sheet_conflicts(
        name: str,
        title: str,
        filename: str,
    ) -> list[dict[str, Any]]:
        scalar = detect_periods(
            [
                PeriodSource(SOURCE_SHEET_NAME, name),
                PeriodSource(SOURCE_FILENAME, filename),
                *([PeriodSource(SOURCE_BLOCK_TITLE, title)] if title else []),
            ]
        )
        conflicts = scalar_conflicts(scalar, dimension=KIND_YEAR)
        conflicts.extend(scalar_conflicts(scalar, dimension=KIND_MONTH))
        return conflicts

    @staticmethod
    def _sheet_warnings(
        grid: SheetGrid,
        blocks: list[BlockModel],
        detection: PeriodDetection,
    ) -> list[str]:
        warnings: list[str] = []
        if not blocks and grid.max_row:
            warnings.append("هیچ بلوک جدولی در این برگه شناسایی نشد.")
        if detection.status != STATUS_DETECTED:
            warnings.append("دوره‌ای با شواهد کافی برای این برگه شناسایی نشد؛ دوره‌ای به‌صورت حدسی ثبت نشد.")
        for block in blocks:
            if block.cells_truncated:
                warnings.append(
                    f"سلول‌های بلوک {block.block_id} برای حجم خروجی محدود شدند؛ مرزها و ستون‌ها کامل ثبت شده‌اند."
                )
        return warnings
