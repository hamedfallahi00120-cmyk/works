"""Context-aware row classification shared by detection, intelligence and extraction.

A row is never classified from a single substring.  Labels are matched with the
word-boundary helpers in :mod:`sheet_analyzer`, and a totals-looking row is
further classified by its position inside the detected block:

* a totals row with data rows below it is a **subtotal**,
* a totals row at the end of the block is the **grand total**,
* text-only rows inside a block are **group headings** (one filled cell) or
  **category rows** (several text cells),
* explanatory text is a **note**, and a note after the last summary row is a
  **footer**.

The module deliberately keeps structural classification separate from numerical
aggregation.  Callers decide which rows may enter a sum; this module only says
what each row is.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Iterable

from local_office_ai.core.models import ColumnProfile, ColumnRole, DataRole, TableCandidate, ValueKind
from local_office_ai.excel.normalization import (
    looks_like_date,
    looks_like_number,
    normalize_digits,
    normalize_text,
    parse_value,
)
from local_office_ai.excel.sheet_analyzer import SheetGrid, summary_label_role

ROW_HEADER = "header"
ROW_SUBHEADER = "subheader"
ROW_DATA = "data"
ROW_DATA_MISSING = "data_missing"
ROW_TOTAL = "total"
ROW_SUBTOTAL = "subtotal"
ROW_DIFFERENCE = "difference"
ROW_PERCENTAGE = "percentage"
ROW_NOTE = "note"
ROW_FOOTER = "footer"
ROW_TITLE = "title"
ROW_BLANK = "blank"
ROW_CATEGORY = "category"
ROW_GROUP_HEADING = "group_heading"

SUMMARY_ROLES = {
    ROW_TOTAL,
    ROW_SUBTOTAL,
    ROW_DIFFERENCE,
    ROW_PERCENTAGE,
}

ROW_ROLE_FA = {
    ROW_HEADER: "عنوان",
    ROW_SUBHEADER: "عنوان فرعی",
    ROW_DATA: "داده",
    ROW_DATA_MISSING: "دوره بدون مقدار",
    ROW_TOTAL: "جمع کل",
    ROW_SUBTOTAL: "جمع میانی",
    ROW_DIFFERENCE: "مقایسه",
    ROW_PERCENTAGE: "درصد",
    ROW_NOTE: "توضیح",
    ROW_FOOTER: "پانویس",
    ROW_TITLE: "عنوان اصلی",
    ROW_BLANK: "خالی",
    ROW_CATEGORY: "ردیف دسته",
    ROW_GROUP_HEADING: "عنوان گروه",
}


@dataclass
class RowRole:
    row: int
    role: str
    label: str = ""
    has_numbers: bool = False

    @property
    def role_label(self) -> str:
        return ROW_ROLE_FA.get(self.role, self.role)

    def to_dict(self) -> dict[str, Any]:
        return {
            "row": self.row,
            "role": self.role,
            "role_label": self.role_label,
            "label": self.label[:160],
            "has_numbers": self.has_numbers,
        }


def _text(raw: Any) -> str:
    if raw is None or isinstance(raw, bool):
        return ""
    return normalize_text(normalize_digits(str(raw))).strip()


def row_label(grid: SheetGrid, row: int, first_col: int, last_col: int) -> str:
    """The leftmost text cell in a row, skipping numbers and errors."""
    for col in range(first_col, min(last_col, first_col + 5) + 1):
        cell = grid.get(row, col)
        if cell is None or cell.is_empty:
            continue
        value = cell.value
        if not isinstance(value, str):
            continue
        if parse_value(value).kind is ValueKind.NUMBER:
            continue
        if value.strip().upper() in {"#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!"}:
            continue
        text = _text(value)
        if text:
            return text
    for col in range(first_col, last_col + 1):
        cell = grid.get(row, col)
        if cell is not None and isinstance(cell.value, str) and cell.value.strip():
            if parse_value(cell.value).kind is ValueKind.TEXT:
                return _text(cell.value)
    return ""


def row_has_numeric(grid: SheetGrid, row: int, columns: Iterable[ColumnProfile]) -> bool:
    for column in columns:
        cell = grid.get(row, column.index)
        if cell is None:
            continue
        if cell.formula:
            return True
        if cell.is_empty:
            continue
        if parse_value(cell.value).kind is ValueKind.NUMBER:
            return True
    return False


def row_has_value(grid: SheetGrid, row: int, columns: Iterable[ColumnProfile]) -> bool:
    for column in columns:
        cell = grid.get(row, column.index)
        if cell is not None and (not cell.is_empty or cell.formula):
            return True
    return False


def row_filled_count(grid: SheetGrid, row: int, first_col: int, last_col: int) -> int:
    return grid.row_non_empty_count(row, first_col, last_col)


def row_is_text_only(grid: SheetGrid, row: int, first_col: int, last_col: int) -> bool:
    cells = [
        cell
        for col in range(first_col, last_col + 1)
        if (cell := grid.get(row, col)) is not None and (not cell.is_empty or cell.formula)
    ]
    if not cells:
        return False
    for cell in cells:
        value = cell.value
        if not isinstance(value, str):
            return False
        if looks_like_number(value) or looks_like_date(value):
            return False
    return True


def row_role_from_label(label: str, *, has_numbers: bool, long_text: bool = False) -> str:
    value = _text(label).lower().strip(" :.-")
    if not value:
        return ROW_BLANK
    summary = summary_label_role(value)
    if summary == "subtotal":
        return ROW_SUBTOTAL
    if summary == "total":
        return ROW_TOTAL
    if summary == "difference":
        return ROW_DIFFERENCE
    if summary == "percentage":
        return ROW_PERCENTAGE
    if summary == "note" or long_text:
        return ROW_NOTE
    if not has_numbers and len(value) > 60:
        return ROW_FOOTER
    return ROW_DATA


def _measure_columns(table: TableCandidate) -> list[ColumnProfile]:
    return [
        column
        for column in table.columns
        if column.statistical_role in {DataRole.NUMERIC_MEASURE, DataRole.CALCULATED}
        or (
            column.statistical_role is DataRole.UNKNOWN
            and column.data_type is ValueKind.NUMBER
            and column.role not in {ColumnRole.ID, ColumnRole.DATE}
        )
    ]


def _temporal_columns(table: TableCandidate) -> list[ColumnProfile]:
    return [
        column
        for column in table.columns
        if column.statistical_role is DataRole.TEMPORAL_KEY
    ]


def _classify_data_row(
    grid: SheetGrid,
    table: TableCandidate,
    row: int,
    label: str,
    *,
    measure_columns: list[ColumnProfile],
    temporal_columns: list[ColumnProfile],
) -> str:
    summary = summary_label_role(label)
    if summary in {"subtotal", "total"}:
        return ROW_SUBTOTAL
    if summary == "difference":
        return ROW_DIFFERENCE
    if summary == "percentage":
        return ROW_PERCENTAGE
    if summary == "note":
        return ROW_NOTE
    if row_has_numeric(grid, row, measure_columns) or row_has_numeric(
        grid, row, temporal_columns
    ):
        return ROW_DATA
    if label and len(label) > 60:
        return ROW_NOTE
    if row_has_value(grid, row, measure_columns):
        return ROW_DATA_MISSING
    if row_has_value(grid, row, temporal_columns):
        return ROW_DATA_MISSING
    if label:
        filled = row_filled_count(grid, row, table.first_col, table.last_col)
        if filled <= 1:
            return ROW_GROUP_HEADING
        return ROW_CATEGORY
    return ROW_DATA_MISSING


def _classify_summary_row(grid: SheetGrid, row: int, label: str, table: TableCandidate) -> str:
    summary = summary_label_role(label)
    if summary:
        return summary
    if row_has_numeric(grid, row, table.columns):
        return ROW_DATA
    if label:
        return ROW_NOTE
    return ROW_BLANK


def classify_block_rows(
    grid: SheetGrid,
    table: TableCandidate,
    header_rows: list[int] | None = None,
    *,
    block_end_row: int | None = None,
) -> list[RowRole]:
    """Classify every row of one detected block, including its summary rows."""
    headers = list(header_rows or table.header_rows or [table.header_row])
    start = min(headers)
    end = block_end_row if block_end_row is not None else (table.block_end_row or table.last_data_row)
    end = max(end, table.last_data_row)
    measure_columns = _measure_columns(table)
    temporal_columns = _temporal_columns(table)
    records: list[RowRole] = []
    for row in range(start, end + 1):
        label = row_label(grid, row, table.first_col, table.last_col)
        has_numbers = row_has_numeric(grid, row, table.columns)
        if grid.row_non_empty_count(row, table.first_col, table.last_col) == 0:
            role = ROW_BLANK
        elif row in headers:
            role = ROW_HEADER if row == headers[0] else ROW_SUBHEADER
        elif row < table.first_data_row:
            role = ROW_TITLE if label else ROW_BLANK
        elif row <= table.last_data_row:
            role = _classify_data_row(
                grid,
                table,
                row,
                label,
                measure_columns=measure_columns,
                temporal_columns=temporal_columns,
            )
        else:
            role = _classify_summary_row(grid, row, label, table)
        records.append(
            RowRole(row=row, role=role, label=label, has_numbers=has_numbers)
        )
    return records


def summary_rows_from_roles(records: Iterable[RowRole]) -> dict[str, list[int]]:
    grouped: dict[str, list[int]] = defaultdict(list)
    for record in records:
        if record.role in SUMMARY_ROLES:
            grouped[record.role].append(record.row)
    return {key: sorted(value) for key, value in grouped.items()}


def role_counts(records: Iterable[RowRole]) -> dict[str, int]:
    counts: dict[str, int] = defaultdict(int)
    for record in records:
        counts[record.role] += 1
    return dict(counts)
