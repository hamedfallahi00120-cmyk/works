"""Sheet-level structural analysis.

Turns a worksheet into a grid, splits it into content regions, scores candidate
header rows, profiles every column (type, fill ratio, formulas) and infers the
semantic role of each column.  Nothing here knows about loans specifically: the
same code must work for a bank statement, an expense list or a payroll sheet.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from local_office_ai.core.models import ColumnProfile, ColumnRole, DataRole, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    looks_like_date,
    looks_like_number,
    normalize_digits,
    normalize_text,
    parse_value,
)

TOTAL_KEYWORDS = ("جمع", "مجموع", "کل", "total", "sum", "grand total", "جمع کل")


def is_total_label(text: object) -> bool:
    """True when a cell text is a totals label (``جمع کل``, ``total`` ...)."""
    if not isinstance(text, str):
        return False
    value = normalize_text(normalize_digits(text)).lower().strip(" :.-")
    return any(keyword in value for keyword in TOTAL_KEYWORDS)
INSTITUTION_TOKENS = (
    "بانک", "صندوق", "مؤسسه", "موسسه", "شرکت", "بیمه", "لیزینگ", "اعتباری",
    "قرض", "تعاونی", "سرمایه", "بورس", "صرافی", "تعاونی", "دانشگاه", "اداره",
    "سازمان", "شهرداری", "وزارت", "بیمه", "کارگزاری", "هلدینگ",
    "bank", "fund", "company", "co.", "ltd", "llc", "inc", "finance", "credit",
    "insurance", "leasing", "institution", "vendor", "supplier",
)

_ROLE_KEYWORDS: tuple[tuple[ColumnRole, tuple[str, ...]], ...] = (
    (
        ColumnRole.DATE,
        ("تاریخ", "سررسید", "مهلت", "date", "due", "deadline", "تاریخ سررسید"),
    ),
    (
        ColumnRole.PAID,
        (
            "اقساط پرداخت شده", "اقساط پرداختشده", "تعداد اقساط پرداخت شده",
            "تعداد اقساط پرداختشده", "پرداخت شده", "پرداختشده", "تعداد پرداخت",
            "paid installments", "paid", "settled", "installments paid",
        ),
    ),
    (
        ColumnRole.INSTALLMENT_COUNT,
        (
            "تعداد کل قسط", "تعداد اقساط", "تعداد قسط", "تعداد کل اقساط",
            "اقساط کل", "مدت بازپرداخت", "تعداد ماه", "installment count",
            "number of installments", "terms", "tenor", "months",
        ),
    ),
    (
        ColumnRole.INSTALLMENT,
        (
            "مبلغ هر قسط", "مبلغ قسط", "هر قسط", "قسط ماهانه", "اقساط", "قسط",
            "installment", "instalment", "monthly payment", "payment", "emi",
        ),
    ),
    (
        ColumnRole.PRINCIPAL,
        (
            "مبلغ وام", "اصل وام", "مبلغ اصل", "اصل تسهیلات", "اصل",
            "principal", "loan amount", "original amount",
        ),
    ),
    (
        ColumnRole.AMOUNT,
        (
            "مبلغ", "مقدار", "amount", "value", "تعهد", "مانده", "balance",
            "هزینه", "درآمد", "حقوق", "دستمزد", "اجاره", "جمع", "مجموع",
            "total", "sum", "بازپرداخت", "repayment", "بدهی", "طلب", "debit",
            "credit", "cost", "price", "مبلغ ماهانه",
        ),
    ),
    (
        ColumnRole.COUNT,
        ("تعداد", "عدد", "دفعات", "count", "qty", "quantity", "number"),
    ),
    (
        ColumnRole.ID,
        ("ردیف", "کد", "شناسه", "شماره", "id", "code", "no.", "#", "row", "index"),
    ),
    (
        ColumnRole.INSTITUTION,
        (
            "بانک", "صندوق", "مؤسسه", "موسسه", "شرکت", "بیمه", "لیزینگ",
            "اعتباری", "قرض", "تعاونی", "institution", "lender", "creditor",
            "bank", "fund", "طرف حساب", "بستانکار", "طلبکار", "فروشنده",
            "supplier", "vendor", "مشتری", "خریدار", "نام وام", "وام دهنده",
        ),
    ),
    (
        ColumnRole.DESCRIPTION,
        (
            "شرح", "توضیح", "بابت", "موضوع", "عنوان", "description", "note",
            "comment", "detail", "remark",
        ),
    ),
)

def header_matches(header: str, keyword: str) -> bool:
    """Match a header keyword as a whole word/phrase, not inside another word.

    ``اصل`` must match ``مبلغ اصل`` but not ``فاصله``; ``جمع`` must not match
    ``تجمعی``.  Multi-word keywords are matched as phrases.
    """
    keyword = keyword.strip()
    if not keyword:
        return False
    if " " in keyword:
        return keyword in header
    return re.search(rf"(?<!\w){re.escape(keyword)}(?!\w)", header) is not None


NUMERIC_ROLES = {
    ColumnRole.PRINCIPAL,
    ColumnRole.INSTALLMENT,
    ColumnRole.INSTALLMENT_COUNT,
    ColumnRole.PAID,
    ColumnRole.AMOUNT,
    ColumnRole.COUNT,
    ColumnRole.ID,
}

ROLE_LABELS_FA = {
    ColumnRole.INSTITUTION: "نام مؤسسه",
    ColumnRole.PRINCIPAL: "اصل وام",
    ColumnRole.INSTALLMENT: "مبلغ قسط",
    ColumnRole.INSTALLMENT_COUNT: "تعداد اقساط",
    ColumnRole.PAID: "مبلغ پرداخت‌شده",
    ColumnRole.AMOUNT: "مبلغ",
    ColumnRole.COUNT: "تعداد",
    ColumnRole.ID: "شناسه",
    ColumnRole.DATE: "تاریخ",
    ColumnRole.DESCRIPTION: "توضیح",
    ColumnRole.UNKNOWN: "بدون نقش",
}


@dataclass
class CellInfo:
    row: int
    column: int
    value: Any
    formula: str | None = None
    data_type: str = "n"
    number_format: str = ""
    hidden: bool = False

    @property
    def address(self) -> str:
        return f"{get_column_letter(self.column)}{self.row}"

    @property
    def is_empty(self) -> bool:
        if self.value is None:
            return True
        return isinstance(self.value, str) and not self.value.strip()


@dataclass
class RowBlock:
    start: int
    end: int

    @property
    def size(self) -> int:
        return self.end - self.start + 1


@dataclass
class HeaderCandidate:
    row: int
    first_col: int
    last_col: int
    score: float
    text_cells: int
    numeric_below: float
    reasons: list[str] = field(default_factory=list)
    band: "HeaderBand | None" = None


JALALI_MONTHS = (
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
)
_TITLE_PERIOD = re.compile(r"(?<!\d)(1[34]\d{2}|19\d{2}|20\d{2})(?!\d)")


@dataclass
class HeaderCell:
    column: int
    letter: str
    parent: str = ""
    sub: str = ""
    period: str = ""
    address: str = ""
    values: list[str] = field(default_factory=list)
    merged: bool = False

    @property
    def standard_name(self) -> str:
        if self.parent and self.sub and self.parent != self.sub:
            return f"{self.parent} | {self.sub}"
        return self.parent or self.sub or self.letter


@dataclass
class HeaderBand:
    """One or more header rows, preserving the parent/sub hierarchy."""

    top_row: int
    bottom_row: int
    first_col: int
    last_col: int
    cells: dict[int, HeaderCell] = field(default_factory=dict)
    groups: list[dict[str, Any]] = field(default_factory=list)
    years: list[str] = field(default_factory=list)
    has_merges: bool = False
    reasons: list[str] = field(default_factory=list)

    @property
    def rows(self) -> list[int]:
        return list(range(self.top_row, self.bottom_row + 1))

    @property
    def width(self) -> int:
        return self.last_col - self.first_col + 1

    def to_dict(self) -> dict[str, Any]:
        return {
            "top_row": self.top_row,
            "bottom_row": self.bottom_row,
            "rows": self.rows,
            "first_col": self.first_col,
            "last_col": self.last_col,
            "has_merges": self.has_merges,
            "years": list(self.years),
            "groups": list(self.groups),
            "reasons": list(self.reasons),
            "cells": {
                str(column): {
                    "letter": cell.letter,
                    "parent": cell.parent,
                    "sub": cell.sub,
                    "period": cell.period,
                    "address": cell.address,
                    "standard_name": cell.standard_name,
                }
                for column, cell in sorted(self.cells.items())
            },
        }


def extract_periods(texts: Iterable[object]) -> list[str]:
    """Year/period tokens in a title, sheet name or subheader, in order."""
    found: list[str] = []
    for text in texts:
        if text is None:
            continue
        normalized = normalize_text(normalize_digits(str(text)))
        for match in _TITLE_PERIOD.finditer(normalized):
            token = match.group(1)
            if token not in found:
                found.append(token)
    return found


def looks_like_period(text: object) -> bool:
    if isinstance(text, bool):
        return False
    if isinstance(text, (int, float)):
        value = str(int(text)) if float(text).is_integer() else str(text)
    else:
        value = normalize_text(normalize_digits(str(text))).strip()
    if not value:
        return False
    if re.fullmatch(r"1[34]\d{2}|19\d{2}|20\d{2}", value):
        return True
    if re.fullmatch(r"[Qq][1-4]", value):
        return True
    if re.fullmatch(r"(نیمه|فصل|سه‌ماهه|شش‌ماهه)\s*\d?", value):
        return True
    return False


class SheetGrid:
    """Sparse cell grid limited to the sheet's real data extent."""

    def __init__(
        self,
        sheet_name: str,
        values_ws: Worksheet,
        formulas_ws: Worksheet | None,
        max_row: int,
        max_col: int,
    ) -> None:
        self.sheet_name = sheet_name
        self.values_ws = values_ws
        self.formulas_ws = formulas_ws
        self.max_row = max(max_row, 0)
        self.max_col = max(max_col, 0)
        self.cells: dict[tuple[int, int], CellInfo] = {}
        self.hidden_rows = {
            int(index)
            for index, dimension in values_ws.row_dimensions.items()
            if getattr(dimension, "hidden", False)
        }
        self.hidden_cols = {
            column_index_from_string(str(index))
            for index, dimension in values_ws.column_dimensions.items()
            if getattr(dimension, "hidden", False) and str(index).isalpha()
        }
        self.merged_anchors: set[tuple[int, int]] = set()
        self.merged_covered: set[tuple[int, int]] = set()
        self.merged_map: dict[tuple[int, int], tuple[int, int]] = {}
        self.merged_ranges: list[tuple[int, int, int, int]] = []
        for merged in values_ws.merged_cells.ranges:
            self.merged_anchors.add((merged.min_row, merged.min_col))
            self.merged_ranges.append((merged.min_row, merged.min_col, merged.max_row, merged.max_col))
            for row in range(merged.min_row, merged.max_row + 1):
                for col in range(merged.min_col, merged.max_col + 1):
                    if (row, col) != (merged.min_row, merged.min_col):
                        self.merged_covered.add((row, col))
                        self.merged_map[(row, col)] = (merged.min_row, merged.min_col)
        self._build()

    def _build(self) -> None:
        if not self.max_row or not self.max_col:
            return
        formula_cells = {}
        if self.formulas_ws is not None:
            for row in self.formulas_ws.iter_rows(
                min_row=1, max_row=self.max_row, max_col=self.max_col
            ):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_cells[(cell.row, cell.column)] = cell.value
        for row in self.values_ws.iter_rows(
            min_row=1, max_row=self.max_row, max_col=self.max_col
        ):
            for cell in row:
                key = (cell.row, cell.column)
                formula = formula_cells.get(key)
                if cell.value is None and formula is None:
                    continue
                if isinstance(cell.value, str) and not cell.value.strip() and formula is None:
                    continue
                self.cells[key] = CellInfo(
                    row=cell.row,
                    column=cell.column,
                    value=cell.value,
                    formula=formula,
                    data_type=str(cell.data_type),
                    number_format=str(getattr(cell, "number_format", "") or ""),
                    hidden=(cell.row in self.hidden_rows or cell.column in self.hidden_cols),
                )

    def get(self, row: int, col: int) -> CellInfo | None:
        return self.cells.get((row, col))

    def value(self, row: int, col: int) -> Any:
        cell = self.cells.get((row, col))
        return None if cell is None else cell.value

    def anchor_value(self, row: int, col: int) -> Any:
        """Value of the merged cell covering ``(row, col)``, else the cell value."""
        anchor = self.merged_map.get((row, col))
        if anchor is None:
            return self.value(row, col)
        return self.value(*anchor)

    def merged_span(self, row: int, col: int) -> tuple[int, int, int, int] | None:
        for min_row, min_col, max_row, max_col in self.merged_ranges:
            if min_row <= row <= max_row and min_col <= col <= max_col:
                return min_row, min_col, max_row, max_col
        return None

    def row_cells(self, row: int, first_col: int | None = None, last_col: int | None = None) -> list[CellInfo]:
        first_col = first_col or 1
        last_col = last_col or self.max_col
        return [
            cell
            for (row_index, col), cell in self.cells.items()
            if row_index == row and first_col <= col <= last_col
        ]

    def row_has_content(self, row: int) -> bool:
        return any(
            not cell.is_empty or cell.formula
            for (row_index, _), cell in self.cells.items()
            if row_index == row
        )

    def row_non_empty_count(self, row: int, first_col: int, last_col: int) -> int:
        return sum(
            1
            for col in range(first_col, last_col + 1)
            if (cell := self.get(row, col)) is not None and (not cell.is_empty or cell.formula)
        )


class SheetAnalyzer:
    def __init__(self, grid: SheetGrid, settings: Settings) -> None:
        self.grid = grid
        self.settings = settings

    # ------------------------------------------------------------------ #
    # Regions and blocks
    # ------------------------------------------------------------------ #
    def regions(self, max_blank_gap: int = 2) -> list[RowBlock]:
        """Rows with content, joining blocks separated by a small blank gap.

        A single blank row often separates a total row from the rows below it,
        or one commitment from the next; joining small gaps keeps those records
        together while a large gap still starts a new region.
        """
        blocks: list[RowBlock] = []
        start: int | None = None
        last_content = 0
        for row in range(1, self.grid.max_row + 1):
            if self.grid.row_has_content(row):
                if start is None:
                    start = row
                elif row - last_content - 1 > max_blank_gap:
                    blocks.append(RowBlock(start, last_content))
                    start = row
                last_content = row
        if start is not None:
            blocks.append(RowBlock(start, last_content))
        return blocks

    def column_groups(
        self,
        row: int,
        gap_tolerance: int = 1,
        region_end: int | None = None,
    ) -> list[tuple[int, int]]:
        """Contiguous groups of non-empty cells in a row (side-by-side tables).

        A single blank column does not split a table when that column holds data
        somewhere below (it is an unlabelled column).  A column that is empty for
        the whole region is a real separator between two tables.
        """
        columns = sorted(
            cell.column
            for cell in self.grid.row_cells(row)
            if not cell.is_empty or cell.formula
        )
        groups: list[tuple[int, int]] = []
        if not columns:
            return groups
        start = previous = columns[0]
        for col in columns[1:]:
            gap = col - previous - 1
            split = gap > gap_tolerance
            if not split and gap == 1 and region_end is not None:
                gap_col = previous + 1
                has_data_below = any(
                    (cell := self.grid.get(candidate, gap_col)) is not None
                    and (not cell.is_empty or cell.formula)
                    for candidate in range(row + 1, region_end + 1)
                )
                split = not has_data_below
            if split:
                groups.append((start, previous))
                start = col
            previous = col
        groups.append((start, previous))
        return groups

    # ------------------------------------------------------------------ #
    # Header detection
    # ------------------------------------------------------------------ #
    def score_header(self, row: int, first_col: int, last_col: int, region_end: int) -> HeaderCandidate | None:
        header_cells = [
            cell
            for col in range(first_col, last_col + 1)
            if (cell := self.grid.get(row, col)) is not None and not cell.is_empty
        ]
        if len(header_cells) < 2:
            return None

        text_like = 0
        numeric_like = 0
        for cell in header_cells:
            value = cell.value
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                numeric_like += 1
            elif isinstance(value, str):
                if looks_like_number(value) or looks_like_date(value):
                    numeric_like += 1
                else:
                    text_like += 1
            else:
                numeric_like += 1
        if text_like < 2:
            return None
        if numeric_like and text_like < 3:
            # A row that is mostly numbers (a totals row or a label/value row) is not a header.
            return None
        for cell in header_cells:
            value = cell.value
            if isinstance(value, (int, float)) and not isinstance(value, bool) and abs(value) > 10_000:
                # A header never carries a real amount; this is a totals row.
                return None

        below_rows = [
            candidate
            for candidate in range(row + 1, min(row + 6, region_end + 1))
            if self.grid.row_non_empty_count(candidate, first_col, last_col) > 0
        ]
        numeric_below = 0.0
        if below_rows:
            ratios = []
            for col in range(first_col, last_col + 1):
                numeric = total = 0
                for candidate in below_rows:
                    cell = self.grid.get(candidate, col)
                    if cell is None or cell.is_empty:
                        continue
                    total += 1
                    parsed = parse_value(cell.value)
                    if parsed.kind is ValueKind.NUMBER:
                        numeric += 1
                if total:
                    ratios.append(numeric / total)
            numeric_below = sum(ratios) / len(ratios) if ratios else 0.0

        non_empty = len(header_cells)
        text_ratio = text_like / non_empty
        row_numeric_ratio = numeric_like / non_empty
        score = 0.5 * text_ratio + 0.6 * numeric_below - 0.4 * row_numeric_ratio + 0.05
        reasons = [
            f"header text cells={text_like}/{non_empty}",
            f"numeric data below={numeric_below:.2f}",
        ]
        if numeric_below < 0.2:
            reasons.append("almost no numeric data below; header is doubtful")
            score -= 0.2
        return HeaderCandidate(
            row=row,
            first_col=first_col,
            last_col=last_col,
            score=max(0.0, min(score, 1.0)),
            text_cells=text_like,
            numeric_below=numeric_below,
            reasons=reasons,
        )

    def _expand_merged_span(self, row: int, first_col: int, last_col: int) -> tuple[int, int]:
        """Grow the header span so a merged parent covering the last child is included."""
        changed = True
        while changed:
            changed = False
            for min_row, min_col, max_row, max_col in self.grid.merged_ranges:
                if not (min_row <= row <= max_row):
                    continue
                if min_col <= last_col + 1 <= max_col and max_col > last_col:
                    last_col = max_col
                    changed = True
                if min_col < first_col <= max_col + 1:
                    first_col = min_col
                    changed = True
        return first_col, last_col

    def find_header(
        self,
        start_row: int,
        region_end: int,
        scan_limit: int | None = None,
    ) -> HeaderCandidate | None:
        """Best header row at or after ``start_row``; skips title rows."""
        limit = scan_limit or self.settings.header_scan_rows
        for row in range(start_row, min(region_end, start_row + limit - 1) + 1):
            best: HeaderCandidate | None = None
            for group_first, group_last in self.column_groups(row, region_end=region_end):
                first_col, last_col = self._expand_merged_span(row, group_first, group_last)
                if last_col - first_col + 1 < self.settings.min_table_columns:
                    continue
                candidate = self.score_header(row, first_col, last_col, region_end)
                if candidate is None or candidate.score < 0.35:
                    continue
                if best is None or candidate.score > best.score + 1e-9:
                    best = candidate
            if best is not None:
                best.band = self.detect_header_band(best, region_end)
                return best
        return None

    def detect_header_band(self, header: HeaderCandidate, region_end: int) -> HeaderBand:
        """Expand a header row into a merged-aware band of 1..N rows.

        A second row is a subheader when it consists of period labels or short
        labels and the row below it is real data.  The evidence requirement
        (merged cells or repeated labels) keeps ordinary flat headers intact.
        """
        rows = [header.row]
        limit = max(1, int(getattr(self.settings, "max_header_rows", 3)))
        while len(rows) < limit:
            candidate_row = rows[-1] + 1
            if candidate_row > region_end:
                break
            if not self._row_is_subheader(candidate_row, header.first_col, header.last_col, region_end):
                break
            rows.append(candidate_row)

        first_col, last_col = header.first_col, header.last_col
        band = HeaderBand(
            top_row=header.row,
            bottom_row=rows[-1],
            first_col=first_col,
            last_col=last_col,
        )
        band.has_merges = any(
            min_row <= header.row <= max_row and (max_col - min_col) >= 1
            for min_row, min_col, max_row, max_col in self.grid.merged_ranges
            if min_col <= last_col and max_col >= first_col
        )
        if len(rows) > 1:
            band.reasons.append(f"header band spans rows {rows[0]}..{rows[-1]}")
        if band.has_merges:
            band.reasons.append("merged cells form the parent header level")

        parent_starts: dict[str, int] = {}
        parent_ends: dict[str, int] = {}
        for col in range(first_col, last_col + 1):
            top_value = self.grid.value(header.row, col)
            top_anchor = self.grid.merged_map.get((header.row, col))
            if top_value is None and top_anchor is not None:
                top_value = self.grid.value(*top_anchor)
            values: list[str] = []
            for row in rows:
                raw = self.grid.value(row, col)
                if raw is None or (isinstance(raw, str) and not raw.strip()):
                    continue
                text = normalize_text(normalize_digits(str(raw))).strip()
                if text and text not in values:
                    values.append(text)
            parent = ""
            if top_value is not None and str(top_value).strip():
                parent = normalize_text(normalize_digits(str(top_value))).strip()
            if not parent and values:
                parent = values[0]
            sub = ""
            for value in values:
                if value != parent:
                    sub = value
                    break
            period = sub if looks_like_period(sub) else (parent if looks_like_period(parent) else "")
            if period and period not in band.years:
                band.years.append(period)
            span = self.grid.merged_span(header.row, col)
            address = f"{get_column_letter(col)}{header.row}"
            if span is not None and span[1] != span[3]:
                address = f"{get_column_letter(span[1])}{span[0]}"
            band.cells[col] = HeaderCell(
                column=col,
                letter=get_column_letter(col),
                parent=parent,
                sub=sub,
                period=period,
                address=address,
                values=values,
                merged=span is not None and (span[1] != span[3] or span[0] != span[2]),
            )
            if parent:
                parent_starts.setdefault(parent, col)
                parent_ends[parent] = col

        for parent, start in parent_starts.items():
            band.groups.append(
                {
                    "parent": parent,
                    "first_col": get_column_letter(start),
                    "last_col": get_column_letter(parent_ends[parent]),
                    "count": parent_ends[parent] - start + 1,
                }
            )
        band.groups.sort(key=lambda item: column_index_from_string(item["first_col"]))
        return band

    def _row_is_subheader(self, row: int, first_col: int, last_col: int, region_end: int) -> bool:
        cells = [
            cell
            for col in range(first_col, last_col + 1)
            if (cell := self.grid.get(row, col)) is not None and not cell.is_empty
        ]
        width = last_col - first_col + 1
        if len(cells) < max(2, int(width * 0.3)):
            return False
        has_merge = any(
            min_row <= row - 1 <= max_row and (max_col - min_col) >= 1
            for min_row, min_col, max_row, max_col in self.grid.merged_ranges
        )
        periods = 0
        for cell in cells:
            value = cell.value
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                if not looks_like_period(value) and (abs(float(value)) > 31 or not float(value).is_integer()):
                    return False
            elif isinstance(value, str):
                if looks_like_date(value):
                    return False
                if looks_like_number(value) and not looks_like_period(value):
                    return False
            else:
                return False
            if looks_like_period(value):
                periods += 1
        if periods / len(cells) < 0.4 and not has_merge:
            return False
        # The row below must look like data, not another header row.
        below_rows = [
            candidate
            for candidate in range(row + 1, min(row + 3, region_end + 1))
            if self.grid.row_non_empty_count(candidate, first_col, last_col) > 0
        ]
        if not below_rows:
            return False
        numeric = total = 0
        for col in range(first_col, last_col + 1):
            for candidate in below_rows:
                cell = self.grid.get(candidate, col)
                if cell is None or cell.is_empty:
                    continue
                total += 1
                if parse_value(cell.value).kind is ValueKind.NUMBER:
                    numeric += 1
        if total and numeric / total < 0.3:
            return False
        # Evidence: a horizontal merge in the row above, or repeated sublabels.
        values = [normalize_text(normalize_digits(str(cell.value))) for cell in cells]
        repeated = len(set(values)) < len(values) * 0.8
        return has_merge or repeated

    def profile_columns_from_band(
        self,
        band: HeaderBand,
        first_data_row: int,
        last_data_row: int,
    ) -> list[ColumnProfile]:
        profiles = self.profile_columns(
            band.top_row, band.first_col, band.last_col, first_data_row, last_data_row
        )
        for profile in profiles:
            cell = band.cells.get(profile.index)
            if cell is None:
                continue
            profile.original_header = cell.parent if not cell.sub else f"{cell.parent} / {cell.sub}"
            profile.header = cell.standard_name
            profile.header_address = cell.address
            profile.parent_header = cell.parent
            profile.sub_header = cell.sub
            profile.period = cell.period
            profile.source_address = f"{cell.letter}{band.top_row}"
            if band.bottom_row > band.top_row:
                profile.source_address = f"{cell.letter}{band.top_row}:{cell.letter}{band.bottom_row}"
                if cell.sub:
                    profile.header_address = f"{cell.letter}{band.bottom_row}"
                if cell.merged:
                    profile.source_address = (
                        f"عنوان والد از ادغام ردیف {band.top_row} + زیرعنوان {cell.letter}{band.bottom_row}"
                    )
            if cell.merged and cell.parent:
                profile.notes.append(
                    "عنوان والد از سلول ادغام‌شده خوانده شد و به همه ستون‌های پوشیده‌شده تعمیم داده شد."
                )
        return profiles

    # ------------------------------------------------------------------ #
    # Column profiling and roles
    # ------------------------------------------------------------------ #
    def profile_columns(
        self,
        header_row: int,
        first_col: int,
        last_col: int,
        first_data_row: int,
        last_data_row: int,
    ) -> list[ColumnProfile]:
        profiles: list[ColumnProfile] = []
        total_rows = max(1, last_data_row - first_data_row + 1)
        for col in range(first_col, last_col + 1):
            header_cell = self.grid.get(header_row, col)
            header_text = "" if header_cell is None else str(header_cell.value or "").strip()
            counts = {"number": 0, "date": 0, "text": 0, "empty": 0, "error": 0, "boolean": 0}
            formulas = 0
            samples: list[str] = []
            distinct: set[str] = set()
            for row in range(first_data_row, last_data_row + 1):
                cell = self.grid.get(row, col)
                if cell is None or cell.is_empty:
                    counts["empty"] += 1
                    continue
                if cell.formula:
                    formulas += 1
                parsed = parse_value(cell.value)
                if parsed.kind is ValueKind.ERROR:
                    counts["error"] += 1
                elif parsed.kind is ValueKind.NUMBER:
                    counts["number"] += 1
                elif parsed.kind is ValueKind.DATE:
                    counts["date"] += 1
                elif parsed.kind is ValueKind.BOOLEAN:
                    counts["boolean"] += 1
                else:
                    counts["text"] += 1
                text = str(cell.value).strip()
                distinct.add(normalize_text(normalize_digits(text))[:80])
                if len(samples) < 5 and text not in samples:
                    samples.append(text[:60])

            non_empty = total_rows - counts["empty"]
            if non_empty <= 0:
                data_type = ValueKind.EMPTY
                confidence = 0.0
            else:
                numeric_ratio = counts["number"] / non_empty
                date_ratio = counts["date"] / non_empty
                text_ratio = counts["text"] / non_empty
                if numeric_ratio >= 0.6:
                    data_type, confidence = ValueKind.NUMBER, numeric_ratio
                elif date_ratio >= 0.6:
                    data_type, confidence = ValueKind.DATE, date_ratio
                elif text_ratio >= 0.6:
                    data_type, confidence = ValueKind.TEXT, text_ratio
                elif counts["error"] / non_empty >= 0.5:
                    data_type, confidence = ValueKind.ERROR, counts["error"] / non_empty
                else:
                    dominant = max(counts["number"], counts["date"], counts["text"])
                    data_type = ValueKind.TEXT
                    confidence = dominant / non_empty
            profiles.append(
                ColumnProfile(
                    index=col,
                    letter=get_column_letter(col),
                    header=header_text,
                    header_address=f"{get_column_letter(col)}{header_row}",
                    data_type=data_type,
                    numeric_ratio=counts["number"] / non_empty if non_empty else 0.0,
                    date_ratio=counts["date"] / non_empty if non_empty else 0.0,
                    text_ratio=counts["text"] / non_empty if non_empty else 0.0,
                    empty_ratio=counts["empty"] / total_rows,
                    formula_ratio=formulas / total_rows,
                    non_empty=non_empty,
                    distinct=len(distinct),
                    samples=samples,
                    confidence=confidence,
                )
            )
        return profiles

    def assign_roles(self, profiles: list[ColumnProfile]) -> list[str]:
        """Assign one semantic role per column using header *and* content.

        Returns human-readable notes whenever the header and the real cell
        content disagree, so the operator sees why a role was rejected instead
        of silently trusting a misleading header.
        """
        warnings: list[str] = []
        for profile in profiles:
            header = normalize_text(normalize_digits(profile.header)).lower()
            if not header:
                continue
            for role, keywords in _ROLE_KEYWORDS:
                if any(header_matches(header, keyword) for keyword in keywords):
                    if role in NUMERIC_ROLES and profile.numeric_ratio < 0.5:
                        warnings.append(
                            f"ستون «{profile.header}» عنوان نقش {ROLE_LABELS_FA.get(role, role.value)} دارد اما "
                            f"{profile.numeric_ratio:.0%} مقادیر آن عددی است؛ نقش نادیده گرفته شد."
                        )
                        continue
                    if role is ColumnRole.DATE and profile.date_ratio < 0.5:
                        # A date header over integers 1..31 is commonly a day-of-month
                        # key.  Keep it as a temporal key instead of a measure.
                        if profile.numeric_ratio >= 0.8:
                            profile.role = ColumnRole.DATE
                            profile.notes.append(
                                "مقادیر عددی ۱ تا ۳۱ به‌عنوان شماره روز ماه (کلید زمانی) خوانده شد، نه متغیر اندازه‌گیری."
                            )
                            profile.confidence = max(profile.confidence, 0.65)
                            break
                        warnings.append(
                            f"ستون «{profile.header}» عنوان تاریخ دارد اما "
                            f"{profile.date_ratio:.0%} مقادیر آن تاریخ است؛ نقش نادیده گرفته شد."
                        )
                        continue
                    if role is ColumnRole.INSTITUTION and profile.text_ratio < 0.5:
                        warnings.append(
                            f"ستون «{profile.header}» عنوان نام مؤسسه دارد اما "
                            f"{profile.text_ratio:.0%} مقادیر آن متنی است؛ نقش نادیده گرفته شد."
                        )
                        continue
                    if role is ColumnRole.DESCRIPTION and profile.text_ratio < 0.5:
                        continue
                    profile.role = role
                    break

        # A text column that holds institution names is the institution column,
        # even when its header is a bad description such as "اقساط قابل پرداخت".
        if not any(profile.role is ColumnRole.INSTITUTION for profile in profiles):
            best: tuple[float, ColumnProfile] | None = None
            for profile in profiles:
                if profile.text_ratio < 0.5 or profile.role in {ColumnRole.DESCRIPTION, ColumnRole.DATE}:
                    continue
                ratio = self._institution_token_ratio(profile.samples)
                if ratio >= 0.3 and (best is None or ratio > best[0]):
                    best = (ratio, profile)
            if best is not None:
                best[1].role = ColumnRole.INSTITUTION
                best[1].confidence = max(best[1].confidence, 0.6)
                warnings.append(
                    f"ستون «{best[1].header or best[1].letter}» بر اساس محتوای متنی به‌عنوان "
                    "نام مؤسسه تفسیر شد، نه فقط بر اساس عنوان ستون."
                )

        # Keep only one column per exclusive role; duplicates become AMOUNT/UNKNOWN.
        # A hierarchical header with distinct parent/period names is a legitimate
        # repetition (e.g. the same measure for 1403 and 1404), not a duplicate.
        seen: dict[ColumnRole, ColumnProfile] = {}
        for profile in profiles:
            role = profile.role
            if role in {ColumnRole.UNKNOWN}:
                continue
            if profile.parent_header or profile.sub_header or profile.period:
                continue
            if role in seen:
                other = seen[role]
                if profile.confidence > other.confidence:
                    other.role = ColumnRole.UNKNOWN
                    seen[role] = profile
                    demoted = other
                else:
                    profile.role = ColumnRole.UNKNOWN
                    demoted = profile
                warnings.append(
                    f"ستون «{demoted.header or demoted.letter}» نقش تکراری {ROLE_LABELS_FA.get(role, role.value)} داشت و "
                    "به‌عنوان ستون بدون نقش ثبت شد."
                )
            else:
                seen[role] = profile
        return warnings

    def assign_statistical_roles(
        self,
        profiles: list[ColumnProfile],
        first_data_row: int,
        last_data_row: int,
    ) -> list[str]:
        """Map business roles to statistical roles and exclude keys from measures."""
        notes: list[str] = []
        for profile in profiles:
            numbers = self._numeric_values(profile, first_data_row, last_data_row)
            if profile.role is ColumnRole.ID:
                profile.statistical_role = DataRole.IDENTIFIER
            elif profile.role is ColumnRole.DATE:
                if profile.date_ratio >= 0.5:
                    profile.statistical_role = DataRole.TEMPORAL_KEY
                elif self._is_day_of_month(numbers):
                    profile.statistical_role = DataRole.TEMPORAL_KEY
                    profile.notes.append(
                        "کلید زمانی روزِ ماه؛ در میانگین، واریانس و همبستگی وارد نمی‌شود."
                    )
                else:
                    profile.statistical_role = DataRole.UNKNOWN
            elif profile.role in NUMERIC_ROLES:
                profile.statistical_role = DataRole.NUMERIC_MEASURE
            elif profile.text_ratio >= 0.5:
                if 0 < profile.distinct <= max(20, int(profile.non_empty * 0.5)):
                    profile.statistical_role = DataRole.CATEGORICAL
                else:
                    profile.statistical_role = DataRole.TEXT
            elif profile.non_empty == 0:
                profile.statistical_role = DataRole.UNKNOWN
            elif profile.data_type is ValueKind.NUMBER:
                if self._is_sequential_identifier(numbers):
                    profile.statistical_role = DataRole.IDENTIFIER
                    profile.notes.append("دنباله یکتا و صعودی ۱..n؛ شناسه ردیف است.")
                else:
                    profile.statistical_role = DataRole.NUMERIC_MEASURE
            if profile.formula_ratio >= 0.5:
                profile.statistical_role = DataRole.CALCULATED
            if profile.statistical_role is DataRole.TEMPORAL_KEY and profile.role is not ColumnRole.DATE:
                notes.append(f"ستون «{profile.header or profile.letter}» کلید زمانی است.")
        return notes

    def _numeric_values(
        self,
        profile: ColumnProfile,
        first_data_row: int,
        last_data_row: int,
    ) -> list[float]:
        values: list[float] = []
        for row in range(first_data_row, last_data_row + 1):
            cell = self.grid.get(row, profile.index)
            if cell is None or cell.is_empty:
                continue
            parsed = parse_value(cell.value)
            if parsed.kind is ValueKind.NUMBER and isinstance(parsed.value, (int, float)):
                values.append(float(parsed.value))
            elif parsed.kind is ValueKind.NUMBER:
                try:
                    values.append(float(parsed.value))
                except (TypeError, ValueError):
                    continue
        return values

    @staticmethod
    def _is_day_of_month(values: list[float]) -> bool:
        if len(values) < 3:
            return False
        if not all(float(value).is_integer() and 1 <= value <= 31 for value in values):
            return False
        return (max(values) - min(values)) >= 5

    @staticmethod
    def _is_sequential_identifier(values: list[float]) -> bool:
        if len(values) < 3:
            return False
        if any(not float(value).is_integer() for value in values):
            return False
        ordered = [int(value) for value in values]
        return ordered == list(range(ordered[0], ordered[0] + len(ordered)))

    @staticmethod
    def _institution_token_ratio(samples: Iterable[str]) -> float:
        samples = [sample for sample in samples if sample]
        if not samples:
            return 0.0
        hits = 0
        for sample in samples:
            lowered = normalize_text(normalize_digits(sample)).lower()
            if any(token in lowered for token in INSTITUTION_TOKENS):
                hits += 1
        return hits / len(samples)

    def fallback_institution(self, profiles: list[ColumnProfile]) -> ColumnProfile | None:
        """Pick a text column as institution when a loan-like table lacks one."""
        candidates = [
            profile
            for profile in profiles
            if profile.role in {ColumnRole.UNKNOWN, ColumnRole.DESCRIPTION}
            and profile.text_ratio >= 0.5
            and profile.non_empty >= 2
        ]
        if not candidates:
            return None
        return max(candidates, key=lambda profile: (profile.distinct, profile.confidence))
