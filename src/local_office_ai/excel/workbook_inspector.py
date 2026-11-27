"""Workbook inventory: real data extent, hidden/protected state, formulas, errors.

The inspector answers the questions the old version never asked: how many
sheets exist, where the data really ends (as opposed to where formatting ends),
what is hidden, what is protected, which cells hold formulas and which hold
Excel errors.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import SheetMeta, WorkbookMeta
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import EXCEL_ERRORS
from local_office_ai.excel.reader import WorkbookBundle

LOGGER = get_logger(__name__)


def used_bounds(worksheet: Worksheet, max_rows: int, max_cols: int) -> tuple[int, int]:
    """Return (last_row, last_col) that actually contain a value.

    Formatting alone (fills, borders, column widths) must not inflate the data
    range; that was the root cause of the old code scanning ``A1:BA511``.
    """
    last_row = last_col = 0
    cells = getattr(worksheet, "_cells", None)
    if isinstance(cells, dict):
        for (row, col), cell in cells.items():
            if row > max_rows or col > max_cols:
                continue
            value = cell.value
            if value is None:
                continue
            if isinstance(value, str) and not value.strip():
                continue
            if row > last_row:
                last_row = row
            if col > last_col:
                last_col = col
        return last_row, last_col

    row_limit = min(worksheet.max_row or 0, max_rows)
    col_limit = min(worksheet.max_column or 0, max_cols)
    for row in worksheet.iter_rows(min_row=1, max_row=row_limit, max_col=col_limit):
        for cell in row:
            if cell.value is None:
                continue
            if isinstance(cell.value, str) and not cell.value.strip():
                continue
            if cell.row > last_row:
                last_row = cell.row
            if cell.column > last_col:
                last_col = cell.column
    return last_row, last_col


def _hidden_count(dimensions: dict, limit: int) -> int:
    count = 0
    for index, dimension in dimensions.items():
        try:
            index_value = column_index_from_string(str(index)) if str(index).isalpha() else int(index)
        except (TypeError, ValueError):
            continue
        if index_value > limit:
            continue
        if getattr(dimension, "hidden", False):
            count += 1
    return count


def inspect_sheet(
    name: str,
    values_ws: Worksheet,
    formulas_ws: Worksheet | None,
    settings: Settings,
) -> SheetMeta:
    max_rows = settings.max_rows_scanned
    max_cols = settings.max_columns_scanned
    last_row, last_col = used_bounds(values_ws, max_rows, max_cols)
    used_dimensions = f"A1:{get_column_letter(max(last_col, 1))}{max(last_row, 1)}" if last_row else "A1:A1"

    formula_count = 0
    error_cells: list[str] = []
    if last_row and last_col:
        formula_source = formulas_ws if formulas_ws is not None else values_ws
        for row in formula_source.iter_rows(min_row=1, max_row=last_row, max_col=last_col):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formula_count += 1
        for row in values_ws.iter_rows(min_row=1, max_row=last_row, max_col=last_col):
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.strip().upper() in EXCEL_ERRORS:
                    if len(error_cells) < 200:
                        error_cells.append(cell.coordinate)

    merged = len(list(values_ws.merged_cells.ranges))
    hidden_rows = _hidden_count(values_ws.row_dimensions, max(last_row, 1))
    hidden_cols = _hidden_count(values_ws.column_dimensions, max(last_col, 1))
    protected = bool(getattr(values_ws.protection, "sheet", False))

    suspicious: list[str] = []
    state = getattr(values_ws, "sheet_state", "visible") or "visible"
    if state != "visible":
        suspicious.append(f"برگه در حالت «{state}» است (مخفی).")
    if last_row == 0:
        suspicious.append("برگه هیچ داده‌ای ندارد.")
    declared_rows = values_ws.max_row or 0
    declared_cols = values_ws.max_column or 0
    if last_col and declared_cols > max(last_col * 5, last_col + 20):
        suspicious.append(
            f"محدوده فرمت‌شده ({declared_cols} ستون) بسیار بزرگ‌تر از داده واقعی ({last_col} ستون) است."
        )
    if last_row and declared_rows > max(last_row * 5, last_row + 200):
        suspicious.append(
            f"محدوده فرمت‌شده ({declared_rows} ردیف) بسیار بزرگ‌تر از داده واقعی ({last_row} ردیف) است."
        )
    if hidden_rows:
        suspicious.append(f"{hidden_rows} ردیف مخفی وجود دارد.")
    if hidden_cols:
        suspicious.append(f"{hidden_cols} ستون مخفی وجود دارد.")
    if protected:
        suspicious.append("برگه محافظت‌شده است؛ فقط خواندن انجام شد.")
    if merged:
        suspicious.append(f"{merged} محدوده ادغام‌شده وجود دارد.")

    return SheetMeta(
        title=name,
        state=state,
        declared_dimensions=str(values_ws.calculate_dimension()),
        used_dimensions=used_dimensions,
        max_row=last_row,
        max_col=last_col,
        merged_ranges=merged,
        hidden_rows=hidden_rows,
        hidden_cols=hidden_cols,
        protected=protected,
        formula_count=formula_count,
        error_cells=error_cells,
        suspicious=suspicious,
    )


def inspect_workbook(bundle: WorkbookBundle, settings: Settings) -> WorkbookMeta:
    sheets: list[SheetMeta] = []
    for name in bundle.values.sheetnames:
        values_ws = bundle.values[name]
        formulas_ws = bundle.formulas[name] if name in bundle.formulas.sheetnames else None
        try:
            sheets.append(inspect_sheet(name, values_ws, formulas_ws, settings))
        except Exception as exc:  # noqa: BLE001 - one broken sheet must not kill the run
            LOGGER.warning("sheet inspection failed for %r: %s", name, exc)
            sheets.append(
                SheetMeta(
                    title=name,
                    state=getattr(values_ws, "sheet_state", "visible") or "visible",
                    declared_dimensions="unknown",
                    used_dimensions="unknown",
                    max_row=0,
                    max_col=0,
                    suspicious=[f"بازرسی برگه ناموفق بود: {exc}"],
                )
            )

    try:
        defined_names = [str(item) for item in bundle.values.defined_names]
    except Exception:  # noqa: BLE001
        defined_names = []

    external_links = 0
    try:
        external_links = len(getattr(bundle.values, "_external_links", []) or [])
    except Exception:  # noqa: BLE001
        external_links = 0

    meta = WorkbookMeta(
        path=str(Path(bundle.path)),
        sha256=bundle.sha256,
        size_bytes=bundle.size_bytes,
        sheets=sheets,
        defined_names=defined_names,
        has_macros=bundle.has_macros,
        external_links=external_links,
        total_formulas=sum(sheet.formula_count for sheet in sheets),
        total_error_cells=sum(len(sheet.error_cells) for sheet in sheets),
        load_seconds=bundle.load_seconds,
    )
    bundle.meta = meta
    LOGGER.info(
        "workbook inspected: %d sheets, %d formulas, %d error cells",
        len(sheets), meta.total_formulas, meta.total_error_cells,
    )
    return meta
