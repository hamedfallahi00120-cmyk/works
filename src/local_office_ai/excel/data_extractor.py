"""Extraction of records and declared totals from detected tables.

Every extracted value keeps a :class:`Provenance` pointing at its source cell,
the raw text and the standardized value, so the output workbook can prove each
figure instead of asserting it.
"""

from __future__ import annotations

from decimal import Decimal
from typing import Any

from openpyxl.utils import get_column_letter

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import (
    ColumnRole,
    DeclaredTotal,
    LoanRecord,
    NormalizedValue,
    ObligationRecord,
    Provenance,
    TableCandidate,
    TableKind,
    ValueKind,
)
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    normalize_digits,
    normalize_institution,
    normalize_text,
    parse_number,
    parse_value,
)
from local_office_ai.excel.sheet_analyzer import SheetGrid, is_total_label

LOGGER = get_logger(__name__)

REPAYMENT_KEYWORDS = ("بازپرداخت", "repayment", "کل بازپرداخت", "مبلغ بازپرداخت")
TOTAL_SCAN_ROWS = 5


def _provenance(grid: SheetGrid, row: int, col: int) -> Provenance:
    cell = grid.get(row, col)
    raw = None if cell is None else cell.value
    formula = None if cell is None else cell.formula
    return Provenance(
        sheet=grid.sheet_name,
        address=f"{get_column_letter(col)}{row}",
        row=row,
        column=col,
        raw=raw,
        formula=formula,
    )


def _as_decimal(value: Any, *, numeric_context: bool = True) -> tuple[Decimal | None, list[str]]:
    parsed = parse_number(value, numeric_context=numeric_context)
    if parsed.kind is ValueKind.NUMBER and isinstance(parsed.value, Decimal):
        return parsed.value, list(parsed.warnings)
    return None, list(parsed.warnings) if parsed.value is None else []


class DataExtractor:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.synonyms = dict(settings.synonyms)

    # ------------------------------------------------------------------ #
    def extract_loans(self, grid: SheetGrid, table: TableCandidate) -> list[LoanRecord]:
        institution_col = table.column_by_role(ColumnRole.INSTITUTION)
        principal_col = table.column_by_role(ColumnRole.PRINCIPAL)
        installment_col = table.column_by_role(ColumnRole.INSTALLMENT)
        count_col = table.column_by_role(ColumnRole.INSTALLMENT_COUNT) or table.column_by_role(ColumnRole.COUNT)
        paid_col = table.column_by_role(ColumnRole.PAID)
        date_col = table.column_by_role(ColumnRole.DATE)
        repayment_col = None
        for column in table.columns:
            header = normalize_text(normalize_digits(column.header))
            if any(keyword in header for keyword in REPAYMENT_KEYWORDS):
                repayment_col = column
                break

        if institution_col is None:
            return []

        records: list[LoanRecord] = []
        for row in range(table.first_data_row, table.last_data_row + 1):
            if grid.row_non_empty_count(row, table.first_col, table.last_col) == 0:
                continue
            institution_cell = grid.get(row, institution_col.index)
            if institution_cell is None or institution_cell.is_empty:
                continue
            institution_text = normalize_text(normalize_digits(str(institution_cell.value)))
            if is_total_label(institution_text):
                continue

            issues: list[str] = []
            fields: dict[str, NormalizedValue] = {}
            parsed_inst = parse_value(institution_cell.value, kind_hint=ValueKind.TEXT)
            fields["institution"] = parsed_inst

            def numeric(column, name: str) -> Decimal | None:
                if column is None:
                    return None
                cell = grid.get(row, column.index)
                if cell is None or cell.is_empty:
                    issues.append(f"{name}: خالی")
                    return None
                parsed = parse_number(cell.value, numeric_context=True)
                fields[name] = parsed
                issues.extend(f"{name}: {warning}" for warning in parsed.warnings)
                if parsed.kind is not ValueKind.NUMBER:
                    issues.append(f"{name}: مقدار عددی معتبر نیست ({cell.value!r})")
                    return None
                return parsed.value if isinstance(parsed.value, Decimal) else None

            principal = numeric(principal_col, "principal")
            installment = numeric(installment_col, "installment")
            count = numeric(count_col, "installment_count")
            paid = numeric(paid_col, "paid_installments") if paid_col is not None else None
            repayment = numeric(repayment_col, "contract_repayment")

            last_date = None
            if date_col is not None:
                date_cell = grid.get(row, date_col.index)
                if date_cell is not None and not date_cell.is_empty:
                    parsed_date = parse_value(date_cell.value, kind_hint=ValueKind.DATE)
                    fields["last_installment_date"] = parsed_date
                    issues.extend(f"date: {warning}" for warning in parsed_date.warnings)
                    if parsed_date.kind is ValueKind.DATE:
                        last_date = parsed_date.value
                    else:
                        issues.append(f"date: تاریخ قابل خواندن نیست ({date_cell.value!r})")

            if count is not None and count != count.to_integral_value():
                issues.append("installment_count: تعداد قسط اعشاری است")
            if count is not None and count <= 0:
                issues.append("installment_count: تعداد قسط صفر یا منفی است")
            for name, value in (("principal", principal), ("installment", installment)):
                if value is not None and value < 0:
                    issues.append(f"{name}: مبلغ منفی است")
            if paid is not None and count is not None and paid > count:
                issues.append(
                    f"paid_installments: تعداد پرداخت‌شده ({paid}) بیشتر از کل اقساط ({count}) است"
                )

            if installment is not None and count is not None:
                computed_repayment = installment * count
                if repayment is None:
                    repayment = computed_repayment
                elif repayment != computed_repayment:
                    issues.append(
                        "contract_repayment: مقدار ستون بازپرداخت با قسط×تعداد برابر نیست "
                        f"({repayment} ≠ {computed_repayment})"
                    )
            elif repayment is None and installment is not None and count is not None:
                repayment = installment * count

            records.append(
                LoanRecord(
                    institution_raw=str(institution_cell.value).strip(),
                    institution_key=normalize_institution(institution_text, self.synonyms),
                    principal=principal,
                    installment=installment,
                    installment_count=count,
                    last_installment_date=last_date,
                    contract_repayment=repayment,
                    provenance=_provenance(grid, row, institution_col.index),
                    fields=fields,
                    issues=issues,
                    paid_installments=paid,
                )
            )
        LOGGER.info("sheet %r: %d loan records extracted", grid.sheet_name, len(records))
        return records

    # ------------------------------------------------------------------ #
    def extract_obligations(self, grid: SheetGrid, table: TableCandidate) -> list[ObligationRecord]:
        institution_col = table.column_by_role(ColumnRole.INSTITUTION)
        amount_col = (
            table.column_by_role(ColumnRole.AMOUNT)
            or table.column_by_role(ColumnRole.PRINCIPAL)
            or table.column_by_role(ColumnRole.INSTALLMENT)
        )
        if institution_col is None or amount_col is None:
            return []

        records: list[ObligationRecord] = []
        for row in range(table.first_data_row, table.last_data_row + 1):
            label_cell = grid.get(row, institution_col.index)
            amount_cell = grid.get(row, amount_col.index)
            if label_cell is None or label_cell.is_empty:
                continue
            if amount_cell is None or amount_cell.is_empty:
                continue
            label = normalize_text(normalize_digits(str(label_cell.value)))
            if is_total_label(label):
                continue
            amount, warnings = _as_decimal(amount_cell.value)
            if amount is None:
                continue
            records.append(
                ObligationRecord(
                    institution_raw=str(label_cell.value).strip(),
                    institution_key=normalize_institution(label, self.synonyms),
                    amount=amount,
                    provenance=_provenance(grid, row, institution_col.index),
                    label=label,
                    issues=warnings,
                )
            )
        LOGGER.info("sheet %r: %d obligation records extracted", grid.sheet_name, len(records))
        return records

    # ------------------------------------------------------------------ #
    def collect_obligations(
        self,
        candidates: list[TableCandidate],
        grids: dict[str, SheetGrid],
        primary_sheet: str,
    ) -> tuple[list[ObligationRecord], list[str]]:
        """Harvest obligations without double-counting the same commitment.

        Sources on the primary sheet win over other sheets, non-report sheets win
        over generated reports, and an identical (institution, amount) pair seen
        in an earlier source is not counted twice.
        """
        from local_office_ai.excel.table_detector import is_report_sheet

        notes: list[str] = []
        sources = [table for table in candidates if table.kind is TableKind.OBLIGATION]
        preferred = [table for table in sources if not is_report_sheet(table.sheet)]
        if preferred:
            skipped = [table for table in sources if is_report_sheet(table.sheet)]
            if skipped:
                notes.append(
                    "برگه‌های گزارش‌مانند از برداشت تعهدات کنار گذاشته شدند: "
                    + "، ".join(sorted({table.sheet for table in skipped}))
                )
            sources = preferred
        sources.sort(key=lambda table: (table.sheet != primary_sheet, -table.confidence))

        records: list[ObligationRecord] = []
        seen: dict[tuple[str, Decimal], ObligationRecord] = {}
        for table in sources:
            grid = grids.get(table.sheet)
            if grid is None:
                continue
            for record in self.extract_obligations(grid, table):
                key = (record.institution_key, record.amount)
                if key in seen:
                    notes.append(
                        f"تعهد تکراری «{record.institution_raw}» با مبلغ {record.amount} در "
                        f"{record.provenance.sheet}!{record.provenance.address} نادیده گرفته شد؛ "
                        f"نمونه اصلی: {seen[key].provenance.sheet}!{seen[key].provenance.address}"
                    )
                    continue
                seen[key] = record
                records.append(record)
        return records, notes

    # ------------------------------------------------------------------ #
    def find_declared_totals(
        self,
        grid: SheetGrid,
        table: TableCandidate,
    ) -> list[DeclaredTotal]:
        """Read totals written in the sheet, next to or inside the table."""
        totals: list[DeclaredTotal] = []
        start = table.first_data_row
        end = min(grid.max_row, table.last_data_row + TOTAL_SCAN_ROWS)
        for row in range(start, end + 1):
            for col in range(table.first_col, table.last_col + 1):
                cell = grid.get(row, col)
                if cell is None or cell.is_empty or not isinstance(cell.value, str):
                    continue
                if not is_total_label(cell.value):
                    continue
                label = normalize_text(normalize_digits(cell.value))
                value, value_col, mapping = self._resolve_total_value(grid, row, col, table)
                if value is None:
                    continue
                measure, confidence = self._classify_measure(label, table, value_col)
                if measure == "unknown":
                    continue
                totals.append(
                    DeclaredTotal(
                        measure=measure,
                        label=label,
                        value=value,
                        provenance=_provenance(grid, row, value_col),
                        mapping=mapping,
                        confidence=confidence,
                    )
                )
        return totals

    @staticmethod
    def _resolve_total_value(
        grid: SheetGrid,
        row: int,
        label_col: int,
        table: TableCandidate,
    ) -> tuple[Decimal | None, int, str]:
        # 1) The value written immediately to the right of the label.
        for col in range(label_col + 1, table.last_col + 2):
            cell = grid.get(row, col)
            if cell is None or cell.is_empty:
                continue
            if isinstance(cell.value, str) and not cell.value.strip():
                continue
            value, _ = _as_decimal(cell.value)
            if value is not None:
                return value, col, "label-next-cell"
            break
        # 2) The value in the same column, aligned under a numeric column header.
        cell = grid.get(row, label_col)
        if cell is not None:
            value, _ = _as_decimal(cell.value)
            if value is not None:
                return value, label_col, "column-aligned"
        return None, label_col, "none"

    @staticmethod
    def _classify_measure(label: str, table: TableCandidate, value_col: int) -> tuple[str, float]:
        text = label.lower()
        if any(keyword in text for keyword in REPAYMENT_KEYWORDS):
            return "contract_repayment", 0.85
        if any(keyword in text for keyword in ("اقساط", "قسط", "installment", "instalment", "payment")):
            return "installments", 0.85
        if any(keyword in text for keyword in ("اصل", "principal")):
            return "principal", 0.85
        if "وام" in text and any(keyword in text for keyword in ("جمع", "کل", "total")):
            return "principal", 0.7
        if any(keyword in text for keyword in ("تعهد", "obligation")):
            return "obligations", 0.8
        if any(keyword in text for keyword in ("تعداد", "count")):
            return "count", 0.6
        for column in table.columns:
            if column.index == value_col:
                if column.role is ColumnRole.PRINCIPAL:
                    return "principal", 0.6
                if column.role is ColumnRole.INSTALLMENT:
                    return "installments", 0.6
                if column.role is ColumnRole.INSTALLMENT_COUNT:
                    return "count", 0.6
                if column.role is ColumnRole.AMOUNT:
                    return "amount", 0.5
        return "unknown", 0.0

    # ------------------------------------------------------------------ #
    def extract_normalized_rows(
        self,
        grid: SheetGrid,
        table: TableCandidate,
        limit: int = 500,
    ) -> list[dict[str, Any]]:
        """Raw + standardized value for every cell of the primary table."""
        rows: list[dict[str, Any]] = []
        for row in range(table.first_data_row, min(table.last_data_row, table.first_data_row + limit - 1) + 1):
            cells: list[dict[str, Any]] = []
            for column in table.columns:
                cell = grid.get(row, column.index)
                raw = None if cell is None else cell.value
                parsed = parse_value(raw)
                cells.append(
                    {
                        "column": column.letter,
                        "header": column.header,
                        "role": column.role.value,
                        "raw": raw if not hasattr(raw, "isoformat") else raw.isoformat(),
                        "value": parsed.value,
                        "kind": parsed.kind.value,
                        "warnings": list(parsed.warnings),
                    }
                )
            rows.append(
                {
                    "sheet": grid.sheet_name,
                    "row": row,
                    "source_row": f"{grid.sheet_name}!{get_column_letter(table.first_col)}{row}",
                    "cells": cells,
                }
            )
        return rows
