"""Build the provenance ledger from values Python already calculated.

The ledger deliberately reads the analysed model instead of recomputing
anything.  Each important figure keeps the workbook, sheet, block, cell or
range, original and normalized labels, original and normalized values, period,
unit, operation, formula and the source cells it was derived from.

Two kinds of record exist:

* a *source* record for a value read from the workbook (a declared total);
* a *derived* record for a statistic or reconciliation, which must point back
  to at least one source cell or range.
"""

from __future__ import annotations

import re
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from openpyxl.utils.cell import range_boundaries

from local_office_ai.quality.models import (
    GateStatus,
    ProvenanceLedger,
    ProvenanceRecord,
    ProvenanceSource,
    status_of,
)

_ADDRESS = re.compile(
    r"(?P<sheet>[^\s!]+)!\s*(?P<cell>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)"
)
_BARE_CELL = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$")

_LOAN_DEFAULTS = {
    "loan_count",
    "principal_sum",
    "installments_sum",
    "obligations_sum",
    "contract_repayment_sum",
    "by_institution",
    "principal_sum_check",
}
_STAT_OPERATIONS = {
    "total": "جمع",
    "mean": "میانگین",
    "median": "میانه",
    "minimum": "کمینه",
    "maximum": "بیشینه",
}


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, bool):
        return "بله" if value else "خیر"
    return str(value)


def _is_number(value: Any) -> bool:
    return isinstance(value, (Decimal, int, float)) and not isinstance(value, bool)


def _looks_numeric(value: Any) -> bool:
    text = _text(value).strip()
    if not text:
        return False
    return bool(re.search(r"[0-9\u0660-\u0669\u06F0-\u06F9]", text))


def _row_from_address(address: str) -> int | None:
    match = re.search(r"(\d+)", address or "")
    return int(match.group(1)) if match else None


def _filename(analysis: Any, source_path: str | Path | None) -> str:
    if source_path:
        return Path(source_path).name
    meta = getattr(analysis, "meta", None)
    path = getattr(meta, "path", "") if meta is not None else ""
    return Path(path).name if path else ""


def _source_path(analysis: Any, source_path: str | Path | None) -> str:
    if source_path:
        return str(source_path)
    meta = getattr(analysis, "meta", None)
    return str(getattr(meta, "path", "") or "")


def _block_for(analysis: Any, sheet: str, address: str = "") -> str:
    extraction = getattr(analysis, "extraction", None)
    if extraction is None:
        return ""
    sheet_model = extraction.sheet(sheet) if hasattr(extraction, "sheet") else None
    if sheet_model is None:
        return ""
    row = _row_from_address(address)
    for block in getattr(sheet_model, "blocks", []):
        boundaries = getattr(block, "boundaries", {}) or {}
        first = boundaries.get("header_row") or boundaries.get("first_data_row") or 1
        last = boundaries.get("block_end_row") or boundaries.get("last_data_row") or first
        if row is None or first <= row <= last:
            return str(getattr(block, "block_id", ""))
    blocks = list(getattr(sheet_model, "blocks", []))
    return str(getattr(blocks[0], "block_id", "")) if blocks else ""


def _single_period(periods: Any) -> str:
    """A period string only when the source names exactly one period."""

    tokens: list[str] = []
    for item in periods or []:
        if isinstance(item, dict):
            token = item.get("token") or item.get("text") or ""
        else:
            token = str(item or "")
        token = str(token).strip()
        if token:
            tokens.append(token)
    return tokens[0] if len(tokens) == 1 else ""


def _column_maps(analysis: Any) -> dict[tuple[str, str], dict[str, str]]:
    """Header, unit and period per sheet/column from extraction + semantics."""

    maps: dict[tuple[str, str], dict[str, str]] = {}
    extraction = getattr(analysis, "extraction", None)
    if extraction is not None:
        for sheet in getattr(extraction, "sheets", []):
            sheet_period = _single_period(getattr(sheet, "periods", []))
            for block in getattr(sheet, "blocks", []):
                block_period = _single_period(getattr(block, "periods", []))
                for column in getattr(block, "columns", []):
                    maps[(sheet.name, column.letter)] = {
                        "header": column.header,
                        "original_header": column.original_header or column.header,
                        "period": column.period or block_period or sheet_period,
                        "block": block.block_id,
                        "unit": "",
                    }
    semantics = getattr(analysis, "semantics", None)
    if semantics is not None:
        for block in getattr(semantics, "blocks", []):
            block_period = _single_period(getattr(block, "periods", []))
            for field in getattr(block, "columns", []):
                key = (block.sheet, field.letter)
                item = maps.setdefault(
                    key,
                    {
                        "header": field.header,
                        "original_header": field.header,
                        "period": field.period or block_period,
                        "block": block.block_id,
                        "unit": "",
                    },
                )
                item["unit"] = field.unit or item.get("unit", "")
                item["period"] = (
                    field.period or item.get("period", "") or block_period
                )
                item["block"] = block.block_id or item.get("block", "")
    return maps


def _primary_metrics(analysis: Any) -> tuple[str, str, dict[str, Any]]:
    primary = getattr(analysis, "primary_table", None)
    if primary is None:
        return "", "", {}
    metrics = {
        "address": getattr(primary, "address", ""),
        "title": getattr(primary, "title", ""),
        "first_data_row": getattr(primary, "first_data_row", None),
        "last_data_row": getattr(primary, "last_data_row", None),
        "kind": getattr(getattr(primary, "kind", ""), "value", getattr(primary, "kind", "")),
        "sheet": getattr(primary, "sheet", ""),
    }
    return metrics["sheet"], metrics["address"], metrics


class _CellIndex:
    """Look up original values and formulas in the extraction model."""

    def __init__(self, analysis: Any) -> None:
        self.values: dict[tuple[str, str], str] = {}
        self.formulas: dict[tuple[str, str], str] = {}
        extraction = getattr(analysis, "extraction", None)
        if extraction is None:
            return
        for sheet in getattr(extraction, "sheets", []) or []:
            for block in getattr(sheet, "blocks", []) or []:
                for cell in getattr(block, "cells", []) or []:
                    key = (getattr(sheet, "name", ""), getattr(cell, "address", ""))
                    self.values[key] = _text(getattr(cell, "value", ""))
                    self.formulas[key] = _text(getattr(cell, "formula", ""))

    @staticmethod
    def _range_cells(address: str) -> list[str]:
        if ":" not in address:
            return [address]
        start, end = address.split(":", 1)
        start_match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", start)
        end_match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", end)
        if not start_match or not end_match:
            return [address]
        column = start_match.group(1).upper()
        first = int(start_match.group(2))
        last = int(end_match.group(2))
        if last < first:
            first, last = last, first
        return [f"{column}{row}" for row in range(first, min(last, first + 200) + 1)]

    def value(self, sheet: str, address: str) -> str:
        cells = self._range_cells(address)
        if len(cells) == 1:
            return self.values.get((sheet, cells[0]), "")
        found = [
            self.values[(sheet, cell)]
            for cell in cells
            if self.values.get((sheet, cell), "") != ""
        ]
        return "، ".join(found[:20])

    def formula(self, sheet: str, address: str) -> str:
        cells = self._range_cells(address)
        if len(cells) == 1:
            return self.formulas.get((sheet, cells[0]), "")
        found = [self.formulas.get((sheet, cell), "") for cell in cells]
        return "؛ ".join(item for item in found if item)[:500]



def _column_range(primary_metrics: dict[str, Any], letter: str) -> str:
    first = primary_metrics.get("first_data_row")
    last = primary_metrics.get("last_data_row")
    if first and last:
        return f"{letter}{first}:{letter}{last}"
    return str(primary_metrics.get("address") or "")


def _calculation_sources(
    analysis: Any,
    calculation: Any,
    filename: str,
    columns: dict[tuple[str, str], dict[str, str]],
    primary_metrics: dict[str, Any],
    cells: _CellIndex,
) -> tuple[list[ProvenanceSource], bool]:
    sources: list[ProvenanceSource] = []
    for item in getattr(calculation, "sources", []) or []:
        sheet = getattr(item, "sheet", "") or ""
        address = getattr(item, "address", "") or ""
        key = (sheet, re.match(r"[A-Za-z]+", address).group(0) if re.match(r"[A-Za-z]+", address) else "")
        info = columns.get(key, {})
        sources.append(
            ProvenanceSource(
                filename=filename,
                sheet=sheet,
                block=_block_for(analysis, sheet, address),
                cell=address,
                label=info.get("header", ""),
                original_value=_text(getattr(item, "raw", "")) or cells.value(sheet, address),
                normalized_value=_text(getattr(item, "raw", "")) or cells.value(sheet, address),
                period=info.get("period", ""),
                unit=info.get("unit", ""),
                formula=getattr(item, "formula", "") or cells.formula(sheet, address),
            )
        )
    if sources:
        return sources, True

    name = str(getattr(calculation, "name", "") or "")
    primary = getattr(analysis, "primary_table", None)
    if primary is None:
        return [], False

    primary_kind = str(primary_metrics.get("kind") or "")
    if name in _LOAN_DEFAULTS and primary_kind != "loan":
        # Loan domain defaults are not figures of a non-loan workbook.
        return [], False
    if name.startswith("loan_") and primary_kind != "loan":
        return [], False

    sheet = str(primary_metrics.get("sheet") or "")
    parts = name.split("::")
    letter = ""
    if len(parts) >= 2 and re.fullmatch(r"[A-Za-z]{1,3}", parts[1] or ""):
        letter = parts[1]
    address = _column_range(primary_metrics, letter) if letter else str(primary_metrics.get("address") or "")
    info = columns.get((sheet, letter), {})
    sources.append(
        ProvenanceSource(
            filename=filename,
            sheet=sheet,
            block=_block_for(analysis, sheet, address),
            cell=address,
            label=info.get("header", "") or str(getattr(primary, "title", "") or ""),
            original_value=cells.value(sheet, address),
            normalized_value=cells.value(sheet, address),
            period=info.get("period", ""),
            unit=info.get("unit", ""),
            formula=cells.formula(sheet, address),
            note="محدوده منبع از جدول اصلی استخراج شد.",
        )
    )
    return sources, True


def _add_calculation_records(
    ledger: ProvenanceLedger,
    analysis: Any,
    filename: str,
    columns: dict[tuple[str, str], dict[str, str]],
    primary_metrics: dict[str, Any],
    cells: _CellIndex,
) -> None:
    for index, calculation in enumerate(getattr(analysis, "calculations", []) or [], start=1):
        value = getattr(calculation, "value", None)
        sources, important = _calculation_sources(
            analysis, calculation, filename, columns, primary_metrics, cells
        )
        if not sources and not important:
            continue
        if not _is_number(value) and not sources:
            continue
        key = f"calculation.{index}.{getattr(calculation, 'name', '')}"
        first = sources[0] if sources else None
        ledger.add(
            ProvenanceRecord(
                key=key,
                filename=filename,
                sheet=first.sheet if first else str(primary_metrics.get("sheet") or ""),
                block=first.block if first else "",
                cell=first.cell if first else "",
                original_label=str(getattr(calculation, "label", "") or ""),
                normalized_label=str(getattr(calculation, "label", "") or ""),
                original_value=_text(getattr(calculation, "verification_value", ""))
                or (first.original_value if first else ""),
                normalized_value=_text(value),
                period=first.period if first else "",
                unit=str(getattr(calculation, "unit", "") or (first.unit if first else "")),
                operation=str(getattr(calculation, "method", "") or ""),
                formula=(first.formula if first else "")
                or str(getattr(calculation, "method", "") or ""),
                validation_status=(
                    GateStatus.PASS.value
                    if getattr(calculation, "verified", None) is not False
                    else GateStatus.FAIL.value
                ),
                sources=sources,
                derived=True,
                important=important,
                notes=list(getattr(calculation, "inputs", []) or [])[:8],
            )
        )


def _add_declared_total_records(ledger: ProvenanceLedger, analysis: Any, filename: str) -> None:
    for index, total in enumerate(getattr(analysis, "declared_totals", []) or [], start=1):
        value = getattr(total, "value", None)
        if not _is_number(value):
            # A formula total without a cached value cannot be verified by
            # Python; the reconciliation layer reports it as a warning.
            continue
        provenance = getattr(total, "provenance", None)
        sheet = getattr(provenance, "sheet", "") if provenance else ""
        address = getattr(provenance, "address", "") if provenance else ""
        if not sheet or not address:
            continue
        ledger.add(
            ProvenanceRecord(
                key=f"declared_total.{index}.{sheet}.{address}",
                filename=filename,
                sheet=sheet,
                block=_block_for(analysis, sheet, address),
                cell=address,
                original_label=str(getattr(total, "label", "") or ""),
                normalized_label=str(getattr(total, "measure", "") or ""),
                original_value=_text(getattr(provenance, "raw", "")),
                normalized_value=_text(value),
                operation="خواندن جمع اعلامی فایل",
                formula=getattr(provenance, "formula", "") or "",
                validation_status=GateStatus.PASS.value,
                derived=False,
                important=True,
            )
        )


def _add_generic_stat_records(
    ledger: ProvenanceLedger,
    analysis: Any,
    filename: str,
    columns: dict[tuple[str, str], dict[str, str]],
    primary_metrics: dict[str, Any],
    cells: _CellIndex,
) -> None:
    generic = getattr(analysis, "generic", None)
    if generic is None:
        return
    sheet = str(getattr(generic, "table_sheet", "") or primary_metrics.get("sheet") or "")
    for column in getattr(generic, "columns", []) or []:
        if not getattr(column, "is_numeric", False):
            continue
        if getattr(column, "is_identifier", False) or getattr(column, "is_temporal_key", False):
            continue
        letter = str(getattr(column, "letter", "") or "")
        if not letter:
            continue
        info = columns.get((sheet, letter), {})
        address = _column_range(primary_metrics, letter)
        for field_name, operation in _STAT_OPERATIONS.items():
            value = getattr(column, field_name, None)
            if value is None:
                continue
            header = str(getattr(column, "header", "") or letter)
            original_header = str(info.get("original_header") or header)
            ledger.add(
                ProvenanceRecord(
                    key=f"stat.{sheet}.{letter}.{field_name}",
                    filename=filename,
                    sheet=sheet,
                    block=info.get("block") or _block_for(analysis, sheet, address),
                    cell=address,
                    original_label=original_header,
                    normalized_label=f"{header} - {operation}",
                    original_value=cells.value(sheet, address),
                    normalized_value=_text(value),
                    period=info.get("period", ""),
                    unit=info.get("unit", ""),
                    operation=operation,
                    formula=f"{operation}({sheet}!{address})",
                    validation_status=GateStatus.PASS.value,
                    sources=[
                        ProvenanceSource(
                            filename=filename,
                            sheet=sheet,
                            block=info.get("block") or _block_for(analysis, sheet, address),
                            cell=address,
                            label=header,
                            original_value=cells.value(sheet, address),
                            normalized_value=cells.value(sheet, address),
                            period=info.get("period", ""),
                            unit=info.get("unit", ""),
                            formula=cells.formula(sheet, address),
                            note="محدوده ستون در جدول اصلی",
                        )
                    ],
                    derived=True,
                    important=True,
                )
            )


def _parse_sources(
    raw_sources: Iterable[str],
    *,
    filename: str,
    analysis: Any,
    cells: _CellIndex,
    fallback_sheet: str = "",
) -> list[ProvenanceSource]:
    sources: list[ProvenanceSource] = []
    for raw in raw_sources:
        text = str(raw or "").strip()
        if not text:
            continue
        matches = list(_ADDRESS.finditer(text))
        if matches:
            for match in matches:
                sheet = match.group("sheet")
                cell = match.group("cell")
                sources.append(
                    ProvenanceSource(
                        filename=filename,
                        sheet=sheet,
                        block=_block_for(analysis, sheet, cell),
                        cell=cell,
                        original_value=cells.value(sheet, cell),
                        normalized_value=cells.value(sheet, cell),
                        formula=cells.formula(sheet, cell),
                    )
                )
        elif _BARE_CELL.match(text) and fallback_sheet:
            sources.append(
                ProvenanceSource(
                    filename=filename,
                    sheet=fallback_sheet,
                    block=_block_for(analysis, fallback_sheet, text),
                    cell=text,
                    original_value=cells.value(fallback_sheet, text),
                    normalized_value=cells.value(fallback_sheet, text),
                    formula=cells.formula(fallback_sheet, text),
                )
            )
    # Keep each cell once; a range already covers its own cells.
    seen: set[tuple[str, str]] = set()
    unique: list[ProvenanceSource] = []
    for item in sources:
        marker = (item.sheet, item.cell)
        if marker in seen:
            continue
        seen.add(marker)
        unique.append(item)
    return unique


def _add_reconciliation_records(
    ledger: ProvenanceLedger, analysis: Any, filename: str, cells: _CellIndex
) -> None:
    report = getattr(analysis, "statistical_reconciliation", None)
    if report is None:
        return
    items = (
        list(getattr(report, "reconciliations", []) or [])
        + list(getattr(report, "comparisons", []) or [])
        + list(getattr(report, "averages", []) or [])
    )
    for index, item in enumerate(items, start=1):
        calculated = getattr(item, "calculated", None)
        expected = getattr(item, "expected", None)
        if calculated is None:
            # A relationship that produced no numeric value is reported by the
            # reconciliation engine itself; it is not a number to trace.
            continue
        sheet = str(getattr(item, "sheet", "") or "")
        raw_sources = list(getattr(item, "sources", []) or [])
        expected_source = str(getattr(item, "expected_source", "") or "")
        if expected_source:
            raw_sources.append(expected_source)
        sources = _parse_sources(
            raw_sources, filename=filename, analysis=analysis, cells=cells, fallback_sheet=sheet
        )
        status = status_of(getattr(item, "status", "")) or GateStatus.WARNING
        label = str(
            getattr(item, "label", "")
            or getattr(item, "kind_label", "")
            or getattr(item, "kind", "")
        )
        ledger.add(
            ProvenanceRecord(
                key=f"reconciliation.{index}.{getattr(item, 'kind', '')}.{label}",
                filename=filename,
                sheet=sheet,
                block=str(getattr(item, "block", "") or ""),
                cell=sources[0].cell if sources else "",
                original_label=label,
                normalized_label=str(getattr(item, "measure", "") or label),
                original_value=_text(expected)
                or "، ".join(
                    source.original_value
                    for source in sources
                    if source.original_value
                )[:500],
                normalized_value=_text(getattr(item, "calculated", None)),
                unit=str(getattr(item, "unit", "") or ""),
                operation=str(getattr(item, "kind_label", "") or getattr(item, "kind", "")),
                formula=str(getattr(item, "kind", "") or ""),
                validation_status=status.value,
                sources=sources,
                derived=True,
                important=True,
                notes=[str(getattr(item, "explanation", "") or "")][:1],
            )
        )


def _add_period_records(
    ledger: ProvenanceLedger,
    analysis: Any,
    filename: str,
    columns: dict[tuple[str, str], dict[str, str]],
    cells: _CellIndex,
) -> None:
    integration = getattr(analysis, "integration", None)
    if integration is None:
        return
    primary_sheet, _, _ = _primary_metrics(analysis)
    for series in getattr(integration, "series", []) or []:
        metric = str(getattr(series, "metric", "") or "")
        for cell in getattr(series, "cells", []) or []:
            total = getattr(cell, "total", None)
            if total is None:
                continue
            sheet = str(getattr(cell, "source", "") or primary_sheet)
            address = str(getattr(cell, "total_address", "") or "")
            period = f"{getattr(cell, 'month', '')} {getattr(cell, 'period', '')}".strip()
            declared = getattr(cell, "declared_total", None)
            verification = status_of(getattr(cell, "verification", "")) or GateStatus.PASS
            info = columns.get((sheet, ""), {})
            ledger.add(
                ProvenanceRecord(
                    key=f"period.{metric}.{getattr(cell, 'period', '')}.{getattr(cell, 'month', '')}",
                    filename=filename,
                    sheet=sheet,
                    block=_block_for(analysis, sheet, address),
                    cell=address,
                    original_label=metric,
                    normalized_label=f"{metric} - {period}",
                    original_value=_text(declared) or cells.value(sheet, address),
                    normalized_value=_text(total),
                    period=period,
                    unit=info.get("unit", ""),
                    operation="جمع ماهانه",
                    formula=f"SUM({sheet}!{address})" if address else "",
                    validation_status=verification.value,
                    sources=[
                        ProvenanceSource(
                            filename=filename,
                            sheet=sheet,
                            block=_block_for(analysis, sheet, address),
                            cell=address,
                            label=metric,
                            original_value=cells.value(sheet, address),
                            normalized_value=_text(total),
                            period=period,
                            formula=cells.formula(sheet, address),
                            note=str(getattr(cell, "note", "") or ""),
                        )
                    ],
                    derived=True,
                    important=True,
                )
            )


def _add_answer_records(
    ledger: ProvenanceLedger,
    analysis: Any,
    filename: str,
    primary_metrics: dict[str, Any],
    cells: _CellIndex,
) -> None:
    primary_table_address = str(primary_metrics.get("address") or "")
    primary_sheet = str(primary_metrics.get("sheet") or "")
    for index, answer in enumerate(getattr(analysis, "request_answers", []) or [], start=1):
        value = getattr(answer, "value", "")
        if str(value or "").strip() == "":
            continue
        if not _looks_numeric(value):
            # Text-only answers (for example "no monthly period was detected")
            # are not numerical results and do not need cell provenance.
            continue
        raw_source = str(getattr(answer, "source", "") or "")
        sources = _parse_sources(
            [raw_source], filename=filename, analysis=analysis, cells=cells, fallback_sheet=primary_sheet
        )
        if not sources and primary_table_address:
            sources = [
                ProvenanceSource(
                    filename=filename,
                    sheet=primary_sheet,
                    block=_block_for(analysis, primary_sheet, primary_table_address),
                    cell=primary_table_address,
                    original_value=cells.value(primary_sheet, primary_table_address),
                    normalized_value=cells.value(primary_sheet, primary_table_address),
                    note="محدوده جدول اصلی به عنوان منبع پاسخ ثبت شد.",
                )
            ]
        title = str(getattr(answer, "title", "") or "")
        ledger.add(
            ProvenanceRecord(
                key=f"answer.{index}",
                filename=filename,
                sheet=sources[0].sheet if sources else primary_sheet,
                block=sources[0].block if sources else "",
                cell=sources[0].cell if sources else "",
                original_label=title,
                normalized_label=title,
                original_value=raw_source,
                normalized_value=_text(value),
                operation="پاسخ درخواست",
                formula=raw_source,
                validation_status=(
                    GateStatus.PASS.value
                    if getattr(answer, "verified", None) is not False
                    else GateStatus.WARNING.value
                ),
                sources=sources,
                derived=True,
                important=True,
            )
        )


def _address_overlaps(first: str, second: str) -> bool:
    """Whether two A1 addresses/ranges share at least one cell."""

    try:
        first_min_col, first_min_row, first_max_col, first_max_row = range_boundaries(first)
        second_min_col, second_min_row, second_max_col, second_max_row = range_boundaries(second)
    except (ValueError, TypeError):
        return False
    return (
        first_min_col <= second_max_col
        and second_min_col <= first_max_col
        and first_min_row <= second_max_row
        and second_min_row <= first_max_row
    )


def _link_related_records(ledger: ProvenanceLedger) -> None:
    """Attach source-record keys and source addresses to derived records."""

    source_records = [item for item in ledger.records if not item.derived]
    for record in ledger.records:
        if not record.derived:
            continue
        links: set[str] = {
            f"{source.sheet}!{source.cell}"
            for source in record.sources
            if source.sheet and source.cell
        }
        for source in record.sources:
            for candidate in source_records:
                if (
                    candidate.filename == source.filename
                    and candidate.sheet == source.sheet
                    and _address_overlaps(candidate.cell, source.cell)
                ):
                    links.add(candidate.key)
        record.related_records = sorted(links)


def build_provenance_ledger(
    analysis: Any,
    *,
    source_path: str | Path | None = None,
) -> ProvenanceLedger:
    """Collect provenance for every important result the analysis produced."""

    filename = _filename(analysis, source_path)
    ledger = ProvenanceLedger(
        filename=filename,
        source_path=_source_path(analysis, source_path),
    )
    columns = _column_maps(analysis)
    _, _, primary_metrics = _primary_metrics(analysis)
    cells = _CellIndex(analysis)

    _add_declared_total_records(ledger, analysis, filename)
    _add_calculation_records(ledger, analysis, filename, columns, primary_metrics, cells)
    _add_generic_stat_records(ledger, analysis, filename, columns, primary_metrics, cells)
    _add_reconciliation_records(ledger, analysis, filename, cells)
    _add_period_records(ledger, analysis, filename, columns, cells)
    _add_answer_records(ledger, analysis, filename, primary_metrics, cells)
    _link_related_records(ledger)
    return ledger
