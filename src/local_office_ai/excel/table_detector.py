"""Table detection: find one or more meaningful tables per sheet.

The detector does not assume a fixed layout.  It walks each content region,
scores header rows, follows the data until a total row, a repeated header or an
empty band, profiles the columns and classifies the table by its roles.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import ColumnProfile, ColumnRole, TableCandidate, TableKind, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    looks_like_date,
    looks_like_number,
    normalize_text,
    parse_value,
)
from local_office_ai.excel.sheet_analyzer import (
    HeaderCandidate,
    SheetAnalyzer,
    SheetGrid,
    is_total_label,
    summary_label_role,
)

LOGGER = get_logger(__name__)

REPORT_SHEET_HINTS = (
    "داشبورد", "dashboard", "گزارش", "report", "خلاصه", "summary", "تحلیل", "analysis",
    "پایش", "monitoring", "performance", "gpu", "trace", "log",
    "جزئیات پردازش", "نمونه‌های", "آداپتور", "adapter",
)


@dataclass
class DetectedStructure:
    tables: list[TableCandidate] = field(default_factory=list)
    key_value_tables: list[TableCandidate] = field(default_factory=list)

    @property
    def all_tables(self) -> list[TableCandidate]:
        return list(self.tables) + list(self.key_value_tables)


def is_report_sheet(name: str) -> bool:
    lowered = normalize_text(name).lower()
    return any(hint in lowered for hint in REPORT_SHEET_HINTS)


class TableDetector:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def detect(self, grid: SheetGrid) -> DetectedStructure:
        analyzer = SheetAnalyzer(grid, self.settings)
        structure = DetectedStructure()
        for region in analyzer.regions():
            position = region.start
            while position <= region.end:
                header = analyzer.find_header(position, region.end)
                if header is None:
                    break
                candidates = [header]
                for first_col, last_col in analyzer.column_groups(header.row, region_end=region.end):
                    if (first_col, last_col) == (header.first_col, header.last_col):
                        continue
                    candidate = analyzer.score_header(header.row, first_col, last_col, region.end)
                    if candidate is not None and candidate.score >= 0.35:
                        candidates.append(candidate)
                last_row = header.row
                for candidate in candidates:
                    if candidate.band is None:
                        candidate.band = analyzer.detect_header_band(candidate, region.end)
                    table, data_end = self._build_table(analyzer, grid, candidate, region.end)
                    last_row = max(last_row, data_end)
                    if table is not None:
                        has_numbers = any(column.numeric_ratio >= 0.3 for column in table.columns)
                        if has_numbers or table.kind in {TableKind.LOAN, TableKind.OBLIGATION}:
                            structure.tables.append(table)
                position = max(last_row, header.row) + 1

            structure.tables = self._dedupe_tables(structure.tables)
            for key_value in self._detect_key_value(analyzer, grid, region.start, region.end, structure.tables):
                structure.key_value_tables.append(key_value)
        LOGGER.info(
            "sheet %r: %d tables, %d key-value lists",
            grid.sheet_name, len(structure.tables), len(structure.key_value_tables),
        )
        return structure

    @staticmethod
    def _dedupe_tables(tables: list[TableCandidate]) -> list[TableCandidate]:
        """Drop header-span variants fully contained in a wider table with the same rows."""
        kept: list[TableCandidate] = []
        for table in tables:
            duplicate = any(
                other is not table
                and other.sheet == table.sheet
                and other.header_row == table.header_row
                and other.first_data_row == table.first_data_row
                and other.last_data_row == table.last_data_row
                and other.block_end == table.block_end
                and other.first_col <= table.first_col
                and other.last_col >= table.last_col
                and (other.first_col, other.last_col) != (table.first_col, table.last_col)
                for other in tables
            )
            if duplicate:
                continue
            if any(
                other is not table
                and other.sheet == table.sheet
                and other.header_row == table.header_row
                and other.first_data_row == table.first_data_row
                and other.last_data_row == table.last_data_row
                and (other.first_col, other.last_col) == (table.first_col, table.last_col)
                for other in kept
            ):
                continue
            kept.append(table)
        return kept

    # ------------------------------------------------------------------ #
    def _build_table(
        self,
        analyzer: SheetAnalyzer,
        grid: SheetGrid,
        header: HeaderCandidate,
        region_end: int,
    ) -> tuple[TableCandidate | None, int]:
        first_col, last_col = header.first_col, header.last_col
        width = last_col - first_col + 1
        band = header.band or analyzer.detect_header_band(header, region_end)
        data_start = band.bottom_row + 1
        last_data = data_start - 1
        blank_streak = 0
        subtotal_rows: list[int] = []
        grand_total_rows: list[int] = []
        terminal_total = 0
        block_end = 0
        for row in range(data_start, region_end + 1):
            filled = grid.row_non_empty_count(row, first_col, last_col)
            if filled == 0:
                blank_streak += 1
                if blank_streak >= 2:
                    break
                continue
            if self._row_is_total(grid, row, first_col, last_col):
                if self._row_has_data_after(
                    analyzer, grid, row, first_col, last_col, region_end
                ):
                    # A subtotal is part of the table, not its end.
                    subtotal_rows.append(row)
                    last_data = row
                    blank_streak = 0
                    continue
                terminal_total = row
                block_end = self._extend_block_end(
                    analyzer, grid, row, first_col, last_col, region_end
                )
                break
            if self._row_is_repeated_header(analyzer, grid, row, first_col, last_col, region_end):
                break
            if (
                last_data >= data_start
                and self._row_is_lone_title(grid, row, first_col, last_col)
                and self._row_starts_new_block(
                    analyzer, grid, row, first_col, last_col, region_end
                )
            ):
                break
            fill_ratio = filled / width
            if fill_ratio < 0.3 and last_data >= data_start:
                break
            if width >= 3 and last_data >= data_start and self._row_is_label_value(grid, row, first_col, last_col):
                label = self._first_text(grid, row, first_col, last_col)
                if label is not None and label.column != first_col:
                    break
            last_data = row
            blank_streak = 0

        if last_data < data_start:
            return None, header.row
        if not block_end:
            block_end = last_data
        if terminal_total:
            grand_total_rows = [
                candidate
                for candidate in range(terminal_total, block_end + 1)
                if self._row_is_total(grid, candidate, first_col, last_col)
            ]
        row_count = last_data - data_start + 1
        if row_count < self.settings.min_table_rows:
            return None, max(last_data, block_end)

        profiles = analyzer.profile_columns_from_band(band, data_start, last_data)
        role_warnings = analyzer.assign_roles(profiles)
        role_warnings.extend(analyzer.assign_statistical_roles(profiles, data_start, last_data))

        if not any(profile.role is ColumnRole.INSTITUTION for profile in profiles) and any(
            profile.role in {ColumnRole.PRINCIPAL, ColumnRole.INSTALLMENT} for profile in profiles
        ):
            fallback = analyzer.fallback_institution(profiles)
            if fallback is not None:
                fallback.role = ColumnRole.INSTITUTION
                fallback.confidence = max(fallback.confidence, 0.55)
                role_warnings.append(
                    f"ستون «{fallback.header or fallback.letter}» با تکیه بر محتوای متنی "
                    "به‌عنوان نام مؤسسه انتخاب شد."
                )

        kind, confidence, reasons = self._classify(profiles, row_count, grid, data_start, last_data)
        title_cell = self._find_title(grid, header.row)
        table = TableCandidate(
            sheet=grid.sheet_name,
            header_row=header.row,
            first_data_row=data_start,
            last_data_row=last_data,
            first_col=first_col,
            last_col=last_col,
            columns=profiles,
            kind=kind,
            confidence=confidence,
            reasons=reasons + header.reasons + role_warnings,
            header_values=[profile.header for profile in profiles],
            title=title_cell,
            header_rows=band.rows,
            header_group_count=len([group for group in band.groups if group["count"] > 1]),
            data_rows_with_values=sum(
                1
                for row in range(data_start, last_data + 1)
                if row not in subtotal_rows
                and any(
                    (cell := grid.get(row, column.index)) is not None
                    and not cell.is_empty
                    and parse_value(cell.value).kind is ValueKind.NUMBER
                    for column in profiles
                )
            ),
            notes=band.reasons,
            block_end_row=block_end,
            subtotal_rows=sorted(subtotal_rows),
            grand_total_rows=sorted(grand_total_rows),
        )
        return table, block_end

    def _find_title(self, grid: SheetGrid, header_row: int) -> str:
        for row in range(header_row - 1, max(header_row - 4, 0), -1):
            cells = [cell for cell in grid.row_cells(row) if not cell.is_empty]
            if len(cells) == 1 and isinstance(cells[0].value, str) and len(cells[0].value.strip()) > 3:
                return cells[0].value.strip()[:120]
        return ""

    @staticmethod
    def _row_is_total(grid: SheetGrid, row: int, first_col: int, last_col: int) -> bool:
        """A totals row leads with a totals label and carries a number.

        Whether it ends the block is decided separately by
        :meth:`_row_has_data_after`; a totals row with data below it is a
        subtotal and stays inside the table.
        """
        first_text = TableDetector._first_text(grid, row, first_col, last_col)
        if first_text is None or not is_total_label(first_text.value):
            return False
        for col in range(first_col, last_col + 1):
            cell = grid.get(row, col)
            if cell is None:
                continue
            if cell.formula:
                return True
            if cell.is_empty:
                continue
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                return True
            if isinstance(cell.value, str) and looks_like_number(cell.value):
                return True
        return False

    @staticmethod
    def _row_has_numeric(grid: SheetGrid, row: int, first_col: int, last_col: int) -> bool:
        for col in range(first_col, last_col + 1):
            cell = grid.get(row, col)
            if cell is None:
                continue
            if cell.formula:
                return True
            if cell.is_empty:
                continue
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                return True
            if isinstance(cell.value, str) and looks_like_number(cell.value):
                return True
        return False

    @staticmethod
    def _row_is_text_only(grid: SheetGrid, row: int, first_col: int, last_col: int) -> bool:
        cells = [
            cell
            for col in range(first_col, last_col + 1)
            if (cell := grid.get(row, col)) is not None and (not cell.is_empty or cell.formula)
        ]
        if not cells:
            return False
        for cell in cells:
            if not isinstance(cell.value, str):
                return False
            if looks_like_number(cell.value) or looks_like_date(cell.value):
                return False
        return True

    def _row_has_data_after(
        self,
        analyzer: SheetAnalyzer,
        grid: SheetGrid,
        row: int,
        first_col: int,
        last_col: int,
        region_end: int,
    ) -> bool:
        """True when real data rows follow a totals row inside the same block."""
        blank_streak = 0
        for candidate in range(row + 1, region_end + 1):
            filled = grid.row_non_empty_count(candidate, first_col, last_col)
            if filled == 0:
                blank_streak += 1
                if blank_streak >= 2:
                    return False
                continue
            blank_streak = 0
            if self._row_is_repeated_header(
                analyzer, grid, candidate, first_col, last_col, region_end
            ):
                return False
            if self._row_is_lone_title(
                grid, candidate, first_col, last_col
            ) and self._row_starts_new_block(
                analyzer, grid, candidate, first_col, last_col, region_end
            ):
                return False
            if self._row_is_total(grid, candidate, first_col, last_col):
                if any(
                    self._row_is_total(grid, later, first_col, last_col)
                    for later in range(candidate + 1, region_end + 1)
                    if grid.row_non_empty_count(later, first_col, last_col) > 0
                ):
                    return True
                continue
            label = self._first_text(grid, candidate, first_col, last_col)
            if label is not None and summary_label_role(label.value) in {
                "total", "subtotal", "difference", "percentage", "note",
            }:
                continue
            if self._row_is_text_only(grid, candidate, first_col, last_col):
                continue
            if self._row_has_numeric(grid, candidate, first_col, last_col):
                return True
            if filled >= 2:
                return True
        return False

    def _extend_block_end(
        self,
        analyzer: SheetAnalyzer,
        grid: SheetGrid,
        row: int,
        first_col: int,
        last_col: int,
        region_end: int,
    ) -> int:
        """Extend a terminal totals row over its comparison/percentage/note rows."""
        end = row
        for candidate in range(row + 1, region_end + 1):
            if grid.row_non_empty_count(candidate, first_col, last_col) == 0:
                break
            if self._row_is_repeated_header(
                analyzer, grid, candidate, first_col, last_col, region_end
            ):
                break
            label = self._first_text(grid, candidate, first_col, last_col)
            role = summary_label_role(label.value) if label is not None else None
            if role == "note":
                end = candidate
                continue
            if self._row_is_lone_title(
                grid, candidate, first_col, last_col
            ) and self._row_starts_new_block(
                analyzer, grid, candidate, first_col, last_col, region_end
            ):
                break
            if role or self._row_is_text_only(grid, candidate, first_col, last_col):
                end = candidate
                continue
            break
        return end

    @staticmethod
    def _first_text(grid: SheetGrid, row: int, first_col: int, last_col: int):
        for col in range(first_col, last_col + 1):
            cell = grid.get(row, col)
            if cell is None or cell.is_empty:
                continue
            if isinstance(cell.value, str) and not looks_like_number(cell.value):
                return cell
        return None

    def _row_is_repeated_header(
        self,
        analyzer: SheetAnalyzer,
        grid: SheetGrid,
        row: int,
        first_col: int,
        last_col: int,
        region_end: int,
    ) -> bool:
        candidate = analyzer.score_header(row, first_col, last_col, region_end)
        if candidate is None or candidate.score < 0.75:
            return False
        numeric_like = 0
        for col in range(first_col, last_col + 1):
            cell = grid.get(row, col)
            if cell is None or cell.is_empty:
                continue
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                numeric_like += 1
        return numeric_like == 0

    def _row_starts_new_block(
        self,
        analyzer: SheetAnalyzer,
        grid: SheetGrid,
        row: int,
        first_col: int,
        last_col: int,
        region_end: int,
    ) -> bool:
        """True when the next non-empty row is a header or the region ends."""
        for candidate in range(row + 1, min(row + 4, region_end + 1)):
            filled = grid.row_non_empty_count(candidate, first_col, last_col)
            if filled == 0:
                continue
            return self._row_is_repeated_header(
                analyzer, grid, candidate, first_col, last_col, region_end
            )
        return True

    @staticmethod
    def _row_is_lone_title(grid: SheetGrid, row: int, first_col: int, last_col: int) -> bool:
        """A single long text cell after data starts is a section title, not a row."""
        cells = [
            cell
            for col in range(first_col, last_col + 1)
            if (cell := grid.get(row, col)) is not None and not cell.is_empty
        ]
        if len(cells) != 1:
            return False
        value = cells[0].value
        return (
            isinstance(value, str)
            and not looks_like_number(value)
            and len(value.strip()) > 3
        )

    @staticmethod
    def _row_is_label_value(grid: SheetGrid, row: int, first_col: int, last_col: int) -> bool:
        """True for a short row shaped like ``label | amount`` (+ optional unit)."""
        filled = [
            cell
            for col in range(first_col, last_col + 1)
            if (cell := grid.get(row, col)) is not None and not cell.is_empty
        ]
        if not (2 <= len(filled) <= 3):
            return False
        label = TableDetector._first_text(grid, row, first_col, last_col)
        if label is None:
            return False
        neighbour = grid.get(row, label.column + 1)
        return neighbour is not None and not neighbour.is_empty and not isinstance(neighbour.value, str)

    # ------------------------------------------------------------------ #
    def _classify(
        self,
        profiles: list[ColumnProfile],
        row_count: int,
        grid: SheetGrid,
        data_start: int,
        last_data: int,
    ) -> tuple[TableKind, float, list[str]]:
        roles = {profile.role for profile in profiles}
        numeric_cols = [profile for profile in profiles if profile.data_type is ValueKind.NUMBER]
        date_cols = [profile for profile in profiles if profile.data_type is ValueKind.DATE]
        reasons: list[str] = []

        if len(numeric_cols) >= 4 and self._first_column_is_time_series(grid, profiles, data_start, last_data):
            return TableKind.TIME_SERIES, 0.85, ["first column is a numeric time axis", "many numeric columns"]

        has_institution = ColumnRole.INSTITUTION in roles
        has_principal = ColumnRole.PRINCIPAL in roles
        has_installment = ColumnRole.INSTALLMENT in roles
        has_count = ColumnRole.INSTALLMENT_COUNT in roles or ColumnRole.COUNT in roles
        has_amount = ColumnRole.AMOUNT in roles
        has_date = ColumnRole.DATE in roles or bool(date_cols)

        if has_institution and (has_principal or has_installment) and (has_count or has_date):
            confidence = 0.8
            reasons.append("institution column with principal/installment and count/date columns")
            if has_principal and has_installment:
                confidence += 0.08
            if has_count and has_date:
                confidence += 0.05
            if row_count >= 3:
                confidence += 0.02
            return TableKind.LOAN, min(confidence, 0.97), reasons

        if has_institution and (has_amount or has_principal) and not has_installment and not has_count:
            confidence = 0.62
            reasons.append("institution column with an amount column and no installment columns")
            if any(
                keyword in normalize_text(profile.header)
                for profile in profiles
                for keyword in ("تعهد", "obligation")
            ):
                confidence += 0.1
            if row_count >= 2:
                confidence += 0.03
            return TableKind.OBLIGATION, min(confidence, 0.9), reasons

        has_date = ColumnRole.DATE in roles or bool(date_cols)
        has_description = ColumnRole.DESCRIPTION in roles
        has_id = ColumnRole.ID in roles
        if (
            row_count <= 10
            and numeric_cols
            and not has_institution
            and not has_installment
            and not has_date
            and not has_description
            and not has_id
        ):
            return TableKind.KPI, 0.5, ["small label/value table without institution, date or id columns"]

        confidence = 0.3 + 0.08 * len(numeric_cols) + (0.05 if date_cols else 0.0)
        return TableKind.GENERIC, min(confidence, 0.7), reasons or ["no financial role combination matched"]

    @staticmethod
    def _first_column_is_time_series(
        grid: SheetGrid,
        profiles: list[ColumnProfile],
        data_start: int,
        last_data: int,
    ) -> bool:
        if not profiles:
            return False
        first = profiles[0]
        if first.data_type is not ValueKind.NUMBER:
            return False
        values = []
        for row in range(data_start, last_data + 1):
            cell = grid.get(row, first.index)
            if cell is None or cell.is_empty:
                continue
            parsed = looks_like_number(cell.value)
            if not parsed:
                return False
            try:
                values.append(float(cell.value))
            except (TypeError, ValueError):
                return False
        if len(values) < 4:
            return False
        increases = sum(1 for left, right in zip(values, values[1:]) if right >= left)
        return increases / max(len(values) - 1, 1) >= 0.9

    # ------------------------------------------------------------------ #
    def _detect_key_value(
        self,
        analyzer: SheetAnalyzer,
        grid: SheetGrid,
        region_start: int,
        region_end: int,
        tables: list[TableCandidate],
    ) -> list[TableCandidate]:
        """Label-value lists without a header (e.g. commitments under a loan table)."""
        covered = set()
        for table in tables:
            if table.sheet != grid.sheet_name:
                continue
            covered.update(range(table.header_row, table.block_end + 1))

        records: list[tuple[int, int, int]] = []  # (row, label_col, value_col)
        for row in range(region_start, region_end + 1):
            if row in covered:
                continue
            cells = [cell for cell in grid.row_cells(row) if not cell.is_empty]
            if not (2 <= len(cells) <= 3):
                continue
            label_cell = None
            value_cell = None
            for cell in cells:
                if isinstance(cell.value, str) and not looks_like_number(cell.value) and not looks_like_date(cell.value):
                    if label_cell is None:
                        label_cell = cell
                elif isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    if value_cell is None:
                        value_cell = cell
            if label_cell is None or value_cell is None:
                continue
            if value_cell.column != label_cell.column + 1:
                continue
            if is_total_label(label_cell.value):
                continue
            if len(str(label_cell.value).strip()) < 2:
                continue
            records.append((row, label_cell.column, value_cell.column))

        if not records:
            return []

        # Group records that share the same label/value columns and are close together.
        grouped: list[list[tuple[int, int, int]]] = []
        for record in records:
            if grouped and grouped[-1][-1][1:] == record[1:] and record[0] - grouped[-1][-1][0] <= 3:
                grouped[-1].append(record)
            else:
                grouped.append([record])

        results: list[TableCandidate] = []
        for group in grouped:
            label_col = group[0][1]
            value_col = group[0][2]
            first_row = group[0][0]
            last_row = group[-1][0]
            label_profile = ColumnProfile(
                index=label_col,
                letter=_letter(label_col),
                header="شرح / مؤسسه",
                header_address=f"{_letter(label_col)}{first_row}",
                data_type=ValueKind.TEXT,
                role=ColumnRole.INSTITUTION,
                text_ratio=1.0,
                non_empty=len(group),
                distinct=len(group),
                samples=[str(grid.value(row, label_col))[:60] for row, _, _ in group[:5]],
                confidence=0.6,
            )
            value_profile = ColumnProfile(
                index=value_col,
                letter=_letter(value_col),
                header="مبلغ",
                header_address=f"{_letter(value_col)}{first_row}",
                data_type=ValueKind.NUMBER,
                role=ColumnRole.AMOUNT,
                numeric_ratio=1.0,
                non_empty=len(group),
                distinct=len(group),
                samples=[str(grid.value(row, value_col))[:30] for row, _, _ in group[:5]],
                confidence=0.6,
            )
            results.append(
                TableCandidate(
                    sheet=grid.sheet_name,
                    header_row=first_row,
                    first_data_row=first_row,
                    last_data_row=last_row,
                    first_col=label_col,
                    last_col=value_col,
                    columns=[label_profile, value_profile],
                    kind=TableKind.OBLIGATION,
                    confidence=0.55,
                    reasons=[
                        "label/value records without a header row",
                        f"{len(group)} records; label in column {_letter(label_col)}",
                    ],
                    header_values=["شرح / مؤسسه", "مبلغ"],
                    title="لیست برچسب-مقدار",
                )
            )
        return results


def _letter(col: int) -> str:
    from openpyxl.utils import get_column_letter

    return get_column_letter(col)
