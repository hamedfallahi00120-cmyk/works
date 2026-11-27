from __future__ import annotations

import pytest

from local_office_ai.core.models import ColumnRole, TableKind, ValueKind
from local_office_ai.excel.reader import WorkbookReader
from local_office_ai.excel.sheet_analyzer import SheetAnalyzer, SheetGrid
from local_office_ai.excel.table_detector import TableDetector, is_total_label
from local_office_ai.excel.workbook_inspector import inspect_workbook


def _grid(bundle, settings, sheet: str) -> SheetGrid:
    meta = inspect_workbook(bundle, settings)
    sheet_meta = next(item for item in meta.sheets if item.title == sheet)
    return SheetGrid(
        sheet,
        bundle.values[sheet],
        bundle.formulas[sheet],
        sheet_meta.max_row,
        sheet_meta.max_col,
    )


def test_regions_join_small_gaps_and_split_large_ones(make_workbook, settings):
    path = make_workbook(
        "regions.xlsx",
        {
            "S": [
                ["a", "b"],
                [1, 2],
                [],
                [3, 4],
                [],
                [],
                [],
                ["c", "d"],
                [5, 6],
            ]
        },
    )
    bundle = WorkbookReader(settings).load(path)
    grid = _grid(bundle, settings, "S")
    blocks = SheetAnalyzer(grid, settings).regions()
    assert [(block.start, block.end) for block in blocks] == [(1, 4), (8, 9)]
    bundle.close()


def test_header_detection_skips_titles_and_totals(sample_path, settings):
    bundle = WorkbookReader(settings).load(sample_path)
    grid = _grid(bundle, settings, "Sheet1")
    analyzer = SheetAnalyzer(grid, settings)
    header = analyzer.find_header(1, 15)
    assert header is not None
    assert header.row == 2
    assert (header.first_col, header.last_col) == (2, 8)
    bundle.close()


def test_roles_on_reference_table(sample_path, settings):
    bundle = WorkbookReader(settings).load(sample_path)
    grid = _grid(bundle, settings, "Sheet1")
    structure = TableDetector(settings).detect(grid)
    table = structure.tables[0]
    roles = {column.header: column.role for column in table.columns}
    assert table.kind is TableKind.LOAN
    assert roles["ردیف"] is ColumnRole.ID
    assert roles["اقساط قابل پرداخت"] is ColumnRole.INSTITUTION
    assert roles["مبلغ وام"] is ColumnRole.PRINCIPAL
    assert roles["مبلغ هر قسط"] is ColumnRole.INSTALLMENT
    assert roles["تعداد کل قسط"] is ColumnRole.INSTALLMENT_COUNT
    assert roles["تاریخ آخرین قسط"] is ColumnRole.DATE
    assert table.row_count == 6
    assert table.confidence >= 0.9
    # The label/value commitment list below the table is detected separately.
    assert structure.key_value_tables
    assert structure.key_value_tables[0].row_count >= 3
    bundle.close()


def test_two_tables_in_one_sheet(make_workbook, settings):
    path = make_workbook(
        "two_tables.xlsx",
        {
            "S": [
                ["گزارش فروش"],
                [],
                ["کالا", "تعداد", "مبلغ"],
                ["الف", 2, 1000],
                ["ب", 3, 2500],
                [],
                [],
                ["مشتری", "مانده"],
                ["الف", 500],
                ["ج", 700],
            ]
        },
    )
    bundle = WorkbookReader(settings).load(path)
    grid = _grid(bundle, settings, "S")
    structure = TableDetector(settings).detect(grid)
    kinds = [table.kind for table in structure.tables + structure.key_value_tables]
    assert structure.tables
    assert structure.tables[0].header_row == 3
    assert len(structure.tables) + len(structure.key_value_tables) >= 2
    assert TableKind.OBLIGATION in kinds or len(structure.key_value_tables) >= 1
    bundle.close()


def test_side_by_side_tables(make_workbook, settings):
    path = make_workbook(
        "side.xlsx",
        {
            "S": [
                ["نام", "مبلغ", None, "شرح", "تعداد"],
                ["الف", 100, None, "x", 1],
                ["ب", 200, None, "y", 2],
            ]
        },
    )
    bundle = WorkbookReader(settings).load(path)
    grid = _grid(bundle, settings, "S")
    structure = TableDetector(settings).detect(grid)
    assert len(structure.tables) == 2
    assert structure.tables[0].last_col == 2
    assert structure.tables[1].first_col == 4
    bundle.close()


def test_time_series_is_classified(make_workbook, settings):
    rows = [["زمان", "CPU", "RAM", "GPU", "Disk"]]
    for index in range(10):
        rows.append([index, 10 + index, 20 + index, 5 + index, index * 2])
    path = make_workbook("ts.xlsx", {"S": rows})
    bundle = WorkbookReader(settings).load(path)
    grid = _grid(bundle, settings, "S")
    structure = TableDetector(settings).detect(grid)
    assert structure.tables[0].kind is TableKind.TIME_SERIES
    bundle.close()


def test_generic_workbook_structure(generic_workbook, settings):
    bundle = WorkbookReader(settings).load(generic_workbook)
    grid = _grid(bundle, settings, "هزینه‌ها")
    structure = TableDetector(settings).detect(grid)
    table = structure.tables[0]
    assert table.header_row == 3
    roles = {column.header: column.role for column in table.columns}
    assert roles["شرح"] is ColumnRole.DESCRIPTION
    assert roles["مبلغ"] is ColumnRole.AMOUNT
    assert roles["تاریخ"] is ColumnRole.DATE
    assert table.columns[2].data_type is ValueKind.NUMBER
    # Hidden column is still analyzed, not ignored.
    assert any(column.letter == "E" for column in table.columns)
    bundle.close()


def test_total_label_helper():
    assert is_total_label("جمع کل وام ها")
    assert is_total_label("Total:")
    assert not is_total_label("تعهدهای ماهانه")
    assert not is_total_label(None)


@pytest.mark.xfail(
    strict=True,
    reason="Known baseline defect G3/R3; remove this marker when total-label matching is fixed.",
)
def test_total_label_needs_a_word_boundary():
    """G3 / R3: a label that merely contains جمع/کل is not a totals row.

    Red until is_total_label switches to the word-boundary rule that the
    statistics channel already uses (P0).
    """
    for label in ("کلانتری ۱۲", "تکلیف", "جمعیت", "کلاسه پرونده"):
        assert not is_total_label(label), label
    for label in ("جمع", "جمع کل", "مجموع", "Total"):
        assert is_total_label(label), label


def test_low_confidence_structure_is_flagged(make_workbook, settings):
    path = make_workbook(
        "weird.xlsx",
        {
            "S": [
                ["یادداشت اول", None, None],
                [None, "متن", 12],
                [None, "متن دو", 15],
            ]
        },
    )
    bundle = WorkbookReader(settings).load(path)
    grid = _grid(bundle, settings, "S")
    structure = TableDetector(settings).detect(grid)
    for table in structure.tables:
        assert 0.0 <= table.confidence <= 1.0
        assert table.reasons
    bundle.close()
