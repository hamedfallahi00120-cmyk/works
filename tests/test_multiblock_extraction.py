from __future__ import annotations

from dataclasses import replace

from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.excel.data_extractor import DataExtractor
from local_office_ai.excel.reader import WorkbookReader
from local_office_ai.excel.row_classification import (
    ROW_CATEGORY,
    ROW_DATA,
    ROW_GROUP_HEADING,
    ROW_NOTE,
    ROW_SUBTOTAL,
    ROW_TOTAL,
    classify_block_rows,
    summary_rows_from_roles,
)
from local_office_ai.excel.sheet_analyzer import (
    SheetGrid,
    is_grand_total_label,
    is_subtotal_label,
    is_total_label,
    summary_label_role,
)
from local_office_ai.excel.table_detector import TableDetector
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


def _detect(make_workbook, settings, rows: list[list]):
    path = make_workbook("blocks.xlsx", {"S": rows})
    bundle = WorkbookReader(settings).load(path)
    grid = _grid(bundle, settings, "S")
    structure = TableDetector(settings).detect(grid)
    return bundle, grid, structure


def test_two_independent_blocks_are_both_detected(make_workbook, settings):
    bundle, _grid_value, structure = _detect(
        make_workbook,
        settings,
        [
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
        ],
    )
    blocks = structure.tables + structure.key_value_tables
    assert len(blocks) >= 2
    assert {3, 8}.issubset({block.header_row for block in blocks})
    bundle.close()


def test_blocks_with_different_header_structures_are_separate(make_workbook, settings):
    bundle, _grid_value, structure = _detect(
        make_workbook,
        settings,
        [
            ["بخش اول"],
            [],
            ["نام", "مبلغ", "تاریخ"],
            ["الف", 100, "1404/01/01"],
            ["ب", 200, "1404/01/02"],
            [],
            [],
            ["دسته", "تعداد", "درصد"],
            ["x", 5, 10],
            ["y", 7, 20],
        ],
    )
    assert len(structure.tables) == 2
    assert {table.header_row for table in structure.tables} == {3, 8}
    headers = {tuple(table.header_values) for table in structure.tables}
    assert ("نام", "مبلغ", "تاریخ") in headers
    assert ("دسته", "تعداد", "درصد") in headers
    bundle.close()


def test_subtotal_in_middle_stays_inside_table(make_workbook, settings):
    bundle, grid, structure = _detect(
        make_workbook,
        settings,
        [
            ["گزارش"],
            [],
            ["گروه", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            ["جمع گروه", 300],
            ["ج", 400],
            ["د", 500],
            ["جمع کل", 900],
        ],
    )
    assert len(structure.tables) == 1
    table = structure.tables[0]
    assert table.first_data_row == 4
    assert table.last_data_row == 8
    assert table.subtotal_rows == [6]
    assert table.grand_total_rows == [9]
    assert table.block_end == 9

    roles = classify_block_rows(
        grid, table, table.header_rows, block_end_row=table.block_end
    )
    role_map = {record.row: record.role for record in roles}
    assert role_map[4] == ROW_DATA
    assert role_map[6] == ROW_SUBTOTAL
    assert role_map[8] == ROW_DATA
    assert role_map[9] == ROW_TOTAL
    assert summary_rows_from_roles(roles) == {"subtotal": [6], "total": [9]}

    totals = DataExtractor(settings).find_declared_totals(grid, table)
    assert [item.label for item in totals] == ["جمع کل"]
    assert totals[0].value == 900
    bundle.close()


def test_multiple_subtotals_remain_inside_the_block(make_workbook, settings):
    bundle, grid, structure = _detect(
        make_workbook,
        settings,
        [
            ["گزارش"],
            [],
            ["گروه", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            ["جمع گروه", 300],
            ["ج", 400],
            ["د", 500],
            ["جمع گروه", 900],
            ["ه", 600],
            ["جمع کل", 1500],
        ],
    )
    table = structure.tables[0]
    assert table.last_data_row == 10
    assert table.block_end == 11
    assert table.subtotal_rows == [6, 9]
    assert table.grand_total_rows == [11]
    roles = {
        record.row: record.role
        for record in classify_block_rows(
            grid, table, table.header_rows, block_end_row=table.block_end
        )
    }
    assert roles[6] == ROW_SUBTOTAL
    assert roles[7] == ROW_DATA
    assert roles[9] == ROW_SUBTOTAL
    assert roles[10] == ROW_DATA
    assert roles[11] == ROW_TOTAL
    bundle.close()


def test_grand_total_ends_block_and_is_not_data(make_workbook, settings):
    bundle, grid, structure = _detect(
        make_workbook,
        settings,
        [
            ["گزارش"],
            [],
            ["کالا", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            ["جمع کل", 300],
        ],
    )
    table = structure.tables[0]
    assert table.last_data_row == 5
    assert table.block_end == 6
    assert table.subtotal_rows == []
    assert table.grand_total_rows == [6]
    roles = classify_block_rows(
        grid, table, table.header_rows, block_end_row=table.block_end
    )
    role_map = {record.row: record.role for record in roles}
    assert role_map[5] == ROW_DATA
    assert role_map[6] == ROW_TOTAL
    bundle.close()


def test_row_classification_distinguishes_data_categories_and_notes(
    make_workbook, settings
):
    bundle, grid, structure = _detect(
        make_workbook,
        settings,
        [
            ["گزارش"],
            [],
            ["گروه", "مبلغ", "توضیح"],
            ["گروه الف", None, None],
            ["الف", 100, "x"],
            ["ب", 200, "y"],
            ["دسته ب", None, "شرح"],
            ["ج", None, "z"],
            ["توضیح: داده‌ها نمونه است", None, None],
            ["جمع کل", 300, None],
        ],
    )
    table = structure.tables[0]
    extended = replace(table, last_data_row=7, block_end_row=7)
    roles = {
        record.row: record.role
        for record in classify_block_rows(
            grid, extended, extended.header_rows, block_end_row=7
        )
    }
    assert roles[4] == ROW_GROUP_HEADING
    assert roles[5] == ROW_DATA
    assert roles[7] == ROW_CATEGORY
    bundle.close()


def test_category_row_with_numeric_code_is_not_data(make_workbook, settings):
    bundle, grid, structure = _detect(
        make_workbook,
        settings,
        [
            ["دسته", "کد", "مبلغ"],
            ["دسته الف", 1, None],
            ["الف", 2, 100],
            ["ب", 3, 200],
            ["جمع کل", None, 300],
        ],
    )
    table = structure.tables[0]
    roles = {
        record.row: record.role
        for record in classify_block_rows(
            grid, table, table.header_rows, block_end_row=table.block_end
        )
    }
    assert roles[2] == ROW_CATEGORY
    assert roles[3] == ROW_DATA
    assert roles[4] == ROW_DATA
    bundle.close()


def test_note_before_grand_total_is_descriptive(make_workbook, settings):
    bundle, grid, structure = _detect(
        make_workbook,
        settings,
        [
            ["کالا", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            ["توضیح: داده‌ها نمونه است", None],
            ["جمع کل", 300],
        ],
    )
    table = structure.tables[0]
    roles = {
        record.row: record.role
        for record in classify_block_rows(
            grid, table, table.header_rows, block_end_row=table.block_end
        )
    }
    assert roles[4] == ROW_NOTE
    assert roles[5] == ROW_TOTAL
    bundle.close()


def test_note_after_grand_total_stays_in_block(make_workbook, settings):
    bundle, grid, structure = _detect(
        make_workbook,
        settings,
        [
            ["کالا", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            ["جمع کل", 300],
            ["توضیح: داده‌ها نمونه است", None],
        ],
    )
    table = structure.tables[0]
    assert table.block_end == 5
    roles = {
        record.row: record.role
        for record in classify_block_rows(
            grid, table, table.header_rows, block_end_row=table.block_end
        )
    }
    assert roles[4] == ROW_TOTAL
    assert roles[5] == ROW_NOTE
    bundle.close()


def test_section_title_after_grand_total_starts_a_new_block(make_workbook, settings):
    bundle, _grid_value, structure = _detect(
        make_workbook,
        settings,
        [
            ["کالا", "مبلغ"],
            ["الف", 100],
            ["ب", 150],
            ["جمع کل", 250],
            ["جمع بندی نتایج"],
            ["کالا", "مبلغ"],
            ["ج", 200],
            ["د", 300],
            ["جمع کل", 500],
        ],
    )
    assert len(structure.tables) == 2
    first, second = structure.tables
    assert first.block_end == 4
    assert second.header_row == 6
    assert second.block_end == 9
    bundle.close()


def test_words_containing_kol_are_not_totals(make_workbook, settings):
    assert not is_total_label("کلانتری ۱۲")
    assert not is_total_label("تکلیف")
    assert not is_total_label("جمعیت")
    assert not is_total_label("کلاسه پرونده")
    assert not is_total_label("کلاسی")
    assert is_total_label("جمع کل")
    assert is_subtotal_label("جمع گروه")
    assert not is_grand_total_label("جمع گروه")
    assert is_grand_total_label("جمع کل")
    assert not is_grand_total_label("جمع کلاسی")
    assert summary_label_role("کلانتری ۱۲") is None
    assert summary_label_role("جمعیت") is None
    assert summary_label_role("جمع گروه") == "subtotal"
    assert summary_label_role("جمع کل") == "total"

    bundle, _grid_value, structure = _detect(
        make_workbook,
        settings,
        [
            ["کلانتری", "تکلیف", "جمعیت", "کلاسه پرونده"],
            [1, 10, 100, 1000],
            [2, 20, 200, 2000],
            ["جمع کل", 3, 300, 3000],
        ],
    )
    table = structure.tables[0]
    assert table.header_row == 1
    assert table.last_data_row == 3
    assert table.grand_total_rows == [4]
    bundle.close()


def test_blank_rows_between_blocks(make_workbook, settings):
    bundle, _grid_value, structure = _detect(
        make_workbook,
        settings,
        [
            ["کالا", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            [],
            [],
            [],
            ["دسته", "تعداد"],
            ["x", 5],
            ["y", 7],
        ],
    )
    assert {table.header_row for table in structure.tables} == {1, 7}
    bundle.close()


def test_repeated_header_splits_blocks(make_workbook, settings):
    bundle, _grid_value, structure = _detect(
        make_workbook,
        settings,
        [
            ["کالا", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            ["کالا", "مبلغ"],
            ["ج", 300],
            ["د", 400],
        ],
    )
    assert {table.header_row for table in structure.tables} == {1, 4}
    first, second = sorted(structure.tables, key=lambda item: item.header_row)
    assert (first.first_data_row, first.last_data_row) == (2, 3)
    assert (second.first_data_row, second.last_data_row) == (5, 6)
    bundle.close()


def test_blocks_with_different_periods_are_separate(make_workbook, settings):
    bundle, _grid_value, structure = _detect(
        make_workbook,
        settings,
        [
            ["فروردین 1403"],
            [],
            ["کالا", "مبلغ"],
            ["الف", 100],
            ["ب", 200],
            [],
            [],
            ["اردیبهشت 1403"],
            [],
            ["کالا", "مبلغ"],
            ["ج", 300],
            ["د", 400],
        ],
    )
    assert {table.title for table in structure.tables} == {
        "فروردین 1403",
        "اردیبهشت 1403",
    }
    relationships = ExcelIntelligenceWorkflow._block_relationships(structure.tables)
    assert any(item["kind"] == "shared_header_signature" for item in relationships)
    bundle.close()


def test_workflow_analyzes_every_block_on_a_sheet(make_workbook, settings):
    path = make_workbook(
        "workflow_two_blocks.xlsx",
        {
            "S": [
                ["گزارش فروش"],
                [],
                ["کالا", "تعداد", "مبلغ"],
                ["الف", 2, 1000],
                ["ب", 3, 2500],
                ["ج", 1, 500],
                ["د", 4, 700],
                ["ه", 2, 300],
                [],
                [],
                ["مشتری", "مانده"],
                ["الف", 500],
                ["ج", 700],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    sheet = analysis.sheets[0]
    assert len(sheet.block_analyses) == 2
    assert {block["header_row"] for block in sheet.block_analyses} == {3, 11}
    assert len(analysis.extraction.sheets[0].blocks) == 2
    assert {
        block.boundaries["header_row"]
        for block in analysis.extraction.sheets[0].blocks
    } == {3, 11}
    assert all(
        block.boundaries["block_address"]
        for block in analysis.extraction.sheets[0].blocks
    )


def test_workflow_analyzes_every_block_and_keeps_subtotal_traceable(
    make_workbook, settings
):
    path = make_workbook(
        "workflow_blocks.xlsx",
        {
            "S": [
                ["گزارش"],
                [],
                ["گروه", "مبلغ"],
                ["الف", 100],
                ["ب", 200],
                ["جمع گروه", 300],
                ["ج", 400],
                ["د", 500],
                ["جمع کل", 900],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    sheet = analysis.sheets[0]
    assert len(sheet.block_analyses) == 1
    block = sheet.block_analyses[0]
    assert block["last_data_row"] == 8
    assert block["block_end_row"] == 9
    assert block["summary_rows"] == {"subtotal": [6], "total": [9]}
    assert block["row_role_counts"]["subtotal"] == 1
    assert block["row_role_counts"]["total"] == 1
    measure = next(item for item in block["statistics"] if item["header"] == "مبلغ")
    assert measure["total"] == "1200"

    roles = {item["row"]: item["role"] for item in sheet.row_roles}
    assert roles[6] == ROW_SUBTOTAL
    assert roles[9] == ROW_TOTAL

    extraction = analysis.extraction
    assert extraction is not None
    extraction_block = extraction.sheets[0].blocks[0]
    assert extraction_block.boundaries["block_end_row"] == 9
    assert extraction_block.boundaries["subtotal_rows"] == [6]
    assert extraction_block.boundaries["grand_total_rows"] == [9]
    assert extraction_block.summary_rows == {"subtotal": [6], "total": [9]}
