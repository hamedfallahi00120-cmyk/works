"""Focused tests for the Prompt 01A extraction model and period evidence."""

from __future__ import annotations

from local_office_ai.app.workflow import ExcelIntelligenceWorkflow


def _analyze(settings, path):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    assert analysis.extraction is not None
    return analysis


def _first_block(analysis, sheet_name):
    sheet = analysis.extraction.sheet(sheet_name)
    assert sheet is not None
    assert sheet.blocks, f"no block detected on sheet {sheet_name!r}"
    return sheet.blocks[0]


def _tokens(periods):
    return {item["token"] for item in periods}


def test_period_detected_from_column_headers(settings, make_workbook):
    """A period written in a multi-level/single header is found with its address."""
    path = make_workbook(
        "headers_only.xlsx",
        {
            "داده": [
                ["شرح", "شاخص", "تعداد", "1403", "1404"],
                ["الف", "x", 10, 11, 21],
                ["ب", "y", 12, 13, 23],
                ["ج", "z", 14, 15, 25],
            ]
        },
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "داده")

    assert {"1403", "1404"} <= _tokens(block.periods)
    for period in block.periods:
        if period["token"] in {"1403", "1404"}:
            assert period["sources"] == ["column_header"]
    assert {item["address"] for item in block.period_evidence} == {"D1", "E1"}
    assert block.period_status == "detected"
    assert analysis.intelligence.sheet("داده").period_status == "detected"


def test_period_detected_from_worksheet_name(settings, make_workbook):
    """When only the worksheet name carries the period, it is still recorded."""
    path = make_workbook(
        "sheet_name_only.xlsx",
        {
            "فروردین 1403": [
                ["شرح", "مقدار", "تعداد"],
                ["الف", 10, 1],
                ["ب", 12, 2],
                ["ج", 14, 3],
            ]
        },
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "فروردین 1403")

    month_year = [item for item in block.periods if item["token"] == "فروردین 1403"]
    assert month_year and month_year[0]["kind"] == "month_year"
    assert month_year[0]["sources"] == ["sheet_name"]
    assert any("1403" in item["token"] for item in analysis.extraction.periods)


def test_period_detected_from_block_title(settings, make_workbook):
    """A title row above the header contributes period evidence with its address."""

    def merge_title(workbook):
        workbook["داده"].merge_cells("A1:C1")

    path = make_workbook(
        "block_title_only.xlsx",
        {
            "داده": [
                ["گزارش روزانه 1403"],
                [],
                ["شرح", "مقدار", "تعداد"],
                ["الف", 10, 1],
                ["ب", 12, 2],
                ["ج", 14, 3],
            ]
        },
        merge_title,
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "داده")

    evidence = [item for item in block.period_evidence if item["source"] == "block_title"]
    assert evidence and evidence[0]["token"] == "1403"
    assert evidence[0]["address"] == "A1"
    assert "1403" in _tokens(block.periods)
    merged = {item.address for item in analysis.extraction.sheet("داده").merged_ranges}
    assert "A1:C1" in merged


def test_scalar_year_evidence_merges_into_month_year(settings, make_workbook):
    """A filename year supports the sheet's month-year token instead of doubling it."""
    path = make_workbook(
        "آمار 1403.xlsx",
        {
            "فروردین 1403": [
                ["شرح", "مقدار", "تعداد"],
                ["الف", 10, 1],
                ["ب", 12, 2],
                ["ج", 14, 3],
            ]
        },
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "فروردین 1403")

    assert _tokens(block.periods) == {"فروردین 1403"}
    period = block.periods[0]
    assert period["sources"] == ["sheet_name", "filename"]
    assert {item["token"] for item in block.period_evidence} == {"فروردین 1403", "1403"}


def test_period_detected_from_filename(settings, make_workbook):
    """The filename is the last evidence source, and it is reported as such."""
    path = make_workbook(
        "آمار روزانه 1404.xlsx",
        {
            "داده": [
                ["شرح", "مقدار", "تعداد"],
                ["الف", 10, 1],
                ["ب", 12, 2],
                ["ج", 14, 3],
            ]
        },
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "داده")

    assert _tokens(block.periods) == {"1404"}
    assert block.periods[0]["sources"] == ["filename"]
    assert analysis.extraction.period_status == "detected"
    assert analysis.extraction.periods[0]["token"] == "1404"


def test_date_and_identifier_columns_are_not_measures(settings, make_workbook):
    """Identifier and temporal columns get their own roles, not numeric measures."""
    path = make_workbook(
        "roles_only.xlsx",
        {
            "داده": [
                ["کد پرسنلی", "تاریخ ثبت", "مبلغ"],
                [2184, "1404/01/15", 1000],
                [2079, "1404/02/03", 2000],
                [2047, "1404/02/10", 3000],
                [2185, "1404/03/01", 4000],
            ]
        },
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "داده")
    roles = {column.header: column.semantic_role for column in block.columns}

    assert roles["کد پرسنلی"] == "identifier"
    assert roles["تاریخ ثبت"] == "temporal_key"
    assert roles["مبلغ"] == "numeric_measure"

    measures = {column.header for column in analysis.intelligence.sheet("داده").measures}
    assert "مبلغ" in measures
    assert not {"کد پرسنلی", "تاریخ ثبت"} & measures


def test_original_cell_references_and_formulas_are_preserved(settings, make_workbook):
    """Cell addresses, raw values and formulas point back to the source cells."""
    path = make_workbook(
        "cell_refs.xlsx",
        {
            "داده": [
                ["شرح", "مقدار", "توضیح"],
                ["الف", 10, "۱۲٬۵۰۰٬۰۰۰"],
                ["ب", 12, "=B2*2"],
                ["ج", 14, "متن"],
            ]
        },
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "داده")

    assert block.boundaries["address"] == "A1:C4"
    cells = {cell.address: cell for cell in block.cells}
    assert cells["A2"].value == "الف"
    assert cells["B2"].value == 10
    assert cells["C2"].value == "۱۲٬۵۰۰٬۰۰۰"
    assert cells["C3"].formula == "=B2*2"
    assert cells["C3"].value is None
    assert cells["C3"].row == 3 and cells["C3"].column == 3


def test_multi_row_headers_and_merged_parent(settings, make_workbook):
    """A merged parent header keeps its span, child identity and cell addresses."""

    def merge_parent(workbook):
        workbook["داده"].merge_cells("C3:D3")

    path = make_workbook(
        "merged_parent.xlsx",
        {
            "داده": [
                ["گزارش آزمایشی"],
                [],
                ["ردیف", "تاریخ", "تعداد اعمال قانون", None],
                [None, None, "1403", "1404"],
                [1, "1403/01/05", 10, 11],
                [2, "1403/01/06", 12, 13],
                [3, "1403/01/07", 14, 15],
            ]
        },
        merge_parent,
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "داده")

    assert block.header_rows == [3, 4]
    assert len(analysis.extraction.sheet("داده").blocks) == 1
    columns = {column.letter: column for column in block.columns}
    assert columns["C"].parent_header == "تعداد اعمال قانون"
    assert columns["C"].sub_header == "1403"
    assert columns["C"].period == "1403"
    assert columns["D"].parent_header == "تعداد اعمال قانون"
    assert columns["D"].sub_header == "1404"
    assert columns["D"].period == "1404"
    assert columns["C"].header_address == "C4"
    assert columns["D"].header_address == "D4"
    assert columns["C"].semantic_role == "numeric_measure"

    merged = {item.address: item for item in block.merged_ranges}
    assert "C3:D3" in merged
    assert merged["C3:D3"].anchor_address == "C3"
    assert merged["C3:D3"].anchor_value == "تعداد اعمال قانون"
    cells = {cell.address: cell for cell in block.cells}
    assert cells["C3"].merged_range == "C3:D3"
    assert {"1403", "1404"} <= _tokens(block.periods)


def test_no_period_is_invented_without_evidence(settings, make_workbook):
    """With no year/month in any source, the model reports not_detected."""
    path = make_workbook(
        "no_period.xlsx",
        {
            "داده": [
                ["شرح", "مقدار", "تعداد"],
                ["الف", 10, 1],
                ["ب", 12, 2],
                ["ج", 14, 3],
            ]
        },
    )
    analysis = _analyze(settings, path)
    block = _first_block(analysis, "داده")

    assert analysis.extraction.period_status == "not_detected"
    assert analysis.extraction.periods == []
    assert analysis.extraction.sheet("داده").periods == []
    assert block.periods == []


def test_filename_and_sheet_name_period_conflict_is_reported(settings, make_workbook):
    """Disagreeing scalar sources are reported, never silently resolved."""
    path = make_workbook(
        "گزارش 1404.xlsx",
        {
            "فروردین 1403": [
                ["شرح", "مقدار", "تعداد"],
                ["الف", 10, 1],
                ["ب", 12, 2],
                ["ج", 14, 3],
            ]
        },
    )
    analysis = _analyze(settings, path)
    sheet = analysis.extraction.sheet("فروردین 1403")
    assert sheet.period_status == "conflicting"
    assert {"فروردین 1403", "1404"} <= _tokens(sheet.periods)

    kinds = {conflict.kind for conflict in analysis.intelligence.sheet("فروردین 1403").metadata_conflicts}
    assert "filename_vs_sheet_name_year" in kinds
