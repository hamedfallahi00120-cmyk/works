"""Regression tests for the generic workbook structure-intelligence layer."""

from __future__ import annotations

from decimal import Decimal

import pytest
from openpyxl import Workbook

from local_office_ai.ai.analyzer import AIAnalyzer
from local_office_ai.ai.context_builder import build_fact_chunks
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.models import ColumnRole, DataRole
from local_office_ai.core.settings import load_settings


def _fill_sheet(ws, offset_col: int, days: int, title_year: str, *, shift: bool = False) -> None:
    c0 = offset_col
    ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 7)
    ws.cell(1, c0, f"گزارش روزانه {title_year}")
    ws.cell(2, c0, "ردیف")
    ws.cell(2, c0 + 1, "تاریخ")
    for start, end, text in [
        (2, 3, "تعداد اعمال قانون"),
        (4, 5, "تعداد تخلفات"),
        (6, 7, "تعداد جریمه"),
    ]:
        ws.merge_cells(start_row=2, start_column=c0 + start, end_row=2, end_column=c0 + end)
        ws.cell(2, c0 + start, text)
    for index, header in enumerate([None, None, "1403", "1404", "1403", "1404", "1403", "1404"]):
        if header:
            ws.cell(3, c0 + index, header)
    for day in range(1, days + 1):
        row = 3 + day
        for index, value in enumerate([day, day, 100 + day, 200 + day, 10 + day, 20 + day, 30 + day, 40 + day]):
            ws.cell(row, c0 + index, value)
    total_row = 4 + days
    ws.cell(total_row, c0, "جمع")
    for index in range(2, 8):
        column = chr(65 + c0 + index - 1)
        ws.cell(total_row, c0 + index, f"=SUM({column}4:{column}{total_row - 1})")
    ws.cell(total_row, c0 + 3, 0)  # deliberately wrong declared total
    percentage_row = total_row + 1
    ws.cell(percentage_row, c0, "درصد")
    ws.cell(percentage_row, c0 + 2, f"={chr(65 + c0 + 3)}{total_row}/{chr(65 + c0 + 2)}{total_row}")
    comparison_row = percentage_row + 1
    ws.cell(comparison_row, c0, "مقایسه")
    ws.cell(comparison_row, c0 + 2, f"={chr(65 + c0 + 2)}{total_row}-{chr(65 + c0 + 3)}{total_row}")
    ws.cell(comparison_row + 2, c0, "توضیح: داده آزمایشی است.")


def _build_workbook(path) -> None:
    workbook = Workbook()
    first = workbook.active
    first.title = "فروردین 1403"
    _fill_sheet(first, 1, 31, "1404")
    first["H7"] = "#DIV/0!"
    second = workbook.create_sheet("اردیبهشت 1403")
    _fill_sheet(second, 3, 31, "1403", shift=True)
    third = workbook.create_sheet("آبان 1403")
    _fill_sheet(third, 1, 18, "1403")
    third["H5"] = None
    notes = workbook.create_sheet("توضیحات")
    notes["A1"] = "این برگه فقط توضیحات است."
    workbook.save(path)


@pytest.fixture(scope="module")
def structured_analysis(tmp_path_factory):
    root = tmp_path_factory.mktemp("structure-project")
    (root / "config").mkdir()
    (root / "data").mkdir()
    (root / "logs").mkdir()
    path = root / "structure_sample.xlsx"
    _build_workbook(path)
    settings = load_settings(project_root=root)
    return ExcelIntelligenceWorkflow(settings).analyze_only(path)


def test_two_level_headers_and_merged_parent(structured_analysis):
    sheet = structured_analysis.intelligence.sheet("فروردین 1403")
    assert sheet.header_rows == [2, 3]
    assert sheet.row_count == 31
    assert sheet.column_count == 8
    columns = {column.letter: column for column in sheet.columns}
    assert columns["A"].role is ColumnRole.ID
    assert columns["B"].role is ColumnRole.DATE
    assert columns["B"].statistical_role is DataRole.TEMPORAL_KEY
    assert columns["C"].header == "تعداد اعمال قانون | 1403"
    assert columns["C"].parent_header == "تعداد اعمال قانون"
    assert columns["C"].sub_header == "1403"
    assert columns["C"].period == "1403"
    assert columns["D"].header == "تعداد اعمال قانون | 1404"
    assert columns["D"].period == "1404"
    assert "ادغام" in columns["D"].source_address


def test_full_span_and_shifted_table(structured_analysis):
    intelligence = structured_analysis.intelligence
    assert intelligence.sheet("فروردین 1403").table_address == "A2:H34"
    shifted = intelligence.sheet("اردیبهشت 1403")
    assert shifted.table_address == "C2:J34"
    assert shifted.column_count == 8
    groups = [group for group in intelligence.similar_groups if "فروردین 1403" in group["sheets"]]
    assert groups and set(groups[0]["sheets"]) == {"فروردین 1403", "اردیبهشت 1403", "آبان 1403"}


def test_overlapping_header_spans_are_not_duplicated(structured_analysis):
    tables = [table for table in structured_analysis.tables if table.sheet == "فروردین 1403"]
    assert [table.address for table in tables] == ["A2:H34"]


def test_trend_uses_observation_order_for_day_keys(structured_analysis):
    trends = structured_analysis.intelligence.sheet("فروردین 1403").generic.trends
    assert trends
    assert trends[0].calendar == "observation"
    assert trends[0].granularity == "observation"
    assert trends[0].date_header == "تاریخ"
    assert trends[0].direction in {"افزایشی", "کاهشی", "بدون تغییر"}
    assert len(trends[0].points) == 31


def test_identifier_and_temporal_key_are_not_measures(structured_analysis):
    sheet = structured_analysis.intelligence.sheet("فروردین 1403")
    generic = sheet.generic
    assert generic.row_count == 31
    assert generic.quality.identifier_columns == ["ردیف"]
    assert generic.quality.temporal_columns == ["تاریخ"]
    measures = [column for column in generic.columns if not column.is_identifier and not column.is_temporal_key]
    assert len(measures) == 6
    assert all(column.header.endswith("1403") or column.header.endswith("1404") for column in measures)


def test_identifier_and_date_do_not_enter_generic_calculations(structured_analysis):
    labels = [item.label for item in structured_analysis.calculations]
    assert not any(label in {"جمع ردیف", "میانگین ردیف"} for label in labels)
    assert not any(label in {"جمع تاریخ", "میانگین تاریخ"} for label in labels)


def test_total_percentage_and_note_rows_are_separated(structured_analysis):
    sheet = structured_analysis.intelligence.sheet("فروردین 1403")
    roles = {item["row"]: item["role"] for item in sheet.row_roles}
    assert roles[35] == "total"
    assert roles[36] == "percentage"
    assert roles[37] == "difference"
    assert roles[39] == "footer"
    assert sheet.coverage.observed_observations == 31
    mismatch = next(audit for audit in sheet.audits if audit.address == "D35")
    assert mismatch.status.value == "mismatch"
    assert mismatch.declared == Decimal("0")
    assert mismatch.computed == Decimal("6696")


def test_incomplete_period_and_blank_are_not_zero(structured_analysis):
    sheet = structured_analysis.intelligence.sheet("آبان 1403")
    coverage = sheet.coverage
    assert coverage.expected_observations == 30
    assert coverage.observed_observations == 18
    assert coverage.missing_observations == 12
    assert coverage.missing_labels == [str(day) for day in range(19, 31)]
    assert coverage.blank_cells == 1
    assert coverage.explicit_zero_cells == 0
    assert coverage.incomplete_period is True


def test_year_conflict_between_sheet_name_and_title(structured_analysis):
    sheet = structured_analysis.intelligence.sheet("فروردین 1403")
    assert sheet.sheet_years == ["1403"]
    assert sheet.title_years == ["1404"]
    kinds = {conflict.kind for conflict in sheet.metadata_conflicts}
    assert "sheet_name_vs_title_year" in kinds
    assert sheet.derived is False


def test_formula_error_is_counted_separately(structured_analysis):
    sheet = structured_analysis.intelligence.sheet("فروردین 1403")
    assert sheet.coverage.error_cells == 1
    assert sheet.coverage.blank_cells == 0
    measure = next(column for column in sheet.generic.columns if column.header == "تعداد جریمه | 1404")
    assert measure.count == 30
    assert measure.missing == 0


def test_period_comparison_and_cross_sheet_aggregation(structured_analysis):
    intelligence = structured_analysis.intelligence
    comparison = next(
        item
        for item in intelligence.sheet("فروردین 1403").period_comparisons
        if item.metric == "تعداد اعمال قانون"
    )
    assert comparison.total_a == Decimal("3596")
    assert comparison.total_b == Decimal("6696")
    assert comparison.percent == Decimal("86.21")

    aggregation = next(
        item
        for item in intelligence.aggregations
        if item.metric == "تعداد اعمال قانون" and item.period == "1403"
    )
    assert aggregation.total == Decimal("7192")
    assert set(aggregation.sheets) == {"فروردین 1403", "اردیبهشت 1403"}
    assert aggregation.excluded_incomplete == ["آبان 1403"]


def test_unanalyzed_footer_sheet_is_reported(structured_analysis):
    intelligence = structured_analysis.intelligence
    assert intelligence.sheet("توضیحات").status == "not_analyzed"
    assert intelligence.status_counts["not_analyzed"] == 1


def test_chunked_ai_context_fits_budget(structured_analysis, ollama_settings, fake_ollama):
    chunks = build_fact_chunks(structured_analysis, ollama_settings)
    assert len(chunks) >= 2
    assert chunks[0][0] == "نمای کلی Workbook"
    analyzer = AIAnalyzer(ollama_settings)
    outcome = analyzer.analyze(structured_analysis, use_ai=True)
    assert outcome.mode == "model"
    assert outcome.metadata["chunk_count"] == len(chunks)
    assert outcome.metadata["successful_chunks"] == len(chunks)
    assert outcome.metadata["failed_chunks"] == 0
    assert outcome.metadata["max_prompt_chars"] <= ollama_settings.max_ai_prompt_chars
    assert outcome.metadata["max_context_chars"] <= ollama_settings.max_ai_context_chars
    assert all(chunk["prompt_chars"] <= ollama_settings.max_ai_prompt_chars for chunk in outcome.metadata["chunks"])
    assert all(chunk["context_chars"] <= ollama_settings.max_ai_context_chars for chunk in outcome.metadata["chunks"])
    assert all(
        chunk["estimated_prompt_tokens"] <= ollama_settings.ollama_num_ctx - 1024
        for chunk in outcome.metadata["chunks"]
    )
    assert len([request for request in fake_ollama.requests if request[0] == "POST"]) == len(chunks)


def test_context_overflow_retries_and_falls_back_cleanly(structured_analysis, ollama_settings, fake_ollama):
    fake_ollama.chat_status = 400
    outcome = AIAnalyzer(ollama_settings).analyze(structured_analysis, use_ai=True)
    assert outcome.mode == "offline"
    assert "400" in outcome.error
    assert outcome.metadata["provenance"]["rule_based_fallback"] == "executed"
    assert outcome.metadata["provenance"]["model_analysis"] == "not_executed"
    assert outcome.metadata["failed_chunks"] == outcome.metadata["chunk_count"] > 0
    assert all(chunk["status"] == "failed" for chunk in outcome.metadata["chunks"])


def test_year_month_day_columns_are_not_measures(settings, make_workbook):
    """G5 / R4: سال، ماه و روز are temporal keys, never summed measures.

    Red until the Excel role vocabulary learns the temporal words the document
    reader already knows (P1).
    """
    path = make_workbook(
        "temporal_keys.xlsx",
        {
            "S": [
                ["گزارش آزمایشی"],
                [],
                ["سال", "ماه", "روز", "تعداد"],
                [1403, 1, 5, 10],
                [1403, 2, 17, 12],
                [1404, 1, 9, 14],
                [1404, 2, 28, 11],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    sheet = analysis.intelligence.sheet("S")
    columns = {column.header: column for column in sheet.columns}
    for header in ("سال", "ماه", "روز"):
        assert columns[header].statistical_role is DataRole.TEMPORAL_KEY
    assert {column.header for column in sheet.measures} == {"تعداد"}
    labels = {item.label for item in analysis.calculations}
    for header in ("سال", "ماه", "روز"):
        assert not any(label.endswith(f" {header}") for label in labels), (header, labels)


def test_code_columns_stay_identifiers(settings, make_workbook):
    """G5 / R5: two code columns stay identifiers and never become measures.

    R5's duplicate-role demotion is unreachable today because parent_header is
    always populated, so both columns already stay identifiers. This guards the
    invariant while the role vocabulary grows in P1.
    """
    path = make_workbook(
        "code_columns.xlsx",
        {
            "S": [
                ["گزارش کدها"],
                [],
                ["کد پلاک", "کد شهرستان", "تعداد"],
                [2184, 11, 5],
                [2079, 12, 7],
                [2047, 13, 9],
                [2185, 14, 4],
                [2071, 15, 6],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    sheet = analysis.intelligence.sheet("S")
    columns = {column.header: column for column in sheet.columns}
    for header in ("کد پلاک", "کد شهرستان"):
        assert columns[header].statistical_role is DataRole.IDENTIFIER
    assert {column.header for column in sheet.measures} == {"تعداد"}
