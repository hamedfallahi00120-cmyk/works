"""Duplicate, derived-sheet and safe-aggregation control.

Every workbook here is a real .xlsx file created by openpyxl; nothing is mocked.
The engine records duplicates and exclusions but never deletes or repairs a row.
"""

from __future__ import annotations

from decimal import Decimal

import openpyxl

from local_office_ai.app.coordinator import Coordinator
from local_office_ai.app.folder_workflow import FolderWorkflow
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.excel.duplicate_control import (
    AggregationDecision,
    AggregationStatus,
    CheckStatus,
    DuplicateControlReport,
    RelationKind,
    SheetRole,
)
from local_office_ai.excel.statistical_reconciliation import ReconciliationStatus

DETAIL = [
    ["گزارش تیر 1403"],
    [],
    ["شعبه", "تعداد پرونده"],
    ["الف", 10],
    ["ب", 20],
    ["ج", 30],
]


def _folder_report(settings, folder, output):
    result = FolderWorkflow(settings).run(folder, use_ai=False, output=output)
    assert result.duplicate_control is not None
    return result


def test_exact_duplicate_tables_in_different_worksheets(make_workbook, settings):
    path = make_workbook("dup.xlsx", {"پرونده‌ها": DETAIL, "کپی": DETAIL})
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    report = analysis.duplicate_control

    assert len(report.duplicate_groups) == 1
    group = report.duplicate_groups[0]
    assert group.identical
    assert {member.sheet for member in group.members} == {"پرونده‌ها", "کپی"}

    sheets = {sheet.sheet: sheet for sheet in report.sheets}
    assert sheets["پرونده‌ها"].role is SheetRole.PRIMARY
    assert sheets["کپی"].role is SheetRole.DUPLICATE
    assert sheets["کپی"].duplicate_of

    blocked = [item for item in report.decisions if item.kind == "duplicate_block"]
    assert blocked and blocked[0].status is AggregationStatus.BLOCKED
    checks = {check.name: check for check in blocked[0].checks}
    assert checks["source_independence"].status is CheckStatus.BLOCK
    assert checks["coverage_non_overlap"].status is CheckStatus.BLOCK

    kinds = {item.kind for item in report.exclusions}
    assert {"duplicate_table", "sheet_duplicate"} <= kinds
    payload = report.to_dict()
    assert payload["counts"]["duplicate_groups"] == 1
    assert payload["counts"]["blocked"] >= 1


def test_copied_summaries_in_different_workbooks(make_workbook, settings, tmp_path):
    summary = [
        ["خلاصه گزارش"],
        [],
        ["شعبه", "تعداد پرونده"],
        ["الف", 10],
        ["ب", 20],
        ["جمع", 30],
    ]
    make_workbook("folder/a.xlsx", {"خلاصه": summary})
    make_workbook("folder/b.xlsx", {"خلاصه": summary})

    result = _folder_report(settings, tmp_path / "folder", tmp_path / "out.xlsx")
    report = result.duplicate_control

    assert len(report.duplicate_groups) == 1
    group = report.duplicate_groups[0]
    assert group.kind == "summary"
    assert {member.source for member in group.members} == {"a.xlsx", "b.xlsx"}
    assert any(item.status is AggregationStatus.BLOCKED for item in report.decisions)
    assert any(item.kind == "duplicate_table" for item in report.exclusions)
    assert (
        report.cross_source_status("تعداد پرونده", ["a.xlsx", "b.xlsx"])
        is AggregationStatus.BLOCKED
    )

    workbook = openpyxl.load_workbook(result.output_path)
    assert "کنترل تکرار و تجمیع بین فایل‌ها" in workbook.sheetnames


def test_formula_based_summary_sheet_is_not_independent(make_workbook, settings, tmp_path):
    path = make_workbook(
        "formula.xlsx",
        {
            "جزئیات": DETAIL,
            "خلاصه": [
                ["خلاصه گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده", "کنترل"],
                ["الف", 10, "=جزئیات!B4"],
                ["ب", 20, "=جزئیات!B5"],
                ["ج", 30, "=جزئیات!B6"],
            ],
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    report = analysis.duplicate_control
    sheets = {sheet.sheet: sheet for sheet in report.sheets}

    assert sheets["خلاصه"].role is SheetRole.SUMMARY
    assert "جزئیات" in sheets["خلاصه"].referenced_sheets
    assert sheets["خلاصه"].cross_sheet_formula_cells >= 3
    assert any("فرمول" in line or "ارجاع" in line for line in sheets["خلاصه"].evidence)

    decision = next(item for item in report.decisions if item.kind == "measure")
    assert decision.status is AggregationStatus.BLOCKED
    independence = next(
        check for check in decision.checks if check.name == "source_independence"
    )
    assert independence.status is CheckStatus.BLOCK
    assert "ارجاع" in independence.explanation

    audit = Coordinator(settings, console=False).run(
        path, use_ai=False, output=tmp_path / "audit.xlsx", profile="audit"
    )
    assert "کنترل تکرار و تجمیع" in openpyxl.load_workbook(audit.output_path).sheetnames


def test_parent_and_child_categories_are_recorded(make_workbook, settings):
    path = make_workbook(
        "parent.xlsx",
        {
            "گروه‌ها": [
                ["گزارش تیر 1403"],
                [],
                ["گروه", "تعداد"],
                ["اداری", 30],
                ["الف", 10],
                ["ب", 20],
                ["عملیاتی", 25],
                ["ج", 10],
                ["د", 15],
                ["جمع", 55],
            ]
        },
    )
    report = ExcelIntelligenceWorkflow(settings).analyze_only(path).duplicate_control

    parents = {
        relation.parent_label: relation
        for relation in report.category_relations
        if relation.kind is RelationKind.PARENT_OF
    }
    assert {"اداری", "عملیاتی"} <= set(parents)
    assert parents["اداری"].child_labels == ["الف", "ب"]
    assert parents["اداری"].confirmed is False
    assert parents["اداری"].parent_value == "30"
    assert parents["اداری"].child_sum == "30"
    assert any("والد/فرزند احتمالی" in warning for warning in report.warnings)

    grand = [
        relation
        for relation in report.category_relations
        if relation.kind is RelationKind.GRAND_TOTAL_OF
    ]
    assert len(grand) == 1
    assert grand[0].parent_label == "جمع"
    assert grand[0].child_labels == ["اداری", "عملیاتی"]
    assert grand[0].confirmed
    assert grand[0].parent_value == grand[0].child_sum == "55"


def test_subtotal_and_grand_total_rows_are_excluded(make_workbook, settings):
    path = make_workbook(
        "subtotal.xlsx",
        {
            "پرونده‌ها": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
                ["جمع", 30],
                ["ج", 5],
                ["د", 5],
                ["جمع کل", 40],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    report = analysis.duplicate_control
    block = analysis.extraction.sheets[0].blocks[0]

    assert block.boundaries["subtotal_rows"] == [6]
    assert block.boundaries["grand_total_rows"] == [9]
    kinds = {item.kind for item in report.exclusions}
    assert {"subtotal_row", "grand_total_row"} <= kinds
    assert [row["row"] for row in analysis.normalized_rows] == [4, 5, 7, 8]

    grand = [
        relation
        for relation in report.category_relations
        if relation.kind is RelationKind.GRAND_TOTAL_OF
    ]
    assert grand and grand[0].confirmed
    assert grand[0].parent_value == grand[0].child_sum == "40"


def test_overlapping_administrative_classifications_block_aggregation(
    make_workbook, settings, tmp_path
):
    make_workbook(
        "folder/a.xlsx",
        {"پرونده‌ها": DETAIL[:3] + [["الف", 10], ["ب", 20]]},
    )
    make_workbook(
        "folder/b.xlsx",
        {"پرونده‌ها": DETAIL[:3] + [["ب", 20], ["ج", 30]]},
    )
    result = _folder_report(settings, tmp_path / "folder", tmp_path / "out.xlsx")
    report = result.duplicate_control

    overlaps = [
        relation
        for relation in report.category_relations
        if relation.kind is RelationKind.OVERLAP
    ]
    assert overlaps
    assert any("ب" in line for line in overlaps[0].evidence)
    assert any(item.kind == "overlapping_group" for item in report.exclusions)

    decision = next(item for item in report.decisions if item.kind == "measure")
    assert decision.status is AggregationStatus.BLOCKED
    coverage = next(
        check for check in decision.checks if check.name == "coverage_non_overlap"
    )
    assert coverage.status is CheckStatus.BLOCK
    assert "مشترک" in coverage.explanation
    assert (
        report.cross_source_status("تعداد پرونده", ["a.xlsx", "b.xlsx"])
        is AggregationStatus.BLOCKED
    )


def test_independent_sources_with_same_metric_are_allowed(
    make_workbook, settings, tmp_path
):
    make_workbook(
        "folder/a.xlsx",
        {"پرونده‌ها": DETAIL[:3] + [["الف", 10], ["ب", 20]]},
    )
    make_workbook(
        "folder/b.xlsx",
        {"پرونده‌ها": DETAIL[:3] + [["ج", 30], ["د", 40]]},
    )
    result = _folder_report(settings, tmp_path / "folder", tmp_path / "out.xlsx")
    report = result.duplicate_control

    assert not report.duplicate_groups
    decisions = [item for item in report.decisions if item.kind == "measure"]
    assert len(decisions) == 1
    assert decisions[0].status is AggregationStatus.ALLOWED
    assert all(check.status is CheckStatus.PASS for check in decisions[0].checks)
    coverage = next(
        check for check in decisions[0].checks if check.name == "coverage_non_overlap"
    )
    assert "هیچ مقدار مشترکی" in coverage.explanation
    assert (
        report.cross_source_status("تعداد پرونده", ["a.xlsx", "b.xlsx"])
        is AggregationStatus.ALLOWED
    )


def test_similar_looking_but_independent_datasets_are_not_duplicates(
    make_workbook, settings, tmp_path
):
    make_workbook(
        "folder/a.xlsx",
        {
            "پرونده‌ها": DETAIL[:2]
            + [["شعبه", "تعداد پرونده ورودی"], ["الف", 10], ["ب", 20]]
        },
    )
    make_workbook(
        "folder/b.xlsx",
        {
            "پرونده‌ها": DETAIL[:2]
            + [["شعبه", "تعداد پرونده خروجی"], ["الف", 10], ["ب", 20]]
        },
    )
    result = _folder_report(settings, tmp_path / "folder", tmp_path / "out.xlsx")
    report = result.duplicate_control

    assert not report.duplicate_groups
    assert all(sheet.role is SheetRole.PRIMARY for sheet in report.sheets)
    assert all(
        decision.status is not AggregationStatus.BLOCKED
        for decision in report.decisions
    )


def test_formula_copy_without_summary_labels_is_derived(make_workbook, settings):
    path = make_workbook(
        "derived.xlsx",
        {
            "جزئیات": DETAIL,
            "بازنمایی": [
                ["بازنمایی داده"],
                [],
                ["شعبه", "تعداد پرونده", "کنترل"],
                ["الف", 10, "=جزئیات!B4"],
                ["ب", 20, "=جزئیات!B5"],
                ["ج", 30, "=جزئیات!B6"],
            ],
        },
    )
    report = ExcelIntelligenceWorkflow(settings).analyze_only(path).duplicate_control
    sheets = {sheet.sheet: sheet for sheet in report.sheets}

    assert sheets["بازنمایی"].role is SheetRole.DERIVED
    assert sheets["جزئیات"].role is SheetRole.PRIMARY
    assert "جزئیات" in sheets["بازنمایی"].referenced_sheets
    assert sheets["بازنمایی"].evidence


def test_record_duplicates_are_flagged_and_not_removed(make_workbook, settings):
    path = make_workbook(
        "records.xlsx",
        {
            "پرونده‌ها": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["الف", 10],
                ["ب", 20],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    report = analysis.duplicate_control

    assert len(report.record_duplicates) == 1
    item = report.record_duplicates[0]
    assert item.kind == "identical"
    assert item.key_values == ["الف"]
    assert item.row_numbers == [4, 5]
    assert any(entry.kind == "record_duplicate" for entry in report.exclusions)
    assert [row["row"] for row in analysis.normalized_rows] == [4, 5, 6]


def test_duplicate_control_is_serialized_and_audit_only(make_workbook, settings, tmp_path):
    path = make_workbook("serialize.xlsx", {"پرونده‌ها": DETAIL, "کپی": DETAIL})
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    payload = analysis.to_dict()

    assert payload["duplicate_control"]["counts"]["duplicate_groups"] == 1
    assert payload["duplicate_control"]["sheets"]
    assert payload["duplicate_control"]["decisions"]

    audit = Coordinator(settings, console=False).run(
        path, use_ai=False, output=tmp_path / "audit.xlsx", profile="audit"
    )
    clean = Coordinator(settings, console=False).run(
        path, use_ai=False, output=tmp_path / "clean.xlsx", profile="clean"
    )
    audit_names = set(openpyxl.load_workbook(audit.output_path).sheetnames)
    clean_book = openpyxl.load_workbook(clean.output_path)
    assert "کنترل تکرار و تجمیع" in audit_names
    assert "کنترل تکرار و تجمیع" not in clean_book.sheetnames

    validation = clean_book["اعتبارسنجی داده"]
    text = " ".join(
        str(cell.value)
        for row in validation.iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "کنترل تکرار و تجمیع ایمن" in text


def test_copied_total_row_with_different_detail_is_a_summary_duplicate(
    make_workbook, settings
):
    path = make_workbook(
        "summary_hash.xlsx",
        {
            "جزئیات": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 5],
                ["ب", 5],
                ["ج", 10],
                ["د", 10],
                ["جمع", 30],
            ],
            "بازبینی": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 10],
                ["ج", 10],
                ["جمع", 30],
            ],
        },
    )
    report = ExcelIntelligenceWorkflow(settings).analyze_only(path).duplicate_control

    assert len(report.duplicate_groups) == 1
    group = report.duplicate_groups[0]
    assert group.kind == "summary"
    assert group.identical
    assert len(group.members) == 2
    assert {member.summary_hash for member in group.members} == {
        group.members[0].summary_hash
    }

    sheets = {sheet.sheet: sheet for sheet in report.sheets}
    assert sheets["جزئیات"].role is SheetRole.PRIMARY
    assert sheets["بازبینی"].role is SheetRole.DUPLICATE
    assert any(
        entry.kind == "duplicate_table" and entry.sheet == "بازبینی"
        for entry in report.exclusions
    )


def test_identical_values_in_different_periods_are_not_duplicates(make_workbook, settings):
    rows = [
        ["شعبه", "تعداد پرونده"],
        ["الف", 10],
        ["ب", 20],
        ["ج", 30],
    ]
    path = make_workbook(
        "periods.xlsx",
        {
            "فروردین 1403": [["گزارش فروردین 1403"], [], *rows],
            "اردیبهشت 1403": [["گزارش اردیبهشت 1403"], [], *rows],
        },
    )
    report = ExcelIntelligenceWorkflow(settings).analyze_only(path).duplicate_control

    assert not report.duplicate_groups
    assert all(sheet.role is SheetRole.PRIMARY for sheet in report.sheets)


def test_sheet_with_one_duplicate_block_and_one_unique_block_stays_primary(
    make_workbook, settings
):
    path = make_workbook(
        "mixed_blocks.xlsx",
        {
            "الف": [
                ["جدول دوم"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["ج", 5],
                ["د", 5],
                ["جمع", 10],
            ],
            "ترکیبی": [
                ["جدول اول"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["الف", 10],
                ["ب", 20],
                ["جمع", 30],
                [],
                [],
                ["جدول دوم"],
                [],
                ["شعبه", "تعداد پرونده"],
                ["ج", 5],
                ["د", 5],
                ["جمع", 10],
            ],
        },
    )
    report = ExcelIntelligenceWorkflow(settings).analyze_only(path).duplicate_control
    sheets = {sheet.sheet: sheet for sheet in report.sheets}

    assert sheets["ترکیبی"].role is SheetRole.PRIMARY
    assert sheets["ترکیبی"].data_blocks == 2
    assert sheets["الف"].role is SheetRole.PRIMARY
    assert "mixed_blocks.xlsx/الف" in sheets["ترکیبی"].duplicate_of
    assert any(
        entry.kind == "duplicate_table" and entry.sheet == "ترکیبی"
        for entry in report.exclusions
    )
    assert any("بازنمایی تکراری" in line for line in sheets["ترکیبی"].evidence)


def test_cross_sheet_total_formula_does_not_make_a_data_sheet_derived(
    make_workbook, settings
):
    rows = [[f"شعبه {index}", index] for index in range(1, 11)]
    other_rows = [[f"واحد {index}", index * 2] for index in range(1, 11)]
    path = make_workbook(
        "total_formula.xlsx",
        {
            "جزئیات": [["گزارش تیر 1403"], [], ["شعبه", "تعداد پرونده"], *rows, ["جمع", 55]],
            "داده": [
                ["گزارش تیر 1403"],
                [],
                ["شعبه", "تعداد پرونده"],
                *other_rows,
                ["جمع", "=SUM(جزئیات!B4:B13)"],
            ],
        },
    )
    report = ExcelIntelligenceWorkflow(settings).analyze_only(path).duplicate_control
    sheets = {sheet.sheet: sheet for sheet in report.sheets}

    assert sheets["داده"].cross_sheet_formula_cells >= 1
    assert sheets["داده"].role is SheetRole.PRIMARY


def test_sheet_without_data_is_unknown_and_flagged(make_workbook, settings):
    path = make_workbook("unknown.xlsx", {"خالی": [["یادداشت بدون داده"]]})
    report = ExcelIntelligenceWorkflow(settings).analyze_only(path).duplicate_control
    sheets = {sheet.sheet: sheet for sheet in report.sheets}

    assert sheets["خالی"].role is SheetRole.UNKNOWN
    assert any("نامشخص" in warning for warning in report.warnings)


def test_parent_category_rows_are_not_added_to_children_in_reconciliation(
    make_workbook, settings
):
    path = make_workbook(
        "parent_reconciliation.xlsx",
        {
            "گروه‌ها": [
                ["گزارش تیر 1403"],
                [],
                ["گروه", "تعداد"],
                ["اداری", 30],
                ["الف", 10],
                ["ب", 20],
                ["عملیاتی", 25],
                ["ج", 10],
                ["د", 15],
                ["جمع", 55],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(path)
    result = next(
        item
        for item in analysis.statistical_reconciliation.reconciliations
        if item.kind == "component_total" and item.measure == "تعداد"
    )

    assert result.calculated == Decimal(55)
    assert result.difference == Decimal(0)
    assert result.status is ReconciliationStatus.WARNING
    assert "والد" in result.explanation
    assert "نیازمند بازبینی" in result.explanation


def test_cross_source_status_only_uses_decisions_inside_the_selected_files():
    report = DuplicateControlReport(
        source_count=3,
        decisions=[
            AggregationDecision(
                decision_id="A1",
                metric="تعداد پرونده",
                kind="measure",
                members=[
                    "a.xlsx/پرونده‌ها/A3:B5/«تعداد پرونده»",
                    "b.xlsx/پرونده‌ها/A3:B5/«تعداد پرونده»",
                ],
                status=AggregationStatus.ALLOWED,
                checks=[],
            ),
            AggregationDecision(
                decision_id="A2",
                metric="تعداد پرونده",
                kind="measure",
                members=[
                    "a.xlsx/پرونده‌ها/A3:B5/«تعداد پرونده»",
                    "c.xlsx/پرونده‌ها/A3:B5/«تعداد پرونده»",
                ],
                status=AggregationStatus.BLOCKED,
                checks=[],
            ),
        ],
    )

    assert (
        report.cross_source_status("تعداد پرونده", ["a.xlsx", "b.xlsx"])
        is AggregationStatus.ALLOWED
    )
    assert (
        report.cross_source_status("تعداد پرونده", ["a.xlsx", "c.xlsx"])
        is AggregationStatus.BLOCKED
    )
