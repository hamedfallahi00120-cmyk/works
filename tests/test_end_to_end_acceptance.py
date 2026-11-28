"""End-to-end acceptance tests: period rows, accident separation, DOCX, honesty.

Every workbook is a real `.xlsx` file created by openpyxl; every DOCX is a real
package that is reopened and validated.  No mock substitutes a parser.
"""

from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from openpyxl import Workbook

from local_office_ai.app.folder_workflow import FolderWorkflow
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.excel.accident_semantics import (
    KIND_FATALITY,
    KIND_FATAL_ACCIDENT,
    build_accident_semantics,
)
from local_office_ai.excel.indicator_periods import build_indicator_periods
from local_office_ai.quality.acceptance import (
    STATUS_ACCEPTED,
    STATUS_FAILED,
    STATUS_UNVERIFIED,
    AcceptanceCheckpoint,
    AcceptanceRunner,
)
from local_office_ai.reporting.formats import ReportFormatError


def _write(path: Path, sheets: dict[str, list[list]]) -> Path:
    workbook = Workbook()
    first = True
    for name, rows in sheets.items():
        sheet = workbook.active if first else workbook.create_sheet(name)
        if first:
            sheet.title = name
            first = False
        for row_index, row in enumerate(rows, start=1):
            for column_index, value in enumerate(row, start=1):
                if value is not None:
                    sheet.cell(row_index, column_index, value)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    return path


def _violations(path: Path, *, scale: int = 1) -> Path:
    return _write(
        path,
        {
            "تخلفات": [
                ["آمار تخلفات پنج ماهه"],
                [],
                ["شاخص", 1404, 1405],
                ["تخلفات منجر به تصادف", 1000 * scale, 900 * scale],
                ["سایر تخلفات", 2000 * scale, 1850 * scale],
                ["عدم استفاده از کلاه کاسکت", 500 * scale, 430 * scale],
                ["نبستن کمربند ایمنی", 800 * scale, 760 * scale],
                ["جمع کل تخلفات", 4300 * scale, 3940 * scale],
            ]
        },
    )


def _accidents(path: Path) -> Path:
    return _write(
        path,
        {
            "تصادفات": [
                ["آمار تصادفات پنج ماهه 1404 و 1405"],
                [],
                [
                    "نوع تصادف",
                    "تعداد تصادف 1404",
                    "تعداد تصادف 1405",
                    "فوتی 1404",
                    "فوتی 1405",
                    "مجروح 1404",
                    "مجروح 1405",
                ],
                ["فوتی", 120, 90, 130, 95, 0, 0],
                ["جرحی", 850, 700, 0, 0, 1200, 950],
                ["خسارتی", 2100, 1800, 0, 0, 0, 0],
                ["جمع", 3070, 2590, 130, 95, 1200, 950],
            ]
        },
    )


def _analyze(settings, path: Path):
    return ExcelIntelligenceWorkflow(settings, progress=lambda message: None).analyze_only(path)


# ---------------------------------------------------------------------- #
# 1. Period header row and row-oriented indicators
# ---------------------------------------------------------------------- #
def test_label_with_year_headers_is_detected_as_a_table(settings, tmp_path):
    analysis = _analyze(settings, _violations(tmp_path / "violations.xlsx"))

    assert len(analysis.tables) == 1
    table = analysis.tables[0]
    assert table.row_count == 4
    assert [column.header for column in table.columns] == ["شاخص", "1404", "1405"]
    block = analysis.extraction.sheets[0].blocks[0]
    assert block.period_status == "detected"
    assert [
        period.get("label") or period.get("token") for period in block.periods
    ] == ["1404", "1405"]
    assert block.period_conflicts == []


def test_a_totals_row_with_year_columns_is_not_promoted_to_header(settings, tmp_path):
    analysis = _analyze(settings, _violations(tmp_path / "violations.xlsx"))

    block = analysis.extraction.sheets[0].blocks[0]
    assert block.header_rows == [3]
    assert block.summary_rows["total"] == [8]
    assert [column.header for column in block.columns][0] == "شاخص"


def test_indicator_comparison_keeps_each_row_separate(settings, tmp_path):
    analysis = _analyze(settings, _violations(tmp_path / "violations.xlsx"))
    report = build_indicator_periods(analysis, source="violations.xlsx")

    assert report.periods == ["1404", "1405"]
    by_label = {item.indicator: item for item in report.comparisons}
    assert by_label["تخلفات منجر به تصادف"].previous == 1000
    assert by_label["تخلفات منجر به تصادف"].current == 900
    assert by_label["تخلفات منجر به تصادف"].percent == Decimal("-10.00")
    assert by_label["سایر تخلفات"].previous == 2000
    assert by_label["جمع کل تخلفات"].role == "total"
    assert by_label["جمع کل تخلفات"].previous == 4300
    assert by_label["جمع کل تخلفات"].current == 3940
    # No comparison is the sum of unrelated indicators.
    assert all(item.previous != 4300 for label, item in by_label.items() if label != "جمع کل تخلفات")


def test_zero_baseline_has_no_percent(settings, tmp_path):
    path = _write(
        tmp_path / "zero.xlsx",
        {
            "داده": [
                ["شاخص", 1404, 1405],
                ["شاخص آزمایشی", 0, 10],
                ["شاخص دوم", 100, 120],
            ]
        },
    )
    analysis = _analyze(settings, path)
    report = build_indicator_periods(analysis, source="zero.xlsx")

    comparison = report.comparison("شاخص آزمایشی")
    assert comparison is not None
    assert comparison.difference == 10
    assert comparison.percent is None
    assert comparison.status == "zero_baseline"


# ---------------------------------------------------------------------- #
# 2. Accident counts are not casualties
# ---------------------------------------------------------------------- #
def test_accident_counts_and_casualties_stay_separate(settings, tmp_path):
    analysis = _analyze(settings, _accidents(tmp_path / "accidents.xlsx"))
    report = build_accident_semantics(analysis, source="accidents.xlsx")

    assert KIND_FATAL_ACCIDENT in report.detected_kinds
    assert KIND_FATALITY in report.detected_kinds
    assert report.missing == []
    fatal_accidents = next(item for item in report.metrics if item.kind == KIND_FATAL_ACCIDENT)
    fatalities = next(item for item in report.metrics if item.kind == KIND_FATALITY and item.label == "فوتی")
    assert fatal_accidents.label != fatalities.label
    assert any(item.code == "count_vs_casualty::fatal_accident_count" for item in report.findings)
    # The declared total matches its components for every measure column.
    assert all(item.status == "PASS" for item in report.findings if item.code == "declared_total_matches_components")


def test_missing_casualty_statistics_are_declared_not_invented(settings, tmp_path):
    path = _write(
        tmp_path / "accident-count-only.xlsx",
        {
            "تصادفات": [
                ["آمار تصادفات"],
                [],
                ["نوع تصادف", "تعداد تصادف 1404", "تعداد تصادف 1405"],
                ["فوتی", 120, 90],
                ["جرحی", 850, 700],
            ],
        },
    )
    analysis = _analyze(settings, path)
    report = build_accident_semantics(analysis, source=path.name)

    assert any("تعداد افراد فوت‌شده" in item for item in report.missing)
    assert any(item.code == "missing_accident_statistics" for item in report.findings)
    assert KIND_FATALITY not in report.detected_kinds


# ---------------------------------------------------------------------- #
# 3. Multi-source DOCX
# ---------------------------------------------------------------------- #
def test_folder_writes_a_valid_multi_source_docx(settings, tmp_path):
    folder = tmp_path / "sources"
    _violations(folder / "الف.xlsx")
    _violations(folder / "ب.xlsx", scale=2)

    result = FolderWorkflow(settings, progress=lambda message: None).run(
        folder,
        request="گزارش مدیریتی تخلفات",
        use_ai=False,
        output=tmp_path / "report.docx",
        output_format="docx",
    )

    output = Path(result.output_path)
    assert output.is_file()
    assert output.suffix == ".docx"
    assert result.report.output_format.value == "docx"
    assert result.report.writer == "DocxReportWriter"
    assert result.report.validation.valid is True, result.report.validation.errors
    with ZipFile(output) as package:
        assert "word/document.xml" in package.namelist()
    document = Document(output)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert "گزارش تحلیلی چندمنبعی" in text
    assert any("مقایسه شاخص" in item for item in result.report.sections_written)
    assert result.quality_gate is not None
    integrity = next(
        stage for stage in result.quality_gate.stages if stage.stage.name == "DOCUMENT_INTEGRITY"
    )
    assert integrity.status.value == "PASS"
    assert result.indicator_periods is not None
    assert result.indicator_periods.comparisons


def test_folder_rejects_pdf_instead_of_writing_excel(settings, tmp_path):
    folder = tmp_path / "sources"
    _violations(folder / "الف.xlsx")
    with pytest.raises(ReportFormatError):
        FolderWorkflow(settings, progress=lambda message: None).run(
            folder,
            request="گزارش",
            use_ai=False,
            output=tmp_path / "report.pdf",
            output_format="pdf",
        )


# ---------------------------------------------------------------------- #
# 4. Acceptance runner
# ---------------------------------------------------------------------- #
def test_acceptance_passes_on_verified_sources(settings, tmp_path):
    source = _violations(tmp_path / "violations.xlsx")
    runner = AcceptanceRunner(settings, progress=lambda message: None)
    report = runner.run(
        [source],
        request="گزارش مدیریتی تخلفات",
        output=tmp_path / "accepted.docx",
        output_dir=tmp_path / "audit",
        checkpoints=[
            AcceptanceCheckpoint(
                label="تخلفات منجر به تصادف",
                previous="1000",
                current="900",
                reported_percent="-10.00",
            ),
            AcceptanceCheckpoint(
                label="جمع کل تخلفات",
                previous="4300",
                current="3940",
                reported_percent="-8.37",
            ),
        ],
    )

    assert report.status == STATUS_ACCEPTED, report.to_dict()
    assert report.check("checkpoint_mismatch").status == "PASS"
    assert all(item.status == "PASS" for item in report.checks)
    assert Path(report.outputs["path"]).suffix == ".docx"
    assert report.docx["valid"] is True
    json_path, markdown_path = report.write(tmp_path / "audit")
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    assert payload["status"] == STATUS_ACCEPTED
    text = markdown_path.read_text(encoding="utf-8")
    persian = "۱۲۳۴۵۶۷۸۹"
    for index, digit in enumerate(persian, start=1):
        assert f"## {digit}." in text
    assert "بررسی checkpoint" in text


def test_acceptance_without_inputs_claims_nothing(settings, tmp_path):
    runner = AcceptanceRunner(settings, progress=lambda message: None)
    report = runner.run(
        [tmp_path / "missing.xlsx"],
        request="گزارش",
        output=tmp_path / "never.docx",
    )

    assert report.status == STATUS_UNVERIFIED
    assert report.outputs == {}
    assert not (tmp_path / "never.docx").exists()
    assert report.unverified_stages
    assert any("موجود نیست" in item for item in report.limitations)


def test_acceptance_checkpoint_mismatch_is_reported_not_forced(settings, tmp_path):
    source = _violations(tmp_path / "violations.xlsx")
    runner = AcceptanceRunner(settings, progress=lambda message: None)
    report = runner.run(
        [source],
        request="گزارش مدیریتی تخلفات",
        output=tmp_path / "report.docx",
        checkpoints=[
            AcceptanceCheckpoint(
                label="تخلفات منجر به تصادف",
                previous="1000",
                current="999",
                reported_percent="-10.00",
            )
        ],
    )

    assert report.status == STATUS_FAILED
    check = report.check("checkpoint_mismatch")
    assert check.status == "FAIL"
    item = report.checkpoints[0]
    assert item.extracted_current == 900
    assert "999" in item.note
    # The document itself still reports the extracted value, not the checkpoint.
    text = Document(report.outputs["path"]).element.xml
    assert "999" not in text


def test_acceptance_marks_conflicting_periods_as_not_successful(settings, tmp_path):
    path = _write(
        tmp_path / "conflict.xlsx",
        {
            "فروردین": [
                ["آمار تخلفات فروردین 1405"],
                [],
                ["شاخص", 1404, 1405],
                ["تخلفات منجر به تصادف", 1000, 900],
                ["سایر تخلفات", 2000, 1850],
            ]
        },
    )
    runner = AcceptanceRunner(settings, progress=lambda message: None)
    report = runner.run(
        [path],
        request="گزارش مدیریتی تخلفات",
        output=tmp_path / "conflict.docx",
    )

    assert report.status in {STATUS_FAILED, "partial"}
    assert report.check("period_not_identified").status == "FAIL"
    assert report.check("critical_validation_failed").status == "FAIL"


def test_acceptance_writes_audit_files_even_when_incomplete(settings, tmp_path):
    source = _violations(tmp_path / "violations.xlsx")
    report = AcceptanceRunner(settings, progress=lambda message: None).run(
        [source, tmp_path / "absent.xlsx"],
        request="گزارش",
        output=tmp_path / "report.docx",
    )

    assert report.status == "partial"
    json_path, markdown_path = report.write(tmp_path / "audit")
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    assert any(item["exists"] is False for item in payload["inputs"])
    assert "جزئی" in markdown_path.read_text(encoding="utf-8")


def test_folder_docx_uses_template_for_layout_not_figures(settings, tmp_path):
    from docx import Document as DocxDocument

    folder = tmp_path / "sources"
    _violations(folder / "الف.xlsx")
    template_path = tmp_path / "template.docx"
    template = DocxDocument()
    template.add_heading("قالب سازمانی", level=1)
    template.add_paragraph("متن ثابت قالب")
    template.add_paragraph("{{content}}")
    template.save(template_path)

    result = FolderWorkflow(settings, progress=lambda message: None).run(
        folder,
        request="گزارش مدیریتی تخلفات",
        use_ai=False,
        output=tmp_path / "report.docx",
        template=template_path,
    )

    assert Path(result.output_path).suffix == ".docx"
    assert result.report.mode == "template"
    assert result.report.validation.valid is True, result.report.validation.errors
    document = Document(result.output_path)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    table_text = "\n".join(
        cell.text for table in document.tables for row in table.rows for cell in row.cells
    )
    assert "قالب سازمانی" in text
    assert "متن ثابت قالب" in text
    assert "{{content}}" not in text
    assert "تخلفات منجر به تصادف" in table_text


def test_acceptance_cli_returns_zero_only_for_accepted(settings, tmp_path):
    from local_office_ai.ui.cli import main

    source = _violations(tmp_path / "violations.xlsx")
    code = main(
        [
            "--root",
            str(settings.project_root),
            "--acceptance",
            "--acceptance-input",
            str(source),
            "--request",
            "گزارش مدیریتی تخلفات",
            "--output",
            str(tmp_path / "cli.docx"),
            "--acceptance-dir",
            str(tmp_path / "audit"),
            "--no-ai",
            "--quiet",
        ]
    )
    assert code == 0
    assert (tmp_path / "cli.docx").is_file()
    assert (tmp_path / "audit" / "acceptance_report.json").is_file()
