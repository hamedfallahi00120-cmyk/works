"""Prompt 04A: output-format routing and DOCX template architecture.

All fixtures are real files: openpyxl workbooks, python-docx templates and the
actual Excel/DOCX writers.  Nothing is mocked.
"""

from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from local_office_ai.ai.analyzer import AIAnalyzer
from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.exceptions import ReportFormatError, ReportTemplateError
from local_office_ai.excel.report_writer import ExcelReportWriter, ReportContext
from local_office_ai.reporting import (
    DocxTemplateRenderer,
    OutputFormat,
    ReportContent,
    ReportGenerator,
    ReportSpecification,
    ReportTemplate,
    ReportValue,
    detect_requested_formats,
    ensure_extension_matches,
    ensure_xlsx_delivery,
    inspect_docx_template,
    resolve_output_intent,
)
from local_office_ai.ui.cli import main

HISTORICAL_FIGURE = "۱۲٬۵۰۰٬۰۰۰"


def _rtl(element, tag: str) -> None:
    element.append(OxmlElement(tag))


@pytest.fixture()
def docx_template(tmp_path: Path) -> Path:
    """A real RTL Word template with styles, numbering, table, header/footer."""

    path = tmp_path / "قالب_گزارش.docx"
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)
    _rtl(section._sectPr, "w:rtlGutter")
    _rtl(section._sectPr, "w:bidi")

    style = document.styles.add_style("متن گزارش", WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = "B Nazanin"
    style.font.size = Pt(13)
    style.font.bold = True
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    document.add_heading("عنوان اصلی {{report.title}}", level=1)
    paragraph = document.add_paragraph(
        "متن {{report.title}} و عدد تاریخی " + HISTORICAL_FIGURE,
        style="متن گزارش",
    )
    _rtl(paragraph._p.get_or_add_pPr(), "w:bidi")
    _rtl(paragraph.runs[0]._r.get_or_add_rPr(), "w:rtl")

    split = document.add_paragraph()
    first = split.add_run("{{report.")
    first.bold = True
    second = split.add_run("source_file}}")
    second.italic = True

    document.add_paragraph("مورد شماره‌دار", style="List Number")

    table = document.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "شرح"
    table.cell(0, 1).text = "مقدار"
    table.cell(1, 0).text = "جمع"
    table.cell(1, 1).text = "{{metric.principal_sum}}"

    section.header.paragraphs[0].text = "سرصفحه {{report.title}}"
    section.footer.paragraphs[0].text = "تاریخ {{report.generated_at}}"
    document.save(path)
    return path


@pytest.fixture()
def template_content() -> ReportContent:
    return ReportContent(
        title="گزارش آزمون",
        subtitle="زیرعنوان",
        values={
            "report.title": ReportValue("report.title", "عنوان", "گزارش آزمون"),
            "report.source_file": ReportValue("report.source_file", "فایل", "costs.xlsx"),
            "report.generated_at": ReportValue("report.generated_at", "تاریخ", "1404/01/01"),
            "metric.principal_sum": ReportValue("metric.principal_sum", "جمع اصل", "9,999"),
        },
    )


# --------------------------------------------------------------------------- #
# 1. Output-format selection
# --------------------------------------------------------------------------- #
def test_output_format_tokens_and_suffixes_are_selected():
    assert OutputFormat.from_token("word") is OutputFormat.DOCX
    assert OutputFormat.from_token("DOCX") is OutputFormat.DOCX
    assert OutputFormat.from_token("ورد") is OutputFormat.DOCX
    assert OutputFormat.from_token("excel") is OutputFormat.XLSX
    assert OutputFormat.from_token("xlsx") is OutputFormat.XLSX
    assert OutputFormat.from_token("pdf") is OutputFormat.PDF
    assert OutputFormat.from_suffix(".XLSX") is OutputFormat.XLSX
    assert OutputFormat.from_suffix(".docx") is OutputFormat.DOCX
    assert OutputFormat.from_suffix(".pdf") is OutputFormat.PDF
    assert OutputFormat.from_suffix("") is None


def test_output_intent_uses_explicit_suffix_request_or_default(tmp_path):
    assert resolve_output_intent(explicit_format="xlsx").output_format is OutputFormat.XLSX
    assert resolve_output_intent(output_path=tmp_path / "r.docx").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="گزارش Word بده").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="خروجی PDF لازم دارم").output_format is OutputFormat.PDF
    assert resolve_output_intent(request_text="گزارش مدیریتی بده").output_format is OutputFormat.XLSX
    assert resolve_output_intent(request_text="گزارش به صورت Word").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="خروجی به فرمت docx").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="گزارش را به Word تبدیل کن").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="فایل docx را بده").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="گزارش Word برای اکسل").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="گزارش برای Word").output_format is OutputFormat.DOCX
    assert resolve_output_intent(request_text="گزارش برای اکسل").output_format is OutputFormat.XLSX
    assert detect_requested_formats("این فایل اکسل را تحلیل کن") == []
    assert detect_requested_formats("این فایل report.docx را تحلیل کن") == []
    assert detect_requested_formats("فایل docx را بررسی کن") == []
    assert detect_requested_formats("docx را تحلیل کن") == []
    assert detect_requested_formats("report.xlsx") == []
    assert detect_requested_formats("docx") == [OutputFormat.DOCX]
    assert detect_requested_formats("فایل docx بساز") == [OutputFormat.DOCX]


def test_specification_appends_the_selected_suffix(tmp_path, docx_template):
    spec = ReportSpecification(
        output_format=OutputFormat.DOCX,
        output_path=tmp_path / "report",
        template_path=docx_template,
    )
    assert spec.output_path.suffix == ".docx"
    spec.validate()

    plan = ReportGenerator(_settings_for(tmp_path)).plan(spec)
    assert plan.output_format is OutputFormat.DOCX
    assert plan.writer == "DocxReportWriter"
    assert plan.requires_template is False


# --------------------------------------------------------------------------- #
# 2. File-extension validation and mismatched formats
# --------------------------------------------------------------------------- #
def test_unknown_output_suffix_is_rejected():
    with pytest.raises(ReportFormatError, match="پسوند خروجی"):
        OutputFormat.from_suffix(".xlsm")
    with pytest.raises(ReportFormatError, match="پسوند خروجی"):
        resolve_output_intent(output_path="report.txt")


def test_mismatched_extension_and_format_are_rejected(tmp_path, docx_template):
    with pytest.raises(ReportFormatError, match="سازگار نیست"):
        ReportSpecification(OutputFormat.XLSX, tmp_path / "report.docx").validate()
    with pytest.raises(ReportFormatError, match="سازگار نیست"):
        ReportSpecification(
            OutputFormat.DOCX,
            tmp_path / "report.xlsx",
            template_path=docx_template,
        ).validate()
    with pytest.raises(ReportFormatError, match="قالب Word فقط"):
        ReportSpecification(
            OutputFormat.XLSX,
            tmp_path / "report.xlsx",
            template_path=docx_template,
        ).validate()


def test_conflicting_format_signals_are_rejected(tmp_path):
    with pytest.raises(ReportFormatError, match="متناقض"):
        resolve_output_intent(explicit_format="xlsx", output_path=tmp_path / "report.docx")
    with pytest.raises(ReportFormatError, match="متناقض"):
        resolve_output_intent(explicit_format="docx", request_text="گزارش اکسل بده")
    with pytest.raises(ReportFormatError, match="متناقض"):
        resolve_output_intent(request_text="گزارش Word و خروجی PDF بده")


def test_extension_helper_and_excel_only_workflows(tmp_path):
    assert ensure_extension_matches(tmp_path / "report", OutputFormat.XLSX).suffix == ".xlsx"
    with pytest.raises(ReportFormatError, match="سازگار نیست"):
        ensure_extension_matches(tmp_path / "report.docx", OutputFormat.XLSX)
    with pytest.raises(ReportFormatError, match="فقط به صورت کارپوشه Excel"):
        ensure_xlsx_delivery(
            output_path=tmp_path / "folder.docx",
            default_output_path=tmp_path / "folder.xlsx",
        )
    with pytest.raises(ReportFormatError, match="فقط به صورت کارپوشه Excel"):
        ensure_xlsx_delivery(
            output_path=None,
            default_output_path=tmp_path / "folder.xlsx",
            request_text="گزارش Word بده",
        )


def test_pdf_generation_is_explicitly_unsupported(tmp_path):
    spec = ReportSpecification(OutputFormat.PDF, tmp_path / "report.pdf")
    with pytest.raises(ReportFormatError, match="تولید سند PDF"):
        spec.validate()
    with pytest.raises(ReportFormatError, match="تولید سند PDF"):
        ReportGenerator(_settings_for(tmp_path)).plan(spec)


# --------------------------------------------------------------------------- #
# 3. DOCX template inspection
# --------------------------------------------------------------------------- #
def test_docx_template_inspection_records_layout_styles_tables_and_rtl(docx_template):
    template = inspect_docx_template(docx_template)

    assert template.valid is True
    assert template.is_data_source is False
    assert template.section_count == 1
    layout = template.page_layouts[0]
    assert layout.top_margin_twips == Cm(2.5).twips
    assert layout.left_margin_twips == Cm(3).twips
    assert layout.rtl_gutter is True
    assert layout.bidi is True

    assert any(style.name == "متن گزارش" for style in template.styles)
    custom = next(style for style in template.styles if style.name == "متن گزارش")
    assert custom.font["name"] == "B Nazanin"
    assert custom.font["size_pt"] == 13.0
    assert custom.font["bold"] is True
    assert custom.alignment == "RIGHT"

    assert any(item.text.startswith("عنوان اصلی") and item.heading_level == 1 for item in template.headings)
    assert any(table.style == "Table Grid" and table.rows == 2 for table in template.tables)
    assert any(item.format == "decimal" for item in template.numbering)
    assert any("report.title" in header.placeholder_keys for header in template.headers)
    assert any("report.generated_at" in footer.placeholder_keys for footer in template.footers)
    assert template.rtl["section_bidi"] is True
    assert template.rtl["section_rtl_gutter"] is True
    assert template.rtl["bidi_paragraphs"] >= 1
    assert template.rtl["rtl_runs"] >= 1
    assert template.rtl["numbering_rtl"] == 0


def test_placeholder_detection_spans_body_table_header_and_footer(docx_template):
    template = inspect_docx_template(docx_template)
    assert ReportTemplate.inspect(docx_template).placeholder_keys == template.placeholder_keys

    assert set(template.placeholder_keys) == {
        "report.title",
        "report.source_file",
        "report.generated_at",
        "metric.principal_sum",
    }
    locations = " ".join(item.location for item in template.placeholders)
    assert "body/paragraph" in locations
    assert "table[" in locations
    assert "header[" in locations
    assert "footer[" in locations
    assert len(template.placeholder_locations("report.title")) == 3


# --------------------------------------------------------------------------- #
# 4. Preservation of styles, 5. placeholders, template/data separation
# --------------------------------------------------------------------------- #
def test_template_rendering_preserves_styles_layout_and_rtl(docx_template, template_content, tmp_path):
    template = inspect_docx_template(docx_template)
    output = tmp_path / "rendered.docx"
    result = DocxTemplateRenderer().render(template, template_content, output)

    assert result.output_path == output
    assert result.replaced == 6
    assert result.unresolved_placeholders == ()

    rendered = Document(str(output))
    section = rendered.sections[0]
    assert section.top_margin.twips == Cm(2.5).twips
    assert section.left_margin.twips == Cm(3).twips
    assert section._sectPr.find(qn("w:rtlGutter")) is not None
    assert section._sectPr.find(qn("w:bidi")) is not None

    heading = rendered.paragraphs[0]
    assert heading.style.name == "Heading 1"
    assert heading.text == "عنوان اصلی گزارش آزمون"

    styled = rendered.paragraphs[1]
    assert styled.style.name == "متن گزارش"
    assert styled.text.startswith("متن گزارش آزمون")
    assert styled._p.pPr.find(qn("w:bidi")) is not None

    split = rendered.paragraphs[2]
    assert split.text == "costs.xlsx"
    assert split.runs[0].bold is True
    assert split.runs[1].italic is True

    assert rendered.paragraphs[3].style.name == "List Number"
    table = rendered.tables[0]
    assert table.style.name == "Table Grid"
    assert table.cell(1, 1).text == "9,999"
    assert rendered.sections[0].header.paragraphs[0].text == "سرصفحه گزارش آزمون"
    assert rendered.sections[0].footer.paragraphs[0].text == "تاریخ 1404/01/01"


def test_renderer_keeps_a_template_that_matches_the_old_temp_name(tmp_path):
    template_path = tmp_path / "report.tmp.docx"
    document = Document()
    document.add_paragraph("{{report.title}}")
    document.save(template_path)
    before = template_path.read_bytes()

    content = ReportContent(
        title="گزارش",
        values={"report.title": ReportValue("report.title", "عنوان", "گزارش")},
    )
    output = tmp_path / "report.docx"
    DocxTemplateRenderer().render(ReportTemplate.inspect(template_path), content, output)

    assert template_path.read_bytes() == before
    assert output.exists()
    assert "گزارش" in Document(str(output)).paragraphs[0].text


def test_unknown_placeholders_are_left_untouched_and_warned(docx_template, tmp_path):
    template = inspect_docx_template(docx_template)
    content = ReportContent(
        title="گزارش",
        values={
            "report.title": ReportValue("report.title", "عنوان", "گزارش"),
            "report.source_file": ReportValue("report.source_file", "فایل", "x.xlsx"),
            "metric.principal_sum": ReportValue("metric.principal_sum", "جمع", "10"),
        },
    )
    output = tmp_path / "partial.docx"
    result = DocxTemplateRenderer().render(template, content, output)

    assert result.unresolved_placeholders == ("report.generated_at",)
    assert any("مقدار معتبر پیدا نشد" in warning for warning in result.warnings)
    rendered = Document(str(output))
    assert "{{report.generated_at}}" in rendered.sections[0].footer.paragraphs[0].text


def test_template_figures_never_become_report_data(settings, make_workbook, docx_template):
    source = make_workbook(
        "costs.xlsx",
        {
            "هزینه‌ها": [
                ["جدول هزینه‌های جاری"],
                [],
                ["ردیف", "شرح", "مبلغ", "دسته"],
                [1, "اجاره", 1250000, "اجاره"],
                [2, "برق", 1250000, "انرژی"],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(source)
    content = ReportContent.from_analysis(analysis, request="گزارش Word بده")
    template = inspect_docx_template(docx_template)

    assert content.source_kind == "python_analysis"
    assert content.validated is True
    assert content.validate() == []
    assert content.placeholder_values()["report.source_file"] == "costs.xlsx"
    assert content.placeholder_values()["stat.C.total"] == "2,500,000"
    assert content.values["stat.C.total"].source == "statistics"
    assert HISTORICAL_FIGURE not in content.placeholder_values().values()
    assert template.is_data_source is False
    assert template.has_static_figures is True
    assert any(HISTORICAL_FIGURE in item.text for item in template.static_numeric_literals)


def test_report_content_preserves_small_decimal_values():
    analysis = SimpleNamespace(
        meta=SimpleNamespace(path="ratio.xlsx"),
        primary_table=None,
        integration=None,
        calculations=[
            SimpleNamespace(
                name="ratio",
                value=Decimal("0.004"),
                unit="",
                sources=[],
            )
        ],
        generic=None,
        request_answers=[],
        anomalies=[],
        warnings=[],
    )

    content = ReportContent.from_analysis(analysis)

    assert content.placeholder_values()["metric.ratio"] == "0.004"


def test_report_content_keeps_zero_answer_available():
    analysis = SimpleNamespace(
        meta=SimpleNamespace(path="answer.xlsx"),
        primary_table=None,
        integration=None,
        calculations=[],
        generic=None,
        request_answers=[
            SimpleNamespace(title="اختلاف", value=0, text="هیچ اختلافی وجود ندارد")
        ],
        anomalies=[],
        warnings=[],
    )

    content = ReportContent.from_analysis(analysis)

    value = content.values["answer.1.value"]
    assert value.value == "0"
    assert value.available is True
    assert content.validate() == []


def test_missing_and_unsupported_templates_are_rejected(tmp_path):
    missing = tmp_path / "missing.docx"
    with pytest.raises(ReportTemplateError, match="پیدا نشد"):
        inspect_docx_template(missing)

    text_file = tmp_path / "template.txt"
    text_file.write_text("not a docx", encoding="utf-8")
    with pytest.raises(ReportTemplateError, match="پشتیبانی نمی‌شود"):
        inspect_docx_template(text_file)

    corrupt = tmp_path / "corrupt.docx"
    corrupt.write_bytes(b"not a zip")
    with pytest.raises(ReportTemplateError, match="ZIP"):
        inspect_docx_template(corrupt)

    empty = tmp_path / "empty.docx"
    empty.write_bytes(b"")
    with pytest.raises(ReportTemplateError, match="خالی"):
        inspect_docx_template(empty)

    # A Word report no longer needs a template; a template that is requested
    # but missing must still fail before any output is written.
    ReportSpecification(OutputFormat.DOCX, tmp_path / "report.docx").validate()
    with pytest.raises(ReportTemplateError, match="پیدا نشد"):
        ReportSpecification(
            OutputFormat.DOCX,
            tmp_path / "report.docx",
            template_path=missing,
        ).validate()


def test_excel_writer_rejects_a_word_suffix(settings, make_workbook, tmp_path):
    source = make_workbook(
        "costs.xlsx",
        {
            "هزینه‌ها": [
                ["جدول هزینه‌های جاری"],
                [],
                ["ردیف", "شرح", "مبلغ"],
                [1, "اجاره", 1250000],
            ]
        },
    )
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(source)
    ai = AIAnalyzer(settings).analyze(analysis, use_ai=False)
    context = ReportContext(
        analysis=analysis,
        ai=ai,
        performance=None,
        output_path=tmp_path / "report.docx",
    )
    with pytest.raises(ReportFormatError, match="سازگار نیست"):
        ExcelReportWriter(settings).build(context)
    assert not (tmp_path / "report.docx").exists()


# --------------------------------------------------------------------------- #
# CLI integration: inspect template, Word output with template, PDF rejection
# --------------------------------------------------------------------------- #
def test_cli_inspects_a_template(project_root, docx_template, capsys):
    code = main(["--root", str(project_root), "--inspect-template", str(docx_template)])
    captured = capsys.readouterr()
    assert code == 0
    assert "جایگاه‌ها" in captured.out
    assert "report.title" in captured.out
    assert "اعداد ثابت قالب" in captured.out


def test_cli_generates_a_word_report_with_a_template(project_root, make_workbook, docx_template):
    source = make_workbook(
        "costs.xlsx",
        {
            "هزینه‌ها": [
                ["جدول هزینه‌های جاری"],
                [],
                ["ردیف", "شرح", "مبلغ", "دسته"],
                [1, "اجاره", 1250000, "اجاره"],
                [2, "برق", 1250000, "انرژی"],
            ]
        },
    )
    output = project_root / "word-report.docx"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(source),
            "--no-ai",
            "--yes",
            "--quiet",
            "--output",
            str(output),
            "--template",
            str(docx_template),
        ]
    )
    assert code == 0
    assert output.exists()
    document = Document(str(output))
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert "costs.xlsx" in text
    assert document.paragraphs[0].style.name == "Heading 1"


def test_configured_word_template_does_not_break_excel_output(project_root, make_workbook, docx_template):
    (project_root / "config" / "settings.json").write_text(
        json.dumps({"docx_template": str(docx_template)}),
        encoding="utf-8",
    )
    source = make_workbook(
        "costs.xlsx",
        {
            "هزینه‌ها": [
                ["جدول هزینه‌های جاری"],
                [],
                ["ردیف", "شرح", "مبلغ"],
                [1, "اجاره", 1250000],
            ]
        },
    )
    output = project_root / "excel-report.xlsx"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(source),
            "--no-ai",
            "--yes",
            "--quiet",
            "--format",
            "xlsx",
            "--output",
            str(output),
        ]
    )
    assert code == 0
    assert output.exists()


def test_cli_rejects_word_template_for_folder(project_root, tmp_path, capsys):
    code = main(
        [
            "--root",
            str(project_root),
            "--folder",
            str(tmp_path),
            "--template",
            str(tmp_path / "template.docx"),
        ]
    )
    captured = capsys.readouterr()
    assert code == 2
    assert "قالب Word" in captured.err


def test_cli_rejects_pdf_and_mismatched_output(project_root, make_workbook):
    source = make_workbook(
        "costs.xlsx",
        {
            "هزینه‌ها": [
                ["جدول هزینه‌های جاری"],
                [],
                ["ردیف", "شرح", "مبلغ"],
                [1, "اجاره", 1250000],
            ]
        },
    )
    pdf = project_root / "report.pdf"
    assert main(
        ["--root", str(project_root), "--file", str(source), "--no-ai", "--yes", "--quiet", "--output", str(pdf)]
    ) == 2
    assert not pdf.exists()

    mismatched = project_root / "report.docx"
    assert main(
        [
            "--root",
            str(project_root),
            "--file",
            str(source),
            "--no-ai",
            "--yes",
            "--quiet",
            "--format",
            "xlsx",
            "--output",
            str(mismatched),
        ]
    ) == 2
    assert not mismatched.exists()


def _settings_for(tmp_path: Path):
    from local_office_ai.core.settings import load_settings

    root = tmp_path / "project"
    (root / "config").mkdir(parents=True, exist_ok=True)
    (root / "data").mkdir(parents=True, exist_ok=True)
    (root / "logs").mkdir(parents=True, exist_ok=True)
    return load_settings(project_root=root)
