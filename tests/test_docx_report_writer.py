"""Prompt 04B: professional DOCX report writing and Persian RTL rendering.

Every fixture is a real file: openpyxl workbooks, python-docx templates and the
actual writer/validator.  Nothing is mocked.
"""

from __future__ import annotations

import json
import zipfile
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.exceptions import ReportFormatError, ReportTemplateError
from local_office_ai.reporting import (
    DocxReportWriter,
    OutputFormat,
    ReportContent,
    ReportGenerator,
    ReportSpecification,
    ReportTable,
    ReportValue,
    inspect_docx_template,
    validate_docx_output,
)
from local_office_ai.ui.cli import main

HISTORICAL_FIGURE = "۱۲٬۵۰۰٬۰۰۰"


# --------------------------------------------------------------------------- #
# Real content and real template fixtures
# --------------------------------------------------------------------------- #
@pytest.fixture()
def report_source(make_workbook) -> Path:
    return make_workbook(
        "report.xlsx",
        {
            "هزینه‌ها": [
                ["گزارش هزینه‌های اداری"],
                [],
                ["ردیف", "شرح", "مبلغ", "تاریخ", "دسته"],
                [1, "اجاره دفتر", 1250000, "1404/01/15", "اجاره"],
                [2, "برق و آب", 1250000, "1404/02/03", "انرژی"],
                ["جمع", None, "=SUM(C4:C5)", None, None],
            ]
        },
    )


@pytest.fixture()
def report_content(settings, report_source):
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(report_source)
    content = ReportContent.from_analysis(analysis, request="گزارش Word بده")
    return analysis, content


@pytest.fixture()
def rtl_template(tmp_path: Path) -> Path:
    path = tmp_path / "قالب_گزارش.docx"
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)
    section._sectPr.append(OxmlElement("w:bidi"))

    style = document.styles.add_style("متن گزارش", WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = "B Nazanin"
    style.font.size = Pt(13)

    document.add_heading("عنوان اصلی {{report.title}}", level=1)
    document.add_paragraph(
        "متن {{report.source_file}} و عدد تاریخی " + HISTORICAL_FIGURE,
        style="متن گزارش",
    )
    document.add_paragraph("{{content}}")

    table = document.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "شرح"
    table.cell(0, 1).text = "مقدار"
    table.cell(1, 0).text = "جمع"
    table.cell(1, 1).text = "{{stat.C.total}}"

    section.header.paragraphs[0].text = "سرصفحه {{report.title}}"
    section.footer.paragraphs[0].text = "تاریخ {{report.generated_at}}"
    document.save(path)
    return path


def _build(content, output: Path, **kwargs):
    return DocxReportWriter().build(content, output, **kwargs)


def _all_text(document) -> str:
    texts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    for section in document.sections:
        for container in (section.header, section.footer):
            texts.extend(paragraph.text for paragraph in container.paragraphs)
    return "\n".join(texts)


# --------------------------------------------------------------------------- #
# 1. Valid DOCX generation
# --------------------------------------------------------------------------- #
def test_standalone_word_report_is_a_real_docx(settings, report_content, tmp_path):
    analysis, content = report_content
    output = tmp_path / "standalone.docx"

    result = _build(content, output)

    assert result.mode == "standalone"
    assert result.validation is not None and result.validation.valid is True
    assert output.exists() and output.stat().st_size > 5000
    assert zipfile.is_zipfile(output)
    with zipfile.ZipFile(output) as archive:
        names = set(archive.namelist())
    assert {"[Content_Types].xml", "_rels/.rels", "word/document.xml"} <= names

    document = Document(str(output))
    assert document.paragraphs
    assert len(document.tables) >= 1
    assert document.paragraphs[0].text == content.title
    assert any("خلاصه اجرایی" in paragraph.text for paragraph in document.paragraphs)


# --------------------------------------------------------------------------- #
# 2. Persian RTL paragraphs
# --------------------------------------------------------------------------- #
def test_persian_paragraphs_carry_real_rtl_properties(report_content, tmp_path):
    _, content = report_content
    output = tmp_path / "rtl.docx"
    result = _build(content, output)

    section = Document(str(output)).sections[0]
    assert section._sectPr.find(qn("w:bidi")) is not None
    assert section._sectPr.find(qn("w:rtlGutter")) is not None

    document = Document(str(output))
    persian_paragraphs = [
        paragraph
        for paragraph in document.paragraphs
        if paragraph.text.strip() and any("\u0600" <= char <= "\u06FF" for char in paragraph.text)
    ]
    assert persian_paragraphs
    for paragraph in persian_paragraphs:
        p_pr = paragraph._p.pPr
        assert p_pr is not None and p_pr.find(qn("w:bidi")) is not None
        assert paragraph.runs
        for run in paragraph.runs:
            r_pr = run._r.rPr
            assert r_pr is not None and r_pr.find(qn("w:rtl")) is not None
            r_fonts = r_pr.find(qn("w:rFonts"))
            assert r_fonts is not None and r_fonts.get(qn("w:cs"))
            assert r_pr.find(qn("w:szCs")) is not None
            assert r_pr.find(qn("w:lang")).get(qn("w:bidi")) == "fa-IR"

    assert result.validation.sections_with_rtl >= 1
    assert result.validation.rtl_runs > 0


# --------------------------------------------------------------------------- #
# 3. RTL tables
# --------------------------------------------------------------------------- #
def test_tables_are_right_to_left_with_repeated_headers(report_content, tmp_path):
    _, content = report_content
    output = tmp_path / "tables.docx"
    result = _build(content, output)

    document = Document(str(output))
    assert document.tables
    for table in document.tables:
        assert table._tbl.tblPr.find(qn("w:bidiVisual")) is not None
        header_tr_pr = table.rows[0]._tr.trPr
        assert header_tr_pr is not None
        assert header_tr_pr.find(qn("w:tblHeader")) is not None
        for cell in table.rows[0].cells:
            tc_pr = cell._tc.tcPr
            assert tc_pr is not None and tc_pr.find(qn("w:shd")) is not None
            for paragraph in cell.paragraphs:
                assert paragraph._p.pPr.find(qn("w:bidi")) is not None
                for run in paragraph.runs:
                    assert run._r.rPr.find(qn("w:rtl")) is not None
                    assert run._r.rPr.find(qn("w:lang")).get(qn("w:bidi")) == "fa-IR"
    assert result.validation.rtl_tables == len(document.tables)


# --------------------------------------------------------------------------- #
# 4. Template preservation
# --------------------------------------------------------------------------- #
def test_template_layout_styles_and_content_are_preserved(report_content, rtl_template, tmp_path):
    _, content = report_content
    before = rtl_template.read_bytes()
    output = tmp_path / "from-template.docx"

    result = _build(content, output, template_path=rtl_template)

    assert result.mode == "template"
    assert result.validation is not None and result.validation.valid is True
    assert rtl_template.read_bytes() == before

    document = Document(str(output))
    section = document.sections[0]
    assert section.top_margin.twips == Cm(2.5).twips
    assert section.left_margin.twips == Cm(3).twips
    assert section._sectPr.find(qn("w:bidi")) is not None

    assert any(style.name == "متن گزارش" for style in document.styles)
    styled = next(style for style in document.styles if style.name == "متن گزارش")
    assert styled.font.name == "B Nazanin"
    assert styled.font.size == Pt(13)

    heading = document.paragraphs[0]
    assert heading.style.name == "Heading 1"
    assert heading.text == f"عنوان اصلی {content.title}"

    table = next(
        item
        for item in document.tables
        if item.style.name == "Table Grid" and item.cell(1, 1).text == "2,500,000"
    )
    assert table.cell(0, 1).text == "مقدار"
    assert "2,500,000" in _all_text(document)


def test_template_run_formatting_is_not_rewritten(report_content, tmp_path):
    _, content = report_content
    template = tmp_path / "format-preserving.docx"
    document = Document()
    paragraph = document.add_paragraph()
    run = paragraph.add_run("متن قالب با قلم اختصاصی")
    run.font.name = "Vazirmatn"
    run.font.size = Pt(11)
    run.bold = False
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "سرستون"
    table.cell(0, 1).text = "مقدار"
    header_run = table.cell(0, 0).paragraphs[0].runs[0]
    header_run.font.name = "Tahoma"
    header_run.bold = False
    document.add_paragraph("{{content}}")
    document.save(template)

    output = tmp_path / "format-preserving-out.docx"
    result = _build(content, output, template_path=template)

    assert result.validation.valid is True
    reopened = Document(str(output))
    styled = next(
        item for item in reopened.paragraphs if item.text == "متن قالب با قلم اختصاصی"
    )
    styled_run = styled.runs[0]
    assert styled_run.font.name == "Vazirmatn"
    assert styled_run.font.size == Pt(11)
    assert styled_run.bold is False

    table = reopened.tables[0]
    header_run = table.cell(0, 0).paragraphs[0].runs[0]
    assert header_run.font.name == "Tahoma"
    assert header_run.bold is False


def test_formatted_static_figure_is_removed_without_collapsing_runs(
    report_content, tmp_path
):
    _, content = report_content
    template = tmp_path / "formatted-figure.docx"
    document = Document()
    paragraph = document.add_paragraph()
    lead = paragraph.add_run("مبلغ تاریخی ")
    lead.bold = True
    figure = paragraph.add_run(HISTORICAL_FIGURE)
    figure.italic = True
    tail = paragraph.add_run(" تومان بود")
    tail.underline = True
    document.add_paragraph("{{content}}")
    document.save(template)

    output = tmp_path / "formatted-figure-out.docx"
    result = _build(content, output, template_path=template)

    assert HISTORICAL_FIGURE in result.removed_static_figures
    reopened = Document(str(output))
    paragraph = next(
        item for item in reopened.paragraphs if "مبلغ تاریخی" in item.text
    )
    assert HISTORICAL_FIGURE not in paragraph.text
    assert paragraph.runs[0].bold is True
    assert paragraph.runs[1].italic is True
    assert paragraph.runs[2].underline is True


# --------------------------------------------------------------------------- #
# 5. Placeholder replacement
# --------------------------------------------------------------------------- #
def test_placeholders_are_replaced_and_historical_figures_removed(
    report_content, rtl_template, tmp_path
):
    _, content = report_content
    output = tmp_path / "placeholders.docx"

    result = _build(content, output, template_path=rtl_template)

    assert "report.title" in result.placeholders_resolved
    assert "report.source_file" in result.placeholders_resolved
    assert "stat.C.total" in result.placeholders_resolved
    assert "content" not in result.placeholders_resolved
    assert result.unresolved_placeholders == ()
    assert HISTORICAL_FIGURE in result.removed_static_figures

    text = _all_text(Document(str(output)))
    assert "{{" not in text and "}}" not in text
    assert HISTORICAL_FIGURE not in text
    assert content.title in text
    assert "report.xlsx" in text
    assert "2,500,000" in text
    assert "صفحه" in text


def test_historical_template_figure_can_be_kept_explicitly(report_content, rtl_template, tmp_path):
    _, content = report_content
    output = tmp_path / "reference-figures.docx"

    result = _build(
        content,
        output,
        template_path=rtl_template,
        strip_static_figures=False,
    )

    assert result.removed_static_figures == ()
    assert HISTORICAL_FIGURE in _all_text(Document(str(output)))
    assert result.validation.valid is True


def test_template_without_anchor_receives_the_body_at_the_end(report_content, tmp_path):
    _, content = report_content
    template = tmp_path / "plain.docx"
    document = Document()
    document.add_paragraph("سند پایه {{report.title}}")
    document.save(template)

    output = tmp_path / "appended.docx"
    result = _build(content, output, template_path=template)

    assert result.mode == "template"
    assert result.validation.valid is True
    document = Document(str(output))
    assert "سند پایه" in document.paragraphs[0].text
    assert any("خلاصه اجرایی" in paragraph.text for paragraph in document.paragraphs)


def test_anchor_with_surrounding_text_keeps_that_text(report_content, tmp_path):
    _, content = report_content
    template = tmp_path / "inline-anchor.docx"
    document = Document()
    document.add_paragraph("بخش تحلیلی در ادامه می‌آید: {{content}}")
    document.save(template)

    output = tmp_path / "inline-anchor-out.docx"
    result = _build(content, output, template_path=template)

    assert result.validation.valid is True
    document = Document(str(output))
    assert document.paragraphs[0].text.startswith("بخش تحلیلی در ادامه می‌آید:")
    assert "{{" not in document.paragraphs[0].text
    assert any("خلاصه اجرایی" in paragraph.text for paragraph in document.paragraphs)
    assert any("شناسنامه گزارش" in paragraph.text for paragraph in document.paragraphs)


def test_english_only_template_still_receives_rtl_sections(report_content, tmp_path):
    _, content = report_content
    template = tmp_path / "english-template.docx"
    document = Document()
    document.add_paragraph("Analysis body")
    document.add_paragraph("{{content}}")
    document.save(template)

    output = tmp_path / "english-template-out.docx"
    result = _build(content, output, template_path=template)

    assert result.validation.valid is True
    assert result.validation.sections_with_rtl >= 1
    assert result.validation.rtl_runs >= 1


def test_template_sections_are_not_duplicated(report_content, tmp_path):
    _, content = report_content
    template = tmp_path / "existing-sections.docx"
    document = Document()
    document.add_heading("خلاصه اجرایی", level=1)
    document.add_paragraph("خلاصه‌ای که قالب از قبل دارد.")
    document.add_paragraph("{{content}}")
    document.save(template)

    output = tmp_path / "no-duplicates.docx"
    result = _build(content, output, template_path=template)

    assert result.validation.valid is True
    document = Document(str(output))
    headings = [
        paragraph.text
        for paragraph in document.paragraphs
        if paragraph.style.name == "Heading 1"
    ]
    assert headings.count("خلاصه اجرایی") == 1
    assert any(
        content.executive_summary[0] in paragraph.text
        for paragraph in document.paragraphs
    )
    assert result.tables_written == len(content.tables)


def test_matching_template_table_is_filled_instead_of_duplicated(
    report_content, tmp_path
):
    _, content = report_content
    template = tmp_path / "existing-table.docx"
    document = Document()
    document.add_heading("شناسنامه گزارش", level=1)
    table = document.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "شرح"
    table.cell(0, 1).text = "مقدار"
    table.cell(1, 0).text = "مقدار قدیمی"
    table.cell(1, 1).text = "۱۲۳"
    document.add_paragraph("{{content}}")
    document.save(template)

    output = tmp_path / "existing-table-out.docx"
    result = _build(content, output, template_path=template)

    assert result.validation.valid is True
    document = Document(str(output))
    matching = [
        item
        for item in document.tables
        if [cell.text for cell in item.rows[0].cells] == ["شرح", "مقدار"]
    ]
    assert len(matching) == 1
    assert "مقدار قدیمی" not in _all_text(document)
    assert any(cell.text == "فایل منبع" for cell in matching[0].columns[0].cells)
    assert result.tables_written == len(content.tables)
    assert any("جلوگیری از تکرار" in warning for warning in result.warnings)


# --------------------------------------------------------------------------- #
# 6. Numerical consistency
# --------------------------------------------------------------------------- #
def test_every_number_in_the_document_is_approved_content(report_content, tmp_path):
    _, content = report_content
    output = tmp_path / "numbers.docx"
    result = _build(content, output)

    validation = validate_docx_output(output, content)
    assert validation.valid is True
    assert validation.unsupported_numbers == ()
    assert validation.numeric_tokens > 0

    document = Document(str(output))
    assert "2,500,000" in _all_text(document)

    document.add_paragraph("مبلغ نامعتبر 987654321 در متن")
    tampered = tmp_path / "tampered.docx"
    document.save(tampered)

    tampered_validation = validate_docx_output(tampered, content)
    assert tampered_validation.valid is False
    assert "987654321" in tampered_validation.unsupported_numbers
    assert any("بدون منبع" in error for error in tampered_validation.errors)


def test_model_narrative_is_dropped_when_it_quotes_an_unapproved_number():
    analysis = SimpleNamespace(
        meta=SimpleNamespace(path="narrative.xlsx"),
        primary_table=None,
        integration=None,
        calculations=[
            SimpleNamespace(
                name="total",
                label="جمع",
                value=Decimal("2500000"),
                unit="",
                sources=[],
            )
        ],
        generic=None,
        request_answers=[],
        quality=None,
        statistical_reconciliation=None,
        anomalies=[],
        warnings=[],
        ai={
            "response": {
                "summary": "جمع کل 2,500,000 است.",
                "findings": ["رقم نامعتبر 7777777 دیده شد."],
                "recommendations": ["عدد 8888888 را بررسی کنید.", "جمع را با منبع تطبیق دهید."],
            }
        },
    )

    content = ReportContent.from_analysis(analysis)

    assert "جمع کل 2,500,000 است." in content.executive_summary
    joined = "\n".join(content.executive_summary + content.recommendations)
    assert "7777777" not in joined
    assert "8888888" not in joined
    assert any("روایت مدل" in warning for warning in content.warnings)


def test_model_date_claim_is_dropped_even_when_its_parts_are_approved():
    analysis = SimpleNamespace(
        meta=SimpleNamespace(path="narrative.xlsx"),
        primary_table=None,
        integration=None,
        calculations=[
            SimpleNamespace(
                name="year",
                label="سال",
                value=Decimal("1403"),
                unit="",
                sources=[],
            ),
            SimpleNamespace(
                name="count",
                label="تعداد",
                value=Decimal("4"),
                unit="",
                sources=[],
            ),
        ],
        generic=None,
        request_answers=[],
        quality=None,
        statistical_reconciliation=None,
        anomalies=[],
        warnings=[],
        ai={
            "response": {
                "summary": "دوره 1403/04 بررسی شد.",
                "findings": [],
                "recommendations": [],
            }
        },
    )

    content = ReportContent.from_analysis(analysis)

    assert "1403/04" not in "\n".join(content.executive_summary)
    assert any("روایت مدل" in warning for warning in content.warnings)


def test_manual_content_tables_must_match_their_headers():
    content = ReportContent(
        title="گزارش",
        tables=[
            ReportTable("جدول", ["شرح", "مقدار"], [["جمع", "1", "اضافه"]]),
        ],
    )
    errors = content.validate()
    assert any("هم‌خوان نیست" in error for error in errors)


# --------------------------------------------------------------------------- #
# 7. Missing template handling
# --------------------------------------------------------------------------- #
def test_missing_template_is_rejected_but_no_template_builds_a_report(
    settings, report_content, tmp_path
):
    _, content = report_content
    missing = tmp_path / "missing.docx"

    with pytest.raises(ReportTemplateError, match="پیدا نشد"):
        _build(content, tmp_path / "out.docx", template_path=missing)

    with pytest.raises(ReportTemplateError, match="پیدا نشد"):
        ReportSpecification(
            OutputFormat.DOCX,
            tmp_path / "out.docx",
            template_path=missing,
        ).validate()

    result = _build(content, tmp_path / "without-template.docx")
    assert result.mode == "standalone"
    assert result.validation.valid is True


# --------------------------------------------------------------------------- #
# 8. Invalid output paths
# --------------------------------------------------------------------------- #
def test_invalid_output_paths_are_rejected(report_content, rtl_template, tmp_path):
    _, content = report_content

    with pytest.raises(ReportFormatError, match="سازگار نیست"):
        _build(content, tmp_path / "report.xlsx")

    with pytest.raises(ReportFormatError, match="سازگار نیست"):
        _build(content, tmp_path / "report.pdf")

    with pytest.raises(ReportFormatError, match="یکسان است"):
        _build(content, rtl_template, template_path=rtl_template)

    directory = tmp_path / "report.docx"
    directory.mkdir()
    with pytest.raises(ReportFormatError, match="پوشه"):
        _build(content, directory)

    blocked_parent = tmp_path / "blocked"
    blocked_parent.write_text("not a directory", encoding="utf-8")
    with pytest.raises(ReportFormatError, match="قابل نوشتن"):
        _build(content, blocked_parent / "report.docx")

    nested = tmp_path / "a" / "b" / "c" / "report.docx"
    result = _build(content, nested)
    assert nested.exists()
    assert result.validation.valid is True


# --------------------------------------------------------------------------- #
# 9. Reopen and validate the generated DOCX
# --------------------------------------------------------------------------- #
def test_generated_docx_is_reopened_and_structurally_validated(report_content, tmp_path):
    _, content = report_content
    output = tmp_path / "validated.docx"
    result = _build(content, output)

    validation = validate_docx_output(
        output,
        content,
        required_sections=("خلاصه اجرایی", "نتیجه‌گیری و پیشنهادها", "یادداشت‌های منبع"),
        minimum_tables=1,
    )
    assert validation.valid is True
    assert validation.paragraph_count > 0
    assert validation.table_count >= 1
    assert validation.heading_titles
    assert validation.has_header is True
    assert validation.has_footer is True
    assert validation.has_page_number_field is True
    assert validation.unresolved_placeholders == ()
    assert result.validation is not None and result.validation.valid is True

    header_validation = validate_docx_output(
        output, content, required_table_headers=[("شرح", "مقدار")]
    )
    assert header_validation.valid is True
    missing_headers = validate_docx_output(
        output, content, required_table_headers=[("ستون ناموجود", "مقدار")]
    )
    assert missing_headers.valid is False
    assert any("سرستون" in error for error in missing_headers.errors)

    plain_text = tmp_path / "fake.docx"
    plain_text.write_text("this is not a Word file", encoding="utf-8")
    fake = validate_docx_output(plain_text)
    assert fake.valid is False
    assert any("ZIP" in error for error in fake.errors)


def test_corrupt_header_xml_is_rejected(report_content, tmp_path):
    _, content = report_content
    output = tmp_path / "valid.docx"
    _build(content, output)

    corrupt = tmp_path / "corrupt-header.docx"
    with zipfile.ZipFile(output) as source, zipfile.ZipFile(corrupt, "w") as target:
        replaced = False
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename.startswith("word/header") and item.filename.endswith(".xml"):
                data = b"<w:hdr"
                replaced = True
            target.writestr(item, data)
    assert replaced is True

    validation = validate_docx_output(corrupt)
    assert validation.valid is False
    assert any("XML" in error for error in validation.errors)


def test_generation_result_serializes_writer_and_validation(report_content, tmp_path):
    _, content = report_content
    specification = ReportSpecification(
        OutputFormat.DOCX,
        tmp_path / "generated.docx",
    )
    result = ReportGenerator(_settings_for(tmp_path)).generate(
        specification,
        analysis=report_content[0],
        content=content,
    )
    payload = result.to_dict()
    assert payload["output_format"] == "docx"
    assert payload["writer"] == "DocxReportWriter"
    assert payload["mode"] == "standalone"
    assert payload["validation"]["valid"] is True
    assert payload["tables_written"] >= 1
    assert result.output_path.exists()


# --------------------------------------------------------------------------- #
# End-to-end CLI quality gate
# --------------------------------------------------------------------------- #
def test_cli_generates_a_word_report_without_a_template(project_root, report_source):
    output = project_root / "cli-report.docx"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(report_source),
            "--no-ai",
            "--yes",
            "--quiet",
            "--output",
            str(output),
        ]
    )
    assert code == 0
    assert output.exists()
    document = Document(str(output))
    assert any("خلاصه اجرایی" in paragraph.text for paragraph in document.paragraphs)
    validation = validate_docx_output(output)
    assert validation.valid is True
    assert validation.has_page_number_field is True


def test_cli_generates_a_word_report_with_a_template_and_keeps_it(
    project_root, report_source, rtl_template
):
    before = rtl_template.read_bytes()
    output = project_root / "cli-template.docx"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(report_source),
            "--no-ai",
            "--yes",
            "--quiet",
            "--output",
            str(output),
            "--template",
            str(rtl_template),
        ]
    )
    assert code == 0
    assert output.exists()
    assert rtl_template.read_bytes() == before
    text = _all_text(Document(str(output)))
    assert HISTORICAL_FIGURE not in text
    assert "2,500,000" in text


def test_cli_can_keep_template_figures_as_explicit_reference(
    project_root, report_source, rtl_template
):
    output = project_root / "cli-template-reference.docx"
    code = main(
        [
            "--root",
            str(project_root),
            "--file",
            str(report_source),
            "--no-ai",
            "--yes",
            "--quiet",
            "--output",
            str(output),
            "--template",
            str(rtl_template),
            "--keep-template-figures",
        ]
    )
    assert code == 0
    assert output.exists()
    assert HISTORICAL_FIGURE in _all_text(Document(str(output)))


def _settings_for(tmp_path: Path):
    from local_office_ai.core.settings import load_settings

    root = tmp_path / "project"
    (root / "config").mkdir(parents=True, exist_ok=True)
    (root / "data").mkdir(exist_ok=True)
    (root / "logs").mkdir(exist_ok=True)
    return load_settings(project_root=root)
