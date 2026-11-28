"""Professional Persian DOCX report writer with real RTL rendering.

Two modes share one validated :class:`~local_office_ai.reporting.specification.ReportContent`:

* standalone -- a complete report is built from scratch (title page, headings,
  executive summary, RTL tables, captions, source/statistical notes, page
  numbers, conclusions and recommendations);
* template -- the supplied template keeps its layout and styles, placeholders
  are replaced, static historical figures are removed, and the analytical body
  is inserted at a ``{{content}}`` anchor or appended at the end.

Every figure comes from ``ReportContent``.  The finished file is reopened and
validated before it is returned.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from local_office_ai.core.exceptions import ReportFormatError, ReportTemplateError
from local_office_ai.reporting.docx_rtl import (
    add_page_number_field,
    contains_persian,
    ensure_paragraph_bidi,
    format_paragraph,
    iter_header_footer_containers,
    oxml,
    paragraph_has_field,
    qn,
    set_run_rtl,
    set_section_rtl,
    set_table_rtl,
)
from local_office_ai.reporting.docx_validation import DocxValidation, validate_docx_output
from local_office_ai.reporting.docx_writer import DocxTemplateRenderer
from local_office_ai.reporting.formats import OutputFormat, ensure_extension_matches
from local_office_ai.reporting.specification import ReportContent, ReportTable
from local_office_ai.reporting.template import PLACEHOLDER_PATTERN, ReportTemplate, inspect_docx_template

FONT_NAME = "B Nazanin"
BODY_PT = 12.0
HEADING_COLOR = "1F3A5F"
MUTED_COLOR = "595959"

_STATIC_FIGURE = re.compile(r"(?<![\w])[0-9۰-۹٠-٩][0-9۰-۹٠-٩,٬٫./\-]{2,}")
_CONTENT_ANCHORS = ("{{content}}", "{{body}}", "[[CONTENT]]")


@dataclass
class DocxReportResult:
    output_path: Path
    mode: str
    sections_written: tuple[str, ...] = ()
    tables_written: int = 0
    placeholders_resolved: tuple[str, ...] = ()
    unresolved_placeholders: tuple[str, ...] = ()
    removed_static_figures: tuple[str, ...] = ()
    warnings: list[str] = field(default_factory=list)
    validation: DocxValidation | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": str(self.output_path),
            "mode": self.mode,
            "sections_written": list(self.sections_written),
            "tables_written": self.tables_written,
            "placeholders_resolved": list(self.placeholders_resolved),
            "unresolved_placeholders": list(self.unresolved_placeholders),
            "removed_static_figures": list(self.removed_static_figures),
            "warnings": list(self.warnings),
            "validation": self.validation.to_dict() if self.validation else None,
        }


class DocxReportWriter:
    """Build a professional Persian Word report or render a template."""

    def __init__(self) -> None:
        self._skipped_template_sections: list[str] = []

    def build(
        self,
        content: ReportContent,
        output_path: Path | str,
        *,
        template_path: Path | str | None = None,
        strip_static_figures: bool = True,
    ) -> DocxReportResult:
        content.ensure_valid()
        output = ensure_extension_matches(Path(output_path), OutputFormat.DOCX)
        self._skipped_template_sections = []
        extra_allowed_numbers: tuple[str, ...] = ()
        if output.exists() and output.is_dir():
            raise ReportFormatError(
                f"مسیر خروجی Word یک پوشه است، نه فایل: {output}"
            )
        if template_path is None or str(template_path).strip() == "":
            result = self._build_standalone(content, output)
        else:
            template = inspect_docx_template(template_path)
            if output.resolve() == Path(template.path).resolve():
                raise ReportFormatError(
                    "مسیر خروجی با مسیر قالب Word یکسان است؛ قالب باید دست‌نخورده بماند."
                )
            result = self._build_from_template(
                content,
                output,
                template,
                strip_static_figures=strip_static_figures,
            )
            if not strip_static_figures:
                extra_allowed_numbers = tuple(
                    item.text for item in template.static_numeric_literals
                )

        validation = validate_docx_output(
            output,
            content,
            required_sections=result.sections_written,
            minimum_tables=1 if content.tables else 0,
            require_rtl=True,
            require_page_numbers=True,
            allow_unresolved_placeholders=not result.unresolved_placeholders,
            allow_benign_numbers=result.mode == "template",
            extra_allowed_numbers=extra_allowed_numbers,
        )
        result.validation = validation
        if not validation.valid:
            raise ReportFormatError(
                "خروجی Word اعتبارسنجی نشد و تحویل داده نمی‌شود: "
                + "؛ ".join(validation.errors)
            )
        return result

    # ------------------------------------------------------------------ #
    def _build_standalone(self, content: ReportContent, output: Path) -> DocxReportResult:
        Document = _document_class()
        document = Document()
        self._configure_document(document, content)
        section = document.sections[0]
        self._write_header_footer(section, content)
        self._write_title_page(document, content)
        builder = _BodyWriter(document)
        sections = self._write_body(builder, content, existing=set())
        _save_atomic(document, output)
        return DocxReportResult(
            output_path=output,
            mode="standalone",
            sections_written=tuple(sections),
            tables_written=len(content.tables),
        )

    def _build_from_template(
        self,
        content: ReportContent,
        output: Path,
        template: ReportTemplate,
        *,
        strip_static_figures: bool,
    ) -> DocxReportResult:
        if not template.valid:
            raise ReportTemplateError(
                "قالب Word معتبر نیست: " + "؛ ".join(template.errors or ["دلیل نامشخص"])
            )
        Document = _document_class()
        try:
            document = Document(str(template.path))
        except Exception as exc:  # noqa: BLE001 - parser errors vary by python-docx version
            raise ReportTemplateError(
                f"قالب Word قابل خواندن نیست: {template.path} ({exc})"
            ) from exc

        anchor = _find_content_anchor(document)
        removed: list[str] = []
        if strip_static_figures:
            removed = _strip_static_figures(document)

        renderer = DocxTemplateRenderer()
        if anchor is not None:
            anchor_keys = {
                key
                for key in template.placeholder_keys
                if f"{{{{{key}}}}}" in _CONTENT_ANCHORS
                or f"[[{key}]]" in _CONTENT_ANCHORS
            }
            if anchor_keys:
                renderer.replace_in_document(document, {key: "" for key in anchor_keys})
        replaced, unresolved = renderer.replace_in_document(document, content.placeholder_values())
        if anchor is not None:
            for token in _CONTENT_ANCHORS:
                if token in anchor.text:
                    for run in anchor.runs:
                        if token in run.text:
                            run.text = run.text.replace(token, "")
        anchor_only = anchor is not None and not anchor.text.strip()

        self._apply_rtl_to_template(document)
        self._ensure_template_page_numbers(document, content)

        existing = {
            _normalize(paragraph.text)
            for paragraph in document.paragraphs
            if _is_heading(paragraph) and paragraph.text.strip()
        }
        builder = _BodyWriter(document, anchor=anchor, after=not anchor_only)
        if anchor is None and any(paragraph.text.strip() for paragraph in document.paragraphs):
            builder.page_break()
        sections = self._write_body(builder, content, existing=existing)
        if anchor is not None and anchor_only:
            _remove_paragraph(anchor)

        warnings: list[str] = list(template.warnings)
        if unresolved:
            warnings.append(
                "برای این جایگاه‌ها مقدار معتبر پیدا نشد و متن اصلی حفظ شد: "
                + "، ".join(sorted(unresolved))
            )
        if removed:
            warnings.append(
                f"{len(removed)} عدد ثابت قالب به‌عنوان داده تاریخی حذف شد: "
                + "، ".join(removed[:10])
            )
        if self._skipped_template_sections:
            warnings.append(
                "این بخش‌ها در قالب از قبل وجود داشتند و برای جلوگیری از تکرار دوباره افزوده نشدند: "
                + "، ".join(self._skipped_template_sections)
            )
        warnings.append(
            "قالب Word فقط برای چیدمان استفاده شد؛ هیچ عددی از قالب به‌عنوان داده گزارش پذیرفته نشد."
        )

        _save_atomic(document, output)
        return DocxReportResult(
            output_path=output,
            mode="template",
            sections_written=tuple(sections),
            tables_written=sum(
                1 for table in content.tables if _normalize(table.title) not in existing
            ),
            placeholders_resolved=tuple(
                sorted(
                    key
                    for key in set(template.placeholder_keys) - set(unresolved)
                    if f"{{{{{key}}}}}" not in _CONTENT_ANCHORS
                )
            ),
            unresolved_placeholders=tuple(sorted(unresolved)),
            removed_static_figures=tuple(removed),
            warnings=warnings,
        )

    # ------------------------------------------------------------------ #
    def _configure_document(self, document: Any, content: ReportContent) -> None:
        from docx.shared import Cm

        section = document.sections[0]
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
        section.header_distance = Cm(1.2)
        section.footer_distance = Cm(1.2)
        section.different_first_page_header_footer = True
        set_section_rtl(section)

        properties = document.core_properties
        properties.title = content.title
        properties.subject = content.subtitle or "گزارش تحلیلی"
        properties.author = "Local Office AI"
        properties.language = "fa-IR"
        properties.comments = "گزارش تولیدشده با محاسبات قطعی Python"

        self._configure_styles(document)

    def _configure_styles(self, document: Any) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt, RGBColor

        styles = document.styles
        heading_color = RGBColor.from_string(HEADING_COLOR)
        muted = RGBColor.from_string(MUTED_COLOR)

        normal = _style(styles, "Normal")
        if normal is not None:
            normal.font.name = FONT_NAME
            normal.font.size = Pt(BODY_PT)
            normal.font.rtl = True
            normal.paragraph_format.line_spacing = 1.5
            normal.paragraph_format.space_after = Pt(6)
            _style_complex_font(normal, FONT_NAME, BODY_PT)

        for name, size, color, before, after in (
            ("Heading 1", 16.0, heading_color, 18, 10),
            ("Heading 2", 14.0, heading_color, 14, 8),
            ("Heading 3", 12.5, heading_color, 10, 6),
        ):
            style = _style(styles, name)
            if style is None:
                continue
            style.font.name = FONT_NAME
            style.font.size = Pt(size)
            style.font.bold = True
            style.font.rtl = True
            style.font.color.rgb = color
            style.paragraph_format.space_before = Pt(before)
            style.paragraph_format.space_after = Pt(after)
            style.paragraph_format.keep_with_next = True
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _style_complex_font(style, FONT_NAME, size)

        title = _style(styles, "Title")
        if title is not None:
            title.font.name = FONT_NAME
            title.font.size = Pt(26)
            title.font.bold = True
            title.font.rtl = True
            title.font.color.rgb = heading_color
            title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title.paragraph_format.space_before = Pt(90)
            title.paragraph_format.space_after = Pt(12)
            _style_complex_font(title, FONT_NAME, 26)

        subtitle = _style(styles, "Subtitle")
        if subtitle is not None:
            subtitle.font.name = FONT_NAME
            subtitle.font.size = Pt(14)
            subtitle.font.rtl = True
            subtitle.font.color.rgb = muted
            subtitle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            subtitle.paragraph_format.space_after = Pt(24)
            _style_complex_font(subtitle, FONT_NAME, 14)

        caption = _style(styles, "Caption")
        if caption is not None:
            caption.font.name = FONT_NAME
            caption.font.size = Pt(9.5)
            caption.font.italic = True
            caption.font.rtl = True
            caption.font.color.rgb = muted
            caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _style_complex_font(caption, FONT_NAME, 9.5)

        bullet = _style(styles, "List Bullet")
        if bullet is not None:
            bullet.font.name = FONT_NAME
            bullet.font.size = Pt(BODY_PT)
            bullet.font.rtl = True
            _style_complex_font(bullet, FONT_NAME, BODY_PT)

    def _write_header_footer(self, section: Any, content: ReportContent) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Pt, RGBColor

        muted = RGBColor.from_string(MUTED_COLOR)
        header = section.header
        header.is_linked_to_previous = False
        paragraph = header.paragraphs[0]
        paragraph.text = ""
        run = paragraph.add_run(content.title)
        set_run_rtl(run, font_name=FONT_NAME, size_pt=9.5, color=muted)
        format_paragraph(
            paragraph,
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            space_after_pt=0,
            line_spacing=1.0,
        )
        _add_bottom_border(paragraph, color="BFBFBF", size=4)

        footer = section.footer
        footer.is_linked_to_previous = False
        footer_paragraph = footer.paragraphs[0]
        footer_paragraph.text = ""
        date_run = footer_paragraph.add_run(f"{content.generated_at}  |  ")
        set_run_rtl(date_run, font_name=FONT_NAME, size_pt=9, color=muted)
        add_page_number_field(footer_paragraph, font_name=FONT_NAME, size_pt=9)
        format_paragraph(
            footer_paragraph,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            space_before_pt=4,
            space_after_pt=0,
            line_spacing=1.0,
        )

    def _write_title_page(self, document: Any, content: ReportContent) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        title = document.add_paragraph(content.title, style=_style_name(document, "Title"))
        format_paragraph(
            title,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            size_pt=26,
            bold=True,
            space_before_pt=90,
            space_after_pt=12,
            line_spacing=1.2,
        )
        if content.subtitle:
            subtitle = document.add_paragraph(
                content.subtitle, style=_style_name(document, "Subtitle")
            )
            format_paragraph(
                subtitle,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                size_pt=14,
                space_after_pt=24,
                line_spacing=1.3,
            )
        meta = document.add_paragraph()
        run = meta.add_run(f"تاریخ گزارش: {content.generated_at}")
        set_run_rtl(run, font_name=FONT_NAME, size_pt=11)
        source = document.add_paragraph()
        source_run = source.add_run(
            f"فایل منبع: {content.values['report.source_file'].value}"
            if "report.source_file" in content.values
            else "فایل منبع: —"
        )
        set_run_rtl(source_run, font_name=FONT_NAME, size_pt=11)
        for paragraph in (meta, source):
            format_paragraph(
                paragraph,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                space_after_pt=4,
                line_spacing=1.3,
            )
        self._page_break(document)

    def _write_body(
        self,
        builder: "_BodyWriter",
        content: ReportContent,
        *,
        existing: set[str],
    ) -> list[str]:
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        written: list[str] = []
        first_heading = True

        if content.executive_summary:
            if _normalize("خلاصه اجرایی") in existing:
                self._skipped_template_sections.append("خلاصه اجرایی")
            else:
                heading = self._heading(builder, "خلاصه اجرایی", 1)
                heading.paragraph_format.page_break_before = not first_heading
                first_heading = False
                written.append("خلاصه اجرایی")
                for text in content.executive_summary:
                    paragraph = builder.paragraph(text)
                    format_paragraph(
                        paragraph,
                        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                        size_pt=BODY_PT,
                        space_after_pt=6,
                        line_spacing=1.5,
                    )

        for table in content.tables:
            if _normalize(table.title) in existing:
                self._skipped_template_sections.append(table.title)
                continue
            heading = self._heading(builder, table.title, 2)
            heading.paragraph_format.page_break_before = False
            written.append(table.title)
            caption = builder.paragraph(
                table.caption or f"جدول {table.title}",
                style=_style_name(builder.document, "Caption"),
            )
            format_paragraph(
                caption,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                size_pt=9.5,
                bold=False,
                space_before_pt=6,
                space_after_pt=4,
                line_spacing=1.0,
            )
            self._write_table(builder, table)
            if table.source_note:
                self._write_note(builder, table.source_note)
            if table.statistical_note:
                self._write_note(builder, table.statistical_note)

        issues = [
            item.value
            for key, item in content.values.items()
            if key.startswith("issue.") and item.value
        ]
        warning_lines = [str(item) for item in content.warnings] + issues
        if warning_lines:
            if _normalize("هشدارها و موارد نیازمند بررسی") in existing:
                self._skipped_template_sections.append("هشدارها و موارد نیازمند بررسی")
            else:
                heading = self._heading(builder, "هشدارها و موارد نیازمند بررسی", 1)
                heading.paragraph_format.page_break_before = not first_heading
                first_heading = False
                written.append("هشدارها و موارد نیازمند بررسی")
                for text in warning_lines:
                    self._write_bullet(builder, text)

        if content.conclusions or content.recommendations:
            if _normalize("نتیجه‌گیری و پیشنهادها") in existing:
                self._skipped_template_sections.append("نتیجه‌گیری و پیشنهادها")
            else:
                heading = self._heading(builder, "نتیجه‌گیری و پیشنهادها", 1)
                heading.paragraph_format.page_break_before = not first_heading
                first_heading = False
                written.append("نتیجه‌گیری و پیشنهادها")
                if content.conclusions:
                    self._heading(builder, "نتیجه‌گیری", 2)
                    for text in content.conclusions:
                        self._write_bullet(builder, text)
                if content.recommendations:
                    self._heading(builder, "پیشنهادها", 2)
                    for text in content.recommendations:
                        self._write_bullet(builder, text)

        if content.source_notes:
            if _normalize("یادداشت‌های منبع") in existing:
                self._skipped_template_sections.append("یادداشت‌های منبع")
            else:
                heading = self._heading(builder, "یادداشت‌های منبع", 1)
                heading.paragraph_format.page_break_before = not first_heading
                first_heading = False
                written.append("یادداشت‌های منبع")
                for text in content.source_notes:
                    self._write_bullet(builder, text)

        if content.statistical_notes:
            if _normalize("یادداشت‌های آماری") in existing:
                self._skipped_template_sections.append("یادداشت‌های آماری")
            else:
                heading = self._heading(builder, "یادداشت‌های آماری", 1)
                heading.paragraph_format.page_break_before = not first_heading
                first_heading = False
                written.append("یادداشت‌های آماری")
                for text in content.statistical_notes:
                    self._write_bullet(builder, text)

        return written

    def _heading(self, builder: "_BodyWriter", text: str, level: int) -> Any:
        heading = builder.heading(text, level)
        ensure_paragraph_bidi(heading)
        size = {1: 16.0, 2: 14.0, 3: 12.5}.get(level)
        for run in heading.runs:
            set_run_rtl(run, font_name=FONT_NAME, size_pt=size, bold=True)
        return heading

    def _write_table(self, builder: "_BodyWriter", table: ReportTable) -> None:
        from docx.enum.table import WD_TABLE_ALIGNMENT

        rows = len(table.rows) + 1
        columns = max(1, len(table.headers))
        doc_table = builder.table(rows=rows, cols=columns)
        doc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        doc_table.autofit = True
        _set_table_width_percent(doc_table, 100)

        style_name = _style_name(builder.document, "Table Grid")
        if style_name:
            try:
                doc_table.style = style_name
            except (KeyError, ValueError):
                pass

        for column, header in enumerate(table.headers):
            cell = doc_table.cell(0, column)
            cell.text = str(header)
        for row_index, values in enumerate(table.rows, start=1):
            for column, value in enumerate(values):
                doc_table.cell(row_index, column).text = str(value)

        set_table_rtl(doc_table)
        for row in doc_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    format_paragraph(
                        paragraph,
                        size_pt=10,
                        space_before_pt=1,
                        space_after_pt=1,
                        line_spacing=1.15,
                    )

    def _write_note(self, builder: "_BodyWriter", text: str) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import RGBColor

        paragraph = builder.paragraph(text, style=_style_name(builder.document, "Caption"))
        format_paragraph(
            paragraph,
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            size_pt=9,
            bold=False,
            color=RGBColor.from_string(MUTED_COLOR),
            space_before_pt=2,
            space_after_pt=2,
            line_spacing=1.1,
        )

    def _write_bullet(self, builder: "_BodyWriter", text: str) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        style_name = _style_name(builder.document, "List Bullet")
        paragraph = builder.paragraph(text, style=style_name)
        format_paragraph(
            paragraph,
            alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            size_pt=BODY_PT,
            space_after_pt=3,
            line_spacing=1.4,
        )

    def _page_break(self, document: Any) -> None:
        from docx.enum.text import WD_BREAK

        paragraph = document.add_paragraph()
        paragraph.add_run().add_break(WD_BREAK.PAGE)
        ensure_paragraph_bidi(paragraph)

    def _apply_rtl_to_template(self, document: Any) -> None:
        paragraphs = list(document.paragraphs)
        for container in iter_header_footer_containers(document):
            paragraphs.extend(container.paragraphs)
        for section in document.sections:
            if section._sectPr.find(qn("w:bidi")) is None:
                set_section_rtl(section)
        for paragraph in paragraphs:
            if contains_persian(paragraph.text):
                ensure_paragraph_bidi(paragraph)
                for run in paragraph.runs:
                    set_run_rtl(run, font_name=FONT_NAME)
        tables = list(document.tables)
        for container in iter_header_footer_containers(document):
            tables.extend(container.tables)
        for table in tables:
            if any(contains_persian(cell.text) for row in table.rows for cell in row.cells):
                set_table_rtl(table, shade_header=False)

    def _ensure_template_page_numbers(self, document: Any, content: ReportContent) -> None:
        if _has_page_field(document):
            return
        for section in document.sections:
            footer = section.footer
            footer.is_linked_to_previous = False
            paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            if paragraph.text.strip():
                paragraph.add_run("  |  ")
            date_run = paragraph.add_run(f"{content.generated_at}  |  ")
            set_run_rtl(date_run, font_name=FONT_NAME, size_pt=9)
            add_page_number_field(paragraph, font_name=FONT_NAME, size_pt=9)
            break


class _BodyWriter:
    """Append content, or insert it next to a template anchor."""

    def __init__(self, document: Any, *, anchor: Any = None, after: bool = False) -> None:
        self.document = document
        self.anchor = anchor
        self.after = after
        self._last = anchor._p if anchor is not None and after else None

    def paragraph(self, text: str = "", style: str | None = None) -> Any:
        paragraph = self.document.add_paragraph(text, style=style) if style else self.document.add_paragraph(text)
        self._move(paragraph._p)
        return paragraph

    def heading(self, text: str, level: int) -> Any:
        paragraph = self.document.add_heading(text, level=level)
        self._move(paragraph._p)
        return paragraph

    def table(self, rows: int, cols: int) -> Any:
        table = self.document.add_table(rows=rows, cols=cols)
        self._move(table._tbl)
        return table

    def page_break(self) -> Any:
        from docx.enum.text import WD_BREAK

        paragraph = self.document.add_paragraph()
        paragraph.add_run().add_break(WD_BREAK.PAGE)
        self._move(paragraph._p)
        return paragraph

    def _move(self, element: Any) -> None:
        if self.anchor is not None:
            if self.after:
                assert self._last is not None
                self._last.addnext(element)
                self._last = element
            else:
                self.anchor._p.addprevious(element)


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _document_class() -> Any:
    try:
        from docx import Document
    except ImportError as exc:  # pragma: no cover - optional dependency branch
        raise ReportFormatError(
            "برای تولید گزارش Word، کتابخانه python-docx نصب نیست."
        ) from exc
    return Document


def _style(styles: Any, name: str) -> Any:
    try:
        return styles[name]
    except KeyError:
        return None


def _style_name(document: Any, name: str) -> str | None:
    return name if _style(document.styles, name) is not None else None


def _style_complex_font(style: Any, name: str, size_pt: float) -> None:
    r_pr = style.element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:cs"), name)
    r_fonts.set(qn("w:ascii"), name)
    r_fonts.set(qn("w:hAnsi"), name)
    sz_cs = r_pr.find(qn("w:szCs"))
    if sz_cs is None:
        sz_cs = oxml("w:szCs")
        r_pr.insert_element_before(
            sz_cs,
            "w:highlight",
            "w:u",
            "w:effect",
            "w:bdr",
            "w:shd",
            "w:fitText",
            "w:vertAlign",
            "w:rtl",
            "w:cs",
            "w:em",
            "w:lang",
            "w:eastAsianLayout",
            "w:specVanish",
            "w:oMath",
            "w:rPrChange",
        )
    sz_cs.set(qn("w:val"), str(int(round(size_pt * 2))))


def _add_bottom_border(paragraph: Any, *, color: str, size: int) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(qn("w:pBdr")) is not None:
        return
    border = oxml("w:pBdr")
    bottom = oxml("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    border.append(bottom)
    p_pr.insert_element_before(border, "w:shd", "w:tabs", "w:spacing", "w:jc", "w:rPr")


def _set_table_width_percent(table: Any, percent: int) -> None:
    tbl_pr = table._tbl.tblPr
    width = tbl_pr.find(qn("w:tblW"))
    if width is None:
        width = oxml("w:tblW")
        tbl_pr.insert_element_before(
            width,
            "w:jc",
            "w:tblCellSpacing",
            "w:tblInd",
            "w:tblBorders",
            "w:shd",
            "w:tblLayout",
            "w:tblCellMar",
            "w:tblLook",
            "w:tblCaption",
            "w:tblDescription",
            "w:tblPrChange",
        )
    width.set(qn("w:type"), "pct")
    width.set(qn("w:w"), str(percent * 50))


def _find_content_anchor(document: Any) -> Any:
    for paragraph in document.paragraphs:
        if any(anchor in paragraph.text for anchor in _CONTENT_ANCHORS):
            return paragraph
    return None


def _remove_paragraph(paragraph: Any) -> None:
    element = paragraph._p
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _strip_static_figures(document: Any) -> list[str]:
    removed: list[str] = []
    for paragraph in document.paragraphs:
        removed.extend(_strip_paragraph(paragraph))
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    removed.extend(_strip_paragraph(paragraph))
    for container in iter_header_footer_containers(document):
        for paragraph in container.paragraphs:
            removed.extend(_strip_paragraph(paragraph))
    return removed


def _strip_paragraph(paragraph: Any) -> list[str]:
    if paragraph_has_field(paragraph):
        return []
    text = paragraph.text
    if not text or not _STATIC_FIGURE.search(PLACEHOLDER_PATTERN.sub(" ", text)):
        return []
    removed: list[str] = []
    pieces: list[str] = []
    cursor = 0

    def _replace(match: re.Match[str]) -> str:
        removed.append(match.group(0))
        return ""

    for placeholder in PLACEHOLDER_PATTERN.finditer(text):
        pieces.append(_STATIC_FIGURE.sub(_replace, text[cursor : placeholder.start()]))
        pieces.append(placeholder.group(0))
        cursor = placeholder.end()
    pieces.append(_STATIC_FIGURE.sub(_replace, text[cursor:]))
    new_text = "".join(pieces)
    new_text = re.sub(r"[ \t]{2,}", " ", new_text).strip()
    if new_text == text:
        return []
    if paragraph.runs and all(_run_is_text_only(run) for run in paragraph.runs):
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    return removed


def _run_is_text_only(run: Any) -> bool:
    allowed = {qn("w:rPr"), qn("w:t")}
    return all(child.tag in allowed for child in run._r)


def _has_page_field(document: Any) -> bool:
    for section in document.sections:
        for container in (
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        ):
            if _is_linked(container):
                continue
            for element in container._element.iter(qn("w:instrText")):
                if "PAGE" in (element.text or "").upper():
                    return True
    return False


def _is_linked(container: Any) -> bool:
    try:
        return bool(container.is_linked_to_previous)
    except (AttributeError, ValueError):
        return False


def _is_heading(paragraph: Any) -> bool:
    name = (getattr(paragraph.style, "name", "") or "").lower()
    return name.startswith("heading") or name in {"title", "subtitle"}


def _normalize(text: str) -> str:
    return " ".join((text or "").replace("\u200c", " ").split())


def _save_atomic(document: Any, output: Path) -> None:
    from uuid import uuid4
    import os
    import zipfile

    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_name(f".{output.stem}.{uuid4().hex}.tmp{output.suffix}")
    try:
        document.save(str(temp))
        if not zipfile.is_zipfile(temp):
            raise ReportFormatError("خروجی Word معتبر ساخته نشد (فایل ZIP نیست).")
        with zipfile.ZipFile(temp) as archive:
            if "word/document.xml" not in archive.namelist():
                raise ReportFormatError(
                    "خروجی Word معتبر ساخته نشد (word/document.xml ندارد)."
                )
        os.replace(temp, output)
    except Exception:
        temp.unlink(missing_ok=True)
        raise
