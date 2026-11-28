"""Output-format decision layer and DOCX report writers."""

from local_office_ai.reporting.docx_report_writer import DocxReportResult, DocxReportWriter
from local_office_ai.reporting.docx_validation import (
    DocxValidation,
    ensure_valid_docx,
    validate_docx_output,
)
from local_office_ai.reporting.docx_writer import DocxTemplateRenderer, TemplateRenderResult
from local_office_ai.reporting.formats import (
    OutputFormat,
    OutputIntent,
    detect_requested_formats,
    ensure_extension_matches,
    ensure_xlsx_delivery,
    resolve_output_intent,
)
from local_office_ai.reporting.generator import (
    ReportGenerationPlan,
    ReportGenerator,
)
from local_office_ai.reporting.specification import (
    ReportContent,
    ReportGenerationResult,
    ReportRow,
    ReportSection,
    ReportSpecification,
    ReportTable,
    ReportValue,
)
from local_office_ai.reporting.template import (
    FontUsage,
    HeaderFooterInfo,
    NumberingInfo,
    PageLayout,
    ParagraphInfo,
    ParagraphStyleInfo,
    PlaceholderInfo,
    ReportTemplate,
    StaticNumericLiteral,
    TableInfo,
    inspect_docx_template,
)

__all__ = [
    "DocxReportResult",
    "DocxReportWriter",
    "DocxTemplateRenderer",
    "DocxValidation",
    "FontUsage",
    "HeaderFooterInfo",
    "NumberingInfo",
    "OutputFormat",
    "OutputIntent",
    "PageLayout",
    "ParagraphInfo",
    "ParagraphStyleInfo",
    "PlaceholderInfo",
    "ReportContent",
    "ReportGenerationPlan",
    "ReportGenerationResult",
    "ReportGenerator",
    "ReportRow",
    "ReportSection",
    "ReportSpecification",
    "ReportTable",
    "ReportTemplate",
    "ReportValue",
    "StaticNumericLiteral",
    "TableInfo",
    "TemplateRenderResult",
    "detect_requested_formats",
    "ensure_extension_matches",
    "ensure_valid_docx",
    "ensure_xlsx_delivery",
    "inspect_docx_template",
    "resolve_output_intent",
    "validate_docx_output",
]
