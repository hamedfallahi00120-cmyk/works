"""Output-format decision layer and DOCX template contract."""

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
    "DocxTemplateRenderer",
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
    "ReportTemplate",
    "ReportValue",
    "StaticNumericLiteral",
    "TableInfo",
    "TemplateRenderResult",
    "detect_requested_formats",
    "ensure_extension_matches",
    "ensure_xlsx_delivery",
    "inspect_docx_template",
    "resolve_output_intent",
]
