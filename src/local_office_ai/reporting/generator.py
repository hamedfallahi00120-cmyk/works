"""Select and run the writer that matches a validated report specification."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from local_office_ai.core.exceptions import ReportFormatError
from local_office_ai.core.settings import Settings
from local_office_ai.reporting.docx_report_writer import DocxReportWriter
from local_office_ai.reporting.formats import OutputFormat
from local_office_ai.reporting.specification import (
    ReportContent,
    ReportGenerationResult,
    ReportSpecification,
)
from local_office_ai.reporting.template import inspect_docx_template


@dataclass(frozen=True)
class ReportGenerationPlan:
    """The decision layer's output: which writer will handle which file."""

    output_format: OutputFormat
    writer: str
    output_path: Path
    template_path: Path | None = None
    requires_template: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_format": self.output_format.value,
            "writer": self.writer,
            "output_path": str(self.output_path),
            "template_path": str(self.template_path) if self.template_path else None,
            "requires_template": self.requires_template,
        }


class ReportGenerator:
    """Route a :class:`ReportSpecification` to exactly one output writer."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def plan(self, specification: ReportSpecification) -> ReportGenerationPlan:
        specification.validate()
        return ReportGenerationPlan(
            output_format=specification.output_format,
            writer=specification.output_format.writer_name,
            output_path=specification.output_path,
            template_path=specification.template_path,
            requires_template=specification.output_format.requires_template,
        )

    def generate(
        self,
        specification: ReportSpecification,
        *,
        analysis: Any = None,
        content: ReportContent | None = None,
        ai: Any = None,
        performance: Any = None,
        monitor: Any = None,
    ) -> ReportGenerationResult:
        plan = self.plan(specification)
        if plan.output_format is OutputFormat.XLSX:
            return self._generate_xlsx(
                specification,
                analysis=analysis,
                ai=ai,
                performance=performance,
                monitor=monitor,
            )
        if plan.output_format is OutputFormat.DOCX:
            return self._generate_docx(specification, analysis=analysis, content=content)
        raise ReportFormatError(
            f"نویسنده‌ای برای {plan.output_format.label_fa} تعریف نشده است."
        )

    def _generate_xlsx(
        self,
        specification: ReportSpecification,
        *,
        analysis: Any,
        ai: Any,
        performance: Any,
        monitor: Any,
    ) -> ReportGenerationResult:
        if analysis is None:
            raise ReportFormatError("برای خروجی Excel، نتیجه تحلیل معتبر لازم است.")
        if ai is None:
            from local_office_ai.ai.analyzer import AIAnalyzer

            ai = AIAnalyzer(self.settings).analyze(analysis, use_ai=False)
        from local_office_ai.excel.report_writer import ExcelReportWriter, ReportContext

        writer = ExcelReportWriter(self.settings)
        output = writer.build(
            ReportContext(
                analysis=analysis,
                ai=ai,
                performance=performance,
                output_path=specification.output_path,
                monitor=monitor,
                profile=specification.profile,
            )
        )
        return ReportGenerationResult(
            output_path=Path(output),
            output_format=OutputFormat.XLSX,
            writer=OutputFormat.XLSX.writer_name,
            content_validated=True,
            evidence=specification.evidence,
        )

    def _generate_docx(
        self,
        specification: ReportSpecification,
        *,
        analysis: Any,
        content: ReportContent | None,
    ) -> ReportGenerationResult:
        if content is None:
            if analysis is None:
                raise ReportFormatError("برای تولید گزارش Word، نتیجه تحلیل معتبر لازم است.")
            content = ReportContent.from_analysis(
                analysis,
                request=specification.request,
            )
        content.ensure_valid()
        result = DocxReportWriter().build(
            content,
            specification.output_path,
            template_path=specification.template_path,
            strip_static_figures=not specification.preserve_template_figures,
        )
        template = (
            inspect_docx_template(specification.template_path)
            if specification.template_path is not None
            else None
        )
        return ReportGenerationResult(
            output_path=result.output_path,
            output_format=OutputFormat.DOCX,
            writer=OutputFormat.DOCX.writer_name,
            template=template,
            placeholders_resolved=result.placeholders_resolved,
            unresolved_placeholders=result.unresolved_placeholders,
            warnings=list(result.warnings),
            notes=list(result.notes),
            content_validated=True,
            evidence=specification.evidence,
            mode=result.mode,
            sections_written=result.sections_written,
            tables_written=result.tables_written,
            validation=result.validation,
        )
