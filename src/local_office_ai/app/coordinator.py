"""Application coordinator: settings, logging, monitor and console reporting."""

from __future__ import annotations

import json
import threading
from pathlib import Path
from typing import Callable

from local_office_ai.app.workflow import ExcelIntelligenceWorkflow, WorkflowOptions
from local_office_ai.core.logging_config import configure_logging, get_logger
from local_office_ai.core.models import AnalysisResult, RunResult, Severity
from local_office_ai.core.settings import Settings, load_settings
from local_office_ai.monitoring.resource_monitor import ResourceMonitor

LOGGER = get_logger(__name__)


class Coordinator:
    def __init__(
        self,
        settings: Settings | None = None,
        *,
        progress: Callable[[str], None] | None = None,
        console: bool = True,
    ) -> None:
        self.settings = settings or load_settings()
        self.log_path = configure_logging(self.settings.log_dir, self.settings.log_level)
        self.progress = progress or (print if console else (lambda message: None))
        self.abort_event = threading.Event()
        self.last_result: RunResult | None = None

    def abort(self) -> None:
        self.abort_event.set()

    # ------------------------------------------------------------------ #
    def run(
        self,
        source: Path | str,
        *,
        model: str | None = None,
        use_ai: bool = True,
        task: str = "general_analysis",
        question: str | None = None,
        goal: str | None = None,
        request: str | None = None,
        output: Path | str | None = None,
        output_format: str | None = None,
        template: Path | str | None = None,
        keep_template_figures: bool = False,
        monitor: bool | None = None,
        profile: str | None = None,
        confirm_structure: Callable[[AnalysisResult], bool] | None = None,
    ) -> RunResult:
        sampling = self.settings.enable_resource_monitoring if monitor is None else monitor
        monitor_instance = ResourceMonitor(self.settings) if sampling else _NullMonitor(self.settings)
        workflow = ExcelIntelligenceWorkflow(
            self.settings,
            monitor=monitor_instance,
            progress=self.progress,
            abort_event=self.abort_event,
        )
        result = workflow.run(
            WorkflowOptions(
                source=Path(source),
                model=model,
                use_ai=use_ai,
                task=task,
                question=question,
                goal=goal or question,
                request=request or goal or question,
                output_path=Path(output) if output else None,
                output_format=output_format,
                template=Path(template) if template else None,
                keep_template_figures=keep_template_figures,
                confirm_structure=confirm_structure,
                report_profile=profile or self.settings.report_profile or "clean",
            )
        )
        self.last_result = result
        self.print_summary(result)
        return result

    # ------------------------------------------------------------------ #
    def print_summary(self, result: RunResult) -> None:
        analysis = result.analysis
        if analysis is None:
            return
        lines = ["", "=" * 62, "خلاصه نتیجه", "=" * 62]
        if analysis.loans:
            for name in ("loan_count", "principal_sum", "installments_sum", "obligations_sum", "contract_repayment_sum"):
                calculation = analysis.calculation(name)
                if calculation is None:
                    continue
                lines.append(f"{calculation.label}: {calculation.value}")
            for audit in analysis.audits:
                if audit.status.value in {"mismatch", "not_verifiable"}:
                    lines.append(f"حسابرسی {audit.label}: {audit.status.value} (اختلاف {audit.difference})")
            for item in analysis.reconciliation:
                if item.status.value in {"outside_loan_table", "probable"}:
                    lines.append(f"تطبیق {item.institution}: {item.status.value}")
        elif analysis.generic is not None:
            generic = analysis.generic
            quality = generic.quality
            lines.append(f"جدول: {generic.table_sheet}!{generic.table_address}")
            lines.append(f"ردیف: {generic.row_count} | ستون: {generic.column_count}")
            if quality:
                lines.append(f"سلول خالی: {quality.missing_ratio:.1%} | ردیف تکراری: {quality.duplicate_rows}")
            if generic.outliers:
                lines.append(f"داده پرت علامت‌گذاری‌شده: {len(generic.outliers)}")
            numeric = [column for column in generic.columns if column.is_numeric and not column.is_identifier][:3]
            for column in numeric:
                if column.mean is not None:
                    lines.append(f"میانگین {column.header}: {column.mean}")
        answers = getattr(analysis, "request_answers", None) or []
        for answer in answers[:5]:
            lines.append(f"{answer.title}: {answer.value or answer.text[:90]}")
        critical = [item for item in analysis.anomalies if item.severity is Severity.CRITICAL]
        if critical:
            lines.append(f"{len(critical)} مورد بحرانی نیازمند بررسی است.")
        lines.append(f"وضعیت: {result.status}")
        lines.append(f"خروجی: {result.output_path}")
        lines.append(f"لاگ: {self.log_path}")
        if result.errors:
            lines.extend(result.errors)
        self.progress("\n".join(lines))

    def export_json(self, result: RunResult, path: Path | str) -> Path:
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(
            json.dumps(result.to_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return target


class _NullMonitor(ResourceMonitor):
    """Monitor used when sampling is explicitly disabled (tests, CI)."""

    def start(self) -> None:  # noqa: D102 - intentional no-op
        return None

    def stop(self) -> float:  # noqa: D102
        return 0.0
