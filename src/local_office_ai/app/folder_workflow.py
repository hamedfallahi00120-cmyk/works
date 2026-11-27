"""Analyse every workbook of a folder and integrate what they share.

The folder path is the "many sources" case: each file is analysed with the same
engine, then the metric totals of all files are put next to each other.  Months
that appear in more than one file are never silently merged; the report says
which file covers which month.
"""

from __future__ import annotations

import os
import zipfile
from dataclasses import dataclass, field
from decimal import Decimal
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
from local_office_ai.core.exceptions import EngineError
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import to_excel_number
from local_office_ai.core.settings import Settings
from local_office_ai.excel.duplicate_control import (
    AGGREGATION_STATUS_FA,
    AggregationStatus,
    DuplicateControlAnalyzer,
)
from local_office_ai.excel.period_integration import PeriodIntegration
from local_office_ai.excel.report_writer import ExcelReportWriter
from local_office_ai.excel.semantic_matching import match_across_sources
from local_office_ai.reporting.formats import (
    OutputFormat,
    ensure_extension_matches,
    ensure_xlsx_delivery,
)

LOGGER = get_logger(__name__)

FONT_NAME = "Tahoma"
DARK = "1F1F1F"
GREEN = "E2EFDA"
RED = "FCE4D6"
YELLOW = "FFF2CC"
GREY = "F2F2F2"
COUNT_FORMAT = "#,##0;[Red]-#,##0"
PERCENT_FORMAT = "0.00%;[Red]-0.00%"

_SKIP_DIRS = {"output", "temp", "processed", "logs"}


@dataclass
class FileOutcome:
    path: str
    name: str
    sheets: int = 0
    rows: int = 0
    columns: int = 0
    numeric_columns: int = 0
    metrics: int = 0
    months: str = ""
    periods: str = ""
    missing_ratio: str = ""
    issues: int = 0
    status: str = ""
    error: str = ""
    analysis: Any = None
    integration: Any = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "name": self.name,
            "sheets": self.sheets,
            "rows": self.rows,
            "columns": self.columns,
            "numeric_columns": self.numeric_columns,
            "metrics": self.metrics,
            "months": self.months,
            "periods": self.periods,
            "missing_ratio": self.missing_ratio,
            "issues": self.issues,
            "status": self.status,
            "error": self.error,
        }


@dataclass
class FolderResult:
    folder: str
    request: str = ""
    output_path: str = ""
    files: list[FileOutcome] = field(default_factory=list)
    metrics: list[str] = field(default_factory=list)
    period: str = ""
    matching: Any = None
    duplicate_control: Any = None
    warnings: list[str] = field(default_factory=list)
    status: str = "completed"

    def to_dict(self) -> dict[str, Any]:
        return {
            "folder": self.folder,
            "request": self.request,
            "output_path": self.output_path,
            "period": self.period,
            "metrics": list(self.metrics),
            "matching": (
                self.matching.to_dict()
                if hasattr(self.matching, "to_dict")
                else self.matching
            ),
            "duplicate_control": (
                self.duplicate_control.to_dict()
                if hasattr(self.duplicate_control, "to_dict")
                else self.duplicate_control
            ),
            "files": [item.to_dict() for item in self.files],
            "warnings": list(self.warnings),
            "status": self.status,
        }


class FolderWorkflow:
    def __init__(
        self,
        settings: Settings,
        *,
        progress: Callable[[str], None] | None = None,
        abort_event: Any = None,
    ) -> None:
        self.settings = settings
        self.progress = progress or (lambda message: LOGGER.info(message))
        self.abort_event = abort_event

    # ------------------------------------------------------------------ #
    def run(
        self,
        folder: Path | str,
        *,
        request: str | None = None,
        model: str | None = None,
        use_ai: bool = True,
        output: Path | str | None = None,
        output_format: str | None = None,
        profile: str = "clean",
        selected: Sequence[Path | str] | None = None,
    ) -> FolderResult:
        source = Path(folder).expanduser().resolve()
        if not source.is_dir():
            raise EngineError(f"پوشه پیدا نشد: {source}")
        files = self.scan(source)
        if selected is not None:
            allowed = {Path(item).expanduser().resolve() for item in selected}
            files = [path for path in files if path.resolve() in allowed]
        result = FolderResult(folder=str(source), request=(request or "").strip())
        if not files:
            result.status = "no_files"
            result.warnings.append("در این پوشه فایل Excel قابل پردازش پیدا نشد.")
            return result

        workflows = ExcelIntelligenceWorkflow(
            self.settings, progress=self.progress, abort_event=self.abort_event
        )
        for path in files:
            if self.abort_event is not None and self.abort_event.is_set():
                break
            self.progress(f"تحلیل فایل: {path.name}")
            outcome = FileOutcome(path=str(path), name=path.name)
            try:
                analysis = workflows.analyze_only(path, request=request)
                outcome.analysis = analysis
                outcome.integration = getattr(analysis, "integration", None)
                outcome.sheets = len(analysis.meta.sheets)
                generic = getattr(analysis, "generic", None)
                if generic is not None:
                    outcome.rows = generic.row_count
                    outcome.columns = generic.column_count
                    outcome.numeric_columns = sum(
                        1
                        for column in generic.columns
                        if column.is_numeric
                        and not column.is_identifier
                        and not column.is_temporal_key
                    )
                    if generic.quality is not None:
                        outcome.missing_ratio = f"{generic.quality.missing_ratio:.1%}"
                if outcome.integration is not None:
                    outcome.metrics = len(outcome.integration.series)
                    outcome.months = "، ".join(outcome.integration.months)
                    outcome.periods = "، ".join(outcome.integration.periods)
                outcome.issues = sum(
                    1
                    for item in analysis.anomalies
                    if item.severity.value in {"critical", "warning"}
                )
                outcome.status = "تحلیل شد"
            except Exception as exc:  # noqa: BLE001 - one bad file must not stop the folder
                outcome.status = "خطا"
                outcome.error = f"{type(exc).__name__}: {exc}"
                result.warnings.append(f"{path.name}: {outcome.error}")
                LOGGER.warning("folder file failed: %s (%s)", path, exc)
            result.files.append(outcome)

        usable = [item for item in result.files if item.integration is not None]
        merged = self._merge(usable)
        result.metrics = merged.metric_names if merged is not None else []
        result.period = merged.main_period() if merged is not None else ""
        result.matching = match_across_sources(
            self.settings,
            [
                (item.name, item.analysis)
                for item in result.files
                if item.analysis is not None
            ],
        )
        result.duplicate_control = DuplicateControlAnalyzer(self.settings).analyze_sources(
            [
                (item.name, item.analysis)
                for item in result.files
                if item.analysis is not None
            ],
            matching=result.matching,
        )

        output_path = ensure_xlsx_delivery(
            output_path=output,
            default_output_path=self._default_output(source),
            explicit_format=output_format,
            request_text=request,
            context="گزارش پوشه",
        )
        writer = FolderReportWriter(self.settings)
        result.output_path = str(writer.build_folder(result, merged, output_path))
        result.status = "completed" if all(item.status == "تحلیل شد" for item in result.files) else "completed_with_warnings"
        return result

    # ------------------------------------------------------------------ #
    def scan(self, folder: Path, *, recursive: bool = True) -> list[Path]:
        root = Path(folder).resolve()
        iterator = root.rglob("*") if recursive else root.iterdir()
        found: list[Path] = []
        for path in iterator:
            try:
                if path.is_symlink() or not path.is_file():
                    continue
                relative = path.relative_to(root)
                if any(part.lower() in _SKIP_DIRS for part in relative.parts[:-1]):
                    continue
                if path.name.startswith(("~$", ".")) or path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    continue
                if not self._is_workbook(path):
                    continue
                found.append(path)
            except (OSError, ValueError):
                continue
        return sorted(found, key=lambda item: str(item).casefold())

    @staticmethod
    def _is_workbook(path: Path) -> bool:
        try:
            with zipfile.ZipFile(path) as archive:
                return "xl/workbook.xml" in archive.namelist()
        except (zipfile.BadZipFile, OSError):
            return False

    def _default_output(self, folder: Path) -> Path:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = self.settings.output_dir / f"{folder.name}_Folder_Analysis_{stamp}.xlsx"
        return target

    @staticmethod
    def _merge(outcomes: list[FileOutcome]) -> PeriodIntegration | None:
        """Rename the source of every cell to ``file::sheet`` before merging."""

        from local_office_ai.excel.period_integration import (
            JALALI_MONTHS,
            MetricSeries,
            PeriodIntegration,
            PeriodTotal,
            _ordered_periods,
        )

        if not outcomes:
            return None
        merged = PeriodIntegration(source_kind="folder")
        buckets: dict[str, MetricSeries] = {}
        months: set[str] = set()
        periods: set[str] = set()
        month_owner: dict[str, str] = {}
        duplicates: list[str] = []
        for outcome in outcomes:
            integration = outcome.integration
            if integration is None:
                continue
            prefix = outcome.name
            for bucket in integration.series:
                target = buckets.setdefault(bucket.metric, MetricSeries(metric=bucket.metric))
                for cell in bucket.cells:
                    target.cells.append(
                        PeriodTotal(
                            metric=cell.metric,
                            period=cell.period,
                            month=cell.month,
                            month_index=cell.month_index,
                            source=f"{prefix}::{cell.source}",
                            total=cell.total,
                            observations=cell.observations,
                            complete=cell.complete,
                            total_address=cell.total_address,
                            declared_total=cell.declared_total,
                            verification=cell.verification,
                            note=cell.note,
                        )
                    )
                months.update(bucket.months())
                periods.update(bucket.periods())
            for month in integration.months:
                owner = month_owner.get(month)
                if owner is not None:
                    duplicates.append(f"{month} در «{owner}» و «{prefix}»")
                else:
                    month_owner[month] = prefix
            merged.excluded_sources.extend(
                f"{prefix}::{name}" for name in integration.excluded_sources
            )
            merged.verified_cells += integration.verified_cells
            merged.mismatched_cells += integration.mismatched_cells
            for warning in integration.warnings:
                if warning not in merged.warnings:
                    merged.warnings.append(warning)
        if duplicates:
            merged.warnings.append(
                "ماه تکراری بین فایل‌ها پیدا شد و در ماتریس ماهانه با هم جمع نشد: "
                + "؛ ".join(duplicates[:10])
            )
        merged.series = list(buckets.values())
        merged.months = [name for name in JALALI_MONTHS if name in months]
        merged.periods = _ordered_periods(periods)
        merged.sources = [outcome.name for outcome in outcomes]
        return merged if merged.series else None


class FolderReportWriter(ExcelReportWriter):
    """Folder-level delivery: one clean workbook for many files."""

    # ------------------------------------------------------------------ #
    def build_folder(
        self,
        result: FolderResult,
        merged: PeriodIntegration | None,
        output_path: Path,
    ) -> Path:
        workbook = Workbook()
        workbook.remove(workbook.active)
        self._folder_summary(workbook, result, merged)
        self._file_list(workbook, result)
        if merged is not None:
            self._file_comparison(workbook, result, merged)
        self._cross_source_matching(workbook, result)
        self._duplicate_control(workbook, result)
        self._folder_issues(workbook, result)
        workbook.properties.title = "Local Office AI - تحلیل پوشه"
        workbook.properties.creator = "Local Office AI Engine"
        workbook.calculation.fullCalcOnLoad = True

        output = ensure_extension_matches(Path(output_path), OutputFormat.XLSX)
        output.parent.mkdir(parents=True, exist_ok=True)
        temp = output.with_suffix(".tmp.xlsx")
        try:
            workbook.save(temp)
            self._verify(temp)
            os.replace(temp, output)
        except Exception:
            if temp.exists():
                temp.unlink(missing_ok=True)
            raise
        return output

    # ------------------------------------------------------------------ #
    def _folder_summary(
        self,
        wb: Workbook,
        result: FolderResult,
        merged: PeriodIntegration | None = None,
    ) -> None:
        ws = wb.create_sheet("خلاصه پوشه")
        self._rtl(ws)
        self._title(ws, "خلاصه تحلیل پوشه", 4)
        row = 3
        self._header(ws, row, ["موضوع", "مقدار"])
        row += 1
        analyzed = sum(1 for item in result.files if item.status == "تحلیل شد")
        failed = len(result.files) - analyzed
        info = [
            ("پوشه", result.folder),
            ("تعداد فایل", len(result.files)),
            ("فایل‌های تحلیل‌شده", analyzed),
            ("فایل‌های ناموفق", failed),
            ("تعداد شاخص مشترک", len(result.metrics)),
            ("تطبیق معنایی بین منابع", self._matching_summary(result)),
            ("دوره تحلیل", result.period or "—"),
            ("وضعیت", "کامل" if failed == 0 else "با هشدار"),
            ("درخواست کاربر", result.request or "—"),
            ("تفسیر درخواست", self._request_reading(result) or "—"),
            ("پاسخ درخواست", self._folder_answer(result, merged) or "—"),
        ]
        for label, value in info:
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, 4, row - 1, 1, 2)
        if result.warnings:
            row += 2
            self._section(ws, row, "هشدارها", 4)
            for warning in result.warnings[:20]:
                row += 1
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
                ws.cell(row, 1, "• " + warning).alignment = Alignment(horizontal="right", wrap_text=True)
        self._widths(ws, {"A": 34, "B": 90, "C": 24, "D": 24})
        self._print_setup(ws)

    @staticmethod
    def _matching_summary(result: FolderResult) -> str:
        report = result.matching
        if report is None:
            return "—"
        counts = report.counts()
        return (
            f"فیلد {counts.get('fields', 0)} — تطبیق قطعی {counts.get('matched', 0)} — "
            f"احتمالی {counts.get('probable', 0)} — مبهم {counts.get('ambiguous', 0)} — "
            f"رد‌شده {counts.get('rejected', 0)}"
        )

    @staticmethod
    def _request_reading(result: FolderResult) -> str:
        for item in result.files:
            intent = getattr(getattr(item, "analysis", None), "request_intent", None)
            if intent is None:
                continue
            parts = []
            labels = getattr(intent, "scope_labels", None) or []
            if labels:
                parts.append("نوع تحلیل: " + "، ".join(labels))
            metrics = list(getattr(intent, "matched_metrics", None) or [])
            if not metrics:
                metrics = list(getattr(intent, "requested_metrics", None) or [])
            parts.append("شاخص‌ها: " + ("، ".join(metrics[:6]) if metrics else "همه شاخص‌های عددی"))
            years = list(getattr(intent, "years", None) or [])
            if years:
                parts.append("سال‌ها: " + "، ".join(years))
            return "؛ ".join(parts) + "."
        return ""

    @staticmethod
    def _folder_answer(result: FolderResult, merged: PeriodIntegration | None) -> str:
        if merged is None:
            return ""
        usable = [item for item in result.files if item.integration is not None]
        if len(usable) < 2:
            return ""
        period = merged.main_period()
        for series in merged.series:
            totals: dict[str, Decimal] = {}
            for cell in series.cells:
                if cell.period != period or cell.total is None or not cell.complete:
                    continue
                name = cell.source.split("::", 1)[0]
                totals[name] = (totals.get(name) or Decimal(0)) + cell.total
            if len(totals) < 2:
                continue
            top = max(totals, key=lambda key: totals[key])
            low = min(totals, key=lambda key: totals[key])
            report = result.duplicate_control
            status = (
                report.cross_source_status(series.metric, totals)
                if report is not None and len(totals) > 1
                else None
            )
            if len(totals) > 1 and status is not AggregationStatus.ALLOWED:
                return (
                    f"شاخص «{series.metric}» در دوره {period}: فایل «{top}» با {totals[top]:,.0f} "
                    f"بیشترین و فایل «{low}» با {totals[low]:,.0f} کمترین مقدار را دارد؛ "
                    "جمع بین فایل‌ها به دلیل احتمال شمارش دوباره ارائه نشد."
                )
            total = sum(totals.values(), Decimal(0))
            if not total:
                continue
            share = totals[top] / total * 100
            return (
                f"شاخص «{series.metric}» در دوره {period}: فایل «{top}» با {totals[top]:,.0f} "
                f"بیشترین و فایل «{low}» با {totals[low]:,.0f} کمترین مقدار را دارد؛ "
                f"سهم فایل بزرگ {share:.2f}٪ از جمع پوشه ({total:,.0f})."
            )
        return ""

    def _file_list(self, wb: Workbook, result: FolderResult) -> None:
        ws = wb.create_sheet("فهرست فایل‌ها")
        self._rtl(ws)
        self._title(ws, "فهرست فایل‌های پوشه و خلاصه هر فایل", 10)
        self._header(
            ws,
            3,
            ["فایل", "برگه", "ردیف", "ستون", "ستون عددی", "شاخص", "ماه‌ها", "دوره‌ها", "سلول خالی", "وضعیت"],
        )
        row = 4
        for item in result.files:
            self._value_cell(ws, row, 1, item.name)
            self._value_cell(ws, row, 2, item.sheets)
            self._value_cell(ws, row, 3, item.rows)
            self._value_cell(ws, row, 4, item.columns)
            self._value_cell(ws, row, 5, item.numeric_columns)
            self._value_cell(ws, row, 6, item.metrics)
            self._value_cell(ws, row, 7, item.months or "—")
            self._value_cell(ws, row, 8, item.periods or "—")
            self._value_cell(ws, row, 9, item.missing_ratio or "—")
            self._value_cell(ws, row, 10, item.status)
            self._fill(ws.cell(row, 10), GREEN if item.status == "تحلیل شد" else RED)
            row += 1
        self._body(ws, 4, row - 1, 1, 10)
        self._widths(ws, {"A": 40, "B": 8, "C": 10, "D": 10, "E": 12, "F": 10, "G": 50, "H": 20, "I": 12, "J": 16})
        ws.freeze_panes = "A4"
        self._print_setup(ws)

    def _file_comparison(self, wb: Workbook, result: FolderResult, merged: PeriodIntegration) -> None:
        ws = wb.create_sheet("مقایسه بین فایل‌ها")
        self._rtl(ws)
        self._title(ws, "مقایسه شاخص‌های مشترک بین فایل‌های پوشه", 9)
        period = merged.main_period()
        files = [item.name for item in result.files if item.integration is not None]
        row = 3
        self._section(ws, row, f"جمع ماه‌های کامل دوره {period} به تفکیک فایل", 8 + len(files))
        row += 1
        self._header(
            ws,
            row,
            ["شاخص"] + files + ["جمع فایل‌ها", "سهم بزرگ‌ترین فایل", "وضعیت تجمیع"],
        )
        header_row = row
        row += 1
        report = result.duplicate_control
        for series in merged.series:
            totals: dict[str, Any] = {}
            for cell in series.cells:
                if cell.period != period or cell.total is None or not cell.complete:
                    continue
                name = cell.source.split("::", 1)[0]
                totals[name] = (totals.get(name) or 0) + cell.total
            if not totals:
                continue
            self._value_cell(ws, row, 1, series.metric)
            for offset, name in enumerate(files, start=2):
                value = totals.get(name)
                cell = ws.cell(row, offset, to_excel_number(value))
                cell.number_format = COUNT_FORMAT
            contributors = [name for name in files if totals.get(name) is not None]
            status = (
                report.cross_source_status(series.metric, contributors)
                if report is not None and len(contributors) > 1
                else None
            )
            safe = len(contributors) <= 1 or status is AggregationStatus.ALLOWED
            if safe:
                total = sum(totals.values())
                self._value_cell(ws, row, 2 + len(files), total)
                top = max(totals.values())
                share = (top / total * 100) if total else None
                share_cell = ws.cell(row, 3 + len(files), to_excel_number(share))
                share_cell.number_format = PERCENT_FORMAT
            else:
                self._value_cell(ws, row, 2 + len(files), "—")
                self._value_cell(ws, row, 3 + len(files), "—")
            status_cell = ws.cell(row, 4 + len(files))
            if safe:
                status_cell.value = "مجاز" if len(contributors) > 1 else "یک منبع"
                self._fill(status_cell, GREEN)
            else:
                status_cell.value = (
                    AGGREGATION_STATUS_FA.get(status, "نیازمند بررسی")
                    if status is not None
                    else "نیازمند بررسی"
                )
                self._fill(
                    status_cell,
                    RED if status is AggregationStatus.BLOCKED else YELLOW,
                )
            row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 4 + len(files))
        self._widths(ws, {"A": 46, **{get_column_letter(index): 20 for index in range(2, 5 + len(files))}})
        ws.freeze_panes = "B5"
        self._print_setup(ws)

    def _cross_source_matching(self, wb: Workbook, result: FolderResult) -> None:
        """Explain which fields of which files were matched, and why not."""
        from local_office_ai.excel.semantic_matching import (
            PERIOD_STATUS_FA,
            RELATION_FA,
            STATUS_FA,
            UNIT_STATE_FA,
        )

        report = result.matching
        if report is None:
            return
        counts = report.counts()
        if not any(
            counts.get(key, 0)
            for key in ("matched", "probable", "ambiguous", "rejected")
        ):
            return

        ws = wb.create_sheet("تطبیق شاخص‌ها بین منابع")
        self._rtl(ws)
        self._title(ws, "تطبیق معنایی شاخص‌ها بین فایل‌های پوشه: شاهد و دلیل تصمیم", 10)
        row = 3
        self._value_cell(ws, row, 1, "خلاصه تصمیم‌ها")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(ws, row, 2, self._matching_summary(result))
        row += 1
        self._value_cell(ws, row, 1, "اصل کار")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(
            ws,
            row,
            2,
            "تطبیق فقط با شاهد نام نرمال‌شده، نقش معنایی، واحد و دوره گزارش‌دهی انجام شده است؛ "
            "این برگه هیچ مقداری را بین فایل‌ها جمع نمی‌کند.",
        )
        row += 2
        if report.warnings:
            self._section(ws, row, "هشدارهای تطبیق", 9)
            row += 1
            for warning in report.warnings:
                self._value_cell(ws, row, 1, "هشدار")
                ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
                self._value_cell(ws, row, 2, warning)
                self._fill(ws.cell(row, 1), YELLOW)
                row += 1
            row += 1

        colors = {"matched": GREEN, "probable": YELLOW, "ambiguous": YELLOW, "rejected": RED}

        def section(title: str, items: list[Any]) -> None:
            nonlocal row
            if not items:
                return
            self._section(ws, row, title, 10)
            row += 1
            self._header(ws, row, [
                "وضعیت", "نوع", "برچسب اصلی ← نرمال‌شده", "منبع/برگه/بلوک/عنوان",
                "رابطه دوره", "واحد", "قابل مقایسه", "اطمینان", "شاهد/دلیل", "توضیح دوره",
            ])
            header_row = row
            row += 1
            for item in items:
                members = "؛ ".join(member.display_label for member in item.members)
                original_labels = " / ".join(
                    dict.fromkeys(member.header for member in item.members)
                )
                display_label = (
                    f"{original_labels} ← {item.normalized_label}"
                    if item.normalized_label
                    else original_labels
                )
                values = [
                    STATUS_FA.get(item.status, item.status.value),
                    item.kind_label,
                    display_label or "—",
                    members,
                    f"{RELATION_FA.get(item.period_relation, item.period_relation or '—')} / "
                    f"{PERIOD_STATUS_FA.get(item.period_status, item.period_status or '—')}",
                    UNIT_STATE_FA.get(item.unit_state, item.unit_state or "—"),
                    "بله" if item.comparable else "خیر",
                    f"{item.confidence:.0%}",
                    (
                        "؛ ".join(item.evidence)
                        + (f" — {item.reason}" if item.reason else "")
                    )
                    or "—",
                    item.period_explanation or "—",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                self._fill(ws.cell(row, 1), colors.get(item.status.value, GREY))
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 10)
            row += 1

        section("تطبیق قطعی بین فایل‌ها", report.matches)
        section("تطبیق احتمالی (برابری کامل اثبات نشد)", report.probable)

        if report.ambiguities:
            self._section(ws, row, "موارد مبهم: گزینه‌ها نگه داشته شده‌اند", 9)
            row += 1
            self._header(ws, row, [
                "فیلد مبنا", "منبع/برگه", "تعداد گزینه", "گزینه‌ها (منبع/برگه/عنوان)",
                "وضعیت گزینه", "امتیاز", "دلیل", "شاهد", "وضعیت دوره",
            ])
            header_row = row
            row += 1
            for item in report.ambiguities:
                anchor = item.members[0]
                candidates = "؛ ".join(
                    f"{candidate.field.source}/{candidate.field.sheet}/«{candidate.field.header}»"
                    f" ({candidate.score:.0%})"
                    for candidate in item.candidates
                )
                values = [
                    anchor.header,
                    f"{anchor.source}/{anchor.sheet}",
                    len(item.candidates),
                    candidates or "—",
                    "؛ ".join(
                        STATUS_FA.get(candidate.status, candidate.status.value)
                        for candidate in item.candidates
                    ) or "—",
                    f"{item.confidence:.0%}",
                    item.reason,
                    "؛ ".join(
                        text for candidate in item.candidates for text in candidate.evidence
                    ) or "—",
                    PERIOD_STATUS_FA.get(item.period_status, item.period_status or "—"),
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                self._fill(ws.cell(row, 1), YELLOW)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 9)
            row += 1

        if report.rejections:
            self._section(ws, row, "موارد رد‌شده: چرا دو فیلد هم‌معنا نیستند", 9)
            row += 1
            self._header(ws, row, [
                "فیلد اول", "منبع/برگه اول", "فیلد دوم", "منبع/برگه دوم",
                "دلیل رد", "امتیاز برچسب", "رابطه دوره", "واحد", "وضعیت دوره",
            ])
            header_row = row
            row += 1
            for item in report.rejections:
                values = [
                    item.left.header,
                    f"{item.left.source}/{item.left.sheet}",
                    item.right.header,
                    f"{item.right.source}/{item.right.sheet}",
                    item.reason,
                    f"{item.label_score:.0%}",
                    RELATION_FA.get(item.period_relation, item.period_relation or "—"),
                    UNIT_STATE_FA.get(item.unit_state, item.unit_state or "—"),
                    PERIOD_STATUS_FA.get(item.period_status, item.period_status or "—"),
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                self._fill(ws.cell(row, 1), RED)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 9)

        self._widths(ws, {
            "A": 22, "B": 20, "C": 26, "D": 48, "E": 16,
            "F": 14, "G": 16, "H": 11, "I": 62, "J": 48,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    def _folder_issues(self, wb: Workbook, result: FolderResult) -> None:
        ws = wb.create_sheet("موارد نیازمند بررسی")
        self._rtl(ws)
        self._title(ws, "موارد نیازمند بررسی در فایل‌های پوشه", 5)
        self._header(ws, 3, ["فایل", "شدت", "دسته", "پیام", "محل"])
        row = 4
        for item in result.files:
            analysis = item.analysis
            if analysis is None:
                self._value_cell(ws, row, 1, item.name)
                self._value_cell(ws, row, 2, "بحرانی")
                self._value_cell(ws, row, 3, "file-error")
                self._value_cell(ws, row, 4, item.error)
                row += 1
                continue
            for anomaly in analysis.anomalies:
                if anomaly.severity.value == "info":
                    continue
                self._value_cell(ws, row, 1, item.name)
                self._value_cell(ws, row, 2, "بحرانی" if anomaly.severity.value == "critical" else "هشدار")
                self._value_cell(ws, row, 3, anomaly.category)
                self._value_cell(ws, row, 4, anomaly.message)
                self._value_cell(ws, row, 5, f"{anomaly.sheet}!{anomaly.address}" if anomaly.sheet else "—")
                self._fill(ws.cell(row, 2), RED if anomaly.severity.value == "critical" else YELLOW)
                row += 1
        if row == 4:
            ws.cell(4, 1, "مورد نیازمند بررسی ثبت نشد.")
            row += 1
        self._body(ws, 4, row - 1, 1, 5)
        self._widths(ws, {"A": 40, "B": 12, "C": 24, "D": 90, "E": 24})
        self._print_setup(ws)

    def _duplicate_control(self, wb: Workbook, result: FolderResult) -> None:
        """Record duplicates, derived sheets and blocked aggregation across files."""
        from local_office_ai.excel.duplicate_control import (
            AGGREGATION_STATUS_FA,
            CHECK_STATUS_FA,
            SHEET_ROLE_FA,
        )

        report = result.duplicate_control
        if report is None:
            return
        ws = wb.create_sheet("کنترل تکرار و تجمیع بین فایل‌ها")
        self._rtl(ws)
        self._title(ws, "کنترل تکرار، برگه مشتق و تجمیع ایمن بین فایل‌های پوشه", 10)
        row = 3
        counts = report.counts()
        roles = report.role_counts()
        self._value_cell(ws, row, 1, "خلاصه")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(
            ws,
            row,
            2,
            f"فایل {counts['sources']} — برگه {counts['sheets']} — گروه محتوای تکراری "
            f"{counts['duplicate_groups']} — رکورد تکراری {counts['record_duplicates']} — "
            f"تصمیم تجمیع: مجاز {counts['allowed']} / نیازمند بررسی {counts['review']} / "
            f"مسدود {counts['blocked']} — کنارگذاشته {counts['exclusions']}",
        )
        row += 1
        self._value_cell(ws, row, 1, "طبقه‌بندی برگه‌ها")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(
            ws,
            row,
            2,
            f"داده اولیه {roles.get('primary', 0)} / خلاصه {roles.get('summary', 0)} / "
            f"مشتق {roles.get('derived', 0)} / تکراری {roles.get('duplicate', 0)} / "
            f"نامشخص {roles.get('unknown', 0)}",
        )
        row += 1
        self._value_cell(ws, row, 1, "اصل کار")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(
            ws,
            row,
            2,
            "هیچ رکورد یا فایلی حذف یا اصلاح نشد. جمع بین فایل‌ها فقط با تأیید شش شرط "
            "(هم‌ارزی شاخص، دوره، واحد، رابطه دسته، استقلال منبع و جدا بودن پوشش) مجاز است؛ "
            "موارد مبهم برای بررسی علامت خورده‌اند.",
        )
        row += 2
        if report.warnings:
            self._section(ws, row, "هشدارها", 9)
            row += 1
            for warning in report.warnings:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
                self._value_cell(ws, row, 1, "• " + warning)
                self._fill(ws.cell(row, 1), YELLOW)
                row += 1
            row += 1

        if report.sheets:
            self._section(ws, row, "نقش برگه‌ها و شاهد طبقه‌بندی", 9)
            row += 1
            self._header(ws, row, ["فایل", "برگه", "نقش", "اطمینان", "فرمول/ارجاع", "تکراری از", "شاهد", "", ""])
            header_row = row
            row += 1
            colors = {"primary": GREEN, "summary": YELLOW, "derived": YELLOW, "duplicate": RED}
            for sheet in report.sheets:
                values = [
                    sheet.source,
                    sheet.sheet,
                    SHEET_ROLE_FA.get(sheet.role, sheet.role.value),
                    f"{sheet.confidence:.0%}",
                    f"{sheet.formula_cells}/{sheet.cross_sheet_formula_cells}",
                    "، ".join(sheet.duplicate_of) or "—",
                    "؛ ".join(sheet.evidence) or "—",
                    "",
                    "",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                self._fill(ws.cell(row, 3), colors.get(sheet.role.value, GREY))
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 7)
            row += 1

        if report.duplicate_groups:
            self._section(ws, row, "گروه‌های محتوای تکراری (جمع آن‌ها مسدود است)", 9)
            row += 1
            self._header(ws, row, ["گروه", "نوع", "اعضا", "دلیل", "شاهد", "", "", "", ""])
            header_row = row
            row += 1
            for group in report.duplicate_groups:
                values = [
                    group.group_id,
                    "جمع تکراری" if group.kind == "summary" else "جدول تکراری",
                    "؛ ".join(member.label for member in group.members),
                    group.reason,
                    "؛ ".join(group.evidence),
                    "",
                    "",
                    "",
                    "",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                self._fill(ws.cell(row, 2), RED if group.identical else YELLOW)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 5)
            row += 1

        if report.decisions:
            self._section(ws, row, "تصمیم‌های تجمیع ایمن بین منابع", 12)
            row += 1
            self._header(ws, row, [
                "شناسه", "شاخص", "اعضا", "نتیجه", "شاخص", "دوره", "واحد",
                "دسته", "استقلال", "پوشش", "دلیل", "شاهد",
            ])
            header_row = row
            row += 1
            decision_colors = {"allowed": GREEN, "review": YELLOW, "blocked": RED}
            for decision in report.decisions[:200]:
                checks = {check.name: check for check in decision.checks}
                values = [
                    decision.decision_id,
                    decision.metric,
                    "؛ ".join(decision.members),
                    AGGREGATION_STATUS_FA.get(decision.status, decision.status.value),
                    CHECK_STATUS_FA.get(checks["indicator_equivalence"].status, "—"),
                    CHECK_STATUS_FA.get(checks["period_compatibility"].status, "—"),
                    CHECK_STATUS_FA.get(checks["unit_compatibility"].status, "—"),
                    CHECK_STATUS_FA.get(checks["category_relationship"].status, "—"),
                    CHECK_STATUS_FA.get(checks["source_independence"].status, "—"),
                    CHECK_STATUS_FA.get(checks["coverage_non_overlap"].status, "—"),
                    "؛ ".join(decision.reasons) or "—",
                    "؛ ".join(decision.evidence) or "—",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                self._fill(ws.cell(row, 4), decision_colors.get(decision.status.value, GREY))
                row += 1
            if len(report.decisions) > 200:
                self._value_cell(ws, row, 1, f"… و {len(report.decisions) - 200} تصمیم دیگر")
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 12)
            row += 1

        if report.exclusions:
            self._section(ws, row, "کنارگذاشته‌ها و دلیل آن‌ها", 9)
            row += 1
            self._header(ws, row, ["نوع", "فایل", "برگه", "محل", "دلیل", "شاهد", "", "", ""])
            header_row = row
            row += 1
            for item in report.exclusions[:300]:
                values = [
                    item.kind,
                    item.source,
                    item.sheet,
                    item.address or "—",
                    item.reason,
                    "؛ ".join(item.evidence) or "—",
                    "",
                    "",
                    "",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
            if len(report.exclusions) > 300:
                self._value_cell(ws, row, 1, f"… و {len(report.exclusions) - 300} مورد دیگر")
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 6)

        self._widths(ws, {
            "A": 20, "B": 28, "C": 40, "D": 14, "E": 16, "F": 16,
            "G": 16, "H": 14, "I": 14, "J": 14, "K": 60, "L": 60,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)
