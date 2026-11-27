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
from local_office_ai.excel.period_integration import PeriodIntegration
from local_office_ai.excel.report_writer import ExcelReportWriter

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
    warnings: list[str] = field(default_factory=list)
    status: str = "completed"

    def to_dict(self) -> dict[str, Any]:
        return {
            "folder": self.folder,
            "request": self.request,
            "output_path": self.output_path,
            "period": self.period,
            "metrics": list(self.metrics),
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

        output_path = Path(output) if output else self._default_output(source)
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
        self._folder_issues(workbook, result)
        workbook.properties.title = "Local Office AI - تحلیل پوشه"
        workbook.properties.creator = "Local Office AI Engine"
        workbook.calculation.fullCalcOnLoad = True

        output = Path(output_path)
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
        self._header(ws, row, ["شاخص"] + files + ["جمع فایل‌ها", "سهم بزرگ‌ترین فایل"])
        header_row = row
        row += 1
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
            total = sum(totals.values())
            self._value_cell(ws, row, 2 + len(files), total)
            top = max(totals.values())
            share = (top / total * 100) if total else None
            share_cell = ws.cell(row, 3 + len(files), to_excel_number(share))
            share_cell.number_format = PERCENT_FORMAT
            row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 3 + len(files))
        self._widths(ws, {"A": 46, **{get_column_letter(index): 20 for index in range(2, 4 + len(files))}})
        ws.freeze_panes = "B5"
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
