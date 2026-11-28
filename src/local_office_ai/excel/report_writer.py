"""Professional Excel output.

The writer always creates a new workbook; the source file is never opened for
writing, so it cannot be damaged.  Every sheet is built from the analyzed
model, never from model-generated numbers.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from local_office_ai.ai.analyzer import AIOutcome
from local_office_ai.core.exceptions import EngineError
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import (
    AnalysisResult,
    AuditStatus,
    ColumnRole,
    MatchStatus,
    Severity,
    TableKind,
    to_excel_number,
)
from local_office_ai.core.settings import Settings
from local_office_ai.excel.generic_report import GenericReportMixin
from local_office_ai.excel.normalization import format_jalali
from local_office_ai.excel.request_report import RequestReportMixin
from local_office_ai.monitoring.performance_report import PerformanceReport, performance_rows
from local_office_ai.reporting.formats import OutputFormat, ensure_extension_matches

LOGGER = get_logger(__name__)

FONT_NAME = "Tahoma"
NAVY = "17365D"
BLUE = "2F5597"
LIGHT_BLUE = "DEEAF6"
GREEN = "E2EFDA"
RED = "FCE4D6"
YELLOW = "FFF2CC"
GREY = "F2F2F2"
WHITE = "FFFFFF"
DARK = "1F1F1F"
THIN = Side(style="thin", color="BFBFBF")
MONEY_FORMAT = '#,##0;[Red]-#,##0'
DECIMAL_FORMAT = '#,##0.00;[Red]-#,##0.00'
PERCENT_FORMAT = '0.00%'

AUDIT_STATUS_FA = {
    AuditStatus.VERIFIED: "تأیید شد",
    AuditStatus.MISMATCH: "مغایرت",
    AuditStatus.NOT_VERIFIABLE: "قابل بررسی نیست",
    AuditStatus.NO_DECLARED_VALUE: "مقدار اعلامی ندارد",
}
MATCH_STATUS_FA = {
    MatchStatus.MATCH: "تطبیق قطعی",
    MatchStatus.PROBABLE: "تطبیق احتمالی - نیازمند تأیید",
    MatchStatus.NO_MATCH: "عدم تطبیق",
    MatchStatus.OUTSIDE_LOAN_TABLE: "تعهد خارج از جدول وام‌ها",
    MatchStatus.NO_OBLIGATION: "تعهد ماهانه پیدا نشد",
}
SEVERITY_FA = {Severity.CRITICAL: "بحرانی", Severity.WARNING: "هشدار", Severity.INFO: "اطلاع"}
TABLE_KIND_FA = {
    TableKind.LOAN: "جدول وام",
    TableKind.OBLIGATION: "جدول تعهدات",
    TableKind.TIME_SERIES: "سری زمانی",
    TableKind.KPI: "شاخص‌ها",
    TableKind.GENERIC: "عمومی",
}
ROLE_FA = {
    ColumnRole.INSTITUTION: "مؤسسه/طرف حساب",
    ColumnRole.PRINCIPAL: "اصل وام",
    ColumnRole.INSTALLMENT: "قسط",
    ColumnRole.INSTALLMENT_COUNT: "تعداد اقساط",
    ColumnRole.DATE: "تاریخ",
    ColumnRole.DESCRIPTION: "شرح",
    ColumnRole.AMOUNT: "مبلغ",
    ColumnRole.COUNT: "تعداد",
    ColumnRole.ID: "شناسه",
    ColumnRole.UNKNOWN: "نامشخص",
}


def _decimal_text(value: Decimal) -> str:
    text = f"{value:,.2f}".rstrip("0").rstrip(".")
    return text or "0"


@dataclass
class ReportContext:
    analysis: AnalysisResult
    ai: AIOutcome
    performance: PerformanceReport | None
    output_path: Path
    monitor: Any = None
    profile: str = "audit"


class ExcelReportWriter(GenericReportMixin, RequestReportMixin):
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    # ------------------------------------------------------------------ #
    def build(self, context: ReportContext) -> Path:
        workbook = Workbook()
        workbook.remove(workbook.active)
        profile = context.profile or self.settings.report_profile or "clean"

        for key in self._plan(context, profile):
            writer = self._SHEET_WRITERS[key]
            writer(self, workbook, context)

        workbook.properties.title = "Local Office AI - گزارش تحلیل Excel"
        workbook.properties.creator = "Local Office AI Engine"
        workbook.properties.description = "گزارش تولیدشده توسط موتور تحلیل Excel"
        workbook.calculation.fullCalcOnLoad = True

        output = ensure_extension_matches(context.output_path, OutputFormat.XLSX)
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
        LOGGER.info("report written: %s", output)
        return output

    # ------------------------------------------------------------------ #
    # Sheet planning
    # ------------------------------------------------------------------ #
    _SHEET_WRITERS = {
        "request": lambda self, wb, context: self._request_sheet(wb, context),
        "dashboard": lambda self, wb, context: self._dashboard(wb, context),
        "management": lambda self, wb, context: self._executive_summary(wb, context),
        "monthly": lambda self, wb, context: self._monthly_sheet(wb, context),
        "comparison": lambda self, wb, context: self._period_compare_sheet(wb, context),
        "descriptive": lambda self, wb, context: self._descriptive_stats(wb, context.analysis),
        "frequency": lambda self, wb, context: self._frequency_sheet(wb, context.analysis),
        "groups": lambda self, wb, context: self._groups_sheet(wb, context.analysis),
        "trend": lambda self, wb, context: self._trend_correlation_sheet(wb, context.analysis),
        "outliers": lambda self, wb, context: self._outliers_sheet(wb, context.analysis),
        "quality": lambda self, wb, context: self._validation(wb, context.analysis),
        "issues": lambda self, wb, context: self._anomalies(wb, context.analysis),
        "audits": lambda self, wb, context: self._audits(wb, context.analysis),
        "reconciliation": lambda self, wb, context: self._reconciliation(wb, context.analysis),
        "ai": lambda self, wb, context: self._ai_sheet(wb, context),
        # audit-only sheets
        "workbook_profile": lambda self, wb, context: self._workbook_profile(wb, context.analysis),
        "data_dictionary": lambda self, wb, context: self._data_dictionary(wb, context.analysis),
        "fields": lambda self, wb, context: self._fields_sheet(wb, context.analysis),
        "blocks": lambda self, wb, context: self._blocks_sheet(wb, context.analysis),
        "coverage": lambda self, wb, context: self._coverage_sheet(wb, context.analysis),
        "metadata_warnings": lambda self, wb, context: self._metadata_warnings_sheet(wb, context.analysis),
        "period_comparison": lambda self, wb, context: self._period_comparison_sheet(wb, context.analysis),
        "reconciliation_sheet": lambda self, wb, context: self._reconciliation_sheet(wb, context.analysis),
        "normalized": lambda self, wb, context: self._normalized_data(wb, context.analysis),
        "raw": lambda self, wb, context: self._raw_data(wb, context.analysis),
        "calculations": lambda self, wb, context: self._calculations(wb, context.analysis),
        "traceability": lambda self, wb, context: self._traceability(wb, context),
        "methodology": lambda self, wb, context: self._methodology_sheet(wb, context.analysis),
        "performance": lambda self, wb, context: self._performance(wb, context),
        "samples": lambda self, wb, context: self._samples(wb, context),
        "limitations": lambda self, wb, context: self._limitations(wb, context),
        "structure": lambda self, wb, context: self._structure(wb, context.analysis),
        "extraction": lambda self, wb, context: self._extraction_model(wb, context.analysis),
        "semantics": lambda self, wb, context: self._semantics_sheet(wb, context.analysis),
        "quality_details": lambda self, wb, context: self._quality_details_sheet(wb, context.analysis),
        "statistical_reconciliation": lambda self, wb, context: self._statistical_reconciliation_sheet(
            wb, context.analysis
        ),
        "semantic_matching": lambda self, wb, context: self._semantic_matching_sheet(
            wb, context.analysis
        ),
        "duplicate_control": lambda self, wb, context: self._duplicate_control_sheet(
            wb, context.analysis
        ),
        "quality_gate": lambda self, wb, context: self._quality_gate_sheet(
            wb, context.analysis
        ),
        "provenance": lambda self, wb, context: self._provenance_sheet(
            wb, context.analysis
        ),
    }

    def _plan(self, context: ReportContext, profile: str) -> list[str]:
        """Choose sheets from the request and the data, never from a template."""

        analysis = context.analysis
        generic = analysis.generic
        integration = getattr(analysis, "integration", None)
        intent = getattr(analysis, "request_intent", None)
        plan: list[str] = []

        if intent is not None and intent.text:
            plan.append("request")
        plan.extend(["dashboard", "management"])
        if integration is not None and getattr(integration, "is_usable", False):
            plan.append("monthly")
            if len(integration.periods) >= 2:
                plan.append("comparison")
        if generic is not None:
            if analysis.intelligence is not None:
                plan.append("data_dictionary")
            plan.append("descriptive")
            has_categorical = any(column.kind.value == "text" for column in generic.columns)
            asks_distribution = intent is not None and intent.has("distribution")
            if generic.frequencies and (has_categorical or asks_distribution):
                plan.append("frequency")
            if generic.groups or generic.crosstabs:
                plan.append("groups")
        asks_trend = intent is not None and intent.needs_generic_trend_sheet()
        if generic is not None and generic.trends and (
            asks_trend or not (integration is not None and getattr(integration, "is_usable", False))
        ):
            plan.append("trend")
        if generic is not None and generic.outliers and asks_trend:
            plan.append("outliers")
        plan.append("quality")
        plan.append("issues")
        if analysis.audits:
            plan.append("audits")
        if analysis.reconciliation:
            plan.append("reconciliation")
        if profile != "audit":
            return plan

        audit_tail = ["workbook_profile", "ai"]
        if analysis.intelligence is not None:
            audit_tail.extend(["fields", "blocks", "coverage", "metadata_warnings"])
            audit_tail.extend(["period_comparison", "reconciliation_sheet"])
        if getattr(analysis, "extraction", None) is not None:
            audit_tail.append("extraction")
        if getattr(analysis, "semantics", None) is not None:
            audit_tail.append("semantics")
        if getattr(analysis, "quality", None) is not None:
            audit_tail.append("quality_details")
        if getattr(analysis, "statistical_reconciliation", None) is not None:
            audit_tail.append("statistical_reconciliation")
        if getattr(analysis, "semantic_matching", None) is not None:
            audit_tail.append("semantic_matching")
        if getattr(analysis, "duplicate_control", None) is not None:
            audit_tail.append("duplicate_control")
        if getattr(analysis, "quality_gate", None) is not None:
            audit_tail.append("quality_gate")
        if getattr(analysis, "provenance", None) is not None:
            audit_tail.append("provenance")
        audit_tail.extend(
            [
                "normalized",
                "raw",
                "calculations",
                "traceability",
                "methodology",
                "performance",
                "samples",
                "limitations",
                "structure",
            ]
        )
        for key in audit_tail:
            if key in {"performance", "samples"} and context.performance is None:
                continue
            if key not in plan:
                plan.append(key)
        return plan

    @staticmethod
    def _verify(path: Path) -> None:
        try:
            check = load_workbook(path, read_only=True)
            names = list(check.sheetnames)
            check.close()
        except Exception as exc:  # noqa: BLE001
            raise EngineError(f"خروجی Excel قابل بازخوانی نیست: {exc}") from exc
        if not names:
            raise EngineError("خروجی Excel هیچ برگه‌ای ندارد.")

    # ------------------------------------------------------------------ #
    # Styling helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def _title(ws: Worksheet, text: str, last_col: int = 6, row: int = 1) -> None:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(last_col, 2))
        cell = ws.cell(row, 1, text)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name=FONT_NAME, size=14, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 30

    @staticmethod
    def _section(ws: Worksheet, row: int, text: str, last_col: int = 6) -> None:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(last_col, 2))
        cell = ws.cell(row, 1, text)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 22

    @staticmethod
    def _header(ws: Worksheet, row: int, headers: Iterable[str], start_col: int = 1) -> None:
        for offset, header in enumerate(headers):
            cell = ws.cell(row, start_col + offset, header)
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
            cell.font = Font(name=FONT_NAME, size=10, bold=True, color=DARK)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(bottom=THIN, top=THIN, left=THIN, right=THIN)
        ws.row_dimensions[row].height = 26

    @staticmethod
    def _body(ws: Worksheet, first_row: int, last_row: int, first_col: int, last_col: int) -> None:
        for row in ws.iter_rows(min_row=first_row, max_row=last_row, min_col=first_col, max_col=last_col):
            for cell in row:
                cell.font = Font(name=FONT_NAME, size=9, color=DARK)
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
                cell.border = Border(bottom=THIN, left=THIN, right=THIN)

    @staticmethod
    def _fill(cell, color: str) -> None:
        cell.fill = PatternFill("solid", fgColor=color)

    @staticmethod
    def _widths(ws: Worksheet, widths: dict[str, float]) -> None:
        for column, width in widths.items():
            ws.column_dimensions[column].width = width

    @staticmethod
    def _rtl(ws: Worksheet) -> None:
        ws.sheet_view.rightToLeft = True

    @staticmethod
    def _print_setup(ws: Worksheet) -> None:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.horizontalCentered = True

    @staticmethod
    def _value_cell(ws: Worksheet, row: int, col: int, value: Any, *, money: bool = False) -> None:
        if isinstance(value, Decimal):
            cell = ws.cell(row, col, to_excel_number(value))
            cell.number_format = MONEY_FORMAT if money else DECIMAL_FORMAT
        elif isinstance(value, date):
            cell = ws.cell(row, col, format_jalali(value))
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif value is None:
            ws.cell(row, col, "—")
        else:
            ws.cell(row, col, value)

    # ------------------------------------------------------------------ #
    def _dashboard(self, wb: Workbook, context: ReportContext) -> None:
        analysis = context.analysis
        generic = analysis.generic
        integration = getattr(analysis, "integration", None)
        ws = wb.create_sheet("داشبورد")
        self._rtl(ws)
        self._title(ws, "داشبورد مدیریتی تحلیل فایل Excel | Local Office AI", 6)
        ws["A2"] = f"فایل: {Path(analysis.meta.path).name}"
        ws.merge_cells("A2:F2")
        ws["A2"].font = Font(name=FONT_NAME, size=9, color=DARK)
        ws["A2"].alignment = Alignment(horizontal="right")
        intent = getattr(analysis, "request_intent", None)
        if intent is not None and getattr(intent, "text", ""):
            ws["A3"] = f"درخواست کاربر: {intent.text}"
            ws.merge_cells("A3:F3")
            ws["A3"].font = Font(name=FONT_NAME, size=9, color=DARK)
            ws["A3"].alignment = Alignment(horizontal="right", wrap_text=True)
        row = 5 if intent is not None and getattr(intent, "text", "") else 4
        charted = False

        if generic is not None and not analysis.loans:
            quality = generic.quality
            self._section(ws, row, "خلاصه داده (محاسبه‌شده با Python)", 6)
            row += 1
            self._header(ws, row, ["شاخص", "مقدار", "واحد", "توضیح", "", ""])
            header_row = row
            row += 1
            summary = [
                ("تعداد ردیف داده", generic.row_count, "ردیف", "ردیف‌های معتبر جدول اصلی"),
                ("تعداد ستون", generic.column_count, "ستون", "همه ستون‌های جدول"),
                (
                    "ستون اندازه‌گیری عددی",
                    sum(1 for c in generic.columns if c.is_numeric and not c.is_identifier and not c.is_temporal_key),
                    "ستون",
                    "قابل محاسبه آماری؛ بدون شناسه و کلید زمانی",
                ),
                (
                    "ستون تاریخ/کلید زمانی",
                    sum(1 for c in generic.columns if c.is_temporal_key or c.kind.value == "date"),
                    "ستون",
                    "مبنای روند زمانی",
                ),
                (
                    "ستون شناسه",
                    sum(1 for c in generic.columns if c.is_identifier),
                    "ستون",
                    "کنارگذاشته از آماره‌ها",
                ),
                ("ستون متنی", sum(1 for c in generic.columns if c.kind.value == "text"), "ستون", "مبنای گروه‌بندی و فراوانی"),
                (
                    "درصد سلول خالی (ستون‌های اندازه‌گیری‌شده)",
                    f"{(quality.measured_missing_ratio if quality and quality.empty_columns else quality.missing_ratio):.1%}"
                    if quality
                    else "0.0%",
                    "",
                    "ستون‌های بدون هیچ مقدار جدا شمرده شدند",
                ),
                (
                    "ستون بدون داده (دوره اندازه‌گیری‌نشده)",
                    len(quality.empty_columns) if quality else 0,
                    "ستون",
                    "مثل ستون سال آینده که خالی است؛ صفر تلقی نشد",
                ),
                ("ردیف تکراری", quality.duplicate_rows if quality else 0, "ردیف", "تکراری کامل؛ حذف نشده"),
            ]
            if integration is not None and getattr(integration, "is_usable", False):
                headline = integration.headline_metric()
                summary.extend(
                    [
                        ("ماه‌های تحلیل‌شده", len(integration.months), "ماه", "، ".join(integration.months)),
                        ("دوره‌های تحلیل‌شده", len(integration.periods), "دوره", integration.periods_summary()),
                        (
                            "شاخص معیار داشبورد",
                            headline.metric if headline is not None else "—",
                            "شاخص",
                            f"بیشترین حجم در دوره {integration.main_period() or '—'}",
                        ),
                    ]
                )
            elif generic.outliers:
                summary.append(("داده پرت", len(generic.outliers), "مقدار", "علامت‌گذاری‌شده؛ حذف نشده"))
            if analysis.intelligence is not None:
                counts = analysis.intelligence.status_counts
                summary.extend(
                    [
                        (
                            "برگه‌های تحلیل‌شده",
                            counts.get("fully_analyzed", 0),
                            "برگه",
                            "کاملاً تحلیل‌شده",
                        ),
                        (
                            "برگه‌های جزئی",
                            counts.get("partially_analyzed", 0),
                            "برگه",
                            "دوره ناقص یا پوشش محدود",
                        ),
                        (
                            "مغایرت جمع/مقایسه",
                            sum(
                                1
                                for item in analysis.intelligence.all_audits
                                if item.status.value == "mismatch"
                                and item.kind in {"total", "difference"}
                            ),
                            "مورد",
                            "بدون اصلاح خودکار",
                        ),
                        (
                            "سلول ردیف درصد با مقدار ثابت",
                            sum(
                                1
                                for item in analysis.intelligence.all_audits
                                if item.status.value == "mismatch" and item.kind == "percentage"
                            ),
                            "سلول",
                            "ردیف نمایشی؛ مبنای نتیجه‌گیری نیست",
                        ),
                    ]
                )
            for label, value, unit, note in summary:
                self._value_cell(ws, row, 1, label)
                self._value_cell(ws, row, 2, value)
                self._value_cell(ws, row, 3, unit)
                self._value_cell(ws, row, 4, note)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 4)

            grouped = [
                item for item in analysis.calculations
                if item.name.startswith("group_sum::") and isinstance(item.value, Decimal)
            ][:6]
            if grouped:
                chart_start = row + 2
                ws.cell(chart_start, 1, "گروه")
                ws.cell(chart_start, 2, "جمع")
                for index, item in enumerate(grouped, start=chart_start + 1):
                    label = item.label.split("«")[-1].split("»")[0] if "«" in item.label else item.label
                    ws.cell(index, 1, label)
                    ws.cell(index, 2, to_excel_number(item.value)).number_format = MONEY_FORMAT
                chart = BarChart()
                chart.type = "bar"
                chart.style = 10
                chart.title = "جمع مقادیر به تفکیک گروه"
                chart.y_axis.title = "مقدار"
                data = Reference(ws, min_col=2, min_row=chart_start, max_row=chart_start + len(grouped))
                categories = Reference(ws, min_col=1, min_row=chart_start + 1, max_row=chart_start + len(grouped))
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(categories)
                chart.height = 8
                chart.width = 16
                ws.add_chart(chart, "H4")
                charted = True
        else:
            kpi_names = [
                "loan_count",
                "principal_sum",
                "installments_sum",
                "obligations_sum",
                "contract_repayment_sum",
                "obligation_gap",
            ]
            kpi_rows = [analysis.calculation(name) for name in kpi_names]
            kpi_rows = [item for item in kpi_rows if item is not None]
            if kpi_rows:
                self._section(ws, row, "شاخص‌های کلیدی (محاسبه‌شده با Python)", 6)
                row += 1
                self._header(ws, row, ["شاخص", "مقدار", "واحد", "روش محاسبه", "کنترل دوباره", "وضعیت"])
                header_row = row
                row += 1
                for calculation in kpi_rows:
                    self._value_cell(ws, row, 1, calculation.label)
                    self._value_cell(ws, row, 2, calculation.value, money=isinstance(calculation.value, Decimal))
                    self._value_cell(ws, row, 3, calculation.unit)
                    self._value_cell(ws, row, 4, calculation.method)
                    verification = "—"
                    if calculation.verified is True:
                        verification = "تأیید شد"
                    elif calculation.verified is False:
                        verification = "تأیید نشد"
                    self._value_cell(ws, row, 5, verification)
                    status = "معتبر" if calculation.verified in (True, None) else "نیازمند بررسی"
                    self._value_cell(ws, row, 6, status)
                    self._fill(ws.cell(row, 6), GREEN if status == "معتبر" else RED)
                    row += 1
                self._body(ws, header_row + 1, row - 1, 1, 6)

            if analysis.audits:
                row += 1
                self._section(ws, row, "حسابرسی جمع‌های اعلامی فایل", 6)
                row += 1
                self._header(ws, row, ["شاخص", "مقدار اعلامی", "محاسبه Python", "اختلاف", "وضعیت", "منبع مقدار اعلامی"])
                header_row = row
                row += 1
                for audit in analysis.audits:
                    self._value_cell(ws, row, 1, audit.label)
                    self._value_cell(ws, row, 2, audit.declared, money=True)
                    self._value_cell(ws, row, 3, audit.computed, money=True)
                    self._value_cell(ws, row, 4, audit.difference, money=True)
                    self._value_cell(ws, row, 5, AUDIT_STATUS_FA.get(audit.status, audit.status.value))
                    source = audit.declared_source
                    self._value_cell(ws, row, 6, f"{source.sheet}!{source.address}" if source else "—")
                    color = {
                        AuditStatus.VERIFIED: GREEN,
                        AuditStatus.MISMATCH: RED,
                        AuditStatus.NOT_VERIFIABLE: YELLOW,
                        AuditStatus.NO_DECLARED_VALUE: GREY,
                    }.get(audit.status, GREY)
                    self._fill(ws.cell(row, 5), color)
                    row += 1
                self._body(ws, header_row + 1, row - 1, 1, 6)

            if analysis.reconciliation:
                row += 1
                self._section(ws, row, "خلاصه تطبیق تعهدات", 6)
                row += 1
                self._header(ws, row, ["مؤسسه", "اقساط وام", "تعهد ماهانه", "اختلاف", "وضعیت", "شباهت نام"])
                header_row = row
                row += 1
                for item in analysis.reconciliation:
                    self._value_cell(ws, row, 1, item.institution)
                    self._value_cell(ws, row, 2, item.loan_total, money=True)
                    self._value_cell(ws, row, 3, item.obligation_total, money=True)
                    self._value_cell(ws, row, 4, item.difference, money=True)
                    self._value_cell(ws, row, 5, MATCH_STATUS_FA.get(item.status, item.status.value))
                    self._value_cell(ws, row, 6, f"{item.similarity:.0%}")
                    color = {
                        MatchStatus.MATCH: GREEN,
                        MatchStatus.PROBABLE: YELLOW,
                        MatchStatus.OUTSIDE_LOAN_TABLE: YELLOW,
                        MatchStatus.NO_OBLIGATION: GREY,
                        MatchStatus.NO_MATCH: RED,
                    }.get(item.status, GREY)
                    self._fill(ws.cell(row, 5), color)
                    row += 1
                self._body(ws, header_row + 1, row - 1, 1, 6)

                chart_rows = [item for item in analysis.reconciliation if item.obligation_total]
                if chart_rows:
                    chart_start = row + 2
                    ws.cell(chart_start, 1, "مؤسسه")
                    ws.cell(chart_start, 2, "تعهد ماهانه")
                    for index, item in enumerate(chart_rows, start=chart_start + 1):
                        ws.cell(index, 1, item.institution)
                        ws.cell(index, 2, to_excel_number(item.obligation_total)).number_format = MONEY_FORMAT
                    chart = BarChart()
                    chart.type = "bar"
                    chart.style = 10
                    chart.title = "تعهد ماهانه به تفکیک مؤسسه"
                    chart.y_axis.title = "مبلغ"
                    data = Reference(ws, min_col=2, min_row=chart_start, max_row=chart_start + len(chart_rows))
                    categories = Reference(ws, min_col=1, min_row=chart_start + 1, max_row=chart_start + len(chart_rows))
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(categories)
                    chart.height = 8
                    chart.width = 16
                    ws.add_chart(chart, "H4")
                    charted = True

        if not charted and integration is not None and getattr(integration, "is_usable", False):
            charted = self._dashboard_monthly_chart(ws, integration, start_row=row + 2)

        if not charted:
            grouped = [
                item for item in analysis.calculations
                if item.name.startswith("group_sum::") and isinstance(item.value, Decimal)
            ][:6]
            if grouped:
                chart_start = row + 2
                ws.cell(chart_start, 1, "گروه")
                ws.cell(chart_start, 2, "جمع")
                for index, item in enumerate(grouped, start=chart_start + 1):
                    label = item.label.split("«")[-1].split("»")[0] if "«" in item.label else item.label
                    ws.cell(index, 1, label)
                    ws.cell(index, 2, to_excel_number(item.value)).number_format = MONEY_FORMAT
                chart = BarChart()
                chart.type = "bar"
                chart.style = 10
                chart.title = "جمع مقادیر به تفکیک گروه"
                chart.y_axis.title = "مقدار"
                data = Reference(ws, min_col=2, min_row=chart_start, max_row=chart_start + len(grouped))
                categories = Reference(ws, min_col=1, min_row=chart_start + 1, max_row=chart_start + len(grouped))
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(categories)
                chart.height = 8
                chart.width = 16
                ws.add_chart(chart, "H4")

        self._widths(ws, {"A": 42, "B": 20, "C": 20, "D": 24, "E": 26, "F": 30})
        ws.freeze_panes = "A6"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    @staticmethod
    def _dashboard_monthly_chart(ws: Worksheet, integration: Any, *, start_row: int) -> bool:
        """Bar chart of the headline metric per month, straight from the matrix."""

        headline = integration.headline_metric()
        if headline is None:
            return False
        period = integration.main_period()
        months = [month for month in headline.months(period) if headline.cell(period, month)]
        if len(months) < 2:
            return False
        ws.cell(start_row, 1, "ماه")
        ws.cell(start_row, 2, headline.metric)
        for offset, month in enumerate(months, start=1):
            cell = headline.cell(period, month)
            label = month if cell.complete else f"{month} (ناقص)"
            ws.cell(start_row + offset, 1, label)
            value_cell = ws.cell(start_row + offset, 2, to_excel_number(cell.total))
            value_cell.number_format = MONEY_FORMAT
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = f"جمع ماهانه «{headline.metric}» — دوره {period}"
        chart.y_axis.title = "مقدار"
        data = Reference(ws, min_col=2, min_row=start_row, max_row=start_row + len(months))
        categories = Reference(ws, min_col=1, min_row=start_row + 1, max_row=start_row + len(months))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        chart.height = 8
        chart.width = 16
        ws.add_chart(chart, "H4")
        return True

    # ------------------------------------------------------------------ #
    def _executive_summary(self, wb: Workbook, context: ReportContext) -> None:
        analysis = context.analysis
        ai = context.ai
        ws = wb.create_sheet("خلاصه اجرایی")
        self._rtl(ws)
        self._title(ws, "خلاصه اجرایی", 4)
        row = 3
        self._header(ws, row, ["موضوع", "مقدار"])
        row += 1
        meta = analysis.meta
        lenders = {loan.institution_raw for loan in analysis.loans if loan.institution_raw}
        matched = sum(1 for item in analysis.reconciliation if item.status is MatchStatus.MATCH)
        probable = sum(1 for item in analysis.reconciliation if item.status is MatchStatus.PROBABLE)
        outside = sum(1 for item in analysis.reconciliation if item.status is MatchStatus.OUTSIDE_LOAN_TABLE)
        verified = sum(1 for audit in analysis.audits if audit.status is AuditStatus.VERIFIED)
        mismatched = sum(1 for audit in analysis.audits if audit.status is AuditStatus.MISMATCH)
        generic = analysis.generic
        profile = context.profile or self.settings.report_profile or "clean"
        audit = profile == "audit"
        intent = getattr(analysis, "request_intent", None)
        info = [("نام فایل", Path(meta.path).name)]
        if audit:
            info.append(("هش فایل (SHA-256)", meta.sha256))
        info.extend(
            [
                ("تعداد برگه", len(meta.sheets)),
                ("اطمینان تشخیص ساختار", f"{analysis.structure_confidence:.0%}"),
            ]
        )
        if intent is not None and getattr(intent, "text", ""):
            info.append(("درخواست کاربر", intent.text))
            info.append(("نوع تحلیل درخواستی", "، ".join(intent.scope_labels)))
        if analysis.loans:
            info.extend([
                ("تعداد ردیف وام شناسایی‌شده", len(analysis.loans)),
                ("تعداد مؤسسه وام‌دهنده", len(lenders)),
                ("تعداد ردیف تعهد شناسایی‌شده", len(analysis.obligations)),
                ("نتیجه تطبیق", f"{matched} مطابق، {probable} احتمالی، {outside} خارج از جدول وام"),
                ("حسابرسی جمع‌های اعلامی", f"{verified} تأیید، {mismatched} مغایرت"),
            ])
        elif generic is not None:
            quality = generic.quality
            info.extend([
                ("جدول تحلیل‌شده", f"{generic.table_sheet}!{generic.table_address}"),
                ("نوع جدول", TABLE_KIND_FA.get(analysis.primary_table.kind, "—") if analysis.primary_table else "—"),
                ("تعداد ردیف داده", generic.row_count),
                ("تعداد ستون", generic.column_count),
                (
                    "درصد سلول خالی (اندازه‌گیری‌شده)",
                    f"{(quality.measured_missing_ratio if quality.empty_columns else quality.missing_ratio):.1%}"
                    if quality
                    else "0.0%",
                ),
                ("ستون بدون داده", len(quality.empty_columns) if quality else 0),
                ("ردیف تکراری", quality.duplicate_rows if quality else 0),
            ])
            if generic.outliers and audit:
                info.append(("داده پرت علامت‌گذاری‌شده", len(generic.outliers)))
            integration = getattr(analysis, "integration", None)
            if integration is not None and getattr(integration, "is_usable", False):
                info.extend(
                    [
                        ("ماه‌های تحلیل‌شده", "، ".join(integration.months)),
                        ("دوره‌های تحلیل‌شده", integration.periods_summary()),
                    ]
                )
                headline = integration.headline_metric()
                if headline is not None:
                    requested = [
                        year
                        for year in (getattr(intent, "years", None) or [])
                        if year in integration.periods
                    ]
                    for period in requested or [integration.main_period()]:
                        total = headline.total(period)
                        if total is None:
                            continue
                        months = len(integration.period_months(period))
                        info.append(
                            (
                                f"جمع «{headline.metric}» در {period} ({months} ماه کامل)",
                                _decimal_text(total),
                            )
                        )
            if analysis.intelligence is not None:
                counts = analysis.intelligence.status_counts
                info.extend([
                    (
                        "پوشش تحلیل برگه‌ها",
                        f"{counts.get('fully_analyzed', 0)} کامل، "
                        f"{counts.get('partially_analyzed', 0)} جزئی، "
                        f"{counts.get('not_analyzed', 0)} بدون تحلیل",
                    ),
                    (
                        "مغایرت جمع/مقایسه",
                        sum(
                            1
                            for item in analysis.intelligence.all_audits
                            if item.status.value == "mismatch" and item.kind in {"total", "difference"}
                        ),
                    ),
                    ("تعارض متادیتا", len(analysis.intelligence.metadata_conflicts)),
                ])
        info.append(
            (
                "تعداد یافته‌های نیازمند بررسی",
                sum(1 for a in analysis.anomalies if a.severity is not Severity.INFO),
            )
        )
        if ai.mode == "model" or audit:
            info.extend(
                [
                    ("حالت تحلیل هوش مصنوعی", "مدل محلی" if ai.mode == "model" else "قواعد قطعی (بدون مدل)"),
                    ("مدل", ai.model or "—"),
                    ("سطح ریسک اعلام‌شده", ai.response.risk_level),
                ]
            )
        for label, value in info:
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, 4, row - 1, 1, 2)

        answers = getattr(analysis, "request_answers", None) or []
        if answers:
            row += 1
            self._section(ws, row, "پاسخ به درخواست (محاسبه‌شده با Python)", 4)
            for answer in answers[:14]:
                row += 1
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
                cell = ws.cell(row, 1, f"• {answer.title}: {answer.text}")
                cell.font = Font(name=FONT_NAME, size=9, color=DARK)
                cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
                ws.row_dimensions[row].height = max(20, min(120, 16 * (len(answer.text) // 95 + 1)))

        row += 1
        self._section(ws, row, "خلاصه مدیریتی", 4)
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        cell = ws.cell(row, 1, ai.response.summary or "خلاصه‌ای تولید نشد.")
        cell.font = Font(name=FONT_NAME, size=10, color=DARK)
        cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
        ws.row_dimensions[row].height = max(45, min(160, 18 * (len(ai.response.summary) // 90 + 2)))

        for title, items in (
            ("یافته‌های کلیدی", ai.response.findings),
            ("اطلاعات ناقص", ai.response.missing_data),
            ("پیشنهادهای عملی", ai.response.recommendations),
        ):
            row += 2
            self._section(ws, row, title, 4)
            if not items:
                row += 1
                ws.cell(row, 1, "موردی ثبت نشد.")
            for item in items:
                row += 1
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
                cell = ws.cell(row, 1, "• " + str(item))
                cell.font = Font(name=FONT_NAME, size=9, color=DARK)
                cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
                ws.row_dimensions[row].height = max(18, min(90, 16 * (len(str(item)) // 90 + 1)))

        self._widths(ws, {"A": 38, "B": 60, "C": 30, "D": 30})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _normalized_data(self, wb: Workbook, analysis: AnalysisResult) -> None:
        ws = wb.create_sheet("داده استانداردشده")
        self._rtl(ws)
        self._title(ws, "داده استانداردشده جدول اصلی (مقدار اصلی در برگه ردیابی منبع)", 8)
        primary = analysis.primary_table
        if primary is None:
            ws.cell(3, 1, "جدول اصلی شناسایی نشد.")
            return
        headers = ["شماره ردیف"] + [column.header or column.letter for column in primary.columns]
        headers += ["بازپرداخت محاسبه‌شده با Python", "آدرس منبع"]
        self._header(ws, 3, headers)
        row = 4
        for record in analysis.normalized_rows:
            ws.cell(row, 1, record["row"])
            for offset, cell in enumerate(record["cells"], start=2):
                self._value_cell(ws, row, offset, cell["value"], money=cell["kind"] == "number")
            loan = next((item for item in analysis.loans if item.provenance.row == record["row"]), None)
            if loan is not None and loan.contract_repayment is not None:
                self._value_cell(ws, row, len(primary.columns) + 2, loan.contract_repayment, money=True)
            source_address = (
                f"{primary.sheet}!{loan.provenance.address}"
                if loan is not None
                else record.get("source_row", "")
            )
            self._value_cell(ws, row, len(primary.columns) + 3, source_address or "—")
            row += 1
        self._body(ws, 4, row - 1, 1, len(headers))
        for index in range(2, len(headers) + 1):
            ws.cell(3, index).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{max(row - 1, 4)}"
        ws.freeze_panes = "A4"
        widths = {"A": 8}
        for index in range(2, len(headers) + 1):
            widths[get_column_letter(index)] = 20
        self._widths(ws, widths)
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _raw_data(self, wb: Workbook, analysis: AnalysisResult) -> None:
        """Every cell of the primary table: raw value, standardized value, source."""
        ws = wb.create_sheet("داده خام")
        self._rtl(ws)
        self._title(ws, "مقدار خام و استانداردشده هر سلول (قابل ردیابی)", 8)
        ws["A2"] = "مقدار اصلی هرگز حذف نمی‌شود؛ ستون «مقدار استانداردشده» فقط نتیجه تبدیل است."
        ws.merge_cells("A2:H2")
        ws["A2"].font = Font(name=FONT_NAME, size=9, color=DARK)
        ws["A2"].alignment = Alignment(horizontal="right")

        headers = ["برگه", "سلول", "ستون", "عنوان", "مقدار خام", "مقدار استانداردشده", "نوع", "هشدارها"]
        self._header(ws, 3, headers)
        row = 4
        for source_row in analysis.normalized_rows:
            for cell in source_row["cells"]:
                self._value_cell(ws, row, 1, source_row["sheet"])
                self._value_cell(ws, row, 2, f"{cell['column']}{source_row['row']}")
                self._value_cell(ws, row, 3, cell["column"])
                self._value_cell(ws, row, 4, cell.get("header") or "—")
                self._value_cell(ws, row, 5, cell.get("raw"))
                value = cell.get("value")
                self._value_cell(ws, row, 6, value, money=isinstance(value, Decimal))
                self._value_cell(ws, row, 7, cell.get("kind"))
                warnings = cell.get("warnings") or []
                self._value_cell(ws, row, 8, "؛ ".join(str(item) for item in warnings) or "—")
                if warnings:
                    self._fill(ws.cell(row, 8), YELLOW)
                row += 1
        if row == 4:
            ws.cell(4, 1, "داده‌ای برای نمایش وجود ندارد.")
        else:
            self._body(ws, 4, row - 1, 1, 8)
        ws.auto_filter.ref = f"A3:H{max(row - 1, 4)}"
        ws.freeze_panes = "A4"
        self._widths(ws, {"A": 18, "B": 10, "C": 8, "D": 24, "E": 28, "F": 28, "G": 12, "H": 34})
        self._print_setup(ws)

    def _validation(self, wb: Workbook, analysis: AnalysisResult) -> None:
        """Pass/fail checklist of every quality rule applied to the data."""
        from local_office_ai.excel.validation import STATUS_LABELS, overall_status

        ws = wb.create_sheet("اعتبارسنجی داده")
        self._rtl(ws)
        self._title(ws, "اعتبارسنجی داده — فهرست بررسی‌های انجام‌شده", 5)
        checks = list(analysis.validation)
        status = overall_status(checks)
        ws["A2"] = (
            f"وضعیت کلی: {STATUS_LABELS.get(status, status)} — "
            f"{sum(1 for c in checks if c.status == 'passed')} قبول، "
            f"{sum(1 for c in checks if c.status == 'warning')} هشدار، "
            f"{sum(1 for c in checks if c.status == 'failed')} رد، "
            f"{sum(1 for c in checks if c.status == 'skipped')} بررسی‌نشده"
        )
        ws.merge_cells("A2:E2")
        ws["A2"].font = Font(name=FONT_NAME, size=9, color=DARK)
        ws["A2"].alignment = Alignment(horizontal="right")
        self._fill(ws["A2"], {"passed": GREEN, "warning": YELLOW, "failed": RED}.get(status, GREY))

        self._header(ws, 4, ["بررسی", "دامنه", "وضعیت", "تعداد اثر", "توضیح"])
        row = 5
        colors = {"passed": GREEN, "warning": YELLOW, "failed": RED, "skipped": GREY}
        for check in checks:
            self._value_cell(ws, row, 1, check.label)
            self._value_cell(ws, row, 2, check.scope)
            self._value_cell(ws, row, 3, STATUS_LABELS.get(check.status, check.status))
            self._value_cell(ws, row, 4, check.affected)
            self._value_cell(ws, row, 5, check.detail)
            self._fill(ws.cell(row, 3), colors.get(check.status, GREY))
            row += 1
        if row == 5:
            ws.cell(5, 1, "اعتبارسنجی برای این تحلیل اجرا نشد.")
        else:
            self._body(ws, 5, row - 1, 1, 5)

        gate = getattr(analysis, "quality_gate", None)
        if gate is not None:
            from local_office_ai.quality import GATE_STATUS_FA, GateStatus

            row += 1
            self._section(ws, row, "دروازه کیفیت یکپارچه (ارزیابی پیش از تولید سند نهایی)", 5)
            row += 1
            self._value_cell(ws, row, 1, "وضعیت کلی دروازه کیفیت")
            self._value_cell(ws, row, 2, GATE_STATUS_FA.get(gate.status, gate.status.value))
            self._fill(
                ws.cell(row, 2),
                {
                    GateStatus.PASS: GREEN,
                    GateStatus.WARNING: YELLOW,
                    GateStatus.FAIL: RED,
                }.get(gate.status, GREY),
            )
            counts = gate.counts()
            self._value_cell(
                ws,
                row,
                3,
                (
                    f"PASS {counts.get('PASS', 0)} / WARNING {counts.get('WARNING', 0)} / "
                    f"FAIL {counts.get('FAIL', 0)} / مسدودکننده {counts.get('blocking', 0)}"
                ),
            )
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
            row += 1
            wrote_gate = False
            for stage in gate.stages:
                if stage.status is GateStatus.PASS:
                    continue
                wrote_gate = True
                self._value_cell(ws, row, 1, stage.label_fa)
                self._value_cell(ws, row, 2, stage.summary)
                self._value_cell(ws, row, 3, GATE_STATUS_FA.get(stage.status, stage.status.value))
                self._value_cell(ws, row, 4, "بله" if stage.blocking else "خیر")
                issues = [check.message for check in stage.checks if check.status is not GateStatus.PASS]
                self._value_cell(ws, row, 5, "؛ ".join(issues[:4]) or "—")
                self._fill(
                    ws.cell(row, 3),
                    {
                        GateStatus.PASS: GREEN,
                        GateStatus.WARNING: YELLOW,
                        GateStatus.FAIL: RED,
                    }.get(stage.status, GREY),
                )
                row += 1
            if not wrote_gate:
                self._value_cell(ws, row, 1, "همه مراحل ارزیابی‌شده PASS هستند.")
                row += 1

        quality = getattr(analysis, "quality", None)
        if quality is not None:
            from local_office_ai.excel.data_quality import QUALITY_STATUS_FA

            row += 1
            self._section(ws, row, "طبقه‌بندی معنایی و کیفیت داده (PASS/WARNING/FAIL)", 5)
            row += 1
            counts = quality.counts()
            self._value_cell(ws, row, 1, "وضعیت کلی کیفیت داده")
            self._value_cell(ws, row, 2, QUALITY_STATUS_FA.get(quality.status, quality.status.value))
            self._fill(
                ws.cell(row, 2),
                {"PASS": GREEN, "WARNING": YELLOW, "FAIL": RED}.get(quality.status.value, GREY),
            )
            self._value_cell(
                ws,
                row,
                3,
                f"PASS {counts.get('PASS', 0)} / WARNING {counts.get('WARNING', 0)} / FAIL {counts.get('FAIL', 0)}",
            )
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
            row += 1
            wrote_finding = False
            for finding in quality.findings:
                if finding.status.value == "PASS":
                    continue
                wrote_finding = True
                self._value_cell(ws, row, 1, finding.title)
                self._value_cell(ws, row, 2, finding.scope)
                self._value_cell(
                    ws, row, 3, QUALITY_STATUS_FA.get(finding.status, finding.status.value)
                )
                self._value_cell(ws, row, 4, finding.count)
                self._value_cell(ws, row, 5, finding.message)
                self._fill(
                    ws.cell(row, 3),
                    {"PASS": GREEN, "WARNING": YELLOW, "FAIL": RED}.get(finding.status.value, GREY),
                )
                row += 1
            if not wrote_finding:
                self._value_cell(ws, row, 1, "یافته‌ای خارج از PASS ثبت نشد.")
                row += 1

        reconciliation = getattr(analysis, "statistical_reconciliation", None)
        if reconciliation is not None:
            from local_office_ai.excel.statistical_reconciliation import STATUS_FA

            row += 1
            self._section(ws, row, "تطبیق آماری و صحت محاسبات (PASS/WARNING/FAIL)", 5)
            row += 1
            counts = reconciliation.counts()
            self._value_cell(ws, row, 1, "وضعیت کلی تطبیق")
            self._value_cell(
                ws, row, 2, STATUS_FA.get(reconciliation.status, reconciliation.status.value)
            )
            self._fill(
                ws.cell(row, 2),
                {"PASS": GREEN, "WARNING": YELLOW, "FAIL": RED}.get(
                    reconciliation.status.value, GREY
                ),
            )
            self._value_cell(
                ws,
                row,
                3,
                f"PASS {counts.get('PASS', 0)} / WARNING {counts.get('WARNING', 0)} / "
                f"FAIL {counts.get('FAIL', 0)} / هشدارهای موتور "
                f"{len(reconciliation.warnings)}",
            )
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
            row += 1
            wrote = False
            for warning in reconciliation.warnings:
                wrote = True
                self._value_cell(ws, row, 1, "هشدار موتور تطبیق")
                self._value_cell(ws, row, 2, "—")
                self._value_cell(ws, row, 3, "هشدار")
                self._value_cell(ws, row, 4, "")
                self._value_cell(ws, row, 5, warning)
                self._fill(ws.cell(row, 3), YELLOW)
                row += 1
            for item in [
                *reconciliation.reconciliations,
                *reconciliation.comparisons,
                *reconciliation.averages,
            ]:
                if item.status.value == "PASS":
                    continue
                wrote = True
                self._value_cell(
                    ws,
                    row,
                    1,
                    getattr(item, "kind_label", "میانگین وزنی"),
                )
                self._value_cell(
                    ws,
                    row,
                    2,
                    getattr(item, "label", "") or getattr(item, "metric", "") or "—",
                )
                self._value_cell(ws, row, 3, STATUS_FA.get(item.status, item.status.value))
                expected = getattr(item, "expected", None)
                if expected is None:
                    expected = getattr(item, "value", None)
                if expected is None:
                    change = getattr(item, "change", None)
                    expected = getattr(change, "current", "") if change is not None else ""
                self._value_cell(ws, row, 4, expected)
                self._value_cell(ws, row, 5, item.explanation)
                self._fill(
                    ws.cell(row, 3),
                    {"PASS": GREEN, "WARNING": YELLOW, "FAIL": RED}.get(item.status.value, GREY),
                )
                row += 1
            if not wrote:
                self._value_cell(ws, row, 1, "تطبیق آماری خارج از PASS ثبت نشد.")
                row += 1

        duplicate_control = getattr(analysis, "duplicate_control", None)
        if duplicate_control is not None:
            from local_office_ai.excel.duplicate_control import AGGREGATION_STATUS_FA

            row += 1
            self._section(ws, row, "کنترل تکرار و تجمیع ایمن (شاهد و تصمیم)", 5)
            row += 1
            counts = duplicate_control.counts()
            roles = duplicate_control.role_counts()
            self._value_cell(ws, row, 1, "خلاصه کنترل تکرار")
            self._value_cell(
                ws,
                row,
                2,
                f"داده اولیه {roles.get('primary', 0)} / خلاصه {roles.get('summary', 0)} / "
                f"مشتق {roles.get('derived', 0)} / تکراری {roles.get('duplicate', 0)}",
            )
            self._value_cell(
                ws,
                row,
                3,
                f"گروه تکراری {counts['duplicate_groups']} / رکورد تکراری "
                f"{counts['record_duplicates']} / مجاز {counts['allowed']} / "
                f"نیازمند بررسی {counts['review']} / مسدود {counts['blocked']}",
            )
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
            row += 1
            for warning in duplicate_control.warnings[:6]:
                self._value_cell(ws, row, 1, "هشدار کنترل تکرار")
                self._value_cell(ws, row, 2, warning)
                ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
                row += 1
            wrote = False
            pending = [
                *duplicate_control.blocked_decisions,
                *duplicate_control.review_decisions,
            ]
            for decision in pending[:40]:
                wrote = True
                self._value_cell(ws, row, 1, decision.metric)
                self._value_cell(ws, row, 2, "؛ ".join(decision.members))
                self._value_cell(
                    ws,
                    row,
                    3,
                    AGGREGATION_STATUS_FA.get(decision.status, decision.status.value),
                )
                self._value_cell(ws, row, 4, len(decision.checks))
                self._value_cell(ws, row, 5, "؛ ".join(decision.reasons) or "—")
                self._fill(
                    ws.cell(row, 3),
                    RED if decision.status.value == "blocked" else YELLOW,
                )
                row += 1
            if len(pending) > 40:
                self._value_cell(
                    ws,
                    row,
                    1,
                    f"… و {len(pending) - 40} تصمیم دیگر در برگه «کنترل تکرار و تجمیع» (نمای audit).",
                )
                row += 1
            if not wrote:
                self._value_cell(ws, row, 1, "تصمیم مسدود یا نیازمند بررسی ثبت نشد.")
                row += 1

        ws.freeze_panes = "A5"
        self._widths(ws, {"A": 34, "B": 26, "C": 12, "D": 12, "E": 60})
        self._print_setup(ws)

    def _calculations(self, wb: Workbook, analysis: AnalysisResult) -> None:
        ws = wb.create_sheet("نتایج محاسبات")
        self._rtl(ws)
        self._title(ws, "نتایج محاسبات قطعی Python با کنترل دوباره", 8)
        self._header(
            ws,
            3,
            ["شاخص", "مقدار", "واحد", "روش محاسبه", "کنترل دوباره", "مقدار کنترل", "وضعیت کنترل", "منابع (نمونه)"],
        )
        row = 4
        for calculation in analysis.calculations:
            self._value_cell(ws, row, 1, calculation.label)
            self._value_cell(ws, row, 2, calculation.value, money=isinstance(calculation.value, Decimal))
            self._value_cell(ws, row, 3, calculation.unit)
            self._value_cell(ws, row, 4, calculation.method)
            self._value_cell(ws, row, 5, calculation.verification_method or "—")
            self._value_cell(ws, row, 6, calculation.verification_value, money=True)
            status = "—" if calculation.verified is None else ("تأیید شد" if calculation.verified else "مغایرت")
            self._value_cell(ws, row, 7, status)
            if calculation.verified is True:
                self._fill(ws.cell(row, 7), GREEN)
            elif calculation.verified is False:
                self._fill(ws.cell(row, 7), RED)
            sources = "، ".join(f"{item.sheet}!{item.address}" for item in calculation.sources[:3])
            self._value_cell(ws, row, 8, sources or "—")
            row += 1
        self._body(ws, 4, row - 1, 1, 8)
        ws.auto_filter.ref = f"A3:H{max(row - 1, 4)}"
        ws.freeze_panes = "A4"
        self._widths(ws, {"A": 40, "B": 20, "C": 14, "D": 52, "E": 40, "F": 20, "G": 18, "H": 44})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _audits(self, wb: Workbook, analysis: AnalysisResult) -> None:
        ws = wb.create_sheet("حسابرسی اعداد")
        self._rtl(ws)
        self._title(ws, "حسابرسی مقدارهای اعلامی فایل در برابر محاسبه مستقل Python", 7)
        self._header(
            ws,
            3,
            ["شاخص", "مقدار اعلامی", "محاسبه Python", "اختلاف", "وضعیت", "منبع مقدار اعلامی", "روش"],
        )
        row = 4
        for audit in analysis.audits:
            self._value_cell(ws, row, 1, audit.label)
            self._value_cell(ws, row, 2, audit.declared, money=True)
            self._value_cell(ws, row, 3, audit.computed, money=True)
            self._value_cell(ws, row, 4, audit.difference, money=True)
            self._value_cell(ws, row, 5, AUDIT_STATUS_FA.get(audit.status, audit.status.value))
            source = audit.declared_source
            self._value_cell(ws, row, 6, f"{source.sheet}!{source.address}" if source else "—")
            self._value_cell(ws, row, 7, audit.method)
            self._fill(
                ws.cell(row, 5),
                {
                    AuditStatus.VERIFIED: GREEN,
                    AuditStatus.MISMATCH: RED,
                    AuditStatus.NOT_VERIFIABLE: YELLOW,
                    AuditStatus.NO_DECLARED_VALUE: GREY,
                }.get(audit.status, GREY),
            )
            row += 1
        if not analysis.audits:
            ws.cell(4, 1, "مقدار اعلامی‌ای در فایل پیدا نشد.")
        self._body(ws, 4, max(row - 1, 4), 1, 7)
        ws.freeze_panes = "A4"
        self._widths(ws, {"A": 40, "B": 20, "C": 20, "D": 18, "E": 22, "F": 26, "G": 60})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _reconciliation(self, wb: Workbook, analysis: AnalysisResult) -> None:
        ws = wb.create_sheet("تطبیق تعهدات")
        self._rtl(ws)
        self._title(ws, "تطبیق اقساط وام با تعهدات ماهانه ثبت‌شده", 10)
        self._header(
            ws,
            3,
            ["مؤسسه", "اقساط وام", "تعهد ماهانه", "اختلاف", "وضعیت", "شباهت نام", "نام در جدول وام", "نام در تعهدات", "توضیح", "منابع"],
        )
        row = 4
        for item in analysis.reconciliation:
            self._value_cell(ws, row, 1, item.institution)
            self._value_cell(ws, row, 2, item.loan_total, money=True)
            self._value_cell(ws, row, 3, item.obligation_total, money=True)
            self._value_cell(ws, row, 4, item.difference, money=True)
            self._value_cell(ws, row, 5, MATCH_STATUS_FA.get(item.status, item.status.value))
            self._value_cell(ws, row, 6, f"{item.similarity:.0%}")
            self._value_cell(ws, row, 7, item.loan_name or "—")
            self._value_cell(ws, row, 8, item.obligation_name or "—")
            self._value_cell(ws, row, 9, item.note)
            sources = "، ".join(
                f"{source.sheet}!{source.address}"
                for source in (item.loan_sources + item.obligation_sources)[:4]
            )
            self._value_cell(ws, row, 10, sources or "—")
            self._fill(
                ws.cell(row, 5),
                {
                    MatchStatus.MATCH: GREEN,
                    MatchStatus.PROBABLE: YELLOW,
                    MatchStatus.OUTSIDE_LOAN_TABLE: YELLOW,
                    MatchStatus.NO_OBLIGATION: GREY,
                    MatchStatus.NO_MATCH: RED,
                }.get(item.status, GREY),
            )
            row += 1
        if not analysis.reconciliation:
            ws.cell(4, 1, "ردیف تطبیقی برای نمایش وجود ندارد.")
        self._body(ws, 4, max(row - 1, 4), 1, 10)
        ws.auto_filter.ref = f"A3:J{max(row - 1, 4)}"
        ws.freeze_panes = "A4"
        self._widths(ws, {"A": 34, "B": 18, "C": 18, "D": 16, "E": 30, "F": 12, "G": 26, "H": 26, "I": 52, "J": 40})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _anomalies(self, wb: Workbook, analysis: AnalysisResult) -> None:
        ws = wb.create_sheet("موارد نیازمند بررسی")
        self._rtl(ws)
        self._title(ws, "موارد نیازمند بررسی (مغایرت‌ها، ناسازگاری‌ها و ریسک‌ها)", 6)
        self._header(ws, 3, ["شدت", "دسته", "پیام", "برگه", "سلول", "جزئیات"])
        row = 4
        for anomaly in analysis.anomalies:
            self._value_cell(ws, row, 1, SEVERITY_FA.get(anomaly.severity, anomaly.severity.value))
            self._value_cell(ws, row, 2, anomaly.category)
            self._value_cell(ws, row, 3, anomaly.message)
            self._value_cell(ws, row, 4, anomaly.sheet or "—")
            self._value_cell(ws, row, 5, anomaly.address or "—")
            self._value_cell(ws, row, 6, anomaly.detail or "—")
            self._fill(
                ws.cell(row, 1),
                {Severity.CRITICAL: RED, Severity.WARNING: YELLOW, Severity.INFO: GREY}.get(anomaly.severity, GREY),
            )
            row += 1
        if not analysis.anomalies:
            ws.cell(4, 1, "موردی برای بررسی پیدا نشد.")
        self._body(ws, 4, max(row - 1, 4), 1, 6)
        ws.auto_filter.ref = f"A3:F{max(row - 1, 4)}"
        ws.freeze_panes = "A4"
        self._widths(ws, {"A": 12, "B": 26, "C": 60, "D": 22, "E": 12, "F": 60})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _ai_sheet(self, wb: Workbook, context: ReportContext) -> None:
        ai = context.ai
        ws = wb.create_sheet("تحلیل هوش مصنوعی")
        self._rtl(ws)
        self._title(ws, "تحلیل هوش مصنوعی و کنترل اعتبار پاسخ", 4)
        row = 3
        self._header(ws, row, ["موضوع", "مقدار"])
        row += 1
        provenance = ai.metadata.get("provenance", {})
        model_ran = ai.mode == "model"
        chunks_known = model_ran or ai.metadata.get("chunk_count") is not None
        rows = [
            ("تحلیل Python (ساختار، آماره‌ها، تطبیق)", "اجرا شد — مستقل از مدل"),
            (
                "تحلیل واقعی Qwen",
                (
                    f"اجرا شد ({ai.metadata.get('successful_chunks', 0)} از "
                    f"{ai.metadata.get('chunk_count', 0)} بخش)"
                    if ai.mode == "model"
                    else "اجرا نشد"
                ),
            ),
            (
                "تحلیل Rule-Based جایگزین",
                "اجرا شد" if provenance.get("rule_based_fallback") == "executed" else "استفاده نشد",
            ),
            (
                "بخش‌های ناموفق مدل",
                ai.metadata.get("failed_chunks", 0) if chunks_known else "—",
            ),
            ("بخش‌های Context", ai.metadata.get("chunk_count", 0) if chunks_known else "—"),
            ("بزرگ‌ترین Context (کاراکتر)", ai.metadata.get("max_context_chars", ai.metadata.get("context_chars", "—"))),
            ("بزرگ‌ترین Prompt (کاراکتر)", ai.metadata.get("max_prompt_chars", ai.metadata.get("prompt_chars", "—"))),
            ("سقف Context مدل (توکن)", self.settings.ollama_num_ctx),
            ("برآورد بزرگ‌ترین Prompt (توکن)", ai.metadata.get("max_estimated_prompt_tokens", "—")),
            ("حالت تحلیل", "مدل محلی" if model_ran else "قواعد قطعی Python (بدون مدل)"),
            ("مدل", ai.model or "—"),
            ("سطح ریسک اعلام‌شده", ai.response.risk_level),
            (
                "پاسخ به داده‌های واقعی متصل است",
                ("بله" if ai.response.is_grounded else "خیر") if model_ran else "— (مدل اجرا نشد)",
            ),
            (
                "پاسخ JSON معتبر بود",
                ("بله" if ai.response.parsed_as_json else "خیر") if model_ran else "—",
            ),
            ("سیاست مقایسه اعداد", ai.response.policy.describe_fa()),
            ("اعداد بدون منبع", ("، ".join(ai.response.unsupported_numbers) or "ندارد") if model_ran else "—"),
            ("تناقض با حسابرسی", ("؛ ".join(ai.response.contradictions) or "ندارد") if model_ran else "—"),
            ("هشدارهای اعتبارسنجی", "؛ ".join(ai.response.warnings) or "ندارد"),
            ("زمان درخواست مدل (ثانیه)", ai.metadata.get("client_elapsed_seconds", "—")),
            ("توکن ورودی", ai.metadata.get("prompt_tokens", "—")),
            ("توکن خروجی", ai.metadata.get("output_tokens", "—")),
            ("سرعت تولید (توکن/ثانیه)", ai.metadata.get("tokens_per_second", "—")),
            ("خطای مدل", ai.error or "ندارد"),
        ]
        for label, value in rows:
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, 4, row - 1, 1, 2)

        chunks = ai.metadata.get("chunks") or []
        if chunks:
            row += 2
            self._section(ws, row, "بخش‌بندی Context و نتیجه هر درخواست", 5)
            row += 1
            self._header(ws, row, ["بخش", "وضعیت", "Context (کاراکتر)", "Prompt (کاراکتر)", "برآورد توکن Prompt"])
            row += 1
            for chunk in chunks:
                values = [
                    chunk.get("label", "—"),
                    "موفق" if chunk.get("status") == "ok" else "ناموفق",
                    chunk.get("context_chars", "—"),
                    chunk.get("prompt_chars", "—"),
                    chunk.get("estimated_prompt_tokens", "—"),
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                if chunk.get("error"):
                    ws.cell(row, 6, chunk["error"])
                row += 1

        for title, items in (
            ("خلاصه مدیریتی", [ai.response.summary] if ai.response.summary else []),
            ("یافته‌ها", ai.response.findings),
            ("اطلاعات ناقص", ai.response.missing_data),
            ("پیشنهادها", ai.response.recommendations),
        ):
            row += 2
            self._section(ws, row, title, 4)
            if not items:
                row += 1
                ws.cell(row, 1, "موردی ثبت نشد.")
            for item in items:
                row += 1
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
                cell = ws.cell(row, 1, "• " + str(item))
                cell.font = Font(name=FONT_NAME, size=9, color=DARK)
                cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
                ws.row_dimensions[row].height = max(18, min(120, 16 * (len(str(item)) // 100 + 1)))
        self._widths(ws, {"A": 42, "B": 70, "C": 30, "D": 30, "E": 28})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _traceability(self, wb: Workbook, context: ReportContext) -> None:
        analysis = context.analysis
        ws = wb.create_sheet("ردیابی منبع")
        self._rtl(ws)
        self._title(ws, "ردیابی منبع داده و مسیر محاسبه", 6)
        row = 3
        self._header(ws, row, ["موضوع", "مقدار"])
        row += 1
        for label, value in (
            ("مسیر فایل ورودی", analysis.meta.path),
            ("هش SHA-256 فایل ورودی", analysis.meta.sha256),
            ("حجم فایل (بایت)", analysis.meta.size_bytes),
            ("زمان خواندن فایل (ثانیه)", round(analysis.meta.load_seconds, 4)),
            ("فایل ماکرو دارد", "بله" if analysis.meta.has_macros else "خیر"),
            ("تعداد فرمول", analysis.meta.total_formulas),
            ("تعداد سلول خطا", analysis.meta.total_error_cells),
            ("پیوند خارجی", analysis.meta.external_links),
            ("نام‌های تعریف‌شده", "، ".join(analysis.meta.defined_names) or "ندارد"),
        ):
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, 4, row - 1, 1, 2)

        row += 1
        self._section(ws, row, "ردیابی شاخص‌های کلیدی تا سلول منبع", 6)
        row += 1
        self._header(ws, row, ["شاخص", "مقدار", "روش", "منبع ۱", "منبع ۲", "منبع ۳"])
        row += 1
        for calculation in analysis.calculations:
            if not calculation.sources:
                continue
            self._value_cell(ws, row, 1, calculation.label)
            self._value_cell(ws, row, 2, calculation.value, money=isinstance(calculation.value, Decimal))
            self._value_cell(ws, row, 3, calculation.method)
            for index in range(3):
                source = calculation.sources[index] if index < len(calculation.sources) else None
                self._value_cell(
                    ws, row, 4 + index,
                    f"{source.sheet}!{source.address}" if source else "—",
                )
            row += 1
        self._body(ws, row - len([c for c in analysis.calculations if c.sources]), row - 1, 1, 6)

        generic = analysis.generic
        if generic is not None and generic.columns:
            row += 1
            self._section(ws, row, "ردیابی آماره‌های عمومی تا ستون و سلول منبع", 6)
            row += 1
            self._header(ws, row, ["شاخص", "مقدار", "روش", "محدوده/سلول منبع", "توضیح", "—"])
            row += 1
            start_row = row
            for column in generic.columns:
                self._value_cell(ws, row, 1, f"پروفایل ستون «{column.header}»")
                self._value_cell(ws, row, 2, f"n={column.count}")
                self._value_cell(ws, row, 3, "پروفایل نوع/نقش/کیفیت")
                self._value_cell(
                    ws, row, 4,
                    f"{generic.table_sheet}!{column.letter}:{column.letter} (جدول {generic.table_address})",
                )
                self._value_cell(ws, row, 5, f"نوع={column.kind.value}؛ نقش={column.role.value}")
                row += 1
            for outlier in generic.outliers:
                self._value_cell(ws, row, 1, f"پرت ستون «{outlier.header}»")
                self._value_cell(ws, row, 2, outlier.value, money=True)
                self._value_cell(ws, row, 3, outlier.method)
                self._value_cell(ws, row, 4, f"{generic.table_sheet}!{outlier.address}")
                self._value_cell(ws, row, 5, f"فنس {outlier.lower_fence} تا {outlier.upper_fence}")
                row += 1
            for trend in generic.trends:
                self._value_cell(ws, row, 1, f"روند «{trend.measure_header}»")
                self._value_cell(ws, row, 2, trend.direction)
                self._value_cell(ws, row, 3, f"{trend.granularity} / {trend.calendar}")
                date_letter = trend.date_key.split(":", 1)[0]
                self._value_cell(ws, row, 4, f"{generic.table_sheet}!{date_letter}:{date_letter}")
                self._value_cell(ws, row, 5, trend.note)
                row += 1
            self._body(ws, start_row, row - 1, 1, 6)

        row += 1
        self._section(ws, row, "ردیابی ردیف‌های وام (مقدار خام و استانداردشده)", 6)
        row += 1
        self._header(ws, row, ["مؤسسه (خام)", "کلید تطبیق", "اصل وام", "قسط", "تعداد", "سلول منبع"])
        row += 1
        for loan in analysis.loans:
            self._value_cell(ws, row, 1, loan.institution_raw)
            self._value_cell(ws, row, 2, loan.institution_key)
            self._value_cell(ws, row, 3, loan.principal, money=True)
            self._value_cell(ws, row, 4, loan.installment, money=True)
            self._value_cell(ws, row, 5, loan.installment_count)
            self._value_cell(ws, row, 6, f"{loan.provenance.sheet}!{loan.provenance.address}")
            row += 1
        self._body(ws, row - len(analysis.loans), row - 1, 1, 6)
        self._widths(ws, {"A": 44, "B": 40, "C": 30, "D": 24, "E": 16, "F": 24})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _performance(self, wb: Workbook, context: ReportContext) -> None:
        report = context.performance
        if report is None:
            return
        ws = wb.create_sheet("گزارش عملکرد")
        self._rtl(ws)
        self._title(ws, "گزارش عملکرد، منابع سیستم و مدل", 5)
        self._header(ws, 3, ["شاخص", "مقدار", "واحد", "دامنه", "توضیح اندازه‌گیری"])
        row = 4
        for label, value, unit, scope, note in performance_rows(report):
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            self._value_cell(ws, row, 3, unit)
            self._value_cell(ws, row, 4, scope)
            self._value_cell(ws, row, 5, note)
            row += 1
        self._body(ws, 4, row - 1, 1, 5)

        row += 1
        self._section(ws, row, "مشخصات محیط اجرا", 5)
        row += 1
        self._header(ws, row, ["مورد", "مقدار"])
        row += 1
        for key, value in report.environment.items():
            self._value_cell(ws, row, 1, key)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, row - len(report.environment), row - 1, 1, 2)
        self._widths(ws, {"A": 50, "B": 30, "C": 18, "D": 22, "E": 70})
        ws.freeze_panes = "A4"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _samples(self, wb: Workbook, context: ReportContext) -> None:
        monitor = context.monitor
        if monitor is None or not getattr(monitor, "samples", None):
            return
        ws = wb.create_sheet("نمونه‌های پایش")
        self._rtl(ws)
        self._title(ws, "نمونه‌های زمانی مصرف منابع (مقدار خالی = اندازه‌گیری نشده)", 8)
        headers = [
            "زمان سپری‌شده (ثانیه)",
            "فاز",
            "CPU سیستم %",
            "RAM مصرفی GB",
            "RAM آزاد GB",
            "CPU Ollama % کل CPU",
            "RAM Ollama MB",
            "RAM برنامه Python MB",
        ]
        self._header(ws, 3, headers)
        row = 4
        for sample in monitor.samples[:5000]:
            self._value_cell(ws, row, 1, sample.get("elapsed_sec"))
            self._value_cell(ws, row, 2, sample.get("phase"))
            self._value_cell(ws, row, 3, sample.get("system_cpu_pct"))
            self._value_cell(ws, row, 4, sample.get("ram_used_gb"))
            self._value_cell(ws, row, 5, sample.get("ram_available_gb"))
            self._value_cell(ws, row, 6, sample.get("ollama_cpu_machine_pct"))
            self._value_cell(ws, row, 7, sample.get("ollama_rss_mb"))
            self._value_cell(ws, row, 8, sample.get("python_rss_mb"))
            row += 1
        if row > 4:
            self._body(ws, 4, row - 1, 1, 8)
            ws.auto_filter.ref = f"A3:H{row - 1}"
            chart = LineChart()
            chart.title = "روند CPU سیستم و Ollama"
            chart.y_axis.title = "درصد"
            chart.x_axis.title = "ثانیه"
            chart.height = 8
            chart.width = 18
            chart.add_data(Reference(ws, min_col=3, max_col=3, min_row=3, max_row=row - 1), titles_from_data=True)
            chart.add_data(Reference(ws, min_col=6, max_col=6, min_row=3, max_row=row - 1), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=1, min_row=4, max_row=row - 1))
            ws.add_chart(chart, "J4")
        self._widths(ws, {"A": 22, "B": 26, "C": 14, "D": 16, "E": 14, "F": 20, "G": 16, "H": 20})
        ws.freeze_panes = "A4"
        self._print_setup(ws)

        gpu_samples = [sample for sample in getattr(monitor, "gpu_samples", []) if sample is not None]
        if not gpu_samples:
            return
        gpu_ws = wb.create_sheet("نمونه‌های GPU")
        self._rtl(gpu_ws)
        self._title(gpu_ws, "نمونه‌های شمارنده GPU (در دسترس نبود ≠ صفر)", 9)
        self._header(
            gpu_ws,
            3,
            [
                "فاز",
                "داده موجود",
                "Ollama Engine Max %",
                "Ollama Engine Sum %",
                "Ollama Dedicated MB",
                "Ollama Shared MB",
                "کل Engine Max %",
                "فاصله نمونه (ثانیه)",
                "خطا",
            ],
        )
        row = 4
        for sample in gpu_samples[:5000]:
            self._value_cell(gpu_ws, row, 1, sample.phase)
            self._value_cell(gpu_ws, row, 2, "بله" if sample.available else "خیر")
            self._value_cell(gpu_ws, row, 3, sample.ollama_engine_max_pct)
            self._value_cell(gpu_ws, row, 4, sample.ollama_engine_sum_pct)
            self._value_cell(gpu_ws, row, 5, sample.ollama_dedicated_mb)
            self._value_cell(gpu_ws, row, 6, sample.ollama_shared_mb)
            self._value_cell(gpu_ws, row, 7, sample.all_engine_max_pct)
            self._value_cell(gpu_ws, row, 8, sample.sample_interval_sec)
            self._value_cell(gpu_ws, row, 9, sample.error or "")
            row += 1
        self._body(gpu_ws, 4, row - 1, 1, 9)
        self._widths(gpu_ws, {"A": 26, "B": 12, "C": 22, "D": 22, "E": 22, "F": 22, "G": 20, "H": 18, "I": 50})
        gpu_ws.freeze_panes = "A4"
        self._print_setup(gpu_ws)

    # ------------------------------------------------------------------ #
    def _limitations(self, wb: Workbook, context: ReportContext) -> None:
        ws = wb.create_sheet("محدودیت‌ها و خطاها")
        self._rtl(ws)
        self._title(ws, "محدودیت‌های اندازه‌گیری، خطاها و هشدارهای پردازش", 4)
        row = 3
        self._section(ws, row, "محدودیت‌های اندازه‌گیری و تفسیر", 4)
        for limitation in (context.performance.limitations if context.performance else []):
            row += 1
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            cell = ws.cell(row, 1, "• " + limitation)
            cell.font = Font(name=FONT_NAME, size=9, color=DARK)
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            ws.row_dimensions[row].height = 28

        row += 2
        self._section(ws, row, "خطاها و هشدارهای ثبت‌شده", 4)
        self._header(ws, row + 1, ["نوع", "پیام", "برگه", "سلول"])
        row += 2
        start = row
        for error in (context.performance.errors if context.performance else []):
            self._value_cell(ws, row, 1, "خطای پایش")
            self._value_cell(ws, row, 2, error)
            row += 1
        for warning in context.analysis.warnings:
            self._value_cell(ws, row, 1, "هشدار پردازش")
            self._value_cell(ws, row, 2, warning)
            row += 1
        for sheet in context.analysis.meta.sheets:
            for item in sheet.suspicious:
                self._value_cell(ws, row, 1, "نکته ساختاری")
                self._value_cell(ws, row, 2, item)
                self._value_cell(ws, row, 3, sheet.title)
                row += 1
        if row == start:
            ws.cell(row, 1, "خطا یا هشدار ثبت‌شده‌ای وجود ندارد.")
            row += 1
        self._body(ws, start, row - 1, 1, 4)
        self._widths(ws, {"A": 22, "B": 90, "C": 26, "D": 14})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _structure(self, wb: Workbook, analysis: AnalysisResult) -> None:
        ws = wb.create_sheet("ساختار فایل")
        self._rtl(ws)
        self._title(ws, "ساختار شناسایی‌شده فایل و جدول‌ها", 9)
        self._header(
            ws,
            3,
            ["برگه", "وضعیت", "محدوده داده", "محدوده فرمت‌شده", "فرمول", "خطا", "ادغام", "مخفی", "نکات"],
        )
        row = 4
        for sheet in analysis.meta.sheets:
            self._value_cell(ws, row, 1, sheet.title)
            self._value_cell(ws, row, 2, "نمایان" if sheet.state == "visible" else sheet.state)
            self._value_cell(ws, row, 3, sheet.used_dimensions)
            self._value_cell(ws, row, 4, sheet.declared_dimensions)
            self._value_cell(ws, row, 5, sheet.formula_count)
            self._value_cell(ws, row, 6, len(sheet.error_cells))
            self._value_cell(ws, row, 7, sheet.merged_ranges)
            self._value_cell(ws, row, 8, f"{sheet.hidden_rows} ردیف / {sheet.hidden_cols} ستون")
            self._value_cell(ws, row, 9, "؛ ".join(sheet.suspicious) or "—")
            row += 1
        self._body(ws, 4, row - 1, 1, 9)

        row += 1
        self._section(ws, row, "جدول‌های شناسایی‌شده", 9)
        row += 1
        self._header(ws, row, ["برگه", "نوع", "محدوده", "ردیف داده", "اطمینان", "ستون‌ها و نقش", "دلیل تشخیص"])
        row += 1
        for table in analysis.tables:
            self._value_cell(ws, row, 1, table.sheet)
            self._value_cell(ws, row, 2, TABLE_KIND_FA.get(table.kind, table.kind.value))
            self._value_cell(ws, row, 3, table.address)
            self._value_cell(ws, row, 4, table.row_count)
            self._value_cell(ws, row, 5, f"{table.confidence:.0%}")
            columns = "؛ ".join(
                f"{column.header or column.letter}={ROLE_FA.get(column.role, column.role.value)}"
                for column in table.columns
            )
            self._value_cell(ws, row, 6, columns)
            self._value_cell(ws, row, 7, "؛ ".join(table.reasons) or "—")
            row += 1
        if not analysis.tables:
            ws.cell(row, 1, "جدولی شناسایی نشد.")
            row += 1
        self._body(ws, row - max(len(analysis.tables), 1), row - 1, 1, 7)
        self._widths(ws, {"A": 22, "B": 16, "C": 18, "D": 12, "E": 12, "F": 70, "G": 50, "H": 20, "I": 50})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _quality_gate_sheet(self, wb: Workbook, analysis: AnalysisResult) -> None:
        """Audit-only view of the central quality gate before final delivery."""

        from local_office_ai.quality import GATE_STATUS_FA, GateStatus

        gate = getattr(analysis, "quality_gate", None)
        ws = wb.create_sheet("دروازه کیفیت")
        self._rtl(ws)
        self._title(ws, "دروازه کیفیت یکپارچه — وضعیت هر مرحله", 6)
        ws["A2"] = (
            "این برگه شواهد مراحل ۱ تا ۱۱ را پیش از نوشتن سند نهایی نشان می‌دهد؛ "
            "درستی قالب خروجی و یکپارچگی سند در JSON نتیجه اجرا ثبت می‌شود."
        )
        ws.merge_cells("A2:F2")
        ws["A2"].font = Font(name=FONT_NAME, size=9, color=DARK)
        ws["A2"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        if gate is None:
            ws.cell(4, 1, "دروازه کیفیت برای این تحلیل اجرا نشد.")
            self._widths(ws, {"A": 34, "B": 40, "C": 14, "D": 14, "E": 14, "F": 70})
            self._print_setup(ws)
            return
        counts = gate.counts()
        self._header(
            ws,
            4,
            ["مرحله", "خلاصه شواهد", "وضعیت", "مسدودکننده", "تعداد بررسی", "یافته‌های غیرقبول"],
        )
        row = 5
        colors = {
            GateStatus.PASS: GREEN,
            GateStatus.WARNING: YELLOW,
            GateStatus.FAIL: RED,
        }
        for stage in gate.stages:
            self._value_cell(ws, row, 1, stage.label_fa)
            self._value_cell(ws, row, 2, stage.summary)
            self._value_cell(ws, row, 3, GATE_STATUS_FA.get(stage.status, stage.status.value))
            self._value_cell(ws, row, 4, "بله" if stage.blocking else "خیر")
            self._value_cell(ws, row, 5, len(stage.checks))
            issues = [check.message for check in stage.checks if check.status is not GateStatus.PASS]
            self._value_cell(ws, row, 6, "؛ ".join(issues[:5]) or "—")
            self._fill(ws.cell(row, 3), colors.get(stage.status, GREY))
            row += 1
        self._body(ws, 5, row - 1, 1, 6)

        row += 1
        self._section(ws, row, "جمع‌بندی دروازه کیفیت", 6)
        row += 1
        self._value_cell(ws, row, 1, "وضعیت کلی")
        self._value_cell(ws, row, 2, GATE_STATUS_FA.get(gate.status, gate.status.value))
        self._fill(ws.cell(row, 2), colors.get(gate.status, GREY))
        self._value_cell(
            ws,
            row,
            3,
            f"PASS {counts.get('PASS', 0)} / WARNING {counts.get('WARNING', 0)} / FAIL {counts.get('FAIL', 0)}",
        )
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        row += 1
        self._value_cell(ws, row, 1, "ادعای گزارش معتبر")
        self._value_cell(ws, row, 2, "بله" if gate.can_claim_valid_report else "خیر")
        self._fill(ws.cell(row, 2), GREEN if gate.can_claim_valid_report else RED)
        self._value_cell(
            ws,
            row,
            3,
            "بدون شکست مسدودکننده" if gate.can_claim_valid_report else "شکست مسدودکننده دارد",
        )
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
        row += 1
        if gate.provenance is not None:
            ledger_counts = gate.provenance.counts()
            self._value_cell(ws, row, 1, "رکوردهای ردیابی منبع")
            self._value_cell(
                ws,
                row,
                2,
                f"مهم: {ledger_counts.get('important', 0)} / کامل: {ledger_counts.get('complete', 0)}",
            )
            self._value_cell(
                ws,
                row,
                3,
                f"بدون منبع مشتق: {len(gate.provenance.derived_without_sources)}",
            )
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
            row += 1
        self._widths(ws, {"A": 34, "B": 52, "C": 16, "D": 14, "E": 14, "F": 80})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _provenance_sheet(self, wb: Workbook, analysis: AnalysisResult) -> None:
        """Audit-only, machine-readable trail from each result to its cells."""

        ledger = getattr(analysis, "provenance", None)
        ws = wb.create_sheet("ردیابی منبع (شواهد)")
        self._rtl(ws)
        self._title(ws, "شواهد ردیابی منبع — هر مقدار مهم تا سلول‌های مبنا", 14)
        headers = [
            "کلید",
            "فایل",
            "برگه",
            "بلوک",
            "سلول/محدوده",
            "عنوان اصلی",
            "عنوان نرمال‌شده",
            "مقدار اصلی",
            "مقدار نرمال‌شده",
            "دوره",
            "واحد",
            "عملیات",
            "فرمول/تحول",
            "منبع‌های وابسته",
            "وضعیت اعتبار",
        ]
        self._header(ws, 3, headers)
        records = list(getattr(ledger, "records", []) or []) if ledger is not None else []
        row = 4
        for record in records:
            self._value_cell(ws, row, 1, record.key)
            self._value_cell(ws, row, 2, record.filename)
            self._value_cell(ws, row, 3, record.sheet)
            self._value_cell(ws, row, 4, record.block)
            self._value_cell(ws, row, 5, record.cell)
            self._value_cell(ws, row, 6, record.original_label)
            self._value_cell(ws, row, 7, record.normalized_label)
            self._value_cell(ws, row, 8, record.original_value)
            self._value_cell(ws, row, 9, record.normalized_value)
            self._value_cell(ws, row, 10, record.period)
            self._value_cell(ws, row, 11, record.unit)
            self._value_cell(ws, row, 12, record.operation)
            self._value_cell(ws, row, 13, record.formula)
            sources = "؛ ".join(
                f"{source.sheet}!{source.cell}".strip("!")
                for source in record.sources
                if source.sheet or source.cell
            )
            self._value_cell(ws, row, 14, sources or "—")
            self._value_cell(ws, row, 15, record.validation_status)
            if not record.complete:
                self._fill(ws.cell(row, 5), YELLOW)
            row += 1
        if row == 4:
            ws.cell(4, 1, "هیچ رکورد ردیابی منبعی وجود ندارد.")
        else:
            self._body(ws, 4, row - 1, 1, 15)
            ws.auto_filter.ref = f"A3:O{row - 1}"
            ws.freeze_panes = "A4"
        widths = {
            "A": 34, "B": 22, "C": 18, "D": 20, "E": 16, "F": 26, "G": 30,
            "H": 22, "I": 24, "J": 18, "K": 12, "L": 26, "M": 34, "N": 48, "O": 14,
        }
        self._widths(ws, widths)
        self._print_setup(ws)

