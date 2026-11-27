"""The sheets that answer the user's Persian request directly.

Three sheets, all built from :class:`PeriodIntegration` and the grounded
answers: the request itself with its answers, the monthly matrix, and the
period comparison.  Charts are drawn from the same cells the reader sees, so a
chart can never show a number that is not in the table.
"""

from __future__ import annotations

from decimal import Decimal
from typing import Any, Iterable

from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

from local_office_ai.core.models import to_excel_number
from local_office_ai.excel.data_qa import comparison_pair
from local_office_ai.excel.period_integration import MetricSeries, PeriodIntegration

FONT_NAME = "Tahoma"
DARK = "1F1F1F"
GREEN = "E2EFDA"
RED = "FCE4D6"
YELLOW = "FFF2CC"
GREY = "F2F2F2"
COUNT_FORMAT = "#,##0;[Red]-#,##0"
PERCENT_FORMAT = "0.00%;[Red]-0.00%"


class RequestReportMixin:
    # ------------------------------------------------------------------ #
    def _request_sheet(self, wb, context) -> None:
        analysis = context.analysis
        intent = getattr(analysis, "request_intent", None)
        answers = getattr(analysis, "request_answers", None) or []
        integration = getattr(analysis, "integration", None)
        ws = wb.create_sheet("پاسخ درخواست")
        self._rtl(ws)
        self._title(ws, "پاسخ به درخواست کاربر", 6)
        row = 3
        self._header(ws, row, ["بخش", "مقدار"])
        row += 1
        request_text = getattr(intent, "text", "") or "درخواستی ثبت نشده است."
        rows = [
            ("درخواست کاربر", request_text),
            (
                "تفسیر درخواست",
                "، ".join(getattr(intent, "scope_labels", [])) or "تحلیل استاندارد",
            ),
            (
                "شاخص‌های هدف",
                "، ".join(getattr(intent, "matched_metrics", [])) or "همه شاخص‌های عددی",
            ),
        ]
        if integration is not None:
            rows.append(("ماه‌های تحلیل‌شده", "، ".join(integration.months) or "—"))
            rows.append(
                ("دوره‌های تحلیل‌شده", integration.periods_summary())
            )
            rows.append(
                (
                    "برگه‌های دوره‌ای استفاده‌شده",
                    "، ".join(integration.sources) or "—",
                )
            )
        for label, value in rows:
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, 4, row - 1, 1, 2)

        row += 1
        self._section(ws, row, "پاسخ‌های محاسبه‌شده با Python", 6)
        row += 1
        self._header(ws, row, ["#", "عنوان", "پاسخ", "مقدار", "منبع محاسبه", "توضیح"])
        header_row = row
        row += 1
        if not answers:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            ws.cell(row, 1, "پاسخی برای این درخواست تولید نشد؛ برگه‌های تحلیل را ببینید.").alignment = Alignment(
                horizontal="right", wrap_text=True
            )
            row += 1
        for index, answer in enumerate(answers, start=1):
            self._value_cell(ws, row, 1, index)
            self._value_cell(ws, row, 2, answer.title)
            self._value_cell(ws, row, 3, answer.text)
            self._value_cell(ws, row, 4, answer.value)
            self._value_cell(ws, row, 5, answer.source)
            self._value_cell(ws, row, 6, answer.detail)
            row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 6)
        for index in range(header_row + 1, row):
            ws.row_dimensions[index].height = 42

        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        cell = ws.cell(
            row,
            1,
            "همه اعداد این برگه با Python و Decimal از سلول‌های فایل محاسبه شده‌اند؛ "
            "مدل زبانی فقط توضیح می‌دهد و هیچ عددی از مدل پذیرفته نمی‌شود.",
        )
        cell.alignment = Alignment(horizontal="right", wrap_text=True)
        cell.font = Font(name=FONT_NAME, size=9, color=DARK)
        self._widths(ws, {"A": 6, "B": 34, "C": 96, "D": 22, "E": 44, "F": 52})
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _monthly_sheet(self, wb, context) -> None:
        integration: PeriodIntegration = getattr(context.analysis, "integration", None)
        ws = wb.create_sheet("جدول ماهانه شاخص‌ها")
        self._rtl(ws)
        self._title(ws, "جدول ماهانه شاخص‌ها؛ جمع هر ماه از سلول‌های همان برگه بازمحاسبه شده است", 8)
        row = 3
        if integration is None or not integration.series:
            self._value_cell(ws, row, 1, "ساختار ماهانه برای این فایل شناسایی نشد.")
            self._print_setup(ws)
            return

        headline = integration.headline_metric()
        for index, period in enumerate(integration.periods):
            months = [
                month
                for month in integration.months
                if any(series.cell(period, month) for series in integration.series)
            ]
            if not months:
                continue
            suspicious = " (برچسب سال مشکوک)" if period in integration.suspicious_periods else ""
            partial = [
                month
                for month in months
                if any(
                    (cell := series.cell(period, month)) is not None and not cell.complete
                    for series in integration.series
                )
            ]
            complete_months = len(months) - len(partial)
            label = f"دوره {period}{suspicious} — "
            label += (
                f"{complete_months} ماه کامل + {len(partial)} ماه ناقص"
                if partial
                else f"{len(months)} ماه کامل"
            )
            self._section(ws, row, label, 8 + len(months))
            row += 1
            self._header(
                ws,
                row,
                ["شاخص معیار"]
                + [
                    month
                    + (" (ناقص)" if self._month_incomplete(integration, period, month) else "")
                    for month in months
                ]
                + ["جمع دوره", "میانگین ماهانه", "بیشترین ماه", "کمترین ماه"],
            )
            header_row = row
            row += 1
            chart_row = 0
            for series in integration.series:
                cells = [series.cell(period, month) for month in months]
                if not any(cell is not None and cell.total is not None for cell in cells):
                    continue
                self._value_cell(ws, row, 1, series.metric)
                if headline is not None and series.metric == headline.metric and chart_row == 0:
                    chart_row = row
                for offset, cell in enumerate(cells, start=2):
                    if cell is None:
                        self._value_cell(ws, row, offset, "—")
                        continue
                    value_cell = ws.cell(row, offset, to_excel_number(cell.total))
                    value_cell.number_format = COUNT_FORMAT
                    if not cell.complete or cell.total is None:
                        self._fill(value_cell, YELLOW)
                total = series.total(period)
                self._value_cell(ws, row, 2 + len(months), total, money=False)
                mean = series.mean(period)
                if mean is not None:
                    mean = mean.quantize(Decimal("0.01"))
                self._value_cell(ws, row + 0, 3 + len(months), mean)
                best = series.best(period)
                worst = series.worst(period)
                self._value_cell(
                    ws, row, 4 + len(months), _extreme_label(best, worst),
                )
                self._value_cell(
                    ws, row, 5 + len(months), _extreme_label(worst, best),
                )
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 5 + len(months))
            for column in range(2, 2 + len(months)):
                for body_row in range(header_row + 1, row):
                    ws.cell(body_row, column).number_format = COUNT_FORMAT
            for column in (2 + len(months), 3 + len(months)):
                for body_row in range(header_row + 1, row):
                    ws.cell(body_row, column).number_format = COUNT_FORMAT

            if chart_row and len(months) >= 2:
                chart = LineChart()
                chart.title = f"روند ماهانه «{headline.metric}» — دوره {period}"
                chart.y_axis.title = "مقدار"
                chart.x_axis.title = "ماه"
                data = Reference(
                    ws, min_col=2, max_col=1 + len(months), min_row=chart_row, max_row=chart_row
                )
                categories = Reference(ws, min_col=2, max_col=1 + len(months), min_row=header_row, max_row=header_row)
                chart.add_data(data, from_rows=True, titles_from_data=False)
                chart.set_categories(categories)
                chart.height = 8.5
                chart.width = 20
                chart.legend = None
                ws.add_chart(chart, f"A{row + 1}")
                row += 18

            missing = [
                month
                for month in months
                if self._month_incomplete(integration, period, month)
            ]
            if missing:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
                ws.cell(
                    row,
                    1,
                    "ماه ناقص (زرد): " + "، ".join(missing)
                    + " — مقدار نمایش داده می‌شود اما در جمع دوره، میانگین و رتبه‌بندی لحاظ نمی‌شود.",
                ).alignment = Alignment(horizontal="right", wrap_text=True)
                row += 1
            row += 2

        excluded = [name for name in integration.excluded_sources if name]
        if excluded:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.cell(
                row,
                1,
                "برگه‌های مشتق یا بدون دوره در ماتریس ماهانه جمع نشدند: " + "، ".join(excluded)
                + "؛ این برگه‌ها خودشان از فرمول برگه‌های دیگر ساخته می‌شوند.",
            ).alignment = Alignment(horizontal="right", wrap_text=True)
        self._widths(ws, {"A": 46, **{chr(66 + index): 16 for index in range(12)}})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _period_compare_sheet(self, wb, context) -> None:
        integration: PeriodIntegration = getattr(context.analysis, "integration", None)
        ws = wb.create_sheet("مقایسه دوره‌ای شاخص‌ها")
        self._rtl(ws)
        self._title(ws, "مقایسه دوره‌ها فقط روی ماه‌های مشترک و کامل", 10)
        row = 3
        intent = getattr(context.analysis, "request_intent", None)
        pair = comparison_pair(intent, integration) if integration is not None else None
        if integration is None or pair is None:
            self._value_cell(ws, row, 1, "برای مقایسه دوره‌ای، حداقل دو دوره هم‌تعریف لازم است.")
            self._print_setup(ws)
            return
        period_a, period_b = pair
        self._section(ws, row, f"مقایسه دوره {period_a} (پایه) با دوره {period_b}", 10)
        row += 1
        self._header(
            ws,
            row,
            [
                "شاخص", f"جمع {period_a}", f"جمع {period_b}", "اختلاف", "درصد تغییر",
                "جهت", "ماه‌های مشترک", "ماه کنارگذاشته", "توضیح",
            ],
        )
        header_row = row
        row += 1
        chart_rows = 0
        for series in integration.series:
            change = series.compare(period_a, period_b)
            if change.total_a is None and change.total_b is None:
                continue
            self._value_cell(ws, row, 1, series.metric)
            self._value_cell(ws, row, 2, change.total_a)
            self._value_cell(ws, row, 3, change.total_b)
            self._value_cell(ws, row, 4, change.difference)
            percent_cell = ws.cell(row, 5, to_excel_number(change.percent))
            percent_cell.number_format = PERCENT_FORMAT
            if change.percent is None:
                percent_cell.value = "—"
            self._value_cell(ws, row, 6, change.direction)
            self._value_cell(ws, row, 7, "، ".join(change.months_used) or "—")
            self._value_cell(ws, row, 8, "، ".join(change.incomplete_excluded) or "—")
            self._value_cell(ws, row, 9, change.note)
            color = GREEN if (change.difference or 0) > 0 else (RED if (change.difference or 0) < 0 else GREY)
            self._fill(ws.cell(row, 6), color)
            if chart_rows < 8 and change.total_a is not None and change.total_b is not None:
                chart_rows += 1
            row += 1
        self._body(ws, header_row + 1, row - 1, 1, 9)
        for body_row in range(header_row + 1, row):
            for column in (2, 3, 4):
                ws.cell(body_row, column).number_format = COUNT_FORMAT

        if chart_rows:
            chart = BarChart()
            chart.type = "col"
            chart.title = f"مقایسه {period_a} و {period_b} — ماه‌های مشترک"
            chart.y_axis.title = "جمع دوره"
            data = Reference(ws, min_col=2, max_col=3, min_row=header_row, max_row=header_row + chart_rows)
            categories = Reference(ws, min_col=1, min_row=header_row + 1, max_row=header_row + chart_rows)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            chart.height = 9
            chart.width = 20
            chart.grouping = "clustered"
            ws.add_chart(chart, f"A{row + 1}")
            row += 19

        row += 1
        self._section(ws, row, "کنترل صحت جمع ماهانه", 10)
        row += 1
        self._header(ws, row, ["بررسی", "وضعیت", "تعداد اثر", "توضیح"])
        header_row = row
        row += 1
        for check in integration.checks:
            self._value_cell(ws, row, 1, check.label)
            status_fa = {"passed": "قبول", "warning": "هشدار", "failed": "رد"}.get(check.status, check.status)
            self._value_cell(ws, row, 2, status_fa)
            self._value_cell(ws, row, 3, check.affected)
            self._value_cell(ws, row, 4, check.detail)
            self._fill(
                ws.cell(row, 2),
                {"passed": GREEN, "warning": YELLOW, "failed": RED}.get(check.status, GREY),
            )
            row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 4)

        if integration.alignment:
            row += 2
            self._section(ws, row, "ناهم‌ترازی سال در برگه خلاصه فایل", 10)
            row += 1
            self._header(
                ws,
                row,
                ["برگه خلاصه", "ستون", "سال اعلامی", "برگه مرجع", "ستون مرجع", "سال مرجع", "تعداد سلول"],
            )
            header_row = row
            row += 1
            for item in integration.alignment[:15]:
                self._value_cell(ws, row, 1, item["summary_sheet"])
                self._value_cell(ws, row, 2, item["summary_column"])
                self._value_cell(ws, row, 3, item["summary_period"])
                self._value_cell(ws, row, 4, item["referenced_sheet"])
                self._value_cell(ws, row, 5, item["referenced_column"])
                self._value_cell(ws, row, 6, item["referenced_period"])
                self._value_cell(ws, row, 7, item["cells"])
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 7)
        self._widths(ws, {"A": 46, "B": 18, "C": 18, "D": 16, "E": 14, "F": 12, "G": 30, "H": 26, "I": 60})
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    @staticmethod
    def _month_incomplete(integration: PeriodIntegration, period: str, month: str) -> bool:
        for series in integration.series:
            cell = series.cell(period, month)
            if cell is not None:
                return not cell.complete
        return False


def _fmt(value: Decimal | None) -> str:
    if value is None:
        return "—"
    if value == value.to_integral_value():
        return f"{int(value):,}"
    return f"{value.quantize(Decimal('0.01')):,f}"


def _extreme_label(target, other) -> str:
    """Best/worst month label; a flat series has no most or least month."""

    if target is None:
        return "—"
    if other is not None and target.total == other.total:
        return "همه ماه‌ها برابر"
    return f"{target.month} ({_fmt(target.total)})"
