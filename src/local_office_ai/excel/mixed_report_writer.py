"""Clean user-facing Excel report for mixed document analysis."""

from __future__ import annotations

import os
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY = "17365D"
BLUE = "2F5597"
LIGHT_BLUE = "DEEAF6"
PALE = "F4F7FB"
WHITE = "FFFFFF"
DARK = "1F2937"
GREEN = "E2F0D9"
YELLOW = "FFF2CC"
RED = "FCE4D6"
THIN = Side(style="thin", color="D8E1EC")


class MixedReportWriter:
    def build(self, result: Any, output: Path | str) -> Path:
        workbook = Workbook()
        workbook.remove(workbook.active)
        self._summary(workbook, result)
        self._sources(workbook, result)
        if result.statistics:
            self._statistics(workbook, result)
        if result.extracted_tables:
            self._tables(workbook, result)
        if result.extracted_text:
            self._text(workbook, result)
        workbook.properties.title = "Local Office AI - گزارش ترکیبی اسناد"
        workbook.properties.creator = "Local Office AI"
        workbook.properties.subject = "تحلیل داده و اسناد اداری"
        target = Path(output)
        target.parent.mkdir(parents=True, exist_ok=True)
        temp = target.with_name(target.stem + ".tmp" + target.suffix)
        try:
            workbook.save(temp)
            check = load_workbook(temp, read_only=True, data_only=False)
            if not check.sheetnames:
                raise ValueError("گزارش خروجی فاقد برگه است.")
            check.close()
            os.replace(temp, target)
        except Exception:
            temp.unlink(missing_ok=True)
            raise
        return target

    def _base(self, ws, title: str, columns: int) -> None:
        ws.sheet_view.rightToLeft = True
        ws.freeze_panes = "A4"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_view.showGridLines = False
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=columns)
        cell = ws.cell(1, 1, title)
        cell.font = Font(name="B Nazanin", size=17, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[1].height = 34
        ws.freeze_panes = "A4"

    def _header(self, ws, row: int, values: list[str]) -> None:
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row, col, value)
            cell.font = Font(name="B Nazanin", size=12, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=BLUE)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(bottom=THIN)
        ws.row_dimensions[row].height = 28

    @staticmethod
    def _value(value: Any) -> Any:
        if isinstance(value, Decimal):
            return float(value)
        if value is None:
            return "—"
        if isinstance(value, (str, int, float, bool)):
            return value
        return str(value)

    def _body(self, ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
        for row in ws.iter_rows(min_row=min_row, max_row=max(max_row, min_row), min_col=min_col, max_col=max_col):
            for cell in row:
                cell.font = Font(name="B Nazanin", size=11, color=DARK)
                cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
                cell.border = Border(bottom=THIN)
                if cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=PALE)

    @staticmethod
    def _widths(ws, widths: dict[str, int]) -> None:
        for column, width in widths.items():
            ws.column_dimensions[column].width = width

    def _summary(self, wb: Workbook, result: Any) -> None:
        ws = wb.create_sheet("خلاصه اجرایی")
        self._base(ws, "گزارش تحلیل چندمنبعی و اسناد اداری", 6)
        self._header(ws, 3, ["موضوع", "نتیجه"])
        values = [
            ("درخواست کاربر", result.request or "تحلیل عمومی منابع انتخاب‌شده"),
            ("تعداد منابع", len(result.files)),
            ("منابع موفق", sum(1 for item in result.files if item.status == "تحلیل شد")),
            ("منابع ناموفق", sum(1 for item in result.files if item.status != "تحلیل شد")),
            ("جدول‌های استخراج‌شده", len(result.extracted_tables)),
            ("شاخص‌های عددی محاسبه‌شده", len(result.statistics)),
        ]
        row = 4
        for label, value in values:
            ws.cell(row, 1, label)
            ws.cell(row, 2, self._value(value))
            row += 1
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        heading = ws.cell(row, 1, "تحلیل و جمع‌بندی")
        heading.font = Font(name="B Nazanin", size=13, bold=True, color=WHITE)
        heading.fill = PatternFill("solid", fgColor=BLUE)
        heading.alignment = Alignment(horizontal="right")
        row += 1
        narrative = result.ai_summary or result.fallback_summary or "برای این منابع جمع‌بندی متنی تولید نشد."
        for paragraph in str(narrative).splitlines() or [str(narrative)]:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            cell = ws.cell(row, 1, paragraph)
            cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
            cell.font = Font(name="B Nazanin", size=12, color=DARK)
            ws.row_dimensions[row].height = 36 if len(paragraph) > 110 else 24
            row += 1
        if result.warnings:
            row += 1
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            cell = ws.cell(row, 1, "موارد نیازمند توجه")
            cell.font = Font(name="B Nazanin", size=13, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor="A66B00")
            cell.alignment = Alignment(horizontal="right")
            row += 1
            for warning in result.warnings:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
                cell = ws.cell(row, 1, "• " + str(warning))
                cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
                row += 1
        self._body(ws, 4, row, 1, 2)
        self._widths(ws, {"A": 30, "B": 105, "C": 18, "D": 18, "E": 18, "F": 18})
        ws.auto_filter.ref = f"A3:B{max(3, row)}"

    def _sources(self, wb: Workbook, result: Any) -> None:
        ws = wb.create_sheet("فهرست منابع")
        self._base(ws, "فهرست فایل‌های بررسی‌شده", 8)
        headers = ["نام فایل", "نوع", "بخش/صفحه/اسلاید", "ردیف/محتوا", "ستون", "شاخص عددی", "وضعیت", "توضیح"]
        self._header(ws, 3, headers)
        for row, item in enumerate(result.files, start=4):
            values = [item.name, item.file_type, item.sheets, item.rows, item.columns, item.metrics, item.status, item.error or item.summary]
            for col, value in enumerate(values, start=1):
                ws.cell(row, col, self._value(value))
            ws.cell(row, 7).fill = PatternFill("solid", fgColor=GREEN if item.status == "تحلیل شد" else YELLOW)
        last = 3 + len(result.files)
        self._body(ws, 4, last, 1, len(headers))
        self._widths(ws, {"A": 38, "B": 18, "C": 16, "D": 15, "E": 12, "F": 14, "G": 16, "H": 70})
        if result.files:
            ws.auto_filter.ref = f"A3:H{last}"

    def _statistics(self, wb: Workbook, result: Any) -> None:
        ws = wb.create_sheet("شاخص‌های آماری")
        self._base(ws, "شاخص‌های عددی استخراج‌شده از منابع", 12)
        headers = ["فایل", "محل", "جدول", "ستون", "تعداد معتبر", "ناقص", "مجموع", "میانگین", "میانه", "انحراف معیار نمونه", "حداقل", "حداکثر"]
        self._header(ws, 3, headers)
        for row, item in enumerate(result.statistics, start=4):
            values = [item.get(key) for key in ("source", "location", "table", "column", "n", "missing", "sum", "mean", "median", "std_sample", "min", "max")]
            for col, value in enumerate(values, start=1):
                ws.cell(row, col, self._value(value))
        last = 3 + len(result.statistics)
        self._body(ws, 4, last, 1, len(headers))
        for row in range(4, last + 1):
            for col in range(7, 13):
                ws.cell(row, col).number_format = '#,##0.00;[Red]-#,##0.00'
        self._widths(ws, {"A": 34, "B": 22, "C": 28, "D": 26, "E": 14, "F": 12, "G": 18, "H": 18, "I": 18, "J": 22, "K": 18, "L": 18})
        if result.statistics:
            ws.auto_filter.ref = f"A3:L{last}"

    def _tables(self, wb: Workbook, result: Any) -> None:
        ws = wb.create_sheet("جداول استخراج‌شده")
        self._base(ws, "ردیف‌های جدول‌های استخراج‌شده", 6)
        self._header(ws, 3, ["فایل", "محل", "عنوان جدول", "ردیف", "عنوان ستون‌ها", "مقادیر ردیف"])
        row = 4
        for table in result.extracted_tables:
            for index, values in enumerate(table.rows, start=1):
                ws.cell(row, 1, table.source.name)
                ws.cell(row, 2, table.location)
                ws.cell(row, 3, table.title)
                ws.cell(row, 4, index)
                ws.cell(row, 5, " | ".join(table.headers))
                ws.cell(row, 6, " | ".join(str(value or "") for value in values))
                row += 1
                if row > 15004:
                    break
            if row > 15004:
                break
        self._body(ws, 4, row - 1, 1, 6)
        self._widths(ws, {"A": 34, "B": 18, "C": 28, "D": 10, "E": 65, "F": 90})
        ws.auto_filter.ref = f"A3:F{max(3, row - 1)}"

    def _text(self, wb: Workbook, result: Any) -> None:
        ws = wb.create_sheet("متن استخراج‌شده")
        self._base(ws, "متن استخراج‌شده از اسناد", 3)
        self._header(ws, 3, ["فایل", "محل", "متن"])
        row = 4
        for source, location, text in result.extracted_text:
            ws.cell(row, 1, source)
            ws.cell(row, 2, location)
            ws.cell(row, 3, text)
            ws.row_dimensions[row].height = 60
            row += 1
            if row > 2004:
                break
        self._body(ws, 4, row - 1, 1, 3)
        self._widths(ws, {"A": 36, "B": 28, "C": 110})
        ws.auto_filter.ref = f"A3:C{max(3, row - 1)}"
