"""Generic statistical report sheets, mixed into the Excel writer.

Keeping these sheets in a mixin avoids touching the established loan report
methods while still using the same styling helpers.
"""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from local_office_ai.core.models import (
    AnalysisResult,
    ColumnRole,
    DataRole,
    TableKind,
    to_excel_number,
)
from local_office_ai.excel.structure_intelligence import (
    ROW_ROLE_FA,
    STATUS_FA,
    STATUS_FULL,
    STATUS_NONE,
    STATUS_PARTIAL,
)

FONT_NAME = "Tahoma"
DARK = "1F1F1F"
GREEN = "E2EFDA"
YELLOW = "FFF2CC"
RED = "FCE4D6"
GREY = "F2F2F2"
TYPE_FA = {
    "number": "عددی",
    "date": "تاریخ",
    "text": "متنی",
    "boolean": "دودویی",
    "empty": "خالی",
    "error": "خطا",
}
ROLE_FA = {
    ColumnRole.INSTITUTION: "مؤسسه/طرف حساب",
    ColumnRole.PRINCIPAL: "اصل وام",
    ColumnRole.INSTALLMENT: "قسط",
    ColumnRole.INSTALLMENT_COUNT: "تعداد اقساط",
    ColumnRole.DATE: "تاریخ",
    ColumnRole.DESCRIPTION: "شرح",
    ColumnRole.AMOUNT: "مبلغ/مقدار",
    ColumnRole.COUNT: "تعداد",
    ColumnRole.ID: "شناسه",
    ColumnRole.UNKNOWN: "نامشخص",
}
KIND_FA = {
    TableKind.LOAN: "جدول وام (ماژول اختیاری)",
    TableKind.OBLIGATION: "جدول تعهدات (ماژول اختیاری)",
    TableKind.TIME_SERIES: "سری زمانی",
    TableKind.KPI: "شاخص‌ها",
    TableKind.GENERIC: "جدول عمومی",
}
DATA_ROLE_FA = {
    DataRole.IDENTIFIER: "شناسه",
    DataRole.TEMPORAL_KEY: "کلید زمانی",
    DataRole.NUMERIC_MEASURE: "متغیر اندازه‌گیری",
    DataRole.CATEGORICAL: "دسته‌بندی",
    DataRole.GROUP: "گروه‌بندی",
    DataRole.TEXT: "متنی",
    DataRole.CALCULATED: "فیلد محاسباتی",
    DataRole.UNKNOWN: "نامشخص",
}
AUDIT_STATUS_FA = {
    "verified": "تأیید شد",
    "mismatch": "مغایرت",
    "not_verifiable": "قابل بازتولید نبود",
    "no_declared_value": "مقدار اعلامی ندارد",
}


class GenericReportMixin:
    # ------------------------------------------------------------------ #
    def _workbook_profile(self, wb, analysis: AnalysisResult) -> None:
        intelligence = analysis.intelligence
        ws = wb.create_sheet("شناسنامه Workbook")
        self._rtl(ws)
        self._title(ws, "شناسنامه Workbook و پوشش تحلیل برگه‌ها", 14)
        if intelligence is None:
            self._value_cell(ws, 3, 1, "اطلاعات ساختاری Workbook در دسترس نیست.")
            self._print_setup(ws)
            return
        counts = intelligence.status_counts
        row = 3
        self._header(ws, row, ["موضوع", "مقدار"])
        row += 1
        info = [
            ("نام فایل", Path(analysis.meta.path).name),
            ("هش فایل", analysis.meta.sha256),
            ("تعداد برگه", intelligence.sheet_count),
            ("برگه‌های کاملاً تحلیل‌شده", counts.get(STATUS_FULL, 0)),
            ("برگه‌های جزئی تحلیل‌شده", counts.get(STATUS_PARTIAL, 0)),
            ("برگه‌های تحلیل‌نشده", counts.get(STATUS_NONE, 0)),
            ("گروه‌های ساختار مشابه", len(intelligence.similar_groups)),
            ("تعارض متادیتا", len(intelligence.metadata_conflicts)),
            ("مغایرت جمع/مقایسه", sum(1 for item in intelligence.audits if item.status.value == "mismatch")),
            ("جمع‌های قابل بازتولید نبود", sum(1 for item in intelligence.audits if item.status.value == "not_verifiable")),
        ]
        for label, value in info:
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, 4, row - 1, 1, 2)

        row += 1
        self._section(ws, row, "وضعیت هر برگه", 14)
        row += 1
        self._header(ws, row, [
            "برگه", "وضعیت", "ماه/دوره", "سال نام برگه", "سال عنوان",
            "سال زیرعنوان", "ردیف داده", "ستون", "پوشش", "مغایرت", "توضیح",
        ])
        header_row = row
        row += 1
        for sheet in intelligence.sheets:
            coverage = sheet.coverage
            coverage_text = "—"
            if coverage is not None:
                if coverage.expected_observations:
                    coverage_text = f"{coverage.observed_observations}/{coverage.expected_observations}"
                else:
                    coverage_text = f"{coverage.observed_observations} مشاهده"
            mismatches = sum(1 for item in sheet.audits if item.status.value == "mismatch")
            values = [
                sheet.sheet,
                STATUS_FA.get(sheet.status, sheet.status),
                sheet.month or "—",
                "، ".join(sheet.sheet_years) or "—",
                "، ".join(sheet.title_years) or "—",
                "، ".join(sheet.subheader_years) or "—",
                sheet.row_count,
                sheet.column_count,
                coverage_text,
                mismatches,
                "؛ ".join(sheet.warnings[:2]) or "—",
            ]
            for offset, value in enumerate(values, start=1):
                self._value_cell(ws, row, offset, value)
            row += 1
        self._body(ws, header_row + 1, row - 1, 1, 11)

        if intelligence.similar_groups:
            row += 2
            self._section(ws, row, "گروه‌های ساختار مشابه", 14)
            row += 1
            self._header(ws, row, ["تعداد ستون", "برگه‌ها", "زوج سال هر برگه", "توضیح"])
            row += 1
            for group in intelligence.similar_groups:
                pairs = " | ".join(
                    f"{item['sheet']}: {'، '.join(item['years']) or '—'}"
                    for item in group["period_pairs"]
                )
                self._value_cell(ws, row, 1, group["column_count"])
                self._value_cell(ws, row, 2, "، ".join(group["sheets"]))
                self._value_cell(ws, row, 3, pairs)
                self._value_cell(ws, row, 4, group["note"])
                row += 1

        row += 2
        self._section(ws, row, "هشدارهای سطح Workbook", 14)
        row += 1
        for warning in intelligence.warnings:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
            ws.cell(row, 1, "• " + warning).alignment = Alignment(horizontal="right", wrap_text=True)
            row += 1
        self._widths(ws, {
            "A": 26, "B": 22, "C": 14, "D": 18, "E": 18, "F": 20,
            "G": 12, "H": 10, "I": 14, "J": 10, "K": 60,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _fields_sheet(self, wb, analysis: AnalysisResult) -> None:
        intelligence = analysis.intelligence
        ws = wb.create_sheet("ساختار جداول و فیلدها")
        self._rtl(ws)
        self._title(ws, "ساختار سلسله‌مراتبی جدول‌ها، عنوان والد/فرعی و نقش هر ستون", 15)
        row = 3
        if intelligence is None:
            self._value_cell(ws, row, 1, "ساختار ستون‌ها در دسترس نیست.")
            self._print_setup(ws)
            return
        for sheet in intelligence.sheets:
            self._section(
                ws, row,
                f"برگه «{sheet.sheet}» — جدول {sheet.table_address or '—'} — {sheet.row_count} ردیف × {sheet.column_count} ستون",
                15,
            )
            row += 1
            self._header(ws, row, [
                "ستون", "آدرس عنوان", "عنوان اصلی", "عنوان والد", "عنوان فرعی",
                "دوره", "نام استاندارد", "نوع داده", "نقش معنایی", "نقش آماری",
                "معتبر", "پوشش", "اطمینان", "یادداشت", "برگه",
            ])
            header_row = row
            row += 1
            for column in sheet.columns:
                values = [
                    column.letter,
                    column.header_address,
                    column.original_header or column.header,
                    column.parent_header or "—",
                    column.sub_header or "—",
                    column.period or "—",
                    column.header,
                    TYPE_FA.get(column.data_type.value, column.data_type.value),
                    ROLE_FA.get(column.role, column.role.value),
                    DATA_ROLE_FA.get(column.statistical_role, column.statistical_role.value),
                    column.non_empty,
                    f"{column.non_empty / max(1, sheet.row_count):.0%}",
                    f"{column.confidence:.0%}",
                    "؛ ".join([*column.notes, *column.samples[:2]]),
                    sheet.sheet,
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 15)
            row += 2
        self._widths(ws, {
            "A": 8, "B": 12, "C": 34, "D": 34, "E": 12, "F": 10,
            "G": 40, "H": 10, "I": 14, "J": 16, "K": 9, "L": 9,
            "M": 9, "N": 55, "O": 16,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _blocks_sheet(self, wb, analysis: AnalysisResult) -> None:
        intelligence = analysis.intelligence
        ws = wb.create_sheet("بلوک‌های استخراج")
        self._rtl(ws)
        self._title(ws, "همه بلوک‌های داده‌ای، مرزهای آن‌ها و جمع‌های میانی/پایانی", 14)
        row = 3
        if intelligence is None:
            self._value_cell(ws, row, 1, "اطلاعات بلوک در دسترس نیست.")
            self._print_setup(ws)
            return
        self._header(ws, row, [
            "برگه", "شناسه بلوک", "آدرس", "نوع", "عنوان", "ردیف سرستون",
            "اولین داده", "آخرین داده", "پایان بلوک", "تعداد ردیف", "تعداد ستون",
            "شاخص‌ها", "جمع‌های میانی", "جمع کل", "بلوک اصل", "اطمینان", "هشدار",
        ])
        header_row = row
        row += 1
        for sheet in intelligence.sheets:
            for block in sheet.block_analyses:
                summary = block.get("summary_rows", {})
                values = [
                    sheet.sheet,
                    block.get("block_id", ""),
                    block.get("block_address") or block.get("address", ""),
                    block.get("kind", ""),
                    block.get("title", ""),
                    block.get("header_row", ""),
                    block.get("first_data_row", ""),
                    block.get("last_data_row", ""),
                    block.get("block_end_row", ""),
                    block.get("row_count", ""),
                    block.get("column_count", ""),
                    "، ".join(block.get("measures", [])[:6]),
                    "، ".join(str(item) for item in summary.get("subtotal", [])),
                    "، ".join(str(item) for item in summary.get("total", [])),
                    "بله" if block.get("is_primary") else "—",
                    f"{float(block.get('confidence', 0)):.0%}",
                    "؛ ".join(block.get("warnings", [])[:3]),
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 17)

        relationships = [
            item for sheet in intelligence.sheets for item in sheet.block_relationships
        ]
        if relationships:
            row += 2
            self._section(ws, row, "رابطه‌های شناسایی‌شده بین بلوک‌ها", 14)
            row += 1
            for item in relationships:
                text = (
                    f"{'، '.join(item.get('blocks', []))} — {item.get('note', '')}"
                    f"{' (مجاور)' if item.get('adjacent') else ''}"
                )
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
                ws.cell(row, 1, "• " + text).alignment = Alignment(horizontal="right", wrap_text=True)
                row += 1
        self._widths(ws, {
            "A": 18, "B": 16, "C": 18, "D": 14, "E": 26, "F": 10,
            "G": 10, "H": 10, "I": 10, "J": 9, "K": 9, "L": 34,
            "M": 14, "N": 14, "O": 10, "P": 10, "Q": 40,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _coverage_sheet(self, wb, analysis: AnalysisResult) -> None:
        intelligence = analysis.intelligence
        ws = wb.create_sheet("کیفیت و پوشش داده")
        self._rtl(ws)
        self._title(ws, "کیفیت و پوشش داده: صفر واقعی، خالی، نامعتبر، خطا و دوره تکمیل‌نشده", 11)
        row = 3
        self._header(ws, row, [
            "برگه", "دوره انتظار", "مشاهده‌شده", "بدون مقدار", "اولین روز", "آخرین روز",
            "آخرین مشاهده تقویمی (کاندید)", "صفر واقعی (سلول)", "سلول خالی", "مقدار نامعتبر",
            "خطای فرمول",
        ])
        header_row = row
        row += 1
        if intelligence is not None:
            for sheet in intelligence.sheets:
                coverage = sheet.coverage
                values = [
                    sheet.sheet,
                    coverage.expected_observations if coverage and coverage.expected_observations else "—",
                    coverage.observed_observations if coverage else 0,
                    coverage.missing_observations if coverage and coverage.missing_observations is not None else "—",
                    coverage.first_period if coverage else "—",
                    coverage.last_period if coverage else "—",
                    self._last_calendar_observation(sheet),
                    coverage.explicit_zero_cells if coverage else 0,
                    coverage.blank_cells if coverage else 0,
                    coverage.invalid_cells if coverage else 0,
                    coverage.error_cells if coverage else 0,
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 11)

        if intelligence is not None:
            row += 2
            self._section(ws, row, "روزهای بدون مقدار و توضیح پوشش", 10)
            row += 1
            for sheet in intelligence.sheets:
                coverage = sheet.coverage
                if coverage is None:
                    continue
                if coverage.missing_labels:
                    text = (
                        f"{sheet.sheet}: روزهای بدون مقدار → "
                        + "، ".join(coverage.missing_labels[:40])
                        + (" ..." if len(coverage.missing_labels) > 40 else "")
                    )
                    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
                    ws.cell(row, 1, "• " + text).alignment = Alignment(horizontal="right", wrap_text=True)
                    row += 1
                for note in coverage.notes:
                    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
                    ws.cell(row, 1, f"• {sheet.sheet}: {note}").alignment = Alignment(horizontal="right", wrap_text=True)
                    row += 1
        self._widths(ws, {"A": 22, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 28, "H": 14, "I": 12, "J": 14, "K": 12})
        self._print_setup(ws)

    @staticmethod
    def _last_calendar_observation(sheet) -> str:
        candidates = []
        for item in sheet.reconstructed_dates:
            last = item.get("last_jalali") or item.get("last_gregorian")
            if last:
                candidates.append(f"{item.get('period', '')}: {last}".strip(": "))
        return "؛ ".join(candidates) if candidates else "—"

    # ------------------------------------------------------------------ #
    def _period_comparison_sheet(self, wb, analysis: AnalysisResult) -> None:
        intelligence = analysis.intelligence
        ws = wb.create_sheet("مقایسه دوره‌ها")
        self._rtl(ws)
        self._title(ws, "مقایسه دوره‌های هم‌تراز، جمع بین‌برگه‌ای ایمن و تاریخ‌های بازسازی‌شده", 10)
        row = 3
        if intelligence is None:
            self._value_cell(ws, row, 1, "اطلاعات دوره‌ای در دسترس نیست.")
            self._print_setup(ws)
            return
        self._section(ws, row, "مقایسه دو دوره در هر برگه (فقط شاخص‌های هم‌تعریف)", 10)
        row += 1
        self._header(ws, row, [
            "برگه", "شاخص", "دوره پایه", "دوره مقایسه", "جمع پایه", "جمع مقایسه",
            "میانگین پایه", "میانگین مقایسه", "اختلاف", "درصد تغییر",
        ])
        header_row = row
        row += 1
        for sheet in intelligence.sheets:
            for item in sheet.period_comparisons:
                values = [
                    item.sheet, item.metric, item.period_a, item.period_b,
                    item.total_a, item.total_b, item.mean_a, item.mean_b,
                    item.difference, item.percent,
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 10)

        row += 2
        self._section(ws, row, "جمع بین‌برگه‌ای (برگه‌های غیرمشتق، ماه‌های متمایز و دوره‌های کامل)", 10)
        row += 1
        self._header(ws, row, ["شاخص", "دوره", "جمع", "تعداد مشاهده", "برگه‌ها", "ماه‌ها", "دوره ناقص کنارگذاشته", "توضیح"])
        header_row = row
        row += 1
        for item in intelligence.aggregations:
            values = [
                item.metric, item.period, item.total, item.observations,
                "، ".join(item.sheets), "، ".join(item.months),
                "، ".join(item.excluded_incomplete) or "—", item.note,
            ]
            for offset, value in enumerate(values, start=1):
                self._value_cell(ws, row, offset, value)
            row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 8)

        row += 2
        self._section(ws, row, "تاریخ‌های بازسازی‌شده (فقط با شواهد صریح ماه و سال)", 10)
        row += 1
        self._header(ws, row, ["برگه", "دوره", "ماه", "اولین روز", "آخرین روز", "میلادی اول", "میلادی آخر", "اطمینان", "یادداشت"])
        row += 1
        found = False
        for sheet in intelligence.sheets:
            for item in sheet.reconstructed_dates:
                found = True
                values = [
                    sheet.sheet, item["period"], item["month"], item["first_jalali"],
                    item["last_jalali"], item["first_gregorian"], item["last_gregorian"],
                    f"{item['confidence']:.0%}", item["note"],
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
        if not found:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
            ws.cell(row, 1, "تاریخ ترکیبی با اطمینان کافی بازسازی نشد؛ مقدار روز بدون حدس نگه داشته شد.").alignment = Alignment(horizontal="right", wrap_text=True)
        self._widths(ws, {"A": 22, "B": 38, "C": 12, "D": 12, "E": 16, "F": 16, "G": 16, "H": 12, "I": 16, "J": 12})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _reconciliation_sheet(self, wb, analysis: AnalysisResult) -> None:
        intelligence = analysis.intelligence
        ws = wb.create_sheet("مغایرت‌های محاسباتی")
        self._rtl(ws)
        self._title(ws, "تطبیق جمع/مقایسه/درصد اعلامی با محاسبه مستقل Python", 12)
        row = 3
        if intelligence is None:
            self._value_cell(ws, row, 1, "اطلاعات تطبیق در دسترس نیست.")
            self._print_setup(ws)
            return
        audits = intelligence.all_audits
        mismatches = sum(1 for item in audits if item.status.value == "mismatch")
        unverifiable = sum(1 for item in audits if item.status.value == "not_verifiable")
        verified = sum(1 for item in audits if item.status.value == "verified")
        self._section(ws, row, f"خلاصه: {verified} تأیید، {mismatches} مغایرت، {unverifiable} قابل بازتولید نبود", 12)
        row += 1
        self._header(ws, row, [
            "نوع", "برگه", "آدرس", "شاخص", "مقدار ثبت‌شده", "مقدار محاسبه‌شده",
            "اختلاف", "وضعیت", "محدوده داده", "فرمول", "محدوده فرمول", "علت احتمالی",
        ])
        header_row = row
        row += 1
        for item in audits:
            values = [
                item.kind, item.sheet, item.address, item.metric, item.declared,
                item.computed, item.difference, AUDIT_STATUS_FA.get(item.status.value, item.status.value),
                item.data_range, item.formula or "—", item.formula_range or "—", item.cause or "—",
            ]
            for offset, value in enumerate(values, start=1):
                self._value_cell(ws, row, offset, value)
            row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 12)
        row += 2
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
        ws.cell(
            row, 1,
            "هیچ اختلافی خودکار اصلاح نشد. جمع‌های بین‌برگه‌ای فقط با برچسب دوره یکسان مقایسه شده‌اند؛ "
            "اگر برگه ارجاع‌شده ستون سال دیگری داشته باشد، نتیجه «قابل بازتولید نبود» یا «مغایرت» ثبت می‌شود.",
        ).alignment = Alignment(horizontal="right", wrap_text=True)
        self._widths(ws, {
            "A": 14, "B": 18, "C": 12, "D": 42, "E": 16, "F": 16,
            "G": 14, "H": 18, "I": 34, "J": 34, "K": 16, "L": 60,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _metadata_warnings_sheet(self, wb, analysis: AnalysisResult) -> None:
        intelligence = analysis.intelligence
        ws = wb.create_sheet("هشدارهای متادیتا")
        self._rtl(ws)
        self._title(ws, "تعارض نام برگه، عنوان اصلی و زیرعنوان‌ها", 8)
        row = 3
        self._header(ws, row, ["برگه", "نوع تعارض", "پیام", "مقادیر", "شدت", "اطمینان"])
        header_row = row
        row += 1
        conflicts = intelligence.metadata_conflicts if intelligence is not None else []
        for item in conflicts:
            values = [
                item.sheet, item.kind, item.message,
                "؛ ".join(f"{key}={value}" for key, value in item.values.items()),
                item.severity, f"{item.confidence:.0%}",
            ]
            for offset, value in enumerate(values, start=1):
                self._value_cell(ws, row, offset, value)
            row += 1
        self._body(ws, header_row + 1, max(header_row + 1, row - 1), 1, 6)
        if not conflicts:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.cell(row, 1, "تعارض صریحی بین نام برگه، عنوان و زیرعنوان پیدا نشد.").alignment = Alignment(horizontal="right", wrap_text=True)
            row += 1
        row += 2
        self._section(ws, row, "زوج سال/دوره هر برگه", 8)
        row += 1
        self._header(ws, row, ["برگه", "سال نام برگه", "سال عنوان", "سال زیرعنوان", "جمع‌بندی"])
        row += 1
        if intelligence is not None:
            for sheet in intelligence.sheets:
                values = [
                    sheet.sheet,
                    "، ".join(sheet.sheet_years) or "—",
                    "، ".join(sheet.title_years) or "—",
                    "، ".join(sheet.subheader_years) or "—",
                    "منبع ترجیحی انتخاب نشد" if sheet.metadata_conflicts else "سازگار",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
        row += 2
        self._section(ws, row, "هشدارهای سطح Workbook", 8)
        row += 1
        for warning in (intelligence.warnings if intelligence is not None else analysis.warnings):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.cell(row, 1, "• " + warning).alignment = Alignment(horizontal="right", wrap_text=True)
            row += 1
        self._widths(ws, {"A": 22, "B": 18, "C": 70, "D": 60, "E": 12, "F": 10})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _data_dictionary(self, wb, analysis: AnalysisResult) -> None:
        generic = analysis.generic
        ws = wb.create_sheet("شناسنامه داده")
        self._rtl(ws)
        self._title(ws, "شناسنامه داده، پروفایل ستون‌ها و کیفیت داده", 12)
        quality = generic.quality if generic else None
        row = 3
        self._header(ws, row, ["موضوع", "مقدار"])
        row += 1
        info = [
            ("نام فایل", Path(analysis.meta.path).name),
            ("هش فایل", analysis.meta.sha256),
            ("برگه تحلیل‌شده", generic.table_sheet if generic else "—"),
            ("جدول تحلیل‌شده", generic.table_address if generic else "—"),
            ("نوع جدول", KIND_FA.get(analysis.primary_table.kind, "—") if analysis.primary_table else "—"),
            ("تعداد ردیف داده", generic.row_count if generic else 0),
            ("تعداد ستون", generic.column_count if generic else 0),
            ("سلول‌های خالی", quality.missing_cells if quality else 0),
            ("درصد خالی", f"{quality.missing_ratio:.1%}" if quality else "0.0%"),
            ("ردیف کاملاً تکراری", quality.duplicate_rows if quality else 0),
            ("ردیف خالی", quality.empty_rows if quality else 0),
            ("ستون شناسه", "، ".join(quality.identifier_columns) if quality and quality.identifier_columns else "—"),
            ("ستون کلید زمانی", "، ".join(quality.temporal_columns) if quality and quality.temporal_columns else "—"),
            ("ستون ثابت", "، ".join(quality.constant_columns) if quality and quality.constant_columns else "—"),
            ("ستون با نوع مختلط", "، ".join(quality.mixed_type_columns) if quality and quality.mixed_type_columns else "—"),
            ("درخواست کاربر", generic.goal if generic and generic.goal else "—"),
        ]
        for label, value in info:
            self._value_cell(ws, row, 1, label)
            self._value_cell(ws, row, 2, value)
            row += 1
        self._body(ws, 4, row - 1, 1, 2)

        row += 1
        self._section(ws, row, "پروفایل ستون‌ها (نوع، نقش، پوشش و نمونه)", 12)
        row += 1
        headers = ["ستون", "حرف", "نوع داده", "نقش", "تعداد معتبر", "خالی", "یکتا", "پوشش", "شناسه؟", "نمونه مقادیر", "هشدار"]
        self._header(ws, row, headers)
        header_row = row
        row += 1
        for column in (generic.columns if generic else []):
            self._value_cell(ws, row, 1, column.header)
            self._value_cell(ws, row, 2, column.letter)
            self._value_cell(ws, row, 3, TYPE_FA.get(column.kind.value, column.kind.value))
            self._value_cell(ws, row, 4, ROLE_FA.get(column.role, column.role.value))
            self._value_cell(ws, row, 5, column.count)
            self._value_cell(ws, row, 6, column.missing)
            self._value_cell(ws, row, 7, column.distinct)
            self._value_cell(ws, row, 8, f"{column.coverage:.0%}")
            self._value_cell(ws, row, 9, "بله" if column.is_identifier else "خیر")
            self._value_cell(ws, row, 10, "، ".join(value for value, _ in column.top_values[:3]))
            self._value_cell(ws, row, 11, "؛ ".join(column.warnings))
            row += 1
        self._body(ws, header_row + 1, row - 1, 1, 11)

        if quality:
            row += 1
            self._section(ws, row, "یادداشت‌های کیفیت داده", 12)
            for note in [*quality.notes, *(generic.warnings if generic else [])]:
                row += 1
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
                ws.cell(row, 1, "• " + note).alignment = Alignment(horizontal="right", wrap_text=True)

        self._widths(ws, {"A": 30, "B": 8, "C": 12, "D": 18, "E": 12, "F": 10, "G": 10, "H": 10, "I": 10, "J": 42, "K": 40})
        ws.freeze_panes = "A6"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _descriptive_stats(self, wb, analysis: AnalysisResult) -> None:
        generic = analysis.generic
        ws = wb.create_sheet("آمار توصیفی")
        self._rtl(ws)
        self._title(ws, "آمار توصیفی نوع‌محور (همه مقادیر معتبر، بدون حذف پرت)", 21)
        row = 3
        self._header(ws, row, [
            "ستون", "نوع", "تعداد", "خالی", "مجموع", "میانگین", "میانه", "مد",
            "کمینه", "بیشینه", "دامنه", "واریانس نمونه", "انحراف معیار نمونه",
            "واریانس جامعه", "انحراف معیار جامعه", "چارک اول", "چارک سوم", "IQR", "P10", "P90", "CV%",
        ])
        header_row = row
        row += 1
        numeric = [
            column
            for column in (generic.columns if generic else [])
            if column.is_numeric and not column.is_identifier and not column.is_temporal_key
        ]
        for column in numeric:
            values = [
                column.header, TYPE_FA.get(column.kind.value, column.kind.value), column.count, column.missing,
                column.total, column.mean, column.median, "، ".join(column.mode),
                column.minimum, column.maximum, column.value_range, column.variance_sample,
                column.std_sample, column.variance_population, column.std_population,
                column.q1, column.q3, column.iqr, column.p10, column.p90, column.cv_percent,
            ]
            for offset, value in enumerate(values, start=1):
                self._value_cell(ws, row, offset, value)
            row += 1
        self._body(ws, header_row + 1, row - 1, 1, 21)

        dates = [column for column in (generic.columns if generic else []) if column.kind.value == "date"]
        if dates:
            row += 2
            self._section(ws, row, "بازه تاریخ‌ها", 6)
            row += 1
            self._header(ws, row, ["ستون", "کمینه", "بیشینه", "تعداد", "بازه (روز)", "توضیح"])
            row += 1
            for column in dates:
                self._value_cell(ws, row, 1, column.header)
                self._value_cell(ws, row, 2, column.date_min)
                self._value_cell(ws, row, 3, column.date_max)
                self._value_cell(ws, row, 4, column.count)
                self._value_cell(ws, row, 5, column.range_days)
                self._value_cell(ws, row, 6, "تاریخ‌ها به شمسی نمایش داده شده‌اند.")
                row += 1

        row += 2
        self._section(ws, row, "راهنمای خواندن", 8)
        for note in (
            "میانگین برای داده‌های عددی مناسب است؛ در توزیع نامتقارن میانه تفسیر بهتری می‌دهد.",
            "واریانس نمونه (n-1) و جامعه (n) جداگانه گزارش شده‌اند تا روش مشخص باشد.",
            "مد فقط وقتی نمایش داده می‌شود که یک مقدار بیش از یک بار تکرار شده باشد.",
            "هیچ مقدار پرتی حذف نشده است؛ فهرست آن‌ها در برگه «داده‌های پرت» است.",
        ):
            row += 1
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            ws.cell(row, 1, "• " + note).alignment = Alignment(horizontal="right", wrap_text=True)

        self._widths(ws, {"A": 30, **{get_column_letter(i): 13 for i in range(2, 22)}})
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _frequency_sheet(self, wb, analysis: AnalysisResult) -> None:
        generic = analysis.generic
        ws = wb.create_sheet("توزیع فراوانی")
        self._rtl(ws)
        self._title(ws, "توزیع فراوانی مقادیر و بازه‌های عددی", 7)
        row = 3
        for table in (generic.frequencies if generic else [])[:10]:
            self._section(ws, row, f"ستون «{table.header}» — {table.total} مقدار معتبر، {table.distinct} مقدار یکتا", 7)
            row += 1
            if table.rows:
                self._header(ws, row, ["مقدار", "فراوانی", "درصد", "درصد تجمعی", "", "", ""])
                row += 1
                start = row
                for item in table.rows:
                    self._value_cell(ws, row, 1, item.value)
                    self._value_cell(ws, row, 2, item.count)
                    self._value_cell(ws, row, 3, item.percent)
                    self._value_cell(ws, row, 4, item.cumulative_percent)
                    row += 1
                self._body(ws, start, row - 1, 1, 4)
            if table.bins:
                self._header(ws, row, ["از", "تا", "فراوانی", "درصد", "", "", ""])
                row += 1
                start = row
                for item in table.bins:
                    self._value_cell(ws, row, 1, item.lower)
                    self._value_cell(ws, row, 2, item.upper)
                    self._value_cell(ws, row, 3, item.count)
                    self._value_cell(ws, row, 4, item.percent)
                    row += 1
                self._body(ws, start, row - 1, 1, 4)
            if table.note:
                row += 1
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
                ws.cell(row, 1, "یادداشت: " + table.note).alignment = Alignment(horizontal="right", wrap_text=True)
            row += 2
        self._widths(ws, {"A": 38, "B": 20, "C": 14, "D": 18, "E": 4, "F": 4, "G": 4})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _groups_sheet(self, wb, analysis: AnalysisResult) -> None:
        generic = analysis.generic
        ws = wb.create_sheet("تحلیل گروه‌ها")
        self._rtl(ws)
        self._title(ws, "مقایسه گروه‌ها و جدول متقاطع", 12)
        row = 3
        for comparison in (generic.groups if generic else []):
            self._section(ws, row, f"مقایسه «{comparison.measure_header}» بر اساس «{comparison.group_header}»", 8)
            row += 1
            self._header(ws, row, ["گروه", "تعداد", "مجموع", "میانگین", "میانه", "کمینه", "بیشینه", "انحراف معیار نمونه"])
            header_row = row
            row += 1
            for group in comparison.groups:
                self._value_cell(ws, row, 1, group.label)
                self._value_cell(ws, row, 2, group.count)
                self._value_cell(ws, row, 3, group.total)
                self._value_cell(ws, row, 4, group.mean)
                self._value_cell(ws, row, 5, group.median)
                self._value_cell(ws, row, 6, group.minimum)
                self._value_cell(ws, row, 7, group.maximum)
                self._value_cell(ws, row, 8, group.std_sample)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 8)
            row += 1
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
            ws.cell(row, 1, comparison.note).alignment = Alignment(horizontal="right", wrap_text=True)
            row += 2

        for table in (generic.crosstabs if generic else []):
            self._section(ws, row, f"جدول متقاطع «{table.row_header}» × «{table.column_header}»", 12)
            row += 1
            self._header(ws, row, ["ردیف", *table.columns, "جمع ردیف"])
            header_row = row
            row += 1
            for item in table.rows:
                self._value_cell(ws, row, 1, item.label)
                for offset, column in enumerate(table.columns, start=2):
                    self._value_cell(ws, row, offset, item.counts.get(column, 0))
                self._value_cell(ws, row, len(table.columns) + 2, item.total)
                row += 1
            self._value_cell(ws, row, 1, "جمع ستون")
            for offset, column in enumerate(table.columns, start=2):
                self._value_cell(ws, row, offset, table.column_totals.get(column, 0))
            self._value_cell(ws, row, len(table.columns) + 2, table.total)
            self._body(ws, header_row + 1, row, 1, len(table.columns) + 2)
            row += 1
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
            ws.cell(row, 1, table.note).alignment = Alignment(horizontal="right", wrap_text=True)
            row += 2

        self._widths(ws, {"A": 32, **{get_column_letter(i): 14 for i in range(2, 14)}})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _trend_correlation_sheet(self, wb, analysis: AnalysisResult) -> None:
        generic = analysis.generic
        ws = wb.create_sheet("روند و همبستگی")
        self._rtl(ws)
        self._title(ws, "تحلیل روند زمانی و همبستگی (توصیفی، نه علت‌ومعلولی)", 10)
        row = 3
        for trend in (generic.trends if generic else []):
            calendar_fa = {"jalali": "شمسی", "observation": "ترتیب مشاهدات"}.get(
                trend.calendar, "میلادی"
            )
            self._section(
                ws, row,
                f"روند «{trend.measure_header}» بر اساس «{trend.date_header}» (تقویم {calendar_fa})",
                8,
            )
            row += 1
            self._header(ws, row, ["دوره", "تعداد", "مجموع", "میانگین", "تغییر مطلق", "تغییر درصد", "جهت", "توضیح"])
            header_row = row
            row += 1
            for point in trend.points:
                self._value_cell(ws, row, 1, point.period_label)
                self._value_cell(ws, row, 2, point.count)
                self._value_cell(ws, row, 3, point.total)
                self._value_cell(ws, row, 4, point.mean)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 8)
            self._value_cell(ws, row, 1, "جمع‌بندی")
            self._value_cell(ws, row, 5, trend.change_absolute)
            self._value_cell(ws, row, 6, trend.change_percent)
            self._value_cell(ws, row, 7, trend.direction)
            self._value_cell(ws, row, 8, trend.note)
            row += 3

        correlations = generic.correlations if generic else []
        if correlations:
            self._section(ws, row, "همبستگی جفتی (پیرسون و اسپیرمن)", 10)
            row += 1
            self._header(ws, row, ["ستون اول", "ستون دوم", "تعداد جفت", "پیرسون", "اسپیرمن", "قدرت", "جهت", "تفسیر"])
            header_row = row
            row += 1
            for pair in correlations:
                self._value_cell(ws, row, 1, pair.header_a)
                self._value_cell(ws, row, 2, pair.header_b)
                self._value_cell(ws, row, 3, pair.count)
                for col, value in ((4, pair.pearson), (5, pair.spearman)):
                    cell = ws.cell(row, col, to_excel_number(value))
                    cell.number_format = "0.000000"
                self._value_cell(ws, row, 6, pair.strength)
                self._value_cell(ws, row, 7, pair.direction)
                self._value_cell(ws, row, 8, pair.note)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 8)

        self._widths(ws, {"A": 32, "B": 14, "C": 16, "D": 16, "E": 14, "F": 14, "G": 14, "H": 50})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _outliers_sheet(self, wb, analysis: AnalysisResult) -> None:
        generic = analysis.generic
        ws = wb.create_sheet("داده‌های پرت")
        self._rtl(ws)
        self._title(ws, "داده‌های پرت علامت‌گذاری‌شده (هیچ مقدار حذف نشده است)", 9)
        row = 3
        self._header(ws, row, ["ستون", "آدرس سلول", "مقدار خام", "مقدار استاندارد", "روش", "حد پایین", "حد بالا", "امتیاز", "اقدام پیشنهادی"])
        header_row = row
        row += 1
        for item in (generic.outliers if generic else []):
            self._value_cell(ws, row, 1, item.header)
            self._value_cell(ws, row, 2, item.address)
            self._value_cell(ws, row, 3, str(item.raw))
            self._value_cell(ws, row, 4, item.value)
            self._value_cell(ws, row, 5, item.method)
            self._value_cell(ws, row, 6, item.lower_fence)
            self._value_cell(ws, row, 7, item.upper_fence)
            self._value_cell(ws, row, 8, item.score)
            self._value_cell(ws, row, 9, "بررسی دستی مقدار اصلی؛ حذف خودکار انجام نشود.")
            row += 1
        self._body(ws, header_row + 1, row - 1, 1, 9)
        row += 2
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
        ws.cell(row, 1, "پرت‌ها با فنس‌های IQR (Q1−1.5×IQR تا Q3+1.5×IQR) شناسایی شده‌اند. پرت بودن دلیل خطا نیست؛ ممکن است مقدار واقعی و مهم باشد.").alignment = Alignment(horizontal="right", wrap_text=True)
        self._widths(ws, {"A": 28, "B": 14, "C": 22, "D": 20, "E": 14, "F": 18, "G": 18, "H": 12, "I": 44})
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _methodology_sheet(self, wb, analysis: AnalysisResult) -> None:
        generic = analysis.generic
        ws = wb.create_sheet("روش‌شناسی آماری")
        self._rtl(ws)
        self._title(ws, "روش‌شناسی، محدودیت‌ها و انتخاب تحلیل", 6)
        row = 3
        self._section(ws, row, "روش‌های به‌کاررفته", 6)
        row += 1
        for line in (generic.methodology if generic else []):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            ws.cell(row, 1, "• " + line).alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
            row += 1

        row += 1
        self._section(ws, row, "انتخاب خودکار تحلیل", 6)
        row += 1
        for name, reason in [
            ("آمار توصیفی", "ستون‌های عددی و تاریخ معتبر"),
            ("توزیع فراوانی", "ستون متنی با دسته‌های محدود یا ستون عددی با تنوع بالا"),
            ("مقایسه گروه‌ها", "ستون متنی با ۲ تا ۲۰ گروه و ستون عددی با حداقل ۳ مشاهده"),
            ("جدول متقاطع", "دو ستون متنی با دسته‌های محدود و حداقل ۴ ردیف مشترک"),
            ("روند زمانی", "ستون تاریخ معتبر یا کلید زمانی و ستون عددی با حداقل دو مشاهده/دوره"),
            ("همبستگی", "حداقل دو ستون عددی با حداقل ۳ جفت مقدار کامل"),
            ("داده پرت", "ستون عددی با حداقل ۸ مقدار و IQR غیرصفر"),
        ]:
            self._value_cell(ws, row, 1, name)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
            self._value_cell(ws, row, 2, reason)
            row += 1

        row += 1
        self._section(ws, row, "محدودیت‌ها و هشدارها", 6)
        row += 1
        limitations = [
            "تحلیل‌های این گزارش اکتشافی و توصیفی هستند؛ برای نتیجه‌گیری استنباطی به طرح نمونه‌گیری و آزمون آماری مناسب نیاز است.",
            "همبستگی به‌معنای علت‌ومعلولی نیست.",
            "داده‌های گمشده و پرت حذف نشده‌اند؛ تعداد و مکان آن‌ها ثبت شده است.",
            "نتیجه مدل زبانی محاسبه مرجع نیست؛ همه اعداد از Python آمده‌اند.",
        ]
        if generic:
            limitations.extend(generic.warnings)
        for line in limitations:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            ws.cell(row, 1, "• " + line).alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
            row += 1

        self._widths(ws, {"A": 28, "B": 30, "C": 20, "D": 20, "E": 20, "F": 20})
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _semantics_sheet(self, wb, analysis: AnalysisResult) -> None:
        """Audit-only view of every field's explicit semantic role and evidence."""
        from local_office_ai.excel.semantic_classification import SEMANTIC_ROLE_FA

        semantics = getattr(analysis, "semantics", None)
        ws = wb.create_sheet("نقش‌های معنایی")
        self._rtl(ws)
        self._title(ws, "نقش معنایی هر فیلد، شاهد تصمیم و نقش جمع‌های میانی/کل", 9)
        row = 3
        if semantics is None:
            self._value_cell(ws, row, 1, "طبقه‌بندی معنایی برای این فایل ساخته نشد.")
            self._print_setup(ws)
            return
        counts = semantics.role_counts()
        summary = "، ".join(
            f"{SEMANTIC_ROLE_FA.get(role, role)}: {count}" for role, count in sorted(counts.items())
        )
        self._value_cell(ws, row, 1, "شمارش نقش‌ها")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(ws, row, 2, summary or "—")
        row += 2

        for block in semantics.blocks:
            self._section(
                ws,
                row,
                f"برگه «{block.sheet}» — بلوک {block.block_id} — {block.address} — {len(block.columns)} ستون",
                9,
            )
            row += 1
            self._header(ws, row, [
                "ستون", "عنوان", "نقش معنایی", "اطمینان", "نقش تجاری", "نقش آماری",
                "واحد", "شاهد تصمیم", "یادداشت",
            ])
            header_row = row
            row += 1
            for column in block.columns:
                values = [
                    column.letter,
                    column.header or "—",
                    SEMANTIC_ROLE_FA.get(column.role, column.role.value),
                    f"{column.confidence:.0%}",
                    column.business_role,
                    column.statistical_role,
                    column.unit or "—",
                    "؛ ".join(column.evidence) or "—",
                    "؛ ".join(column.notes) or "—",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
            if row > header_row + 1:
                self._body(ws, header_row + 1, row - 1, 1, 9)
            for item in block.rows:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
                self._value_cell(
                    ws,
                    row,
                    1,
                    f"• ردیف {item.row} نقش «{SEMANTIC_ROLE_FA.get(item.role, item.role.value)}» دارد"
                    + (f" — {item.label}" if item.label else ""),
                )
                row += 1
            row += 1
        self._widths(ws, {
            "A": 8, "B": 30, "C": 18, "D": 10, "E": 14, "F": 18,
            "G": 12, "H": 62, "I": 40,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _quality_details_sheet(self, wb, analysis: AnalysisResult) -> None:
        """Audit-only view of the PASS/WARNING/FAIL quality findings."""
        from local_office_ai.excel.data_quality import QUALITY_STATUS_FA

        report = getattr(analysis, "quality", None)
        ws = wb.create_sheet("کیفیت داده (تفصیلی)")
        self._rtl(ws)
        self._title(ws, "یافته‌های کیفیت داده با وضعیت PASS/WARNING/FAIL و محل دقیق", 10)
        row = 3
        if report is None:
            self._value_cell(ws, row, 1, "یافته کیفیت داده برای این فایل ثبت نشد.")
            self._print_setup(ws)
            return
        counts = report.counts()
        self._value_cell(ws, row, 1, "وضعیت کلی")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=10)
        self._value_cell(
            ws,
            row,
            2,
            f"{QUALITY_STATUS_FA.get(report.status, report.status.value)} — "
            f"PASS {counts.get('PASS', 0)}، WARNING {counts.get('WARNING', 0)}، FAIL {counts.get('FAIL', 0)}",
        )
        row += 2
        self._header(ws, row, [
            "وضعیت", "بررسی", "پیام", "دامنه", "برگه", "بلوک", "سلول", "ستون",
            "تعداد", "نمونه/شاهد",
        ])
        header_row = row
        row += 1
        colors = {"PASS": GREEN, "WARNING": YELLOW, "FAIL": RED}
        for finding in report.findings:
            samples = list(finding.samples)
            evidence = list(finding.evidence)
            combined = "؛ ".join([*samples, *evidence])
            values = [
                QUALITY_STATUS_FA.get(finding.status, finding.status.value),
                finding.title,
                finding.message,
                finding.scope or "—",
                finding.sheet or "—",
                finding.block or "—",
                finding.address or "—",
                finding.column or "—",
                finding.count,
                combined or "—",
            ]
            for offset, value in enumerate(values, start=1):
                self._value_cell(ws, row, offset, value)
            self._fill(ws.cell(row, 1), colors.get(finding.status.value, GREY))
            row += 1
        self._body(ws, header_row + 1, row - 1, 1, 10)
        ws.auto_filter.ref = f"A{header_row}:J{max(row - 1, header_row + 1)}"
        self._widths(ws, {
            "A": 12, "B": 30, "C": 56, "D": 30, "E": 16, "F": 12,
            "G": 10, "H": 8, "I": 8, "J": 48,
        })
        ws.freeze_panes = f"A{header_row + 1}"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _statistical_reconciliation_sheet(self, wb, analysis: AnalysisResult) -> None:
        """Audit-only view of every reconciliation and comparison result."""
        from local_office_ai.excel.statistical_reconciliation import STATUS_FA

        report = getattr(analysis, "statistical_reconciliation", None)
        ws = wb.create_sheet("تطبیق آماری و صحت محاسبات")
        self._rtl(ws)
        self._title(ws, "تطبیق آماری: مقدار مورد انتظار، مقدار محاسبه‌شده، اختلاف و تحمل", 11)
        row = 3
        if report is None:
            self._value_cell(ws, row, 1, "تطبیق آماری برای این فایل ساخته نشد.")
            self._print_setup(ws)
            return

        counts = report.counts()
        self._value_cell(ws, row, 1, "وضعیت کلی")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=11)
        self._value_cell(
            ws,
            row,
            2,
            f"{STATUS_FA.get(report.status, report.status.value)} — "
            f"PASS {counts.get('PASS', 0)}، WARNING {counts.get('WARNING', 0)}، "
            f"FAIL {counts.get('FAIL', 0)}، هشدارهای موتور {len(report.warnings)}",
        )
        row += 1
        self._value_cell(ws, row, 1, "اصل محاسبه")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=11)
        self._value_cell(
            ws,
            row,
            2,
            "همه اعداد با Decimal در Python محاسبه شده‌اند؛ مدل زبانی مجاز به محاسبه مجدد نیست.",
        )
        row += 2

        if report.warnings:
            self._section(ws, row, "هشدارهای موتور تطبیق", 11)
            row += 1
            for warning in report.warnings:
                self._value_cell(ws, row, 1, "هشدار")
                ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=11)
                self._value_cell(ws, row, 2, warning)
                self._fill(ws.cell(row, 1), YELLOW)
                row += 1
            row += 1

        colors = {"PASS": GREEN, "WARNING": YELLOW, "FAIL": RED}
        self._section(ws, row, "تطبیق جمع‌ها، اجزا، جمع میانی، گروه و دوره", 11)
        row += 1
        self._header(ws, row, [
            "وضعیت", "نوع", "شاخص/برچسب", "مقدار مورد انتظار", "مقدار محاسبه‌شده",
            "اختلاف", "تحمل", "توضیح", "برگه", "بلوک", "منابع",
        ])
        header_row = row
        row += 1
        for item in report.reconciliations:
            values = [
                STATUS_FA.get(item.status, item.status.value),
                item.kind_label,
                item.label or item.measure or "—",
                item.expected,
                item.calculated,
                item.difference,
                item.tolerance.describe_fa(item.expected),
                item.explanation,
                item.sheet or "—",
                item.block or "—",
                "؛ ".join(item.sources[:6]) or "—",
            ]
            for offset, value in enumerate(values, start=1):
                money = offset in {4, 5, 6} and isinstance(value, Decimal)
                self._value_cell(ws, row, offset, value, money=money)
            self._fill(ws.cell(row, 1), colors.get(item.status.value, GREY))
            row += 1
        if row == header_row + 1:
            self._value_cell(ws, row, 1, "رابطه جمع قابل توجیهی پیدا نشد.")
            row += 1
        else:
            self._body(ws, header_row + 1, row - 1, 1, 11)

        if report.comparisons:
            row += 1
            self._section(ws, row, "مقایسه بین‌دوره‌ای (تغییر مطلق و درصدی)", 11)
            row += 1
            self._header(ws, row, [
                "وضعیت", "شاخص", "دوره مبنا", "دوره فعلی", "مقدار مبنا", "مقدار فعلی",
                "تغییر مطلق", "تغییر درصد", "جهت", "توضیح", "منبع",
            ])
            comparison_header = row
            row += 1
            for item in report.comparisons:
                change = item.change
                values = [
                    STATUS_FA.get(item.status, item.status.value),
                    item.metric or "—",
                    item.previous_label,
                    item.current_label,
                    change.previous,
                    change.current,
                    change.absolute,
                    change.percent,
                    change.direction,
                    item.explanation,
                    item.source or "—",
                ]
                for offset, value in enumerate(values, start=1):
                    money = offset in {5, 6, 7} and isinstance(value, Decimal)
                    self._value_cell(ws, row, offset, value, money=money)
                self._fill(ws.cell(row, 1), colors.get(item.status.value, GREY))
                row += 1
            self._body(ws, comparison_header + 1, row - 1, 1, 11)

        if report.averages:
            row += 1
            self._section(ws, row, "میانگین ساده در برابر میانگین وزنی", 11)
            row += 1
            self._header(ws, row, [
                "وضعیت", "شاخص", "مقدار گزارش‌شده", "میانگین ساده", "میانگین وزنی",
                "وزن لازم", "وزن موجود", "تعداد مشاهده", "توضیح", "واحد", "—",
            ])
            average_header = row
            row += 1
            for item in report.averages:
                values = [
                    STATUS_FA.get(item.status, item.status.value),
                    item.label or "—",
                    item.value,
                    item.unweighted,
                    item.weighted,
                    "بله" if item.weights_required else "خیر",
                    "بله" if item.weights_available else "خیر",
                    item.observations,
                    item.explanation,
                    item.unit or "—",
                    "—",
                ]
                for offset, value in enumerate(values, start=1):
                    money = offset in {3, 4, 5} and isinstance(value, Decimal)
                    self._value_cell(ws, row, offset, value, money=money)
                self._fill(ws.cell(row, 1), colors.get(item.status.value, GREY))
                row += 1
            self._body(ws, average_header + 1, row - 1, 1, 11)

        self._widths(ws, {
            "A": 12, "B": 18, "C": 30, "D": 18, "E": 18, "F": 14,
            "G": 26, "H": 64, "I": 18, "J": 14, "K": 44,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)

    # ------------------------------------------------------------------ #
    def _extraction_model(self, wb, analysis: AnalysisResult) -> None:
        """Audit-only view of the structured extraction model and its evidence."""
        model = getattr(analysis, "extraction", None)
        ws = wb.create_sheet("مدل استخراج")
        self._rtl(ws)
        self._title(ws, "مدل استخراج: مرز بلوک‌ها، سرستون‌ها، ادغام‌ها، سلول‌ها و شواهد دوره", 12)
        row = 3
        if model is None:
            self._value_cell(ws, row, 1, "مدل استخراج برای این فایل ساخته نشد.")
            self._print_setup(ws)
            return
        status_fa = {
            "detected": "شناسایی شد",
            "not_detected": "شناسایی نشد",
            "conflicting": "شواهد متناقض",
        }
        summary = [
            ("نام فایل", model.source_filename),
            ("وضعیت دوره", status_fa.get(model.period_status, model.period_status)),
            ("اطمینان دوره", f"{model.period_confidence:.0%}"),
            ("دوره‌های شناسایی‌شده", "، ".join(item["token"] for item in model.periods) or "—"),
            ("برگه‌ها / بلوک‌ها", f"{len(model.sheets)} / {sum(len(sheet.blocks) for sheet in model.sheets)}"),
        ]
        for label, value in summary:
            self._value_cell(ws, row, 1, label)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=12)
            self._value_cell(ws, row, 2, value)
            row += 1

        for sheet in model.sheets:
            row += 1
            self._section(
                ws,
                row,
                f"برگه «{sheet.name}» — {len(sheet.blocks)} بلوک — دوره: "
                f"{'، '.join(item['token'] for item in sheet.periods) or 'شناسایی نشد'}",
                12,
            )
            row += 1
            self._header(ws, row, [
                "شناسه بلوک", "نوع", "محدوده", "ردیف‌های عنوان", "ستون‌ها",
                "دوره‌های بلوک", "منابع دوره", "اطمینان دوره", "شاهد دوره",
                "ادغام‌ها", "سلول‌های ثبت‌شده", "اطمینان ساختار",
            ])
            header_row = row
            row += 1
            if not sheet.blocks:
                self._value_cell(ws, row, 1, "بلوکی شناسایی نشد.")
                row += 1
            for block in sheet.blocks:
                evidence = "؛ ".join(
                    f"{item['token']} ← {item['source_label']}"
                    + (f" ({item['address']})" if item.get("address") else "")
                    for item in block.period_evidence
                )
                merged = "، ".join(item.address for item in block.merged_ranges)
                values = [
                    block.block_id,
                    block.kind,
                    block.boundaries.get("address", ""),
                    "، ".join(str(item) for item in block.header_rows),
                    "، ".join(column.letter for column in block.columns),
                    "، ".join(item["token"] for item in block.periods) or "—",
                    "، ".join(
                        sorted({source for item in block.periods for source in item["sources"]})
                    ) or "—",
                    f"{block.period_confidence:.0%}" if block.periods else "—",
                    evidence or "—",
                    merged or "—",
                    f"{len(block.cells)}{' (محدودشده)' if block.cells_truncated else ''}",
                    f"{block.confidence:.0%}",
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 12)
            for warning in sheet.warnings:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
                ws.cell(row, 1, "• " + warning).alignment = Alignment(horizontal="right", wrap_text=True)
                row += 1

        if model.warnings:
            row += 1
            self._section(ws, row, "هشدارهای سطح Workbook", 12)
            row += 1
            for warning in model.warnings:
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
                ws.cell(row, 1, "• " + warning).alignment = Alignment(horizontal="right", wrap_text=True)
                row += 1

        self._widths(ws, {
            "A": 22, "B": 12, "C": 14, "D": 12, "E": 22, "F": 24,
            "G": 18, "H": 11, "I": 55, "J": 18, "K": 14, "L": 12,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)


    # ------------------------------------------------------------------ #
    def _semantic_matching_sheet(self, wb, analysis: AnalysisResult) -> None:
        """Audit-only evidence for every matched, probable or rejected field pair."""
        from local_office_ai.excel.semantic_matching import (
            PERIOD_STATUS_FA,
            RELATION_FA,
            STATUS_FA,
            UNIT_STATE_FA,
        )

        report = getattr(analysis, "semantic_matching", None)
        ws = wb.create_sheet("تطبیق معنایی منابع")
        self._rtl(ws)
        self._title(ws, "تطبیق معنایی منابع: برچسب، واحد، دوره و شاهد تصمیم", 10)
        row = 3
        if report is None:
            self._value_cell(ws, row, 1, "تطبیق معنایی برای این فایل ساخته نشد.")
            self._print_setup(ws)
            return

        counts = report.counts()
        self._value_cell(ws, row, 1, "خلاصه تصمیم‌ها")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(
            ws,
            row,
            2,
            f"فیلد {counts.get('fields', 0)} — تطبیق قطعی {counts.get('matched', 0)} — "
            f"احتمالی {counts.get('probable', 0)} — مبهم {counts.get('ambiguous', 0)} — "
            f"رد‌شده {counts.get('rejected', 0)}",
        )
        row += 1
        self._value_cell(ws, row, 1, "اصل کار")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        self._value_cell(
            ws,
            row,
            2,
            "تطبیق فقط با شاهد نام نرمال‌شده، نقش معنایی، واحد و دوره گزارش‌دهی انجام شده است؛ "
            "هیچ مقدار یا جمعی بین منابع تجمیع نشده است.",
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

        colors = {
            "matched": GREEN,
            "probable": YELLOW,
            "ambiguous": YELLOW,
            "rejected": RED,
        }

        def write_match_section(title: str, items: list[Any]) -> None:
            nonlocal row
            if not items:
                return
            self._section(ws, row, title, 10)
            row += 1
            self._header(ws, row, [
                "وضعیت", "نوع", "برچسب اصلی ← نرمال‌شده", "اعضا (منبع/برگه/بلوک/عنوان)",
                "رابطه دوره", "واحد", "قابل مقایسه", "اطمینان", "شاهد/دلیل", "توضیح دوره",
            ])
            header_row = row
            row += 1
            for item in items:
                members = "؛ ".join(
                    f"{member.source}/{member.sheet}/{member.block}/«{member.header}»"
                    for member in item.members
                )
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

        write_match_section("تطبیق قطعی فیلدهای هم‌معنا", report.matches)
        write_match_section("تطبیق احتمالی (برابری کامل اثبات نشد)", report.probable)

        if report.ambiguities:
            self._section(ws, row, "موارد مبهم: گزینه‌ها نگه داشته شده‌اند و تطبیق خودکار انجام نشد", 9)
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
                statuses = "؛ ".join(
                    STATUS_FA.get(candidate.status, candidate.status.value)
                    for candidate in item.candidates
                )
                reasons = "؛ ".join(candidate.reason for candidate in item.candidates if candidate.reason)
                evidence = "؛ ".join(
                    text for candidate in item.candidates for text in candidate.evidence
                )
                values = [
                    anchor.header,
                    f"{anchor.source}/{anchor.sheet}",
                    len(item.candidates),
                    candidates or "—",
                    statuses or "—",
                    f"{item.confidence:.0%}",
                    reasons or item.reason,
                    evidence or "—",
                    PERIOD_STATUS_FA.get(item.period_status, item.period_status or "—"),
                ]
                for offset, value in enumerate(values, start=1):
                    self._value_cell(ws, row, offset, value)
                self._fill(ws.cell(row, 1), YELLOW)
                row += 1
            self._body(ws, header_row + 1, row - 1, 1, 9)
            row += 1

        if report.rejections:
            self._section(ws, row, "موارد رد‌شده: چرا دو فیلد با نام نزدیک هم‌معنا نیستند", 9)
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
            "A": 22, "B": 20, "C": 26, "D": 46, "E": 16,
            "F": 14, "G": 16, "H": 11, "I": 60, "J": 50,
        })
        ws.freeze_panes = "A5"
        self._print_setup(ws)
