"""Result panel for a completed analysis.

Every figure shown here is read from the engine's ``RunResult`` object; nothing
is recomputed in the interface.  Tabs mirror what the delivered workbook
contains: request answers, findings, tables, a chart, and limits/warnings.
"""

from __future__ import annotations

import html
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    Badge,
    BarChart,
    CallbackButton,
    Card,
    EmptyState,
    IconButton,
    MetricTile,
    SectionHeader,
    horizontal_line,
    make_label,
)
from local_office_ai.ui.desktop.formatting import (
    align_table_header,
    display_value,
    elide,
    format_cell,
    format_number,
    format_percent,
    ltr,
    soften_numbers,
    mode_label,
    severity_label,
    severity_tone,
    validation_label,
    validation_tone,
)
from local_office_ai.ui.desktop.preview import ReportPreviewDialog, open_path
from local_office_ai.ui.desktop.theme import Palette


def rich_text(text: Any) -> str:
    return html.escape(str(text or "")).replace("\n", "<br/>")


class ResultsView(Card):
    followup_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent, margins=(20, 18, 20, 18), spacing=14)
        self._palette: Palette | None = None
        self._result: Any = None
        self._path: str = ""
        self._chart: BarChart | None = None
        self._chart_note: QLabel | None = None
        self._series_combo: QComboBox | None = None
        self._chart_series: list[Any] = []
        self._table_options: list[tuple[str, list[str], list[list[str]]]] = []
        self._table_combo: QComboBox | None = None
        self._table_widget: QTableWidget | None = None
        self._icon_buttons: list[IconButton] = []
        self._build()
        self.setVisible(False)

    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        header = QHBoxLayout()
        header.setSpacing(8)
        self.status_badge = Badge("تکمیل شد", "success")
        self.file_label = QLabel("")
        self.file_label.setObjectName("CardTitle")
        self.mode_badge = Badge("", "accent")
        header.addWidget(self.status_badge)
        header.addWidget(self.file_label, 1)
        header.addWidget(self.mode_badge)

        self.preview_button = self._icon_button("eye", "پیش‌نمایش گزارش")
        self.open_button = self._icon_button("external", "بازکردن فایل Excel")
        self.folder_button = self._icon_button("folder-open", "بازکردن پوشه خروجی")
        self.copy_button = self._icon_button("copy", "کپی خلاصه نتیجه")
        self.followup_button = CallbackButton("درخواست تکمیلی", self.followup_requested.emit, object_name="Ghost", icon_name="sparkles")
        header.addWidget(self.preview_button)
        header.addWidget(self.open_button)
        header.addWidget(self.folder_button)
        header.addWidget(self.copy_button)
        header.addWidget(self.followup_button)
        self.body.addLayout(header)

        self.path_label = QLabel("")
        self.path_label.setObjectName("Caption")
        self.path_label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.path_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.body.addWidget(self.path_label)
        self.body.addWidget(horizontal_line())

        self.tabs = QTabWidget()
        self.summary_tab = self._scroll_tab()
        self.findings_tab = self._scroll_tab()
        self.tables_tab = self._scroll_tab()
        self.chart_tab = self._scroll_tab()
        self.notes_tab = self._scroll_tab()
        self.tabs.addTab(self.summary_tab, "خلاصه")
        self.tabs.addTab(self.findings_tab, "یافته‌ها")
        self.tabs.addTab(self.tables_tab, "جداول")
        self.tabs.addTab(self.chart_tab, "نمودار")
        self.tabs.addTab(self.notes_tab, "هشدارها")
        self.tabs.setMinimumHeight(430)
        self.body.addWidget(self.tabs)

        self.preview_button.clicked.connect(self._preview)
        self.open_button.clicked.connect(self._open_excel)
        self.folder_button.clicked.connect(self._open_folder)
        self.copy_button.clicked.connect(self._copy)

    def _icon_button(self, name: str, tooltip: str) -> IconButton:
        button = IconButton(name, tooltip)
        self._icon_buttons.append(button)
        return button

    @staticmethod
    def _scroll_tab() -> QScrollArea:
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.Shape.NoFrame)
        content = QWidget()
        area.setWidget(content)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(4, 8, 4, 8)
        layout.setSpacing(12)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        area.setProperty("contentLayout", True)
        return area

    @staticmethod
    def _tab_layout(area: QScrollArea) -> QVBoxLayout:
        widget = area.widget()
        assert widget is not None
        return widget.layout()  # type: ignore[return-value]

    # ------------------------------------------------------------------ #
    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        for button in self._icon_buttons:
            button.apply_palette(palette)
        if self._chart is not None:
            self._chart.apply_palette(palette)

    # ------------------------------------------------------------------ #
    def set_result(self, result: Any, palette: Palette | None = None) -> None:
        self._result = result
        if palette is not None:
            self.apply_palette(palette)
        analysis = getattr(result, "analysis", None)
        if analysis is None:
            return
        self._path = str(getattr(result, "output_path", "") or "")
        self.file_label.setText(Path(self._path).name if self._path else "نتیجه تحلیل")
        self.path_label.setText(ltr(self._path))
        self.path_label.setToolTip(self._path)
        status = str(getattr(result, "status", ""))
        self.status_badge.set_status(
            "تکمیل شد" if status == "completed" else "با هشدار تکمیل شد",
            "success" if status == "completed" else "warning",
        )
        ai = getattr(analysis, "ai", {}) or {}
        self.mode_badge.set_status(
            f"{mode_label(ai.get('mode'))}" + (f" — {ltr(ai.get('model'))}" if ai.get("model") else ""),
            "accent" if ai.get("mode") == "model" else "muted",
        )
        self._populate_summary(analysis, ai)
        self._populate_findings(analysis, ai)
        self._populate_tables(analysis)
        self._populate_chart(analysis)
        self._populate_notes(analysis, ai)
        self.setVisible(True)

    # -- summary -------------------------------------------------------- #
    def _populate_summary(self, analysis: Any, ai: dict[str, Any]) -> None:
        layout = self._clear_tab(self.summary_tab)
        response = ai.get("response") or {}
        summary = response.get("summary") or ""
        layout.addWidget(SectionHeader("خلاصه مدیریتی", "متن مدیریتی از تحلیل مدل یا قواعد قطعی Python"))
        confidence = float(getattr(analysis, "structure_confidence", 0.0) or 0.0)
        context_row = QHBoxLayout()
        context_row.addWidget(Badge(f"اطمینان ساختار {confidence:.0%}", "info" if confidence >= 0.6 else "warning"))
        primary = getattr(analysis, "primary_table", None)
        if primary is not None:
            context_row.addWidget(Badge(f"جدول اصلی: {primary.sheet}!{primary.address}", "muted"))
        context_row.addStretch(1)
        layout.addLayout(context_row)
        summary_label = QLabel(
            rich_text(soften_numbers(summary)) if summary else "خلاصه‌ای از تحلیل در دسترس نیست."
        )
        summary_label.setWordWrap(True)
        summary_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(summary_label)
        error = str(ai.get("error") or "")
        if ai.get("mode") != "model" and error:
            if "به‌درخواست کاربر" in error or "به‌ درخواست" in error:
                message = (
                    "تحلیل مدل محلی در این اجرا انجام نشد (Ollama در دسترس نبود یا مدلی انتخاب نشده بود)؛ "
                    "گزارش با قواعد قطعی Python ساخته شد."
                )
            else:
                message = f"تحلیل مدل انجام نشد: {error}"
            note = QLabel(message)
            note.setObjectName("Caption")
            note.setWordWrap(True)
            note.setToolTip(error)
            layout.addWidget(note)

        generic = getattr(analysis, "generic", None)
        quality = getattr(generic, "quality", None) if generic is not None else None
        anomalies = [item for item in (getattr(analysis, "anomalies", []) or []) if item.severity.value != "info"]
        tiles = QHBoxLayout()
        tiles.setSpacing(10)
        if generic is not None:
            tiles.addWidget(MetricTile("ردیف‌های جدول اصلی", format_number(generic.row_count)))
            tiles.addWidget(MetricTile("ستون‌ها", format_number(generic.column_count)))
            if quality is not None:
                tiles.addWidget(
                    MetricTile(
                        "سلول‌های خالی",
                        format_percent(quality.missing_ratio),
                        f"{format_number(quality.missing_cells)} از {format_number(quality.total_cells)} سلول",
                    )
                )
        else:
            for calculation in list(getattr(analysis, "calculations", []) or [])[:3]:
                tiles.addWidget(MetricTile(calculation.label, format_number(calculation.value), calculation.unit))
        tiles.addWidget(MetricTile("موارد نیازمند بررسی", format_number(len(anomalies))))
        layout.addLayout(tiles)

        answers = list(getattr(analysis, "request_answers", []) or [])
        if answers:
            layout.addWidget(horizontal_line())
            layout.addWidget(SectionHeader("پاسخ درخواست شما", "برگرفته از برگه «پاسخ درخواست» خروجی"))
            for answer in answers[:4]:
                layout.addWidget(self._answer_block(answer))

        layout.addStretch(1)

    @staticmethod
    def _answer_block(answer: Any) -> QFrame:
        frame = QFrame()
        frame.setObjectName("CardAlt")
        column = QVBoxLayout(frame)
        column.setContentsMargins(14, 12, 14, 12)
        column.setSpacing(4)
        head = QHBoxLayout()
        title = QLabel(str(getattr(answer, "title", "")))
        title.setObjectName("CardTitle")
        head.addWidget(title)
        value = str(getattr(answer, "value", "") or "").strip()
        if value and value not in {"0", "0 ماه"}:
            shown = elide(soften_numbers(display_value(value)), 32)
            badge = Badge(shown, "accent")
            badge.setToolTip(soften_numbers(display_value(value)))
            head.addWidget(badge)
        head.addStretch(1)
        column.addLayout(head)
        text = QLabel(rich_text(soften_numbers(getattr(answer, "text", ""))))
        text.setWordWrap(True)
        text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        column.addWidget(text)
        source = str(getattr(answer, "source", "") or "")
        if source:
            caption = QLabel("منبع: " + source)
            caption.setObjectName("Caption")
            caption.setWordWrap(True)
            column.addWidget(caption)
        return frame

    # -- findings ------------------------------------------------------- #
    def _populate_findings(self, analysis: Any, ai: dict[str, Any]) -> None:
        layout = self._clear_tab(self.findings_tab)
        response = ai.get("response") or {}
        answers = list(getattr(analysis, "request_answers", []) or [])
        layout.addWidget(
            SectionHeader(
                "یافته‌های اصلی",
                "هر عدد از محاسبات Python و منبع آن ذکر شده است؛ کسرهای بسیار طولانی فقط در نمایش کوتاه شده‌اند.",
            )
        )
        if answers:
            for answer in answers:
                layout.addWidget(self._answer_block(answer))
        else:
            layout.addWidget(make_label("پاسخ محاسبه‌شده‌ای برای این درخواست پیدا نشد.", "Muted", wrap=True))

        findings = list(response.get("findings") or [])
        if findings:
            layout.addWidget(horizontal_line())
            layout.addWidget(SectionHeader("نکات تحلیل مدیریتی"))
            for item in findings:
                row = QLabel("• " + rich_text(soften_numbers(item)))
                row.setWordWrap(True)
                row.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                layout.addWidget(row)

        anomalies = list(getattr(analysis, "anomalies", []) or [])
        if anomalies:
            layout.addWidget(horizontal_line())
            layout.addWidget(SectionHeader("موارد نیازمند بررسی"))
            for anomaly in anomalies[:12]:
                row = QHBoxLayout()
                row.addWidget(Badge(severity_label(anomaly.severity), severity_tone(anomaly.severity)))
                text = QLabel(rich_text(soften_numbers(anomaly.message)))
                text.setWordWrap(True)
                row.addWidget(text, 1)
                holder = QWidget()
                holder.setLayout(row)
                layout.addWidget(holder)
                if anomaly.detail:
                    detail = QLabel(rich_text(soften_numbers(anomaly.detail)))
                    detail.setObjectName("Caption")
                    detail.setWordWrap(True)
                    layout.addWidget(detail)
        layout.addStretch(1)

    # -- tables --------------------------------------------------------- #
    def _populate_tables(self, analysis: Any) -> None:
        layout = self._clear_tab(self.tables_tab)
        options: list[tuple[str, list[str], list[list[str]]]] = []
        answers = list(getattr(analysis, "request_answers", []) or [])
        if answers:
            options.append(
                (
                    "پاسخ درخواست",
                    ["موضوع", "مقدار", "پاسخ"],
                    [
                        [
                            str(answer.title),
                            display_value(answer.value),
                            format_cell(answer.text, limit=220),
                        ]
                        for answer in answers
                    ],
                )
            )
        calculations = list(getattr(analysis, "calculations", []) or [])
        if calculations:
            options.append(
                (
                    "محاسبات",
                    ["شاخص", "مقدار", "واحد", "روش"],
                    [
                        [
                            item.label,
                            format_number(item.value),
                            str(item.unit or ""),
                            str(item.method or ""),
                        ]
                        for item in calculations
                    ],
                )
            )
        generic = getattr(analysis, "generic", None)
        if generic is not None:
            columns = list(getattr(generic, "columns", []) or [])
            if columns:
                options.append(
                    (
                        "آمار ستون‌ها",
                        ["ستون", "نوع", "تعداد", "خالی", "میانگین", "میانه", "کمینه", "بیشینه", "جمع"],
                        [
                            [
                                column.header,
                                getattr(column.kind, "value", str(column.kind)),
                                format_number(column.count),
                                format_number(column.missing),
                                format_number(column.mean),
                                format_number(column.median),
                                format_number(column.minimum),
                                format_number(column.maximum),
                                format_number(column.total),
                            ]
                            for column in columns
                        ],
                    )
                )
            frequencies = list(getattr(generic, "frequencies", []) or [])
            if frequencies:
                frequency = frequencies[0]
                options.append(
                    (
                        f"فراوانی — {frequency.header}",
                        ["مقدار", "تعداد", "درصد", "درصد تجمعی"],
                        [
                            [
                                row.value,
                                format_number(row.count),
                                format_percent(row.percent),
                                format_percent(row.cumulative_percent),
                            ]
                            for row in frequency.rows
                        ],
                    )
                )
            groups = list(getattr(generic, "groups", []) or [])
            if groups:
                group = groups[0]
                options.append(
                    (
                        f"مقایسه گروه‌ها — {group.group_header}",
                        ["گروه", "تعداد", "جمع", "میانگین", "میانه", "کمینه", "بیشینه"],
                        [
                            [
                                item.label,
                                format_number(item.count),
                                format_number(item.total),
                                format_number(item.mean),
                                format_number(item.median),
                                format_number(item.minimum),
                                format_number(item.maximum),
                            ]
                            for item in group.groups
                        ],
                    )
                )
        audits = list(getattr(analysis, "audits", []) or [])
        if audits:
            options.append(
                (
                    "حسابرسی جمع‌های اعلامی",
                    ["شاخص", "مقدار اعلامی", "مقدار محاسبه", "اختلاف", "وضعیت"],
                    [
                        [
                            item.label,
                            format_number(item.declared),
                            format_number(item.computed),
                            format_number(item.difference),
                            str(getattr(item.status, "value", item.status)),
                        ]
                        for item in audits
                    ],
                )
            )
        reconciliation = list(getattr(analysis, "reconciliation", []) or [])
        if reconciliation:
            options.append(
                (
                    "تطبیق تعهدات",
                    ["مؤسسه", "جمع وام", "جمع تعهد", "اختلاف", "وضعیت"],
                    [
                        [
                            item.institution,
                            format_number(item.loan_total),
                            format_number(item.obligation_total),
                            format_number(item.difference),
                            str(getattr(item.status, "value", item.status)),
                        ]
                        for item in reconciliation
                    ],
                )
            )
        validation = list(getattr(analysis, "validation", []) or [])
        if validation:
            options.append(
                (
                    "کیفیت داده",
                    ["بررسی", "دامنه", "وضعیت", "تعداد", "توضیح"],
                    [
                        [
                            item.label,
                            str(item.scope),
                            validation_label(item.status),
                            format_number(item.affected),
                            str(item.detail),
                        ]
                        for item in validation
                    ],
                )
            )

        self._table_options = options
        if not options:
            layout.addWidget(EmptyState("database", "جدولی برای نمایش نیست", "این فایل جدول داده قابل اتکایی نداشت.", self._palette))
            return
        selector_row = QHBoxLayout()
        selector_row.addWidget(make_label("جدول:", "Muted"))
        self._table_combo = QComboBox()
        self._table_combo.addItems([option[0] for option in options])
        self._table_combo.setMinimumWidth(260)
        self._table_combo.currentIndexChanged.connect(self._show_table)
        selector_row.addWidget(self._table_combo, 1)
        selector_row.addStretch(1)
        layout.addLayout(selector_row)
        self._table_widget = QTableWidget()
        self._table_widget.setMinimumHeight(320)
        self._table_widget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table_widget.setAlternatingRowColors(True)
        layout.addWidget(self._table_widget)
        self._show_table(0)

    def _show_table(self, index: int) -> None:
        if not (0 <= index < len(self._table_options)) or self._table_widget is None:
            return
        _, headers, rows = self._table_options[index]
        table = self._table_widget
        table.clear()
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        align_table_header(table)
        table.setRowCount(min(len(rows), 300))
        for row_index, row in enumerate(rows[:300]):
            for column_index, value in enumerate(row):
                formatted = format_cell(value)
                item = QTableWidgetItem(formatted)
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter)
                if len(formatted) > 60:
                    item.setToolTip(formatted)
                table.setItem(row_index, column_index, item)
        table.resizeColumnsToContents()
        for column_index in range(len(headers)):
            if table.columnWidth(column_index) > 340:
                table.setColumnWidth(column_index, 340)
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)

    # -- chart ---------------------------------------------------------- #
    def _populate_chart(self, analysis: Any) -> None:
        layout = self._clear_tab(self.chart_tab)
        self._chart_series = []
        self._series_combo = None
        integration = getattr(analysis, "integration", None)
        layout.addWidget(SectionHeader("نمای نمودار", "نمودار از اعداد همان تحلیل ساخته می‌شود"))
        if integration is not None and getattr(integration, "is_usable", False):
            series = list(getattr(integration, "series", []) or [])
            if series:
                self._chart_series = series
                selector = QHBoxLayout()
                selector.addWidget(make_label("شاخص:", "Muted"))
                self._series_combo = QComboBox()
                self._series_combo.addItems([item.metric for item in series])
                self._series_combo.currentIndexChanged.connect(self._render_integration_chart)
                selector.addWidget(self._series_combo, 1)
                selector.addStretch(1)
                layout.addLayout(selector)
                self._chart = BarChart()
                layout.addWidget(self._chart)
                self._chart_note = QLabel("")
                self._chart_note.setObjectName("Caption")
                self._chart_note.setWordWrap(True)
                layout.addWidget(self._chart_note)
                if self._palette is not None:
                    self._chart.apply_palette(self._palette)
                self._render_integration_chart(0)
                return

        generic = getattr(analysis, "generic", None)
        if generic is not None and getattr(generic, "trends", None):
            trend = generic.trends[0]
            if trend.points:
                self._chart = BarChart()
                layout.addWidget(self._chart)
                self._chart_note = QLabel(
                    f"روند «{trend.measure_header}» بر اساس «{trend.date_header}»؛ "
                    "مقدارها جمع هر دوره است."
                )
                self._chart_note.setObjectName("Caption")
                self._chart_note.setWordWrap(True)
                layout.addWidget(self._chart_note)
                if self._palette is not None:
                    self._chart.apply_palette(self._palette)
                self._chart.set_data(
                    [(point.period_label, point.total) for point in trend.points],
                    self._palette,
                )
                return
        if generic is not None and getattr(generic, "groups", None):
            group = generic.groups[0]
            if group.groups:
                self._chart = BarChart()
                layout.addWidget(self._chart)
                self._chart_note = QLabel(
                    f"جمع «{group.measure_header}» به تفکیک «{group.group_header}»."
                )
                self._chart_note.setObjectName("Caption")
                self._chart_note.setWordWrap(True)
                layout.addWidget(self._chart_note)
                if self._palette is not None:
                    self._chart.apply_palette(self._palette)
                self._chart.set_data(
                    [(item.label, item.total) for item in group.groups],
                    self._palette,
                )
                return
        if generic is not None:
            numeric = [
                column
                for column in (getattr(generic, "columns", []) or [])
                if column.kind.value == "number" and not column.is_identifier and column.total is not None
            ][:8]
            if numeric:
                self._chart = BarChart()
                layout.addWidget(self._chart)
                self._chart_note = QLabel("جمع ستون‌های عددی جدول اصلی؛ ستون‌های شناسه کنار گذاشته شده‌اند.")
                self._chart_note.setObjectName("Caption")
                self._chart_note.setWordWrap(True)
                layout.addWidget(self._chart_note)
                if self._palette is not None:
                    self._chart.apply_palette(self._palette)
                self._chart.set_data([(column.header, column.total) for column in numeric], self._palette)
                return
        self._chart = None
        layout.addWidget(
            EmptyState(
                "activity",
                "نمودار در دسترس نیست",
                "در این فایل سری زمانی، روند یا گروه عددی قابل نمودار پیدا نشد.",
                self._palette,
            )
        )

    def _render_integration_chart(self, index: int) -> None:
        if self._chart is None or not (0 <= index < len(self._chart_series)):
            return
        series = self._chart_series[index]
        integration = getattr(self._result.analysis, "integration", None) if self._result else None
        if integration is None:
            return
        period = integration.main_period()
        months = list(getattr(integration, "months", []) or [])
        totals: dict[str, float] = {}
        incomplete: set[str] = set()
        for cell in getattr(series, "cells", []):
            if getattr(cell, "period", "") != period:
                continue
            total = getattr(cell, "total", None)
            if total is None:
                continue
            if not getattr(cell, "complete", True):
                incomplete.add(cell.month)
                continue
            totals[cell.month] = totals.get(cell.month, 0.0) + float(total)
        data = [(month, totals[month]) for month in months if month in totals]
        self._chart.set_data(data, self._palette or self._chart.palette)
        excluded = "، ".join(month for month in months if month in incomplete)
        note = f"دوره {period}؛ جمع ماهانه شاخص «{series.metric}»."
        if excluded:
            note += f" ماه‌های ناقص نمایش داده نشدند: {excluded}."
        if not data:
            note = "برای این شاخص در دوره انتخابی، ماه کامل و قابل جمع‌بندی پیدا نشد."
        if self._chart_note is not None:
            self._chart_note.setText(note)

    # -- notes ---------------------------------------------------------- #
    def _populate_notes(self, analysis: Any, ai: dict[str, Any]) -> None:
        layout = self._clear_tab(self.notes_tab)
        response = ai.get("response") or {}
        layout.addWidget(SectionHeader("هشدارها و محدودیت‌ها", "شفاف درباره چیزی که اندازه‌گیری یا محاسبه نشده است"))

        warnings = list(getattr(analysis, "warnings", []) or [])
        if warnings:
            for warning in warnings:
                row = QLabel("• " + rich_text(soften_numbers(warning)))
                row.setWordWrap(True)
                layout.addWidget(row)
            layout.addWidget(horizontal_line())

        anomalies = list(getattr(analysis, "anomalies", []) or [])
        if anomalies:
            for anomaly in anomalies:
                frame = QFrame()
                frame.setObjectName("CardAlt")
                column = QVBoxLayout(frame)
                column.setContentsMargins(12, 10, 12, 10)
                column.setSpacing(3)
                head = QHBoxLayout()
                head.addWidget(Badge(severity_label(anomaly.severity), severity_tone(anomaly.severity)))
                location = " · ".join(
                    item for item in (getattr(anomaly, "sheet", ""), getattr(anomaly, "address", "")) if item
                )
                if location:
                    head.addWidget(make_label(ltr(location), "Caption"), 1)
                head.addStretch(1)
                column.addLayout(head)
                message = QLabel(rich_text(soften_numbers(anomaly.message)))
                message.setWordWrap(True)
                column.addWidget(message)
                if anomaly.detail:
                    detail = QLabel(rich_text(soften_numbers(anomaly.detail)))
                    detail.setObjectName("Caption")
                    detail.setWordWrap(True)
                    column.addWidget(detail)
                layout.addWidget(frame)
            layout.addWidget(horizontal_line())

        for title, key, tone in (
            ("داده‌های ناقص از دید مدل", "missing_data", "warning"),
            ("پیشنهادهای مدل", "recommendations", "info"),
            ("هشدارهای اعتبارسنجی پاسخ", "warnings", "warning"),
        ):
            items = list(response.get(key) or [])
            if not items:
                continue
            layout.addWidget(SectionHeader(title))
            for item in items:
                row = QLabel("• " + rich_text(soften_numbers(item)))
                row.setWordWrap(True)
                row.setObjectName("WarningText" if tone == "warning" else "Muted")
                layout.addWidget(row)

        unsupported = list(response.get("unsupported_numbers") or [])
        if unsupported:
            layout.addWidget(horizontal_line())
            layout.addWidget(SectionHeader("اعداد بدون منبع در پاسخ مدل"))
            caption = QLabel("این اعداد در محاسبات Python پیدا نشدند و مبنای تصمیم نیستند.")
            caption.setObjectName("Caption")
            caption.setWordWrap(True)
            layout.addWidget(caption)
            layout.addWidget(QLabel("، ".join(str(item) for item in unsupported[:20])))

        generic = getattr(analysis, "generic", None)
        quality = getattr(generic, "quality", None) if generic is not None else None
        if quality is not None and quality.notes:
            layout.addWidget(horizontal_line())
            layout.addWidget(SectionHeader("یادداشت روش محاسبه"))
            for note in quality.notes:
                row = QLabel("• " + rich_text(note))
                row.setObjectName("Caption")
                row.setWordWrap(True)
                layout.addWidget(row)
        if not warnings and not anomalies and not unsupported:
            layout.addWidget(make_label("هشدار ثبت‌شده‌ای وجود ندارد.", "Muted"))
        layout.addStretch(1)

    # ------------------------------------------------------------------ #
    @staticmethod
    def _clear_tab(area: QScrollArea) -> QVBoxLayout:
        layout = ResultsView._tab_layout(area)
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        return layout

    def _preview(self) -> None:
        if self._path and Path(self._path).exists():
            ReportPreviewDialog(self._path, self).exec()

    def _open_excel(self) -> None:
        if self._path:
            open_path(self._path)

    def _open_folder(self) -> None:
        if self._path:
            open_path(Path(self._path).parent)

    def _copy(self) -> None:
        result = self._result
        if result is None:
            return
        analysis = getattr(result, "analysis", None)
        ai = getattr(analysis, "ai", {}) or {}
        response = ai.get("response") or {}
        lines = [
            f"گزارش تحلیل: {Path(self._path).name if self._path else ''}",
            f"وضعیت: {getattr(result, 'status', '')}",
            f"خلاصه: {response.get('summary', '')}",
        ]
        answers = list(getattr(analysis, "request_answers", []) or [])
        if answers:
            lines.append("پاسخ درخواست:")
            for answer in answers:
                lines.append(
                    f"- {answer.title}: {answer.value or ''} — {soften_numbers(answer.text)}"
                )
        findings = list(response.get("findings") or [])
        if findings:
            lines.append("یافته‌ها:")
            lines.extend(f"- {item}" for item in findings)
        if self._path:
            lines.append(f"خروجی: {self._path}")
        QApplication.clipboard().setText("\n".join(lines))
        self.copy_button.setToolTip("کپی شد")
