"""Report specification and validated report content.

``ReportSpecification`` is the single decision record for delivery: format,
path, profile and optional DOCX template.  ``ReportContent`` is built from a
validated :class:`~local_office_ai.core.models.AnalysisResult`; a template is
never allowed to supply numerical content.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from local_office_ai.core.exceptions import ReportFormatError, ReportTemplateError
from local_office_ai.excel.normalization import (
    canonical_number,
    extract_numeric_claims,
    format_jalali,
    normalize_digits,
    parse_number,
)
from local_office_ai.reporting.formats import OutputFormat, ensure_extension_matches

SUPPORTED_PROFILES = {"clean", "audit"}


def _decimal_text(value: Any) -> str:
    if value is None:
        return "در دسترس نبود"
    if isinstance(value, (Decimal, int, float)) and not isinstance(value, bool):
        number = value if isinstance(value, Decimal) else Decimal(str(value))
        text = f"{number:,f}"
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        return text or "0"
    return str(value)


def _safe_key(value: str) -> str:
    return "".join(char if char.isalnum() or char in "._-" else "_" for char in value).strip("_")


def _cell_address(sheet: Any, address: Any) -> str:
    sheet_text = str(sheet or "").strip()
    address_text = str(address or "").strip()
    if sheet_text and address_text:
        return f"{sheet_text}!{address_text}"
    return address_text


def _status_value(value: Any) -> str:
    return str(getattr(value, "value", value) or "").strip()


def _status_fa(value: Any) -> str:
    raw = _status_value(value)
    return {
        "PASS": "قبول",
        "WARNING": "هشدار",
        "FAIL": "رد",
        "pass": "قبول",
        "warning": "هشدار",
        "fail": "رد",
        "ALLOWED": "مجاز",
        "BLOCKED": "مسدود",
        "REVIEW": "نیازمند بررسی",
    }.get(raw, raw or "—")


_DATE_CLAIM = re.compile(r"(?<!\d)(\d{4})[/-](\d{1,2})(?:[/-](\d{1,4}))?(?!\d)")


def _canonical_numbers_in(text: str) -> set[str]:
    """Canonical forms of every numeric claim in a text.

    A numeric date is also kept as one claim, so ``1403/04`` cannot pass merely
    because ``1403`` and ``4`` are approved separately.
    """

    found: set[str] = set()
    for token in extract_numeric_claims(text or ""):
        parsed = parse_number(token, numeric_context=True)
        if isinstance(parsed.value, Decimal):
            try:
                found.add(canonical_number(parsed.value))
            except (InvalidOperation, ValueError):
                continue
    for match in _DATE_CLAIM.finditer(normalize_digits(text or "")):
        parts = [part for part in match.groups() if part]
        found.add("/".join(str(int(part)) for part in parts))
    return found


def _text_is_grounded(text: str, allowed: set[str]) -> bool:
    """A narrative is safe only when every number it quotes is approved."""

    if not text.strip():
        return False
    return _canonical_numbers_in(text) <= allowed


@dataclass
class ReportSpecification:
    """What to generate, where, and with which presentation profile."""

    output_format: OutputFormat
    output_path: Path
    profile: str = "clean"
    template_path: Path | None = None
    preserve_template_figures: bool = False
    request: str = ""
    evidence: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not isinstance(self.output_format, OutputFormat):
            self.output_format = OutputFormat.from_token(str(self.output_format)) or OutputFormat.XLSX
        self.output_path = Path(self.output_path).expanduser()
        if self.template_path is None or str(self.template_path).strip() == "":
            self.template_path = None
        else:
            self.template_path = Path(self.template_path).expanduser()
        if not self.output_path.suffix:
            self.output_path = self.output_path.with_suffix(self.output_format.extension)

    def validate(self) -> None:
        """Reject an unsupported format, a mismatched suffix or a bad template."""

        if not self.output_format.supports_generation:
            raise ReportFormatError(
                f"تولید {self.output_format.label_fa} در این مرحله پشتیبانی نمی‌شود. "
                "خروجی Excel یا گزارش Word (.docx) انتخاب کنید."
            )
        ensure_extension_matches(self.output_path, self.output_format)
        if self.profile not in SUPPORTED_PROFILES:
            raise ReportFormatError(
                f"نمایه گزارش «{self.profile}» پشتیبانی نمی‌شود؛ نمایه‌های مجاز: clean، audit."
            )
        if self.output_format is OutputFormat.DOCX:
            if self.template_path is not None:
                self._validate_template_file()
                if self.output_path.resolve() == self.template_path.resolve():
                    raise ReportFormatError(
                        "مسیر خروجی با مسیر قالب Word یکسان است؛ قالب باید دست‌نخورده بماند."
                    )
        elif self.template_path is not None:
            raise ReportFormatError(
                f"قالب Word فقط با خروجی {OutputFormat.DOCX.label_fa} قابل استفاده است، "
                f"نه با {self.output_format.label_fa}."
            )

    def _validate_template_file(self) -> None:
        path = self.template_path
        assert path is not None
        if not path.exists():
            raise ReportTemplateError(f"قالب Word پیدا نشد: {path}")
        if not path.is_file():
            raise ReportTemplateError(f"مسیر قالب Word یک فایل نیست: {path}")
        if path.suffix.lower() != ".docx":
            raise ReportTemplateError(
                f"قالب پشتیبانی نمی‌شود: {path.suffix or 'بدون پسوند'}؛ قالب باید .docx باشد."
            )
        if path.stat().st_size <= 0:
            raise ReportTemplateError(f"قالب Word خالی است: {path}")
        if not zipfile.is_zipfile(path):
            raise ReportTemplateError(f"قالب Word معتبر نیست (فایل ZIP نیست): {path}")
        try:
            with zipfile.ZipFile(path) as archive:
                names = set(archive.namelist())
        except (OSError, zipfile.BadZipFile) as exc:
            raise ReportTemplateError(f"قالب Word قابل خواندن نیست: {path}") from exc
        if "word/document.xml" not in names:
            raise ReportTemplateError(
                f"قالب Word معتبر نیست (word/document.xml ندارد): {path}"
            )

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_format": self.output_format.value,
            "output_path": str(self.output_path),
            "profile": self.profile,
            "template_path": str(self.template_path) if self.template_path else None,
            "preserve_template_figures": self.preserve_template_figures,
            "request": self.request,
            "evidence": list(self.evidence),
        }


@dataclass(frozen=True)
class ReportValue:
    """One validated figure or text value with its provenance."""

    key: str
    label: str
    value: str
    unit: str = ""
    available: bool = True
    source: str = ""
    source_address: str = ""
    validated: bool = True

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "label": self.label,
            "value": self.value,
            "unit": self.unit,
            "available": self.available,
            "source": self.source,
            "source_address": self.source_address,
            "validated": self.validated,
        }


@dataclass(frozen=True)
class ReportRow:
    label: str
    value: str
    unit: str = ""
    source_address: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "value": self.value,
            "unit": self.unit,
            "source_address": self.source_address,
        }


@dataclass
class ReportSection:
    title: str
    rows: list[ReportRow] = field(default_factory=list)
    level: int = 1

    def to_dict(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "level": self.level,
            "rows": [row.to_dict() for row in self.rows],
        }


@dataclass
class ReportTable:
    """A table whose cells already contain validated, display-ready values."""

    title: str
    headers: list[str]
    rows: list[list[str]] = field(default_factory=list)
    caption: str = ""
    source_note: str = ""
    statistical_note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "headers": list(self.headers),
            "rows": [list(row) for row in self.rows],
            "caption": self.caption,
            "source_note": self.source_note,
            "statistical_note": self.statistical_note,
        }


@dataclass
class ReportContent:
    """Validated content for a report; never sourced from a template."""

    title: str
    subtitle: str = ""
    sections: list[ReportSection] = field(default_factory=list)
    values: dict[str, ReportValue] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    source_kind: str = "python_analysis"
    validated: bool = True
    generated_at: str = ""
    tables: list[ReportTable] = field(default_factory=list)
    executive_summary: list[str] = field(default_factory=list)
    conclusions: list[str] = field(default_factory=list)
    recommendations: list[str] = field(default_factory=list)
    source_notes: list[str] = field(default_factory=list)
    statistical_notes: list[str] = field(default_factory=list)
    quality_gate: dict[str, Any] | None = None
    provenance_summary: dict[str, Any] | None = None

    @classmethod
    def from_analysis(
        cls,
        analysis: Any,
        *,
        title: str | None = None,
        subtitle: str = "",
        request: str = "",
        generated_at: date | datetime | None = None,
    ) -> "ReportContent":
        """Build content only from the analysed workbook, never from a template."""

        if analysis is None:
            raise ReportFormatError("برای ساخت محتوای گزارش، نتیجه تحلیل لازم است.")

        meta = getattr(analysis, "meta", None)
        source_file = Path(getattr(meta, "path", "") or "کارپوشه").name
        primary = getattr(analysis, "primary_table", None)
        integration = getattr(analysis, "integration", None)
        periods = list(getattr(integration, "periods", []) or []) if integration is not None else []
        months = list(getattr(integration, "months", []) or []) if integration is not None else []
        request_answers = list(getattr(analysis, "request_answers", []) or [])
        report_title = title or f"گزارش تحلیل {source_file}"
        stamp = format_jalali(generated_at or date.today())
        if isinstance(generated_at, datetime):
            stamp = f"{stamp} {generated_at.strftime('%H:%M')}"

        values: dict[str, ReportValue] = {}

        def add(
            key: str,
            label: str,
            value: Any,
            *,
            unit: str = "",
            source: str = "analysis",
            source_address: str = "",
            available: bool | None = None,
        ) -> None:
            text = _decimal_text(value)
            values[key] = ReportValue(
                key=key,
                label=label,
                value=text,
                unit=unit,
                available=(value is not None) if available is None else available,
                source=source,
                source_address=source_address,
                validated=True,
            )

        add("report.title", "عنوان گزارش", report_title, source="report")
        add("report.subtitle", "زیرعنوان گزارش", subtitle, source="report")
        add("report.generated_at", "تاریخ گزارش", stamp, source="report")
        add("report.source_file", "فایل منبع", source_file, source="workbook")
        add("report.request", "درخواست کاربر", request or "—", source="request")
        if primary is not None:
            add("data.sheet", "برگه جدول اصلی", primary.sheet, source="table")
            add("data.table_address", "نشانی جدول اصلی", primary.address, source="table")
            add("data.rows", "تعداد ردیف", getattr(primary, "row_count", None), source="table")
            add("data.columns", "تعداد ستون", getattr(primary, "column_count", None), source="table")
        add(
            "data.periods",
            "دوره‌های شناسایی‌شده",
            "، ".join(periods) if periods else "در دسترس نبود",
            source="periods",
            available=bool(periods),
        )
        add(
            "data.months",
            "ماه‌های شناسایی‌شده",
            "، ".join(months) if months else "در دسترس نبود",
            source="periods",
            available=bool(months),
        )

        for calculation in list(getattr(analysis, "calculations", []) or []):
            name = _safe_key(str(getattr(calculation, "name", "")))
            if not name:
                continue
            address = ""
            sources = list(getattr(calculation, "sources", []) or [])
            if sources:
                first = sources[0]
                address = _cell_address(
                    getattr(first, "sheet", ""),
                    getattr(first, "address", ""),
                )
            add(
                f"metric.{name}",
                str(getattr(calculation, "label", "") or name),
                getattr(calculation, "value", None),
                unit=str(getattr(calculation, "unit", "") or ""),
                source="calculation",
                source_address=address,
            )

        generic = getattr(analysis, "generic", None)
        for column in list(getattr(generic, "columns", []) or []) if generic is not None else []:
            if not getattr(column, "is_numeric", False):
                continue
            if getattr(column, "is_identifier", False) or getattr(column, "is_temporal_key", False):
                continue
            letter = _safe_key(str(getattr(column, "letter", "") or ""))
            if not letter:
                continue
            header = str(getattr(column, "header", "") or letter)
            for field_name, label in (
                ("total", "جمع"),
                ("mean", "میانگین"),
                ("median", "میانه"),
                ("minimum", "کمینه"),
                ("maximum", "بیشینه"),
            ):
                value = getattr(column, field_name, None)
                if value is None:
                    continue
                add(
                    f"stat.{letter}.{field_name}",
                    f"{header} - {label}",
                    value,
                    source="statistics",
                    source_address=_cell_address(
                        getattr(generic, "table_sheet", ""),
                        letter,
                    ),
                )

        for index, answer in enumerate(list(getattr(analysis, "request_answers", []) or []), start=1):
            add(f"answer.{index}.title", "عنوان پاسخ", getattr(answer, "title", ""), source="request_answer")
            answer_value = getattr(answer, "value", None)
            add(
                f"answer.{index}.value",
                "مقدار پاسخ",
                answer_value,
                source="request_answer",
                available=answer_value is not None and str(answer_value) != "",
            )
            add(f"answer.{index}.text", "متن پاسخ", getattr(answer, "text", ""), source="request_answer")

        warnings = list(getattr(analysis, "warnings", []) or [])
        for index, warning in enumerate(warnings, start=1):
            add(f"warning.{index}", "هشدار", str(warning), source="warning")
        anomalies = list(getattr(analysis, "anomalies", []) or [])
        for index, anomaly in enumerate(anomalies, start=1):
            message = getattr(anomaly, "message", "")
            severity_raw = getattr(anomaly, "severity", "")
            severity = getattr(severity_raw, "value", severity_raw) or ""
            add(f"issue.{index}", f"مورد نیازمند بررسی ({severity})", message, source="anomaly")

        identity_rows = [
            ReportRow("فایل منبع", source_file),
            ReportRow("تاریخ گزارش", stamp),
            ReportRow("درخواست کاربر", request or "—"),
        ]
        sections: list[ReportSection] = [ReportSection("شناسنامه گزارش", identity_rows)]
        metric_rows = [
            ReportRow(item.label, item.value, item.unit, item.source_address)
            for key, item in values.items()
            if key.startswith("metric.")
        ]
        if metric_rows:
            sections.append(ReportSection("شاخص‌های محاسبه‌شده", metric_rows))
        stat_rows = [
            ReportRow(item.label, item.value, item.unit, item.source_address)
            for key, item in values.items()
            if key.startswith("stat.")
        ]
        if stat_rows:
            sections.append(ReportSection("آمار توصیفی", stat_rows))
        answer_rows = [
            ReportRow(item.label, item.value, "", "")
            for key, item in values.items()
            if key.startswith("answer.") and key.endswith(".text")
        ]
        if answer_rows:
            sections.append(ReportSection("پاسخ درخواست", answer_rows))
        if warnings:
            sections.append(ReportSection("هشدارها", [ReportRow("هشدار", str(item)) for item in warnings]))

        tables: list[ReportTable] = [
            ReportTable(
                "شناسنامه گزارش",
                ["شرح", "مقدار"],
                [[row.label, row.value] for row in identity_rows],
                caption="جدول شناسنامه گزارش",
            )
        ]
        if metric_rows:
            tables.append(
                ReportTable(
                    "شاخص‌های محاسبه‌شده",
                    ["شاخص", "مقدار", "واحد", "منبع"],
                    [
                        [row.label, row.value, row.unit or "—", row.source_address or "—"]
                        for row in metric_rows
                    ],
                    caption="جدول شاخص‌های محاسبه‌شده",
                    source_note="هر مقدار از محاسبه قطعی Python و نشانی منبع همان سطر گرفته شده است.",
                )
            )
        if stat_rows:
            tables.append(
                ReportTable(
                    "آمار توصیفی",
                    ["شاخص", "مقدار", "واحد", "منبع"],
                    [
                        [row.label, row.value, row.unit or "—", row.source_address or "—"]
                        for row in stat_rows
                    ],
                    caption="جدول آمار توصیفی",
                    statistical_note=(
                        "آمار توصیفی فقط روی ستون‌های عددی غیرشناسه و غیرتاریخی محاسبه شده است."
                    ),
                )
            )
        if request_answers:
            answer_table_rows: list[list[str]] = []
            for answer in request_answers:
                answer_table_rows.append(
                    [
                        str(getattr(answer, "title", "") or "—"),
                        _decimal_text(getattr(answer, "value", None)),
                        str(getattr(answer, "text", "") or "—"),
                    ]
                )
            tables.append(
                ReportTable(
                    "پاسخ درخواست",
                    ["عنوان", "مقدار", "متن پاسخ"],
                    answer_table_rows[:60],
                    caption="جدول پاسخ درخواست کاربر",
                    source_note="پاسخ‌ها از تحلیل داده و ماتریس دوره‌ها ساخته شده‌اند؛ مدل فقط متن را بیان می‌کند.",
                )
            )

        quality = getattr(analysis, "quality", None)
        if quality is not None and getattr(quality, "findings", None):
            counts = quality.counts() if hasattr(quality, "counts") else {}
            quality_rows = [
                [status_label, str(counts.get(status_key, 0))]
                for status_key, status_label in (
                    ("PASS", "قبول"),
                    ("WARNING", "هشدار"),
                    ("FAIL", "رد"),
                )
            ]
            tables.append(
                ReportTable(
                    "کیفیت داده",
                    ["وضعیت", "تعداد بررسی"],
                    quality_rows,
                    caption=f"جدول خلاصه کیفیت داده (وضعیت کلی: {quality.status_label})",
                )
            )
            findings = [
                item
                for item in getattr(quality, "findings", [])
                if _status_value(getattr(item, "status", "")) != "PASS"
            ]
            if findings:
                tables.append(
                    ReportTable(
                        "یافته‌های کیفیت داده",
                        ["بررسی", "وضعیت", "تعداد", "توضیح", "محل"],
                        [
                            [
                                str(getattr(item, "title", "") or getattr(item, "code", "")),
                                _status_fa(getattr(item, "status", "")),
                                str(getattr(item, "count", 0)),
                                str(getattr(item, "message", "") or "—"),
                                _cell_address(
                                    getattr(item, "sheet", ""),
                                    getattr(item, "address", ""),
                                )
                                or "—",
                            ]
                            for item in findings[:40]
                        ],
                        caption="جدول یافته‌های کیفیت داده",
                        source_note="هیچ مقدار مشکوکی حذف یا اصلاح نشده است؛ فقط ثبت و گزارش شده است.",
                    )
                )

        statistical = getattr(analysis, "statistical_reconciliation", None)
        if statistical is not None:
            reconciliation_rows: list[list[str]] = []
            for item in getattr(statistical, "reconciliations", []) or []:
                if _status_value(getattr(item, "status", "")) == "PASS":
                    continue
                reconciliation_rows.append(
                    [
                        "تطبیق جمع",
                        str(getattr(item, "label", "") or getattr(item, "measure", "") or "—"),
                        (
                            f"انتظار {_decimal_text(getattr(item, 'expected', None))} / "
                            f"محاسبه {_decimal_text(getattr(item, 'calculated', None))} / "
                            f"اختلاف {_decimal_text(getattr(item, 'difference', None))}"
                        ),
                        _status_fa(getattr(item, "status", "")),
                        str(getattr(item, "explanation", "") or "—"),
                    ]
                )
            for item in getattr(statistical, "comparisons", []) or []:
                if _status_value(getattr(item, "status", "")) == "PASS":
                    continue
                change = getattr(item, "change", None)
                reconciliation_rows.append(
                    [
                        "مقایسه دوره",
                        str(getattr(item, "metric", "") or "—"),
                        (
                            f"{_decimal_text(getattr(change, 'previous', None))} به "
                            f"{_decimal_text(getattr(change, 'current', None))} / "
                            f"تغییر {_decimal_text(getattr(change, 'absolute', None))} / "
                            f"درصد {_decimal_text(getattr(change, 'percent', None))}"
                        ),
                        _status_fa(getattr(item, "status", "")),
                        str(getattr(item, "explanation", "") or "—"),
                    ]
                )
            for item in getattr(statistical, "averages", []) or []:
                if _status_value(getattr(item, "status", "")) == "PASS":
                    continue
                reconciliation_rows.append(
                    [
                        "میانگین",
                        str(getattr(item, "label", "") or "—"),
                        f"مقدار {_decimal_text(getattr(item, 'value', None))}",
                        _status_fa(getattr(item, "status", "")),
                        str(getattr(item, "explanation", "") or "—"),
                    ]
                )
            if reconciliation_rows:
                tables.append(
                    ReportTable(
                        "تطبیق آماری و صحت محاسبات",
                        ["نوع", "شاخص", "نتیجه", "وضعیت", "توضیح"],
                        reconciliation_rows[:60],
                        caption="جدول موارد تطبیق آماری نیازمند بررسی",
                        statistical_note=(
                            "هر نتیجه شامل مقدار انتظاری، مقدار محاسبه‌شده و اختلاف است؛ "
                            "مقدار غایب هرگز صفر تلقی نشده است."
                        ),
                    )
                )

        # ---------------------------------------------------------------- #
        # Narrative: Python facts first, then model text only when grounded.
        # ---------------------------------------------------------------- #
        allowed_numbers: set[str] = set()
        for item in values.values():
            allowed_numbers |= _canonical_numbers_in(item.value)
        for table in tables:
            for row in table.rows:
                for cell in row:
                    allowed_numbers |= _canonical_numbers_in(cell)

        ai_summary = ""
        ai_findings: list[str] = []
        ai_recommendations: list[str] = []
        ai_payload = getattr(analysis, "ai", None)
        if isinstance(ai_payload, dict):
            response = ai_payload.get("response")
            if isinstance(response, dict):
                ai_summary = str(response.get("summary") or "").strip()
                ai_findings = [str(item) for item in response.get("findings") or [] if str(item).strip()]
                ai_recommendations = [
                    str(item) for item in response.get("recommendations") or [] if str(item).strip()
                ]

        executive_summary: list[str] = [
            f"این گزارش بر پایه تحلیل قطعی کارپوشه «{source_file}» تهیه شده است."
        ]
        if primary is not None:
            executive_summary.append(
                f"جدول اصلی در برگه «{primary.sheet}» با "
                f"{getattr(primary, 'row_count', 0)} ردیف و "
                f"{getattr(primary, 'column_count', 0)} ستون شناسایی شد."
            )
        if periods:
            executive_summary.append("دوره‌های شناسایی‌شده: " + "، ".join(periods) + ".")
        if months:
            executive_summary.append("ماه‌های شناسایی‌شده: " + "، ".join(months) + ".")
        executive_summary.append(
            f"{len(metric_rows)} شاخص محاسبه‌شده و {len(stat_rows)} آماره توصیفی "
            "از داده‌های معتبر استخراج شد."
        )
        if quality is not None:
            executive_summary.append(
                f"وضعیت کلی کیفیت داده «{getattr(quality, 'status_label', '—')}» است."
            )
        if statistical is not None:
            executive_summary.append(
                f"وضعیت تطبیق آماری «{getattr(statistical, 'status_label', '—')}» است."
            )
        dropped_model_text = False
        if ai_summary:
            if _text_is_grounded(ai_summary, allowed_numbers):
                executive_summary.append(ai_summary)
            else:
                dropped_model_text = True
        for finding in ai_findings:
            if _text_is_grounded(finding, allowed_numbers):
                executive_summary.append(finding)
            else:
                dropped_model_text = True

        conclusions: list[str] = []
        if quality is not None:
            quality_status = _status_value(getattr(quality, "status", ""))
            if quality_status == "FAIL":
                conclusions.append(
                    "کیفیت داده در وضعیت «رد» است؛ پیش از استناد به نتایج، "
                    "یافته‌های کیفیت داده باید بررسی و اصلاح شوند."
                )
            elif quality_status == "WARNING":
                conclusions.append(
                    "کیفیت داده هشدار دارد؛ نتایج قابل استفاده است اما محدودیت‌های ثبت‌شده باید در تفسیر لحاظ شود."
                )
            else:
                conclusions.append("بررسی‌های کیفیت داده بدون هشدار مؤثر به پایان رسید.")
        if statistical is not None:
            statistical_status = _status_value(getattr(statistical, "status", ""))
            if statistical_status == "FAIL":
                conclusions.append(
                    "دست‌کم یک تطبیق آماری رد شده است؛ جمع‌های اعلام‌شده با محاسبه مستقل Python هم‌خوان نیستند."
                )
            elif statistical_status == "WARNING":
                conclusions.append(
                    "تطبیق آماری هشدار دارد؛ برخی جمع‌ها یا مقایسه‌ها به دلیل داده ناقص یا رابطه نامشخص نیازمند بازبینی‌اند."
                )
            else:
                conclusions.append("تطبیق‌های آماری انجام‌شده با محاسبات مستقل Python هم‌خوان بودند.")
        if anomalies:
            conclusions.append(
                f"{len(anomalies)} مورد نیازمند بررسی ثبت شده است که باید پیش از تصمیم‌گیری نهایی دیده شود."
            )
        if not conclusions:
            conclusions.append(
                "هیچ ناسازگاری قطعی در محاسبات ثبت نشد؛ با این حال نتایج باید در بستر محدودیت‌های داده تفسیر شوند."
            )

        recommendations: list[str] = []
        if quality is not None and _status_value(getattr(quality, "status", "")) in {"WARNING", "FAIL"}:
            recommendations.append(
                "یافته‌های کیفیت داده و نشانی سلول‌های آن‌ها را بررسی کنید و مقدارهای مشکوک را در فایل منبع اصلاح کنید."
            )
        if statistical is not None and _status_value(getattr(statistical, "status", "")) in {"WARNING", "FAIL"}:
            recommendations.append(
                "جمع‌ها و مقایسه‌های علامت‌خورده را با ستون‌های جزء و دوره‌های کامل دوباره بررسی کنید."
            )
        if warnings:
            recommendations.append(
                "هشدارهای فهرست‌شده را پیش از انتشار گزارش بازبینی کنید؛ هیچ هشداری به‌صورت خودکار حذف نشده است."
            )
        recommendations.append(
            "برای تصمیم‌گیری، هر عدد نهایی را با نشانی منبع همان سطر در فایل منبع تطبیق دهید."
        )
        recommendations.append(
            "مقادیر «در دسترس نبود» را صفر تلقی نکنید؛ نبود داده باید جداگانه تأمین یا اعلام شود."
        )
        for item in ai_recommendations:
            if _text_is_grounded(item, allowed_numbers):
                recommendations.append(item)
            else:
                dropped_model_text = True
        if dropped_model_text:
            warnings.append("روایت مدل به دلیل داشتن عدد بدون منبع معتبر درج نشد.")

        source_notes: list[str] = [
            "همه اعداد این گزارش با محاسبات قطعی Python تولید شده‌اند.",
            f"فایل منبع: {source_file}",
            "مقادیر «در دسترس نبود» نشانه نبود داده معتبر است و صفر تلقی نمی‌شود.",
        ]
        if primary is not None:
            source_notes.append(
                f"جدول اصلی: {_cell_address(getattr(primary, 'sheet', ''), getattr(primary, 'address', ''))}"
            )
        statistical_notes: list[str] = [
            "آمار توصیفی فقط روی ستون‌های عددی غیرشناسه و غیرتاریخی محاسبه شده است.",
            "شناسه‌ها و کلیدهای زمانی در میانگین، واریانس و همبستگی وارد نمی‌شوند.",
        ]
        methodology = list(getattr(generic, "methodology", []) or []) if generic is not None else []
        statistical_notes.extend(str(item) for item in methodology[:12])
        if generic is not None and getattr(generic, "correlations", None):
            statistical_notes.append("همبستگی رابطه علّی نیست و فقط شدت همراهی را نشان می‌دهد.")

        gate = getattr(analysis, "quality_gate", None)
        provenance = getattr(analysis, "provenance", None)
        quality_gate_payload = (
            gate.to_dict() if hasattr(gate, "to_dict") else gate
        )
        provenance_payload = (
            provenance.to_dict() if hasattr(provenance, "to_dict") else provenance
        )
        if gate is not None and getattr(gate, "finalized", False):
            gate_status = _status_value(getattr(gate, "status", ""))
            if gate_status == "FAIL":
                warnings.append(
                    "دروازه کیفیت شکست ثبت کرده است؛ این گزارش سند معتبر نهایی تلقی نمی‌شود."
                )
            elif gate_status == "WARNING":
                warnings.append(
                    "دروازه کیفیت هشدار دارد؛ گزارش موفقیت کامل نیست و باید با قید بازبینی استفاده شود."
                )

        return cls(
            title=report_title,
            subtitle=subtitle,
            sections=sections,
            values=values,
            warnings=warnings,
            source_kind="python_analysis",
            validated=True,
            generated_at=stamp,
            tables=tables,
            executive_summary=executive_summary,
            conclusions=conclusions,
            recommendations=recommendations,
            source_notes=source_notes,
            statistical_notes=statistical_notes,
            quality_gate=quality_gate_payload,
            provenance_summary=provenance_payload,
        )

    # ------------------------------------------------------------------ #
    def placeholder_values(self) -> dict[str, str]:
        """Return only validated values that a template may display."""

        result: dict[str, str] = {"report.title": self.title, "report.subtitle": self.subtitle}
        for key, item in self.values.items():
            result[key] = item.value
            result[f"{key}.label"] = item.label
            if item.unit:
                result[f"{key}.unit"] = item.unit
        return result

    def approved_numbers(self) -> set[str]:
        """Every number this validated content is allowed to display."""

        numbers: set[str] = set()
        for text in (self.title, self.subtitle, self.generated_at):
            numbers |= _canonical_numbers_in(text)
        for item in self.values.values():
            numbers |= _canonical_numbers_in(item.value)
        for section in self.sections:
            for row in section.rows:
                numbers |= _canonical_numbers_in(row.label)
                numbers |= _canonical_numbers_in(row.value)
        for table in self.tables:
            for text in (table.title, table.caption, table.source_note, table.statistical_note):
                numbers |= _canonical_numbers_in(text)
            for header in table.headers:
                numbers |= _canonical_numbers_in(header)
            for row in table.rows:
                for cell in row:
                    numbers |= _canonical_numbers_in(cell)
        for group in (
            self.executive_summary,
            self.conclusions,
            self.recommendations,
            self.source_notes,
            self.statistical_notes,
            [str(item) for item in self.warnings],
        ):
            for text in group:
                numbers |= _canonical_numbers_in(text)
        return numbers

    def validate(self) -> list[str]:
        errors: list[str] = []
        if not self.title.strip():
            errors.append("عنوان گزارش خالی است.")
        if self.source_kind != "python_analysis":
            errors.append("محتوای گزارش باید از تحلیل معتبر Python ساخته شود.")
        if not self.validated:
            errors.append("محتوای گزارش به‌عنوان معتبر علامت‌گذاری نشده است.")
        for key, item in self.values.items():
            if not item.validated:
                errors.append(f"مقدار «{key}» از منبع معتبر ساخته نشده است.")
            if item.value is None:
                errors.append(f"مقدار «{key}» تهی است؛ باید «در دسترس نبود» ثبت شود.")
        for table in self.tables:
            if not table.title.strip():
                errors.append("عنوان جدول خالی است.")
            if not table.headers:
                errors.append(f"جدول «{table.title}» ستون ندارد.")
            for row in table.rows:
                if len(row) != len(table.headers):
                    errors.append(
                        f"ردیف جدول «{table.title}» با تعداد ستون‌ها هم‌خوان نیست."
                    )
                    break
                if any(cell is None for cell in row):
                    errors.append(f"جدول «{table.title}» سلول تهی دارد.")
                    break
        return errors

    def ensure_valid(self) -> None:
        errors = self.validate()
        if errors:
            raise ReportFormatError("محتوای گزارش معتبر نیست: " + "؛ ".join(errors))

    def to_dict(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "subtitle": self.subtitle,
            "source_kind": self.source_kind,
            "validated": self.validated,
            "generated_at": self.generated_at,
            "sections": [section.to_dict() for section in self.sections],
            "tables": [table.to_dict() for table in self.tables],
            "executive_summary": list(self.executive_summary),
            "conclusions": list(self.conclusions),
            "recommendations": list(self.recommendations),
            "source_notes": list(self.source_notes),
            "statistical_notes": list(self.statistical_notes),
            "values": {key: item.to_dict() for key, item in self.values.items()},
            "warnings": list(self.warnings),
            "quality_gate": dict(self.quality_gate) if self.quality_gate else None,
            "provenance_summary": (
                dict(self.provenance_summary) if self.provenance_summary else None
            ),
        }


@dataclass
class ReportGenerationResult:
    """Outcome of the format decision plus the selected writer."""

    output_path: Path
    output_format: OutputFormat
    writer: str
    template: Any = None
    placeholders_resolved: tuple[str, ...] = ()
    unresolved_placeholders: tuple[str, ...] = ()
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    content_validated: bool = True
    evidence: tuple[str, ...] = ()
    mode: str = ""
    sections_written: tuple[str, ...] = ()
    tables_written: int = 0
    validation: Any = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": str(self.output_path),
            "output_format": self.output_format.value,
            "writer": self.writer,
            "template": (
                self.template.to_dict()
                if hasattr(self.template, "to_dict")
                else self.template
            ),
            "placeholders_resolved": list(self.placeholders_resolved),
            "unresolved_placeholders": list(self.unresolved_placeholders),
            "warnings": list(self.warnings),
            "notes": list(self.notes),
            "content_validated": self.content_validated,
            "evidence": list(self.evidence),
            "mode": self.mode,
            "sections_written": list(self.sections_written),
            "tables_written": self.tables_written,
            "validation": (
                self.validation.to_dict()
                if hasattr(self.validation, "to_dict")
                else self.validation
            ),
        }
