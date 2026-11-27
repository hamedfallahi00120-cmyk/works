"""Report specification and validated report content.

``ReportSpecification`` is the single decision record for delivery: format,
path, profile and optional DOCX template.  ``ReportContent`` is built from a
validated :class:`~local_office_ai.core.models.AnalysisResult`; a template is
never allowed to supply numerical content.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from local_office_ai.core.exceptions import ReportFormatError, ReportTemplateError
from local_office_ai.excel.normalization import format_jalali
from local_office_ai.reporting.formats import OutputFormat, ensure_extension_matches

SUPPORTED_PROFILES = {"clean", "audit"}


def _decimal_text(value: Any) -> str:
    if value is None:
        return "در دسترس نبود"
    if isinstance(value, Decimal):
        text = f"{value:,.2f}".rstrip("0").rstrip(".")
        return text or "0"
    if isinstance(value, float):
        text = f"{value:,.2f}".rstrip("0").rstrip(".")
        return text or "0"
    return str(value)


def _safe_key(value: str) -> str:
    return "".join(char if char.isalnum() or char in "._-" else "_" for char in value).strip("_")


@dataclass
class ReportSpecification:
    """What to generate, where, and with which presentation profile."""

    output_format: OutputFormat
    output_path: Path
    profile: str = "clean"
    template_path: Path | None = None
    request: str = ""
    evidence: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not isinstance(self.output_format, OutputFormat):
            self.output_format = OutputFormat.from_token(str(self.output_format)) or OutputFormat.XLSX
        self.output_path = Path(self.output_path).expanduser()
        if self.template_path is not None:
            self.template_path = Path(self.template_path).expanduser()
        if not self.output_path.suffix:
            self.output_path = self.output_path.with_suffix(self.output_format.extension)

    def validate(self) -> None:
        """Reject an unsupported format, a mismatched suffix or a bad template."""

        if not self.output_format.supports_generation:
            raise ReportFormatError(
                f"تولید {self.output_format.label_fa} در این مرحله پشتیبانی نمی‌شود. "
                "خروجی Excel یا گزارش Word با قالب DOCX انتخاب کنید."
            )
        ensure_extension_matches(self.output_path, self.output_format)
        if self.profile not in SUPPORTED_PROFILES:
            raise ReportFormatError(
                f"نمایه گزارش «{self.profile}» پشتیبانی نمی‌شود؛ نمایه‌های مجاز: clean، audit."
            )
        if self.output_format.requires_template:
            if self.template_path is None:
                raise ReportTemplateError(
                    "برای تولید گزارش Word باید یک قالب DOCX مشخص شود (--template). "
                    "بدون قالب، فایل Excel تحت نام Word ساخته نمی‌شود."
                )
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
                address = f"{getattr(first, 'sheet', '')}!{getattr(first, 'address', '')}".strip("!")
            add(
                f"metric.{name}",
                str(getattr(calculation, "label", name)),
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
            letter = _safe_key(str(getattr(column, "letter", "")))
            if not letter:
                continue
            header = str(getattr(column, "header", letter))
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
                    source_address=f"{getattr(generic, 'table_sheet', '')}!{letter}",
                )

        for index, answer in enumerate(list(getattr(analysis, "request_answers", []) or []), start=1):
            add(f"answer.{index}.title", "عنوان پاسخ", getattr(answer, "title", ""), source="request_answer")
            add(f"answer.{index}.value", "مقدار پاسخ", getattr(answer, "value", "") or None, source="request_answer")
            add(f"answer.{index}.text", "متن پاسخ", getattr(answer, "text", ""), source="request_answer")

        warnings = list(getattr(analysis, "warnings", []) or [])
        for index, warning in enumerate(warnings, start=1):
            add(f"warning.{index}", "هشدار", warning, source="warning")
        anomalies = list(getattr(analysis, "anomalies", []) or [])
        for index, anomaly in enumerate(anomalies, start=1):
            message = getattr(anomaly, "message", "")
            severity = getattr(getattr(anomaly, "severity", None), "value", "")
            add(f"issue.{index}", f"مورد نیازمند بررسی ({severity})", message, source="anomaly")

        sections: list[ReportSection] = [
            ReportSection(
                "شناسنامه گزارش",
                [
                    ReportRow("فایل منبع", source_file),
                    ReportRow("تاریخ گزارش", stamp),
                    ReportRow("درخواست کاربر", request or "—"),
                ],
            )
        ]
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
            sections.append(ReportSection("هشدارها", [ReportRow("هشدار", item) for item in warnings]))

        return cls(
            title=report_title,
            subtitle=subtitle,
            sections=sections,
            values=values,
            warnings=warnings,
            source_kind="python_analysis",
            validated=True,
            generated_at=stamp,
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
            "values": {key: item.to_dict() for key, item in self.values.items()},
            "warnings": list(self.warnings),
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
    content_validated: bool = True
    evidence: tuple[str, ...] = ()

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
            "content_validated": self.content_validated,
            "evidence": list(self.evidence),
        }
