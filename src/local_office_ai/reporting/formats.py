"""Output-format contract and deterministic format resolution.

The engine has more than one output writer.  A requested format must be
resolved before a writer is selected so an Excel workbook can never be written
under a ``.docx`` name (or the other way round).  This module only decides and
validates formats; it does not write a report.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from local_office_ai.core.exceptions import ReportFormatError
from local_office_ai.excel.normalization import normalize_digits, normalize_text


class OutputFormat(str, Enum):
    """Formats the application can distinguish at the delivery boundary."""

    XLSX = "xlsx"
    DOCX = "docx"
    PDF = "pdf"

    @property
    def extension(self) -> str:
        return f".{self.value}"

    @property
    def label_fa(self) -> str:
        return {
            OutputFormat.XLSX: "کارپوشه Excel",
            OutputFormat.DOCX: "گزارش Word",
            OutputFormat.PDF: "سند PDF",
        }[self]

    @property
    def writer_name(self) -> str:
        return {
            OutputFormat.XLSX: "ExcelReportWriter",
            OutputFormat.DOCX: "DocxTemplateRenderer",
            OutputFormat.PDF: "PDF writer",
        }[self]

    @property
    def supports_generation(self) -> bool:
        return self is not OutputFormat.PDF

    @property
    def requires_template(self) -> bool:
        return self is OutputFormat.DOCX

    @classmethod
    def from_suffix(cls, suffix: str | None) -> "OutputFormat | None":
        """Map a file suffix to a format; ``None`` for an empty suffix.

        An unknown, non-empty suffix is rejected rather than guessed.
        """

        text = (suffix or "").strip().lower()
        if not text:
            return None
        if not text.startswith("."):
            text = f".{text}"
        mapping = {
            ".xlsx": cls.XLSX,
            ".docx": cls.DOCX,
            ".pdf": cls.PDF,
        }
        found = mapping.get(text)
        if found is None:
            supported = "، ".join(sorted(mapping))
            raise ReportFormatError(
                f"پسوند خروجی «{suffix}» پشتیبانی نمی‌شود. پسوندهای مجاز: {supported}."
            )
        return found

    @classmethod
    def from_token(cls, token: str | None) -> "OutputFormat | None":
        """Parse an explicit format token (English or Persian)."""

        raw = (token or "").strip().lower()
        if not raw:
            return None
        if raw.startswith("."):
            return cls.from_suffix(raw)
        normalized = normalize_text(normalize_digits(raw)).replace("‌", " ")
        normalized = re.sub(r"[\s_\-]+", "", normalized)
        aliases = {
            "xlsx": cls.XLSX,
            "excel": cls.XLSX,
            "اکسل": cls.XLSX,
            "کارپوشه": cls.XLSX,
            "کارپوشهاکسل": cls.XLSX,
            "docx": cls.DOCX,
            "word": cls.DOCX,
            "ورد": cls.DOCX,
            "سندword": cls.DOCX,
            "گزارشword": cls.DOCX,
            "pdf": cls.PDF,
            "پیدیاف": cls.PDF,
        }
        found = aliases.get(normalized)
        if found is None:
            raise ReportFormatError(
                f"قالب خروجی «{token}» شناخته نمی‌شود. قالب‌های مجاز: xlsx، docx، pdf."
            )
        return found


# Persian/English output-intent hints.  A bare "اکسل" usually describes the
# *input* file, so a format word only counts when it is attached to an output
# verb or an output noun.
_FORMAT_WORDS = {
    OutputFormat.DOCX: r"(?:word|docx|ورد)",
    OutputFormat.XLSX: r"(?:excel|xlsx|اکسل)",
    OutputFormat.PDF: r"(?:pdf|پی\s*دی\s*اف)",
}
_OUTPUT_CONTEXT_BEFORE = r"(?:گزارش|خروجی|سند|مدرک|نسخه|قالب|تهیه|تولید|ساخت|بساز|بده)"
_OUTPUT_CONTEXT_AFTER = r"(?:بساز|تولید|خروجی|بده|تهیه|می\s*خواهم|میخوام|لازم|می‌خواهم)"
_STRONG_TOKENS = {"docx", "xlsx", "pdf"}


@dataclass(frozen=True)
class OutputIntent:
    """The resolved output format plus the evidence used to choose it."""

    output_format: OutputFormat
    evidence: tuple[str, ...] = ()
    explicit_format: OutputFormat | None = None
    suffix_format: OutputFormat | None = None
    request_format: OutputFormat | None = None

    def to_dict(self) -> dict[str, object]:
        return {
            "output_format": self.output_format.value,
            "evidence": list(self.evidence),
            "explicit_format": self.explicit_format.value if self.explicit_format else None,
            "suffix_format": self.suffix_format.value if self.suffix_format else None,
            "request_format": self.request_format.value if self.request_format else None,
        }


def detect_requested_formats(text: str | None) -> list[OutputFormat]:
    """Find explicit output-format requests in a Persian/English request."""

    raw = (text or "").strip()
    if not raw:
        return []
    normalized = normalize_text(normalize_digits(raw)).lower()
    found: list[OutputFormat] = []
    for output_format, word in _FORMAT_WORDS.items():
        patterns = [
            rf"{_OUTPUT_CONTEXT_BEFORE}\s*(?:به\s*)?(?:قالب\s*)?{word}",
            rf"{word}\s*{_OUTPUT_CONTEXT_AFTER}",
        ]
        matched = any(re.search(pattern, normalized) for pattern in patterns)
        if not matched and output_format is not OutputFormat.PDF:
            matched = re.search(rf"(?<![\w]){re.escape(output_format.value)}(?![\w])", normalized) is not None
        if not matched:
            for token in _STRONG_TOKENS:
                if token == output_format.value and re.search(rf"(?<![\w]){token}(?![\w])", normalized):
                    matched = True
                    break
        if matched:
            found.append(output_format)
    return found


def resolve_output_intent(
    *,
    explicit_format: str | None = None,
    output_path: Path | str | None = None,
    request_text: str | None = None,
    default: OutputFormat = OutputFormat.XLSX,
) -> OutputIntent:
    """Resolve one output format or fail with an explicit conflict error."""

    explicit = OutputFormat.from_token(explicit_format)
    path = Path(output_path) if output_path else None
    suffix = OutputFormat.from_suffix(path.suffix) if path is not None and path.suffix else None
    requested = detect_requested_formats(request_text)

    evidence: list[str] = []
    if explicit is not None:
        evidence.append(f"قالب صریح: {explicit.label_fa}")
    if suffix is not None:
        evidence.append(f"پسوند خروجی: {path.suffix}")  # type: ignore[union-attr]
    if requested:
        evidence.append("درخواست کاربر: " + "، ".join(item.label_fa for item in requested))

    candidates = [item for item in (explicit, suffix, *requested) if item is not None]
    distinct = set(candidates)
    if len(distinct) > 1:
        labels = "، ".join(sorted(item.label_fa for item in distinct))
        raise ReportFormatError(
            f"قالب خروجی متناقض است ({labels}). یک قالب واحد انتخاب کنید یا پسوند فایل را اصلاح کنید."
        )

    chosen = candidates[0] if candidates else default
    request_format = requested[0] if requested else None
    return OutputIntent(
        output_format=chosen,
        evidence=tuple(evidence or [f"قالب پیش‌فرض: {default.label_fa}"]),
        explicit_format=explicit,
        suffix_format=suffix,
        request_format=request_format,
    )


def ensure_extension_matches(output_path: Path | str, output_format: OutputFormat) -> Path:
    """Validate that a writer's file suffix belongs to the selected format."""

    path = Path(output_path)
    suffix = path.suffix.lower()
    if not suffix:
        path = path.with_suffix(output_format.extension)
        suffix = output_format.extension
    if suffix != output_format.extension:
        raise ReportFormatError(
            f"پسوند خروجی «{suffix}» با نویسنده {output_format.label_fa} سازگار نیست؛ "
            f"پسوند درست «{output_format.extension}» است."
        )
    return path


def ensure_xlsx_delivery(
    *,
    output_path: Path | str | None,
    default_output_path: Path | str,
    explicit_format: str | None = None,
    request_text: str | None = None,
    context: str = "این گزارش",
) -> Path:
    """Resolve an Excel-only workflow (folder/mixed-document report)."""

    intent = resolve_output_intent(
        explicit_format=explicit_format,
        output_path=output_path,
        request_text=request_text,
    )
    if intent.output_format is not OutputFormat.XLSX:
        raise ReportFormatError(
            f"{context} فقط به صورت کارپوشه Excel (.xlsx) تولید می‌شود؛ "
            f"قالب درخواستی: {intent.output_format.label_fa}."
        )
    target = Path(output_path) if output_path else Path(default_output_path)
    return ensure_extension_matches(target, OutputFormat.XLSX)

