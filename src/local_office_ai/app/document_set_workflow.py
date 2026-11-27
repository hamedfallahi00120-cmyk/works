"""Analyze a selected set of mixed administrative files in one workflow.

Excel files use the existing Structural Intelligence engine. CSV, DOCX, PPTX,
PDF and images use bounded readers and deterministic table statistics. A single
local Ollama request interprets the combined evidence; Python remains the
source of every computed statistic.
"""

from __future__ import annotations

import hashlib
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Sequence

from local_office_ai.ai.ollama_client import OllamaClient
from local_office_ai.core.exceptions import AIServiceError, EngineError, ProcessingAborted
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.settings import Settings
from local_office_ai.ingestion.models import ExtractedDocument, ExtractedTable
from local_office_ai.ingestion.reader import DocumentReader, EXCEL_SUFFIXES, SUPPORTED_DOCUMENT_SUFFIXES
from local_office_ai.ingestion.statistics import profile_excel_analysis, profile_table

LOGGER = get_logger(__name__)
_SKIP_DIRS = {".git", ".venv", "venv", "__pycache__", "output", "outputs", "processed", "temp", "tmp", "logs", "reports", "reference", "legacy", "tools"}


@dataclass
class SourceOutcome:
    path: str
    name: str
    file_type: str
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
    summary: str = ""
    analysis: Any = None


@dataclass
class DocumentSetResult:
    folder: str
    request: str = ""
    output_path: str = ""
    files: list[SourceOutcome] = field(default_factory=list)
    metrics: list[str] = field(default_factory=list)
    period: str = ""
    warnings: list[str] = field(default_factory=list)
    status: str = "completed"
    ai_summary: str = ""
    fallback_summary: str = ""
    ai_mode: str = "offline"
    statistics: list[dict[str, Any]] = field(default_factory=list)
    extracted_tables: list[ExtractedTable] = field(default_factory=list)
    extracted_text: list[tuple[str, str, str]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "folder": self.folder,
            "request": self.request,
            "output_path": self.output_path,
            "files": [item.__dict__.copy() | {"analysis": None} for item in self.files],
            "metrics": list(self.metrics),
            "period": self.period,
            "warnings": list(self.warnings),
            "status": self.status,
            "ai_summary": self.ai_summary,
            "ai_mode": self.ai_mode,
            "statistics": list(self.statistics),
        }


class DocumentSetWorkflow:
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
        self.reader = DocumentReader(settings)

    @staticmethod
    def scan(folder: Path | str, *, recursive: bool = True) -> list[Path]:
        root = Path(folder).expanduser().resolve()
        if not root.is_dir():
            raise EngineError(f"پوشه پیدا نشد: {root}")
        iterator = root.rglob("*") if recursive else root.iterdir()
        found: list[Path] = []
        for path in iterator:
            try:
                if path.is_symlink() or not path.is_file():
                    continue
                relative = path.relative_to(root)
                if any(part.lower() in _SKIP_DIRS for part in relative.parts[:-1]):
                    continue
                if path.name.startswith(("~$", ".")):
                    continue
                if path.suffix.lower() in SUPPORTED_DOCUMENT_SUFFIXES:
                    found.append(path.resolve())
            except (OSError, ValueError):
                continue
        return sorted(found, key=lambda item: str(item).casefold())

    def run(
        self,
        paths: Sequence[Path | str] | None = None,
        *,
        folder: Path | str | None = None,
        recursive: bool = True,
        request: str | None = None,
        model: str | None = None,
        use_ai: bool = True,
        output: Path | str | None = None,
        label: str | None = None,
    ) -> DocumentSetResult:
        if paths is None:
            if folder is None:
                raise EngineError("فایل یا پوشه‌ای برای تحلیل مشخص نشده است.")
            source_paths = self.scan(folder, recursive=recursive)
            source_label = str(Path(folder).resolve())
        else:
            source_paths = [Path(item).expanduser().resolve() for item in paths]
            source_label = label or (str(Path(folder).resolve()) if folder else "منابع انتخاب‌شده")
        unique: list[Path] = []
        seen: set[str] = set()
        for path in source_paths:
            key = str(path).casefold()
            if key not in seen:
                unique.append(path)
                seen.add(key)
        if not unique:
            return DocumentSetResult(folder=source_label, request=request or "", status="no_files", warnings=["فایل پشتیبانی‌شده‌ای برای تحلیل پیدا نشد."])
        if len(unique) > 100:
            raise EngineError("در هر نوبت حداکثر ۱۰۰ فایل قابل پردازش است. فایل‌ها را در چند دسته جدا کنید.")

        result = DocumentSetResult(folder=source_label, request=(request or "").strip())
        contexts: list[str] = []
        content_hashes: dict[str, str] = {}
        excel_workflow = None
        start = time.perf_counter()
        for index, path in enumerate(unique, start=1):
            self._check_abort()
            self.progress(f"شروع مرحله: دریافت منبع {index} از {len(unique)}")
            outcome = SourceOutcome(path=str(path), name=path.name, file_type=path.suffix.lower().lstrip(".").upper())
            try:
                self._validate_source(path)
                digest = self._sha256(path)
                prior = content_hashes.get(digest)
                if prior:
                    outcome.status = "تکراری - نادیده گرفته شد"
                    outcome.error = f"محتوای این فایل با «{prior}» یکسان است؛ برای جلوگیری از دوباره‌شماری وارد تحلیل نشد."
                    result.warnings.append(f"{path.name}: {outcome.error}")
                    result.files.append(outcome)
                    continue
                content_hashes[digest] = path.name
                self.progress(f"خواندن فایل: {path.name}")
                if path.suffix.lower() in EXCEL_SUFFIXES:
                    if excel_workflow is None:
                        from local_office_ai.app.workflow import ExcelIntelligenceWorkflow
                        excel_workflow = ExcelIntelligenceWorkflow(self.settings, progress=self.progress, abort_event=self.abort_event)
                    analysis = excel_workflow.analyze_only(path, request=result.request or None)
                    outcome.analysis = analysis
                    outcome.sheets = len(getattr(analysis.meta, "sheets", []) or [])
                    generic = getattr(analysis, "generic", None)
                    if generic is not None:
                        outcome.rows = int(getattr(generic, "row_count", 0) or 0)
                        outcome.columns = int(getattr(generic, "column_count", 0) or 0)
                        numeric = [column for column in getattr(generic, "columns", []) or [] if getattr(column, "is_numeric", False) and not getattr(column, "is_identifier", False) and not getattr(column, "is_temporal_key", False)]
                        outcome.numeric_columns = len(numeric)
                        outcome.metrics = len(numeric)
                        quality = getattr(generic, "quality", None)
                        if quality is not None:
                            outcome.missing_ratio = f"{quality.missing_ratio:.1%}"
                    integration = getattr(analysis, "integration", None)
                    if integration is not None:
                        outcome.metrics = max(outcome.metrics, len(getattr(integration, "series", []) or []))
                        outcome.months = "، ".join(getattr(integration, "months", []) or [])
                        outcome.periods = "، ".join(getattr(integration, "periods", []) or [])
                    outcome.issues = sum(1 for issue in getattr(analysis, "anomalies", []) or [] if getattr(getattr(issue, "severity", None), "value", "") in {"critical", "warning"})
                    outcome.status = "تحلیل شد"
                    numeric_profiles = profile_excel_analysis(path.name, analysis)
                    result.statistics.extend(numeric_profiles)
                    contexts.append(self._excel_context(path, analysis, numeric_profiles))
                    result.metrics.extend(item["column"] for item in numeric_profiles if item.get("column"))
                    outcome.summary = f"{outcome.sheets} برگه، {outcome.rows} ردیف در جدول اصلی، {outcome.numeric_columns} ستون عددی معنادار"
                else:
                    document = self.reader.read(path)
                    outcome.sheets = document.page_count or document.slide_count or len(document.tables)
                    outcome.rows = sum(table.row_count for table in document.tables)
                    outcome.columns = max((table.column_count for table in document.tables), default=0)
                    table_profiles: list[dict[str, Any]] = []
                    for table in document.tables:
                        result.extracted_tables.append(table)
                        table_profiles.extend(profile_table(path.name, table.location, table.title, table.headers, table.rows))
                    result.statistics.extend(table_profiles)
                    outcome.numeric_columns = len({(item["table"], item["column"]) for item in table_profiles})
                    outcome.metrics = outcome.numeric_columns
                    outcome.issues = len(document.warnings) + sum(1 for table in document.tables for _ in table.warnings)
                    outcome.status = "تحلیل شد"
                    outcome.summary = self._document_summary(document, table_profiles)
                    for location, text in document.text_blocks:
                        if text.strip():
                            result.extracted_text.append((path.name, location, text[:20_000]))
                    contexts.append(self._document_context(document, table_profiles))
                    result.warnings.extend(f"{path.name}: {warning}" for warning in document.warnings)
                self.progress(f"پایان مرحله: دریافت منبع {index} ({0.0:.1f} ثانیه)")
            except Exception as exc:  # noqa: BLE001 - one unreadable file must not stop the set
                outcome.status = "خطا"
                outcome.error = f"{type(exc).__name__}: {exc}"
                result.warnings.append(f"{path.name}: {outcome.error}")
                LOGGER.warning("document ingestion failed: %s (%s)", path, exc)
            result.files.append(outcome)

        successful = [item for item in result.files if item.status == "تحلیل شد"]
        result.metrics = list(dict.fromkeys(result.metrics))
        result.fallback_summary = self._fallback_summary(result)
        if use_ai and successful:
            self.progress("شروع مرحله: تفسیر چندمنبعی با هوش مصنوعی")
            try:
                result.ai_summary = self._ask_model(
                    request=result.request,
                    contexts=contexts,
                    model=model,
                )
                result.ai_mode = "model"
                self.progress("پایان مرحله: تفسیر چندمنبعی با هوش مصنوعی (0.0 ثانیه)")
            except AIServiceError as exc:
                result.ai_mode = "offline"
                result.warnings.append(f"تحلیل زبانی Qwen انجام نشد؛ خروجی محاسباتی حفظ شد: {exc}")
                LOGGER.warning("multi-source AI failed: %s", exc)
        elif not use_ai:
            result.ai_mode = "offline"
        output_path = Path(output) if output else self._default_output()
        from local_office_ai.excel.mixed_report_writer import MixedReportWriter

        result.output_path = str(MixedReportWriter().build(result, output_path))
        result.status = "completed" if successful and len(successful) == len(result.files) and not result.warnings else ("completed_with_warnings" if successful else "failed")
        LOGGER.info("document-set run completed files=%d successful=%d elapsed=%.2fs", len(result.files), len(successful), time.perf_counter() - start)
        self.progress(f"خروجی ساخته شد: {result.output_path}")
        return result

    def _validate_source(self, path: Path) -> None:
        if not path.is_file():
            raise EngineError(f"فایل پیدا نشد: {path}")
        if path.suffix.lower() not in SUPPORTED_DOCUMENT_SUFFIXES:
            raise EngineError(f"قالب فایل پشتیبانی نمی‌شود: {path.suffix}")
        if path.stat().st_size > self.settings.max_file_mb * 1024 * 1024:
            raise EngineError(f"حجم فایل {path.name} از سقف مجاز بیشتر است.")

    @staticmethod
    def _sha256(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    @staticmethod
    def _document_summary(document: ExtractedDocument, profiles: list[dict[str, Any]]) -> str:
        parts = [f"نوع: {document.kind}"]
        if document.page_count:
            parts.append(f"صفحه: {document.page_count}")
        if document.slide_count:
            parts.append(f"اسلاید: {document.slide_count}")
        parts.append(f"جدول: {len(document.tables)}")
        parts.append(f"ردیف استخراج‌شده: {document.extracted_rows}")
        parts.append(f"شاخص عددی معتبر: {len(profiles)}")
        return "؛ ".join(parts)

    @staticmethod
    def _excel_context(path: Path, analysis: Any, profiles: list[dict[str, Any]]) -> str:
        parts = [f"منبع Excel: {path.name}"]
        for sheet in getattr(analysis.meta, "sheets", []) or []:
            parts.append(f"برگه: {sheet.title}، ابعاد: {sheet.max_row}×{sheet.max_col}")
        generic = getattr(analysis, "generic", None)
        if generic is not None:
            parts.append(f"جدول اصلی: {generic.table_sheet}!{generic.table_address}، ردیف: {generic.row_count}، ستون: {generic.column_count}")
        for profile in profiles[:20]:
            parts.append(
                f"شاخص «{profile['column']}»: n={profile['n']}، میانگین={profile['mean']}، میانه={profile['median']}، حداقل={profile['min']}، حداکثر={profile['max']}، مجموع={profile['sum']}"
            )
        warnings = list(getattr(analysis, "warnings", []) or [])
        if warnings:
            parts.append("هشدارهای منبع: " + "؛ ".join(warnings[:6]))
        return "\n".join(parts)

    @staticmethod
    def _document_context(document: ExtractedDocument, profiles: list[dict[str, Any]]) -> str:
        parts = [f"منبع {document.kind}: {document.source.name}"]
        parts.append(f"عنوان: {document.title}")
        if document.page_count:
            parts.append(f"تعداد صفحات: {document.page_count}")
        if document.slide_count:
            parts.append(f"تعداد اسلایدها: {document.slide_count}")
        if profiles:
            parts.append("شاخص‌های عددی محاسبه‌شده توسط Python:")
            for profile in profiles[:20]:
                parts.append(
                    f"{profile['table']} / {profile['column']}: تعداد={profile['n']}، میانگین={profile['mean']}، میانه={profile['median']}، حداقل={profile['min']}، حداکثر={profile['max']}، مجموع={profile['sum']}"
                )
        text = document.compact_text(1800)
        if text:
            parts.append("گزیده محتوای منبع:\n" + text)
        if document.warnings:
            parts.append("محدودیت استخراج: " + "؛ ".join(document.warnings[:5]))
        return "\n".join(parts)

    def _ask_model(self, *, request: str, contexts: list[str], model: str | None) -> str:
        selected_model = model
        if not selected_model:
            from local_office_ai.ai.analyzer import AIAnalyzer
            selected = AIAnalyzer(self.settings).available_models()
            selected_model = selected[0].name if selected else ""
        if not selected_model:
            raise AIServiceError("مدل سازگار Ollama پیدا نشد.")
        configured_budget = int(getattr(self.settings, "max_ai_context_chars", 9_000) or 9_000)
        prompt_budget = int(getattr(self.settings, "max_ai_prompt_chars", 11_000) or 11_000)
        budget = min(10_000, max(2_000, configured_budget), max(2_000, prompt_budget - 1_500))
        used: list[str] = []
        remaining = budget
        truncated = False
        for context in contexts:
            if remaining <= 0:
                truncated = True
                break
            part = context[:remaining]
            if part:
                used.append(part)
                remaining -= len(part)
            if len(part) < len(context):
                truncated = True
        context_text = "\n\n--- منبع بعدی ---\n\n".join(used)
        if truncated:
            context_text += "\n\nتوجه: به دلیل سقف متن مدل، بخشی از جزئیات بعضی منابع در این درخواست ارسال نشده است. فهرست همه منابع و نتایج محاسباتی در گزارش Excel آمده است؛ درباره منبعی که جزئیاتش در متن حاضر نیست، نتیجه قطعی نساز."
        messages = [
            {
                "role": "system",
                "content": (
                    "تو دستیار تحلیل اداری و آماری فارسی هستی. فقط بر اساس شواهد منابع و آمار محاسبه‌شده پاسخ بده. "
                    "اعداد و شاخص‌های محاسبه‌شده توسط Python مرجع هستند و نباید دوباره حدس زده یا تغییر داده شوند. "
                    "واقعیت مشاهده‌شده را از تفسیر و فرضیه جدا کن. علت قطعی را بدون شواهد ادعا نکن. "
                    "اگر منابع با هم ناسازگارند، اختلاف را آشکارا گزارش کن. پاسخ را حرفه‌ای، روشن و فارسی بنویس."
                ),
            },
            {
                "role": "user",
                "content": (
                    f"درخواست کاربر:\n{request or 'منابع را از نظر اداری و آماری بررسی و مهم‌ترین یافته‌ها را جمع‌بندی کن.'}\n\n"
                    f"اطلاعات استخراج‌شده و آمار محاسبه‌شده:\n{context_text}\n\n"
                    "لطفاً خلاصه مدیریتی، یافته‌های مهم، مقایسه یا روند قابل پشتیبانی، محدودیت‌ها و پیشنهادهای مبتنی بر شواهد را ارائه کن. "
                    "اگر اطلاعات برای نتیجه‌ای کافی نیست، صریحاً بگو. از ساختن عدد یا جزئیات منبع خودداری کن."
                ),
            },
        ]
        response = OllamaClient(self.settings).chat(
            selected_model,
            messages,
            json_format=False,
            temperature=0.1,
            num_ctx=self.settings.ollama_num_ctx,
        )
        content = response.content.strip()
        if not content:
            raise AIServiceError("پاسخ متنی مدل خالی بود.")
        return content

    @staticmethod
    def _fallback_summary(result: DocumentSetResult) -> str:
        successful = [item for item in result.files if item.status == "تحلیل شد"]
        types: dict[str, int] = {}
        for item in successful:
            types[item.file_type] = types.get(item.file_type, 0) + 1
        type_text = "، ".join(f"{name}: {count}" for name, count in sorted(types.items())) or "—"
        return (
            f"از {len(result.files)} منبع، {len(successful)} منبع با موفقیت خوانده شد. "
            f"ترکیب منابع: {type_text}. تعداد جدول‌های استخراج‌شده {len(result.extracted_tables)} و "
            f"تعداد شاخص‌های عددی قابل محاسبه {len(result.statistics)} است. "
            "برای تفسیر مدیریتی و نتیجه‌گیری معنایی، اتصال مدل محلی باید فعال باشد."
        )

    def _default_output(self) -> Path:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = self.settings.output_dir / f"گزارش_چندمنبعی_{stamp}.xlsx"
        counter = 1
        while target.exists():
            target = self.settings.output_dir / f"گزارش_چندمنبعی_{stamp}_{counter}.xlsx"
            counter += 1
        return target

    def _check_abort(self) -> None:
        if self.abort_event is not None and self.abort_event.is_set():
            raise ProcessingAborted("پردازش منابع متوقف شد.")
