"""Qt threads that drive the existing engine without blocking the window.

Worker threads never touch widgets: they emit Qt signals and the main thread
renders them.  The engine callbacks only forward progress text, exactly like
the CLI does.
"""

from __future__ import annotations

import re
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PySide6.QtCore import QThread, Signal

from local_office_ai.core.exceptions import AIServiceError, EngineError, ProcessingAborted
from local_office_ai.core.settings import Settings

_STAGE_START = re.compile(r"^شروع مرحله:\s*(.+?)\s*$")
_STAGE_END = re.compile(r"^پایان مرحله:\s*(.+?)\s*\((\d+(?:\.\d+)?)\s*ثانیه\)\s*$")
_STAGE_FAIL = re.compile(r"^خطا در مرحله:\s*(.+?)\s*—\s*(.*)$")


@dataclass
class StageEvent:
    name: str
    state: str  # running | done | failed
    seconds: float = 0.0
    error: str = ""


@dataclass
class ChatReply:
    analysis: Any
    answers: list[Any] = field(default_factory=list)
    ai: dict[str, Any] = field(default_factory=dict)
    source: str = ""
    output_path: str = ""
    followup: bool = True
    error: str = ""
    result: Any = None


def friendly_error(exc: BaseException) -> str:
    if isinstance(exc, ProcessingAborted):
        return f"پردازش متوقف شد: {exc}"
    if isinstance(exc, (EngineError, AIServiceError)):
        return str(exc)
    if isinstance(exc, PermissionError):
        return "دسترسی به فایل یا پوشه ممکن نشد. مسیر و مجوز دسترسی را بررسی کنید."
    if isinstance(exc, FileNotFoundError):
        return f"فایل یا پوشه پیدا نشد: {exc.filename or exc}"
    return f"خطای غیرمنتظره ({type(exc).__name__}): {exc}"


class _StageMixin:
    stage = Signal(object)
    progress = Signal(str)

    def _emit_stage(self, name: str, state: str, seconds: float = 0.0, error: str = "") -> None:
        self.stage.emit(StageEvent(name=name, state=state, seconds=seconds, error=error))

    def _on_message(self, message: str) -> None:
        self.progress.emit(str(message))
        text = str(message).strip()
        start = _STAGE_START.match(text)
        if start:
            self._emit_stage(start.group(1).strip(), "running")
            return
        end = _STAGE_END.match(text)
        if end:
            self._emit_stage(end.group(1).strip(), "done", float(end.group(2)))
            return
        failure = _STAGE_FAIL.match(text)
        if failure:
            self._emit_stage(failure.group(1).strip(), "failed", error=failure.group(2).strip())


class AnalysisWorker(_StageMixin, QThread):
    completed = Signal(object)
    failed = Signal(str)
    confirm_needed = Signal(object)

    def __init__(
        self,
        settings: Settings,
        source: Path,
        *,
        request: str = "",
        model: str | None = None,
        use_ai: bool = True,
        task: str = "general_analysis",
        parent: Any = None,
    ) -> None:
        super().__init__(parent)
        self.settings = settings
        self.source = Path(source)
        self.request = request
        self.model = model
        self.use_ai = use_ai
        self.task = task
        self._coordinator: Any = None
        self._cancel_requested = False
        self._confirm_event: threading.Event | None = None
        self._confirm_value = False

    # ------------------------------------------------------------------ #
    def run(self) -> None:  # noqa: D102 - QThread entry point
        try:
            from local_office_ai.app.coordinator import Coordinator

            coordinator = Coordinator(self.settings, progress=self._on_message, console=False)
            self._coordinator = coordinator
            if self._cancel_requested:
                coordinator.abort()
            result = coordinator.run(
                self.source,
                model=self.model,
                use_ai=self.use_ai,
                task=self.task,
                request=self.request or None,
                confirm_structure=self._confirm,
            )
            self.completed.emit(result)
        except Exception as exc:  # noqa: BLE001 - the operator must see the reason
            self.failed.emit(friendly_error(exc))
        finally:
            self._coordinator = None

    def _confirm(self, analysis: Any) -> bool:
        if analysis.primary_table is None or analysis.structure_confidence >= 0.6:
            return True
        self._confirm_event = threading.Event()
        self._confirm_value = False
        self.confirm_needed.emit(analysis)
        self._confirm_event.wait(120)
        return self._confirm_value

    def resolve_confirm(self, value: bool) -> None:
        self._confirm_value = bool(value)
        if self._confirm_event is not None:
            self._confirm_event.set()

    def cancel(self) -> None:
        self._cancel_requested = True
        coordinator = self._coordinator
        if coordinator is not None:
            coordinator.abort()


class FolderWorker(_StageMixin, QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        settings: Settings,
        folder: Path,
        *,
        request: str = "",
        model: str | None = None,
        use_ai: bool = True,
        selected_files: list[Path] | None = None,
        recursive: bool = True,
        parent: Any = None,
    ) -> None:
        super().__init__(parent)
        self.settings = settings
        self.folder = Path(folder)
        self.request = request
        self.model = model
        self.use_ai = use_ai
        self.selected_files = list(selected_files) if selected_files is not None else None
        self.recursive = bool(recursive)
        self._abort_event = threading.Event()

    def run(self) -> None:  # noqa: D102 - QThread entry point
        try:
            from local_office_ai.app.document_set_workflow import DocumentSetWorkflow
            from local_office_ai.ingestion.reader import EXCEL_SUFFIXES

            selected = self.selected_files
            if selected is None:
                selected = DocumentSetWorkflow.scan(self.folder, recursive=self.recursive)
            if selected and any(Path(item).suffix.lower() not in EXCEL_SUFFIXES for item in selected):
                workflow = DocumentSetWorkflow(self.settings, progress=self._on_message, abort_event=self._abort_event)
                result = workflow.run(
                    paths=selected, folder=self.folder, recursive=self.recursive,
                    request=self.request or None, model=self.model, use_ai=self.use_ai,
                )
            else:
                from local_office_ai.app.folder_workflow import FolderWorkflow
                workflow = FolderWorkflow(
                    self.settings, progress=self._on_message, abort_event=self._abort_event,
                )
                result = workflow.run(
                    self.folder, request=self.request or None, model=self.model,
                    use_ai=self.use_ai, selected=selected,
                )
            self.completed.emit(result)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(friendly_error(exc))

    def cancel(self) -> None:
        self._abort_event.set()


class FolderScanWorker(QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, settings: Settings, folder: Path, *, recursive: bool = True, parent: Any = None) -> None:
        super().__init__(parent)
        self.settings = settings
        self.folder = Path(folder)
        self.recursive = bool(recursive)

    def run(self) -> None:  # noqa: D102
        try:
            from local_office_ai.app.document_set_workflow import DocumentSetWorkflow

            files = DocumentSetWorkflow.scan(self.folder, recursive=self.recursive)
            self.completed.emit(
                [
                    {
                        "path": str(path),
                        "name": path.name,
                        "size": path.stat().st_size if path.exists() else 0,
                    }
                    for path in files
                ]
            )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(friendly_error(exc))


class OllamaWorker(QThread):
    completed = Signal(bool, object, str)

    def __init__(self, settings: Settings, parent: Any = None) -> None:
        super().__init__(parent)
        self.settings = settings

    def run(self) -> None:  # noqa: D102
        try:
            from local_office_ai.ai.ollama_client import OllamaClient

            models = OllamaClient(self.settings).list_models()
            payload = [
                {
                    "name": model.name,
                    "size_gb": model.size_gb,
                    "family": model.family,
                    "parameter_size": model.parameter_size,
                }
                for model in models
            ]
            self.completed.emit(True, payload, "")
        except AIServiceError as exc:
            self.completed.emit(False, [], str(exc))
        except Exception as exc:  # noqa: BLE001
            self.completed.emit(False, [], friendly_error(exc))


class ChatWorker(_StageMixin, QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        settings: Settings,
        question: str,
        *,
        analysis: Any = None,
        source: Path | None = None,
        document_set_result: Any = None,
        model: str | None = None,
        use_ai: bool = True,
        parent: Any = None,
    ) -> None:
        super().__init__(parent)
        self.settings = settings
        self.question = question
        self.analysis = analysis
        self.source = Path(source) if source else None
        self.document_set_result = document_set_result
        self.model = model
        self.use_ai = use_ai
        self._coordinator: Any = None

    def run(self) -> None:  # noqa: D102
        try:
            if self.document_set_result is not None:
                reply = self._ask_document_set_followup()
            else:
                reply = self._ask_without_data() if self.analysis is None else self._ask_followup()
            self.completed.emit(reply)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(friendly_error(exc))
        finally:
            self._coordinator = None

    # ------------------------------------------------------------------ #
    def _ask_without_data(self) -> ChatReply:
        if self.source is None:
            raise EngineError("برای گفت‌وگو ابتدا یک فایل Excel انتخاب و تحلیل کنید.")
        self._emit_stage("دریافت فایل", "running")
        from local_office_ai.app.coordinator import Coordinator

        coordinator = Coordinator(self.settings, progress=self._on_message, console=False)
        self._coordinator = coordinator
        result = coordinator.run(
            self.source,
            model=self.model,
            use_ai=self.use_ai,
            request=self.question,
        )
        analysis = result.analysis
        return ChatReply(
            analysis=analysis,
            answers=list(getattr(analysis, "request_answers", []) or []),
            ai=dict(getattr(analysis, "ai", {}) or {}),
            source=str(self.source),
            output_path=str(result.output_path),
            followup=False,
            result=result,
        )

    def _ask_document_set_followup(self) -> ChatReply:
        """Answer from an already-computed mixed-format result; never recalculate in the LLM."""
        from local_office_ai.ai.ollama_client import OllamaClient

        result = self.document_set_result
        if not self.use_ai or not self.model:
            raise AIServiceError("برای پرسش تکمیلی روی چند منبع، اتصال مدل محلی باید فعال باشد.")
        files = list(getattr(result, "files", []) or [])
        lines = [
            f"درخواست اولیه: {getattr(result, 'request', '') or 'تحلیل عمومی منابع'}",
            f"خلاصه قبلی: {getattr(result, 'ai_summary', '') or getattr(result, 'fallback_summary', '')}",
            "منابع:",
        ]
        for item in files[:40]:
            lines.append(
                f"- {getattr(item, 'name', '')}: {getattr(item, 'status', '')}; "
                f"{getattr(item, 'summary', '')}; خطا/هشدار: {getattr(item, 'error', '')}"
            )
        lines.append("شاخص‌های محاسبه‌شده توسط Python:")
        for item in list(getattr(result, "statistics", []) or [])[:60]:
            lines.append(
                f"- {item.get('source', '')} / {item.get('table', '')} / {item.get('column', '')}: "
                f"n={item.get('n')}, mean={item.get('mean')}, median={item.get('median')}, "
                f"min={item.get('min')}, max={item.get('max')}, sum={item.get('sum')}"
            )
        warnings = list(getattr(result, "warnings", []) or [])
        if warnings:
            lines.append("محدودیت‌ها و هشدارها:")
            lines.extend(f"- {item}" for item in warnings[:12])
        for source, location, text in list(getattr(result, "extracted_text", []) or [])[:20]:
            snippet = str(text)[:700]
            if snippet:
                lines.append(f"گزیده {source} / {location}: {snippet}")
        context = "\n".join(lines)[:6500]
        self._emit_stage("پرسش از نتایج چندمنبعی", "running")
        response = OllamaClient(self.settings).chat(
            self.model,
            [
                {
                    "role": "system",
                    "content": (
                        "تو دستیار اداری و آماری فارسی هستی. فقط بر اساس متن و ارقام موجود در زمینه پاسخ بده. "
                        "اعداد محاسبه‌شده توسط Python را تغییر نده و عدد تازه نساز. واقعیت، محاسبه و برداشت احتمالی را جدا کن. "
                        "اگر اطلاعات برای پاسخ کافی نیست، صریح بگو. پاسخ کوتاه، روشن و مستند باشد."
                    ),
                },
                {
                    "role": "user",
                    "content": f"زمینه تحلیل چندمنبعی قبلی:\n{context}\n\nپرسش تکمیلی کاربر:\n{self.question}",
                },
            ],
            json_format=False,
            temperature=0.1,
            num_ctx=self.settings.ollama_num_ctx,
        )
        answer = response.content.strip()
        if not answer:
            raise AIServiceError("پاسخ تکمیلی مدل خالی بود.")
        self._emit_stage("پرسش از نتایج چندمنبعی", "done")
        return ChatReply(
            analysis=None,
            ai={"mode": "model", "response": {"summary": answer, "is_grounded": True}},
            source=str(getattr(result, "folder", "منابع انتخاب‌شده")),
            output_path=str(getattr(result, "output_path", "")),
            followup=True,
        )

    def _ask_followup(self) -> ChatReply:
        from local_office_ai.ai.analyzer import AIAnalyzer
        from local_office_ai.ai.request_analyzer import RequestAnalyzer
        from local_office_ai.excel.data_qa import answer_request

        analysis = self.analysis
        integration = getattr(analysis, "integration", None)
        metric_names: list[str] = []
        if integration is not None:
            metric_names.extend(getattr(integration, "metric_names", []) or [])
        metric_names.extend(
            profile.header for profile in (getattr(analysis, "column_profiles", []) or [])
        )

        self._emit_stage("تفسیر درخواست", "running")
        intent = RequestAnalyzer().parse(
            self.question,
            available_metrics=list(dict.fromkeys(metric_names)),
        )
        self._emit_stage("تفسیر درخواست", "done")

        self._emit_stage("پاسخ از داده‌های محاسبه‌شده", "running")
        answers = answer_request(intent, analysis, integration)
        self._emit_stage("پاسخ از داده‌های محاسبه‌شده", "done")

        self._emit_stage("تحلیل مدل", "running")
        outcome = AIAnalyzer(self.settings).analyze(
            analysis,
            model=self.model if self.use_ai else None,
            question=self.question,
            use_ai=self.use_ai,
        )
        self._emit_stage("تحلیل مدل", "done")
        return ChatReply(
            analysis=analysis,
            answers=list(answers),
            ai=outcome.to_dict(),
            source=str(getattr(analysis.meta, "path", "") or ""),
            followup=True,
        )

    def cancel(self) -> None:
        coordinator = self._coordinator
        if coordinator is not None:
            coordinator.abort()


class DocumentSetWorker(_StageMixin, QThread):
    """Analyze one or more mixed-format sources in a single bounded workflow."""

    completed = Signal(object)
    failed = Signal(str)

    def __init__(
        self, settings: Settings, *, paths: list[Path] | None = None, folder: Path | None = None,
        recursive: bool = True, request: str = "", model: str | None = None, use_ai: bool = True, parent: Any = None,
    ) -> None:
        super().__init__(parent)
        self.settings = settings
        self.paths = list(paths) if paths is not None else None
        self.folder = Path(folder) if folder else None
        self.recursive = bool(recursive)
        self.request = request
        self.model = model
        self.use_ai = use_ai
        self._abort_event = threading.Event()

    def run(self) -> None:
        try:
            from local_office_ai.app.document_set_workflow import DocumentSetWorkflow
            workflow = DocumentSetWorkflow(self.settings, progress=self._on_message, abort_event=self._abort_event)
            result = workflow.run(
                paths=self.paths, folder=self.folder, recursive=self.recursive, request=self.request or None,
                model=self.model, use_ai=self.use_ai,
            )
            self.completed.emit(result)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(friendly_error(exc))

    def cancel(self) -> None:
        self._abort_event.set()


class HardwareWorker(QThread):
    snapshot = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        *,
        enable_gpu: bool = True,
        interval: float = 2.0,
        parent: Any = None,
    ) -> None:
        super().__init__(parent)
        self.enable_gpu = enable_gpu
        self.interval = max(1.0, float(interval))
        self._stop_event = threading.Event()

    def run(self) -> None:  # noqa: D102
        try:
            from local_office_ai.ui.desktop.hardware import HardwareMonitor

            monitor = HardwareMonitor(enable_gpu=self.enable_gpu)
            self._stop_event.wait(0.25)
            while not self._stop_event.is_set():
                started = __import__("time").monotonic()
                try:
                    self.snapshot.emit(monitor.sample())
                except Exception as exc:  # noqa: BLE001
                    self.failed.emit(friendly_error(exc))
                elapsed = __import__("time").monotonic() - started
                self._stop_event.wait(max(0.2, self.interval - elapsed))
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(friendly_error(exc))

    def stop(self) -> None:
        self._stop_event.set()


class PreviewWorker(QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, path: Path, *, max_rows: int = 60, max_cols: int = 14, parent: Any = None) -> None:
        super().__init__(parent)
        self.path = Path(path)
        self.max_rows = max_rows
        self.max_cols = max_cols

    def run(self) -> None:  # noqa: D102
        try:
            from local_office_ai.ui.desktop.preview import read_workbook_preview

            self.completed.emit(
                read_workbook_preview(self.path, max_rows=self.max_rows, max_cols=self.max_cols)
            )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(friendly_error(exc))


class ReportsScanWorker(QThread):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, output_dir: Path, parent: Any = None) -> None:
        super().__init__(parent)
        self.output_dir = Path(output_dir)

    def run(self) -> None:  # noqa: D102
        try:
            entries = []
            if self.output_dir.is_dir():
                for path in self.output_dir.glob("*.xlsx"):
                    try:
                        stat = path.stat()
                    except OSError:
                        continue
                    entries.append(
                        {
                            "path": str(path),
                            "name": path.name,
                            "modified": stat.st_mtime,
                            "size": stat.st_size,
                        }
                    )
            entries.sort(key=lambda item: item["modified"], reverse=True)
            self.completed.emit(entries)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(friendly_error(exc))
