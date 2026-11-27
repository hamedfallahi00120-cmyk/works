"""File analysis page: source, request, live engine stages and results."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    BusyBar,
    Card,
    SectionHeader,
    StepList,
    make_label,
)
from local_office_ai.ui.desktop.formatting import elide, ltr
from local_office_ai.ui.desktop.pages.base import Page
from local_office_ai.ui.desktop.results import ResultsView
from local_office_ai.ui.desktop.source import RequestComposer, SourceListCard
from local_office_ai.ui.desktop.theme import Palette
from local_office_ai.ui.desktop.workers import StageEvent


class ProgressCard(Card):
    def __init__(self, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent, spacing=10)
        self.ai_enabled = False
        self.add(
            SectionHeader(
                "مراحل پردازش",
                "فقط مرحله‌هایی که موتور واقعاً گزارش می‌کند نمایش داده می‌شوند؛ درصد ساختگی وجود ندارد.",
            )
        )
        self.busy = BusyBar()
        self.steps = StepList()
        self.last_event = make_label("", "Caption", wrap=True)
        self.add(self.busy)
        self.add(self.steps)
        self.add(self.last_event)
        note = make_label(
            "بررسی کیفیت داده و محاسبات آماری داخل مرحله «تشخیص جدول و استخراج» اجرا می‌شوند "
            "و نتیجه‌شان در بخش جداول و هشدارها دیده می‌شود.",
            "Caption",
            wrap=True,
        )
        self.add(note)
        self.setVisible(False)

    def begin(self) -> None:
        self.steps.reset()
        self.last_event.setText("در انتظار شروع…")
        self.busy.start()
        self.setVisible(True)

    def _stage_name(self, name: str) -> str:
        if name == "تحلیل مدل" and not self.ai_enabled:
            return "تحلیل با قواعد قطعی"
        return name

    def on_stage(self, event: StageEvent) -> None:
        name = self._stage_name(event.name)
        if event.state == "running":
            self.steps.start(name)
            self.last_event.setText(f"در حال اجرا: {name}")
        elif event.state == "done":
            self.steps.finish(name, event.seconds)
            self.last_event.setText(f"پایان مرحله: {name}")
        elif event.state == "failed":
            self.steps.fail(name, event.error)
            self.last_event.setText(f"خطا در مرحله {name}: {event.error}")

    def on_progress(self, message: str) -> None:
        text = str(message).strip()
        if not text or "\n" in text or text.startswith("===="):
            return
        self.last_event.setText(elide(text, 180))

    def finish(self) -> None:
        self.busy.stop()


class ErrorCard(Card):
    def __init__(self, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._palette = palette
        head = QHBoxLayout()
        self.icon = QLabel()
        self.icon.setPixmap(icons.pixmap("alert", palette.danger, 22))
        title = QLabel("پردازش انجام نشد")
        title.setObjectName("CardTitle")
        head.addWidget(self.icon)
        head.addWidget(title, 1)
        self.add_layout(head)
        self.message = QLabel("")
        self.message.setObjectName("ErrorText")
        self.message.setWordWrap(True)
        self.message.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.add(self.message)
        self.add(
            make_label(
                "اگر خطا تکرار شد، مسیر فایل و متن خطا را بررسی کنید؛ گزارش کامل در پوشه logs است.",
                "Caption",
                wrap=True,
            )
        )
        self.setVisible(False)

    def show_error(self, message: str) -> None:
        self.message.setText(message)
        self.setVisible(True)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.icon.setPixmap(icons.pixmap("alert", palette.danger, 22))


class FileAnalysisPage(Page):
    browse_requested = Signal()
    add_folder_requested = Signal()
    remove_requested = Signal(str)
    analyze_requested = Signal(str)
    stop_requested = Signal()
    followup_requested = Signal()
    file_dropped = Signal(str)
    files_dropped = Signal(list)

    def __init__(self, state: Any, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self._palette = palette
        self.source_card = SourceListCard(palette)
        self.source_card.browse_requested.connect(self.browse_requested.emit)
        self.source_card.browse_folder_requested.connect(self.add_folder_requested.emit)
        self.source_card.remove_requested.connect(self.remove_requested.emit)
        self.source_card.file_dropped.connect(self.file_dropped.emit)
        self.source_card.files_dropped.connect(self.files_dropped.emit)
        self.add(self.source_card)

        self.composer = RequestComposer(palette)
        self.composer.start_requested.connect(self.analyze_requested.emit)
        self.composer.stop_requested.connect(self.stop_requested.emit)
        self.add(self.composer)

        self.progress_card = ProgressCard(palette)
        self.add(self.progress_card)

        self.error_card = ErrorCard(palette)
        self.add(self.error_card)

        self.results = ResultsView()
        self.results.apply_palette(palette)
        self.results.followup_requested.connect(self.followup_requested.emit)
        self.add(self.results)

        state.source_changed.connect(self._refresh)
        self._refresh()

    def _output_note(self) -> str:
        settings = getattr(self.state, "settings", None)
        output = getattr(settings, "output_dir", None)
        root = getattr(settings, "project_root", None)
        if output is None:
            return ""
        try:
            shown = Path(output).relative_to(root) if root is not None else Path(output)
        except ValueError:
            shown = Path(output)
        return f"خروجی: {ltr(str(shown).replace(chr(92), '/'))} • فایل‌های ورودی تغییر نمی‌کنند."

    def _apply_output_note(self) -> None:
        text = self._output_note()
        if not text:
            return
        settings = self.state.settings
        self.composer.set_output_note(text, str(settings.output_dir))

    # ------------------------------------------------------------------ #
    def _refresh(self) -> None:
        self.source_card.set_sources(
            list(getattr(self.state, "sources", []) or []),
            getattr(self.state, "folder", None),
        )
        self.results.setVisible(self._result_matches_source())
        self._apply_output_note()

    def _result_matches_source(self) -> bool:
        result = self.state.last_result
        source = self.state.source
        if result is None or source is None:
            return False
        analysis = getattr(result, "analysis", None)
        meta = getattr(analysis, "meta", None)
        path = getattr(meta, "path", "")
        if not path:
            return False
        try:
            return Path(str(path)).resolve() == source
        except OSError:
            return False

    def set_source(self, path: Path | None) -> None:
        self.source_card.set_file(path)

    def set_sources(self, paths: list[Path]) -> None:
        self.source_card.set_files(paths)

    def set_request(self, text: str) -> None:
        self.composer.set_request(text)

    def begin_run(self) -> None:
        self.error_card.setVisible(False)
        self.results.setVisible(False)
        self.progress_card.ai_enabled = bool(self.state.use_ai)
        self.composer.set_running(True)
        self.progress_card.begin()

    def on_stage(self, event: StageEvent) -> None:
        self.progress_card.on_stage(event)

    def on_progress(self, message: str) -> None:
        self.progress_card.on_progress(message)

    def show_result(self, result: Any) -> None:
        self.error_card.setVisible(False)
        self.results.apply_palette(self._palette)
        self.results.set_result(result, self._palette)

    def finish_success(self, result: Any) -> None:
        self.composer.input.clearFocus()
        self.composer.start_button.clearFocus()
        self.set_running(False)
        self.setFocus()
        self.progress_card.finish()
        self.progress_card.setVisible(False)
        self.show_result(result)
        self.composer.set_hint("تحلیل کامل شد؛ نتیجه در پایین صفحه است.", "success")

    def finish_error(self, message: str) -> None:
        self.composer.set_running(False)
        self.progress_card.finish()
        self.error_card.show_error(message)
        self.composer.set_hint("پردازش انجام نشد.", "danger")

    def set_running(self, running: bool) -> None:
        self.composer.set_running(running)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.source_card.apply_palette(palette)
        self.composer.apply_palette(palette)
        self.error_card.apply_palette(palette)
        self.results.apply_palette(palette)
