"""Landing page: pick a source, write the Persian request, start the analysis."""

from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from local_office_ai.ui.desktop.components import (
    Badge,
    CallbackButton,
    Card,
    SectionHeader,
    make_label,
)
from local_office_ai.ui.desktop.formatting import elide, ltr, rtl
from local_office_ai.ui.desktop.pages.base import Page
from local_office_ai.ui.desktop.source import RequestComposer, SourceListCard
from local_office_ai.ui.desktop.theme import Palette


class HomePage(Page):
    browse_file = Signal()
    browse_folder = Signal()
    analyze_requested = Signal(str)
    navigate = Signal(str)
    open_last = Signal()
    remove_source = Signal(str)
    file_dropped = Signal(str)
    files_dropped = Signal(list)

    def __init__(self, state: Any, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self._palette = palette
        self._build()
        state.source_changed.connect(self._refresh_source)
        state.folder_changed.connect(self._refresh_source)
        state.analysis_changed.connect(self._refresh_last)
        state.model_changed.connect(self._refresh_status)
        state.ollama_changed.connect(self._refresh_status)
        state.folder_result_changed.connect(self._refresh_last)
        self._refresh_source()
        self._refresh_last()

    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        hero = Card(variant="hero", margins=(26, 22, 26, 22), spacing=10, shadow=True)
        title = QLabel("دستیار هوشمند اداری")
        title.setObjectName("HeroTitle")
        self.hero_title = title
        self.hero_subtitle = QLabel(
            "یک یا چند فایل اداری انتخاب کنید، درخواست خود را فارسی بنویسید و گزارش مدیریتی بگیرید."
        )
        self.hero_subtitle.setObjectName("HeroSubtitle")
        self.hero_subtitle.setWordWrap(True)
        hero.add(title)
        hero.add(self.hero_subtitle)

        self.status_note = make_label("", "Caption", wrap=True)
        hero.add(self.status_note)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        file_button = CallbackButton(
            "انتخاب فایل / چند فایل", self.browse_file.emit, object_name="Primary", icon_name="file"
        )
        folder_button = CallbackButton(
            "انتخاب پوشه", self.browse_folder.emit, object_name="Ghost", icon_name="folder"
        )
        actions.addWidget(file_button)
        actions.addWidget(folder_button)
        actions.addStretch(1)
        hero.add_layout(actions)
        self.add(hero)

        self.sources = SourceListCard(self._palette)
        self.sources.browse_requested.connect(self.browse_file.emit)
        self.sources.browse_folder_requested.connect(self.browse_folder.emit)
        self.sources.remove_requested.connect(self.remove_source.emit)
        self.sources.file_dropped.connect(self.file_dropped.emit)
        self.sources.files_dropped.connect(self.files_dropped.emit)
        self.add(self.sources)

        self.composer = RequestComposer(self._palette, compact=True)
        self.composer.start_requested.connect(self.analyze_requested.emit)
        self.add(self.composer)

        last_card = Card()
        head = QHBoxLayout()
        head.setSpacing(10)
        head.addWidget(SectionHeader("آخرین نتیجه", "جدیدترین تحلیل همین نشست"))
        self.last_chip = Badge("بدون نتیجه", "muted")
        head.addWidget(self.last_chip, 0, Qt.AlignmentFlag.AlignBottom)
        head.addStretch(1)
        self.open_last_button = CallbackButton(
            "مشاهده نتیجه", self.open_last.emit, object_name="Ghost", icon_name="eye"
        )
        head.addWidget(self.open_last_button, 0, Qt.AlignmentFlag.AlignBottom)
        last_card.add_layout(head)
        self.last_summary = QLabel("هنوز تحلیلی در این نشست انجام نشده است.")
        self.last_summary.setObjectName("Muted")
        self.last_summary.setWordWrap(True)
        last_card.add(self.last_summary)
        self.last_path = QLabel("")
        self.last_path.setObjectName("Caption")
        self.last_path.setWordWrap(True)
        self.last_path.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.last_path.setAlignment(Qt.AlignmentFlag.AlignLeft)
        last_card.add(self.last_path)
        self.add(last_card)

    def set_compact(self, compact: bool) -> None:
        """Drop the marketing line on short screens so the request box stays reachable."""
        self.hero_subtitle.setVisible(not compact)

    # ------------------------------------------------------------------ #
    def _refresh_source(self, *_: Any) -> None:
        files = list(getattr(self.state, "sources", []) or [])
        folder = getattr(self.state, "folder", None)
        self.sources.set_sources(files, folder)
        if len(files) > 1:
            self.composer.set_hint("چند منبع انتخاب شده‌اند؛ درخواست را بنویسید و شروع کنید.", "muted")
        elif len(files) == 1:
            self.composer.set_hint("منبع انتخاب شده است؛ درخواست را بنویسید و شروع کنید.", "muted")
        elif folder is not None:
            self.composer.set_hint(
                "پوشه انتخاب شده است؛ فایل‌های تیک‌خورده آن تحلیل و مقایسه می‌شوند.", "muted"
            )
        else:
            self.composer.set_hint("ابتدا یک فایل یا پوشه انتخاب کنید.", "warning")
        self._refresh_status()

    def _refresh_status(self, *_: Any) -> None:
        if self.state.use_ai:
            self.status_note.setText(
                rtl(
                    f"وضعیت موتور: Ollama متصل • مدل فعال {ltr(self.state.model)} • "
                    "اعداد با Python محاسبه می‌شوند."
                )
            )
            self.status_note.setObjectName("SuccessText")
        elif self.state.ollama_online:
            self.status_note.setText(
                rtl("وضعیت موتور: Ollama متصل است اما مدلی انتخاب نشده؛ تحلیل با قواعد قطعی Python انجام می‌شود.")
            )
            self.status_note.setObjectName("Caption")
        else:
            self.status_note.setText(
                rtl(
                    "وضعیت موتور: Ollama در دسترس نیست؛ تحلیل با قواعد قطعی Python انجام می‌شود "
                    "و همین در گزارش ذکر می‌شود."
                )
            )
            self.status_note.setObjectName("Caption")
        self.status_note.style().unpolish(self.status_note)
        self.status_note.style().polish(self.status_note)

    def _refresh_last(self, *_: Any) -> None:
        result = self.state.last_result
        folder_result = self.state.last_folder_result
        if self.state.last_result_kind == "folder" and folder_result is not None:
            self._show_folder_result(folder_result)
        elif result is not None:
            analysis = getattr(result, "analysis", None)
            ai = getattr(analysis, "ai", {}) if analysis is not None else {}
            response = ai.get("response") or {}
            summary = response.get("summary") or "نتیجه آماده است."
            self.last_summary.setText(elide(summary, 320))
            path = str(getattr(result, "output_path", ""))
            self._set_path(path)
            self.last_chip.set_status("آماده", "success")
            self.open_last_button.setEnabled(True)
        elif folder_result is not None:
            self._show_folder_result(folder_result)
        else:
            self.last_summary.setText("هنوز تحلیلی در این نشست انجام نشده است.")
            self.last_path.setText("")
            self.last_path.setToolTip("")
            self.last_chip.set_status("بدون نتیجه", "muted")
            self.open_last_button.setEnabled(False)
        self.composer.set_running(self.state.busy)

    def _show_folder_result(self, folder_result: Any) -> None:
        files = list(getattr(folder_result, "files", []) or [])
        self.last_summary.setText(
            f"تحلیل پوشه با {len(files)} فایل انجام شد. "
            f"{len(getattr(folder_result, 'metrics', []) or [])} شاخص مشترک پیدا شد."
        )
        self._set_path(str(getattr(folder_result, "output_path", "")))
        self.last_chip.set_status("آماده", "success")
        self.open_last_button.setEnabled(True)

    def _set_path(self, path: str) -> None:
        self.last_path.setText(("خروجی: " + elide(ltr(path), 96)) if path else "")
        self.last_path.setToolTip(path)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.composer.apply_palette(palette)
        self.sources.apply_palette(palette)
