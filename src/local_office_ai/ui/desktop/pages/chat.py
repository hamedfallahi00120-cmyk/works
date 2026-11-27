"""Persian chat surface grounded in the analysis the engine already produced."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from local_office_ai.ui.desktop import icons
from local_office_ai.ui.desktop.components import (
    Badge,
    BusyBar,
    CallbackButton,
    Card,
    ChatInput,
    Chip,
    IconButton,
    PromptChips,
    make_label,
)
from local_office_ai.ui.desktop.formatting import (
    elide,
    format_number,
    ltr,
    mode_label,
    soften_numbers,
)
from local_office_ai.ui.desktop.preview import ReportPreviewDialog, open_path
from local_office_ai.ui.desktop.results import rich_text
from local_office_ai.ui.desktop.theme import Palette
from local_office_ai.ui.desktop.workers import ChatReply

GREETING = (
    "سلام. پس از تحلیل فایل‌ها یا پوشه، می‌توانید درباره همان داده پرسش تکمیلی بپرسید. "
    "اعداد با Python محاسبه می‌شوند و مدل محلی آن‌ها را تفسیر می‌کند."
)
FOLLOWUP_PROMPTS = (
    "چرا این نتیجه مهم است؟",
    "کدام ماه بیشترین تغییر را داشت؟",
    "برای مدیریت چه پیشنهادی داری؟",
)


class ChatBubble(QFrame):
    def __init__(
        self,
        role: str,
        title: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.role = role
        self.setObjectName("ChatError" if role == "error" else "ChatBubble")
        if role != "error":
            self.setProperty("role", role)
        self.setMinimumWidth(320)
        self.setMaximumWidth(700)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Minimum)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)
        head = QHBoxLayout()
        self.title = QLabel(title)
        self.title.setObjectName("BubbleAuthor")
        head.addWidget(self.title, 1)
        self.copy_button = IconButton("copy", "کپی این پاسخ", size=26, icon_size=14)
        self.copy_button.clicked.connect(self._copy)
        head.addWidget(self.copy_button)
        layout.addLayout(head)
        self.content = QVBoxLayout()
        self.content.setContentsMargins(0, 0, 0, 0)
        self.content.setSpacing(6)
        layout.addLayout(self.content)
        self._text = ""

    def add_text(self, text: str, object_name: str = "") -> QLabel:
        text = soften_numbers(text)
        label = QLabel(rich_text(text))
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        if object_name:
            label.setObjectName(object_name)
        self._text += text + "\n"
        self.content.addWidget(label)
        return label

    def add_widget(self, widget: QWidget) -> None:
        self.content.addWidget(widget)

    def _copy(self) -> None:
        QApplication.clipboard().setText(self._text.strip())
        self.copy_button.setToolTip("کپی شد")

    def apply_palette(self, palette: Palette) -> None:
        self.copy_button.apply_palette(palette)


class ChatPage(QWidget):
    ask_requested = Signal(str)
    attach_requested = Signal()

    def __init__(self, state: Any, palette: Palette, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self._palette = palette
        self._analysis: Any = None
        self._document_set_result: Any = None
        self._source: Path | None = None
        self._bubbles: list[ChatBubble] = []
        self._pending: ChatBubble | None = None
        self._build()
        state.analysis_changed.connect(self._refresh_chip)
        state.model_changed.connect(self._refresh_chip)
        state.source_changed.connect(self._source_changed)
        self._refresh_chip()

    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        header = Card(margins=(18, 14, 18, 14), spacing=8)
        head_row = QHBoxLayout()
        self.data_chip = Chip("داده‌ای متصل نیست", "muted")
        self.data_chip.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        self.mode_chip = Chip("", "muted")
        self.attach_button = CallbackButton(
            "تحلیل فایل برای گفت‌وگو", self.attach_requested.emit, object_name="Ghost", icon_name="file"
        )
        self.clear_button = CallbackButton(
            "پاک کردن گفت‌وگو", self.clear_conversation, object_name="Quiet", icon_name="trash"
        )
        head_row.addWidget(self.data_chip)
        head_row.addWidget(self.mode_chip)
        head_row.addStretch(1)
        head_row.addWidget(self.attach_button)
        head_row.addWidget(self.clear_button)
        header.add_layout(head_row)
        self.data_note = make_label(
            "پاسخ‌ها روی همان داده متصل ساخته می‌شوند؛ با تغییر فایل، گفت‌وگو بازنشانی می‌شود.",
            "Caption",
            wrap=True,
        )
        header.add(self.data_note)
        layout.addWidget(header)

        self.scroll = QScrollArea()
        self.scroll.setObjectName("ChatScroll")
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        chat_body = QWidget()
        chat_body.setObjectName("PageBody")
        self.chat_layout = QVBoxLayout(chat_body)
        self.chat_layout.setContentsMargins(14, 14, 14, 14)
        self.chat_layout.setSpacing(10)
        self.chat_layout.addStretch(1)
        self.scroll.setWidget(chat_body)
        self.scroll.setMinimumHeight(260)
        layout.addWidget(self.scroll, 1)

        input_card = Card(margins=(14, 12, 14, 12), spacing=8)
        self.prompts = PromptChips(FOLLOWUP_PROMPTS)
        self.prompts.chosen.connect(self._fill_prompt)
        input_card.add(self.prompts)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.input = ChatInput()
        self.input.submitted.connect(self._send)
        self.send_button = CallbackButton("ارسال", self._send, object_name="Primary", icon_name="send")
        row.addWidget(self.input, 1)
        row.addWidget(self.send_button)
        input_card.add_layout(row)
        status_row = QHBoxLayout()
        self.busy = BusyBar()
        self.status = make_label("آماده", "Caption")
        hint = make_label("ارسال با Enter", "Caption")
        status_row.addWidget(hint)
        status_row.addWidget(self.busy, 1)
        status_row.addWidget(self.status)
        input_card.add_layout(status_row)
        layout.addWidget(input_card)

        self._add_greeting()

    # ------------------------------------------------------------------ #
    def _add_greeting(self) -> None:
        self._add_bubble("assistant", "دستیار", GREETING)

    def _add_bubble(self, role: str, title: str, text: str = "") -> ChatBubble:
        if role == "error":
            bubble = ChatBubble("error", title)
        else:
            bubble = ChatBubble(role, title)
        bubble.apply_palette(self._palette)
        if text:
            bubble.add_text(text)
        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        if role == "user":
            row.addStretch(1)
            row.addWidget(bubble)
        else:
            row.addWidget(bubble, 1)
            row.addStretch(0)
            bubble.setMinimumWidth(360)
            bubble.setMaximumWidth(720)
        self.chat_layout.insertWidget(self.chat_layout.count() - 1, container)
        self._bubbles.append(bubble)
        self._scroll_to_end()
        return bubble

    def _scroll_to_end(self) -> None:
        def go() -> None:
            bar = self.scroll.verticalScrollBar()
            bar.setValue(bar.maximum())

        QTimer.singleShot(0, go)
        QTimer.singleShot(90, go)

    # ------------------------------------------------------------------ #
    def attach_dataset(self, analysis: Any, source: Path | None, *, reset: bool = True) -> None:
        changed = analysis is not self._analysis or (source is not None and source != self._source)
        self._document_set_result = None
        self._analysis = analysis
        self._source = source
        self._refresh_chip()
        if analysis is None:
            return
        if changed and reset:
            self.clear_conversation()
        if changed:
            generic = getattr(analysis, "generic", None)
            rows = getattr(generic, "row_count", None) if generic is not None else None
            primary = getattr(analysis, "primary_table", None)
            details = []
            if primary is not None:
                details.append(f"جدول اصلی: {primary.sheet}!{primary.address}")
            if rows is not None:
                details.append(f"{format_number(rows)} ردیف")
            self._add_bubble(
                "assistant",
                "داده متصل شد",
                f"{source.name if source else 'فایل'} آماده است. "
                + " • ".join(details)
                + " — می‌توانید پرسش تکمیلی بنویسید.",
            )

    def attach_document_set(self, result: Any, *, reset: bool = True) -> None:
        """Attach a combined result so follow-up questions can reference every input."""
        changed = result is not self._document_set_result
        self._analysis = None
        self._source = None
        self._document_set_result = result
        self._refresh_chip()
        if changed and reset:
            self.clear_conversation()
        if changed:
            files = list(getattr(result, "files", []) or [])
            names = "، ".join(str(getattr(item, "name", "")) for item in files[:4] if getattr(item, "name", ""))
            if len(files) > 4:
                names += f" و {len(files) - 4} فایل دیگر"
            self._add_bubble(
                "assistant",
                "منابع متصل شدند",
                f"نتیجه ترکیبی {len(files)} منبع به گفت‌وگو متصل شد. {names} — می‌توانید درباره همه منابع پرسش تکمیلی بپرسید.",
            )

    def clear_conversation(self) -> None:
        while self.chat_layout.count() > 1:
            item = self.chat_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self._bubbles.clear()
        self._pending = None
        self._add_greeting()

    def _source_changed(self, *_: Any) -> None:
        current = self.state.source
        if self._source is not None and (current is None or current != self._source):
            self._analysis = None
            self._source = None
            self.clear_conversation()
        self._refresh_chip()

    def _refresh_chip(self, *_: Any) -> None:
        if self._document_set_result is not None:
            files = list(getattr(self._document_set_result, "files", []) or [])
            self.data_chip.set_status(f"منابع متصل: {len(files)} فایل", "accent")
        elif self._analysis is not None:
            name = self._source.name if self._source else Path(str(getattr(self._analysis.meta, "path", ""))).name
            generic = getattr(self._analysis, "generic", None)
            rows = getattr(generic, "row_count", 0) if generic is not None else 0
            self.data_chip.set_status(f"داده متصل: {ltr(elide(name, 34))} • {format_number(rows)} ردیف", "accent")
        elif self.state.source is not None:
            self.data_chip.set_status(f"فایل انتخاب‌شده: {ltr(self.state.source.name)}", "info")
        else:
            self.data_chip.set_status("داده‌ای متصل نیست", "muted")
        if self.state.use_ai:
            self.mode_chip.set_status(f"مدل: {ltr(self.state.model)}", "accent")
        else:
            self.mode_chip.set_status("قواعد قطعی Python", "muted")

    # ------------------------------------------------------------------ #
    def session_analysis(self) -> Any:
        return self._analysis

    def session_document_set(self) -> Any:
        return self._document_set_result

    def ask(self, text: str) -> None:
        self.input.setPlainText(text)
        self._send()

    def _fill_prompt(self, prompt: str) -> None:
        self.input.setPlainText(prompt)
        self.input.setFocus()

    def _send(self) -> None:
        text = self.input.toPlainText().strip()
        if not text:
            return
        if self.state.busy:
            self.status.setText("پردازش دیگری در جریان است؛ کمی بعد تلاش کنید.")
            return
        self.input.clear()
        self._add_bubble("user", "شما", text)
        self._pending = self._add_bubble("assistant", "دستیار", "در حال بررسی داده‌ها…")
        self.set_busy(True)
        self.ask_requested.emit(text)

    def set_busy(self, busy: bool) -> None:
        self.busy.setVisible(busy)
        if busy:
            self.busy.start()
            self.status.setText("در حال پاسخ‌گویی…")
        else:
            self.busy.stop()
            self.status.setText("آماده")
        self.input.setEnabled(not busy)
        self.send_button.setEnabled(not busy)

    def _remove_pending(self) -> None:
        if self._pending is None:
            return
        bubble = self._pending
        container = bubble.parentWidget()
        self._pending = None
        if bubble in self._bubbles:
            self._bubbles.remove(bubble)
        if container is not None:
            container.deleteLater()

    def show_reply(self, reply: ChatReply) -> None:
        self._remove_pending()
        self.set_busy(False)
        if reply.analysis is not None:
            self.attach_dataset(
                reply.analysis,
                Path(reply.source) if reply.source else self._source,
                reset=False,
            )
        bubble = self._add_bubble("assistant", "دستیار")
        ai = reply.ai or {}
        response = ai.get("response") or {}
        summary = response.get("summary") or ""
        if summary:
            bubble.add_text(summary)
        findings = list(response.get("findings") or [])
        if findings:
            bubble.add_text("یافته‌های کلیدی", "Caption")
            for item in findings[:6]:
                bubble.add_text("• " + str(item))
        answers = list(reply.answers or [])
        if answers:
            bubble.add_text("پاسخ‌های محاسبه‌شده", "Caption")
            for answer in answers[:8]:
                title = getattr(answer, "title", "")
                value = getattr(answer, "value", "")
                text = getattr(answer, "text", "")
                line = f"• {title}" + (f" — {value}" if value else "") + (f": {text}" if text else "")
                bubble.add_text(elide(line, 420))
        if not summary and not answers and not findings:
            bubble.add_text("پاسخی از داده‌ها پیدا نشد؛ درخواست را دقیق‌تر بنویسید یا فایل را تحلیل کنید.")

        footer = QHBoxLayout()
        footer.setSpacing(6)
        footer.addWidget(Badge(mode_label(ai.get("mode")), "accent" if ai.get("mode") == "model" else "muted"))
        grounded = response.get("is_grounded")
        if grounded is False:
            footer.addWidget(Badge("اعداد بدون منبع دارد", "warning"))
        if reply.output_path:
            open_button = CallbackButton(
                "بازکردن گزارش", lambda path=reply.output_path: open_path(path), object_name="Quiet", icon_name="external"
            )
            open_button.apply_palette(self._palette)
            footer.addWidget(open_button)
        footer.addStretch(1)
        holder = QWidget()
        holder.setLayout(footer)
        bubble.add_widget(holder)
        self._scroll_to_end()

    def show_stage(self, event: Any) -> None:
        if event.state == "running":
            self.status.setText(f"در حال اجرا: {event.name}")
        elif event.state == "done":
            self.status.setText(f"پایان: {event.name}")

    def show_error(self, message: str) -> None:
        self._remove_pending()
        self.set_busy(False)
        bubble = self._add_bubble("error", "پاسخ دریافت نشد", message)
        bubble.add_text("اگر ارتباط Ollama قطع است، تحلیل با قواعد قطعی Python انجام می‌شود.")
        self._scroll_to_end()

    def set_analysis(self, analysis: Any, source: Path | None) -> None:
        self.attach_dataset(analysis, source)

    def apply_palette(self, palette: Palette) -> None:
        self._palette = palette
        self.attach_button.apply_palette(palette)
        self.clear_button.apply_palette(palette)
        self.send_button.apply_palette(palette)
        for bubble in list(self._bubbles):
            try:
                bubble.apply_palette(palette)
            except RuntimeError:
                continue
