"""Scrollable page scaffold used by every screen."""

from __future__ import annotations

from typing import Any

from PySide6.QtCore import QTimer, Qt
from PySide6.QtWidgets import QFrame, QScrollArea, QVBoxLayout, QWidget


class Page(QScrollArea):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("PageScroll")
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.content = QWidget()
        self.content.setObjectName("PageBody")
        self.layout = QVBoxLayout(self.content)
        self.layout.setContentsMargins(0, 0, 4, 10)
        self.layout.setSpacing(14)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.setWidget(self.content)

    def add(self, widget: QWidget, stretch: int = 0) -> QWidget:
        self.layout.addWidget(widget, stretch)
        return widget

    def add_layout(self, layout: Any, stretch: int = 0) -> None:
        self.layout.addLayout(layout, stretch)

    def add_stretch(self) -> None:
        self.layout.addStretch(1)

    def reveal(self, widget: QWidget) -> None:
        def scroll_to_widget() -> None:
            self.layout.activate()
            bar = self.verticalScrollBar()
            bar.setValue(max(0, widget.y() - 8))

        scroll_to_widget()
        QTimer.singleShot(120, scroll_to_widget)
