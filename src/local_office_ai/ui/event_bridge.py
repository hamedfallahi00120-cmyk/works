"""Thread-safe bridge between worker threads and the Tk main loop.

Worker threads never call Tkinter. They post immutable events to a queue; the
main loop drains that queue inside ``root.after`` and is the only code that
touches widgets. Confirmation dialogs are request/response events: the worker
waits on an ``Event`` that the main loop sets after the operator answers.
"""

from __future__ import annotations

import queue
import threading
from dataclasses import dataclass, field
from typing import Any


@dataclass
class UiEvent:
    kind: str
    payload: Any = None
    token: int | None = None
    ack: threading.Event | None = None
    result: dict[str, Any] = field(default_factory=dict)


class UiBridge:
    def __init__(self) -> None:
        self.queue: "queue.Queue[UiEvent]" = queue.Queue()
        self._counter = 0
        self._counter_lock = threading.Lock()

    # -- worker side ---------------------------------------------------- #
    def log(self, message: str) -> None:
        self.queue.put(UiEvent("log", str(message)))

    def done(self, result: Any) -> None:
        self.queue.put(UiEvent("done", result))

    def error(self, message: str) -> None:
        self.queue.put(UiEvent("error", str(message)))

    def models(self, names: list[str]) -> None:
        self.queue.put(UiEvent("models", list(names)))

    def checkpoints(self, names: list[str]) -> None:
        self.queue.put(UiEvent("checkpoints", list(names)))

    def image_status(self, message: str) -> None:
        self.queue.put(UiEvent("image_status", str(message)))

    def confirm(self, text: str, timeout: float | None = None) -> bool:
        """Ask the main loop for a yes/no answer and block until it arrives."""
        with self._counter_lock:
            self._counter += 1
            token = self._counter
        event = UiEvent("confirm", text, token, threading.Event(), {"value": False})
        self.queue.put(event)
        event.ack.wait(timeout)  # type: ignore[union-attr]
        return bool(event.result.get("value"))

    # -- main-loop side ------------------------------------------------- #
    def drain(self, limit: int = 200) -> list[UiEvent]:
        events: list[UiEvent] = []
        for _ in range(limit):
            try:
                events.append(self.queue.get_nowait())
            except queue.Empty:
                break
        return events

    @staticmethod
    def resolve(event: UiEvent, value: Any) -> None:
        event.result["value"] = value
        if event.ack is not None:
            event.ack.set()
