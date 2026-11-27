"""The GUI bridge must let workers communicate without touching Tk widgets."""

from __future__ import annotations

import threading

from local_office_ai.ui.event_bridge import UiBridge


def test_worker_events_are_queued_for_the_main_loop():
    bridge = UiBridge()
    bridge.log("hello")
    bridge.models(["qwen3.5:9b"])
    bridge.done("result")
    bridge.error("boom")

    events = bridge.drain()
    assert [event.kind for event in events] == ["log", "models", "done", "error"]
    assert events[0].payload == "hello"
    assert events[1].payload == ["qwen3.5:9b"]


def test_confirmation_is_a_request_response_event():
    bridge = UiBridge()
    answers: dict[str, bool] = {}

    def worker() -> None:
        answers["value"] = bridge.confirm("continue?", timeout=3)

    thread = threading.Thread(target=worker)
    thread.start()

    events = []
    for _ in range(100):
        events = bridge.drain()
        if events:
            break
        threading.Event().wait(0.01)
    assert len(events) == 1
    assert events[0].kind == "confirm"
    assert events[0].payload == "continue?"
    bridge.resolve(events[0], True)

    thread.join(timeout=3)
    assert not thread.is_alive()
    assert answers["value"] is True


def test_confirmation_times_out_safely():
    bridge = UiBridge()
    assert bridge.confirm("no answer", timeout=0.05) is False


def test_wizard_worker_does_not_touch_widgets():
    """Static guard: the worker must use only the bridge, never Tk names."""
    import ast
    from pathlib import Path

    source = Path(__file__).resolve().parents[1] / "src" / "local_office_ai" / "ui" / "wizard.py"
    tree = ast.parse(source.read_text(encoding="utf-8"))
    launch = next(
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name == "launch"
    )
    functions = {
        node.name: node
        for node in launch.body
        if isinstance(node, ast.FunctionDef)
    }
    forbidden = {
        "root",
        "log_box",
        "progress",
        "run_button",
        "stop_button",
        "open_button",
        "model_var",
        "task_var",
        "question_var",
        "request_var",
        "file_var",
        "folder_button",
        "messagebox",
        "filedialog",
        "ttk",
    }
    for name in ("worker", "folder_worker", "confirm_structure"):
        used = {node.id for node in ast.walk(functions[name]) if isinstance(node, ast.Name)}
        assert not (used & forbidden), f"{name} touches Tk widgets: {used & forbidden}"
    pump = functions["pump_events"]
    calls = [
        node.func.attr
        for node in ast.walk(pump)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
    ]
    assert "after" in calls


def test_worker_forwards_the_persian_request_and_folder_mode_is_wired():
    """The request text must reach the workflow, and folder analysis must stay bridge-only."""
    import ast
    from pathlib import Path

    source = Path(__file__).resolve().parents[1] / "src" / "local_office_ai" / "ui" / "wizard.py"
    tree = ast.parse(source.read_text(encoding="utf-8"))
    launch = next(
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name == "launch"
    )
    functions = {
        node.name: node
        for node in launch.body
        if isinstance(node, ast.FunctionDef)
    }

    worker = functions["worker"]
    coord_calls = [
        node
        for node in ast.walk(worker)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "run"
    ]
    assert coord_calls, "worker باید coordinator.run را صدا بزند"
    keywords = {keyword.arg for keyword in coord_calls[0].keywords}
    assert "request" in keywords, "درخواست فارسی به workflow منتقل نمی‌شود"

    requests = [
        node
        for node in ast.walk(launch)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "Entry"
        and any(
            keyword.arg == "textvariable"
            and isinstance(keyword.value, ast.Name)
            and keyword.value.id == "request_var"
            for keyword in node.keywords
        )
    ]
    assert requests, "فیلد درخواست در فرم رابط گرافیکی ساخته نمی‌شود"

    folder = functions["folder_worker"]
    assert any(
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "done"
        for node in ast.walk(folder)
    ), "پردازش پوشه نتیجه را به bridge نمی‌دهد"
