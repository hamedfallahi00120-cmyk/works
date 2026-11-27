"""Shared fixtures.

No mocks: the Ollama tests run a real HTTP server in-process, and every workbook
used by the tests is a real .xlsx file created by openpyxl.
"""

from __future__ import annotations

import json
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from openpyxl import Workbook  # noqa: E402

from local_office_ai.core.settings import load_settings  # noqa: E402

SAMPLE = ROOT / "data" / "input" / "Book2_AI_V4_2_Monitored.xlsx"


@pytest.fixture()
def project_root(tmp_path: Path) -> Path:
    root = tmp_path / "project"
    (root / "config").mkdir(parents=True)
    (root / "data").mkdir()
    (root / "logs").mkdir()
    return root


@pytest.fixture()
def settings(project_root: Path):
    # Regression tests cover the full engine surface, so they run the audit
    # profile (all traceability sheets). Clean delivery has its own tests.
    loaded = load_settings(project_root=project_root)
    loaded.report_profile = "audit"
    return loaded


@pytest.fixture()
def sample_path() -> Path:
    assert SAMPLE.exists(), "reference workbook is missing"
    return SAMPLE


def build_workbook(path: Path, sheets: dict[str, list[list]], post=None) -> Path:
    """Create a workbook from ``{sheet_name: rows}``; ``post`` may style it."""
    workbook = Workbook()
    first = True
    for name, rows in sheets.items():
        if first:
            worksheet = workbook.active
            worksheet.title = name
            first = False
        else:
            worksheet = workbook.create_sheet(name)
        for row_index, row in enumerate(rows, start=1):
            for col_index, value in enumerate(row, start=1):
                if value is not None:
                    worksheet.cell(row_index, col_index, value)
    if post is not None:
        post(workbook)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    return path


@pytest.fixture()
def make_workbook(tmp_path: Path):
    def factory(name: str, sheets: dict[str, list[list]], post=None) -> Path:
        return build_workbook(tmp_path / name, sheets, post)

    return factory


@pytest.fixture()
def generic_workbook(make_workbook) -> Path:
    """Two sheets, two tables, Persian digits, Jalali dates, formulas, an error."""
    sheets = {
        "خلاصه": [
            ["گزارش هزینه‌های اداری"],
            [],
            ["شاخص", "مقدار", "توضیح"],
            ["تعداد اقلام", 4, "شمارش دستی"],
            ["جمع کل", "=SUM(هزینه‌ها!C3:C6)", "فرمول"],
        ],
        "هزینه‌ها": [
            ["جدول هزینه‌های جاری"],
            [],
            ["ردیف", "شرح", "مبلغ", "تاریخ", "دسته"],
            [1, "اجاره دفتر", "۱۲٬۵۰۰٬۰۰۰", "1404/01/15", "اجاره"],
            [2, "برق و آب", "1,250,000", "1404/02/03", "انرژی"],
            [3, "اینترنت", "۸۵۰٬۰۰۰", "1404/02/10", "ارتباطات"],
            [4, "اجاره دفتر", "۱۲٬۵۰۰٬۰۰۰", "1404/01/15", "اجاره"],
            ["جمع", None, "=SUM(C4:C7)", None, None],
            [],
            ["شرح", "مبلغ"],
            ["حق بیمه", "۲٬۰۰۰٬۰۰۰"],
            ["مالیات", "۱٬۵۰۰٬۰۰۰"],
        ],
    }

    def post(workbook):
        summary = workbook["خلاصه"]
        summary.merge_cells("A1:C1")
        costs = workbook["هزینه‌ها"]
        costs.merge_cells("A1:E1")
        costs.column_dimensions["E"].hidden = True
        costs["F4"] = "#DIV/0!"

    return make_workbook("generic.xlsx", sheets, post)


@pytest.fixture()
def loans_english_workbook(make_workbook) -> Path:
    """English headers, different column order, obligations on their own sheet."""
    sheets = {
        "Loans": [
            ["Loan portfolio"],
            ["Creditor", "Monthly payment", "Loan amount", "Installments", "Due date"],
            ["Bank A", 1200, 24000, 20, "2027-01-15"],
            ["Bank B", 800, 19200, 24, "2026-12-01"],
            ["Fund C", 450, 10800, 24, "2027-03-20"],
        ],
        "Commitments": [
            ["Monthly commitments"],
            ["Creditor", "Amount"],
            ["Bank A", 1200],
            ["Bank B", 850],
            ["Fund C", 450],
            ["Charity", 300],
        ],
    }
    return make_workbook("loans_en.xlsx", sheets)


@pytest.fixture()
def no_table_workbook(make_workbook) -> Path:
    return make_workbook("empty.xlsx", {"Sheet1": [["فقط یک یادداشت"]]})


# --------------------------------------------------------------------------- #
# Real HTTP server standing in for Ollama
# --------------------------------------------------------------------------- #
class _Handler(BaseHTTPRequestHandler):
    server: "FakeOllamaServer"

    def log_message(self, *args):  # silence request logging
        return

    def _send(self, status: int, payload: dict | str) -> None:
        body = payload if isinstance(payload, str) else json.dumps(payload, ensure_ascii=False)
        data = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):  # noqa: N802
        owner = self.server.owner
        owner.requests.append(("GET", self.path))
        if self.path == "/api/tags":
            if owner.fail_tags:
                self._send(500, {"error": "boom"})
                return
            self._send(200, {"models": owner.models})
            return
        self._send(404, {"error": "not found"})

    def do_POST(self):  # noqa: N802
        owner = self.server.owner
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        try:
            body = json.loads(raw)
        except json.JSONDecodeError:
            body = {}
        owner.requests.append(("POST", self.path))
        owner.last_body = body
        if self.path == "/api/show":
            self._send(200, {"details": {"family": "qwen"}})
            return
        if self.path != "/api/chat":
            self._send(404, {"error": "not found"})
            return
        if owner.chat_delay:
            time.sleep(owner.chat_delay)
        if owner.chat_status != 200:
            self._send(owner.chat_status, {"error": "model failure"})
            return
        if owner.chat_raw is not None:
            self._send(200, owner.chat_raw)
            return
        self._send(
            200,
            {
                "model": body.get("model", "qwen3.5:9b"),
                "message": {"role": "assistant", "content": owner.chat_content},
                "total_duration": 2_000_000_000,
                "load_duration": 500_000_000,
                "prompt_eval_count": 700,
                "prompt_eval_duration": 300_000_000,
                "eval_count": 120,
                "eval_duration": 1_000_000_000,
            },
        )


class FakeOllamaServer:
    def __init__(self) -> None:
        self.models = [
            {"name": "qwen3.5:9b", "size": 6_000_000_000, "details": {"family": "qwen3", "parameter_size": "9B"}},
            {"name": "llama3:8b", "size": 4_000_000_000, "details": {"family": "llama"}},
        ]
        self.chat_content = json.dumps(
            {
                "summary": "خلاصه آزمایشی",
                "findings": ["یافته یک"],
                "missing_data": [],
                "recommendations": ["پیشنهاد یک"],
                "risk_level": "low",
            },
            ensure_ascii=False,
        )
        self.chat_status = 200
        self.chat_raw: dict | None = None
        self.chat_delay = 0.0
        self.fail_tags = False
        self.requests: list[tuple[str, str]] = []
        self.last_body: dict | None = None
        self._httpd: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None

    def start(self) -> str:
        self._httpd = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
        self._httpd.owner = self  # type: ignore[attr-defined]
        self._thread = threading.Thread(target=self._httpd.serve_forever, daemon=True)
        self._thread.start()
        host, port = self._httpd.server_address[:2]
        return f"http://{host}:{port}"

    def stop(self) -> None:
        if self._httpd is not None:
            self._httpd.shutdown()
            self._httpd.server_close()
        if self._thread is not None:
            self._thread.join(timeout=3)


@pytest.fixture()
def fake_ollama():
    server = FakeOllamaServer()
    yield server
    server.stop()


@pytest.fixture()
def ollama_settings(settings, fake_ollama):
    settings.ollama_base_url = fake_ollama.start()
    settings.ollama_retries = 1
    settings.ollama_connect_timeout_seconds = 2.0
    settings.ollama_timeout_seconds = 5.0
    return settings


# --------------------------------------------------------------------------- #
# Real HTTP server standing in for ComfyUI
# --------------------------------------------------------------------------- #
PNG_BYTES = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01"
    b"\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
)


class _ComfyHandler(BaseHTTPRequestHandler):
    server: "FakeComfyServer"

    def log_message(self, *args):
        return

    def _send(self, status: int, body: bytes, content_type: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _json(self, payload: dict, status: int = 200) -> None:
        self._send(status, json.dumps(payload).encode("utf-8"), "application/json")

    def do_GET(self):  # noqa: N802
        owner = self.server.owner
        path = self.path.split("?")[0]
        owner.requests.append(("GET", self.path))
        if path == "/models/checkpoints":
            if owner.fail:
                self._json({"error": "boom"}, 500)
                return
            self._json(owner.checkpoints)
            return
        if path.startswith("/history/"):
            prompt_id = path.rsplit("/", 1)[-1]
            owner.history_calls += 1
            if owner.history_calls <= owner.empty_history_calls:
                self._json({})
                return
            self._json(
                {
                    prompt_id: {
                        "status": {"status_str": owner.status_str},
                        "outputs": {"7": {"images": [{"filename": "out.png", "subfolder": "", "type": "output"}]}},
                    }
                }
            )
            return
        if path == "/view":
            self._send(200, PNG_BYTES, "image/png")
            return
        self._json({"error": "not found"}, 404)

    def do_POST(self):  # noqa: N802
        owner = self.server.owner
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        owner.requests.append(("POST", self.path))
        owner.last_body = json.loads(raw) if raw else {}
        if self.path == "/prompt":
            if owner.reject_prompt:
                self._json({"error": "invalid prompt"}, 400)
                return
            self._json({"prompt_id": owner.prompt_id, "number": 1})
            return
        self._json({"error": "not found"}, 404)


class FakeComfyServer:
    def __init__(self) -> None:
        self.checkpoints = ["sd15.safetensors", "dreamshaper_8.safetensors"]
        self.prompt_id = "pid-1"
        self.empty_history_calls = 1
        self.history_calls = 0
        self.status_str = "success"
        self.fail = False
        self.reject_prompt = False
        self.requests: list[tuple[str, str]] = []
        self.last_body: dict | None = None
        self._httpd: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None

    def start(self) -> str:
        self._httpd = ThreadingHTTPServer(("127.0.0.1", 0), _ComfyHandler)
        self._httpd.owner = self  # type: ignore[attr-defined]
        self._thread = threading.Thread(target=self._httpd.serve_forever, daemon=True)
        self._thread.start()
        host, port = self._httpd.server_address[:2]
        return f"http://{host}:{port}"

    def stop(self) -> None:
        if self._httpd is not None:
            self._httpd.shutdown()
            self._httpd.server_close()
        if self._thread is not None:
            self._thread.join(timeout=3)


@pytest.fixture()
def fake_comfy():
    server = FakeComfyServer()
    yield server
    server.stop()


@pytest.fixture()
def comfy_settings(settings, fake_comfy):
    settings.comfyui_base_url = fake_comfy.start()
    settings.comfyui_timeout_seconds = 5.0
    settings.comfyui_generation_timeout_seconds = 10.0
    return settings
