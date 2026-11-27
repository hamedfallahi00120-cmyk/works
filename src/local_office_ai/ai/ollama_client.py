"""Minimal, robust Ollama client.

Uses only the standard library so the project does not grow a dependency for a
single HTTP endpoint.  Errors become :class:`AIServiceError` with an actionable
Persian message; connection problems are retried, model errors are not.
"""

from __future__ import annotations

import json
import socket
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from local_office_ai.core.exceptions import AIServiceError
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.settings import Settings

LOGGER = get_logger(__name__)


@dataclass
class ModelInfo:
    name: str
    size_bytes: int = 0
    family: str = ""
    parameter_size: str = ""
    quantization: str = ""
    modified_at: str = ""

    @property
    def size_gb(self) -> float:
        return round(self.size_bytes / (1024 ** 3), 2)

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "size_gb": self.size_gb,
            "family": self.family,
            "parameter_size": self.parameter_size,
            "quantization": self.quantization,
            "modified_at": self.modified_at,
        }


@dataclass
class ChatResult:
    content: str
    model: str
    attempts: int
    client_elapsed_seconds: float
    raw: dict[str, Any] = field(default_factory=dict)
    timings: dict[str, Any] = field(default_factory=dict)


class OllamaClient:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.base_url = settings.ollama_base_url.rstrip("/")

    def _request(
        self,
        path: str,
        *,
        payload: dict[str, Any] | None = None,
        method: str = "GET",
        timeout: float | None = None,
    ) -> Any:
        url = f"{self.base_url}{path}"
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
        headers = {"Content-Type": "application/json"} if data else {}
        request = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(
                request, timeout=timeout or self.settings.ollama_connect_timeout_seconds
            ) as response:
                raw = response.read()
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                detail = exc.read().decode("utf-8", errors="replace")[:400]
            except Exception:  # noqa: BLE001
                detail = ""
            if exc.code == 404:
                raise AIServiceError(f"مدل یا مسیر در Ollama پیدا نشد (HTTP 404). {detail}".strip()) from exc
            if 400 <= exc.code < 500:
                raise AIServiceError(f"Ollama درخواست را رد کرد (HTTP {exc.code}). {detail}".strip()) from exc
            raise AIServiceError(f"خطای سرور Ollama (HTTP {exc.code}). {detail}".strip()) from exc
        except (urllib.error.URLError, socket.timeout, TimeoutError) as exc:
            raise AIServiceError(
                f"ارتباط با Ollama برقرار نشد ({self.base_url}). آیا سرویس Ollama در حال اجراست؟ جزئیات: {exc}"
            ) from exc
        if not raw:
            return {}
        try:
            return json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise AIServiceError("پاسخ Ollama JSON معتبر نبود.") from exc

    def health(self) -> bool:
        try:
            self._request("/api/tags", timeout=self.settings.ollama_connect_timeout_seconds)
            return True
        except AIServiceError:
            return False

    def list_models(self) -> list[ModelInfo]:
        payload = self._request("/api/tags", timeout=self.settings.ollama_connect_timeout_seconds)
        models: list[ModelInfo] = []
        for item in payload.get("models", []) or []:
            if not isinstance(item, dict) or not item.get("name"):
                continue
            details = item.get("details") or {}
            models.append(
                ModelInfo(
                    name=str(item.get("name")),
                    size_bytes=int(item.get("size") or 0),
                    family=str(details.get("family") or ""),
                    parameter_size=str(details.get("parameter_size") or ""),
                    quantization=str(details.get("quantization_level") or ""),
                    modified_at=str(item.get("modified_at") or ""),
                )
            )
        models.sort(key=lambda model: model.name)
        return models

    def preferred_models(self) -> list[ModelInfo]:
        models = self.list_models()
        prefixes = tuple(prefix.lower() for prefix in self.settings.model_name_prefixes)
        matching = [model for model in models if model.name.lower().startswith(prefixes)]
        return matching or models

    def show_model(self, name: str) -> dict[str, Any]:
        return self._request("/api/show", payload={"name": name}, method="POST")

    def chat(
        self,
        model: str,
        messages: list[dict[str, str]],
        *,
        json_format: bool = True,
        temperature: float | None = None,
        num_ctx: int | None = None,
        timeout: float | None = None,
        keep_alive: str | int | None = None,
    ) -> ChatResult:
        body: dict[str, Any] = {
            "model": model,
            "stream": False,
            "messages": messages,
            "options": {
                "temperature": self.settings.ollama_temperature if temperature is None else temperature,
                "num_ctx": self.settings.ollama_num_ctx if num_ctx is None else num_ctx,
            },
        }
        if json_format:
            body["format"] = "json"
        if keep_alive is not None:
            body["keep_alive"] = keep_alive

        attempts = 0
        last_error: Exception | None = None
        started = time.perf_counter()
        while attempts <= self.settings.ollama_retries:
            attempts += 1
            try:
                payload = self._request(
                    "/api/chat",
                    payload=body,
                    method="POST",
                    timeout=timeout or self.settings.ollama_timeout_seconds,
                )
                elapsed = time.perf_counter() - started
                content = str((payload.get("message") or {}).get("content") or "")
                timings = self._extract_timings(payload)
                LOGGER.info(
                    "ollama chat ok model=%s attempts=%d elapsed=%.2fs tokens_out=%s",
                    model, attempts, elapsed, timings.get("output_tokens"),
                )
                return ChatResult(
                    content=content,
                    model=model,
                    attempts=attempts,
                    client_elapsed_seconds=elapsed,
                    raw=payload,
                    timings=timings,
                )
            except AIServiceError as exc:
                last_error = exc
                retryable = "برقرار نشد" in str(exc) or "HTTP 5" in str(exc)
                if not retryable or attempts > self.settings.ollama_retries:
                    raise
                wait = min(2.0 * attempts, 6.0)
                LOGGER.warning("ollama chat retry %d after error: %s", attempts, exc)
                time.sleep(wait)
        raise AIServiceError(f"درخواست مدل ناموفق ماند: {last_error}")

    @staticmethod
    def _extract_timings(payload: dict[str, Any]) -> dict[str, Any]:
        def seconds(key: str) -> float | None:
            value = payload.get(key)
            return value / 1e9 if isinstance(value, (int, float)) else None

        output_tokens = payload.get("eval_count")
        eval_duration = payload.get("eval_duration")
        tokens_per_second = None
        if isinstance(output_tokens, (int, float)) and isinstance(eval_duration, (int, float)) and eval_duration > 0:
            tokens_per_second = output_tokens / (eval_duration / 1e9)
        return {
            "total_duration_seconds": seconds("total_duration"),
            "load_duration_seconds": seconds("load_duration"),
            "prompt_eval_duration_seconds": seconds("prompt_eval_duration"),
            "eval_duration_seconds": seconds("eval_duration"),
            "prompt_tokens": payload.get("prompt_eval_count"),
            "output_tokens": output_tokens,
            "tokens_per_second": tokens_per_second,
        }


def parse_timestamp(value: str) -> str:
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).strftime("%Y-%m-%d %H:%M")
    except (ValueError, AttributeError):
        return value
