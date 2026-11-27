"""Optional ComfyUI image studio (preserved from V4.2).

Image generation is delegated to a local ComfyUI server; the language model only
prepares the English prompt.  Everything here is optional: if ComfyUI is not
running the Excel pipeline is unaffected and the error is explicit.
"""

from __future__ import annotations

import json
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from local_office_ai.ai.ollama_client import OllamaClient
from local_office_ai.core.exceptions import ComfyUIError
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import AnalysisResult
from local_office_ai.core.settings import Settings

LOGGER = get_logger(__name__)

DEFAULT_NEGATIVE = (
    "text, watermark, logo, signature, low quality, blurry, deformed, extra fingers, "
    "duplicate subject, oversaturated"
)
PROMPT_SYSTEM_MESSAGE = (
    "You are an expert Stable Diffusion prompt designer. Convert the user's request into one "
    "detailed, coherent English image-generation prompt. Preserve the requested subject and intent. "
    "Specify composition, lighting, materials, environment, camera/viewpoint and visual style when "
    "appropriate. Do not add text, logos or watermarks. Return only the final prompt."
)


@dataclass
class ImageRequest:
    prompt: str
    negative_prompt: str = DEFAULT_NEGATIVE
    checkpoint: str | None = None
    width: int = 512
    height: int = 512
    steps: int = 20
    cfg: float = 7.0
    seed: int | None = None


@dataclass
class ImageResult:
    prompt: str
    checkpoint: str
    prompt_id: str
    images: list[Path] = field(default_factory=list)
    seconds: float = 0.0
    model_used: str = ""
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "prompt": self.prompt,
            "checkpoint": self.checkpoint,
            "prompt_id": self.prompt_id,
            "images": [str(path) for path in self.images],
            "seconds": self.seconds,
            "model_used": self.model_used,
            "warnings": list(self.warnings),
        }


class ComfyUIClient:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.base_url = settings.comfyui_base_url.rstrip("/")

    # ------------------------------------------------------------------ #
    def _request(
        self,
        path: str,
        *,
        method: str = "GET",
        payload: dict[str, Any] | None = None,
        timeout: float | None = None,
    ) -> bytes:
        url = f"{self.base_url}{path}"
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
        request = urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"} if data else {},
            method=method,
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout or self.settings.comfyui_timeout_seconds) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "replace")[:400]
            raise ComfyUIError(f"ComfyUI پاسخ {exc.code} داد: {detail}") from exc
        except urllib.error.URLError as exc:
            raise ComfyUIError(f"اتصال به ComfyUI برقرار نشد ({self.base_url}): {exc.reason}") from exc
        except TimeoutError as exc:
            raise ComfyUIError(f"زمان پاسخ ComfyUI تمام شد: {path}") from exc

    def _json(self, path: str, **kwargs) -> Any:
        raw = self._request(path, **kwargs)
        try:
            return json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ComfyUIError(f"پاسخ ComfyUI قابل خواندن نبود: {path}") from exc

    # ------------------------------------------------------------------ #
    def health(self) -> bool:
        try:
            self.checkpoints()
            return True
        except ComfyUIError:
            return False

    def checkpoints(self) -> list[str]:
        try:
            value = self._json("/models/checkpoints", timeout=12)
            names = _checkpoint_names(value)
            if names:
                return names
        except ComfyUIError:
            pass
        try:
            info = self._json("/object_info/CheckpointLoaderSimple", timeout=12)
            node = info.get("CheckpointLoaderSimple", info) if isinstance(info, dict) else {}
            required = (node.get("input") or {}).get("required") or {}
            values = required.get("ckpt_name") or []
            if values and isinstance(values[0], list):
                return [str(item) for item in values[0] if item]
        except ComfyUIError as exc:
            raise ComfyUIError(f"فهرست checkpoint ها خوانده نشد: {exc}") from exc
        return []

    @staticmethod
    def build_workflow(request: ImageRequest, checkpoint: str) -> dict[str, Any]:
        seed = request.seed if request.seed is not None else secrets.randbelow(2**32)
        return {
            "1": {"class_type": "CheckpointLoaderSimple", "inputs": {"ckpt_name": checkpoint}},
            "2": {"class_type": "CLIPTextEncode", "inputs": {"text": request.prompt, "clip": ["1", 1]}},
            "3": {"class_type": "CLIPTextEncode", "inputs": {"text": request.negative_prompt, "clip": ["1", 1]}},
            "4": {
                "class_type": "EmptyLatentImage",
                "inputs": {"width": int(request.width), "height": int(request.height), "batch_size": 1},
            },
            "5": {
                "class_type": "KSampler",
                "inputs": {
                    "seed": int(seed),
                    "steps": int(request.steps),
                    "cfg": float(request.cfg),
                    "sampler_name": "euler",
                    "scheduler": "normal",
                    "denoise": 1.0,
                    "model": ["1", 0],
                    "positive": ["2", 0],
                    "negative": ["3", 0],
                    "latent_image": ["4", 0],
                },
            },
            "6": {"class_type": "VAEDecode", "inputs": {"samples": ["5", 0], "vae": ["1", 2]}},
            "7": {
                "class_type": "SaveImage",
                "inputs": {"filename_prefix": "LocalOfficeAI", "images": ["6", 0]},
            },
        }

    # ------------------------------------------------------------------ #
    def generate(
        self,
        request: ImageRequest,
        *,
        progress: Callable[[str], None] | None = None,
    ) -> ImageResult:
        started = time.perf_counter()
        checkpoint = request.checkpoint or self.settings.comfyui_default_checkpoint
        if not checkpoint:
            available = self.checkpoints()
            if not available:
                raise ComfyUIError("هیچ checkpoint در ComfyUI نصب نیست.")
            checkpoint = available[0]

        workflow = self.build_workflow(request, checkpoint)
        client_id = str(uuid.uuid4())
        response = self._json("/prompt", method="POST", payload={"prompt": workflow, "client_id": client_id}, timeout=60)
        prompt_id = str((response or {}).get("prompt_id") or "")
        if not prompt_id:
            raise ComfyUIError("ComfyUI درخواست را نپذیرفت (prompt_id دریافت نشد).")
        if progress:
            progress(f"کار ComfyUI ثبت شد: {prompt_id[:12]}")

        images = self._wait_for_images(prompt_id, progress=progress)
        saved = self._download(images)
        return ImageResult(
            prompt=request.prompt,
            checkpoint=checkpoint,
            prompt_id=prompt_id,
            images=saved,
            seconds=time.perf_counter() - started,
        )

    def _wait_for_images(
        self,
        prompt_id: str,
        *,
        progress: Callable[[str], None] | None = None,
    ) -> list[dict[str, Any]]:
        deadline = time.monotonic() + self.settings.comfyui_generation_timeout_seconds
        while time.monotonic() < deadline:
            history = self._json(f"/history/{prompt_id}", timeout=20)
            entry = history.get(prompt_id, {}) if isinstance(history, dict) else {}
            if isinstance(entry, dict):
                status = entry.get("status") or {}
                if isinstance(status, dict) and str(status.get("status_str", "")).lower() in {"error", "failed"}:
                    raise ComfyUIError(
                        "اجرای workflow در ComfyUI شکست خورد: "
                        + json.dumps(status.get("messages", []), ensure_ascii=False)[:600]
                    )
                items: list[dict[str, Any]] = []
                for node_output in (entry.get("outputs") or {}).values():
                    if isinstance(node_output, dict):
                        items.extend(node_output.get("images", []) or [])
                if items:
                    return items
            if progress:
                progress("ComfyUI در حال تولید تصویر است…")
            time.sleep(2.0)
        raise ComfyUIError(f"زمان انتظار تولید تصویر تمام شد (prompt_id={prompt_id}).")

    def _download(self, items: list[dict[str, Any]]) -> list[Path]:
        target_dir = self.settings.output_dir / "images"
        target_dir.mkdir(parents=True, exist_ok=True)
        saved: list[Path] = []
        for index, item in enumerate(items, start=1):
            filename = str(item.get("filename") or f"image_{index}.png")
            query = urllib.parse.urlencode(
                {
                    "filename": filename,
                    "subfolder": item.get("subfolder", ""),
                    "type": item.get("type", "output"),
                }
            )
            data = self._request(f"/view?{query}", timeout=120)
            safe_name = Path(filename).name or f"image_{index}.png"
            path = target_dir / f"{time.strftime('%Y%m%d_%H%M%S')}_{index}_{safe_name}"
            path.write_bytes(data)
            saved.append(path)
        return saved


def _checkpoint_names(value: Any) -> list[str]:
    names: list[str] = []
    if isinstance(value, list):
        for item in value:
            if isinstance(item, str):
                names.append(item)
            elif isinstance(item, dict):
                names.append(str(item.get("name") or item.get("filename") or ""))
    elif isinstance(value, dict):
        for item in value.get("models", []) or []:
            if isinstance(item, str):
                names.append(item)
            elif isinstance(item, dict):
                names.append(str(item.get("name") or item.get("filename") or ""))
    return sorted({name for name in names if name})


def describe_analysis(analysis: AnalysisResult | None) -> str:
    """Deterministic image idea derived from the audited figures."""
    if analysis is None:
        return "A tidy Persian office desk with a financial report, warm daylight, clean composition"
    parts = ["A clean modern Persian office scene"]
    loan_count = analysis.calculation("loan_count")
    if loan_count is not None:
        parts.append(f"{loan_count.value} loan files arranged on a desk")
    if any(item.status.value == "outside_loan_table" for item in analysis.reconciliation):
        parts.append("one document highlighted as needing review")
    parts.append("soft window light, shallow depth of field, professional editorial photography")
    return ", ".join(parts)


class ImageStudio:
    """Prompt preparation plus generation, with a deterministic fallback prompt."""

    def __init__(self, settings: Settings, client: ComfyUIClient | None = None) -> None:
        self.settings = settings
        self.comfy = client or ComfyUIClient(settings)

    def prepare_prompt(
        self,
        user_prompt: str,
        *,
        analysis: AnalysisResult | None = None,
        model: str | None = None,
        use_ai: bool = True,
    ) -> tuple[str, str, list[str]]:
        base = (user_prompt or "").strip() or describe_analysis(analysis)
        warnings: list[str] = []
        if not use_ai:
            return base, "", warnings
        try:
            ollama = OllamaClient(self.settings)
            selected = model or next(
                (item.name for item in ollama.preferred_models()), ""
            )
            if not selected:
                warnings.append("مدلی برای بهبود پرامپت پیدا نشد؛ پرامپت ورودی استفاده شد.")
                return base, "", warnings
            result = ollama.chat(
                selected,
                [
                    {"role": "system", "content": PROMPT_SYSTEM_MESSAGE},
                    {"role": "user", "content": base},
                ],
                json_format=False,
                temperature=0.35,
            )
            text = result.content.strip().strip('"').strip()
            if not text:
                warnings.append("پاسخ مدل خالی بود؛ پرامپت ورودی استفاده شد.")
                return base, selected, warnings
            return text, selected, warnings
        except Exception as exc:  # noqa: BLE001 - image prompt is optional
            warnings.append(f"بهبود پرامپت با مدل انجام نشد: {exc}")
            return base, "", warnings

    def generate(
        self,
        user_prompt: str,
        *,
        analysis: AnalysisResult | None = None,
        model: str | None = None,
        use_ai: bool = True,
        request: ImageRequest | None = None,
        progress: Callable[[str], None] | None = None,
    ) -> ImageResult:
        prompt, used_model, warnings = self.prepare_prompt(
            user_prompt, analysis=analysis, model=model, use_ai=use_ai
        )
        payload = request or ImageRequest(prompt=prompt)
        payload.prompt = prompt
        result = self.comfy.generate(payload, progress=progress)
        result.model_used = used_model
        result.warnings.extend(warnings)
        LOGGER.info("image generated: %s (%.1fs)", result.images, result.seconds)
        return result
