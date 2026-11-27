"""ComfyUI image studio (optional V4.2 capability) against a real local server."""

from __future__ import annotations

import pytest

from local_office_ai.ai.image_studio import (
    ComfyUIClient,
    ImageRequest,
    ImageStudio,
    describe_analysis,
)
from local_office_ai.core.exceptions import ComfyUIError


def test_checkpoint_list(comfy_settings):
    assert ComfyUIClient(comfy_settings).checkpoints() == [
        "dreamshaper_8.safetensors",
        "sd15.safetensors",
    ]
    assert ComfyUIClient(comfy_settings).health() is True


def test_workflow_graph_is_valid_sd15(settings):
    request = ImageRequest(prompt="a desk", width=640, height=384, steps=12, cfg=6.5, seed=7)
    workflow = ComfyUIClient.build_workflow(request, "sd15.safetensors")
    assert workflow["1"]["class_type"] == "CheckpointLoaderSimple"
    assert workflow["1"]["inputs"]["ckpt_name"] == "sd15.safetensors"
    assert workflow["2"]["inputs"]["text"] == "a desk"
    assert workflow["4"]["inputs"]["width"] == 640
    assert workflow["5"]["inputs"]["steps"] == 12
    assert workflow["5"]["inputs"]["seed"] == 7
    assert workflow["5"]["inputs"]["positive"] == ["2", 0]
    assert workflow["7"]["class_type"] == "SaveImage"


def test_generate_saves_png(comfy_settings, fake_comfy):
    result = ComfyUIClient(comfy_settings).generate(ImageRequest(prompt="an office"))
    assert result.prompt_id == "pid-1"
    assert result.checkpoint == "dreamshaper_8.safetensors"
    assert len(result.images) == 1
    assert result.images[0].exists()
    assert result.images[0].read_bytes().startswith(b"\x89PNG")
    sent = fake_comfy.last_body
    assert sent and sent["prompt"]["2"]["inputs"]["text"] == "an office"
    assert any(item[0] == "GET" and item[1].startswith("/view?") for item in fake_comfy.requests)


def test_workflow_error_status_is_reported(comfy_settings, fake_comfy):
    fake_comfy.status_str = "error"
    with pytest.raises(ComfyUIError):
        ComfyUIClient(comfy_settings).generate(ImageRequest(prompt="x"))


def test_rejected_prompt_is_reported(comfy_settings, fake_comfy):
    fake_comfy.reject_prompt = True
    with pytest.raises(ComfyUIError):
        ComfyUIClient(comfy_settings).generate(ImageRequest(prompt="x"))


def test_comfy_unavailable(settings):
    settings.comfyui_base_url = "http://127.0.0.1:9"
    with pytest.raises(ComfyUIError):
        ComfyUIClient(settings).checkpoints()
    assert ComfyUIClient(settings).health() is False


def test_prompt_improvement_with_model(ollama_settings, fake_ollama):
    fake_ollama.chat_content = "a detailed cinematic office scene, warm light"
    prompt, model, warnings = ImageStudio(ollama_settings).prepare_prompt("دفتر اداری", use_ai=True)
    assert prompt == "a detailed cinematic office scene, warm light"
    assert model == "qwen3.5:9b"
    assert warnings == []
    assert "format" not in fake_ollama.last_body


def test_prompt_falls_back_without_model(settings):
    settings.ollama_base_url = "http://127.0.0.1:9"
    settings.ollama_retries = 0
    studio = ImageStudio(settings)
    prompt, model, warnings = studio.prepare_prompt("دفتر اداری", use_ai=True)
    assert prompt == "دفتر اداری"
    assert model == ""
    assert warnings and "بهبود پرامپت" in warnings[0]


def test_generate_uses_analysis_prompt_when_empty(settings, fake_comfy, sample_path):
    from local_office_ai.app.workflow import ExcelIntelligenceWorkflow

    fake_comfy.start()
    settings.comfyui_base_url = f"http://127.0.0.1:{fake_comfy._httpd.server_address[1]}"
    analysis = ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)
    result = ImageStudio(settings).generate("", analysis=analysis, use_ai=False)
    assert "office" in result.prompt.lower()
    assert result.images
    assert describe_analysis(analysis).startswith("A clean modern Persian office")
