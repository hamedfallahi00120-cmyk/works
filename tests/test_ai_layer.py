from __future__ import annotations

import json

import pytest

from local_office_ai.ai.analyzer import AIAnalyzer
from local_office_ai.ai.context_builder import allowed_numbers, build_facts
from local_office_ai.ai.ollama_client import OllamaClient
from local_office_ai.ai.response_validator import validate_response
from local_office_ai.core.exceptions import AIServiceError


@pytest.fixture()
def sample_analysis(settings, sample_path):
    from local_office_ai.app.workflow import ExcelIntelligenceWorkflow

    return ExcelIntelligenceWorkflow(settings).analyze_only(sample_path)


def test_list_and_preferred_models(ollama_settings):
    client = OllamaClient(ollama_settings)
    assert client.health()
    names = [model.name for model in client.list_models()]
    assert names == ["llama3:8b", "qwen3.5:9b"]
    preferred = [model.name for model in client.preferred_models()]
    assert preferred == ["qwen3.5:9b"]
    assert client.list_models()[1].size_gb == pytest.approx(5.59, abs=0.01)


def test_chat_returns_content_and_timings(ollama_settings, fake_ollama):
    result = OllamaClient(ollama_settings).chat("qwen3.5:9b", [{"role": "user", "content": "سلام"}])
    assert "خلاصه آزمایشی" in result.content
    assert result.attempts == 1
    assert result.timings["prompt_tokens"] == 700
    assert result.timings["output_tokens"] == 120
    assert result.timings["tokens_per_second"] == pytest.approx(120.0)
    assert result.timings["total_duration_seconds"] == pytest.approx(2.0)
    assert fake_ollama.last_body["format"] == "json"


def test_http_404_is_reported(ollama_settings, fake_ollama):
    fake_ollama.chat_status = 404
    with pytest.raises(AIServiceError) as info:
        OllamaClient(ollama_settings).chat("missing", [{"role": "user", "content": "x"}])
    assert "404" in str(info.value)


def test_server_error_is_retried_then_fails(ollama_settings, fake_ollama):
    fake_ollama.chat_status = 500
    with pytest.raises(AIServiceError):
        OllamaClient(ollama_settings).chat("qwen3.5:9b", [{"role": "user", "content": "x"}])
    chat_calls = [item for item in fake_ollama.requests if item == ("POST", "/api/chat")]
    assert len(chat_calls) == 2  # one try + one retry


def test_invalid_json_response(ollama_settings, fake_ollama):
    fake_ollama.chat_raw = "not-json"
    with pytest.raises(AIServiceError):
        OllamaClient(ollama_settings).chat("qwen3.5:9b", [{"role": "user", "content": "x"}])


def test_connection_refused(settings):
    settings.ollama_base_url = "http://127.0.0.1:9"
    settings.ollama_retries = 0
    with pytest.raises(AIServiceError) as info:
        OllamaClient(settings).list_models()
    assert "برقرار نشد" in str(info.value)


def test_validator_accepts_grounded_answer():
    response = validate_response(
        json.dumps(
            {
                "summary": "جمع اصل ۳۲۵,۰۰۰,۰۰۰ است.",
                "findings": ["مبلغ ۴,۲۵۰,۰۰۰ خارج از جدول است."],
                "missing_data": [],
                "recommendations": ["بررسی"],
                "risk_level": "medium",
            },
            ensure_ascii=False,
        ),
        allowed_numbers={"325000000", "4250000"},
        mismatch_exists=True,
    )
    assert response.is_grounded
    assert response.unsupported_numbers == []
    assert response.risk_level == "medium"
    assert response.parsed_as_json


def test_validator_flags_invented_numbers():
    response = validate_response(
        json.dumps({"summary": "جمع کل ۹۹۹,۹۹۹,۹۹۹ است.", "findings": [], "missing_data": [], "recommendations": []}, ensure_ascii=False),
        allowed_numbers={"325000000"},
        mismatch_exists=False,
    )
    assert not response.is_grounded
    assert "999,999,999" in response.unsupported_numbers


def test_validator_flags_contradiction():
    response = validate_response(
        json.dumps({"summary": "مغایرت جدی وجود دارد.", "findings": [], "missing_data": [], "recommendations": []}, ensure_ascii=False),
        allowed_numbers=set(),
        mismatch_exists=False,
    )
    assert not response.is_grounded
    assert response.contradictions


def test_validator_non_json_falls_back():
    response = validate_response("این پاسخ JSON نیست", allowed_numbers=set(), mismatch_exists=True)
    assert not response.parsed_as_json
    assert not response.is_grounded
    assert response.summary


def test_facts_contain_only_verified_calculations(sample_analysis, settings):
    facts, meta = build_facts(sample_analysis, settings)
    calculations = facts["verified_calculations_from_python"]
    assert calculations["principal_sum"]["value"] == "325000000"
    assert calculations["obligations_sum"]["value"] == "18999000"
    assert facts["primary_table"]["kind"] == "loan"
    assert meta["chars"] > 0
    assert "325000000" in allowed_numbers(sample_analysis)


def test_facts_are_trimmed_to_budget(sample_analysis, settings):
    settings.max_ai_context_chars = 2500
    facts, meta = build_facts(sample_analysis, settings)
    assert meta["chars"] <= 2500
    assert meta["trimmed_rows"] > 0


def test_analyzer_uses_model(ollama_settings, sample_analysis, fake_ollama):
    outcome = AIAnalyzer(ollama_settings).analyze(sample_analysis, use_ai=True)
    assert outcome.mode == "model"
    assert outcome.model == "qwen3.5:9b"
    assert outcome.response.is_grounded
    assert outcome.metadata["output_tokens"] == 120
    assert outcome.metadata["tokens_per_second"] == pytest.approx(120.0)


def test_analyzer_falls_back_to_offline(settings, sample_analysis):
    settings.ollama_base_url = "http://127.0.0.1:9"
    settings.ollama_retries = 0
    outcome = AIAnalyzer(settings).analyze(sample_analysis, use_ai=True)
    assert outcome.mode == "offline"
    assert outcome.error
    assert "325,000,000" in outcome.response.summary
    assert outcome.response.is_grounded
    assert any("Python" in warning for warning in outcome.response.warnings)


def test_analyzer_respects_no_ai(ollama_settings, sample_analysis):
    outcome = AIAnalyzer(ollama_settings).analyze(sample_analysis, use_ai=False)
    assert outcome.mode == "offline"
    assert "غیرفعال" in outcome.error


def test_analyzer_accepts_scaled_claim_through_real_client(ollama_settings, sample_analysis, fake_ollama):
    fake_ollama.chat_content = json.dumps(
        {
            "summary": "مجموع اصل وام‌ها ۳۲۵ میلیون تومان است.",
            "findings": [],
            "missing_data": [],
            "recommendations": [],
            "risk_level": "low",
        },
        ensure_ascii=False,
    )
    outcome = AIAnalyzer(ollama_settings).analyze(sample_analysis, use_ai=True)
    assert outcome.mode == "model"
    assert outcome.response.is_grounded
    assert outcome.response.unsupported_numbers == []


def test_analyzer_rejects_one_unit_off_through_real_client(ollama_settings, sample_analysis, fake_ollama):
    fake_ollama.chat_content = json.dumps(
        {
            "summary": "مجموع اصل وام‌ها ۳۲۵,۰۰۰,۰۰۱ تومان است.",
            "findings": [],
            "missing_data": [],
            "recommendations": [],
            "risk_level": "low",
        },
        ensure_ascii=False,
    )
    outcome = AIAnalyzer(ollama_settings).analyze(sample_analysis, use_ai=True)
    assert outcome.mode == "model"
    assert not outcome.response.is_grounded
    assert outcome.response.unsupported_numbers
