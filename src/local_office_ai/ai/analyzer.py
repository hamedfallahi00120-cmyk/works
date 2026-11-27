"""AI orchestration: choose a model, send verified facts, validate the answer.

The analyzer never blocks the pipeline.  If Ollama is missing or the model
fails, it produces a deterministic Persian summary from the Python facts and
records exactly why the model was not used.
"""

from __future__ import annotations

import math
import time
from dataclasses import dataclass, field
from typing import Any

from local_office_ai.ai.context_builder import (
    allowed_numbers,
    build_fact_chunks,
    build_facts,
    trim_facts,
)
from local_office_ai.ai.ollama_client import ChatResult, ModelInfo, OllamaClient
from local_office_ai.ai.prompt_manager import build_messages, offline_summary
from local_office_ai.ai.response_validator import ValidatedResponse, validate_response
from local_office_ai.core.exceptions import AIServiceError
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import AnalysisResult, AuditStatus, MatchStatus
from local_office_ai.core.settings import Settings

LOGGER = get_logger(__name__)


@dataclass
class AIOutcome:
    mode: str  # "model" | "offline"
    response: ValidatedResponse
    model: str = ""
    metadata: dict[str, Any] = field(default_factory=dict)
    error: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode,
            "model": self.model,
            "response": self.response.to_dict(),
            "metadata": dict(self.metadata),
            "error": self.error,
        }


class AIAnalyzer:
    def __init__(self, settings: Settings, client: OllamaClient | None = None) -> None:
        self.settings = settings
        self.client = client or OllamaClient(settings)

    def available_models(self) -> list[ModelInfo]:
        return self.client.preferred_models()

    def analyze(
        self,
        analysis: AnalysisResult,
        *,
        model: str | None = None,
        task: str = "general_analysis",
        question: str | None = None,
        use_ai: bool = True,
    ) -> AIOutcome:
        facts, meta = build_facts(analysis, self.settings)
        mismatch_exists = any(audit.status is AuditStatus.MISMATCH for audit in analysis.audits) or any(
            item.status in {MatchStatus.PROBABLE, MatchStatus.OUTSIDE_LOAN_TABLE}
            for item in analysis.reconciliation
        )
        if analysis.intelligence is not None:
            mismatch_exists = mismatch_exists or any(
                item.status is AuditStatus.MISMATCH for item in analysis.intelligence.all_audits
            )
        statistical = getattr(analysis, "statistical_reconciliation", None)
        if statistical is not None:
            mismatch_exists = mismatch_exists or any(
                item.status.value == "FAIL" for item in statistical.reconciliations
            )

        if not use_ai:
            return self._offline(facts, meta, reason="تحلیل مدل به‌درخواست کاربر غیرفعال شد.")

        try:
            selected = model or self._select_model()
            if not selected:
                return self._offline(facts, meta, reason="هیچ مدل سازگاری در Ollama نصب نیست.")

            chunks = build_fact_chunks(analysis, self.settings)
            allowed = allowed_numbers(analysis)
            responses: list[tuple[str, ValidatedResponse]] = []
            stages: list[dict[str, Any]] = []
            errors: list[str] = []
            aggregate = {
                "context_chars": 0,
                "prompt_chars": 0,
                "max_context_chars": 0,
                "max_prompt_chars": 0,
                "max_estimated_prompt_tokens": 0,
                "trimmed_rows": 0,
                "trimmed_anomalies": 0,
                "trimmed_generic": 0,
                "attempts": 0,
                "output_tokens": 0,
                "prompt_tokens": 0,
                "eval_duration_seconds": 0.0,
                "total_duration_seconds": 0.0,
                "client_elapsed_seconds": 0.0,
                "degraded": False,
            }
            for label, payload in chunks:
                try:
                    validated, stage_meta = self._request_chunk(
                        selected, payload, task, question, label, allowed, mismatch_exists
                    )
                except AIServiceError as exc:
                    LOGGER.warning("AI chunk failed (%s): %s", label, exc)
                    stages.append({"label": label, "status": "failed", "error": str(exc)})
                    errors.append(f"{label}: {exc}")
                    continue
                responses.append((label, validated))
                stages.append({"label": label, "status": "ok", **stage_meta})
                aggregate["max_context_chars"] = max(
                    aggregate["max_context_chars"], int(stage_meta.get("context_chars") or 0)
                )
                aggregate["max_prompt_chars"] = max(
                    aggregate["max_prompt_chars"], int(stage_meta.get("prompt_chars") or 0)
                )
                aggregate["max_estimated_prompt_tokens"] = max(
                    aggregate["max_estimated_prompt_tokens"],
                    int(stage_meta.get("estimated_prompt_tokens") or 0),
                )
                for key in (
                    "context_chars",
                    "prompt_chars",
                    "trimmed_rows",
                    "trimmed_anomalies",
                    "trimmed_generic",
                    "attempts",
                    "output_tokens",
                    "prompt_tokens",
                ):
                    aggregate[key] += int(stage_meta.get(key) or 0)
                for key in (
                    "eval_duration_seconds",
                    "total_duration_seconds",
                    "client_elapsed_seconds",
                ):
                    aggregate[key] += float(stage_meta.get(key) or 0.0)
                aggregate["degraded"] = aggregate["degraded"] or bool(stage_meta.get("degraded"))

            if not responses:
                reason = "; ".join(errors) or "هیچ بخشی از درخواست مدل موفق نشد."
                outcome = self._offline(facts, meta, reason=reason)
                outcome.metadata.update(
                    {
                        "chunk_count": len(chunks),
                        "failed_chunks": len(errors),
                        "chunks": stages,
                        "max_ai_context_chars": self.settings.max_ai_context_chars,
                        "max_ai_prompt_chars": self.settings.max_ai_prompt_chars,
                    }
                )
                return outcome

            merged = self._merge_responses([response for _, response in responses])
            if aggregate["eval_duration_seconds"]:
                aggregate["tokens_per_second"] = round(
                    aggregate["output_tokens"] / aggregate["eval_duration_seconds"],
                    2,
                )
            else:
                aggregate["tokens_per_second"] = 0.0
            metadata = {
                "mode": "model",
                "model": selected,
                **aggregate,
                "chunk_count": len(chunks),
                "successful_chunks": len(responses),
                "failed_chunks": len(errors),
                "chunks": stages,
                "max_ai_context_chars": self.settings.max_ai_context_chars,
                "max_ai_prompt_chars": self.settings.max_ai_prompt_chars,
                "provenance": {
                    "python_analysis": "executed",
                    "model_analysis": "executed" if responses else "failed",
                    "rule_based_fallback": "not_used" if not errors else "partial",
                },
            }
            return AIOutcome(
                mode="model",
                response=merged,
                model=selected,
                metadata=metadata,
                error="; ".join(errors) if errors else "",
            )
        except AIServiceError as exc:
            LOGGER.warning("AI unavailable, falling back to offline summary: %s", exc)
            return self._offline(facts, meta, reason=str(exc))
        except Exception as exc:  # noqa: BLE001 - AI must never break the pipeline
            LOGGER.exception("unexpected AI failure")
            return self._offline(facts, meta, reason=f"خطای غیرمنتظره در تحلیل مدل: {exc}")

    # ------------------------------------------------------------------ #
    def _request_chunk(
        self,
        model: str,
        payload: dict[str, Any],
        task: str,
        question: str | None,
        label: str,
        allowed: set[str],
        mismatch_exists: bool,
    ) -> tuple[ValidatedResponse, dict[str, Any]]:
        facts, context_meta = trim_facts(payload, self.settings.max_ai_context_chars)
        messages, prompt_chars, context_meta = self._fit_prompt(facts, context_meta, task, question, label)
        attempts = 0
        result: ChatResult | None = None
        for attempt in range(2):
            try:
                started = time.perf_counter()
                result = self.client.chat(model, messages)
                break
            except AIServiceError as exc:
                # A 400 from Ollama usually means the prompt is longer than the
                # model context.  Shrink once and retry instead of crashing.
                if attempt == 0 and "400" in str(exc):
                    safe_budget = max(800, int(context_meta.get("chars", 2000) * self.settings.ai_chunk_retry_ratio))
                    facts, context_meta = trim_facts(facts, safe_budget)
                    context_meta["degraded"] = True
                    messages, prompt_chars, context_meta = self._fit_prompt(
                        facts, context_meta, task, question, label
                    )
                    attempts += 1
                    continue
                raise
        if result is None:
            raise AIServiceError("درخواست مدل اجرا نشد.")
        validated = validate_response(
            result.content,
            allowed_numbers=allowed,
            mismatch_exists=mismatch_exists,
        )
        timings = result.timings or {}
        return validated, {
            "attempts": result.attempts + attempts,
            "context_chars": context_meta.get("chars") or 0,
            "prompt_chars": prompt_chars,
            "estimated_prompt_tokens": context_meta.get("prompt_tokens_estimated") or 0,
            "trimmed_rows": context_meta.get("trimmed_rows") or 0,
            "trimmed_anomalies": context_meta.get("trimmed_anomalies") or 0,
            "trimmed_generic": context_meta.get("trimmed_generic") or 0,
            "degraded": bool(context_meta.get("degraded")),
            "output_tokens": timings.get("output_tokens") or 0,
            "prompt_tokens": timings.get("prompt_tokens") or 0,
            "eval_duration_seconds": timings.get("eval_duration_seconds") or 0.0,
            "total_duration_seconds": timings.get("total_duration_seconds") or 0.0,
            "client_elapsed_seconds": result.client_elapsed_seconds or 0.0,
        }

    def _fit_prompt(
        self,
        facts: dict[str, Any],
        context_meta: dict[str, Any],
        task: str,
        question: str | None,
        label: str,
    ) -> tuple[list[dict[str, str]], int, dict[str, Any]]:
        budget = self.settings.max_ai_context_chars
        messages = build_messages(facts, task=task, question=question, chunk_label=label)
        prompt_chars = sum(len(message["content"]) for message in messages)
        prompt_tokens = self._estimate_tokens(prompt_chars)
        token_budget = self._prompt_token_budget()
        guard = 0
        while (
            prompt_chars > self.settings.max_ai_prompt_chars or prompt_tokens > token_budget
        ) and guard < 6:
            budget = max(800, int(budget * 0.6))
            facts, context_meta = trim_facts(facts, budget)
            context_meta["degraded"] = True
            messages = build_messages(facts, task=task, question=question, chunk_label=label)
            prompt_chars = sum(len(message["content"]) for message in messages)
            prompt_tokens = self._estimate_tokens(prompt_chars)
            guard += 1
        context_meta["prompt_tokens_estimated"] = prompt_tokens
        return messages, prompt_chars, context_meta

    @staticmethod
    def _estimate_tokens(chars: int) -> int:
        # Qwen tokenization of mixed Persian/numbers is conservatively estimated
        # at 1.5 characters per token; rounding up keeps the request below the
        # configured context window rather than relying on a 400 retry.
        return max(1, math.ceil(chars / 1.5))

    def _prompt_token_budget(self) -> int:
        reserve = 1024
        return max(512, int(self.settings.ollama_num_ctx) - reserve)

    @staticmethod
    def _merge_responses(responses: list[ValidatedResponse]) -> ValidatedResponse:
        def unique(values: list[str]) -> list[str]:
            seen: set[str] = set()
            result: list[str] = []
            for value in values:
                if value and value not in seen:
                    seen.add(value)
                    result.append(value)
            return result

        summaries = unique([item.summary for item in responses])
        findings = unique([value for item in responses for value in item.findings])
        missing = unique([value for item in responses for value in item.missing_data])
        recommendations = unique([value for item in responses for value in item.recommendations])
        warnings = unique([value for item in responses for value in item.warnings])
        unsupported = unique([value for item in responses for value in item.unsupported_numbers])
        contradictions = unique([value for item in responses for value in item.contradictions])
        rank = {"low": 0, "medium": 1, "high": 2}
        risk = max(
            (item.risk_level for item in responses),
            key=lambda value: rank.get(value, -1),
        )
        return ValidatedResponse(
            summary=" ".join(summaries)[:2000],
            findings=findings[:12],
            missing_data=missing[:10],
            recommendations=recommendations[:10],
            risk_level=risk,
            unsupported_numbers=unsupported,
            contradictions=contradictions,
            warnings=warnings,
            is_grounded=all(item.is_grounded for item in responses),
            parsed_as_json=all(item.parsed_as_json for item in responses),
        )

    # ------------------------------------------------------------------ #
    def _select_model(self) -> str | None:
        models = self.client.preferred_models()
        if not models:
            return None
        for model in models:
            if model.name.endswith(":9b"):
                return model.name
        return models[0].name

    def _offline(self, facts: dict[str, Any], meta: dict[str, Any], reason: str) -> AIOutcome:
        payload = offline_summary(facts)
        response = ValidatedResponse(
            summary=payload["summary"],
            findings=payload["findings"],
            missing_data=payload["missing_data"],
            recommendations=payload["recommendations"],
            risk_level=payload["risk_level"],
            warnings=[f"تحلیل مدل انجام نشد: {reason}", "این گزارش با قواعد قطعی Python تولید شده است."],
            is_grounded=True,
            parsed_as_json=True,
        )
        return AIOutcome(
            mode="offline",
            response=response,
            model="",
            metadata={
                "mode": "offline",
                "context_chars": meta.get("chars"),
                "prompt_chars": 0,
                "trimmed_rows": meta.get("trimmed_rows"),
                "trimmed_anomalies": meta.get("trimmed_anomalies"),
                "chunk_count": None,
                "successful_chunks": None,
                "failed_chunks": None,
                "reason": reason,
                "provenance": {
                    "python_analysis": "executed",
                    "model_analysis": "not_executed",
                    "rule_based_fallback": "executed",
                },
            },
            error=reason,
        )

    @staticmethod
    def _metadata(
        result: ChatResult,
        elapsed: float,
        context_meta: dict[str, Any],
        model: str,
    ) -> dict[str, Any]:
        timings = result.timings
        return {
            "mode": "model",
            "model": model,
            "client_elapsed_seconds": elapsed,
            "attempts": result.attempts,
            "context_chars": context_meta.get("chars"),
            "trimmed_rows": context_meta.get("trimmed_rows"),
            "trimmed_anomalies": context_meta.get("trimmed_anomalies"),
            **timings,
        }
