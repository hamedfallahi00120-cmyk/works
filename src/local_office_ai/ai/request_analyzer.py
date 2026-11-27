"""Deterministic understanding of a Persian analysis request.

The user writes one or two sentences ("مقایسه ماهانه شاخص‌های تردد و گزارش
مدیریتی بده") and the engine must decide which analysis and which sheets are
actually needed.  This module only *classifies* the request: no figure is ever
produced or changed here.  Everything numeric still comes from Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable

from local_office_ai.excel.normalization import normalize_digits, normalize_text

# --------------------------------------------------------------------------- #
# Intent vocabulary
# --------------------------------------------------------------------------- #
SCOPE_SUMMARY = "summary"
SCOPE_COMPARISON = "period_comparison"
SCOPE_TREND = "trend"
SCOPE_RANKING = "ranking"
SCOPE_QUALITY = "data_quality"
SCOPE_GROUP = "group_comparison"
SCOPE_CORRELATION = "correlation"
SCOPE_DISTRIBUTION = "distribution"
SCOPE_DISCREPANCY = "discrepancy"

SCOPE_FA = {
    SCOPE_SUMMARY: "خلاصه و تحلیل کلی",
    SCOPE_COMPARISON: "مقایسه دوره‌ها",
    SCOPE_TREND: "روند زمانی",
    SCOPE_RANKING: "رتبه‌بندی و بیشترین/کمترین",
    SCOPE_QUALITY: "کیفیت داده",
    SCOPE_GROUP: "مقایسه گروه‌ها",
    SCOPE_CORRELATION: "همبستگی",
    SCOPE_DISTRIBUTION: "توزیع فراوانی",
    SCOPE_DISCREPANCY: "مغایرت‌ها و کنترل جمع",
}

_PATTERNS: dict[str, tuple[str, ...]] = {
    SCOPE_COMPARISON: (
        "مقایسه", "مقایسه‌ای", "نسبت به", "در برابر", "تفاوت", "اختلاف",
        "درصد تغییر", "رشد", "کاهش", "افزایش", "سال قبل", "سال گذشته", "دوره قبل",
    ),
    SCOPE_TREND: ("روند", "رویه", "تغییرات زمانی", "سیر", "طی ماه", "طی سال", "trend"),
    SCOPE_RANKING: (
        "بیشترین", "کمترین", "رتبه", "رتبه‌بندی", "بهترین", "بدترین",
        "بالاترین", "پایین‌ترین", "اول", "top", "rank",
    ),
    SCOPE_QUALITY: (
        "کیفیت", "ناقص", "تکراری", "خالی", "نامعتبر", "خطا", "کنترل کیفیت",
        "اعتبارسنجی", "پاکسازی",
    ),
    SCOPE_GROUP: ("گروه", "به تفکیک", "بر اساس", "دسته", "طبقه"),
    SCOPE_CORRELATION: ("همبستگی", "رابطه", "correlation"),
    SCOPE_DISTRIBUTION: ("توزیع", "فراوانی", "بازه", "هیستوگرام"),
    SCOPE_DISCREPANCY: (
        "مغایرت", "تطبیق", "جمع کل", "کنترل جمع", "حسابرسی", "اختلاف جمع",
    ),
}

_PERIOD_WORDS = {
    "ماهانه": "month",
    "ماهی": "month",
    "روزانه": "day",
    "روزی": "day",
    "سالانه": "year",
    "سال": "year",
    "فصلی": "quarter",
}

_AGGREGATION_WORDS = {
    "جمع": "sum",
    "مجموع": "sum",
    "میانگین": "mean",
    "میانه": "median",
    "تعداد": "count",
    "درصد": "percent",
}

_OUTPUT_HINTS = {
    "table_only": ("فقط جدول", "فقط یک جدول", "جدول بده", "فقط خروجی جدولی", "table only"),
    "dashboard": ("داشبورد", "dashboard"),
    "chart": ("نمودار", "چارت", "char", "بصری", "گراف"),
    "management": ("گزارش مدیریتی", "گزارش مدیران", "خلاصه مدیریتی", "مدیریتی"),
    "report": ("گزارش", "report"),
    "audit": ("حسابرسی کامل", "جزئیات فنی", "audit", "ردیابی منبع", "داده خام"),
}

# Words that carry no discriminating power when matching a metric header.
_STOPWORDS = {
    "تعداد", "کل", "مجموع", "جمع", "مقدار", "درصد", "نسبت", "شاخص", "میزان",
    "و", "در", "به", "از", "برای", "با", "که", "را", "های", "ها", "یک", "دو",
    "ماه", "ماهه", "سال", "روز", "روزانه", "کل", "ساز", "عملکرد",
    "لطفا", "لطفاً", "می", "خواهم", "بده", "بدهید", "تهیه", "کن", "کنید",
}

_SHEET_LANGUAGE = {
    "dashboard": "داشبورد",
    "management": "خلاصه مدیریتی",
    "request": "پاسخ درخواست",
    "monthly": "جدول ماهانه شاخص‌ها",
    "comparison": "مقایسه دوره‌ای شاخص‌ها",
    "descriptive": "آمار توصیفی",
    "frequency": "توزیع فراوانی",
    "groups": "تحلیل گروه‌ها",
    "trend": "روند و همبستگی",
    "quality": "کیفیت داده",
    "issues": "موارد نیازمند بررسی",
    "audit": "حسابرسی کامل",
}


@dataclass
class RequestIntent:
    """What the user asked for, with the evidence used to decide it."""

    text: str = ""
    normalized: str = ""
    scopes: list[str] = field(default_factory=list)
    requested_metrics: list[str] = field(default_factory=list)
    matched_metrics: list[str] = field(default_factory=list)
    period: str = ""
    aggregation: str = ""
    wants_table_only: bool = False
    wants_chart: bool = False
    wants_dashboard: bool = False
    wants_management: bool = False
    wants_audit: bool = False
    ranking_wants_least: bool = False
    years: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def primary_scope(self) -> str:
        return self.scopes[0] if self.scopes else SCOPE_SUMMARY

    def has(self, scope: str) -> bool:
        return scope in self.scopes

    @property
    def scope_labels(self) -> list[str]:
        return [SCOPE_FA.get(scope, scope) for scope in self.scopes]

    def needs_generic_trend_sheet(self) -> bool:
        """The per-day generic trend is only shown when explicitly asked."""

        return self.has(SCOPE_TREND) or self.has(SCOPE_CORRELATION)

    def wants(self, sheet_key: str) -> bool:
        return sheet_key == "audit" and self.wants_audit

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "normalized": self.normalized,
            "scopes": list(self.scopes),
            "scope_labels": self.scope_labels,
            "requested_metrics": list(self.requested_metrics),
            "matched_metrics": list(self.matched_metrics),
            "period": self.period,
            "aggregation": self.aggregation,
            "wants_table_only": self.wants_table_only,
            "wants_chart": self.wants_chart,
            "wants_dashboard": self.wants_dashboard,
            "wants_management": self.wants_management,
            "wants_audit": self.wants_audit,
            "ranking_wants_least": self.ranking_wants_least,
            "years": list(self.years),
            "notes": list(self.notes),
        }


class RequestAnalyzer:
    """Classify a Persian request into scopes, target metrics and output shape."""

    def __init__(self) -> None:
        self._scope_order = [
            SCOPE_DISCREPANCY,
            SCOPE_QUALITY,
            SCOPE_RANKING,
            SCOPE_COMPARISON,
            SCOPE_TREND,
            SCOPE_GROUP,
            SCOPE_CORRELATION,
            SCOPE_DISTRIBUTION,
        ]

    # ------------------------------------------------------------------ #
    def parse(self, text: str, available_metrics: Iterable[str] = ()) -> RequestIntent:
        raw = (text or "").strip()
        normalized = self._normalize(raw)
        intent = RequestIntent(text=raw, normalized=normalized)
        if not normalized:
            intent.notes.append("درخواست خالی است؛ تحلیل استاندارد و کامل اجرا می‌شود.")
            return intent

        found = [
            scope
            for scope in self._scope_order
            if any(self._contains(normalized, word) for word in _PATTERNS[scope])
        ]
        intent.scopes = found or [SCOPE_SUMMARY]

        for word, period in _PERIOD_WORDS.items():
            if self._contains(normalized, word):
                intent.period = period
                break
        for word, aggregation in _AGGREGATION_WORDS.items():
            if self._contains(normalized, word):
                intent.aggregation = aggregation
                break

        for key, words in _OUTPUT_HINTS.items():
            if not any(self._contains(normalized, word) for word in words):
                continue
            if key == "table_only":
                intent.wants_table_only = True
            elif key == "dashboard":
                intent.wants_dashboard = True
            elif key == "chart":
                intent.wants_chart = True
            elif key == "management":
                intent.wants_management = True
            elif key == "audit":
                intent.wants_audit = True
            elif key == "report":
                intent.wants_management = True

        if any(self._contains(normalized, word) for word in ("کمترین", "پایین‌ترین", "بدترین")):
            intent.ranking_wants_least = True

        intent.years = self._extract_years(normalized)

        intent.requested_metrics = self._extract_metric_phrases(normalized)
        intent.matched_metrics = self.bind_metrics(intent.requested_metrics, available_metrics)
        return intent

    # ------------------------------------------------------------------ #
    def bind_metrics(self, phrases: Iterable[str], available: Iterable[str]) -> list[str]:
        """Match requested phrases to real column headers (never invent one)."""

        headers = self._unique([self._normalize(header) for header in available if header])
        if not headers:
            return []
        matched: list[str] = []
        for phrase in phrases:
            score, best = self._best_header(phrase, headers)
            if best is not None and score >= 0.5:
                matched.append(best)
        return self._unique(matched)

    def resolve_metrics(self, intent: RequestIntent, available: Iterable[str]) -> None:
        headers = [self._normalize(header) for header in available if header]
        if intent.requested_metrics:
            intent.matched_metrics = self.bind_metrics(intent.requested_metrics, headers)
            if not intent.matched_metrics:
                intent.notes.append(
                    "شاخص خواسته‌شده در ستون‌های فایل پیدا نشد؛ تحلیل روی همه شاخص‌ها انجام شد."
                )
        elif headers:
            intent.matched_metrics = []

    # ------------------------------------------------------------------ #
    @staticmethod
    def _normalize(text: str) -> str:
        value = normalize_text(normalize_digits(str(text)))
        value = re.sub(r"[«»\"'`()\[\]{}،؛,.:!?؟…]+", " ", value)
        value = re.sub(r"\s+", " ", value)
        return value.strip()

    @staticmethod
    def _contains(haystack: str, needle: str) -> bool:
        needle = needle.strip()
        if not needle:
            return False
        if " " in needle:
            return needle in haystack
        return re.search(rf"(?<!\w){re.escape(needle)}(?!\w)", haystack) is not None

    @staticmethod
    def _unique(values: Iterable[str]) -> list[str]:
        seen: set[str] = set()
        result: list[str] = []
        for value in values:
            if value and value not in seen:
                seen.add(value)
                result.append(value)
        return result

    @staticmethod
    def _tokens(text: str) -> list[str]:
        tokens = [token for token in text.split(" ") if token]
        return [token for token in tokens if token not in _STOPWORDS and len(token) > 1]

    @staticmethod
    def _extract_years(normalized: str) -> list[str]:
        """Jalali years mentioned in the request, in numeric order."""

        found = {
            match
            for match in re.findall(r"(?<!\d)(1[34]\d{2})(?!\d)", normalized)
        }
        return sorted(found)

    def _extract_metric_phrases(self, normalized: str) -> list[str]:
        """Pull noun phrases out of the request without inventing meaning."""

        phrases: list[str] = []
        # A metric is usually introduced by a quantity word.
        for match in re.finditer(
            r"(?:تعداد|مقدار|مجموع|جمع|میانگین|شاخص|میزان|درصد)\s+([^،؛.]{2,60})", normalized
        ):
            candidate = match.group(0).strip()
            candidate = re.split(
                r"\s+(?:را|رو|در|برای|بین|از|به|با|طی|و)\s+", candidate, maxsplit=1
            )[0]
            phrases.append(candidate.strip())
        if not phrases:
            tokens = self._tokens(normalized)
            if len(tokens) >= 1:
                phrases.append(" ".join(tokens[:4]))
        return self._unique(phrases)

    def _best_header(self, phrase: str, headers: list[str]) -> tuple[float, str | None]:
        tokens = self._tokens(phrase)
        best_score = 0.0
        best_header: str | None = None
        for header in headers:
            if header in phrase or phrase in header:
                score = 1.0
            else:
                header_tokens = self._tokens(header)
                if not header_tokens or not tokens:
                    continue
                shared = len(set(tokens) & set(header_tokens))
                score = shared / max(len(header_tokens), 1)
            if score > best_score:
                best_score, best_header = score, header
        return best_score, best_header

    # ------------------------------------------------------------------ #
    def sheet_plan(self, intent: RequestIntent, available: dict[str, bool]) -> list[str]:
        """Pick the sheets a clean delivery needs for this request.

        ``available`` maps a sheet key to whether the data supports it.  The
        order of the result is the order the sheets are written in.
        """

        if intent.wants_table_only:
            plan = ["request", "monthly", "comparison"]
            return [key for key in plan if available.get(key) or key == "request"]

        plan = ["request", "dashboard", "management"]
        for key in ("monthly", "comparison", "descriptive", "frequency", "groups"):
            if available.get(key):
                plan.append(key)
        if available.get("trend") and (intent.needs_generic_trend_sheet() or not available.get("monthly")):
            plan.append("trend")
        if available.get("quality"):
            plan.append("quality")
        if available.get("issues"):
            plan.append("issues")
        if intent.wants_audit and available.get("audit"):
            plan.append("audit")
        return self._unique(plan)


def describe_plan(plan: Iterable[str]) -> str:
    """Persian one-line description of the chosen sheets."""

    labels = [_SHEET_LANGUAGE.get(key, key) for key in plan]
    return "، ".join(labels)
