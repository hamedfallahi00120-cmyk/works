"""Answer the user's question with numbers Python already computed.

This is the grounded layer behind the "پاسخ درخواست" sheet.  Every answer is
built from :class:`PeriodIntegration` or from the generic column profile, keeps
the cell it came from, and never accepts a number from a language model.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Iterable

from local_office_ai.ai.request_analyzer import (
    SCOPE_COMPARISON,
    SCOPE_DISCREPANCY,
    SCOPE_QUALITY,
    SCOPE_RANKING,
    SCOPE_TREND,
    RequestIntent,
)
from local_office_ai.excel.normalization import normalize_text
from local_office_ai.excel.period_integration import (
    MetricSeries,
    PeriodIntegration,
    PeriodTotal,
)

MAX_METRICS = 5


@dataclass
class RequestAnswer:
    kind: str
    title: str
    text: str
    value: str = ""
    source: str = ""
    detail: str = ""
    verified: bool | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "kind": self.kind,
            "title": self.title,
            "text": self.text,
            "value": self.value,
            "source": self.source,
            "detail": self.detail,
            "verified": self.verified,
        }


def answer_request(
    intent: RequestIntent,
    analysis: Any,
    integration: PeriodIntegration | None,
) -> list[RequestAnswer]:
    answers: list[RequestAnswer] = []
    answers.append(_scope_answer(intent, analysis, integration))

    intelligence = getattr(analysis, "intelligence", None)
    conflicts = list(getattr(intelligence, "metadata_conflicts", []) or [])
    if conflicts:
        answers.append(
            RequestAnswer(
                kind="integrity",
                title="ناسازگاری برچسب سال در فایل",
                text=(
                    f"{len(conflicts)} برگه نام سال و عنوان سال ناهم‌خوان دارند: "
                    + "، ".join(
                        f"{item.sheet} (نام برگه {_years(_years_of(item, 'sheet'))}، "
                        f"عنوان {_years(_years_of(item, 'title'))})"
                        for item in conflicts[:4]
                    )
                    + ". هیچ منبعی بر دیگری ترجیح داده نشد؛ دوره‌ها همان‌طور که در سرستون سال "
                    "نوشته شده تحلیل شدند و مقایسه فقط روی برگه‌هایی انجام شد که هر دو سال "
                    "درخواستی را دارند."
                ),
                value=f"{len(conflicts)} برگه",
                source="مقایسه نام برگه، عنوان اصلی و سرستون سال",
                verified=True,
            )
        )

    if integration is not None and integration.is_usable:
        metrics = select_metrics(intent, integration)
        period = ranking_period(intent, integration)
        if intent.has(SCOPE_RANKING) or not intent.scopes or intent.primary_scope == "summary":
            answers.extend(_ranking_answers(metrics, period))
        if intent.has(SCOPE_COMPARISON) or intent.primary_scope == "summary":
            answers.extend(_comparison_answers(metrics, integration, intent))
        if intent.has(SCOPE_TREND):
            answers.extend(_trend_answers(metrics, period))
        if intent.primary_scope == "summary" and not intent.has(SCOPE_COMPARISON):
            answers.extend(_total_answers(metrics, period))
        answers.extend(_period_coverage_answers(integration, metrics, period))
    elif analysis is not None:
        answers.extend(_generic_answers(analysis))

    if intent.has(SCOPE_QUALITY):
        answers.extend(_quality_answers(analysis))
    if intent.has(SCOPE_DISCREPANCY) or intent.primary_scope == "summary":
        answers.extend(_integrity_answers(analysis, integration))
    return answers


# --------------------------------------------------------------------------- #
# Selection helpers
# --------------------------------------------------------------------------- #
def ranking_period(intent: RequestIntent, integration: PeriodIntegration) -> str:
    """The period the user asked about, else the best-covered period."""

    requested = [year for year in intent.years if year in integration.periods]
    if requested:
        return max(requested, key=lambda year: (len(integration.period_months(year)), year))
    return integration.main_period()


def comparison_pair(intent: RequestIntent, integration: PeriodIntegration) -> tuple[str, str] | None:
    """Two comparable periods; years named in the request win."""

    requested = [year for year in intent.years if year in integration.periods]
    if len(requested) >= 2:
        ordered = sorted(requested)
        pair = (ordered[0], ordered[1])
        if integration.series and any(
            series.compare(pair[0], pair[1]).months_used for series in integration.series
        ):
            return pair
    periods = integration.comparison_periods()
    if len(periods) < 2:
        return None
    return periods[-2], periods[-1]


def select_metrics(intent: RequestIntent, integration: PeriodIntegration) -> list[MetricSeries]:
    if intent.matched_metrics:
        wanted = [normalize_text(name) for name in intent.matched_metrics]
        chosen: list[MetricSeries] = []
        for item in integration.series:
            normal = normalize_text(item.metric)
            if any(name == normal or name in normal or normal in name for name in wanted):
                chosen.append(item)
        if chosen:
            return chosen[:MAX_METRICS]
    headline = integration.headline_metric()
    ordered = list(integration.series)
    if headline is not None and headline in ordered:
        ordered.remove(headline)
        ordered.insert(0, headline)
    return ordered[:MAX_METRICS]


# --------------------------------------------------------------------------- #
# Answers
# --------------------------------------------------------------------------- #
def _scope_answer(
    intent: RequestIntent, analysis: Any, integration: PeriodIntegration | None
) -> RequestAnswer:
    parts: list[str] = []
    if analysis is not None:
        meta = getattr(analysis, "meta", None)
        if meta is not None:
            parts.append(f"{len(meta.sheets)} برگه")
    if integration is not None:
        if integration.months:
            parts.append(f"{len(integration.months)} ماه")
        period_labels = [
            f"{period} ({len(integration.period_months(period))} ماه)"
            for period in integration.periods
        ]
        if period_labels:
            parts.append(f"{len(integration.periods)} دوره: {'، '.join(period_labels)}")
        if integration.series:
            parts.append(f"{len(integration.series)} شاخص عددی")
    scope = "، ".join(intent.scope_labels) or "خلاصه و تحلیل کلی"
    text = f"دامنه تحلیل: {'، '.join(parts)}. نوع تحلیل: {scope}."
    if intent.matched_metrics:
        text += " شاخص‌های هدف: " + "، ".join(intent.matched_metrics) + "."
    monthly = len(integration.months) if integration is not None else 0
    return RequestAnswer(
        kind="scope",
        title="دامنه تحلیل و تفسیر درخواست",
        text=text,
        value=f"{monthly} ماه" if monthly else "دوره ماهانه شناسایی نشد",
        source="ساختار فایل و ستون‌های شناسایی‌شده",
    )


def _ranking_answers(metrics: Iterable[MetricSeries], period: str) -> list[RequestAnswer]:
    answers: list[RequestAnswer] = []
    for item in metrics:
        best = item.best(period)
        worst = item.worst(period)
        if best is None or worst is None or best.month == worst.month:
            continue
        gap = (best.total or Decimal(0)) - (worst.total or Decimal(0))
        text = (
            f"در دوره {period}، بیشترین «{item.metric}» در ماه {best.month} با "
            f"{_fmt(best.total)} و کمترین در ماه {worst.month} با {_fmt(worst.total)} ثبت شده است؛ "
            f"اختلاف دو ماه {_fmt(gap)} است."
        )
        excluded_cells = [cell for cell in item.cells if cell.period == period and not cell.complete]
        detail = (
            "ماه ناقص از رتبه‌بندی کنار گذاشته شد: "
            + "، ".join(f"{cell.month} ({cell.observations} روز ثبت‌شده)" for cell in excluded_cells)
            if excluded_cells
            else "همه ماه‌های این دوره کامل بودند؛ رتبه‌بندی روی همه آن‌ها انجام شد."
        )
        answers.append(
            RequestAnswer(
                kind="ranking",
                title=f"بیشترین و کمترین ماه | {item.metric}",
                text=text,
                value=f"{best.month}: {_fmt(best.total)}",
                source=_source(best) + "؛ " + _source(worst),
                detail=detail,
                verified=True,
            )
        )
    return answers


def _comparison_answers(
    metrics: Iterable[MetricSeries],
    integration: PeriodIntegration,
    intent: RequestIntent,
) -> list[RequestAnswer]:
    pair = comparison_pair(intent, integration)
    if pair is None:
        return []
    period_a, period_b = pair
    answers: list[RequestAnswer] = []
    for item in metrics:
        change = item.compare(period_a, period_b)
        if change.difference is None:
            continue
        percent = "قابل محاسبه نیست" if change.percent is None else f"{change.percent}%"
        text = (
            f"«{item.metric}» در ماه‌های مشترک ({'، '.join(change.months_used)}) دوره {period_b} "
            f"برابر {_fmt(change.total_b)} و در دوره {period_a} برابر {_fmt(change.total_a)} بوده است؛ "
            f"{change.direction} {_fmt(abs(change.difference))} واحد ({percent})."
        )
        answers.append(
            RequestAnswer(
                kind="comparison",
                title=f"مقایسه {period_a} و {period_b} | {item.metric}",
                text=text,
                value=f"{change.direction} {percent}",
                source=(
                    "جمع ستون‌های ماهانه ماه‌های مشترک: " + "، ".join(change.months_used)
                ),
                detail=change.note,
                verified=True,
            )
        )
    return answers


def _trend_answers(metrics: Iterable[MetricSeries], period: str) -> list[RequestAnswer]:
    answers: list[RequestAnswer] = []
    for item in metrics:
        changes = item.month_changes(period)
        if len(changes) < 2:
            continue
        first_month, first_value, _, _ = changes[0]
        last_month, last_value, _, _ = changes[-1]
        difference = last_value - first_value
        percent = (
            (difference / abs(first_value) * Decimal(100)).quantize(Decimal("0.01"))
            if first_value != 0
            else None
        )
        direction = "افزایش" if difference > 0 else ("کاهش" if difference < 0 else "بدون تغییر")
        text = (
            f"«{item.metric}» از ماه {first_month} ({_fmt(first_value)}) تا ماه {last_month} "
            f"({_fmt(last_value)}) {direction} {_fmt(abs(difference))} واحد"
            + (f" معادل {percent}%." if percent is not None else ".")
        )
        answers.append(
            RequestAnswer(
                kind="trend",
                title=f"روند ماهانه | {item.metric}",
                text=text,
                value=direction,
                source=f"جمع ماهانه دوره {period}",
                detail="مقایسه ماه‌به‌ماه توصیفی است و علت‌ومعلول را اثبات نمی‌کند.",
                verified=True,
            )
        )
    return answers


def _total_answers(metrics: Iterable[MetricSeries], period: str) -> list[RequestAnswer]:
    answers: list[RequestAnswer] = []
    for item in metrics:
        total = item.total(period)
        if total is None:
            continue
        excluded = item.excluded_incomplete(period)
        detail = ""
        if excluded:
            detail = "ماه ناقص در جمع لحاظ نشد: " + "، ".join(excluded) + "."
        text = (
            f"جمع «{item.metric}» در ماه‌های کامل دوره {period} برابر {_fmt(total)} است "
            f"({len(item.months_used(period))} ماه)."
        )
        answers.append(
            RequestAnswer(
                kind="total",
                title=f"جمع دوره {period} | {item.metric}",
                text=text,
                value=_fmt(total),
                source="جمع ماهانه ستون‌های هم‌تعریف",
                detail=detail,
                verified=True,
            )
        )
    return answers


def _period_coverage_answers(
    integration: PeriodIntegration, metrics: Iterable[MetricSeries], period: str
) -> list[RequestAnswer]:
    answers: list[RequestAnswer] = []
    if integration.alignment:
        total_cells = sum(item["cells"] for item in integration.alignment)
        sample = integration.alignment[0]
        text = (
            f"برگه «{sample['summary_sheet']}» جمع خود را برای سال {sample['summary_period']} از "
            f"ستون {sample['referenced_column']} برگه «{sample['referenced_sheet']}» می‌خواند که "
            f"سال {sample['referenced_period']} است. این ناهم‌ترازی در {total_cells} سلول دیده شد؛ "
            "بنابراین جمع برگه خلاصه با جمع ماهانه هم‌تعریف نیست و مبنای گزارش قرار نگرفت."
        )
        answers.append(
            RequestAnswer(
                kind="integrity",
                title="ناهم‌ترازی سال در برگه خلاصه فایل",
                text=text,
                value=f"{total_cells} سلول",
                source="فرمول‌های برگه خلاصه در برابر برچسب سال برگه‌های ماهانه",
                verified=True,
            )
        )
    if integration.suspicious_periods:
        details = "، ".join(
            f"{period} ({len(integration.period_months(period))} ماه)" for period in integration.suspicious_periods
        )
        answers.append(
            RequestAnswer(
                kind="integrity",
                title="دوره با برچسب مشکوک",
                text=(
                    f"ستون‌های با برچسب {details} در فایل فقط برای یک ماه مقدار دارند و با نام برگه‌ها "
                    "هم‌خوان نیستند؛ این دوره‌ها مبنای مقایسه قرار نگرفتند و مقدارشان جداگانه نشان داده شده است."
                ),
                value="، ".join(integration.suspicious_periods),
                source="برچسب سال در سرستون در برابر نام برگه و پوشش داده",
                verified=True,
            )
        )
    excluded = [item for item in metrics if item.excluded_incomplete(period)]
    if excluded:
        names = "، ".join(item.metric for item in excluded[:5])
        answers.append(
            RequestAnswer(
                kind="coverage",
                title="دوره ناقص",
                text=(
                    f"در دوره {period} ماه ناقص وجود دارد؛ شاخص‌های «{names}» در جمع دوره و رتبه‌بندی "
                    "لحاظ نشدند و مقدار عددی آن‌ها فقط به‌صورت ثبت‌شده نمایش داده می‌شود."
                ),
                value="ماه ناقص کنار گذاشته شد",
                source="پوشش روزهای هر برگه",
                verified=True,
            )
        )
    return answers


def _generic_answers(analysis: Any) -> list[RequestAnswer]:
    generic = getattr(analysis, "generic", None)
    if generic is None:
        return []
    answers: list[RequestAnswer] = []
    numeric = [
        column
        for column in generic.columns
        if column.is_numeric and not column.is_identifier and not column.is_temporal_key
    ]
    numeric.sort(key=lambda column: (len(str(column.header)), column.count), reverse=True)
    for column in numeric[:3]:
        if column.maximum is None or column.minimum is None:
            continue
        text = (
            f"ستون «{column.header}» در {column.count} ردیف، مجموع {_fmt(column.total)}، "
            f"میانگین {_fmt(column.mean)}، کمینه {_fmt(column.minimum)} و بیشینه {_fmt(column.maximum)} دارد."
        )
        answers.append(
            RequestAnswer(
                kind="descriptive",
                title=f"آمار توصیفی | {column.header}",
                text=text,
                value=_fmt(column.mean),
                source=f"{generic.table_sheet}!{generic.table_address} ستون {column.letter}",
                verified=True,
            )
        )
    return answers


def _quality_answers(analysis: Any) -> list[RequestAnswer]:
    answers: list[RequestAnswer] = []
    generic = getattr(analysis, "generic", None)
    if generic is not None and generic.quality is not None:
        quality = generic.quality
        ratio = quality.measured_missing_ratio if quality.empty_columns else quality.missing_ratio
        empty_columns = (
            f" ({len(quality.empty_columns)} ستون بدون هیچ مقدار در دوره‌ای خالی است و جدا شمرده شد.)"
            if quality.empty_columns
            else ""
        )
        error_cells = getattr(getattr(analysis, "meta", None), "total_error_cells", 0)
        text = (
            f"درصد سلول خالی {ratio:.1%}{empty_columns}، ردیف کاملاً تکراری "
            f"{quality.duplicate_rows}، سلول نامعتبر {quality.invalid_cells} و سلول خطای فرمول "
            f"{error_cells} است. هیچ مقدار خالی به صفر تبدیل نشد و هیچ ردیفی حذف نشد."
        )
        answers.append(
            RequestAnswer(
                kind="quality",
                title="کیفیت داده جدول اصلی",
                text=text,
                value=f"خالی {ratio:.1%}",
                source=f"{generic.table_sheet}!{generic.table_address}",
                verified=True,
            )
        )
    checks = getattr(analysis, "validation", None) or []
    failed = [check for check in checks if getattr(check, "status", "") == "failed"]
    if failed:
        text = "؛ ".join(f"{check.label}: {check.detail}" for check in failed[:4])
        answers.append(
            RequestAnswer(
                kind="quality",
                title="بررسی‌های رد‌شده کیفیت",
                text=text,
                value=f"{len(failed)} بررسی",
                source="چک‌لیست اعتبارسنجی داده",
                verified=True,
            )
        )
    return answers


def _integrity_answers(analysis: Any, integration: PeriodIntegration | None) -> list[RequestAnswer]:
    answers: list[RequestAnswer] = []
    intelligence = getattr(analysis, "intelligence", None)
    if intelligence is not None:
        material = [
            audit
            for audit in intelligence.all_audits
            if getattr(audit, "status", None) is not None
            and getattr(audit.status, "value", "") == "mismatch"
            and audit.kind in {"total", "difference"}
        ]
        percent_rows = [
            audit
            for audit in intelligence.all_audits
            if getattr(audit.status, "value", "") == "mismatch" and audit.kind == "percentage"
        ]
        if material:
            sample = material[0]
            text = (
                f"{len(material)} جمع/مقایسه اعلامی با محاسبه مستقل Python یکسان نیست؛ "
                f"نمونه: برگه «{sample.sheet}» سلول {sample.address} برای «{sample.metric}» "
                f"مقدار {_fmt(sample.declared)} در برابر {_fmt(sample.computed)}. "
                "هیچ مقدار خودکار اصلاح نشد."
            )
            answers.append(
                RequestAnswer(
                    kind="integrity",
                    title="مغایرت جمع اعلامی با محاسبه",
                    text=text,
                    value=f"{len(material)} مورد",
                    source=f"{sample.sheet}!{sample.address}",
                    verified=True,
                )
            )
        if percent_rows:
            answers.append(
                RequestAnswer(
                    kind="integrity",
                    title="ردیف درصد فایل",
                    text=(
                        f"{len(percent_rows)} سلول ردیف «درصد» مقدار ثابت دارد و با نسبت محاسبه‌شده "
                        "هم‌خوان نیست؛ این ردیف نمایشی است و در نتیجه‌گیری استفاده نشد."
                    ),
                    value=f"{len(percent_rows)} سلول",
                    source="ردیف درصد برگه‌های ماهانه",
                    verified=True,
                )
            )
    if integration is not None and integration.mismatched_cells:
        answers.append(
            RequestAnswer(
                kind="integrity",
                title="کنترل دوباره جمع ماهانه",
                text=(
                    f"{integration.mismatched_cells} ستون از {integration.verified_cells} ستون "
                    "با بازمحاسبه مستقل اختلاف داشت؛ عدد بازمحاسبه‌شده مبنا قرار گرفت."
                ),
                value=f"{integration.mismatched_cells} ستون",
                source="بازمحاسبه سلول‌به‌سلول",
                verified=True,
            )
        )
    return answers


# --------------------------------------------------------------------------- #
# Formatting
# --------------------------------------------------------------------------- #
def _years_of(conflict: Any, key: str) -> list[str]:
    values = getattr(conflict, "values", None) or {}
    for name in (f"{key}_years", key, f"{key}Year"):
        if name in values:
            return list(values[name] or [])
    return []


def _years(values: Any) -> str:
    years = [str(item) for item in (values or []) if str(item).strip()]
    return "، ".join(years) if years else "بدون سال"


def _fmt(value: Decimal | None) -> str:
    if value is None:
        return "—"
    if value == value.to_integral_value():
        return f"{int(value):,}"
    return f"{value.normalize():,f}"


def _source(cell: PeriodTotal) -> str:
    address = cell.total_address or "جمع ستون"
    return f"{cell.source}!{address}"
