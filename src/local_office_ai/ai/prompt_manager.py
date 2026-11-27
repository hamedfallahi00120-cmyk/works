"""Prompt construction, kept separate from calculation and from HTTP."""

from __future__ import annotations

import json
from decimal import Decimal
from typing import Any

SYSTEM_PROMPT = """تو یک تحلیل‌گر ارشد داده‌های اداری و آماری هستی که برای یک مدیر فارسی‌زبان گزارش می‌نویسد.
قواعد قطعی:
1) همه محاسبات قبلاً با Python انجام شده است. هیچ جمع، تفریق، ضرب، تقسیم، درصد، میانگین، واریانس یا همبستگی جدیدی محاسبه نکن.
2) فقط اعدادی را نقل کن که در بسته داده آمده‌اند. هیچ عدد تازه‌ای نساز و هیچ عددی را گرد یا تغییر نده.
3) همبستگی رابطه را نشان می‌دهد، نه علت‌ومعلول؛ این را صریح رعایت کن.
4) داده گمشده و داده پرت حذف نشده‌اند؛ اگر به آن‌ها اشاره می‌کنی، همان وضعیت ثبت‌شده را بگو.
5) مغایرت را فقط وقتی اعلام کن که در بخش audits یا reconciliation وضعیت mismatch/probable/outside ثبت شده باشد.
6) اگر داده‌ای برای یک نتیجه‌گیری وجود ندارد، آن را در missing_data بنویس و حدس نزن.
7) بخش وام/تعهد فقط یک ماژول اختیاری است و باید در صورت وجود داده، نه به‌عنوان هسته اصلی تحلیل، تفسیر شود.
پاسخ را فقط به صورت JSON معتبر و با این ساختار بده:
{"summary": "متن فارسی", "findings": ["..."], "missing_data": ["..."], "recommendations": ["..."], "risk_level": "low|medium|high"}
همه متن‌ها فارسی و کوتاه و قابل ارائه به مدیر باشند."""

TASKS: dict[str, str] = {
    "management_analysis": (
        "این مجموعه داده اداری/آماری را تحلیل مدیریتی کن: کیفیت داده، توزیع‌ها، گروه‌های مهم، روند، "
        "همبستگی‌های قابل توجه، داده‌های پرت و پیشنهادهای عملی. اگر بخش وام/تعهد در بسته هست، "
        "آن را فقط به‌عنوان یک دیدگاه اختیاری و با اعداد همان بسته توضیح بده."
    ),
    "general_analysis": (
        "یک تحلیل اکتشافی کامل بده: ساختار داده، کیفیت، آمار توصیفی، گروه‌ها، روند، همبستگی و پرت‌ها. "
        "روش‌ها و محدودیت‌ها را کوتاه توضیح بده و از نتیجه‌گیری علّی پرهیز کن."
    ),
    "data_quality": (
        "فقط کیفیت داده را تحلیل کن: سلول خالی، ردیف تکراری، ستون ثابت، نوع مختلط، شناسه‌ها و پرت‌ها. "
        "برای هر مورد اقدام پیشنهادی بده؛ هیچ ردیفی را حذف‌شده فرض نکن."
    ),
    "trend_analysis": (
        "روند زمانی را بر اساس دوره‌های موجود توضیح بده: جهت، تغییر میانگین و محدودیت‌های تعداد دوره. "
        "علت تغییر را بدون داده اضافی به قطعیت نسبت نده."
    ),
    "group_comparison": (
        "گروه‌های اصلی را با هم مقایسه کن: تفاوت میانگین و میانه، تعداد هر گروه و پراکندگی. "
        "اگر گروه کوچک است، آن را به‌عنوان محدودیت ذکر کن."
    ),
    "correlation_analysis": (
        "همبستگی‌های گزارش‌شده را تفسیر کن: جهت، قدرت، تعداد جفت و تفاوت پیرسون/اسپیرمن. "
        "صریح بگو همبستگی علت‌ومعلول نیست."
    ),
    "explain_discrepancies": (
        "فقط مغایرت‌ها و تطبیق‌های نیازمند تأیید را توضیح بده: طرفین، اختلاف و اقدام پیشنهادی. "
        "اگر مغایرتی ثبت نشده، صریح بنویس."
    ),
    "column_meaning": (
        "معنای هر ستون جدول اصلی را بر اساس عنوان، نوع داده و نمونه مقادیر توضیح بده. "
        "بگو هر ستون برای چه نوع تحلیلی مناسب است و کجا محدودیت دارد. هیچ عدد تازه‌ای نساز."
    ),
    "executive_summary": (
        "یک خلاصه اجرایی حداکثر ۱۲۰ کلمه‌ای برای مدیر بنویس: مهم‌ترین یافته‌های داده، "
        "کیفیت داده و مهم‌ترین اقدام بعدی."
    ),
}


def build_messages(
    facts: dict[str, Any],
    *,
    task: str = "general_analysis",
    question: str | None = None,
    extra_rules: list[str] | None = None,
    chunk_label: str | None = None,
) -> list[dict[str, str]]:
    task_text = TASKS.get(task, TASKS["general_analysis"])
    user_parts = [task_text]
    if chunk_label:
        user_parts.append("بخش درخواست: " + chunk_label)
    if extra_rules:
        user_parts.append("قواعد تکمیلی: " + " | ".join(extra_rules))
    if question:
        user_parts.append("پرسش کاربر: " + question)
    user_parts.append("بسته داده (JSON):")
    user_parts.append(json.dumps(facts, ensure_ascii=False))
    return [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": "\n\n".join(user_parts)},
    ]


def _offline_generic_summary(facts: dict[str, Any], dataset: dict[str, Any]) -> dict[str, Any]:
    def number(value: Any) -> str:
        try:
            return f"{Decimal(str(value)):,.2f}".rstrip("0").rstrip(".")
        except (TypeError, ValueError, ArithmeticError):
            return str(value)

    quality = dataset.get("quality", {})
    table = dataset.get("table", {})
    rows = table.get("rows", 0)
    cols = table.get("columns", 0)
    missing = (
        quality.get("measured_missing_ratio")
        if quality.get("empty_columns")
        else quality.get("missing_ratio", 0)
    ) or 0
    duplicates = quality.get("duplicate_rows", 0)
    outliers = dataset.get("outliers", {})
    outlier_count = outliers.get("count", 0)
    descriptive = dataset.get("descriptive", [])
    groups = dataset.get("groups", [])
    trends = dataset.get("trends", [])
    correlations = dataset.get("correlations", [])
    columns = dataset.get("columns", [])

    workbook = facts.get("workbook_intelligence") or {}
    summary = []
    if workbook:
        counts = workbook.get("status_counts", {})
        audit_counts = workbook.get("audit_counts", {})
        summary.append(
            f"Workbook {workbook.get('sheet_count', 0)} برگه دارد: "
            f"{counts.get('fully_analyzed', 0)} برگه کامل، "
            f"{counts.get('partially_analyzed', 0)} برگه جزئی و "
            f"{counts.get('not_analyzed', 0)} برگه بدون تحلیل."
        )
        if audit_counts.get("mismatch"):
            summary.append(
                f"{audit_counts['mismatch']} مغایرت بین جمع/مقایسه اعلامی و محاسبه مستقل پیدا شد."
            )
    summary.extend([
        f"جدول اصلی {rows} ردیف و {cols} ستون دارد.",
        f"درصد سلول‌های خالی {float(missing):.1%} و ردیف‌های کاملاً تکراری {duplicates} است.",
    ])
    period = facts.get("period_analysis") or {}
    if period:
        months = period.get("months") or []
        periods = period.get("periods") or []
        summary.append(
            f"تحلیل ماهانه روی {len(months)} ماه ({'، '.join(months)}) و دوره‌های "
            f"{'، '.join(periods)} انجام شد؛ جمع هر ماه از سلول‌های همان برگه بازمحاسبه شده است."
        )
    numeric_columns = [column for column in columns if column.get("type") == "number"]
    if numeric_columns:
        picks = descriptive[:3]
        bits = []
        for column in picks:
            if column.get("mean") is not None:
                bits.append(f"میانگین «{column['header']}» {number(column['mean'])}")
        if bits:
            summary.append("؛ ".join(bits) + ".")
    if trends:
        trend = trends[0]
        summary.append(
            f"روند «{trend.get('measure')}» بر اساس «{trend.get('date')}» {trend.get('direction')} است"
            + (f" و تغییر میانگین دوره‌ها {number(trend.get('change_percent'))} درصد است." if trend.get("change_percent") is not None else ".")
        )
    if groups:
        group = groups[0]
        values = group.get("values", [])
        if values:
            top = max(values, key=lambda item: Decimal(str(item.get("mean") or "0")))
            summary.append(
                f"در مقایسه «{group.get('measure')}» بر اساس «{group.get('group')}»، "
                f"گروه «{top.get('label')}» بیشترین میانگین را دارد."
            )
    if outlier_count:
        summary.append(f"{outlier_count} مقدار پرت علامت‌گذاری شده است؛ هیچ مقداری حذف نشده است.")

    findings: list[str] = []
    for column in descriptive[:4]:
        if column.get("mean") is not None and column.get("median") is not None:
            findings.append(
                f"«{column['header']}»: میانگین {number(column['mean'])}، میانه {number(column['median'])}، "
                f"انحراف معیار نمونه {number(column.get('std_sample'))}."
            )
    if groups:
        group = groups[0]
        values = group.get("values", [])
        if len(values) >= 2:
            low = min(values, key=lambda item: Decimal(str(item.get("mean") or "0")))
            high = max(values, key=lambda item: Decimal(str(item.get("mean") or "0")))
            findings.append(
                f"تفاوت میانگین «{group.get('measure')}» بین «{low.get('label')}» و «{high.get('label')}» "
                "در برگه تحلیل گروه‌ها ثبت شده است."
            )
    for pair in correlations[:3]:
        if pair.get("pearson") is not None:
            findings.append(
                f"همبستگی «{pair.get('a')}» و «{pair.get('b')}» {pair.get('strength')} است "
                f"(r={number(pair.get('pearson'))}, n={pair.get('n')})؛ این رابطه علت‌ومعلولی نیست."
            )
    if period:
        main_period = period.get("main_period") or (period.get("periods") or [""])[-1]
        for entry in (period.get("metrics") or [])[:4]:
            best = (entry.get("best") or {}).get(main_period)
            worst = (entry.get("worst") or {}).get(main_period)
            if best and worst:
                findings.append(
                    f"«{entry['metric']}» در دوره {main_period}: بیشترین ماه {best['month']} با "
                    f"{number(best['total'])} و کمترین ماه {worst['month']} با {number(worst['total'])}."
                )
        for item in (period.get("comparison") or [])[:3]:
            if item.get("percent") is None:
                continue
            findings.append(
                f"مقایسه «{item['metric']}» در ماه‌های مشترک ({'، '.join(item.get('months_used') or [])}): "
                f"دوره {item['period_b']} {number(item.get('total_b'))} در برابر دوره {item['period_a']} "
                f"{number(item.get('total_a'))}؛ {item.get('direction')} {number(item.get('percent'))} درصد."
            )
        if period.get("suspicious_periods"):
            findings.append(
                "دوره با برچسب مشکوک (کمتر از نیمی از ماه‌ها داده دارد): "
                + "، ".join(period["suspicious_periods"])
                + "؛ مبنای مقایسه قرار نگرفت."
            )
    if duplicates:
        findings.append(f"{duplicates} ردیف کاملاً تکراری وجود دارد؛ حذف خودکار انجام نشده است.")
    if quality.get("identifier_columns"):
        findings.append("ستون‌های شناسه: " + "، ".join(quality["identifier_columns"]) + ".")

    recommendations: list[str] = []
    if missing >= 0.1:
        recommendations.append("الگوی مقادیر خالی بررسی و روش برخورد با آن مستند شود.")
    if duplicates:
        recommendations.append("ردیف‌های تکراری به‌صورت دستی بررسی شوند؛ حذف خودکار مجاز نیست.")
    if outlier_count:
        recommendations.append("مقادیر پرت با مقدار اصلی هر سلول مقایسه شوند؛ پرت بودن دلیل خطا نیست.")
    if not trends and not period:
        recommendations.append("برای تحلیل روند، یک ستون تاریخ معتبر به داده اضافه یا اصلاح شود.")
    recommendations.append("برای نتیجه‌گیری استنباطی، آزمون آماری متناسب با نوع داده و طرح نمونه‌گیری انتخاب شود.")

    workbook = facts.get("workbook_intelligence") or {}
    audit_counts = workbook.get("audit_counts", {})
    for item in (workbook.get("audit_mismatches") or [])[:4]:
        findings.append(
            f"مغایرت «{item.get('metric')}» در {item.get('sheet')}!{item.get('address')}: "
            f"ثبت‌شده {number(item.get('declared'))} در برابر محاسبه {number(item.get('computed'))}."
        )
    for item in (workbook.get("metadata_conflicts") or [])[:3]:
        findings.append(f"تعارض متادیتا در «{item.get('sheet')}»: {item.get('message')}")
    if workbook.get("audit_counts", {}).get("mismatch"):
        recommendations.append("مغایرت‌های جمع و درصد به‌صورت دستی با فرمول اصلی فایل مقایسه شوند؛ اصلاح خودکار انجام نشود.")
    if workbook.get("metadata_conflicts"):
        recommendations.append("تعارض سال نام برگه/عنوان/زیرعنوان پیش از هر مقایسه دوره‌ای تأیید شود.")
    incomplete = [
        sheet for sheet in (workbook.get("sheets") or [])
        if (sheet.get("coverage") or {}).get("incomplete")
    ]
    if incomplete:
        missing_data = [
            f"{sheet['sheet']}: دوره ناقص ({sheet['coverage'].get('observed')} از {sheet['coverage'].get('expected')})"
            for sheet in incomplete[:6]
        ]
    else:
        missing_data = list(facts.get("data_gaps", []))

    risk = "high" if audit_counts.get("mismatch") else "medium" if (missing >= 0.1 or duplicates or outlier_count) else "low"
    return {
        "summary": " ".join(summary),
        "findings": findings[:10],
        "missing_data": missing_data[:8],
        "recommendations": recommendations[:6],
        "risk_level": risk,
        "mode": "offline",
    }


def offline_summary(facts: dict[str, Any]) -> dict[str, Any]:
    """Deterministic Persian summary used when no model is available.

    Generic datasets get a data-quality and statistical narrative.  The
    optional loan module still gets its own factual wording when loan figures
    are present, so Book2 compatibility is preserved without making finance the
    core of the engine.
    """

    def money(value: Any) -> str:
        try:
            return f"{int(Decimal(str(value))):,}"
        except (TypeError, ValueError, ArithmeticError):
            return str(value)

    dataset = facts.get("dataset_profile")
    if dataset and not facts.get("loan_institutions"):
        return _offline_generic_summary(facts, dataset)

    calculations = facts.get("verified_calculations_from_python", {})
    audits = facts.get("audits_of_declared_totals", [])
    reconciliation = facts.get("reconciliation", [])
    anomalies = facts.get("important_anomalies", [])
    counts = facts.get("reconciliation_counts", {})
    lenders = facts.get("loan_institutions", [])

    summary_bits: list[str] = []
    count = calculations.get("loan_count", {}).get("value")
    principal = calculations.get("principal_sum", {}).get("value")
    installments = calculations.get("installments_sum", {}).get("value")
    obligations = calculations.get("obligations_sum", {}).get("value")
    repayment = calculations.get("contract_repayment_sum", {}).get("value")

    if count is not None:
        lender_text = f" از {len(lenders)} مؤسسه" if lenders else ""
        summary_bits.append(f"در این فایل {count} وام{lender_text} شناسایی شد.")
    if principal is not None and installments is not None:
        summary_bits.append(
            f"مجموع اصل وام‌ها {money(principal)} تومان و مجموع اقساط ماهانه {money(installments)} تومان است."
        )
    if obligations is not None:
        summary_bits.append(
            f"مجموع تعهدهای ماهانه ثبت‌شده خارج از جدول وام‌ها {money(obligations)} تومان است."
        )
    if repayment is not None:
        summary_bits.append(
            f"کل بازپرداخت قراردادی وام‌ها بر مبنای مبلغ هر قسط × تعداد اقساط {money(repayment)} تومان محاسبه شد."
        )
    if counts:
        summary_bits.append(
            "نتیجه تطبیق: "
            f"{counts.get('matched', 0)} مؤسسه مطابقت کامل، "
            f"{counts.get('probable', 0)} تطبیق احتمالی نیازمند تأیید و "
            f"{counts.get('outside_loan_table', 0)} تعهد خارج از جدول وام‌ها شناسایی شد."
        )

    mismatches = [audit for audit in audits if audit.get("status") == "mismatch"]
    if audits:
        summary_bits.append(
            "حسابرسی اعداد اعلام‌شده فایل: همه ارقام با محاسبات مستقل Python تأیید شد."
            if not mismatches
            else f"در حسابرسی اعداد اعلام‌شده {len(mismatches)} مغایرت ثبت شد؛ جزئیات در برگه «حسابرسی اعداد» است."
        )
    if anomalies:
        summary_bits.append(
            f"در مجموع {len(anomalies)} مورد نیازمند بررسی ثبت شد؛ فهرست کامل در برگه «موارد نیازمند بررسی» است."
        )

    findings: list[str] = []
    covered = {"total-mismatch", "probable-match", "obligation-outside-table"}
    for audit in audits:
        if audit.get("status") == "mismatch":
            findings.append(f"مغایرت {audit.get('label')} با اختلاف {money(audit.get('difference'))}.")
    for item in reconciliation:
        if item.get("status") == "outside_loan_table":
            findings.append(
                f"تعهد «{item.get('institution')}» به مبلغ {money(item.get('obligation_total'))} تومان "
                "در جدول وام‌ها ثبت نشده است و باید بررسی شود."
            )
        elif item.get("status") == "probable":
            findings.append(f"تطبیق «{item.get('institution')}» نیازمند تأیید است.")
    for anomaly in anomalies:
        if anomaly.get("category") in covered:
            continue
        message = anomaly.get("message", "")
        if message and message not in findings:
            findings.append(message)
        if len(findings) >= 8:
            break

    return {
        "summary": " ".join(summary_bits),
        "findings": findings,
        "missing_data": list(facts.get("data_gaps", [])),
        "recommendations": [
            "ردیف‌های نیازمند بررسی در برگه «موارد نیازمند بررسی» مرور شود.",
            "برای اطمینان، جمع‌های اعلامی فایل با دفاتر رسمی مطابقت داده شود.",
        ],
        "risk_level": "medium" if mismatches or anomalies else "low",
        "mode": "offline",
    }
