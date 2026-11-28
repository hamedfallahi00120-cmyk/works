"""Build one validated Word report from a multi-file folder analysis.

The single-workbook builder in :mod:`reporting.specification` cannot describe a
folder: several files, cross-source matching, duplicate decisions, per-file
quality findings and a comparison that is only allowed when the evidence
supports it.  This module assembles that content from values Python already
calculated.  It never reads a workbook again and never lets a template become a
data source.

Every table cell is display-ready, so the DOCX validator can confirm that each
number in the document is one of the approved report values.
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from typing import Any

from local_office_ai.excel.indicator_periods import STATUS_COMPARED
from local_office_ai.reporting.specification import (
    ReportContent,
    ReportRow,
    ReportSection,
    ReportTable,
)


def _text(value: Any) -> str:
    if value is None:
        return "—"
    if isinstance(value, Decimal):
        return format(value, "f")
    text = str(value).strip()
    return text or "—"


def _number(value: Any) -> str:
    return "" if value is None else _text(value)


def _status(value: Any) -> str:
    if value is None:
        return "—"
    label = getattr(value, "label_fa", None)
    if label:
        return str(label)
    return str(getattr(value, "value", value))


def _percent(value: Any) -> str:
    if value is None:
        return "—"
    return f"{_text(value)}٪"


def _name(item: Any) -> str:
    return str(getattr(item, "name", "") or getattr(item, "path", "") or "—")


def _sheet_names(analysis: Any) -> list[str]:
    meta = getattr(analysis, "meta", None)
    sheets = list(getattr(meta, "sheets", []) or []) if meta is not None else []
    return [str(getattr(sheet, "title", "") or "") for sheet in sheets if getattr(sheet, "title", "")]


def _blocks(analysis: Any) -> list[Any]:
    extraction = getattr(analysis, "extraction", None)
    blocks: list[Any] = []
    for sheet in getattr(extraction, "sheets", []) or []:
        blocks.extend(list(getattr(sheet, "blocks", []) or []))
    return blocks


def _gate_summary(gate: Any) -> dict[str, int]:
    if gate is None:
        return {}
    counts: dict[str, int] = {}
    for stage in getattr(gate, "stages", []) or []:
        status = _status(getattr(stage, "status", "")).upper()
        counts[status] = counts.get(status, 0) + 1
    return counts


def build_folder_content(
    result: Any,
    *,
    title: str | None = None,
    request: str = "",
    generated_at: date | datetime | None = None,
) -> ReportContent:
    """Assemble validated content for a multi-source Word report."""

    stamp = generated_at or datetime.now()
    stamp_text = stamp.strftime("%Y-%m-%d %H:%M") if isinstance(stamp, datetime) else stamp.isoformat()
    files = list(getattr(result, "files", []) or [])
    analyses = [item for item in files if getattr(item, "analysis", None) is not None]
    matching = getattr(result, "matching", None)
    duplicate = getattr(result, "duplicate_control", None)
    indicators = getattr(result, "indicator_periods", None)
    accidents = getattr(result, "accident_semantics", None)
    gate = getattr(result, "quality_gate", None)

    report_title = title or "گزارش تحلیلی چندمنبعی"
    analyses_count = len(analyses)
    failed = [item for item in files if getattr(item, "status", "") != "تحلیل شد"]
    all_blocks = [block for item in analyses for block in _blocks(item.analysis)]
    periods = sorted({period for item in analyses for period in getattr(item.analysis.integration, "periods", []) or []})
    if not periods and indicators is not None:
        periods = list(getattr(indicators, "periods", []) or [])
    matching_counts = matching.counts() if matching is not None and hasattr(matching, "counts") else {}
    duplicate_counts = duplicate.counts() if duplicate is not None and hasattr(duplicate, "counts") else {}
    provenance = getattr(gate, "provenance", None) if gate is not None else None
    provenance_counts = provenance.counts() if provenance is not None and hasattr(provenance, "counts") else {}

    values: dict[str, Any] = {
        "report.title": (report_title, "عنوان گزارش", "report", ""),
        "report.generated_at": (stamp_text, "تاریخ گزارش", "report", ""),
        "report.source_file": (
            "، ".join(_name(item) for item in files) or "—",
            "فایل‌های منبع",
            "folder",
            "",
        ),
        "report.request": (request or getattr(result, "request", "") or "—", "درخواست کاربر", "request", ""),
        "files.total": (str(len(files)), "تعداد فایل منبع", "folder", ""),
        "files.analyzed": (str(analyses_count), "فایل‌های تحلیل‌شده", "folder", ""),
        "files.failed": (str(len(failed)), "فایل‌های ناموفق", "folder", ""),
        "blocks.detected": (str(len(all_blocks)), "بلوک‌های استخراج‌شده", "extraction", ""),
        "periods.count": (str(len(periods)), "دوره‌های شناسایی‌شده", "periods", "، ".join(periods)),
        "matching.matched": (str(matching_counts.get("matched", 0)), "تطبیق قطعی شاخص", "matching", ""),
        "matching.probable": (str(matching_counts.get("probable", 0)), "تطبیق احتمالی شاخص", "matching", ""),
        "matching.ambiguous": (str(matching_counts.get("ambiguous", 0)), "تطبیق مبهم شاخص", "matching", ""),
        "matching.rejected": (str(matching_counts.get("rejected", 0)), "تطبیق رد‌شده", "matching", ""),
        "duplicates.groups": (str(duplicate_counts.get("duplicate_groups", 0)), "گروه تکرار", "duplicate_control", ""),
        "duplicates.blocked": (str(duplicate_counts.get("blocked", 0)), "تصمیم تجمیع مسدود", "duplicate_control", ""),
        "provenance.coverage": (
            f"{provenance.coverage() * 100:.1f}٪" if provenance is not None else "—",
            "پوشش ردیابی منبع",
            "provenance",
            "",
        ),
        "provenance.records": (
            str(provenance_counts.get("important", 0)),
            "رکوردهای مهم ردیابی",
            "provenance",
            "",
        ),
        "gate.status": (_status(getattr(gate, "status", None)), "وضعیت دروازه کیفیت", "quality_gate", ""),
    }

    sections: list[ReportSection] = []
    identity_rows = [
        ReportRow("فایل‌های منبع", "، ".join(_name(item) for item in files) or "—"),
        ReportRow("فایل‌های تحلیل‌شده", _text(analyses_count)),
        ReportRow("فایل‌های ناموفق", _text(len(failed))),
        ReportRow("دوره‌های شناسایی‌شده", "، ".join(periods) or "در دسترس نبود"),
        ReportRow("درخواست کاربر", request or getattr(result, "request", "") or "—"),
        ReportRow("تاریخ گزارش", stamp_text),
    ]
    sections.append(ReportSection("شناسنامه گزارش", identity_rows))

    tables: list[ReportTable] = []

    file_rows = []
    for item in files:
        analysis = getattr(item, "analysis", None)
        blocks = _blocks(analysis)
        file_periods = list(getattr(getattr(analysis, "integration", None), "periods", []) or [])
        if not file_periods:
            file_periods = sorted(
                {
                    str(period.get("label") or period.get("token") or "")
                    for block in blocks
                    for period in getattr(block, "periods", []) or []
                    if isinstance(period, dict)
                }
                - {""}
            )
        file_rows.append(
            [
                _name(item),
                _text(getattr(item, "status", "")),
                str(len(_sheet_names(analysis))) if analysis is not None else "0",
                str(len(blocks)),
                "، ".join(file_periods) or "—",
                _text(getattr(item, "error", "") or "—"),
            ]
        )
    tables.append(
        ReportTable(
            "فایل‌های منبع",
            ["فایل", "وضعیت", "برگه‌های بازرسی‌شده", "بلوک‌های استخراج‌شده", "دوره‌ها", "خطا"],
            file_rows,
            caption="جدول وضعیت پردازش فایل‌های منبع",
            source_note="حضور یک فایل در این فهرست به‌معنای موفقیت استخراج آن نیست؛ ستون وضعیت معیار است.",
        )
    )

    comparisons = list(getattr(indicators, "comparisons", []) or [])
    compared = [item for item in comparisons if item.status == STATUS_COMPARED]
    if comparisons:
        rows = []
        for item in comparisons[:60]:
            rows.append(
                [
                    _text(item.indicator),
                    "جمع ثبت‌شده" if getattr(item, "role", "data") == "total" else "داده",
                    _number(item.previous),
                    _number(item.current),
                    _number(item.difference),
                    _percent(item.percent),
                    _text(getattr(item, "status_label", "")),
                    f"{item.sheet}!{item.previous_address} ← {item.current_address}",
                ]
            )
        tables.append(
            ReportTable(
                "مقایسه شاخص‌ها بین دوره‌ها",
                ["شاخص", "نقش", "دوره قبل", "دوره جاری", "اختلاف", "درصد تغییر", "وضعیت", "منبع (سلول)"],
                rows,
                caption="جدول مقایسه شاخص‌های ردیفی در دو دوره گزارش‌دهی",
                source_note=(
                    "هر مقدار یک سلول مشخص است؛ شاخص‌ها با هم جمع نشده‌اند. ردیف جمع/کل فقط "
                    "به‌عنوان مقدار ثبت‌شده با دوره دیگر مقایسه می‌شود و به اجزا اضافه نمی‌شود؛ "
                    "ردیف‌های درصد و توضیح وارد مقایسه نشده‌اند."
                ),
                statistical_note=(
                    "درصد تغییر فقط وقتی محاسبه می‌شود که مقدار دوره قبل موجود و غیرصفر باشد."
                ),
            )
        )

    if matching is not None:
        rows = []
        for decision in list(getattr(matching, "decisions", []) or [])[:60]:
            members = "؛ ".join(
                _text(getattr(member, "display_label", "")) for member in getattr(decision, "members", []) or []
            )
            rows.append(
                [
                    _text(getattr(decision, "status", "")),
                    _text(getattr(decision, "kind_label", "")),
                    _text(getattr(decision, "normalized_label", "")),
                    members or "—",
                    "بله" if getattr(decision, "comparable", False) else "خیر",
                    _text(getattr(decision, "reason", "")),
                ]
            )
        if rows and len(analyses) > 1:
            tables.append(
                ReportTable(
                    "تطبیق شاخص‌ها بین منابع",
                    ["وضعیت", "نوع", "برچسب نرمال‌شده", "اعضا", "قابل مقایسه", "دلیل"],
                    rows,
                    caption="جدول تصمیم تطبیق معنایی شاخص‌ها بین فایل‌ها",
                    source_note="این جدول فقط تصمیم تطبیق را ثبت می‌کند و هیچ مقداری را جمع نمی‌کند.",
                )
            )

    if duplicate is not None:
        decisions = list(getattr(duplicate, "decisions", []) or [])
        blocked_or_review = [
            item for item in decisions if _status(getattr(item, "status", "")).upper() != "مجاز"
        ]
        rows = [
            [
                _text(getattr(item, "kind", "")),
                _text(getattr(item, "status_label", "")),
                "، ".join(sorted(getattr(item, "headers", set()) or [])),
                "، ".join(sorted(getattr(item, "sources", set()) or [])),
                _text(getattr(item, "reason", "")),
            ]
            for item in blocked_or_review[:40]
        ]
        if rows:
            tables.append(
                ReportTable(
                    "کنترل تکرار و تجمیع",
                    ["نوع", "وضعیت", "ستون‌ها", "منابع", "دلیل"],
                    rows,
                    caption="جدول تصمیم‌های تجمیع بین منابع",
                    source_note="جمع بین فایل‌ها فقط با تصمیم «مجاز» گزارش می‌شود؛ هیچ مقدار حذف یا اصلاح نشد.",
                )
            )
        exclusions = list(getattr(duplicate, "exclusions", []) or [])
        if exclusions:
            rows = [
                [
                    _text(getattr(item, "kind", "")),
                    _text(getattr(item, "source", "")),
                    _text(getattr(item, "sheet", "")),
                    _text(getattr(item, "label", "")),
                    _text(getattr(item, "reason", "")),
                ]
                for item in exclusions[:40]
            ]
            tables.append(
                ReportTable(
                    "موارد خارج‌شده از تجمیع",
                    ["نوع", "منبع", "برگه", "برچسب", "دلیل"],
                    rows,
                    caption="جدول موارد خارج‌شده از تجمیع",
                    source_note="این موارد از جمع کنار گذاشته شده‌اند و در گزارش به‌عنوان داده حذف نشده‌اند.",
                )
            )

    reconciliation_rows: list[list[str]] = []
    quality_rows: list[list[str]] = []
    for item in analyses:
        analysis = item.analysis
        statistical = getattr(analysis, "statistical_reconciliation", None)
        for record in getattr(statistical, "reconciliations", []) or []:
            status = _status(getattr(record, "status", ""))
            if status.upper() == "PASS":
                continue
            reconciliation_rows.append(
                [
                    _name(item),
                    _text(getattr(record, "label", "") or getattr(record, "measure", "")),
                    _number(getattr(record, "expected", None)),
                    _number(getattr(record, "calculated", None)),
                    _number(getattr(record, "difference", None)),
                    status,
                ]
            )
        quality = getattr(analysis, "quality", None)
        for finding in getattr(quality, "findings", []) or []:
            status = _status(getattr(finding, "status", ""))
            if status.upper() == "PASS":
                continue
            quality_rows.append(
                [
                    _name(item),
                    _text(getattr(finding, "code", "") or getattr(finding, "title", "")),
                    status,
                    str(getattr(finding, "count", 0)),
                    _text(getattr(finding, "message", "")),
                ]
            )
    if reconciliation_rows:
        tables.append(
            ReportTable(
                "تطبیق آماری و صحت محاسبات",
                ["فایل", "شاخص", "مقدار انتظار", "مقدار محاسبه", "اختلاف", "وضعیت"],
                reconciliation_rows[:40],
                caption="جدول تطبیق‌های آماری غیرقبول",
                source_note="مقدار انتظار از ردیف جمع ثبت‌شده و مقدار محاسبه از جمع اجزای مستقل می‌آید.",
            )
        )
    if quality_rows:
        tables.append(
            ReportTable(
                "کیفیت داده",
                ["فایل", "بررسی", "وضعیت", "تعداد", "توضیح"],
                quality_rows[:40],
                caption="جدول یافته‌های کیفیت داده غیرقبول",
                source_note="هیچ مقدار مشکوکی اصلاح یا حذف نشده است؛ فقط ثبت شده است.",
            )
        )

    if accidents is not None and getattr(accidents, "tables", 0):
        metric_rows = [
            [item.label, item.kind_label, f"{item.sheet}!{item.address}"]
            for item in list(getattr(accidents, "metrics", []) or [])[:40]
        ]
        tables.append(
            ReportTable(
                "تفکیک آمار تصادف و تلفات",
                ["برچسب منبع", "نوع شاخص", "منبع (سلول)"],
                metric_rows or [["—", "—", "—"]],
                caption="جدول تفکیک شاخص‌های تصادف و تلفات",
                source_note=(
                    "«تعداد تصادف فوتی/جرحی» با «تعداد افراد فوت‌شده/مجروح» یکسان نیست و "
                    "این دو با هم جمع نمی‌شوند."
                ),
            )
        )
        findings = [
            [
                _text(item.status),
                _text(item.code),
                _text(item.message),
            ]
            for item in list(getattr(accidents, "findings", []) or [])
        ]
        if findings:
            tables.append(
                ReportTable(
                    "یافته‌های معنایی تصادف",
                    ["وضعیت", "کد بررسی", "توضیح"],
                    findings[:40],
                    caption="جدول یافته‌های تفکیک تصادف و تلفات",
                )
            )

    if gate is not None:
        stage_rows = [
            [stage.stage.label_fa, _status(stage.status), _text(stage.summary)]
            for stage in getattr(gate, "stages", []) or []
        ]
        tables.append(
            ReportTable(
                "دروازه کیفیت",
                ["مرحله", "وضعیت", "خلاصه"],
                stage_rows,
                caption="جدول مراحل دروازه کیفیت",
                source_note=(
                    "مراحل قالب خروجی و سلامت سند نهایی پس از تولید فایل ارزیابی می‌شوند؛ "
                    "نتیجه نهایی در گزارش حسابرسی ماشین‌خوان ثبت است."
                ),
            )
        )
    if provenance is not None:
        counts = provenance_counts
        tables.append(
            ReportTable(
                "ردیابی منبع",
                ["موضوع", "مقدار"],
                [
                    ["رکوردهای مهم", str(counts.get("important", 0))],
                    ["رکوردهای دارای منبع", str(counts.get("with_sources", 0))],
                    ["رکوردهای کامل", str(counts.get("complete", 0))],
                    ["رکوردهای ناقص", str(counts.get("incomplete", 0))],
                    ["پوشش ردیابی", f"{provenance.coverage() * 100:.1f}٪"],
                ],
                caption="جدول پوشش ردیابی منبع",
                source_note="هر رقم محاسبه‌شده باید به سلول منبع یا محدوده مشخص متصل باشد.",
            )
        )

    warnings: list[str] = list(getattr(result, "warnings", []) or [])
    for source in (matching, duplicate, indicators, accidents):
        if source is matching and len(analyses) < 2:
            continue
        for warning in list(getattr(source, "warnings", []) or []) if source is not None else []:
            if warning not in warnings:
                warnings.append(str(warning))
    if failed:
        warnings.append(
            "گزارش جزئی است: "
            + "، ".join(_name(item) for item in failed)
            + " تحلیل نشد و داده آن در هیچ جمع یا مقایسه‌ای وارد نشده است."
        )
    if not comparisons:
        warnings.append(
            "برای هیچ شاخص ردیفی دو دوره قابل مقایسه پیدا نشد؛ مقایسه دوره‌ها در این گزارش وجود ندارد."
        )
    if gate is not None:
        prior_stages = [
            stage
            for stage in getattr(gate, "stages", []) or []
            if getattr(getattr(stage, "stage", ""), "name", "")
            not in {"OUTPUT_FORMAT", "DOCUMENT_INTEGRITY"}
        ]
        if any(
            str(getattr(getattr(stage, "status", None), "value", "PASS")).upper() != "PASS"
            for stage in prior_stages
        ):
            warnings.append(
                "دروازه کیفیت پیش از تولید سند وضعیت قبول ندارد؛ این سند موفقیت کامل گزارش نمی‌شود."
            )

    executive_summary = [
        f"{len(files)} فایل منبع شناسایی شد و {analyses_count} فایل تحلیل شد.",
        f"{len(all_blocks)} بلوک داده استخراج شد و {len(periods)} دوره شناسایی شد.",
        f"{len(compared)} شاخص در دو دوره با مقدار هر دو دوره مقایسه شد.",
        f"{matching_counts.get('ambiguous', 0)} تطبیق مبهم و {duplicate_counts.get('blocked', 0)} تصمیم تجمیع مسدود ثبت شد.",
    ]
    if not periods:
        executive_summary.append("دوره گزارش‌دهی به‌طور کامل شناسایی نشد؛ نتیجه مقایسه دوره‌ای معتبر نیست.")

    conclusions: list[str] = []
    decreases = [item for item in compared if item.difference is not None and item.difference < 0]
    increases = [item for item in compared if item.difference is not None and item.difference > 0]
    if decreases:
        item = min(decreases, key=lambda entry: entry.percent if entry.percent is not None else Decimal("0"))
        conclusions.append(
            f"بیشترین کاهش مربوط به «{item.indicator}» با {_percent(item.percent)} است."
        )
    if increases:
        item = max(increases, key=lambda entry: entry.percent if entry.percent is not None else Decimal("0"))
        conclusions.append(
            f"بیشترین افزایش مربوط به «{item.indicator}» با {_percent(item.percent)} است."
        )
    if accidents is not None and getattr(accidents, "missing", None):
        conclusions.append(
            "شاخص‌های تصادف که در منابع نبودند به‌صراحت به‌عنوان محدودیت اعلام شده‌اند و حدس زده نشدند."
        )

    recommendations = [
        "پیش از استفاده اداری، هر رقم کلیدی را با سلول منبع درج‌شده در جدول مقایسه مطابقت دهید.",
        "تفکیک «تعداد تصادف» از «تعداد افراد فوتی و مجروح» در همه گزارش‌های بعدی حفظ شود.",
        "برای شاخص‌هایی که تطبیق مبهم دارند، تعریف واحدی از شاخص در دو منبع تأیید شود.",
        "جمع بین فایل‌ها فقط پس از تصمیم «مجاز» کنترل تکرار گزارش شود.",
    ]

    source_notes = [
        "هر جدول، نشانی برگه و سلول مقدار خود را دارد؛ فایل مرجع Word فقط قالب است و داده آن وارد گزارش نشد.",
        "فایل‌های منبع در این اجرا تغییر نکردند.",
    ]
    statistical_notes = [
        "همه محاسبات با Decimal در Python انجام شده است و درصد تغییر فقط با مبنای غیرصفر محاسبه می‌شود.",
        "مقدار ناموجود «در دسترس نبود» است و هرگز صفر فرض نشده است.",
        "شاخص‌های ردیفی با هم جمع نشده‌اند؛ فقط هر شاخص در دو دوره مقایسه شده است.",
    ]

    content_values = {}
    for key, (value, label, source, address) in values.items():
        content_values[key] = _report_value(key, label, value, source, address)

    has_accident = accidents is not None and bool(getattr(accidents, "tables", 0))
    return ReportContent(
        title=report_title,
        subtitle="گزارش چندمنبعی تصادفات و تخلفات" if has_accident else "گزارش چندمنبعی",
        sections=sections,
        values=content_values,
        warnings=warnings,
        source_kind="python_analysis",
        validated=True,
        generated_at=stamp_text,
        tables=tables,
        executive_summary=executive_summary,
        conclusions=conclusions,
        recommendations=recommendations,
        source_notes=source_notes,
        statistical_notes=statistical_notes,
        quality_gate=gate.to_dict() if gate is not None and hasattr(gate, "to_dict") else None,
        provenance_summary=(
            provenance.to_dict() if provenance is not None and hasattr(provenance, "to_dict") else None
        ),
    )


def _report_value(key: str, label: str, value: Any, source: str, address: str):
    from local_office_ai.reporting.specification import ReportValue

    text = _text(value)
    return ReportValue(
        key=key,
        label=label,
        value=text,
        available=text not in {"", "—"},
        source=source,
        source_address=address,
        validated=True,
    )
