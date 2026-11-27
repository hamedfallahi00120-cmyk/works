"""Build a compact, verified fact package for the model.

Only values computed by Python are sent.  The package is trimmed to a hard
character budget so a large workbook can never overflow the model context.
"""

from __future__ import annotations

import json
from datetime import date, datetime
from decimal import Decimal
from typing import Any

from local_office_ai.core.models import AnalysisResult
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import canonical_number

KEY_CALCULATIONS = (
    "loan_count",
    "principal_sum",
    "installments_sum",
    "obligations_sum",
    "contract_repayment_sum",
    "obligation_gap",
    "installment_to_principal_ratio",
)


def _cell_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return value


def _num(value: Any) -> Any:
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (int, float)):
        return value
    return None


def _generic_facts(analysis: AnalysisResult, generic: Any = None) -> dict[str, Any] | None:
    generic = generic if generic is not None else analysis.generic
    if generic is None:
        return None
    quality = generic.quality
    columns = [
        {
            "header": column.header,
            "type": column.kind.value,
            "role": column.role.value,
            "count": column.count,
            "missing": column.missing,
            "distinct": column.distinct,
            "identifier": column.is_identifier,
            "samples": [value for value, _ in column.top_values[:3]],
        }
        for column in generic.columns
    ]
    descriptive = []
    for column in generic.columns:
        if column.is_identifier:
            continue
        if column.kind.value == "date":
            descriptive.append(
                {
                    "header": column.header,
                    "type": "date",
                    "count": column.count,
                    "missing": column.missing,
                    "date_min": column.date_min,
                    "date_max": column.date_max,
                    "date_calendar": column.date_calendar,
                    "range_days": column.range_days,
                }
            )
            continue
        if not column.is_numeric:
            continue
        descriptive.append(
            {
                "header": column.header,
                "type": column.kind.value,
                "count": column.count,
                "missing": column.missing,
                "sum": _num(column.total),
                "mean": _num(column.mean),
                "median": _num(column.median),
                "mode": column.mode,
                "min": _num(column.minimum),
                "max": _num(column.maximum),
                "range": _num(column.value_range),
                "variance_sample": _num(column.variance_sample),
                "variance_population": _num(column.variance_population),
                "std_sample": _num(column.std_sample),
                "std_population": _num(column.std_population),
                "q1": _num(column.q1),
                "q3": _num(column.q3),
                "iqr": _num(column.iqr),
                "p10": _num(column.p10),
                "p90": _num(column.p90),
                "cv_percent": _num(column.cv_percent),
            }
        )
    frequencies = [
        {
            "header": table.header,
            "total": table.total,
            "distinct": table.distinct,
            "top": [[row.value, row.count, _num(row.percent)] for row in table.rows[:8]],
            "bins": [
                {"from": _num(bin.lower), "to": _num(bin.upper), "count": bin.count, "percent": _num(bin.percent)}
                for bin in table.bins[:12]
            ],
        }
        for table in generic.frequencies[:6]
    ]
    groups = [
        {
            "group": item.group_header,
            "measure": item.measure_header,
            "values": [
                {
                    "label": group.label,
                    "count": group.count,
                    "sum": _num(group.total),
                    "mean": _num(group.mean),
                    "median": _num(group.median),
                    "min": _num(group.minimum),
                    "max": _num(group.maximum),
                    "std_sample": _num(group.std_sample),
                }
                for group in item.groups[:12]
            ],
        }
        for item in generic.groups[:4]
    ]
    crosstabs = [
        {
            "rows": item.row_header,
            "columns": item.column_header,
            "matrix": {row.label: dict(row.counts) for row in item.rows[:12]},
            "total": item.total,
        }
        for item in generic.crosstabs[:2]
    ]
    trends = [
        {
            "date": item.date_header,
            "measure": item.measure_header,
            "granularity": item.granularity,
            "calendar": item.calendar,
            "direction": item.direction,
            "change_absolute": _num(item.change_absolute),
            "change_percent": _num(item.change_percent),
            "points": [
                {"period": point.period_label, "count": point.count, "mean": _num(point.mean)}
                for point in item.points[:24]
            ],
        }
        for item in generic.trends[:3]
    ]
    correlations = [
        {
            "a": item.header_a,
            "b": item.header_b,
            "n": item.count,
            "pearson": _num(item.pearson),
            "spearman": _num(item.spearman),
            "strength": item.strength,
            "direction": item.direction,
        }
        for item in generic.correlations[:8]
    ]
    outliers = [
        {
            "column": item.header,
            "address": item.address,
            "value": _num(item.value),
            "method": item.method,
        }
        for item in generic.outliers[:12]
    ]
    return {
        "table": {
            "sheet": generic.table_sheet,
            "address": generic.table_address,
            "kind": generic.table_kind,
            "rows": generic.row_count,
            "columns": generic.column_count,
        },
        "quality": {
            "missing_cells": quality.missing_cells if quality else 0,
            "missing_ratio": round(quality.missing_ratio, 4) if quality else 0,
            "measured_missing_ratio": round(quality.measured_missing_ratio, 4) if quality else 0,
            "empty_columns": quality.empty_columns if quality else [],
            "duplicate_rows": quality.duplicate_rows if quality else 0,
            "empty_rows": quality.empty_rows if quality else 0,
            "constant_columns": quality.constant_columns if quality else [],
            "mixed_type_columns": quality.mixed_type_columns if quality else [],
            "identifier_columns": quality.identifier_columns if quality else [],
            "invalid_cells": quality.invalid_cells if quality else 0,
        },
        "columns": columns,
        "descriptive": descriptive,
        "frequencies": frequencies,
        "groups": groups,
        "crosstabs": crosstabs,
        "trends": trends,
        "correlations": correlations,
        "outliers": {"count": len(generic.outliers), "items": outliers},
        "methodology": generic.methodology,
        "goal": generic.goal,
    }


def _intelligence_facts(analysis: AnalysisResult) -> dict[str, Any] | None:
    intelligence = analysis.intelligence
    if intelligence is None:
        return None
    mismatches = [
        item for item in intelligence.all_audits if item.status.value == "mismatch"
    ]
    unverifiable = [
        item for item in intelligence.all_audits if item.status.value == "not_verifiable"
    ]
    verified = [item for item in intelligence.all_audits if item.status.value == "verified"]
    sheets = []
    for sheet in intelligence.sheets:
        coverage = sheet.coverage
        sheets.append(
            {
                "sheet": sheet.sheet,
                "status": sheet.status,
                "month": sheet.month,
                "sheet_years": sheet.sheet_years,
                "title_years": sheet.title_years,
                "subheader_years": sheet.subheader_years,
                "rows": sheet.row_count,
                "columns": sheet.column_count,
                "derived": sheet.derived,
                "coverage": None
                if coverage is None
                else {
                    "expected": coverage.expected_observations,
                    "observed": coverage.observed_observations,
                    "missing": coverage.missing_observations,
                    "missing_labels": coverage.missing_labels[:15],
                    "explicit_zeros": coverage.explicit_zero_cells,
                    "blanks": coverage.blank_cells,
                    "invalid": coverage.invalid_cells,
                    "errors": coverage.error_cells,
                    "incomplete": coverage.incomplete_period,
                },
                "audit_mismatches": sum(
                    1 for item in sheet.audits if item.status.value == "mismatch"
                ),
                "audit_unverifiable": sum(
                    1 for item in sheet.audits if item.status.value == "not_verifiable"
                ),
                "metadata_conflicts": len(sheet.metadata_conflicts),
            }
        )
    return {
        "sheet_count": intelligence.sheet_count,
        "status_counts": dict(intelligence.status_counts),
        "sheets": sheets,
        "similar_groups": [
            {
                "sheets": group["sheets"],
                "column_count": group["column_count"],
                "period_pairs": group["period_pairs"],
            }
            for group in intelligence.similar_groups[:8]
        ],
        "metadata_conflicts": [
            {"sheet": item.sheet, "kind": item.kind, "message": item.message}
            for item in intelligence.metadata_conflicts[:20]
        ],
        "audit_counts": {
            "verified": len(verified),
            "mismatch": len(mismatches),
            "not_verifiable": len(unverifiable),
            "cross_sheet": len(intelligence.cross_sheet_audits),
        },
        "audit_mismatches": [
            {
                "sheet": item.sheet,
                "address": item.address,
                "kind": item.kind,
                "metric": item.metric,
                "declared": _num(item.declared),
                "computed": _num(item.computed),
                "difference": _num(item.difference),
                "cause": item.cause,
            }
            for item in mismatches[:20]
        ],
        "audit_unverifiable": [
            {
                "sheet": item.sheet,
                "address": item.address,
                "metric": item.metric,
                "declared": _num(item.declared),
                "cause": item.cause,
            }
            for item in unverifiable[:12]
        ],
        "period_comparisons": [
            {
                "sheet": item.sheet,
                "metric": item.metric,
                "period_a": item.period_a,
                "period_b": item.period_b,
                "total_a": _num(item.total_a),
                "total_b": _num(item.total_b),
                "difference": _num(item.difference),
                "percent": _num(item.percent),
            }
            for item in intelligence.period_comparisons[:30]
        ],
        "aggregations": [
            {
                "metric": item.metric,
                "period": item.period,
                "total": _num(item.total),
                "months": item.months,
                "note": item.note,
            }
            for item in intelligence.aggregations[:20]
        ],
        "warnings": intelligence.warnings[:12],
    }


def build_facts(
    analysis: AnalysisResult,
    settings: Settings,
    *,
    max_rows: int | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Return (facts, meta) where meta records what had to be trimmed."""
    max_rows = max_rows or settings.max_ai_rows
    calculations = {}
    for name in KEY_CALCULATIONS:
        calculation = analysis.calculation(name)
        if calculation is not None:
            calculations[name] = {
                "label": calculation.label,
                "value": _num(calculation.value),
                "unit": calculation.unit,
                "verified": calculation.verified,
            }

    audits = [
        {
            "label": audit.label,
            "declared": _num(audit.declared),
            "computed": _num(audit.computed),
            "difference": _num(audit.difference),
            "status": audit.status.value,
        }
        for audit in analysis.audits
    ]
    reconciliation = [
        {
            "institution": item.institution,
            "loan_total": _num(item.loan_total),
            "obligation_total": _num(item.obligation_total),
            "difference": _num(item.difference),
            "status": item.status.value,
            "similarity": round(item.similarity, 2),
            "note": item.note,
        }
        for item in analysis.reconciliation
    ]
    anomalies = [
        {"severity": item.severity.value, "category": item.category, "message": item.message}
        for item in analysis.anomalies
        if item.severity.value != "info"
    ][:25]

    rows: list[dict[str, Any]] = []
    primary = analysis.primary_table
    if primary is not None:
        for row in analysis.normalized_rows[:max_rows]:
            rows.append(
                {
                    "row": row["row"],
                    "values": {
                        cell["header"] or cell["column"]: _cell_value(cell["value"])
                        for cell in row["cells"]
                    },
                }
            )

    facts: dict[str, Any] = {
        "file": {
            "name": analysis.meta.path.rsplit("/", 1)[-1].rsplit("\\", 1)[-1],
            "sheet_count": len(analysis.meta.sheets),
            "structure_confidence": round(analysis.structure_confidence, 2),
            "formula_cells": analysis.meta.total_formulas,
            "error_cells": analysis.meta.total_error_cells,
        },
        "verified_calculations_from_python": calculations,
        "dataset_profile": _generic_facts(analysis),
        "workbook_intelligence": _intelligence_facts(analysis),
        "period_analysis": _period_facts(analysis),
        "request_answers": [
            {"title": answer.title, "value": answer.value, "text": answer.text}
            for answer in (getattr(analysis, "request_answers", None) or [])[:14]
        ],
        "request_text": getattr(getattr(analysis, "request_intent", None), "text", ""),
        "audits_of_declared_totals": audits,
        "reconciliation": reconciliation,
        "statistical_reconciliation": _reconciliation_facts(analysis),
        "reconciliation_counts": {
            "matched": sum(1 for item in analysis.reconciliation if item.status.value == "match"),
            "probable": sum(1 for item in analysis.reconciliation if item.status.value == "probable"),
            "outside_loan_table": sum(
                1 for item in analysis.reconciliation if item.status.value == "outside_loan_table"
            ),
            "no_obligation": sum(
                1 for item in analysis.reconciliation if item.status.value == "no_obligation"
            ),
        },
        "loan_institutions": sorted({loan.institution_raw for loan in analysis.loans if loan.institution_raw}),
        "important_anomalies": anomalies,
        "primary_table": None
        if primary is None
        else {
            "sheet": primary.sheet,
            "address": primary.address,
            "kind": primary.kind.value,
            "headers": [column.header for column in primary.columns],
            "row_count": primary.row_count,
        },
        "sample_rows": rows,
        "data_gaps": _data_gaps(analysis),
        "rules": [
            "همه اعداد و محاسبات از Python آمده‌اند؛ محاسبه مجدد انجام نده.",
            "فقط اعدادی را نقل کن که در همین بسته آمده‌اند.",
            "همبستگی را رابطه علت‌ومعلولی معرفی نکن؛ پرت‌ها حذف نشده‌اند.",
            "مغایرت را فقط اگر در audits یا reconciliation آمده باشد اعلام کن.",
            "اگر داده‌ای نیست، آن را در missing_data بنویس.",
        ],
    }
    return _trim(facts, settings.max_ai_context_chars)


def _reconciliation_facts(analysis: AnalysisResult) -> dict[str, Any] | None:
    report = getattr(analysis, "statistical_reconciliation", None)
    if report is None:
        return None
    return {
        "status": report.status.value,
        "counts": report.counts(),
        "reconciliations": [
            {
                "kind": item.kind,
                "label": item.label,
                "measure": item.measure,
                "status": item.status.value,
                "expected": _num(item.expected),
                "calculated": _num(item.calculated),
                "difference": _num(item.difference),
                "explanation": item.explanation,
                "source": item.expected_source or (item.sources[0] if item.sources else ""),
            }
            for item in report.reconciliations[:12]
        ],
        "comparisons": [
            {
                "metric": item.metric,
                "previous_period": item.previous_label,
                "current_period": item.current_label,
                "status": item.status.value,
                "previous": _num(item.change.previous),
                "current": _num(item.change.current),
                "absolute": _num(item.change.absolute),
                "percent": _num(item.change.percent),
                "direction": item.change.direction,
                "months_used": item.months_used,
            }
            for item in report.comparisons[:12]
        ],
        "averages": [
            {
                "label": item.label,
                "status": item.status.value,
                "value": _num(item.value),
                "unweighted": _num(item.unweighted),
                "weighted": _num(item.weighted),
                "weighted_used": item.weighted_used,
                "explanation": item.explanation,
            }
            for item in report.averages[:8]
        ],
        "note": (
            "همه مقادیر تطبیق در Python با Decimal محاسبه شده‌اند؛ مدل فقط می‌تواند "
            "همین اعداد و وضعیت‌ها را توضیح دهد و مجاز به محاسبه مجدد نیست."
        ),
    }


def _period_facts(analysis: AnalysisResult) -> dict[str, Any] | None:
    integration = getattr(analysis, "integration", None)
    if integration is None or not getattr(integration, "is_usable", False):
        return None
    metrics = []
    for index, series in enumerate(integration.series[:8]):
        entry: dict[str, Any] = {"metric": series.metric, "totals": {}, "best": {}, "worst": {}}
        for period in integration.periods:
            entry["totals"][period] = _num(series.total(period))
            best = series.best(period)
            worst = series.worst(period)
            if best is not None:
                entry["best"][period] = {"month": best.month, "total": _num(best.total)}
            if worst is not None:
                entry["worst"][period] = {"month": worst.month, "total": _num(worst.total)}
        if index < 3:
            entry["month_totals"] = {
                period: {
                    cell.month: _num(cell.total)
                    for cell in series.cells
                    if cell.period == period and cell.total is not None
                }
                for period in integration.periods
            }
        metrics.append(entry)
    from local_office_ai.excel.data_qa import comparison_pair, ranking_period

    intent = getattr(analysis, "request_intent", None)
    pair = comparison_pair(intent, integration) if intent is not None else None
    periods = list(pair) if pair else integration.comparison_periods()
    comparison = []
    if len(periods) >= 2:
        for series in integration.series[:8]:
            change = series.compare(periods[0], periods[1])
            comparison.append(
                {
                    "metric": series.metric,
                    "period_a": change.period_a,
                    "period_b": change.period_b,
                    "total_a": _num(change.total_a),
                    "total_b": _num(change.total_b),
                    "difference": _num(change.difference),
                    "percent": _num(change.percent),
                    "months_used": change.months_used,
                    "direction": change.direction,
                }
            )
    main_period = ranking_period(intent, integration) if intent is not None else integration.main_period()
    return {
        "months": list(integration.months),
        "periods": list(integration.periods),
        "main_period": main_period,
        "comparison_periods": periods,
        "suspicious_periods": list(integration.suspicious_periods),
        "sources": list(integration.sources),
        "comparison": comparison,
        "metrics": metrics,
        "note": (
            "جمع هر ماه از سلول‌های همان برگه بازمحاسبه شده است؛ ماه ناقص در جمع و رتبه‌بندی نیست. "
            "مقایسه فقط روی ماه‌های مشترک و هم‌پوشش انجام شده است."
        ),
    }


def _data_gaps(analysis: AnalysisResult) -> list[str]:
    gaps: list[str] = []
    generic = analysis.generic
    if analysis.primary_table is None:
        gaps.append("جدول داده اصلی شناسایی نشد.")
    if generic is not None:
        if not any(column.is_numeric for column in generic.columns):
            gaps.append("ستون عددی معتبری برای تحلیل آماری وجود ندارد.")
        if generic.quality:
            ratio = (
                generic.quality.measured_missing_ratio
                if generic.quality.empty_columns
                else generic.quality.missing_ratio
            )
            if ratio >= 0.2:
                gaps.append(f"{ratio:.0%} سلول‌های اندازه‌گیری‌شده خالی است.")
            if generic.quality.empty_columns:
                gaps.append(
                    "ستون‌های بدون هیچ مقدار (دوره اندازه‌گیری‌نشده): "
                    + "، ".join(generic.quality.empty_columns[:6])
                )
        if generic.outliers:
            gaps.append(f"{len(generic.outliers)} مقدار پرت علامت‌گذاری شده و حذف نشده است.")
    if analysis.loans:
        for loan in analysis.loans:
            if loan.principal is None or loan.installment is None or loan.installment_count is None:
                gaps.append(f"برای «{loan.institution_raw}» برخی فیلدهای کلیدی خالی است.")
    if analysis.meta.total_error_cells:
        gaps.append(f"{analysis.meta.total_error_cells} سلول خطای فرمول در فایل وجود دارد.")
    return gaps[:12]


def _trim(facts: dict[str, Any], budget: int) -> tuple[dict[str, Any], dict[str, Any]]:
    """Shrink the payload until it fits the character budget.

    Trimming is progressive and recorded, so the report can say exactly how much
    of the sample data the model did not see.
    """
    meta: dict[str, Any] = {
        "chars": 0,
        "trimmed_rows": 0,
        "trimmed_anomalies": 0,
        "trimmed_reconciliation": 0,
        "trimmed_audits": 0,
        "trimmed_generic": 0,
        "degraded": False,
    }
    for key in ("sample_rows", "important_anomalies", "reconciliation", "audits_of_declared_totals", "rules", "data_gaps"):
        facts.setdefault(key, [])
    facts.setdefault("file", {})
    facts.setdefault("verified_calculations_from_python", {})
    payload = json.dumps(facts, ensure_ascii=False)

    while len(payload) > budget and facts["sample_rows"]:
        facts["sample_rows"] = facts["sample_rows"][: max(0, len(facts["sample_rows"]) - 5)]
        meta["trimmed_rows"] += 5
        payload = json.dumps(facts, ensure_ascii=False)
    while len(payload) > budget and facts["important_anomalies"]:
        facts["important_anomalies"] = facts["important_anomalies"][:-1]
        meta["trimmed_anomalies"] += 1
        payload = json.dumps(facts, ensure_ascii=False)
    while len(payload) > budget and len(facts["reconciliation"]) > 5:
        facts["reconciliation"] = facts["reconciliation"][:-1]
        meta["trimmed_reconciliation"] += 1
        payload = json.dumps(facts, ensure_ascii=False)
    while len(payload) > budget and len(facts["audits_of_declared_totals"]) > 5:
        facts["audits_of_declared_totals"] = facts["audits_of_declared_totals"][:-1]
        meta["trimmed_audits"] += 1
        payload = json.dumps(facts, ensure_ascii=False)
    dataset = facts.get("dataset_profile")
    if dataset:
        for key in ("frequencies", "groups", "crosstabs", "trends", "correlations", "descriptive", "columns"):
            while len(payload) > budget and dataset.get(key):
                dataset[key] = dataset[key][:-1]
                meta["trimmed_generic"] += 1
                payload = json.dumps(facts, ensure_ascii=False)
        outliers = dataset.get("outliers")
        if outliers and outliers.get("items"):
            while len(payload) > budget and outliers["items"]:
                outliers["items"] = outliers["items"][:-1]
                meta["trimmed_generic"] += 1
                payload = json.dumps(facts, ensure_ascii=False)
    profiles = facts.get("dataset_profiles")
    if profiles:
        for key in ("frequencies", "groups", "crosstabs", "trends", "correlations", "descriptive", "columns"):
            index = 0
            while len(payload) > budget and index < len(profiles):
                profile = profiles[index]
                if profile and profile.get(key):
                    profile[key] = profile[key][:-1]
                    meta["trimmed_generic"] += 1
                    payload = json.dumps(facts, ensure_ascii=False)
                    continue
                index += 1
    period = facts.get("period_analysis")
    if period:
        while len(payload) > budget and len(period.get("metrics") or []) > 3:
            period["metrics"] = period["metrics"][:-1]
            meta["trimmed_generic"] += 1
            payload = json.dumps(facts, ensure_ascii=False)
        while len(payload) > budget and len(period.get("comparison") or []) > 3:
            period["comparison"] = period["comparison"][:-1]
            meta["trimmed_generic"] += 1
            payload = json.dumps(facts, ensure_ascii=False)
        if len(payload) > budget:
            for entry in period.get("metrics") or []:
                entry.pop("month_totals", None)
            payload = json.dumps(facts, ensure_ascii=False)
    reconciliation = facts.get("statistical_reconciliation")
    if reconciliation:
        for key in ("reconciliations", "comparisons", "averages"):
            while len(payload) > budget and len(reconciliation.get(key) or []) > 3:
                reconciliation[key] = reconciliation[key][:-1]
                meta["trimmed_reconciliation"] += 1
                payload = json.dumps(facts, ensure_ascii=False)
    if len(payload) > budget:
        facts["sample_rows"] = []
        facts["rules"] = facts["rules"][:2]
        facts["data_gaps"] = facts["data_gaps"][:5]
        if facts.get("primary_table"):
            facts["primary_table"]["headers"] = facts["primary_table"]["headers"][:10]
        payload = json.dumps(facts, ensure_ascii=False)
    if len(payload) > budget:
        # Last resort: keep only the verified figures and a note.
        compact_period = None
        if period:
            compact_period = {
                "months": period.get("months"),
                "periods": period.get("periods"),
                "main_period": period.get("main_period"),
                "comparison_periods": period.get("comparison_periods"),
                "suspicious_periods": period.get("suspicious_periods"),
                "comparison": [
                    {
                        key: item.get(key)
                        for key in (
                            "metric", "period_a", "period_b", "total_a",
                            "total_b", "percent", "months_used", "direction",
                        )
                    }
                    for item in (period.get("comparison") or [])[:6]
                ],
                "metrics": [
                    {key: entry.get(key) for key in ("metric", "totals", "best", "worst")}
                    for entry in (period.get("metrics") or [])[:6]
                ],
                "note": period.get("note"),
            }
        facts = {
            "file": facts["file"],
            "verified_calculations_from_python": facts["verified_calculations_from_python"],
            "dataset_profile": (
                {"table": dataset.get("table"), "quality": dataset.get("quality")}
                if dataset else None
            ),
            "period_analysis": compact_period,
            "request_answers": [
                {"title": item.get("title"), "value": item.get("value")}
                for item in (facts.get("request_answers") or [])[:10]
            ],
            "request_text": facts.get("request_text"),
            "statistical_reconciliation": facts.get("statistical_reconciliation"),
            "note": "بسته داده به دلیل محدودیت طول خلاصه شد؛ محاسبات تأییدشده و جمع‌های ماهانه حفظ شد.",
            "data_gaps": facts["data_gaps"][:5],
        }
        meta["degraded"] = True
        payload = json.dumps(facts, ensure_ascii=False)
    meta["chars"] = len(payload)
    meta["payload"] = payload
    return facts, meta


def trim_facts(facts: dict[str, Any], budget: int) -> tuple[dict[str, Any], dict[str, Any]]:
    """Public wrapper used by the analyzer to shrink a payload to a budget."""
    return _trim(facts, budget)


def _sheet_facts(
    analysis: AnalysisResult,
    sheet: Any,
    base: dict[str, Any],
) -> dict[str, Any]:
    intelligence = analysis.intelligence
    audits = [item for item in sheet.audits if item.status.value != "verified"]
    return {
        "file": base.get("file"),
        "workbook": {
            "sheet_count": intelligence.sheet_count if intelligence else 1,
            "status_counts": dict(intelligence.status_counts) if intelligence else {},
            "audit_counts": (
                {
                    "verified": sum(
                        1 for item in intelligence.all_audits if item.status.value == "verified"
                    ),
                    "mismatch": sum(
                        1 for item in intelligence.all_audits if item.status.value == "mismatch"
                    ),
                    "not_verifiable": sum(
                        1
                        for item in intelligence.all_audits
                        if item.status.value == "not_verifiable"
                    ),
                }
                if intelligence
                else {}
            ),
            "warnings": (intelligence.warnings[:6] if intelligence else []),
        },
        "sheet": {
            "name": sheet.sheet,
            "status": sheet.status,
            "title": sheet.title,
            "month": sheet.month,
            "sheet_years": sheet.sheet_years,
            "title_years": sheet.title_years,
            "subheader_years": sheet.subheader_years,
            "table_address": sheet.table_address,
            "header_rows": sheet.header_rows,
            "data_start": sheet.data_start,
            "data_end": sheet.data_end,
            "row_count": sheet.row_count,
            "column_count": sheet.column_count,
            "derived": sheet.derived,
            "referenced_sheets": sheet.referenced_sheets[:15],
            "coverage": sheet.coverage.to_dict() if sheet.coverage else None,
            "metadata_conflicts": [item.to_dict() for item in sheet.metadata_conflicts],
            "reconstructed_dates": sheet.reconstructed_dates[:4],
            "warnings": sheet.warnings[:8],
        },
        "verified_calculations_from_python": {},
        "dataset_profile": _generic_facts(analysis, sheet.generic),
        "sheet_audits": [item.to_dict() for item in audits[:20]],
        "sheet_period_comparisons": [
            item.to_dict() for item in sheet.period_comparisons[:20]
        ],
        "primary_table": None,
        "sample_rows": [],
        "important_anomalies": [],
        "reconciliation": [],
        "audits_of_declared_totals": [],
        "data_gaps": [],
        "rules": list(base.get("rules") or [])[:5],
    }


def build_fact_chunks(
    analysis: AnalysisResult,
    settings: Settings,
) -> list[tuple[str, dict[str, Any]]]:
    """Split the fact package into logical requests.

    A single-table workbook gets one request.  A multi-sheet workbook gets one
    workbook overview plus chunks of at most ``max_ai_sheets_per_chunk`` sheets;
    no request carries raw rows for every sheet.
    """
    facts, _meta = build_facts(analysis, settings)
    sheets = [sheet for sheet in (analysis.sheets or []) if getattr(sheet, "generic", None) is not None]
    if len(sheets) <= 1:
        # Keep the established single-request behaviour and its sample rows.
        return [("تحلیل کامل", facts)]

    overview = {key: value for key, value in facts.items() if key != "sample_rows"}
    overview["sample_rows"] = []
    overview["chunking_note"] = (
        "این بسته فقط نمای کلی Workbook است؛ جزئیات هر برگه در درخواست جداگانه ارسال می‌شود."
    )
    chunks: list[tuple[str, dict[str, Any]]] = [("نمای کلی Workbook", overview)]
    budget = max(1500, int(getattr(settings, "max_ai_context_chars", 6000)))
    max_sheets = max(1, int(getattr(settings, "max_ai_sheets_per_chunk", 3)))
    current: list[Any] = []
    current_len = 0
    for sheet in sheets:
        payload = _sheet_facts(analysis, sheet, facts)
        payload_len = len(json.dumps(payload, ensure_ascii=False))
        if current and (len(current) >= max_sheets or current_len + payload_len > budget):
            chunks.append(
                (
                    "برگه‌های " + "، ".join(item.sheet for item in current),
                    _merge_sheet_facts(analysis, current, facts),
                )
            )
            current = []
            current_len = 0
        current.append(sheet)
        current_len += payload_len
    if current:
        chunks.append(
            (
                "برگه‌های " + "، ".join(item.sheet for item in current),
                _merge_sheet_facts(analysis, current, facts),
            )
        )
    return chunks


def _merge_sheet_facts(
    analysis: AnalysisResult,
    sheets: list[Any],
    base: dict[str, Any],
) -> dict[str, Any]:
    payloads = [_sheet_facts(analysis, sheet, base) for sheet in sheets]
    return {
        "file": payloads[0].get("file"),
        "workbook": payloads[0].get("workbook"),
        "sheets": [item["sheet"] for item in payloads],
        "verified_calculations_from_python": {},
        "dataset_profile": None,
        "dataset_profiles": [item["dataset_profile"] for item in payloads],
        "sheet_audits": [audit for item in payloads for audit in item["sheet_audits"]],
        "sheet_period_comparisons": [
            item for payload in payloads for item in payload["sheet_period_comparisons"]
        ],
        "primary_table": None,
        "sample_rows": [],
        "important_anomalies": [],
        "reconciliation": [],
        "audits_of_declared_totals": [],
        "statistical_reconciliation": base.get("statistical_reconciliation"),
        "data_gaps": [],
        "rules": list(base.get("rules") or [])[:5],
        "chunking_note": "این درخواست شامل خلاصه ساختاری و آماری چند برگه است؛ داده خام ارسال نشده است.",
    }


def allowed_numbers(analysis: AnalysisResult) -> set[str]:
    """Every number the model is allowed to quote, in canonical form."""
    from local_office_ai.excel.normalization import canonical_number

    allowed: set[str] = set()
    for calculation in analysis.calculations:
        if isinstance(calculation.value, Decimal):
            allowed.add(canonical_number(calculation.value))
        elif isinstance(calculation.value, int):
            allowed.add(str(calculation.value))
        if isinstance(calculation.verification_value, Decimal):
            allowed.add(canonical_number(calculation.verification_value))
    for audit in analysis.audits:
        for value in (audit.declared, audit.computed, audit.difference):
            if isinstance(value, Decimal):
                allowed.add(canonical_number(value))
    for item in analysis.reconciliation:
        for value in (item.loan_total, item.obligation_total, item.difference):
            if isinstance(value, Decimal):
                allowed.add(canonical_number(value))
    statistical = getattr(analysis, "statistical_reconciliation", None)
    if statistical is not None:
        for item in statistical.reconciliations:
            for value in (item.expected, item.calculated, item.difference):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
            for component in item.components:
                value = component.get("value")
                if value not in (None, ""):
                    allowed.add(canonical_number(value))
        for item in statistical.comparisons:
            for value in (
                item.change.previous,
                item.change.current,
                item.change.absolute,
                item.change.percent,
            ):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
            allowed.add(str(len(item.months_used)))
        for item in statistical.averages:
            for value in (item.value, item.unweighted, item.weighted):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
        for count in statistical.counts().values():
            allowed.add(str(count))
    for loan in analysis.loans:
        for value in (loan.principal, loan.installment, loan.installment_count, loan.contract_repayment):
            if isinstance(value, Decimal):
                allowed.add(canonical_number(value))
    for item in analysis.obligations:
        allowed.add(canonical_number(item.amount))
    generic = analysis.generic
    if generic is not None:
        for value in (generic.row_count, generic.column_count):
            allowed.add(str(value))
        quality = generic.quality
        if quality is not None:
            for value in (
                quality.row_count, quality.column_count, quality.total_cells,
                quality.missing_cells, quality.duplicate_rows, quality.empty_rows,
                quality.invalid_cells,
            ):
                allowed.add(str(value))
            allowed.add(canonical_number(quality.missing_ratio))
            allowed.add(canonical_number(quality.missing_ratio * 100))
            allowed.add(canonical_number(quality.measured_missing_ratio))
            allowed.add(canonical_number(quality.measured_missing_ratio * 100))
            allowed.add(str(quality.measured_missing_cells))
        for table in generic.frequencies:
            for value in (table.total, table.distinct, table.other_count):
                allowed.add(str(value))
            for row in table.rows:
                allowed.add(str(row.count))
                allowed.add(canonical_number(row.percent))
                allowed.add(canonical_number(row.cumulative_percent))
            for bin in table.bins:
                allowed.add(str(bin.count))
                allowed.add(canonical_number(bin.percent))
                allowed.add(canonical_number(bin.lower))
                allowed.add(canonical_number(bin.upper))
        for crosstab in generic.crosstabs:
            allowed.add(str(crosstab.total))
            for row in crosstab.rows:
                allowed.add(str(row.total))
                for count in row.counts.values():
                    allowed.add(str(count))
            for count in crosstab.column_totals.values():
                allowed.add(str(count))
        for column in generic.columns:
            for value in (
                column.total, column.mean, column.median, column.minimum, column.maximum,
                column.value_range, column.variance_population, column.variance_sample,
                column.std_population, column.std_sample, column.cv_percent,
                column.q1, column.q3, column.iqr, column.p10, column.p90,
            ):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
            for value in (column.count, column.missing, column.distinct):
                if isinstance(value, int):
                    allowed.add(str(value))
            for date_text in (column.date_min, column.date_max):
                if date_text:
                    for part in str(date_text).replace("-", "/").split("/"):
                        if part.strip().isdigit():
                            allowed.add(str(int(part)))
        for comparison in generic.groups:
            for group in comparison.groups:
                allowed.add(str(group.count))
                for value in (group.total, group.mean, group.median, group.minimum, group.maximum, group.std_sample):
                    if isinstance(value, Decimal):
                        allowed.add(canonical_number(value))
        for trend in generic.trends:
            if isinstance(trend.change_absolute, Decimal):
                allowed.add(canonical_number(trend.change_absolute))
            if isinstance(trend.change_percent, Decimal):
                allowed.add(canonical_number(trend.change_percent))
            for point in trend.points:
                allowed.add(str(point.count))
                for part in str(point.period_label).replace("-", "/").split("/"):
                    if part.strip().isdigit():
                        allowed.add(str(int(part)))
                for value in (point.total, point.mean):
                    allowed.add(canonical_number(value))
        for pair in generic.correlations:
            allowed.add(str(pair.count))
            for value in (pair.pearson, pair.spearman):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
        for integration in [getattr(analysis, "integration", None)]:
            if integration is None:
                continue
            for period in integration.periods:
                allowed.add(str(period))
                allowed.add(str(len(integration.period_months(period))))
            for month in integration.months:
                allowed.add(str(len(integration.period_months(month))))
            for series in integration.series:
                for cell in series.cells:
                    if isinstance(cell.total, Decimal):
                        allowed.add(canonical_number(cell.total))
                    allowed.add(str(cell.observations))
                for period in integration.periods:
                    total = series.total(period)
                    if isinstance(total, Decimal):
                        allowed.add(canonical_number(total))
                    mean = series.mean(period)
                    if isinstance(mean, Decimal):
                        allowed.add(canonical_number(mean))
                for index in range(len(integration.periods) - 1):
                    change = series.compare(
                        integration.periods[index], integration.periods[index + 1]
                    )
                    for value in (change.total_a, change.total_b, change.difference, change.percent):
                        if isinstance(value, Decimal):
                            allowed.add(canonical_number(value))
                for month, _value, difference, percent in series.month_changes(
                    integration.main_period()
                ):
                    if isinstance(difference, Decimal):
                        allowed.add(canonical_number(difference))
                    if isinstance(percent, Decimal):
                        allowed.add(canonical_number(percent))
            allowed.add(str(len(integration.months)))
            allowed.add(str(len(integration.periods)))
            allowed.add(str(len(integration.series)))
        for outlier in generic.outliers:
            allowed.add(canonical_number(outlier.value))
    for sheet in analysis.sheets or []:
        if getattr(sheet, "generic", None) is not None:
            _extend_generic_numbers(sheet.generic, allowed)
        for value in (sheet.row_count, sheet.column_count):
            allowed.add(str(value))
        for value in (sheet.sheet_years, sheet.title_years, sheet.subheader_years):
            for year in value:
                if str(year).isdigit():
                    allowed.add(str(int(year)))
        coverage = sheet.coverage
        if coverage is not None:
            for value in (
                coverage.expected_observations,
                coverage.observed_observations,
                coverage.missing_observations,
                coverage.explicit_zero_cells,
                coverage.blank_cells,
                coverage.invalid_cells,
                coverage.error_cells,
                coverage.pending_rows,
            ):
                if isinstance(value, int):
                    allowed.add(str(value))
            for label in coverage.missing_labels:
                if str(label).isdigit():
                    allowed.add(str(int(label)))
        for audit in sheet.audits:
            for value in (audit.declared, audit.computed, audit.difference):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
            for part in str(audit.address).replace("$", "").split(":"):
                digits = "".join(char for char in part if char.isdigit())
                if digits:
                    allowed.add(str(int(digits)))
        for item in sheet.period_comparisons:
            for value in (item.total_a, item.total_b, item.mean_a, item.mean_b, item.difference, item.percent):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
            for value in (item.count_a, item.count_b):
                allowed.add(str(value))
            for period in (item.period_a, item.period_b):
                if str(period).isdigit():
                    allowed.add(str(int(period)))
        for item in sheet.reconstructed_dates:
            for key in ("period", "first_day", "last_day"):
                if str(item.get(key, "")).isdigit():
                    allowed.add(str(int(item[key])))
    intelligence = analysis.intelligence
    if intelligence is not None:
        for item in intelligence.all_audits:
            for value in (item.declared, item.computed, item.difference):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
        for item in intelligence.period_comparisons:
            for value in (item.total_a, item.total_b, item.mean_a, item.mean_b, item.difference, item.percent):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
            for value in (item.count_a, item.count_b):
                allowed.add(str(value))
            for period in (item.period_a, item.period_b):
                if str(period).isdigit():
                    allowed.add(str(int(period)))
        for item in intelligence.aggregations:
            if isinstance(item.total, Decimal):
                allowed.add(canonical_number(item.total))
            allowed.add(str(item.observations))
            if str(item.period).isdigit():
                allowed.add(str(int(item.period)))
        counts = intelligence.status_counts
        for value in counts.values():
            allowed.add(str(value))
    allowed.add("0")
    return allowed


def _extend_generic_numbers(generic: Any, allowed: set[str]) -> None:
    for value in (generic.row_count, generic.column_count):
        allowed.add(str(value))
    quality = generic.quality
    if quality is not None:
        for value in (
            quality.row_count,
            quality.column_count,
            quality.total_cells,
            quality.missing_cells,
            quality.duplicate_rows,
            quality.empty_rows,
            quality.invalid_cells,
        ):
            allowed.add(str(value))
        allowed.add(canonical_number(quality.missing_ratio))
        allowed.add(canonical_number(quality.missing_ratio * 100))
    for table in generic.frequencies:
        for value in (table.total, table.distinct, table.other_count):
            allowed.add(str(value))
        for item in table.rows:
            allowed.add(str(item.count))
            allowed.add(canonical_number(item.percent))
        for bin in table.bins:
            allowed.add(str(bin.count))
            allowed.add(canonical_number(bin.lower))
            allowed.add(canonical_number(bin.upper))
    for column in generic.columns:
        for value in (
            column.total,
            column.mean,
            column.median,
            column.minimum,
            column.maximum,
            column.value_range,
            column.variance_population,
            column.variance_sample,
            column.std_population,
            column.std_sample,
            column.cv_percent,
            column.q1,
            column.q3,
            column.iqr,
            column.p10,
            column.p90,
        ):
            if isinstance(value, Decimal):
                allowed.add(canonical_number(value))
        for value in (column.count, column.missing, column.distinct):
            if isinstance(value, int):
                allowed.add(str(value))
        for date_text in (column.date_min, column.date_max):
            if date_text:
                for part in str(date_text).replace("-", "/").split("/"):
                    if part.strip().isdigit():
                        allowed.add(str(int(part)))
    for comparison in generic.groups:
        for group in comparison.groups:
            allowed.add(str(group.count))
            for value in (group.total, group.mean, group.median, group.minimum, group.maximum, group.std_sample):
                if isinstance(value, Decimal):
                    allowed.add(canonical_number(value))
    for trend in generic.trends:
        for value in (trend.change_absolute, trend.change_percent):
            if isinstance(value, Decimal):
                allowed.add(canonical_number(value))
        for point in trend.points:
            allowed.add(str(point.count))
            for part in str(point.period_label).replace("-", "/").split("/"):
                if part.strip().isdigit():
                    allowed.add(str(int(part)))
            for value in (point.total, point.mean):
                allowed.add(canonical_number(value))
    for pair in generic.correlations:
        allowed.add(str(pair.count))
        for value in (pair.pearson, pair.spearman):
            if isinstance(value, Decimal):
                allowed.add(canonical_number(value))
    for outlier in generic.outliers:
        allowed.add(canonical_number(outlier.value))
