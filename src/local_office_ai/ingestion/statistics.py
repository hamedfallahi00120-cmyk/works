"""Deterministic numeric profiling for extracted document tables."""

from __future__ import annotations

from decimal import Decimal
from statistics import median
from typing import Any

from local_office_ai.excel.normalization import parse_number

_IDENTIFIER_WORDS = ("شناسه", "کد", "ردیف", "شماره", "سریال", "پلاک", "کدپستی", "id", "identifier", "serial", "code", "index", "no.")
_TEMPORAL_WORDS = ("تاریخ", "روز", "ماه", "سال", "زمان", "ساعت", "date", "day", "month", "year", "time")


def _excluded_numeric_role(header: str) -> bool:
    value = str(header or "").strip().casefold().replace("ي", "ی").replace("ك", "ک")
    compact = "".join(char for char in value if char.isalnum() or "\u0600" <= char <= "\u06ff")
    return any(token in value for token in _IDENTIFIER_WORDS + _TEMPORAL_WORDS) or any(token in compact for token in ("شناسه", "تاریخ", "شماره", "ردیف", "کد"))


def profile_table(source: str, location: str, title: str, headers: list[str], rows: list[list[Any]]) -> list[dict[str, Any]]:
    profiles: list[dict[str, Any]] = []
    if not headers or not rows:
        return profiles
    width = max(len(headers), max((len(row) for row in rows), default=0))
    safe_headers = list(headers) + [f"ستون {i + 1}" for i in range(len(headers), width)]
    for index in range(width):
        header = safe_headers[index]
        if _excluded_numeric_role(header):
            continue
        values = [row[index] if index < len(row) else None for row in rows]
        nonempty = [value for value in values if value is not None and str(value).strip()]
        if not nonempty:
            continue
        parsed: list[Decimal] = []
        for value in nonempty:
            normalized = parse_number(value, numeric_context=True)
            if normalized.value is not None and isinstance(normalized.value, Decimal):
                parsed.append(normalized.value)
        ratio = len(parsed) / len(nonempty) if nonempty else 0.0
        # Only profile columns that are predominantly numeric; IDs and dates are filtered later by intent.
        if len(parsed) < 2 or ratio < 0.65:
            continue
        ordered = sorted(parsed)
        avg = sum(ordered, Decimal(0)) / Decimal(len(ordered))
        med = Decimal(str(median(ordered)))
        if len(ordered) > 1:
            variance = sum(((value - avg) ** 2 for value in ordered), Decimal(0)) / Decimal(len(ordered) - 1)
            std = variance.sqrt() if variance >= 0 else None
        else:
            variance = None
            std = None
        profiles.append(
            {
                "source": source,
                "location": location,
                "table": title,
                "column": safe_headers[index],
                "n": len(parsed),
                "missing": len(values) - len(parsed),
                "mean": avg,
                "median": med,
                "std_sample": std,
                "min": min(ordered),
                "max": max(ordered),
                "sum": sum(ordered, Decimal(0)),
            }
        )
    return profiles


def profile_excel_analysis(source: str, analysis: Any) -> list[dict[str, Any]]:
    profiles: list[dict[str, Any]] = []
    generic = getattr(analysis, "generic", None)
    if generic is None:
        return profiles
    for column in getattr(generic, "columns", []) or []:
        if not getattr(column, "is_numeric", False):
            continue
        if getattr(column, "is_identifier", False) or getattr(column, "is_temporal_key", False):
            continue
        profiles.append(
            {
                "source": source,
                "location": f"{getattr(generic, 'table_sheet', '')}!{getattr(generic, 'table_address', '')}",
                "table": "جدول اصلی Excel",
                "column": getattr(column, "header", "ستون"),
                "n": getattr(column, "count", 0),
                "missing": getattr(column, "missing", 0),
                "mean": getattr(column, "mean", None),
                "median": getattr(column, "median", None),
                "std_sample": getattr(column, "std_sample", None),
                "min": getattr(column, "minimum", None),
                "max": getattr(column, "maximum", None),
                "sum": getattr(column, "total", None),
            }
        )
    return profiles
