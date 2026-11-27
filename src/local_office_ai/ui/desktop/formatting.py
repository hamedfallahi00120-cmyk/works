"""Display formatting helpers shared by the desktop widgets.

Numbers always come from the engine as ``Decimal``/int; this module only
rounds for presentation and never feeds a formatted value back into a
calculation.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any

from PySide6.QtCore import Qt

UNAVAILABLE = "در دسترس نیست"
_LTR_EMBED = "\u202a"
_POP_DIRECTIONAL = "\u202c"
_NUMERIC = re.compile(r"^-?\d+(?:\.\d+)?$")
_LONG_FRACTION = re.compile(r"(?<![\d.,])(\d[\d,]*\.\d{6,})(?![\d.,])")


def format_number(value: Any, *, decimals: int = 2) -> str:
    if value is None or value == "":
        return UNAVAILABLE
    try:
        if isinstance(value, bool):
            return "بله" if value else "خیر"
        if isinstance(value, Decimal):
            amount = value
        elif isinstance(value, (int, float)):
            amount = Decimal(str(value))
        else:
            amount = Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return str(value)
    if amount == amount.to_integral_value():
        return f"{int(amount):,}"
    return f"{amount:,.{decimals}f}".rstrip("0").rstrip(".")


def format_percent(value: Any, *, decimals: int = 1) -> str:
    if value is None:
        return UNAVAILABLE
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if -1.0 <= number <= 1.0 and not float(number).is_integer():
        number *= 100
    return f"{number:.{decimals}f}%"


def format_bytes(value: Any) -> str:
    if value is None:
        return UNAVAILABLE
    try:
        size = float(value)
    except (TypeError, ValueError):
        return str(value)
    units = ("B", "KB", "MB", "GB", "TB")
    index = 0
    while size >= 1024 and index < len(units) - 1:
        size /= 1024
        index += 1
    if index == 0:
        return f"{int(size)} {units[index]}"
    return f"{size:,.1f} {units[index]}"


def format_mb(value: Any) -> str:
    if value is None:
        return UNAVAILABLE
    try:
        amount = float(value)
    except (TypeError, ValueError):
        return str(value)
    if amount >= 1024:
        return f"{amount / 1024:,.1f} GB"
    return f"{amount:,.0f} MB"


def format_duration(seconds: float | None) -> str:
    if seconds is None:
        return ""
    if seconds < 60:
        return f"{seconds:.1f} ثانیه"
    minutes = int(seconds // 60)
    remainder = int(seconds % 60)
    if minutes < 60:
        return f"{minutes} دقیقه و {remainder} ثانیه"
    hours = minutes // 60
    return f"{hours} ساعت و {minutes % 60} دقیقه"


def format_timestamp(value: Any) -> str:
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, (int, float)):
        return datetime.fromtimestamp(value).strftime("%Y-%m-%d %H:%M")
    return str(value or "")


def format_file_size(path: str) -> str:
    try:
        from pathlib import Path

        return format_bytes(Path(path).stat().st_size)
    except OSError:
        return UNAVAILABLE


def severity_label(value: Any) -> str:
    raw = getattr(value, "value", value)
    return {"critical": "بحرانی", "warning": "هشدار", "info": "اطلاع"}.get(str(raw), str(raw))


def severity_tone(value: Any) -> str:
    raw = getattr(value, "value", value)
    return {"critical": "danger", "warning": "warning", "info": "info"}.get(str(raw), "info")


def validation_tone(status: Any) -> str:
    raw = getattr(status, "value", status)
    return {
        "passed": "success",
        "warning": "warning",
        "failed": "danger",
        "not_applicable": "muted",
        "skipped": "muted",
    }.get(str(raw), "muted")


def validation_label(status: Any) -> str:
    raw = getattr(status, "value", status)
    return {
        "passed": "قبول",
        "warning": "هشدار",
        "failed": "رد",
        "not_applicable": "بی‌ربط",
        "skipped": "انجام نشد",
    }.get(str(raw), str(raw))


def ltr(text: Any) -> str:
    """Embed a path or timestamp so an RTL label keeps its reading order."""
    return f"{_LTR_EMBED}{text}{_POP_DIRECTIONAL}"


def rtl(text: Any) -> str:
    """Force a right-to-left paragraph when a sentence starts with a Latin word."""
    return f"\u200f{text}"


def display_value(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if _NUMERIC.match(text):
        return format_number(Decimal(text))
    return text


def soften_numbers(text: Any) -> str:
    """Round only long decimal fractions in prose for readability."""
    if text is None:
        return ""

    def replace(match: re.Match[str]) -> str:
        try:
            return format_number(Decimal(match.group(1).replace(",", "")))
        except InvalidOperation:
            return match.group(1)

    return _LONG_FRACTION.sub(replace, str(text))


def format_cell(value: Any, *, limit: int = 240) -> str:
    text = " ".join(str(value or "").split())
    if _NUMERIC.match(text):
        text = format_number(Decimal(text))
    text = soften_numbers(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def align_table_header(table: Any) -> None:
    """Right-align header captions; Qt does not mirror them inside RTL sections."""
    table.horizontalHeader().setHighlightSections(False)
    for index in range(table.columnCount()):
        item = table.horizontalHeaderItem(index)
        if item is not None:
            item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute | Qt.AlignmentFlag.AlignVCenter)


def elide(text: str, limit: int = 90) -> str:
    clean = " ".join(str(text or "").split())
    return clean if len(clean) <= limit else clean[: limit - 1] + "…"


def mode_label(mode: Any) -> str:
    return "مدل محلی" if str(mode) == "model" else "قواعد قطعی Python"


def pretty_path(path: Any, limit: int = 52) -> str:
    text = str(path or "")
    if len(text) <= limit:
        return text
    return "…" + text[-(limit - 1):]
