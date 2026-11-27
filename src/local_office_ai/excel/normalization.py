"""Standardization of real-world Persian/Arabic/English spreadsheet values.

Design rules
------------
* The raw value is never destroyed; every conversion returns a
  :class:`NormalizedValue` that keeps the original next to the standardized form.
* Ambiguous conversions are never silent: they add a warning and, where the
  interpretation changes the meaning (day/month swap, thousands separator),
  they say which rule was applied.
* Money is produced as :class:`decimal.Decimal`, never as binary float.
"""

from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from local_office_ai.core.models import NormalizedValue, Provenance, ValueKind

PERSIAN_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
ARABIC_DIGITS = "٠١٢٣٤٥٦٧٨٩"
_DIGIT_TRANSLATION = {ord(char): str(index) for index, char in enumerate(PERSIAN_DIGITS)}
_DIGIT_TRANSLATION.update({ord(char): str(index) for index, char in enumerate(ARABIC_DIGITS)})

# ZWNJ (\u200c) is handled separately: it becomes a space, not a deletion.
INVISIBLE_CHARS = (
    "\u200b\u200d\u200e\u200f\u202a\u202b\u202c\u202d\u202e"
    "\u2066\u2067\u2068\u2069\ufeff\u00ad\u2060"
)
_ZWNJ = "\u200c"

THOUSAND_SEPARATORS = {",", "٬", "،", "'", "’", " ", "\u00a0", "\u202f", "_"}
DECIMAL_SEPARATORS = {".", "٫"}

CURRENCY_TOKENS = (
    "تومان", "تومن", "ریال", "ريال", "درهم", "دلار", "یورو",
    "tooman", "toman", "rial", "irr", "usd", "eur", "$", "€", "£",
)

# Scale words carry a magnitude, not a currency. They must multiply the parsed
# figure, otherwise "۲ میلیون تومان" silently becomes 2 instead of 2,000,000.
SCALE_UNITS: dict[str, Decimal] = {
    "هزار": Decimal("1000"),
    "thousand": Decimal("1000"),
    "میلیون": Decimal("1000000"),
    "million": Decimal("1000000"),
    "میلیارد": Decimal("1000000000"),
    "billion": Decimal("1000000000"),
    "همت": Decimal("1000000000000"),
    "تریلیون": Decimal("1000000000000"),
    "trillion": Decimal("1000000000000"),
}
_SCALE_WORD_RE = re.compile(
    r"(?<!\w)(" + "|".join(sorted(SCALE_UNITS, key=len, reverse=True)) + r")(?!\w)",
    re.IGNORECASE,
)
_SCALE_SUFFIX_RE = re.compile(
    r"^\s*(" + "|".join(sorted(SCALE_UNITS, key=len, reverse=True)) + r")(?!\w)",
    re.IGNORECASE,
)

# Conservative default entity synonyms. They are used ONLY to build a matching
# key; the source text is never rewritten.
DEFAULT_SYNONYMS: dict[str, str] = {"صندق": "صندوق"}


def load_synonyms(path: Path | str | None) -> dict[str, str]:
    """Read an optional ``{"wrong": "correct"}`` map used only for matching."""
    if not path:
        return {}
    target = Path(path)
    if not target.is_file():
        return {}
    try:
        data = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {
        str(key): str(value)
        for key, value in data.items()
        if not str(key).startswith("_") and isinstance(value, str)
    }


def _detect_scale(text: str) -> tuple[str, Decimal | None]:
    """Remove the first scale word from ``text`` and return its multiplier."""
    match = _SCALE_WORD_RE.search(text)
    if match is None:
        return text, None
    word = match.group(1)
    multiplier = SCALE_UNITS.get(word) or SCALE_UNITS.get(word.lower())
    if multiplier is None:
        return text, None
    return _SCALE_WORD_RE.sub(" ", text, count=1), multiplier


_SCALE_ALT = "|".join(sorted(SCALE_UNITS, key=len, reverse=True))
_SCALED_PART_RE = re.compile(
    rf"(?P<num>[-+]?\d[\d\s.,٬٫'’_]*\d|[-+]?\d)\s*(?P<scale>{_SCALE_ALT})?",
    re.IGNORECASE,
)
_NEGATIVE_WORDS = ("منفی", "بدهکار")


def _parse_scaled_sum(text: str) -> tuple[Decimal, int] | None:
    """Parse compound Persian amounts such as ``۱۲ هزار و ۵۰۰``.

    Only pure numeric expressions qualify; prose falls back to single-token
    extraction so a scale word is never applied to an unrelated sentence.
    """
    cleaned = text.strip()
    if not _SCALE_WORD_RE.search(cleaned):
        return None
    total = Decimal(0)
    position = 0
    parts = 0
    while position < len(cleaned):
        match = _SCALED_PART_RE.match(cleaned, position)
        if match is None:
            return None
        token, scale = match.group("num").strip(), match.group("scale")
        numeric_text, _, _ = _detect_separators(token)
        if not re.fullmatch(r"[-+]?\d+(\.\d+)?", numeric_text):
            return None
        amount = Decimal(numeric_text)
        if scale:
            multiplier = SCALE_UNITS.get(scale) or SCALE_UNITS.get(scale.lower())
            if multiplier is None:
                return None
            amount *= multiplier
        total += amount
        parts += 1
        position = match.end()
        connector = re.match(r"\s*(?:و\s*)?", cleaned[position:])
        position += connector.end() if connector else 0
        if position < len(cleaned) and not re.match(r"[-+]?\d", cleaned[position:]):
            return None
    if parts < 2:
        return None
    return total, parts

EXCEL_ERRORS = {
    "#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NULL!", "#NUM!",
    "#GETTING_DATA", "#SPILL!", "#CALC!",
}

_JALALI_MONTHS = (
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
)
_GREGORIAN_MONTHS = (
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
)

_NUMBER_TOKEN = re.compile(r"[-+]?\d[\d\s.,٬٫'’_]*\d|[-+]?\d")
_DATE_SEPARATED = re.compile(r"^\s*(\d{1,4})\s*[/\-.\\]\s*(\d{1,2})\s*[/\-.\\]\s*(\d{1,4})\s*$")
_DATE_WITH_MONTH_NAME = re.compile(r"^\s*(\d{1,2})\s*([^\W\d_]+)\s*(\d{2,4})\s*$", re.UNICODE)


def _div(a: int, b: int) -> int:
    return a // b


def jalali_to_gregorian(jy: int, jm: int, jd: int) -> tuple[int, int, int]:
    """Convert a Jalali (Solar Hijri) date to Gregorian (33-year cycle arithmetic)."""
    jy += 1595
    days = -355668 + (365 * jy) + (_div(jy, 33) * 8) + _div((jy % 33) + 3, 4) + jd
    days += (jm - 1) * 31 if jm < 7 else ((jm - 7) * 30) + 186
    gy = 400 * _div(days, 146097)
    days %= 146097
    if days > 36524:
        days -= 1
        gy += 100 * _div(days, 36524)
        days %= 36524
        if days >= 365:
            days += 1
    gy += 4 * _div(days, 1461)
    days %= 1461
    if days > 365:
        gy += _div(days - 1, 365)
        days = (days - 1) % 365
    gd = days + 1
    leap = (gy % 4 == 0 and gy % 100 != 0) or gy % 400 == 0
    month_days = [0, 31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    gm = 1
    while gm < 13 and gd > month_days[gm]:
        gd -= month_days[gm]
        gm += 1
    return gy, gm, gd


def gregorian_to_jalali(gy: int, gm: int, gd: int) -> tuple[int, int, int]:
    """Convert a Gregorian date to Jalali."""
    g_days_in_month = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    j_days_in_month = [0, 31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
    gy2, _, gd2 = gy - 1600, gm - 1, gd - 1
    g_day_no = (
        365 * gy2
        + (gy2 + 3) // 4
        - (gy2 + 99) // 100
        + (gy2 + 399) // 400
        + gd2
        + sum(g_days_in_month[1:gm])
    )
    if gm > 2 and ((gy % 4 == 0 and gy % 100 != 0) or gy % 400 == 0):
        g_day_no += 1
    j_day_no = g_day_no - 79
    j_np = j_day_no // 12053
    j_day_no %= 12053
    jy = 979 + 33 * j_np + 4 * (j_day_no // 1461)
    j_day_no %= 1461
    if j_day_no >= 366:
        jy += (j_day_no - 1) // 365
        j_day_no = (j_day_no - 1) % 365
    jm = 1
    while jm <= 12 and j_day_no >= j_days_in_month[jm]:
        j_day_no -= j_days_in_month[jm]
        jm += 1
    jd = j_day_no + 1
    return jy, jm, jd


def is_valid_jalali(jy: int, jm: int, jd: int) -> bool:
    if not (1178 <= jy <= 1633 and 1 <= jm <= 12 and 1 <= jd <= 31):
        return False
    gy, gm, gd = jalali_to_gregorian(jy, jm, jd)
    return gregorian_to_jalali(gy, gm, gd) == (jy, jm, jd)


def format_jalali(value: date | datetime) -> str:
    """Render a Gregorian date as ``YYYY/MM/DD`` in the Jalali calendar."""
    if isinstance(value, datetime):
        value = value.date()
    jy, jm, jd = gregorian_to_jalali(value.year, value.month, value.day)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"


def excel_serial_to_date(serial: float) -> date:
    """Convert an Excel 1900-system serial number to a date."""
    whole = int(serial)
    if whole < 60:
        return date(1899, 12, 31) + timedelta(days=whole)
    return date(1899, 12, 30) + timedelta(days=whole)


def normalize_digits(text: str) -> str:
    """Map Persian and Arabic-Indic digits to ASCII digits."""
    return str(text).translate(_DIGIT_TRANSLATION)


def strip_invisible(text: str) -> str:
    return "".join(char for char in str(text) if char not in INVISIBLE_CHARS)


def unify_letters(text: str) -> str:
    """Unify Arabic/Persian letter variants that mean the same letter."""
    # Only true Arabic/Persian variants are unified.  ``ئ`` and ``ؤ`` are
    # legitimate Persian letters and are left untouched so display text keeps
    # its correct spelling.
    translation = {
        "ي": "ی", "ى": "ی", "ﻯ": "ی", "ك": "ک", "ﮐ": "ک", "ﮏ": "ک",
        "أ": "ا", "إ": "ا", "ٱ": "ا", "ة": "ه", "ۀ": "ه",
    }
    return "".join(translation.get(char, char) for char in str(text))


def normalize_text(text: str, *, keep_zwnj: bool = False) -> str:
    """Standardize whitespace and invisible characters but keep the wording."""
    value = strip_invisible(text)
    if not keep_zwnj:
        value = value.replace(_ZWNJ, " ")
    value = unify_letters(value)
    value = re.sub(r"[\t\r\n]+", " ", value)
    value = re.sub(r" {2,}", " ", value)
    return value.strip()


_LEGAL_WORDS = (
    "بانک", "صندوق", "مؤسسه", "موسسه", "شرکت", "قرض الحسنه", "قرض‌الحسنه",
    "اعتباری", "مالی", "تعاونی", "ltd", "co", "company", "bank", "fund",
)


def normalize_institution(text: Any, synonyms: dict[str, str] | None = None) -> str:
    """Build a comparison key for an institution name.

    Unifies characters and drops legal-form words so ``بانك مهر ايران`` and
    ``مهر ایران`` compare equal.  Configured synonyms are applied to the key
    only; the displayed/source value is never modified.
    """
    value = normalize_text(normalize_digits(str(text or "")), keep_zwnj=False)
    mapping = dict(DEFAULT_SYNONYMS)
    mapping.update(synonyms or {})
    for source, target in mapping.items():
        if source and source in value:
            value = value.replace(source, target)
    value = re.sub(r"[.,;:!?()\[\]{}«»\"'`ـ_/\\|+*&^%$#@~=<>]", " ", value)
    value = re.sub(r" {2,}", " ", value).strip()
    lowered = value.lower()
    for word in _LEGAL_WORDS:
        lowered = re.sub(rf"(^|\s){re.escape(word)}(\s|$)", " ", lowered)
    lowered = re.sub(r" {2,}", " ", lowered).strip()
    return lowered or value.lower().strip()


def _bigrams(text: str) -> set[str]:
    compact = re.sub(r"\s+", "", text)
    if len(compact) < 2:
        return {compact} if compact else set()
    return {compact[index : index + 2] for index in range(len(compact) - 1)}


def institution_similarity(left: str, right: str, synonyms: dict[str, str] | None = None) -> float:
    """Similarity in [0, 1] from token overlap, character bigrams and containment."""
    key_left = normalize_institution(left, synonyms)
    key_right = normalize_institution(right, synonyms)
    if not key_left or not key_right:
        return 0.0
    if key_left == key_right:
        return 1.0

    tokens_left = set(key_left.split())
    tokens_right = set(key_right.split())
    token_score = len(tokens_left & tokens_right) / max(min(len(tokens_left), len(tokens_right)), 1)

    grams_left = _bigrams(key_left)
    grams_right = _bigrams(key_right)
    gram_score = len(grams_left & grams_right) / max(len(grams_left | grams_right), 1)

    # Sequence ratio copes with a single typo inside an otherwise identical name,
    # which token overlap alone would punish too hard.
    sequence_score = SequenceMatcher(None, key_left, key_right).ratio()
    containment = 1.0 if (key_left in key_right or key_right in key_left) else 0.0
    return max(
        sequence_score,
        0.5 * token_score + 0.5 * gram_score,
        0.7 * sequence_score + 0.3 * token_score,
        containment * 0.9,
    )


def _detect_separators(text: str) -> tuple[str, str | None, list[str]]:
    """Return (numeric_text_with_dot, thousands_char, warnings)."""
    warnings: list[str] = []
    has_dot = "." in text
    comma_like = [char for char in text if char in {",", "٬", "،", "'", "’", "_"}]
    space_like = any(char in {" ", "\u00a0", "\u202f"} for char in text)

    if "٫" in text:
        text = text.replace("٫", ".")
        has_dot = True

    if has_dot and comma_like:
        dot_pos = text.rfind(".")
        comma_pos = max(text.rfind(char) for char in set(comma_like))
        if comma_pos > dot_pos:
            text = re.sub(r"[.٬،'’_]", "", text)
            text = text.replace(",", ".", 1)
            text = text.replace(",", "")
            warnings.append("comma used as decimal separator")
            return text, ",", warnings
        text = re.sub(r"[٬،'’_]", "", text)
        return text.replace(",", ""), ",", warnings

    if comma_like and not has_dot:
        char = comma_like[0]
        parts = text.split(char)
        groups = parts[1:]
        if len(parts) == 2 and groups and len(groups[0]) != 3:
            warnings.append("single comma treated as decimal separator")
            return text.replace(char, "."), char, warnings
        if groups and all(len(group) == 3 for group in groups):
            return re.sub(r"[,٬،'’_]", "", text), char, warnings
        warnings.append("comma grouping is irregular; treated as decimal separator")
        return text.replace(char, "."), char, warnings

    if has_dot:
        parts = text.split(".")
        if len(parts) > 2:
            if all(len(group) == 3 for group in parts[1:]):
                warnings.append("multiple dots with 3-digit groups treated as thousands")
                return text.replace(".", ""), ".", warnings
            warnings.append("multiple dots; only the last is treated as decimal")
            return f"{''.join(parts[:-1])}.{parts[-1]}", ".", warnings
        integer_part, _, fraction = text.partition(".")
        if len(fraction) == 3 and 1 <= len(integer_part.lstrip("+-")) <= 3:
            warnings.append("ambiguous '1.234' read as decimal; verify the source locale")
        return text, ".", warnings

    if space_like:
        stripped = re.sub(r"[\s\u00a0\u202f]", "", text)
        if re.fullmatch(r"[-+]?\d+", stripped):
            warnings.append("space used as thousands separator")
            return stripped, " ", warnings

    return text, None, warnings


def parse_number(
    raw: Any,
    *,
    numeric_context: bool = False,
    provenance: Provenance | None = None,
) -> NormalizedValue:
    """Parse a number from any supported representation into ``Decimal``."""
    warnings: list[str] = []
    method = "direct"

    if raw is None:
        return NormalizedValue(raw, None, ValueKind.EMPTY, "empty", provenance=provenance)
    if isinstance(raw, bool):
        return NormalizedValue(raw, None, ValueKind.BOOLEAN, "not-a-number", provenance=provenance)
    if isinstance(raw, Decimal):
        return NormalizedValue(raw, raw, ValueKind.NUMBER, "decimal", provenance=provenance)
    if isinstance(raw, int):
        return NormalizedValue(raw, Decimal(raw), ValueKind.NUMBER, "integer", provenance=provenance)
    if isinstance(raw, float):
        if raw != raw or raw in (float("inf"), float("-inf")):
            return NormalizedValue(raw, None, ValueKind.ERROR, "non-finite float", ["value is not finite"], provenance)
        return NormalizedValue(raw, Decimal(repr(raw)), ValueKind.NUMBER, "float", provenance=provenance)
    if isinstance(raw, (datetime, date)):
        return NormalizedValue(raw, None, ValueKind.DATE, "not-a-number", ["value is a date"], provenance)

    text = strip_invisible(normalize_digits(str(raw))).strip()
    if not text:
        return NormalizedValue(raw, None, ValueKind.EMPTY, "empty", provenance=provenance)
    if text.startswith("="):
        return NormalizedValue(
            raw, None, ValueKind.TEXT, "formula-without-value",
            ["متن فرمول است و مقدار محاسبه‌شده ندارد"], provenance,
        )
    if text.upper() in EXCEL_ERRORS:
        return NormalizedValue(raw, None, ValueKind.ERROR, "excel-error", [f"Excel error value {text}"], provenance)

    negative = False
    negative_word = False
    for word in _NEGATIVE_WORDS:
        if text == word or text.startswith(f"{word} "):
            negative = True
            negative_word = True
            text = text[len(word):].strip()
            method = "negative-word"
            break
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()
        method = "parentheses-negative"
    elif "(" in text and ")" in text:
        negative = True
        text = text.replace("(", "").replace(")", "").strip()
        method = "parentheses-negative"

    lowered = text.lower()
    unit = ""
    for token in CURRENCY_TOKENS:
        if token in lowered:
            unit = token
            lowered = lowered.replace(token, " ")
            break
    percent = "%" in lowered or "٪" in lowered
    if percent:
        lowered = lowered.replace("%", " ").replace("٪", " ")

    scaled_sum = _parse_scaled_sum(lowered)
    if scaled_sum is not None:
        total, parts = scaled_sum
        if negative:
            total = -total
            warnings.append("negative word applied")
        if percent:
            total = total / Decimal(100)
            warnings.append("percent converted to fraction")
        if unit:
            warnings.append(f"currency unit removed: {unit}")
        warnings.append(f"compound scale expression summed from {parts} parts")
        return NormalizedValue(raw, total, ValueKind.NUMBER, "scaled-unit-sum", warnings, provenance)

    lowered, scale_multiplier = _detect_scale(lowered)
    if scale_multiplier is not None:
        warnings.append(f"scale unit applied: x{scale_multiplier}")
        method = "scaled-unit"
        if negative_word:
            warnings.append("negative word applied")
            method = "negative-word-scaled-unit"

    lowered = re.sub(r"[−–—](?=\d)", "-", lowered)
    match = _NUMBER_TOKEN.search(lowered)
    if not match:
        return NormalizedValue(raw, None, ValueKind.TEXT, "no-number-token", ["no numeric token found"], provenance)
    token = match.group(0)
    if token != lowered.strip():
        warnings.append(f"number extracted from text: {str(raw)[:60]!r}")
        method = "embedded-in-text"

    sign = ""
    if token.startswith(("-", "+")):
        sign = token[0]
        token = token[1:]
    if any(char in token for char in "−–—"):
        sign = "-"
        for char in "−–—":
            token = token.replace(char, "")
    if negative:
        sign = "-"
    token = token.strip()

    cleaned, separator, sep_warnings = _detect_separators(token)
    warnings.extend(sep_warnings)
    if not re.fullmatch(r"\d+(\.\d+)?", cleaned):
        return NormalizedValue(
            raw, None, ValueKind.TEXT, "unparseable", warnings + ["unparseable numeric token"], provenance
        )

    integer_part = cleaned.split(".")[0]
    if (
        not numeric_context
        and separator is None
        and len(integer_part) > 1
        and integer_part.startswith("0")
        and "." not in cleaned
    ):
        return NormalizedValue(
            raw,
            normalize_text(normalize_digits(str(raw))),
            ValueKind.TEXT,
            "leading-zero-identifier",
            warnings + ["leading zeros kept: value treated as identifier, not amount"],
            provenance,
        )

    try:
        value = Decimal(f"{sign}{cleaned}")
    except InvalidOperation:
        return NormalizedValue(raw, None, ValueKind.ERROR, "invalid", warnings + ["Decimal conversion failed"], provenance)

    if scale_multiplier is not None:
        value = value * scale_multiplier
    if percent:
        value = value / Decimal(100)
        warnings.append("percent converted to fraction")
    if unit:
        warnings.append(f"currency unit removed: {unit}")
    return NormalizedValue(raw, value, ValueKind.NUMBER, method, warnings, provenance)


def _resolve_year(raw_year: int, warnings: list[str]) -> tuple[int, str]:
    if 1700 <= raw_year <= 2200:
        return raw_year, "gregorian"
    if 1178 <= raw_year <= 1599:
        return raw_year, "jalali"
    if 1600 <= raw_year <= 1699:
        warnings.append(f"year {raw_year} is ambiguous; treated as Jalali")
        return raw_year, "jalali"
    if 0 <= raw_year <= 99:
        expanded = 2000 + raw_year if raw_year <= 69 else 1900 + raw_year
        warnings.append(f"two-digit year {raw_year} expanded to {expanded}")
        return expanded, "gregorian"
    warnings.append(f"year {raw_year} is outside the supported range")
    return raw_year, "unknown"


def _build_date(year: int, month: int, day: int, calendar: str) -> date | None:
    if calendar == "jalali":
        if not is_valid_jalali(year, month, day):
            return None
        gy, gm, gd = jalali_to_gregorian(year, month, day)
        try:
            return date(gy, gm, gd)
        except ValueError:
            return None
    try:
        return date(year, month, day)
    except ValueError:
        return None


def parse_date(
    raw: Any,
    *,
    day_first: bool = True,
    provenance: Provenance | None = None,
) -> NormalizedValue:
    """Parse Jalali, Gregorian or Excel-serial dates into a ``date``."""
    warnings: list[str] = []
    if isinstance(raw, datetime):
        return NormalizedValue(raw, raw.date(), ValueKind.DATE, "datetime", provenance=provenance)
    if isinstance(raw, date):
        return NormalizedValue(raw, raw, ValueKind.DATE, "date", provenance=provenance)
    if raw is None or isinstance(raw, bool):
        return NormalizedValue(raw, None, ValueKind.EMPTY, "not-a-date", provenance=provenance)
    if isinstance(raw, (int, float, Decimal)):
        value = float(raw)
        if 1 <= value <= 2_958_465:
            if value < 20_000:
                warnings.append("small numeric value read as Excel serial date; verify")
            try:
                converted = excel_serial_to_date(value)
            except (ValueError, OverflowError):
                return NormalizedValue(raw, None, ValueKind.ERROR, "serial-out-of-range", provenance=provenance)
            return NormalizedValue(raw, converted, ValueKind.DATE, "excel-serial", warnings, provenance)
        return NormalizedValue(raw, None, ValueKind.TEXT, "not-a-date", ["number is not a plausible date serial"], provenance)

    text = normalize_text(normalize_digits(str(raw)))
    if not text:
        return NormalizedValue(raw, None, ValueKind.EMPTY, "empty", provenance=provenance)

    name_match = _DATE_WITH_MONTH_NAME.match(text)
    if name_match:
        day_text, month_name, year_text = name_match.groups()
        month_name = month_name.strip()
        month = None
        if month_name in _JALALI_MONTHS:
            month = _JALALI_MONTHS.index(month_name) + 1
            year, calendar = _resolve_year(int(year_text), warnings)
            calendar = "jalali"
        elif month_name.lower() in _GREGORIAN_MONTHS:
            month = _GREGORIAN_MONTHS.index(month_name.lower()) + 1
            year, calendar = _resolve_year(int(year_text), warnings)
            calendar = "gregorian"
        if month:
            converted = _build_date(year, month, int(day_text), calendar)
            if converted:
                return NormalizedValue(raw, converted, ValueKind.DATE, f"{calendar}-month-name", warnings, provenance)
        return NormalizedValue(raw, None, ValueKind.TEXT, "unparsed-date", warnings + ["month name not recognized"], provenance)

    match = _DATE_SEPARATED.match(text)
    if not match:
        return NormalizedValue(raw, None, ValueKind.TEXT, "not-a-date", ["date pattern not recognized"], provenance)

    first, second, third = (int(group) for group in match.groups())
    if len(match.group(1)) == 4:
        year, month, day = first, second, third
        order = "YMD"
    elif len(match.group(3)) == 4:
        if first > 12:
            day, month, year = first, second, third
            order = "DMY"
        elif second > 12:
            month, day, year = first, second, third
            order = "MDY"
        elif day_first:
            day, month, year = first, second, third
            order = "DMY"
            warnings.append("day/month order ambiguous; read as day/month/year")
        else:
            month, day, year = first, second, third
            order = "MDY"
            warnings.append("day/month order ambiguous; read as month/day/year")
    else:
        return NormalizedValue(raw, None, ValueKind.TEXT, "not-a-date", ["no four-digit year found"], provenance)

    year, calendar = _resolve_year(year, warnings)
    converted = _build_date(year, month, day, calendar)
    if converted is None:
        return NormalizedValue(
            raw, None, ValueKind.TEXT, "invalid-date",
            warnings + [f"{calendar} date {year}-{month:02d}-{day:02d} is not valid"], provenance,
        )
    return NormalizedValue(raw, converted, ValueKind.DATE, f"{calendar}-{order}", warnings, provenance)


def looks_like_number(text: Any) -> bool:
    if isinstance(text, (int, float, Decimal)) and not isinstance(text, bool):
        return True
    if not isinstance(text, str):
        return False
    value = strip_invisible(normalize_digits(text)).strip()
    if not value or value.upper() in EXCEL_ERRORS or not re.search(r"\d", value):
        return False
    letters = re.sub(r"[\d\s.,٬٫'’_%()\[\]+*/\-−–—]", "", value)
    if letters:
        lowered = value.lower()
        for token in CURRENCY_TOKENS:
            lowered = lowered.replace(token, " ")
        for token in SCALE_UNITS:
            lowered = lowered.replace(token.lower(), " ")
        for token in ("حدود", "تقریبا", "تقریباً", "بالغ"):
            lowered = lowered.replace(token, " ")
        letters = re.sub(r"[\d\s.,٬٫'’_%()\[\]+*/\-−–—]", "", lowered)
        return not letters
    return True


def looks_like_date(text: Any) -> bool:
    if isinstance(text, (date, datetime)):
        return True
    if not isinstance(text, str):
        return False
    value = normalize_text(normalize_digits(text))
    if _DATE_SEPARATED.match(value):
        return True
    if _DATE_WITH_MONTH_NAME.match(value):
        parts = value.split()
        month_name = parts[1] if len(parts) > 1 else ""
        return month_name in _JALALI_MONTHS or month_name.lower() in _GREGORIAN_MONTHS
    return False


def parse_value(
    raw: Any,
    *,
    kind_hint: ValueKind | None = None,
    numeric_context: bool = False,
    day_first: bool = True,
    provenance: Provenance | None = None,
) -> NormalizedValue:
    """Standardize one cell value using an optional expected kind."""
    if raw is None:
        return NormalizedValue(raw, None, ValueKind.EMPTY, "empty", provenance=provenance)
    if isinstance(raw, str):
        stripped = strip_invisible(raw).strip()
        if not stripped:
            return NormalizedValue(raw, None, ValueKind.EMPTY, "empty", provenance=provenance)
        if stripped.upper() in EXCEL_ERRORS:
            return NormalizedValue(raw, None, ValueKind.ERROR, "excel-error", [f"Excel error value {stripped}"], provenance)

    if kind_hint is ValueKind.DATE:
        return parse_date(raw, day_first=day_first, provenance=provenance)
    if kind_hint is ValueKind.NUMBER:
        return parse_number(raw, numeric_context=numeric_context, provenance=provenance)
    if kind_hint is ValueKind.TEXT:
        if isinstance(raw, (date, datetime)):
            return NormalizedValue(raw, format_jalali(raw), ValueKind.TEXT, "date-as-text", provenance=provenance)
        return NormalizedValue(
            raw, normalize_text(normalize_digits(str(raw))), ValueKind.TEXT, "text", provenance=provenance
        )

    if isinstance(raw, bool):
        return NormalizedValue(raw, raw, ValueKind.BOOLEAN, "boolean", provenance=provenance)
    if isinstance(raw, (date, datetime)):
        value = raw.date() if isinstance(raw, datetime) else raw
        return NormalizedValue(raw, value, ValueKind.DATE, "date", provenance=provenance)
    if isinstance(raw, (int, float, Decimal)):
        return parse_number(raw, numeric_context=numeric_context, provenance=provenance)
    if looks_like_date(raw):
        parsed = parse_date(raw, day_first=day_first, provenance=provenance)
        if parsed.kind is ValueKind.DATE:
            return parsed
    if looks_like_number(raw):
        parsed = parse_number(raw, numeric_context=numeric_context, provenance=provenance)
        if parsed.kind is ValueKind.NUMBER:
            return parsed
    return NormalizedValue(
        raw, normalize_text(normalize_digits(str(raw))), ValueKind.TEXT, "text", provenance=provenance
    )


def canonical_number(value: Decimal | int | float | str) -> str:
    """Canonical text for a number, used to compare quoted figures exactly."""
    decimal_value = value if isinstance(value, Decimal) else Decimal(str(value))
    if decimal_value == decimal_value.to_integral_value():
        return str(int(decimal_value))
    return format(decimal_value.normalize(), "f")


def extract_numeric_tokens(text: str) -> list[str]:
    """Raw numeric tokens in a text, with Persian/Arabic digits converted."""
    prepared = strip_invisible(normalize_digits(text))
    return _NUMBER_TOKEN.findall(prepared)


def extract_numeric_claims(text: str) -> list[str]:
    """Numeric claims including a trailing scale word.

    ``"حدود ۲ میلیون تومان"`` yields ``"2 میلیون"`` so the comparison layer can
    resolve the real magnitude instead of comparing a bare ``2``.
    """
    prepared = strip_invisible(normalize_digits(text))
    claims: list[str] = []
    for match in _NUMBER_TOKEN.finditer(prepared):
        token = match.group(0).strip()
        suffix = _SCALE_SUFFIX_RE.match(prepared[match.end() :])
        if suffix is not None:
            token = f"{token} {suffix.group(1)}"
        claims.append(token)
    return claims
