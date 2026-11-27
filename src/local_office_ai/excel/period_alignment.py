"""Normalize reporting periods and decide whether two periods are comparable.

Administrative workbooks describe the same reporting period in many ways:
``تیر 1403``, ``1403/04``, ``1403-04-01``, ``July 2024`` or an explicit
``از 1403/01/01 تا 1403/03/31`` range.  Two values may only be treated as the
same figure when their reporting periods are compatible, so this module turns
every reference into one :class:`NormalizedPeriod` with real start/end dates and
then explains the relationship between two periods.

Rules kept from the rest of the engine:

* Jalali and Gregorian references are both supported; a Jalali date is converted
  to a real Gregorian date with the shared converter, never by a rough offset.
* A month or year is a range, not a point: ``تیر 1403`` is 1403-04-01..1403-04-31.
* A missing period stays ``not_detected``; it is never guessed from an unrelated
  number in the title.
* Comparing a Jalali month with a Gregorian month yields ``overlapping`` (the
  ranges are not identical), never ``equivalent``.  Only identical ranges are
  equivalent.
"""

from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta
from typing import Any, Iterable

from local_office_ai.excel.normalization import (
    gregorian_to_jalali,
    is_valid_jalali,
    jalali_to_gregorian,
    normalize_digits,
    normalize_text,
)

CALENDAR_JALALI = "jalali"
CALENDAR_GREGORIAN = "gregorian"
CALENDAR_MIXED = "mixed"
CALENDAR_UNKNOWN = "unknown"

GRANULARITY_DAY = "day"
GRANULARITY_MONTH = "month"
GRANULARITY_YEAR = "year"
GRANULARITY_RANGE = "range"
GRANULARITY_UNKNOWN = "unknown"

STATUS_DETECTED = "detected"
STATUS_NOT_DETECTED = "not_detected"

RELATION_EQUIVALENT = "equivalent"
RELATION_OVERLAPPING = "overlapping"
RELATION_CONTAINS = "contains"
RELATION_WITHIN = "within"
RELATION_INCOMPATIBLE = "incompatible"
RELATION_UNKNOWN = "unknown"

RELATION_FA = {
    RELATION_EQUIVALENT: "یکسان",
    RELATION_OVERLAPPING: "هم‌پوشان",
    RELATION_CONTAINS: "شامل",
    RELATION_WITHIN: "داخل بازه دیگر",
    RELATION_INCOMPATIBLE: "ناسازگار",
    RELATION_UNKNOWN: "نامشخص",
}

CALENDAR_FA = {
    CALENDAR_JALALI: "شمسی",
    CALENDAR_GREGORIAN: "میلادی",
    CALENDAR_MIXED: "ترکیبی",
    CALENDAR_UNKNOWN: "نامشخص",
}

JALALI_MONTHS = (
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
)

GREGORIAN_MONTHS = (
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
)

GREGORIAN_MONTH_ALIASES = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}

# Persian seasons and half-year/quarter phrases.  Each entry is an inclusive
# month range inside the same year.
_SEASONS = {
    "بهار": (1, 3),
    "تابستان": (4, 6),
    "پاییز": (7, 9),
    "زمستان": (10, 12),
    "spring": (1, 3),
    "summer": (4, 6),
    "autumn": (7, 9),
    "fall": (7, 9),
    "winter": (10, 12),
}

_QUARTER_ORDINALS = {
    "اول": (1, 3),
    "دوم": (4, 6),
    "سوم": (7, 9),
    "چهارم": (10, 12),
    "۱": (1, 3),
    "۲": (4, 6),
    "۳": (7, 9),
    "۴": (10, 12),
    "1": (1, 3),
    "2": (4, 6),
    "3": (7, 9),
    "4": (10, 12),
}

_HALF_ORDINALS = {
    "اول": (1, 6),
    "دوم": (7, 12),
    "۱": (1, 6),
    "۲": (7, 12),
    "1": (1, 6),
    "2": (7, 12),
}

_PERSIAN_LETTERS = "آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءأإؤئ"

_FOUR_DIGIT_YEAR = re.compile(r"(?<!\d)(1[1-6]\d{2}|1[7-9]\d{2}|20\d{2}|21\d{2})(?!\d)")
_NUMERIC_DATE = re.compile(
    r"^\s*(\d{1,4})\s*[/\-.\\]\s*(\d{1,2})\s*[/\-.\\]\s*(\d{1,4})\s*$"
)
_NUMERIC_MONTH = re.compile(r"^\s*(\d{1,4})\s*[/\-.\\]\s*(\d{1,4})\s*$")
_EMBEDDED_DATE = re.compile(
    r"(?<!\d)(\d{4})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})(?!\d)"
)
_EMBEDDED_MONTH_YMD = re.compile(
    r"(?<!\d)(\d{4})\s*[/\-.]\s*(\d{1,2})(?![\d/.\-])"
)
_EMBEDDED_MONTH_DMY = re.compile(
    r"(?<![\d/.\-])(\d{1,2})\s*[/\-.]\s*(\d{4})(?!\d)"
)


@dataclass(frozen=True)
class NormalizedPeriod:
    """One reporting period with a real date range and its source evidence."""

    label: str
    calendar: str
    granularity: str
    start: date | None
    end: date | None
    year: int | None = None
    month: int | None = None
    day: int | None = None
    status: str = STATUS_DETECTED
    source: str = ""
    evidence: str = ""

    @property
    def detected(self) -> bool:
        return self.status == STATUS_DETECTED and self.start is not None and self.end is not None

    @property
    def calendar_label(self) -> str:
        return CALENDAR_FA.get(self.calendar, self.calendar)

    @property
    def key(self) -> str:
        if not self.detected:
            return "not_detected"
        return f"{self.start.isoformat()}..{self.end.isoformat()}"

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "calendar": self.calendar,
            "calendar_label": self.calendar_label,
            "granularity": self.granularity,
            "start": self.start.isoformat() if self.start else None,
            "end": self.end.isoformat() if self.end else None,
            "year": self.year,
            "month": self.month,
            "day": self.day,
            "status": self.status,
            "source": self.source,
            "evidence": self.evidence,
        }


@dataclass
class PeriodCompatibility:
    """The explained relationship between two periods."""

    relation: str
    left: NormalizedPeriod | None
    right: NormalizedPeriod | None
    explanation: str = ""
    overlap_start: date | None = None
    overlap_end: date | None = None

    @property
    def compatible(self) -> bool:
        return self.relation in {
            RELATION_EQUIVALENT,
            RELATION_OVERLAPPING,
            RELATION_CONTAINS,
            RELATION_WITHIN,
        }

    @property
    def equivalent(self) -> bool:
        return self.relation == RELATION_EQUIVALENT

    @property
    def relation_label(self) -> str:
        return RELATION_FA.get(self.relation, self.relation)

    def to_dict(self) -> dict[str, Any]:
        return {
            "relation": self.relation,
            "relation_label": self.relation_label,
            "compatible": self.compatible,
            "equivalent": self.equivalent,
            "explanation": self.explanation,
            "overlap_start": self.overlap_start.isoformat() if self.overlap_start else None,
            "overlap_end": self.overlap_end.isoformat() if self.overlap_end else None,
            "left": self.left.to_dict() if self.left else None,
            "right": self.right.to_dict() if self.right else None,
        }


# --------------------------------------------------------------------------- #
# Calendar ranges
# --------------------------------------------------------------------------- #
def _jalali_month_bounds(year: int, month: int) -> tuple[date, date] | None:
    if not (1178 <= year <= 1633 and 1 <= month <= 12):
        return None
    try:
        start = date(*jalali_to_gregorian(year, month, 1))
        if month == 12:
            next_start = date(*jalali_to_gregorian(year + 1, 1, 1))
        else:
            next_start = date(*jalali_to_gregorian(year, month + 1, 1))
    except (ValueError, OverflowError):
        return None
    return start, next_start - timedelta(days=1)


def _gregorian_month_bounds(year: int, month: int) -> tuple[date, date] | None:
    if not (1 <= month <= 12):
        return None
    try:
        last = calendar.monthrange(year, month)[1]
        return date(year, month, 1), date(year, month, last)
    except ValueError:
        return None


def _jalali_year_bounds(year: int) -> tuple[date, date] | None:
    start_bounds = _jalali_month_bounds(year, 1)
    if start_bounds is None:
        return None
    try:
        next_start = date(*jalali_to_gregorian(year + 1, 1, 1))
    except (ValueError, OverflowError):
        return None
    return start_bounds[0], next_start - timedelta(days=1)


def _gregorian_year_bounds(year: int) -> tuple[date, date] | None:
    try:
        return date(year, 1, 1), date(year, 12, 31)
    except ValueError:
        return None


def _resolve_year(value: int) -> str:
    if 1700 <= value <= 2200:
        return CALENDAR_GREGORIAN
    if 1178 <= value <= 1633:
        # 1600..1633 is the upper edge of the shared Jalali converter; both
        # stages use the same valid range so a period is never invented.
        return CALENDAR_JALALI
    return CALENDAR_UNKNOWN


def _month_bounds(year: int, month: int, calendar_name: str) -> tuple[date, date] | None:
    if calendar_name == CALENDAR_JALALI:
        return _jalali_month_bounds(year, month)
    return _gregorian_month_bounds(year, month)


def _year_bounds(year: int, calendar_name: str) -> tuple[date, date] | None:
    if calendar_name == CALENDAR_JALALI:
        return _jalali_year_bounds(year)
    return _gregorian_year_bounds(year)


def _months_in(text: str) -> list[tuple[str, int]]:
    """Jalali month names that appear as whole words, with their month number."""
    found: list[tuple[str, int]] = []
    for index, month in enumerate(JALALI_MONTHS, start=1):
        pattern = rf"(?<![{_PERSIAN_LETTERS}]){month}(?:ماه)?(?![{_PERSIAN_LETTERS}])"
        if re.search(pattern, text):
            found.append((month, index))
    return found


def _gregorian_months_in(text: str) -> list[tuple[str, int]]:
    found: list[tuple[str, int]] = []
    lowered = text.lower()
    for index, month in enumerate(GREGORIAN_MONTHS, start=1):
        if re.search(rf"(?<![a-z]){month}(?![a-z])", lowered):
            found.append((month, index))
    for alias, index in GREGORIAN_MONTH_ALIASES.items():
        if re.search(rf"(?<![a-z]){alias}\.?(?![a-z])", lowered):
            name = GREGORIAN_MONTHS[index - 1]
            if all(existing != name for existing, _ in found):
                found.append((name, index))
    return found


def _period(
    *,
    label: str,
    calendar_name: str,
    granularity: str,
    bounds: tuple[date, date] | None,
    year: int | None = None,
    month: int | None = None,
    day: int | None = None,
    source: str = "",
    evidence: str = "",
) -> NormalizedPeriod:
    if bounds is None:
        return NormalizedPeriod(
            label=label,
            calendar=calendar_name,
            granularity=GRANULARITY_UNKNOWN,
            start=None,
            end=None,
            year=year,
            month=month,
            day=day,
            status=STATUS_NOT_DETECTED,
            source=source,
            evidence=evidence or "بازه تاریخ معتبر نبود",
        )
    start, end = bounds
    return NormalizedPeriod(
        label=label,
        calendar=calendar_name,
        granularity=granularity,
        start=start,
        end=end,
        year=year,
        month=month,
        day=day,
        status=STATUS_DETECTED,
        source=source,
        evidence=evidence,
    )


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #
def _clean(text: Any) -> str:
    return normalize_text(normalize_digits(str(text or ""))).strip()


def _parse_ordinal_range(text: str, source: str) -> NormalizedPeriod | None:
    """``سه‌ماهه اول 1403``, ``نیمه دوم 1403``, ``فصل 2 1403`` and seasons."""
    year_match = _FOUR_DIGIT_YEAR.search(text)
    if year_match is None:
        return None
    year = int(year_match.group(1))
    calendar_name = _resolve_year(year)
    if calendar_name == CALENDAR_UNKNOWN:
        return None

    lowered = text.lower()
    for season, (first, last) in _SEASONS.items():
        if re.search(rf"(?<![{_PERSIAN_LETTERS}a-z]){season}(?![{_PERSIAN_LETTERS}a-z])", lowered):
            start = _month_bounds(year, first, calendar_name)
            end = _month_bounds(year, last, calendar_name)
            if start and end:
                return _period(
                    label=f"{season} {year}",
                    calendar_name=calendar_name,
                    granularity=GRANULARITY_RANGE,
                    bounds=(start[0], end[1]),
                    year=year,
                    source=source,
                    evidence=f"فصل/بازه فصلی «{season}» با سال {year}",
                )

    for phrase, mapping in (
        (r"سه\s*ماهه|سه ماهه", _QUARTER_ORDINALS),
        (r"شش\s*ماهه|شش ماهه|نیمه|نیم\s*سال", _HALF_ORDINALS),
        (r"فصل", _QUARTER_ORDINALS),
        (r"quarter", _QUARTER_ORDINALS),
    ):
        if not re.search(phrase, text):
            continue
        for ordinal, (first, last) in mapping.items():
            if re.search(
                rf"(?<![\w{_PERSIAN_LETTERS}]){ordinal}(?![\w{_PERSIAN_LETTERS}])",
                text,
            ):
                start = _month_bounds(year, first, calendar_name)
                end = _month_bounds(year, last, calendar_name)
                if start and end:
                    return _period(
                        label=f"{first}-{last} {year}",
                        calendar_name=calendar_name,
                        granularity=GRANULARITY_RANGE,
                        bounds=(start[0], end[1]),
                        year=year,
                        source=source,
                        evidence=f"دوره چندماهه ({first} تا {last}) با سال {year}",
                    )

    quarter = re.fullmatch(
        r"\s*(?:q|quarter)\s*([1-4])\s*[-/.]?\s*(\d{4})\s*", text, flags=re.IGNORECASE
    )
    if quarter:
        number, year_text = int(quarter.group(1)), int(quarter.group(2))
        calendar_name = _resolve_year(year_text)
        first = (number - 1) * 3 + 1
        last = first + 2
        start = _month_bounds(year_text, first, calendar_name)
        end = _month_bounds(year_text, last, calendar_name)
        if start and end:
            return _period(
                label=f"Q{number} {year_text}",
                calendar_name=calendar_name,
                granularity=GRANULARITY_RANGE,
                bounds=(start[0], end[1]),
                year=year_text,
                source=source,
                evidence=f"سه‌ماهه Q{number} سال {year_text}",
            )
    return None


def _parse_named_month(text: str, source: str) -> NormalizedPeriod | None:
    year_match = _FOUR_DIGIT_YEAR.search(text)
    if year_match is None:
        return None
    year = int(year_match.group(1))
    calendar_name = _resolve_year(year)
    if calendar_name == CALENDAR_UNKNOWN:
        return None

    months = _months_in(text)
    if months and calendar_name == CALENDAR_JALALI:
        name, month = months[0]
        bounds = _jalali_month_bounds(year, month)
        if bounds:
            return _period(
                label=f"{name} {year}",
                calendar_name=CALENDAR_JALALI,
                granularity=GRANULARITY_MONTH,
                bounds=bounds,
                year=year,
                month=month,
                source=source,
                evidence=f"نام ماه شمسی «{name}» با سال {year}",
            )

    gregorian = _gregorian_months_in(text)
    if gregorian:
        name, month = gregorian[0]
        bounds = _gregorian_month_bounds(year, month)
        if bounds:
            return _period(
                label=f"{name.title()} {year}",
                calendar_name=CALENDAR_GREGORIAN,
                granularity=GRANULARITY_MONTH,
                bounds=bounds,
                year=year,
                month=month,
                source=source,
                evidence=f"نام ماه میلادی «{name}» با سال {year}",
            )
    return None


def _parse_numeric(text: str, source: str) -> NormalizedPeriod | None:
    match = _NUMERIC_DATE.match(text)
    if match:
        first_text, second_text, third_text = match.groups()
        first, second, third = int(first_text), int(second_text), int(third_text)
        if len(first_text) == 4:
            year, month, day = first, second, third
        elif len(third_text) == 4:
            if first > 12 and second <= 12:
                day, month, year = first, second, third
            elif second > 12 and first <= 12:
                month, day, year = first, second, third
            else:
                day, month, year = first, second, third
        else:
            return None
        calendar_name = _resolve_year(year)
        if calendar_name == CALENDAR_UNKNOWN:
            return None
        if calendar_name == CALENDAR_JALALI:
            if not is_valid_jalali(year, month, day):
                return None
            try:
                gy, gm, gd = jalali_to_gregorian(year, month, day)
                start = date(gy, gm, gd)
            except (ValueError, OverflowError):
                return None
        else:
            try:
                start = date(year, month, day)
            except ValueError:
                return None
        return _period(
            label=text.strip(),
            calendar_name=calendar_name,
            granularity=GRANULARITY_DAY,
            bounds=(start, start),
            year=year,
            month=month,
            day=day,
            source=source,
            evidence=f"تاریخ کامل {CALENDAR_FA.get(calendar_name, calendar_name)}",
        )

    match = _NUMERIC_MONTH.match(text)
    if match:
        first_text, second_text = match.groups()
        first, second = int(first_text), int(second_text)
        if len(first_text) == 4:
            year, month = first, second
        elif len(second_text) == 4:
            month, year = first, second
        else:
            return None
        calendar_name = _resolve_year(year)
        if calendar_name == CALENDAR_UNKNOWN:
            return None
        bounds = _month_bounds(year, month, calendar_name)
        if bounds:
            return _period(
                label=text.strip(),
                calendar_name=calendar_name,
                granularity=GRANULARITY_MONTH,
                bounds=bounds,
                year=year,
                month=month,
                source=source,
                evidence=f"سال و ماه {CALENDAR_FA.get(calendar_name, calendar_name)}",
            )
        return None

    year_match = _FOUR_DIGIT_YEAR.fullmatch(text.strip())
    if year_match:
        year = int(year_match.group(1))
        calendar_name = _resolve_year(year)
        bounds = _year_bounds(year, calendar_name) if calendar_name != CALENDAR_UNKNOWN else None
        if bounds:
            return _period(
                label=text.strip(),
                calendar_name=calendar_name,
                granularity=GRANULARITY_YEAR,
                bounds=bounds,
                year=year,
                source=source,
                evidence=f"سال {CALENDAR_FA.get(calendar_name, calendar_name)}",
            )
    return None


def _parse_range(text: str, source: str) -> NormalizedPeriod | None:
    """Explicit ``از ... تا ...`` or ``... .. ...`` ranges."""
    parts: list[str] = []
    if " تا " in text:
        left, _, right = text.partition(" تا ")
        parts = [left, right]
    elif ".." in text:
        left, _, right = text.partition("..")
        parts = [left, right]
    elif re.search(r"\s+-\s+", text):
        left, right = re.split(r"\s+-\s+", text, maxsplit=1)
        parts = [left, right]
    elif re.search(r"\s+to\s+", text, flags=re.IGNORECASE):
        left, right = re.split(r"\s+to\s+", text, maxsplit=1, flags=re.IGNORECASE)
        parts = [left, right]
    if len(parts) != 2:
        # A compact range such as ``1403/01/01-1403/03/31`` or
        # ``2024-01-01-2024-03-31`` has no spaces around the separator.
        compact = re.search(
            r"(?<!\d)(\d{4}[/\-.]\d{1,2}(?:[/\-.]\d{1,2})?)"
            r"\s*-\s*"
            r"(\d{4}[/\-.]\d{1,2}(?:[/\-.]\d{1,2})?)(?!\d)",
            text,
        )
        if compact:
            parts = [compact.group(1), compact.group(2)]
    if len(parts) != 2:
        year_range = re.search(r"(?<!\d)(\d{4})\s*-\s*(\d{4})(?!\d)", text)
        if year_range:
            parts = [year_range.group(1), year_range.group(2)]
    if len(parts) != 2:
        return None
    left = re.sub(r"^\s*(از|from)\s+", "", parts[0].strip(), flags=re.IGNORECASE)
    right = re.sub(r"^\s*(تا|to)\s+", "", parts[1].strip(), flags=re.IGNORECASE)
    # ``فروردین تا خرداد 1403`` carries the year only once, on the right side.
    if _FOUR_DIGIT_YEAR.search(right) and not _FOUR_DIGIT_YEAR.search(left):
        year_match = _FOUR_DIGIT_YEAR.search(right)
        left = f"{left} {year_match.group(1)}"
    elif _FOUR_DIGIT_YEAR.search(left) and not _FOUR_DIGIT_YEAR.search(right):
        year_match = _FOUR_DIGIT_YEAR.search(left)
        right = f"{right} {year_match.group(1)}"
    start_period = parse_period(left, source=source)
    end_period = parse_period(right, source=source)
    if start_period is None or end_period is None:
        return None
    if not start_period.detected or not end_period.detected:
        return None
    if start_period.start is None or end_period.end is None:
        return None
    if start_period.start > end_period.end:
        start_period, end_period = end_period, start_period
    calendars = {start_period.calendar, end_period.calendar}
    calendar_name = calendars.pop() if len(calendars) == 1 else CALENDAR_MIXED
    return _period(
        label=text.strip(),
        calendar_name=calendar_name,
        granularity=GRANULARITY_RANGE,
        bounds=(start_period.start, end_period.end),
        year=start_period.year,
        source=source,
        evidence=f"بازه صریح از {start_period.label} تا {end_period.label}",
    )


def _parse_embedded_numeric(text: str, source: str) -> NormalizedPeriod | None:
    """A numeric date or year-month inside a header such as ``تعداد 1403/04``."""
    match = _EMBEDDED_DATE.search(text)
    if match:
        period = _parse_numeric(
            f"{match.group(1)}/{match.group(2)}/{match.group(3)}", source
        )
        if period:
            return replace(
                period,
                label=text.strip(),
                evidence=f"تاریخ کامل داخل برچسب: {match.group(0)}",
            )
    match = _EMBEDDED_MONTH_YMD.search(text)
    if match:
        period = _parse_numeric(f"{match.group(1)}/{match.group(2)}", source)
        if period:
            return replace(
                period,
                label=text.strip(),
                evidence=f"سال و ماه داخل برچسب: {match.group(0)}",
            )
    match = _EMBEDDED_MONTH_DMY.search(text)
    if match:
        period = _parse_numeric(f"{match.group(2)}/{match.group(1)}", source)
        if period:
            return replace(
                period,
                label=text.strip(),
                evidence=f"ماه و سال داخل برچسب: {match.group(0)}",
            )
    return None


def _parse_embedded_year(text: str, source: str) -> NormalizedPeriod | None:
    """A single year inside a short title such as ``گزارش روزانه 1403``.

    A long free-text cell is never read as a period; only a short label with one
    unambiguous four-digit year qualifies, matching the period detector's title
    behaviour.
    """
    if len(text) > 120:
        return None
    years = {int(match.group(1)) for match in _FOUR_DIGIT_YEAR.finditer(text)}
    if len(years) != 1:
        return None
    year = next(iter(years))
    if not (1300 <= year <= 1500 or 1900 <= year <= 2100):
        # A bare four-digit number outside the administrative year bands is a
        # code, not a reporting year.
        return None
    calendar_name = _resolve_year(year)
    bounds = _year_bounds(year, calendar_name) if calendar_name != CALENDAR_UNKNOWN else None
    if bounds is None:
        return None
    return _period(
        label=text.strip(),
        calendar_name=calendar_name,
        granularity=GRANULARITY_YEAR,
        bounds=bounds,
        year=year,
        source=source,
        evidence=f"تنها سال موجود در متن: {year}",
    )


def parse_period(text: Any, *, source: str = "") -> NormalizedPeriod | None:
    """Parse one period reference; ``None`` when nothing supports a period."""

    if isinstance(text, datetime):
        value = text.date()
        return _period(
            label=value.isoformat(),
            calendar_name=CALENDAR_GREGORIAN,
            granularity=GRANULARITY_DAY,
            bounds=(value, value),
            year=value.year,
            month=value.month,
            day=value.day,
            source=source,
            evidence="مقدار تاریخ سلول (میلادی)",
        )
    if isinstance(text, date):
        return _period(
            label=text.isoformat(),
            calendar_name=CALENDAR_GREGORIAN,
            granularity=GRANULARITY_DAY,
            bounds=(text, text),
            year=text.year,
            month=text.month,
            day=text.day,
            source=source,
            evidence="مقدار تاریخ سلول (میلادی)",
        )
    if text is None or isinstance(text, bool):
        return None
    if isinstance(text, (int, float)):
        if float(text).is_integer():
            text = str(int(text))
        else:
            return None

    cleaned = _clean(text)
    if not cleaned:
        return None

    for parser in (
        _parse_range,
        _parse_ordinal_range,
        _parse_named_month,
        _parse_numeric,
        _parse_embedded_numeric,
        _parse_embedded_year,
    ):
        period = parser(cleaned, source)
        if period is not None:
            return period
    return None


def parse_periods(
    texts: Iterable[Any],
    *,
    source: str = "",
) -> list[NormalizedPeriod]:
    """Parse every reference and drop duplicates that describe the same range."""
    result: list[NormalizedPeriod] = []
    seen: set[str] = set()
    for text in texts:
        period = parse_period(text, source=source)
        if period is None or not period.detected:
            continue
        key = period.key
        if key in seen:
            continue
        seen.add(key)
        result.append(period)
    return result


# --------------------------------------------------------------------------- #
# Comparison
# --------------------------------------------------------------------------- #
def compare_periods(
    left: NormalizedPeriod | None,
    right: NormalizedPeriod | None,
) -> PeriodCompatibility:
    """Explain whether two periods are equivalent, compatible or incompatible."""

    if left is None or right is None or not left.detected or not right.detected:
        missing = []
        if left is None or not left.detected:
            missing.append("مبنا")
        if right is None or not right.detected:
            missing.append("مقصد")
        return PeriodCompatibility(
            relation=RELATION_UNKNOWN,
            left=left,
            right=right,
            explanation="دوره گزارش‌دهی برای " + " و ".join(missing) + " مشخص نیست؛ برابری اثبات‌شدنی نیست.",
        )

    assert left.start and left.end and right.start and right.end
    if left.start == right.start and left.end == right.end:
        return PeriodCompatibility(
            relation=RELATION_EQUIVALENT,
            left=left,
            right=right,
            overlap_start=left.start,
            overlap_end=left.end,
            explanation=(
                f"دوره «{left.label}» ({left.calendar_label}) و «{right.label}» "
                f"({right.calendar_label}) هر دو بازه {left.start.isoformat()} تا "
                f"{left.end.isoformat()} را پوشش می‌دهند."
            ),
        )

    overlap_start = max(left.start, right.start)
    overlap_end = min(left.end, right.end)
    if overlap_start > overlap_end:
        return PeriodCompatibility(
            relation=RELATION_INCOMPATIBLE,
            left=left,
            right=right,
            explanation=(
                f"دوره «{left.label}» ({left.start.isoformat()} تا {left.end.isoformat()}) با "
                f"دوره «{right.label}» ({right.start.isoformat()} تا {right.end.isoformat()}) "
                "هیچ هم‌پوشانی ندارد؛ مقایسه مجاز نیست."
            ),
        )

    if left.start <= right.start and left.end >= right.end:
        relation = RELATION_CONTAINS
        detail = f"دوره «{left.label}» کل دوره «{right.label}» را در بر می‌گیرد."
    elif right.start <= left.start and right.end >= left.end:
        relation = RELATION_WITHIN
        detail = f"دوره «{left.label}» داخل دوره «{right.label}» قرار دارد."
    else:
        relation = RELATION_OVERLAPPING
        detail = (
            f"دوره‌ها هم‌پوشانی جزئی دارند ({overlap_start.isoformat()} تا "
            f"{overlap_end.isoformat()}); بازه‌ها یکسان نیستند."
        )
    return PeriodCompatibility(
        relation=relation,
        left=left,
        right=right,
        overlap_start=overlap_start,
        overlap_end=overlap_end,
        explanation=detail + " برای برابری کامل باید بازه‌ها یکسان باشند.",
    )


def _merge_intervals(periods: list[NormalizedPeriod]) -> list[tuple[date, date]]:
    intervals = sorted(
        (period.start, period.end)
        for period in periods
        if period.start is not None and period.end is not None
    )
    merged: list[tuple[date, date]] = []
    for start, end in intervals:
        if merged and start <= merged[-1][1] + timedelta(days=1):
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def compare_period_sets(
    left: Iterable[NormalizedPeriod],
    right: Iterable[NormalizedPeriod],
) -> PeriodCompatibility:
    """Compare two period sets without claiming equivalence for partial coverage."""

    left_list = [item for item in left if item.detected]
    right_list = [item for item in right if item.detected]
    if not left_list or not right_list:
        return compare_periods(
            left_list[0] if left_list else None,
            right_list[0] if right_list else None,
        )
    if len(left_list) == 1 and len(right_list) == 1:
        return compare_periods(left_list[0], right_list[0])

    left_intervals = _merge_intervals(left_list)
    right_intervals = _merge_intervals(right_list)
    left_label = period_label(left_list)
    right_label = period_label(right_list)
    if left_intervals == right_intervals:
        return PeriodCompatibility(
            relation=RELATION_EQUIVALENT,
            left=left_list[0],
            right=right_list[0],
            overlap_start=left_intervals[0][0],
            overlap_end=left_intervals[-1][1],
            explanation=(
                f"مجموعه دوره‌های «{left_label}» و «{right_label}» بازه‌های "
                "گزارش‌دهی یکسانی را پوشش می‌دهند."
            ),
        )
    overlap = any(
        max(left_start, right_start) <= min(left_end, right_end)
        for left_start, left_end in left_intervals
        for right_start, right_end in right_intervals
    )
    if not overlap:
        return PeriodCompatibility(
            relation=RELATION_INCOMPATIBLE,
            left=left_list[0],
            right=right_list[0],
            explanation=(
                f"هیچ‌یک از دوره‌های «{left_label}» با دوره‌های «{right_label}» "
                "هم‌پوشانی ندارد؛ مقایسه مجاز نیست."
            ),
        )
    return PeriodCompatibility(
        relation=RELATION_OVERLAPPING,
        left=left_list[0],
        right=right_list[0],
        explanation=(
            f"مجموعه دوره‌های «{left_label}» و «{right_label}» فقط بخشی از بازه‌ها را "
            "مشترک دارند؛ برابری کامل اثبات نشد."
        ),
    )


def period_label(periods: Iterable[NormalizedPeriod]) -> str:
    labels = [item.label for item in periods if item.label]
    return "، ".join(dict.fromkeys(labels))


def jalali_label(value: date) -> str:
    """Render a Gregorian date as a Jalali ``YYYY/MM/DD`` label."""
    jy, jm, jd = gregorian_to_jalali(value.year, value.month, value.day)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"
