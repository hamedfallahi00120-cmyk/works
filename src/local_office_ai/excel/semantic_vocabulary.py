"""Shared vocabulary and value-pattern evidence for semantic classification.

This is a leaf module: it depends only on normalization and the core value
types, so both the detector (``sheet_analyzer``) and the explicit semantic
classifier can use one source of truth without an import cycle.

The central rule is that a numeric-looking value is not automatically a
measure.  An administrative code, a postal code, a personnel number or a year
can be stored as a number and must stay a key, never enter a mean or a sum.
"""

from __future__ import annotations

import re
from typing import Iterable

from local_office_ai.excel.normalization import normalize_digits, normalize_text

# Words that make a column an identifier regardless of its numeric content.
STRONG_IDENTIFIER_WORDS = (
    "شناسه",
    "شناسه‌ها",
    "شناسه‌های",
    "شناسایی",
    "کد",
    "کدها",
    "کدهای",
    "شماره",
    "شماره‌ها",
    "شماره‌های",
    "ردیف",
    "ردیف‌ها",
    "ردیف‌های",
    "سریال",
    "بارکد",
    "پلاک",
    "پرسنلی",
    "پستی",
    "id",
    "code",
    "serial",
    "barcode",
    "postal",
    "zip",
    "uuid",
    "sku",
    "ref",
    "reference",
    "index",
    "row",
    "no",
)

# Words that identify a key only when the content is numeric/code-like.  A text
# column called «منطقه» with province names is a category, not an identifier.
WEAK_IDENTIFIER_WORDS = (
    "منطقه",
    "ناحیه",
    "پرونده",
    "سند",
    "رسید",
    "حساب",
    "مرجع",
    "تلفن",
    "همراه",
    "موبایل",
    "پرسنل",
    "کارمند",
    "کارکنان",
    "خودرو",
    "وسیله",
    "national",
    "personnel",
    "employee",
    "staff",
    "vehicle",
    "region",
    "district",
    "record",
    "plate",
    "key",
    "account",
    "phone",
    "mobile",
    "tel",
)

# Words that mark a column as a measurable quantity.  They suppress the weaker
# "many distinct integers" identifier heuristic so a column of distinct amounts
# is not mistaken for a code.
MEASURE_WORDS = (
    "مبلغ",
    "مقدار",
    "تعداد",
    "جمع",
    "مجموع",
    "مانده",
    "حقوق",
    "دستمزد",
    "هزینه",
    "درآمد",
    "قیمت",
    "نرخ",
    "درصد",
    "سهم",
    "امتیاز",
    "عدد",
    "دفعات",
    "وزن",
    "حجم",
    "مساحت",
    "فاصله",
    "مصرف",
    "تولید",
    "فروش",
    "خرید",
    "موجودی",
    "بدهی",
    "طلب",
    "سود",
    "زیان",
    "پرداخت",
    "دریافت",
    "اعتبار",
    "تسهیلات",
    "وام",
    "amount",
    "value",
    "total",
    "sum",
    "count",
    "quantity",
    "qty",
    "price",
    "cost",
    "salary",
    "wage",
    "balance",
    "rate",
    "percent",
    "ratio",
    "score",
    "weight",
    "volume",
    "area",
    "distance",
    "usage",
    "consumption",
    "production",
    "sales",
    "purchase",
    "stock",
    "debt",
    "credit",
    "profit",
    "loss",
    "payment",
    "income",
    "expense",
    "budget",
)

# A negative value is legitimate in these measures; elsewhere it is reported.
NEGATIVE_ALLOWED_WORDS = (
    "اختلاف",
    "تفاضل",
    "تغییر",
    "مقایسه",
    "مانده",
    "بدهکار",
    "بستانکار",
    "بدهی",
    "طلب",
    "درصد",
    "نرخ",
    "سود",
    "زیان",
    "خالص",
    "تراز",
    "difference",
    "delta",
    "change",
    "variance",
    "balance",
    "debit",
    "credit",
    "profit",
    "loss",
    "net",
    "margin",
)

# Headers that describe text even when they mention a measure word, e.g.
# «توضیحات مبلغ».  They must not turn a descriptive column into a numeric one.
DESCRIPTIVE_HEADER_WORDS = (
    "توضیح",
    "توضیحات",
    "شرح",
    "ملاحظات",
    "یادداشت",
    "آدرس",
    "نام",
    "عنوان",
    "علت",
    "اقدام",
    "پیوست",
    "نوع",
    "وضعیت",
    "دسته",
    "گروه",
    "طبقه",
    "description",
    "comment",
    "note",
    "notes",
    "address",
    "name",
    "title",
    "reason",
    "action",
    "type",
    "status",
    "category",
    "group",
)

# Identifier columns where uniqueness is an explicit requirement.
UNIQUE_IDENTIFIER_WORDS = (
    "شناسه ملی",
    "کد ملی",
    "کد پرسنلی",
    "شماره پرسنلی",
    "شناسه",
    "شناسه‌ها",
    "شناسه‌های",
    "شناسه‌های ملی",
    "کدهای ملی",
    "سریال",
    "بارکد",
    "ردیف",
    "یکتا",
    "شماره پرونده",
    "شماره‌های پرونده",
    "شناسه پرونده",
    "شناسه‌های پرونده",
    "شماره سند",
    "شماره نامه",
    "شماره رکورد",
    "کد رهگیری",
    "unique",
    "national id",
    "personnel id",
    "serial",
    "barcode",
    "uuid",
    "record id",
    "record number",
    "case number",
    "tracking code",
)

SCALE_WORDS = (
    "هزار",
    "میلیون",
    "میلیارد",
    "همت",
    "thousand",
    "million",
    "billion",
)

CURRENCY_WORDS = (
    "ریال",
    "تومان",
    "دلار",
    "یورو",
    "درهم",
    "پوند",
    "usd",
    "eur",
    "irr",
    "toman",
    "rial",
    "dollar",
    "euro",
)

UNIT_WORDS = (
    "ریال",
    "تومان",
    "درصد",
    "نفر",
    "عدد",
    "متر",
    "مترمربع",
    "متر مربع",
    "کیلومتر",
    "کیلوگرم",
    "کیلو",
    "گرم",
    "لیتر",
    "ساعت",
    "دقیقه",
    "روز",
    "ماه",
    "سال",
    "هکتار",
    "دلار",
    "یورو",
    "usd",
    "eur",
    "irr",
    "percent",
    "kg",
    "km",
    "hour",
    "day",
    "month",
    "year",
    "unit",
)

_LEADING_ZERO = re.compile(r"^0\d+$")
_DIGIT_ONLY = re.compile(r"^\d+$")
_YEAR = re.compile(r"^(?:13|14|19|20)\d{2}$")


def matches_word(text: object, keyword: object) -> bool:
    """Whole-word/phrase match after digit and letter normalization.

    ``کد`` matches ``کد پستی`` but not ``کدگذاری``; ``جمع`` does not match
    inside ``جمعیت``.
    """
    if not isinstance(text, str) or not isinstance(keyword, str):
        return False
    prepared = normalize_text(normalize_digits(text)).lower()
    key = normalize_text(normalize_digits(keyword)).lower().strip()
    if not prepared or not key:
        return False
    pattern = r"(?<!\w)" + r"\s+".join(re.escape(part) for part in key.split()) + r"(?!\w)"
    return re.search(pattern, prepared) is not None


def has_word(text: object, words: Iterable[str]) -> bool:
    return any(matches_word(text, word) for word in words)


def has_identifier_word(header: object, *, strong_only: bool = False) -> bool:
    if has_word(header, STRONG_IDENTIFIER_WORDS):
        return True
    return not strong_only and has_word(header, WEAK_IDENTIFIER_WORDS)


def has_measure_word(header: object) -> bool:
    return has_word(header, MEASURE_WORDS)


def has_descriptive_word(header: object) -> bool:
    return has_word(header, DESCRIPTIVE_HEADER_WORDS)


def allows_negative(header: object) -> bool:
    return has_word(header, NEGATIVE_ALLOWED_WORDS)


def requires_unique_identifier(header: object) -> bool:
    return has_word(header, UNIQUE_IDENTIFIER_WORDS)


def unit_from_header(header: object) -> str:
    if not isinstance(header, str):
        return ""
    prepared = normalize_text(normalize_digits(header)).lower()
    for word in UNIT_WORDS:
        if matches_word(prepared, word):
            return word
    return ""


def scale_words_in(text: object) -> set[str]:
    if not isinstance(text, str):
        return set()
    prepared = normalize_text(normalize_digits(text)).lower()
    return {word for word in SCALE_WORDS if matches_word(prepared, word)}


def currency_words_in(text: object) -> set[str]:
    if not isinstance(text, str):
        return set()
    prepared = normalize_text(normalize_digits(text)).lower()
    return {word for word in CURRENCY_WORDS if matches_word(prepared, word)}


def is_sequential_integers(values: list[int]) -> bool:
    """True for a set of unique integers that forms a contiguous run (1..n)."""
    if len(values) < 3:
        return False
    unique = sorted(set(values))
    if len(unique) != len(values):
        return False
    return unique == list(range(unique[0], unique[0] + len(unique)))


def identifier_value_evidence(
    numbers: Iterable[float] | None,
    samples: Iterable[str] | None,
    *,
    header_has_measure: bool = False,
) -> list[str]:
    """Evidence that the values themselves look like codes, not measurements."""
    evidence: list[str] = []
    values = list(numbers or [])
    texts = [
        normalize_digits(str(sample)).strip()
        for sample in (samples or [])
        if sample is not None and str(sample).strip()
    ]

    digit_texts = [text for text in texts if _DIGIT_ONLY.fullmatch(text)]
    if any(_LEADING_ZERO.fullmatch(text) for text in digit_texts):
        evidence.append("مقادیر با صفر پیشرو؛ متن کد است نه عدد اندازه‌گیری")
    # An all-digit sample is only code evidence when the column does not already
    # parse as numbers.  Otherwise the samples are just the measure's own values.
    if (
        not header_has_measure
        and not values
        and digit_texts
        and len(digit_texts) == len(texts)
        and len(texts) >= 2
    ):
        lengths = {len(text) for text in digit_texts}
        unique_ratio = len(set(digit_texts)) / len(digit_texts)
        if min(lengths) >= 4 and (len(lengths) == 1 or unique_ratio >= 0.9):
            evidence.append("متن تمام‌رقمی با طول کد؛ شناسه است")

    integers = [int(value) for value in values if float(value).is_integer()]
    if (
        not header_has_measure
        and values
        and len(integers) == len(values)
        and len(integers) >= 3
    ):
        widths = {len(str(abs(value))) for value in integers}
        if len(widths) == 1 and next(iter(widths)) >= 5:
            evidence.append("کد عددی با طول ثابت")
    return evidence


def identifier_evidence(
    header: object,
    numbers: Iterable[float] | None,
    samples: Iterable[str] | None,
) -> list[str]:
    """Evidence that a column is an identifier and must not be a measure."""
    values = list(numbers or [])
    texts = [str(sample) for sample in (samples or []) if sample is not None and str(sample).strip()]
    if not values and not texts:
        return []
    evidence: list[str] = []
    strong = has_word(header, STRONG_IDENTIFIER_WORDS)
    weak = has_word(header, WEAK_IDENTIFIER_WORDS)
    measure = has_measure_word(header)
    numeric_texts = any(_DIGIT_ONLY.fullmatch(normalize_digits(text).strip()) for text in texts)
    if strong:
        evidence.append(f"عنوان «{header}» واژه شناسه دارد")
    elif weak and (values or numeric_texts):
        evidence.append(f"عنوان «{header}» واژه کلید دارد و مقادیر آن عددی/کدی است")
    evidence.extend(
        identifier_value_evidence(values, texts, header_has_measure=measure)
    )
    return evidence


def looks_like_year_value(value: object) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return float(value).is_integer() and bool(_YEAR.fullmatch(str(int(value))))
    if isinstance(value, str):
        return bool(_YEAR.fullmatch(normalize_digits(value).strip()))
    return False
