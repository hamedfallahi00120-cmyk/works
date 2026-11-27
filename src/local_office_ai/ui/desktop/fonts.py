"""Persian-aware font discovery and application.

Qt already knows every font installed on the machine.  A family is recommended
only when its cmap maps every Persian-specific letter to a real glyph.  No font
file is downloaded or bundled: the user may load a file from their own disk,
and otherwise the best installed family wins.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QFont, QFontDatabase, QRawFont
from PySide6.QtWidgets import QApplication

PERSIAN_PROBE = "پچگژکیۀآئءاب"
PERSIAN_SAMPLE = (
    "گزارش مدیریتی داده‌های اداری؛ بررسی شاخص‌ها، مقایسه دوره‌ها و کیفیت اطلاعات. "
    "۱۲۳۴۵۶۷۸۹۰"
)
FALLBACK_FAMILIES = (
    "Vazirmatn",
    "Vazirmatn UI",
    "Vazir",
    "IRANSansX",
    "IRANSans",
    "IRANYekan",
    "Sahel",
    "Shabnam",
    "Yekan Bakh",
    "B Nazanin",
    "B Yekan",
    "Noto Naskh Arabic",
    "Noto Sans Arabic",
    "Segoe UI",
    "Tahoma",
    "Arial",
)
MIN_SIZE = 9
MAX_SIZE = 22
MIN_HEADING = 13
MAX_HEADING = 32

_capable_cache: list[str] | None = None
_custom_families: dict[str, str] = {}


def _supports_persian(family: str) -> bool:
    if family in {"Sans Serif", "Serif", "Monospace", "Cursive", "Fantasy"}:
        return False
    try:
        if QFontDatabase.WritingSystem.Arabic not in QFontDatabase.writingSystems(family):
            return False
        font = QFont(family)
        font.setPointSize(14)
        raw = QRawFont.fromFont(font)
        if not raw.isValid():
            return False
        indexes = list(raw.glyphIndexesForString(PERSIAN_PROBE))
        return len(indexes) == len(PERSIAN_PROBE) and all(index != 0 for index in indexes)
    except Exception:  # noqa: BLE001 - a broken font entry must not crash the UI
        return False


def persian_families(refresh: bool = False) -> list[str]:
    """Installed families that can render Persian letters, preferred first."""
    global _capable_cache
    if _capable_cache is not None and not refresh:
        return list(_capable_cache)
    families = list(QFontDatabase.families())
    capable = [family for family in families if _supports_persian(family)]
    ordered: list[str] = []
    for preferred in FALLBACK_FAMILIES:
        if preferred in capable and preferred not in ordered:
            ordered.append(preferred)
    ordered.extend(sorted(family for family in capable if family not in ordered))
    _capable_cache = ordered
    return list(ordered)


def all_families() -> list[str]:
    return sorted(QFontDatabase.families())


def resolve_family(preferred: str = "", custom_file: str = "") -> str:
    """Best usable family: the saved choice, else a Persian fallback."""
    installed = set(QFontDatabase.families())
    if preferred and preferred in installed:
        return preferred
    candidates = persian_families()
    for fallback in FALLBACK_FAMILIES:
        if fallback in installed and (not candidates or fallback in candidates):
            return fallback
    if candidates:
        return candidates[0]
    if custom_file and Path(custom_file).is_file():
        family, _ = load_font_file(custom_file)
        if family:
            return family
    return ""


def load_font_file(path: str | Path) -> tuple[str, str]:
    """Load a local .ttf/.otf and return (family, error message)."""
    target = Path(path)
    if not target.is_file():
        return "", "فایل فونت پیدا نشد."
    if target.suffix.lower() not in {".ttf", ".otf", ".ttc", ".otc"}:
        return "", "فقط فایل‌های ttf، otf، ttc یا otc پشتیبانی می‌شوند."
    font_id = QFontDatabase.addApplicationFont(str(target))
    if font_id < 0:
        return "", "Qt نتوانست این فایل فونت را بخواند."
    families = QFontDatabase.applicationFontFamilies(font_id)
    if not families:
        return "", "هیچ خانواده فونتی داخل فایل پیدا نشد."
    family = families[0]
    _custom_families[family] = str(target)
    persian_families(refresh=True)
    return family, ""


def apply_font(app: QApplication, family: str, base_size: int) -> str:
    """Apply the family/size to the whole application; return the real family."""
    resolved = resolve_family(family)
    font = QFont(resolved) if resolved else app.font()
    font.setPointSize(max(MIN_SIZE, min(MAX_SIZE, int(base_size))))
    font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
    app.setFont(font)
    return font.family()


def recommended_default() -> str:
    return resolve_family("")
