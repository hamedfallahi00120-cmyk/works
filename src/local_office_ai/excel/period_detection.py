"""Period detection from every place a workbook may carry a period.

A period is not always a column header.  Real administrative workbooks write it
in a worksheet name (``فروردین 1404``), a block title (``گزارش روزانه 1403``),
a row label or the source filename.  This module gathers those five evidence
sources, keeps the evidence for every detected token and reports
``not_detected`` instead of inventing a period when nothing supports it.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable

from local_office_ai.excel.normalization import normalize_digits, normalize_text
from local_office_ai.excel.sheet_analyzer import JALALI_MONTHS, extract_periods

SOURCE_COLUMN_HEADER = "column_header"
SOURCE_SHEET_NAME = "sheet_name"
SOURCE_BLOCK_TITLE = "block_title"
SOURCE_ROW_LABEL = "row_label"
SOURCE_FILENAME = "filename"

SOURCE_LABELS_FA = {
    SOURCE_COLUMN_HEADER: "عنوان ستون",
    SOURCE_SHEET_NAME: "نام برگه",
    SOURCE_BLOCK_TITLE: "عنوان جدول/بلوک",
    SOURCE_ROW_LABEL: "برچسب ردیف",
    SOURCE_FILENAME: "نام فایل",
}

# Explicit header evidence outranks a title, a title outranks a row label, and a
# sheet name outranks the filename.  A single weak source is still reported, but
# with the confidence of that source only.
SOURCE_CONFIDENCE = {
    SOURCE_COLUMN_HEADER: 0.95,
    SOURCE_BLOCK_TITLE: 0.85,
    SOURCE_ROW_LABEL: 0.80,
    SOURCE_SHEET_NAME: 0.70,
    SOURCE_FILENAME: 0.60,
}

KIND_YEAR = "year"
KIND_MONTH = "month"
KIND_MONTH_YEAR = "month_year"

STATUS_DETECTED = "detected"
STATUS_NOT_DETECTED = "not_detected"
STATUS_CONFLICTING = "conflicting"

# Sources that describe a whole sheet/block rather than one column.  Disagreement
# between two of them is a real conflict; a multi-year header is not.
SCALAR_SOURCES = (SOURCE_SHEET_NAME, SOURCE_BLOCK_TITLE, SOURCE_FILENAME)


@dataclass(frozen=True)
class PeriodSource:
    """One text that may carry a period, plus where it came from."""

    source: str
    text: str
    address: str = ""
    confidence: float | None = None

    @property
    def effective_confidence(self) -> float:
        if self.confidence is not None:
            return self.confidence
        return SOURCE_CONFIDENCE.get(self.source, 0.5)


@dataclass
class PeriodEvidence:
    token: str
    kind: str
    source: str
    text: str
    address: str = ""
    confidence: float = 0.5

    @property
    def source_label(self) -> str:
        return SOURCE_LABELS_FA.get(self.source, self.source)

    def to_dict(self) -> dict[str, Any]:
        return {
            "token": self.token,
            "kind": self.kind,
            "source": self.source,
            "source_label": self.source_label,
            "text": self.text,
            "address": self.address,
            "confidence": round(self.confidence, 3),
        }


@dataclass
class DetectedPeriod:
    token: str
    kind: str
    confidence: float
    sources: list[str] = field(default_factory=list)
    evidence: list[PeriodEvidence] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "token": self.token,
            "kind": self.kind,
            "confidence": round(self.confidence, 3),
            "sources": list(self.sources),
            "evidence": [item.to_dict() for item in self.evidence],
        }


@dataclass
class PeriodDetection:
    periods: list[DetectedPeriod] = field(default_factory=list)
    years: list[str] = field(default_factory=list)
    months: list[str] = field(default_factory=list)
    evidence: list[PeriodEvidence] = field(default_factory=list)
    status: str = STATUS_NOT_DETECTED
    confidence: float = 0.0
    conflicts: list[dict[str, Any]] = field(default_factory=list)

    @property
    def tokens(self) -> list[str]:
        return [item.token for item in self.periods]

    @property
    def sources(self) -> list[str]:
        return sorted({item.source for item in self.evidence})

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "confidence": round(self.confidence, 3),
            "periods": [item.to_dict() for item in self.periods],
            "years": list(self.years),
            "months": list(self.months),
            "sources": self.sources,
            "evidence": [item.to_dict() for item in self.evidence],
            "conflicts": [dict(item) for item in self.conflicts],
        }


_PERSIAN_LETTERS = "آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءأإؤئ"


def _months_in(text: str) -> list[str]:
    """Month names that appear as words, never inside another word.

    ``دی`` must match ``دی 1403`` but not ``ردیف``; ``فروردین‌ماه`` still counts.
    """
    found: list[str] = []
    for month in JALALI_MONTHS:
        pattern = rf"(?<![{_PERSIAN_LETTERS}]){month}(?:ماه)?(?![{_PERSIAN_LETTERS}])"
        if re.search(pattern, text):
            found.append(month)
    return found


def _tokens(text: str) -> list[tuple[str, str]]:
    """Most specific period token(s) inside one text, with their kind."""
    normalized = normalize_text(normalize_digits(str(text or "")))
    if not normalized:
        return []
    years = extract_periods([normalized])
    months = _months_in(normalized)
    if months and len(years) == 1:
        return [(f"{month} {years[0]}", KIND_MONTH_YEAR) for month in months]
    tokens: list[tuple[str, str]] = []
    if len(years) > 1:
        tokens.extend((year, KIND_YEAR) for year in years)
        tokens.extend((month, KIND_MONTH) for month in months)
        return tokens
    if years:
        return [(years[0], KIND_YEAR)]
    if months:
        return [(month, KIND_MONTH) for month in months]
    return []


def detect_periods(sources: Iterable[PeriodSource]) -> PeriodDetection:
    """Aggregate period evidence from all supplied sources."""
    detection = PeriodDetection()
    by_token: dict[str, DetectedPeriod] = {}
    for source in sources:
        if not source.text:
            continue
        for token, kind in _tokens(source.text):
            evidence = PeriodEvidence(
                token=token,
                kind=kind,
                source=source.source,
                text=source.text,
                address=source.address,
                confidence=source.effective_confidence,
            )
            detection.evidence.append(evidence)
            period = by_token.get(token)
            if period is None:
                period = DetectedPeriod(token=token, kind=kind, confidence=evidence.confidence)
                by_token[token] = period
            period.confidence = max(period.confidence, evidence.confidence)
            period.evidence.append(evidence)
            if source.source not in period.sources:
                period.sources.append(source.source)

    detection.periods = _merge_scalar_years(list(by_token.values()))
    if detection.periods:
        detection.status = STATUS_DETECTED
        detection.confidence = max(item.confidence for item in detection.periods)

    years: list[str] = []
    months: list[str] = []
    for period in detection.periods:
        if period.kind == KIND_MONTH_YEAR:
            month, _, year = period.token.rpartition(" ")
            if month and month not in months:
                months.append(month)
            if year and year not in years:
                years.append(year)
        elif period.kind == KIND_YEAR and period.token not in years:
            years.append(period.token)
        elif period.kind == KIND_MONTH and period.token not in months:
            months.append(period.token)
    detection.years = years
    detection.months = months
    return detection


def _merge_scalar_years(periods: list[DetectedPeriod]) -> list[DetectedPeriod]:
    """Fold a scalar-only bare year into the matching ``month_year`` token.

    ``فروردین 1403`` on the sheet and ``1403`` in the filename describe one
    period, not two.  A bare year that came from a column header is kept, because
    it may be a real year column next to a month-specific token.
    """
    month_years: dict[str, list[DetectedPeriod]] = {}
    for period in periods:
        if period.kind == KIND_MONTH_YEAR:
            month_years.setdefault(period.token.rpartition(" ")[2], []).append(period)
    if not month_years:
        return periods
    kept: list[DetectedPeriod] = []
    for period in periods:
        if period.kind != KIND_YEAR or period.token not in month_years:
            kept.append(period)
            continue
        if any(evidence.source not in SCALAR_SOURCES for evidence in period.evidence):
            kept.append(period)
            continue
        for target in month_years[period.token]:
            for evidence in period.evidence:
                if evidence.source not in target.sources:
                    target.sources.append(evidence.source)
    return kept


def years_per_source(detection: PeriodDetection) -> dict[str, list[str]]:
    """Years grouped by the source that produced them (scalar sources only)."""
    grouped: dict[str, list[str]] = {}
    for evidence in detection.evidence:
        if evidence.source not in SCALAR_SOURCES:
            continue
        year = evidence.token if evidence.kind == KIND_YEAR else evidence.token.rpartition(" ")[2]
        bucket = grouped.setdefault(evidence.source, [])
        if year and year not in bucket:
            bucket.append(year)
    return grouped


def months_per_source(detection: PeriodDetection) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = {}
    for evidence in detection.evidence:
        if evidence.source not in SCALAR_SOURCES:
            continue
        if evidence.kind == KIND_MONTH:
            month = evidence.token
        elif evidence.kind == KIND_MONTH_YEAR:
            month = evidence.token.rpartition(" ")[0]
        else:
            continue
        bucket = grouped.setdefault(evidence.source, [])
        if month and month not in bucket:
            bucket.append(month)
    return grouped


def scalar_conflicts(detection: PeriodDetection, *, dimension: str) -> list[dict[str, Any]]:
    """Conflicts between scalar sources for one dimension (``year``/``month``)."""
    grouped = years_per_source(detection) if dimension == KIND_YEAR else months_per_source(detection)
    conflicts: list[dict[str, Any]] = []
    keys = [key for key in grouped if grouped[key]]
    for index, left in enumerate(keys):
        for right in keys[index + 1:]:
            if set(grouped[left]) & set(grouped[right]):
                continue
            conflicts.append(
                {
                    "dimension": dimension,
                    "sources": [left, right],
                    "values": {left: list(grouped[left]), right: list(grouped[right])},
                }
            )
    if conflicts:
        detection.conflicts.extend(conflicts)
    return conflicts


def filename_periods(filename: str) -> PeriodDetection:
    return detect_periods([PeriodSource(SOURCE_FILENAME, filename)])
