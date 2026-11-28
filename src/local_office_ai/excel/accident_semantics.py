"""Separate accident counts from casualty counts and name what is missing.

An accident table mixes several quantities that look alike:

* the number of fatal **accidents**,
* the number of people killed,
* the number of injury **accidents**,
* the number of people injured,
* damage-only accidents.

Adding or relabelling them is a factual error, so this module classifies each
label with explicit word evidence and records the pairs that must stay apart.
It also reports which of the required accident statistics the sources do not
contain, and whether a declared total equals its non-overlapping components.
Nothing is summed, repaired or renamed here.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Iterable

from local_office_ai.core.models import DataRole
from local_office_ai.excel.normalization import normalize_digits, normalize_text
from local_office_ai.excel.row_classification import ROW_TOTAL

KIND_ACCIDENT_COUNT = "accident_count"
KIND_FATAL_ACCIDENT = "fatal_accident_count"
KIND_INJURY_ACCIDENT = "injury_accident_count"
KIND_DAMAGE_ACCIDENT = "damage_accident_count"
KIND_FATALITY = "fatality_count"
KIND_INJURY = "injury_count"
KIND_VEHICLE = "vehicle_count"
KIND_OTHER = "other"

KIND_FA = {
    KIND_ACCIDENT_COUNT: "تعداد تصادف",
    KIND_FATAL_ACCIDENT: "تعداد تصادف فوتی",
    KIND_INJURY_ACCIDENT: "تعداد تصادف جرحی",
    KIND_DAMAGE_ACCIDENT: "تعداد تصادف خسارتی",
    KIND_FATALITY: "تعداد افراد فوت‌شده",
    KIND_INJURY: "تعداد افراد مجروح",
    KIND_VEHICLE: "تعداد وسیله نقلیه",
    KIND_OTHER: "سایر",
}

STATUS_PASS = "PASS"
STATUS_WARNING = "WARNING"
STATUS_FAIL = "FAIL"

_ACCIDENT_WORDS = ("تصادف", "تصادفات", "حادثه", "accident", "collision")
_FATAL_WORDS = ("فوتی", "فوت", "کشته", "جانباخت", "fatal", "death")
_INJURY_WORDS = ("جرحی", "مجروح", "زخمی", "مصدوم", "injury", "injured")
_DAMAGE_WORDS = ("خسارتی", "خسارت", "damage")
_PERSON_WORDS = ("افراد", "نفر", "نفرات", "اشخاص", "شده", "شدگان", "persons", "people")
_VEHICLE_WORDS = ("وسیله", "خودرو", "موتور", "vehicle")
_OVERLAP_WORDS = ("سایر", "other", "متفرقه")
_TOTAL_WORDS = ("جمع", "کل", "total", "مجموع")

_REQUIRED = {
    KIND_ACCIDENT_COUNT: "تعداد کل تصادفات",
    KIND_FATAL_ACCIDENT: "تعداد تصادفات فوتی",
    KIND_INJURY_ACCIDENT: "تعداد تصادفات جرحی",
    KIND_DAMAGE_ACCIDENT: "تعداد تصادفات خسارتی",
    KIND_FATALITY: "تعداد افراد فوت‌شده",
    KIND_INJURY: "تعداد افراد مجروح",
}


def _normalize(text: Any) -> str:
    return normalize_text(normalize_digits(str(text or ""))).strip().lower()


def _has(text: str, words: Iterable[str]) -> bool:
    return any(word in text for word in words)


def classify_accident_label(label: str) -> str:
    """Classify an accident-table label; the evidence is the label's own words."""

    text = _normalize(label)
    if not text:
        return KIND_OTHER
    if _has(text, ("تخلف", "violation")) and not _has(
        text, ("تعداد تصادف", "تعداد حادثه", "accident count")
    ):
        # A violation that led to an accident is a violation count, not an accident count.
        return KIND_OTHER
    accident = _has(text, _ACCIDENT_WORDS)
    fatal = _has(text, _FATAL_WORDS)
    injury = _has(text, _INJURY_WORDS)
    damage = _has(text, _DAMAGE_WORDS)
    vehicle = _has(text, _VEHICLE_WORDS)
    if damage and not accident:
        return KIND_DAMAGE_ACCIDENT
    if fatal and injury and not accident:
        return KIND_OTHER
    if accident and fatal:
        return KIND_FATAL_ACCIDENT
    if accident and injury:
        return KIND_INJURY_ACCIDENT
    if accident and damage:
        return KIND_DAMAGE_ACCIDENT
    if accident:
        return KIND_ACCIDENT_COUNT
    if fatal:
        return KIND_FATALITY
    if injury:
        return KIND_INJURY
    if vehicle:
        return KIND_VEHICLE
    return KIND_OTHER


def classify_indicator(label: str, measure: str = "") -> str:
    """Classify a combined ``row - measure`` indicator without guessing.

    A severity row multiplied by a count column is an accident count
    (``فوتی - تعداد تصادف``).  The same severity row next to a casualty column
    is a casualty count (``فوتی - فوتی``).
    """

    row_kind = classify_accident_label(label)
    measure_kind = classify_accident_label(measure) if measure else KIND_OTHER
    if measure_kind == KIND_ACCIDENT_COUNT:
        if row_kind == KIND_FATALITY:
            return KIND_FATAL_ACCIDENT
        if row_kind == KIND_INJURY:
            return KIND_INJURY_ACCIDENT
        if row_kind == KIND_DAMAGE_ACCIDENT:
            return KIND_DAMAGE_ACCIDENT
        return KIND_ACCIDENT_COUNT
    if measure_kind != KIND_OTHER:
        return measure_kind
    return row_kind if row_kind in {
        KIND_ACCIDENT_COUNT,
        KIND_FATAL_ACCIDENT,
        KIND_INJURY_ACCIDENT,
        KIND_DAMAGE_ACCIDENT,
        KIND_FATALITY,
        KIND_INJURY,
    } else KIND_OTHER


@dataclass
class AccidentFinding:
    code: str
    status: str
    message: str
    sheet: str = ""
    block: str = ""
    evidence: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "status": self.status,
            "message": self.message,
            "sheet": self.sheet,
            "block": self.block,
            "evidence": dict(self.evidence),
        }


@dataclass
class AccidentMetric:
    sheet: str
    block: str
    label: str
    kind: str
    address: str = ""

    @property
    def kind_label(self) -> str:
        return KIND_FA.get(self.kind, self.kind)

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "block": self.block,
            "label": self.label,
            "kind": self.kind,
            "kind_label": self.kind_label,
            "address": self.address,
        }


@dataclass
class AccidentSemanticReport:
    source: str = ""
    tables: int = 0
    metrics: list[AccidentMetric] = field(default_factory=list)
    findings: list[AccidentFinding] = field(default_factory=list)
    missing: list[str] = field(default_factory=list)

    @property
    def detected_kinds(self) -> set[str]:
        return {item.kind for item in self.metrics}

    @property
    def is_accident_source(self) -> bool:
        return bool(self.metrics)

    @property
    def status(self) -> str:
        if any(item.status == STATUS_FAIL for item in self.findings):
            return STATUS_FAIL
        if any(item.status == STATUS_WARNING for item in self.findings):
            return STATUS_WARNING
        return STATUS_PASS

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "tables": self.tables,
            "status": self.status,
            "metrics": [item.to_dict() for item in self.metrics],
            "findings": [item.to_dict() for item in self.findings],
            "missing": list(self.missing),
        }


def _block_is_accident_table(block: Any) -> bool:
    labels = [str(getattr(block, "title", "") or "")]
    labels.extend(str(getattr(column, "header", "") or "") for column in block.columns)
    labels.extend(_data_labels(block).values())
    accident_kinds = {
        KIND_ACCIDENT_COUNT,
        KIND_FATAL_ACCIDENT,
        KIND_INJURY_ACCIDENT,
        KIND_DAMAGE_ACCIDENT,
        KIND_FATALITY,
        KIND_INJURY,
    }
    return any(classify_accident_label(label) in accident_kinds for label in labels)


def _data_labels(block: Any) -> dict[int, str]:
    from local_office_ai.excel.indicator_periods import _cells_by_address

    cells = _cells_by_address(block)
    labels: dict[int, str] = {}
    label_columns = [
        column
        for column in block.columns
        if column.semantic_role
        in {
            DataRole.IDENTIFIER.value,
            DataRole.GROUP.value,
            DataRole.CATEGORICAL.value,
            DataRole.TEXT.value,
        }
    ]
    for column in label_columns:
        for (row, index), cell in cells.items():
            if index != column.index or cell.value in (None, ""):
                continue
            text = str(cell.value).strip()
            if text:
                labels.setdefault(row, text)
    return labels


def _decimal(value: Any) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    try:
        return Decimal(normalize_digits(str(value)).replace(",", "").strip())
    except Exception:  # noqa: BLE001 - non-numeric cell
        return None


def build_accident_semantics(analysis: Any, *, source: str = "") -> AccidentSemanticReport:
    from local_office_ai.excel.indicator_periods import build_indicator_periods

    report = AccidentSemanticReport(source=source)
    extraction = getattr(analysis, "extraction", None)
    if extraction is None:
        return report

    indicators = build_indicator_periods(analysis, source=source)
    seen: set[tuple[str, str, str]] = set()
    accident_blocks: set[tuple[str, str]] = set()
    for observation in indicators.observations:
        kind = classify_indicator(observation.row_label or observation.indicator, observation.measure)
        if kind == KIND_OTHER:
            continue
        key = (observation.sheet, observation.block, observation.indicator)
        if key in seen:
            continue
        seen.add(key)
        accident_blocks.add((observation.sheet, observation.block))
        report.metrics.append(
            AccidentMetric(
                sheet=observation.sheet,
                block=observation.block,
                label=observation.indicator,
                kind=kind,
                address=observation.address,
            )
        )

    for sheet in getattr(extraction, "sheets", []) or []:
        for block in getattr(sheet, "blocks", []) or []:
            if not _block_is_accident_table(block):
                continue
            for column in getattr(block, "columns", []) or []:
                label = str(getattr(column, "header", "") or "")
                kind = classify_accident_label(label)
                if kind == KIND_OTHER:
                    continue
                accident_blocks.add((getattr(block, "sheet", ""), getattr(block, "block_id", "")))
                key = (getattr(block, "sheet", ""), getattr(block, "block_id", ""), label)
                if key in seen:
                    continue
                seen.add(key)
                report.metrics.append(
                    AccidentMetric(
                        sheet=getattr(block, "sheet", ""),
                        block=getattr(block, "block_id", ""),
                        label=label,
                        kind=kind,
                        address=str(getattr(column, "header_address", "") or ""),
                    )
                )
            _inspect_block(block, report)

    report.tables = len(accident_blocks)
    if report.tables:
        missing = [
            description
            for kind, description in _REQUIRED.items()
            if kind not in report.detected_kinds
        ]
        report.missing = missing
        if missing:
            report.findings.append(
                AccidentFinding(
                    code="missing_accident_statistics",
                    status=STATUS_WARNING,
                    message=(
                        "این شاخص‌ها در منابع ارائه‌شده موجود نیستند و نباید حدس زده شوند: "
                        + "، ".join(missing)
                    ),
                    evidence={"missing": missing},
                )
            )
        for accident_kind, casualty_kind in (
            (KIND_FATAL_ACCIDENT, KIND_FATALITY),
            (KIND_INJURY_ACCIDENT, KIND_INJURY),
        ):
            if accident_kind in report.detected_kinds and casualty_kind in report.detected_kinds:
                report.findings.append(
                    AccidentFinding(
                        code=f"count_vs_casualty::{accident_kind}",
                        status=STATUS_PASS,
                        message=(
                            f"«{KIND_FA[accident_kind]}» و «{KIND_FA[casualty_kind]}» جدا "
                            "نگه داشته شدند؛ این دو یک عدد نیستند و با هم جمع نمی‌شوند."
                        ),
                        evidence={
                            "accident_metric": KIND_FA[accident_kind],
                            "casualty_metric": KIND_FA[casualty_kind],
                        },
                    )
                )
            elif casualty_kind in report.detected_kinds and accident_kind not in report.detected_kinds:
                report.findings.append(
                    AccidentFinding(
                        code=f"casualty_without_accident_count::{casualty_kind}",
                        status=STATUS_WARNING,
                        message=(
                            f"فقط «{KIND_FA[casualty_kind]}» موجود است؛ "
                            f"«{KIND_FA[accident_kind]}» در منبع نیست و جای آن حدس زده نمی‌شود."
                        ),
                    )
                )
    return report


def _inspect_block(block: Any, report: AccidentSemanticReport) -> None:
    from local_office_ai.excel.indicator_periods import _cells_by_address

    cells = _cells_by_address(block)
    labels = _data_labels(block)
    measure_columns = [
        column
        for column in block.columns
        if column.semantic_role
        in {DataRole.NUMERIC_MEASURE.value, DataRole.CALCULATED.value}
    ]
    total_rows = {
        int(item.get("row"))
        for item in getattr(block, "row_roles", []) or []
        if isinstance(item, dict) and item.get("role") == ROW_TOTAL
    }
    overlap_labels = [
        label
        for label in labels.values()
        if _has(_normalize(label), _OVERLAP_WORDS)
    ]
    total_labels = [
        label
        for label in labels.values()
        if _has(_normalize(label), _TOTAL_WORDS)
    ]
    if overlap_labels and total_labels:
        report.findings.append(
            AccidentFinding(
                code="overlapping_categories_present",
                status=STATUS_WARNING,
                message=(
                    "دسته‌های «سایر» و ردیف «جمع/کل» در یک جدول دیده شد؛ این ردیف‌ها فقط "
                    "وقتی جمع می‌شوند که استقلال دسته‌ها اثبات شود. مقایسه شاخص، ردیف "
                    "جمع/کل را وارد محاسبه نمی‌کند."
                ),
                sheet=getattr(block, "sheet", ""),
                block=getattr(block, "block_id", ""),
                evidence={
                    "overlap_labels": overlap_labels,
                    "total_labels": total_labels,
                },
            )
        )

    if measure_columns and labels:
        _check_declared_total(block, cells, measure_columns, total_rows, report)


def _check_declared_total(
    block: Any,
    cells: dict[tuple[int, int], Any],
    measure_columns: list[Any],
    total_rows: set[int],
    report: AccidentSemanticReport,
) -> None:
    if not total_rows:
        return
    roles = {
        int(item.get("row"))
        for item in getattr(block, "row_roles", []) or []
        if isinstance(item, dict) and item.get("role") in {"data", "data_missing", "category"}
    }
    for column in measure_columns:
        for total_row in total_rows:
            cell = cells.get((total_row, column.index))
            declared = _decimal(getattr(cell, "value", None)) if cell is not None else None
            if declared is None:
                continue
            components = [
                _decimal(getattr(cells.get((row, column.index)), "value", None))
                for row in sorted(roles)
            ]
            values = [value for value in components if value is not None]
            if not values:
                continue
            computed = sum(values, Decimal("0"))
            if computed == declared:
                report.findings.append(
                    AccidentFinding(
                        code="declared_total_matches_components",
                        status=STATUS_PASS,
                        message=(
                            f"جمع ردیف‌های تفصیلی ستون «{getattr(column, 'header', '')}» با "
                            "ردیف جمع/کل برابر است."
                        ),
                        sheet=getattr(block, "sheet", ""),
                        block=getattr(block, "block_id", ""),
                        evidence={
                            "column": getattr(column, "letter", ""),
                            "declared": str(declared),
                            "computed": str(computed),
                            "cells": f"{getattr(cell, 'address', '')}",
                        },
                    )
                )
            else:
                report.findings.append(
                    AccidentFinding(
                        code="declared_total_mismatch",
                        status=STATUS_WARNING,
                        message=(
                            f"جمع ردیف‌های تفصیلی ستون «{getattr(column, 'header', '')}» با "
                            f"ردیف جمع/کل برابر نیست (ثبت‌شده {declared}، محاسبه‌شده {computed})."
                        ),
                        sheet=getattr(block, "sheet", ""),
                        block=getattr(block, "block_id", ""),
                        evidence={
                            "column": getattr(column, "letter", ""),
                            "declared": str(declared),
                            "computed": str(computed),
                            "cells": f"{getattr(cell, 'address', '')}",
                        },
                    )
                )


def combine_accident_semantics(
    reports: Iterable[AccidentSemanticReport],
) -> AccidentSemanticReport:
    combined = AccidentSemanticReport(source="folder")
    for report in reports:
        combined.tables += report.tables
        combined.metrics.extend(report.metrics)
        combined.findings.extend(report.findings)
        combined.missing.extend(
            item for item in report.missing if item not in combined.missing
        )
    return combined
