"""Duplicate detection, derived-sheet classification and safe aggregation.

Repeating the same records in more than one worksheet or file is normal in
administrative workbooks: a detail sheet, a monthly copy, a summary that reads
the detail with formulas, and a folder that contains last month's file again.
Adding those representations together counts the same record several times.

This module records the evidence needed to refuse that addition:

* every extracted block gets a content fingerprint and the identifier/category
  values it covers, so identical tables, copied summaries and overlapping
  record sets are detected with addresses instead of guessed from titles;
* every worksheet is classified as ``primary``, ``summary``, ``derived``,
  ``duplicate`` or ``unknown`` with the formula references, aggregation rows
  and duplicate matches that justify the class;
* parent/child categories and subtotal/grand-total rows are recorded as
  relations, so a parent is never added to its own children;
* each matched indicator pair receives an aggregation decision with six
  explicit checks (indicator equivalence, period compatibility, unit
  compatibility, category relationship, source independence and non-overlapping
  coverage).  Aggregation is only *allowed* when all six pass; uncertain cases
  are ``review`` and are never removed automatically;
* every inclusion, exclusion and blocked decision is kept in the report.

Nothing here aggregates, repairs or deletes a value.  Python remains the only
authority for arithmetic and a later stage may use ``allowed`` decisions only.
"""

from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from enum import Enum
from typing import Any

from local_office_ai.core.models import DataRole, ValueKind
from local_office_ai.core.settings import Settings
from local_office_ai.excel.normalization import (
    normalize_digits,
    normalize_text,
    parse_value,
)
from local_office_ai.excel.period_alignment import compare_period_sets, parse_periods
from local_office_ai.excel.row_classification import (
    ROW_CATEGORY,
    ROW_DATA,
    ROW_DATA_MISSING,
    ROW_DIFFERENCE,
    ROW_GROUP_HEADING,
    ROW_PERCENTAGE,
    ROW_SUBTOTAL,
    ROW_TOTAL,
)
from local_office_ai.excel.sheet_analyzer import is_total_label, summary_label_role

ZERO = Decimal(0)
_SHEET_REF = re.compile(r"(?:'([^']+)'|([^\s'!()=+*/-]+))!")
_MAX_PROFILES = 250
_MAX_RECORD_GROUPS = 200
_MAX_RELATIONS = 200
_MAX_EXCLUSIONS = 600
_MAX_DECISIONS = 300
_MAX_VALUES_PER_COLUMN = 40
_MAX_SHARED_VALUES = 12

# Semantic roles (DataRole values) that identify a row or a group instead of a
# quantity.  They are the only safe basis for record coverage.
_KEY_ROLES = {DataRole.IDENTIFIER.value, DataRole.TEMPORAL_KEY.value}
_CATEGORY_ROLES = {DataRole.CATEGORICAL.value, DataRole.GROUP.value}
_MEASURE_ROLES = {DataRole.NUMERIC_MEASURE.value, DataRole.CALCULATED.value}
_DATA_ROLES = {ROW_DATA, ROW_DATA_MISSING, ROW_CATEGORY, ROW_GROUP_HEADING}
_SUMMARY_ROLES = {ROW_SUBTOTAL, ROW_TOTAL, ROW_DIFFERENCE, ROW_PERCENTAGE}

_SUMMARY_TITLE_WORDS = (
    "خلاصه",
    "جمع‌بندی",
    "جمع بندی",
    "گزارش کل",
    "summary",
    "aggregate",
    "consolidated",
)
_AGGREGATION_TOKENS = {"total", "average", "percent", "rate", "share"}
_POSITIONAL_HEADERS = {
    "ردیف",
    "شماره ردیف",
    "ردیف شماره",
    "row",
    "row number",
    "row no",
    "no",
    "number",
    "#",
}


class SheetRole(str, Enum):
    PRIMARY = "primary"
    SUMMARY = "summary"
    DERIVED = "derived"
    DUPLICATE = "duplicate"
    UNKNOWN = "unknown"


SHEET_ROLE_FA = {
    SheetRole.PRIMARY: "داده اولیه",
    SheetRole.SUMMARY: "خلاصه",
    SheetRole.DERIVED: "مشتق",
    SheetRole.DUPLICATE: "تکراری",
    SheetRole.UNKNOWN: "نامشخص",
}


class CheckStatus(str, Enum):
    PASS = "pass"
    REVIEW = "review"
    BLOCK = "block"


CHECK_STATUS_FA = {
    CheckStatus.PASS: "تأیید",
    CheckStatus.REVIEW: "نیازمند بررسی",
    CheckStatus.BLOCK: "مسدود",
}


class AggregationStatus(str, Enum):
    ALLOWED = "allowed"
    REVIEW = "review"
    BLOCKED = "blocked"


AGGREGATION_STATUS_FA = {
    AggregationStatus.ALLOWED: "مجاز",
    AggregationStatus.REVIEW: "نیازمند بررسی",
    AggregationStatus.BLOCKED: "مسدود",
}


class RelationKind(str, Enum):
    PARENT_OF = "parent_of"
    GRAND_TOTAL_OF = "grand_total_of"
    OVERLAP = "overlap"


RELATION_KIND_FA = {
    RelationKind.PARENT_OF: "والد → فرزندان",
    RelationKind.GRAND_TOTAL_OF: "جمع کل → اجزا",
    RelationKind.OVERLAP: "هم‌پوشانی گروه‌ها",
}


CHECK_LABELS = {
    "indicator_equivalence": "هم‌ارزی شاخص",
    "period_compatibility": "سازگاری دوره",
    "unit_compatibility": "سازگاری واحد",
    "category_relationship": "رابطه دسته/والد-فرزند",
    "source_independence": "استقلال منبع",
    "coverage_non_overlap": "عدم هم‌پوشانی پوشش داده",
}


# --------------------------------------------------------------------------- #
# Value normalization and small helpers
# --------------------------------------------------------------------------- #
def _decimal_text(value: Decimal) -> str:
    if value == 0:
        return "0"
    try:
        text = format(value.normalize(), "f")
    except (ValueError, InvalidOperation):  # pragma: no cover - defensive
        return str(value)
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def _normalize_value(value: Any) -> str:
    """Canonical text used by fingerprints and record comparisons."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, Decimal):
        return _decimal_text(value)
    if isinstance(value, (int, float)):
        try:
            return _decimal_text(Decimal(str(value)))
        except InvalidOperation:  # pragma: no cover - defensive
            return str(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    text = normalize_text(normalize_digits(str(value))).lower().strip()
    if not text:
        return ""
    try:
        number = Decimal(text.replace(",", "").replace("٬", ""))
    except (InvalidOperation, ValueError):
        return text
    return _decimal_text(number)


def _as_decimal(value: Any) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    if isinstance(value, str):
        parsed = parse_value(value)
        if parsed.kind is ValueKind.NUMBER and isinstance(parsed.value, Decimal):
            return parsed.value
    return None


def _close(left: Decimal, right: Decimal) -> bool:
    difference = abs(left - right)
    if difference == 0:
        return True
    return difference <= max(Decimal("0.01"), abs(right) * Decimal("0.005"))


def _digest(payload: Any) -> str:
    text = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:32]


def _header_key(column: Any) -> str:
    parts = [
        str(getattr(column, name, "") or "")
        for name in ("header", "parent_header", "sub_header", "period")
    ]
    return " | ".join(_normalize_value(part) for part in parts if str(part).strip())


def _is_positional(header: str) -> bool:
    return _normalize_value(header) in _POSITIONAL_HEADERS


def _display_header(column: Any) -> str:
    header = str(getattr(column, "header", "") or "").strip()
    if header:
        return header
    letter = str(getattr(column, "letter", "") or "")
    return f"ستون {letter}" if letter else "—"


def _trim(values: list[str]) -> tuple[str, ...]:
    end = len(values)
    while end and not values[end - 1]:
        end -= 1
    return tuple(values[:end])


def _label_tokens(text: str) -> set[str]:
    cleaned = re.sub(
        r"[^\w\u0600-\u06FF]+", " ", normalize_text(normalize_digits(text)).lower()
    )
    return {word for word in cleaned.split() if word}


def _unique(values: Iterable[str], limit: int = _MAX_VALUES_PER_COLUMN) -> list[str]:
    seen: list[str] = []
    for value in values:
        if not value or value in seen:
            continue
        seen.append(value)
        if len(seen) >= limit:
            break
    return seen


# --------------------------------------------------------------------------- #
# Report records
# --------------------------------------------------------------------------- #
@dataclass
class BlockProfile:
    source: str
    sheet: str
    block: str
    address: str
    title: str
    kind: str
    row_count: int
    column_count: int
    table_hash: str
    content_hash: str
    summary_hash: str
    measure_headers: list[str] = field(default_factory=list)
    key_headers: list[str] = field(default_factory=list)
    category_headers: list[str] = field(default_factory=list)
    key_values: dict[str, list[str]] = field(default_factory=dict)
    category_values: dict[str, list[str]] = field(default_factory=dict)
    row_numbers: list[int] = field(default_factory=list)
    summary_row_numbers: list[int] = field(default_factory=list)
    numeric_values: int = 0
    truncated: bool = False
    formula_cells: int = 0
    cross_sheet_formula_cells: int = 0
    referenced_sheets: list[str] = field(default_factory=list)
    period_tokens: list[str] = field(default_factory=list)
    records: list[dict[str, Any]] = field(default_factory=list, repr=False)
    row_signatures: list[tuple[str, ...]] = field(default_factory=list, repr=False)
    summary_signatures: list[tuple[str, ...]] = field(default_factory=list, repr=False)

    @property
    def identity(self) -> str:
        return f"{self.source}::{self.sheet}::{self.block}"

    @property
    def label(self) -> str:
        title = self.title or "—"
        return f"{self.source}/{self.sheet}/{self.block}/«{title}»"

    @property
    def summary_only(self) -> bool:
        return bool(self.summary_signatures) and not self.row_signatures

    def values_for(self, header: str) -> set[str]:
        normalized = _normalize_value(header)
        for key, values in self.key_values.items():
            if _normalize_value(key) == normalized:
                return set(values)
        for key, values in self.category_values.items():
            if _normalize_value(key) == normalized:
                return set(values)
        return set()

    def to_dict(self) -> dict[str, Any]:
        return {
            "identity": self.identity,
            "source": self.source,
            "sheet": self.sheet,
            "block": self.block,
            "address": self.address,
            "title": self.title,
            "kind": self.kind,
            "row_count": self.row_count,
            "column_count": self.column_count,
            "table_hash": self.table_hash,
            "content_hash": self.content_hash,
            "summary_hash": self.summary_hash,
            "summary_only": self.summary_only,
            "measure_headers": list(self.measure_headers),
            "key_headers": list(self.key_headers),
            "category_headers": list(self.category_headers),
            "key_values": {key: list(values) for key, values in self.key_values.items()},
            "category_values": {
                key: list(values) for key, values in self.category_values.items()
            },
            "row_numbers": list(self.row_numbers),
            "summary_row_numbers": list(self.summary_row_numbers),
            "numeric_values": self.numeric_values,
            "truncated": self.truncated,
            "formula_cells": self.formula_cells,
            "cross_sheet_formula_cells": self.cross_sheet_formula_cells,
            "referenced_sheets": list(self.referenced_sheets),
            "period_tokens": list(self.period_tokens),
        }


@dataclass
class SheetProfile:
    source: str
    sheet: str
    role: SheetRole
    confidence: float
    evidence: list[str] = field(default_factory=list)
    referenced_sheets: list[str] = field(default_factory=list)
    formula_cells: int = 0
    cross_sheet_formula_cells: int = 0
    data_blocks: int = 0
    duplicate_of: list[str] = field(default_factory=list)
    summary_rows: int = 0
    title: str = ""

    @property
    def role_label(self) -> str:
        return SHEET_ROLE_FA.get(self.role, self.role.value)

    @property
    def independent(self) -> bool:
        return self.role is SheetRole.PRIMARY

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "sheet": self.sheet,
            "title": self.title,
            "role": self.role.value,
            "role_label": self.role_label,
            "confidence": round(self.confidence, 4),
            "evidence": list(self.evidence),
            "referenced_sheets": list(self.referenced_sheets),
            "formula_cells": self.formula_cells,
            "cross_sheet_formula_cells": self.cross_sheet_formula_cells,
            "data_blocks": self.data_blocks,
            "duplicate_of": list(self.duplicate_of),
            "summary_rows": self.summary_rows,
        }


@dataclass
class DuplicateGroup:
    group_id: str
    kind: str
    identical: bool
    similarity: float
    members: list[BlockProfile]
    evidence: list[str] = field(default_factory=list)
    reason: str = ""

    @property
    def safe_to_aggregate(self) -> bool:
        return False

    def to_dict(self) -> dict[str, Any]:
        return {
            "group_id": self.group_id,
            "kind": self.kind,
            "identical": self.identical,
            "similarity": round(self.similarity, 4),
            "safe_to_aggregate": self.safe_to_aggregate,
            "members": [member.identity for member in self.members],
            "member_labels": [member.label for member in self.members],
            "evidence": list(self.evidence),
            "reason": self.reason,
        }


@dataclass
class RecordDuplicate:
    source: str
    sheet: str
    block: str
    key_headers: list[str]
    key_values: list[str]
    row_numbers: list[int]
    measure_values: list[str]
    kind: str = "identical"
    evidence: list[str] = field(default_factory=list)

    @property
    def label(self) -> str:
        return f"{self.source}/{self.sheet}/{self.block}"

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "sheet": self.sheet,
            "block": self.block,
            "key_headers": list(self.key_headers),
            "key_values": list(self.key_values),
            "row_numbers": list(self.row_numbers),
            "measure_values": list(self.measure_values),
            "kind": self.kind,
            "evidence": list(self.evidence),
        }


@dataclass
class CategoryRelation:
    relation_id: str
    kind: RelationKind
    source: str
    sheet: str
    block: str
    parent_label: str
    child_labels: list[str]
    column: str
    parent_value: str | None = None
    child_sum: str | None = None
    difference: str | None = None
    confirmed: bool = True
    evidence: list[str] = field(default_factory=list)
    addresses: list[str] = field(default_factory=list)
    parent_row: int = 0
    child_rows: list[int] = field(default_factory=list)

    @property
    def kind_label(self) -> str:
        return RELATION_KIND_FA.get(self.kind, self.kind.value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "relation_id": self.relation_id,
            "kind": self.kind.value,
            "kind_label": self.kind_label,
            "source": self.source,
            "sheet": self.sheet,
            "block": self.block,
            "parent_label": self.parent_label,
            "child_labels": list(self.child_labels),
            "column": self.column,
            "parent_value": self.parent_value,
            "child_sum": self.child_sum,
            "difference": self.difference,
            "confirmed": self.confirmed,
            "evidence": list(self.evidence),
            "addresses": list(self.addresses),
            "parent_row": self.parent_row,
            "child_rows": list(self.child_rows),
        }


@dataclass
class AggregationCheck:
    name: str
    status: CheckStatus
    explanation: str

    @property
    def label(self) -> str:
        return CHECK_LABELS.get(self.name, self.name)

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "label": self.label,
            "status": self.status.value,
            "status_label": CHECK_STATUS_FA.get(self.status, self.status.value),
            "explanation": self.explanation,
        }


@dataclass
class AggregationDecision:
    decision_id: str
    metric: str
    kind: str
    members: list[str]
    status: AggregationStatus
    checks: list[AggregationCheck]
    reasons: list[str] = field(default_factory=list)
    evidence: list[str] = field(default_factory=list)
    period_relation: str = ""
    unit_state: str = ""

    @property
    def status_label(self) -> str:
        return AGGREGATION_STATUS_FA.get(self.status, self.status.value)

    @property
    def sources(self) -> set[str]:
        return {member.split("/", 1)[0] for member in self.members if "/" in member}

    @property
    def cross_source(self) -> bool:
        return len(self.sources) > 1

    @property
    def headers(self) -> set[str]:
        found: set[str] = set()
        for member in self.members:
            if "«" in member and "»" in member:
                found.add(member.rsplit("«", 1)[-1].rstrip("»"))
        return found

    def to_dict(self) -> dict[str, Any]:
        return {
            "decision_id": self.decision_id,
            "metric": self.metric,
            "kind": self.kind,
            "members": list(self.members),
            "sources": sorted(self.sources),
            "cross_source": self.cross_source,
            "status": self.status.value,
            "status_label": self.status_label,
            "checks": [check.to_dict() for check in self.checks],
            "reasons": list(self.reasons),
            "evidence": list(self.evidence),
            "period_relation": self.period_relation,
            "unit_state": self.unit_state,
        }


@dataclass
class ExclusionRecord:
    source: str
    sheet: str
    block: str
    address: str
    kind: str
    reason: str
    evidence: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "sheet": self.sheet,
            "block": self.block,
            "address": self.address,
            "kind": self.kind,
            "reason": self.reason,
            "evidence": list(self.evidence),
        }


@dataclass
class DuplicateControlReport:
    source_count: int = 0
    sheets: list[SheetProfile] = field(default_factory=list)
    blocks: list[BlockProfile] = field(default_factory=list)
    duplicate_groups: list[DuplicateGroup] = field(default_factory=list)
    record_duplicates: list[RecordDuplicate] = field(default_factory=list)
    category_relations: list[CategoryRelation] = field(default_factory=list)
    decisions: list[AggregationDecision] = field(default_factory=list)
    exclusions: list[ExclusionRecord] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def allowed_decisions(self) -> list[AggregationDecision]:
        return [item for item in self.decisions if item.status is AggregationStatus.ALLOWED]

    @property
    def review_decisions(self) -> list[AggregationDecision]:
        return [item for item in self.decisions if item.status is AggregationStatus.REVIEW]

    @property
    def blocked_decisions(self) -> list[AggregationDecision]:
        return [item for item in self.decisions if item.status is AggregationStatus.BLOCKED]

    def role_counts(self) -> dict[str, int]:
        counts = {role.value: 0 for role in SheetRole}
        for sheet in self.sheets:
            counts[sheet.role.value] = counts.get(sheet.role.value, 0) + 1
        return counts

    def cross_source_status(
        self, metric: str, sources: Iterable[str] | None = None
    ) -> AggregationStatus | None:
        """Safety of adding one metric across the given files.

        Only decisions whose two members come from different files are
        relevant; a pair inside one block says nothing about adding two files.
        """
        target = _normalize_value(metric)
        selected = {str(item) for item in sources} if sources else None
        statuses: list[AggregationStatus] = []
        for decision in self.decisions:
            if decision.kind != "measure" or not decision.cross_source:
                continue
            if selected is not None and not (decision.sources & selected):
                continue
            candidates = [decision.metric, *decision.headers]
            if target and not any(
                target == _normalize_value(candidate)
                or target in _normalize_value(candidate)
                or _normalize_value(candidate) in target
                for candidate in candidates
                if candidate
            ):
                continue
            statuses.append(decision.status)
        if not statuses:
            return None
        if any(status is AggregationStatus.BLOCKED for status in statuses):
            return AggregationStatus.BLOCKED
        if any(status is AggregationStatus.REVIEW for status in statuses):
            return AggregationStatus.REVIEW
        return AggregationStatus.ALLOWED

    def counts(self) -> dict[str, int]:
        return {
            "sources": self.source_count,
            "sheets": len(self.sheets),
            "blocks": len(self.blocks),
            "duplicate_groups": len(self.duplicate_groups),
            "duplicate_blocks": sum(len(group.members) for group in self.duplicate_groups),
            "record_duplicates": len(self.record_duplicates),
            "category_relations": len(self.category_relations),
            "decisions": len(self.decisions),
            "allowed": len(self.allowed_decisions),
            "review": len(self.review_decisions),
            "blocked": len(self.blocked_decisions),
            "exclusions": len(self.exclusions),
            "warnings": len(self.warnings),
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "counts": self.counts(),
            "role_counts": self.role_counts(),
            "sheets": [item.to_dict() for item in self.sheets],
            "blocks": [item.to_dict() for item in self.blocks],
            "duplicate_groups": [item.to_dict() for item in self.duplicate_groups],
            "record_duplicates": [item.to_dict() for item in self.record_duplicates],
            "category_relations": [item.to_dict() for item in self.category_relations],
            "decisions": [item.to_dict() for item in self.decisions],
            "exclusions": [item.to_dict() for item in self.exclusions],
            "warnings": list(self.warnings),
        }


# --------------------------------------------------------------------------- #
# Block profiles
# --------------------------------------------------------------------------- #
def _cell_value(cells: dict[tuple[int, int], Any], row: int, column: int) -> Any:
    cell = cells.get((row, column))
    return None if cell is None else getattr(cell, "value", None)


def _role_of(column: Any) -> str:
    return str(getattr(column, "semantic_role", "") or "")


def _active_columns(
    columns: list[Any], cells: dict[tuple[int, int], Any], rows: list[int]
) -> list[Any]:
    active: list[Any] = []
    for column in columns:
        if _header_key(column):
            active.append(column)
            continue
        index = int(getattr(column, "index", 0) or 0)
        if any(_normalize_value(_cell_value(cells, row, index)) for row in rows):
            active.append(column)
    return active or columns


def _row_label(cells: dict[tuple[int, int], Any], row: int, columns: list[Any]) -> str:
    for column in columns:
        if _role_of(column) in _MEASURE_ROLES:
            continue
        value = _cell_value(cells, row, int(getattr(column, "index", 0) or 0))
        text = str(value or "").strip()
        if text and _as_decimal(value) is None:
            return text
    return ""


def _block_profile(source: str, block: Any, known_sheets: set[str]) -> BlockProfile:
    bounds = dict(getattr(block, "boundaries", {}) or {})
    first_data = int(bounds.get("first_data_row") or 0)
    last_data = int(bounds.get("last_data_row") or 0)
    block_end = int(bounds.get("block_end_row") or last_data)
    columns = list(getattr(block, "columns", []) or [])
    cells = {
        (int(getattr(cell, "row", 0)), int(getattr(cell, "column", 0))): cell
        for cell in (getattr(block, "cells", []) or [])
    }
    roles: dict[int, str] = {}
    labels: dict[int, str] = {}
    for record in getattr(block, "row_roles", []) or []:
        if not isinstance(record, dict):
            continue
        row = int(record.get("row") or 0)
        if not row:
            continue
        roles[row] = str(record.get("role") or "")
        labels[row] = str(record.get("label") or "")

    data_rows = [
        row for row in range(first_data, last_data + 1) if roles.get(row) in _DATA_ROLES
    ]
    if not data_rows:
        data_rows = [
            row
            for row in range(first_data, last_data + 1)
            if any(
                _normalize_value(_cell_value(cells, row, int(getattr(column, "index", 0) or 0)))
                for column in columns
            )
        ]
    summary_rows = [
        row for row in range(first_data, block_end + 1) if roles.get(row) in _SUMMARY_ROLES
    ]
    active = _active_columns(columns, cells, data_rows)

    key_columns = [column for column in active if _role_of(column) in _KEY_ROLES]
    category_columns = [column for column in active if _role_of(column) in _CATEGORY_ROLES]
    measure_columns = [column for column in active if _role_of(column) in _MEASURE_ROLES]
    if not measure_columns:
        measure_columns = [
            column
            for column in active
            if str(getattr(column, "data_type", "")) == ValueKind.NUMBER.value
            and column not in key_columns
        ]
    if not key_columns:
        key_columns = category_columns
    key_columns = [
        column for column in key_columns if not _is_positional(_display_header(column))
    ]

    def signature(row: int) -> tuple[str, ...]:
        return _trim(
            [
                _normalize_value(_cell_value(cells, row, int(getattr(column, "index", 0) or 0)))
                for column in active
            ]
        )

    measure_indexes = [
        int(getattr(column, "index", 0) or 0) for column in measure_columns
    ]
    row_signatures = [value for value in (signature(row) for row in data_rows) if any(value)]
    summary_signatures = [
        value
        for row in summary_rows
        if not measure_indexes
        or any(
            _normalize_value(_cell_value(cells, row, index))
            for index in measure_indexes
        )
        for value in [signature(row)]
        if any(value)
    ]
    headers = [_header_key(column) for column in active]

    key_values: dict[str, list[str]] = {}
    for column in key_columns:
        header = _display_header(column)
        index = int(getattr(column, "index", 0) or 0)
        values = [
            _normalize_value(_cell_value(cells, row, index))
            for row in data_rows
        ]
        key_values[header] = _unique([*key_values.get(header, []), *values])

    category_values: dict[str, list[str]] = {}
    for column in category_columns:
        header = _display_header(column)
        index = int(getattr(column, "index", 0) or 0)
        values = [
            _normalize_value(_cell_value(cells, row, index))
            for row in data_rows
        ]
        category_values[header] = _unique([*category_values.get(header, []), *values])

    records: list[dict[str, Any]] = []
    for row in data_rows:
        key = tuple(
            _normalize_value(
                _cell_value(cells, row, int(getattr(column, "index", 0) or 0))
            )
            for column in key_columns
        )
        if not any(key):
            continue
        measures = tuple(
            _normalize_value(
                _cell_value(cells, row, int(getattr(column, "index", 0) or 0))
            )
            for column in measure_columns
        )
        records.append(
            {
                "row": row,
                "key": key,
                "measures": measures,
                "label": labels.get(row) or _row_label(cells, row, active),
            }
        )

    formula_cells = 0
    cross_sheet_formula_cells = 0
    referenced_sheets: list[str] = []
    for cell in cells.values():
        formula = getattr(cell, "formula", None)
        if not isinstance(formula, str) or not formula.startswith("="):
            continue
        formula_cells += 1
        names = {
            (match.group(1) or match.group(2) or "").strip()
            for match in _SHEET_REF.finditer(formula)
        }
        external = {
            name
            for name in names
            if name in known_sheets and name != str(getattr(block, "sheet", "") or "")
        }
        if external:
            cross_sheet_formula_cells += 1
            for name in sorted(external):
                if name not in referenced_sheets:
                    referenced_sheets.append(name)

    numeric_values = sum(
        1
        for row in data_rows
        for column in measure_columns
        if _as_decimal(
            _cell_value(cells, row, int(getattr(column, "index", 0) or 0))
        )
        is not None
    )

    return BlockProfile(
        source=source,
        sheet=str(getattr(block, "sheet", "") or ""),
        block=str(getattr(block, "block_id", "") or ""),
        address=str(bounds.get("address") or ""),
        title=str(getattr(block, "title", "") or ""),
        kind=str(getattr(block, "kind", "") or ""),
        row_count=len(row_signatures),
        column_count=len(active),
        table_hash=_digest({"headers": headers, "rows": row_signatures}),
        content_hash=_digest(
            {"headers": headers, "rows": [*row_signatures, *summary_signatures]}
        ),
        summary_hash=(
            _digest({"headers": headers, "rows": summary_signatures})
            if summary_signatures
            else ""
        ),
        measure_headers=[_display_header(column) for column in measure_columns],
        key_headers=[_display_header(column) for column in key_columns],
        category_headers=[_display_header(column) for column in category_columns],
        key_values=key_values,
        category_values=category_values,
        row_numbers=[record["row"] for record in records],
        summary_row_numbers=summary_rows,
        numeric_values=numeric_values,
        truncated=bool(getattr(block, "cells_truncated", False)),
        formula_cells=formula_cells,
        cross_sheet_formula_cells=cross_sheet_formula_cells,
        referenced_sheets=referenced_sheets,
        period_tokens=[
            str(item.get("token") or "")
            for item in (getattr(block, "periods", []) or [])
            if isinstance(item, dict) and item.get("token")
        ],
        records=records,
        row_signatures=row_signatures,
        summary_signatures=summary_signatures,
    )


def profiles_from_analysis(analysis: Any, source: str = "") -> list[BlockProfile]:
    """Fingerprint every extracted block of one analysis."""
    extraction = getattr(analysis, "extraction", None)
    if extraction is None:
        return []
    source_name = source or str(getattr(extraction, "source_filename", "") or "")
    known_sheets = {
        str(getattr(sheet, "name", "") or "")
        for sheet in (getattr(extraction, "sheets", []) or [])
    }
    profiles: list[BlockProfile] = []
    for sheet in getattr(extraction, "sheets", []) or []:
        for block in getattr(sheet, "blocks", []) or []:
            profiles.append(_block_profile(source_name, block, known_sheets))
    return profiles


# --------------------------------------------------------------------------- #
# Duplicate detection
# --------------------------------------------------------------------------- #
def _eligible(profile: BlockProfile) -> bool:
    signatures = len(profile.row_signatures) + len(profile.summary_signatures)
    return bool(
        profile.content_hash
        and signatures >= 2
        and profile.column_count >= 2
        and (profile.numeric_values >= 1 or profile.summary_signatures)
    )


def _duplicate_kind(profile: BlockProfile) -> str:
    if profile.summary_only:
        return "summary"
    if is_total_label(profile.title) or summary_label_role(profile.title):
        return "summary"
    text = f"{_normalize_value(profile.title)} {_normalize_value(profile.sheet)}"
    if text.strip() and any(word in text for word in _SUMMARY_TITLE_WORDS):
        return "summary"
    return "table"


def _similarity(left: BlockProfile, right: BlockProfile) -> float:
    if left.column_count != right.column_count:
        return 0.0
    # Header token similarity keeps two unrelated tables with the same shape apart.
    left_tokens = {
        token
        for header in [*left.measure_headers, *left.key_headers, *left.category_headers]
        for token in _label_tokens(header)
    }
    right_tokens = {
        token
        for header in [*right.measure_headers, *right.key_headers, *right.category_headers]
        for token in _label_tokens(header)
    }
    if left_tokens and right_tokens:
        union = left_tokens | right_tokens
        if union and len(left_tokens & right_tokens) / len(union) < 0.75:
            return 0.0
    left_rows = set(left.row_signatures)
    right_rows = set(right.row_signatures)
    if len(left_rows) < 2 or len(right_rows) < 2:
        return 0.0
    union = left_rows | right_rows
    if not union:
        return 0.0
    return len(left_rows & right_rows) / len(union)


def _compatible_clusters(members: list[BlockProfile]) -> list[list[BlockProfile]]:
    """Split one fingerprint bucket by reporting period.

    Two months with the same total are not a copied summary; only blocks that
    describe the same period can repeat the same records.
    """
    clusters: list[list[BlockProfile]] = []
    for profile in members:
        for cluster in clusters:
            if all(_periods_compatible(profile, other) for other in cluster):
                cluster.append(profile)
                break
        else:
            clusters.append([profile])
    return clusters


def _duplicate_groups(profiles: list[BlockProfile]) -> list[DuplicateGroup]:
    eligible = [profile for profile in profiles if _eligible(profile)]
    groups: list[DuplicateGroup] = []
    grouped: set[str] = set()
    buckets: dict[str, list[BlockProfile]] = {}
    for profile in eligible:
        buckets.setdefault(profile.content_hash, []).append(profile)
    for members in buckets.values():
        for cluster in _compatible_clusters(members):
            if len(cluster) < 2:
                continue
            cluster = sorted(cluster, key=lambda item: item.identity)
            kind = (
                "summary"
                if all(_duplicate_kind(item) == "summary" for item in cluster)
                else "table"
            )
            groups.append(
                DuplicateGroup(
                    group_id=f"D{len(groups) + 1}",
                    kind=kind,
                    identical=True,
                    similarity=1.0,
                    members=cluster,
                    evidence=[
                        (
                            f"اثر انگشت محتوا یکسان است ({cluster[0].content_hash[:12]}…)؛ "
                            f"{cluster[0].row_count} ردیف داده، {cluster[0].column_count} ستون، "
                            f"{cluster[0].numeric_values} مقدار عددی."
                        )
                    ],
                    reason=(
                        "همان رکوردها/جمع‌ها در چند محل تکرار شده‌اند؛ جمع آن‌ها شمارش "
                        "دوباره است و مجاز نیست."
                    ),
                )
            )
            grouped.update(member.identity for member in cluster)

    # A summary sheet may copy only the totals of a detail table.  Its content
    # hash differs (the detail rows are absent) but the summary fingerprint matches.
    summary_buckets: dict[str, list[BlockProfile]] = {}
    for profile in eligible:
        if profile.summary_hash and profile.identity not in grouped:
            summary_buckets.setdefault(profile.summary_hash, []).append(profile)
    for members in summary_buckets.values():
        for cluster in _compatible_clusters(members):
            if len(cluster) < 2:
                continue
            cluster = sorted(cluster, key=lambda item: item.identity)
            groups.append(
                DuplicateGroup(
                    group_id=f"D{len(groups) + 1}",
                    kind="summary",
                    identical=True,
                    similarity=1.0,
                    members=cluster,
                    evidence=[
                        (
                            f"اثر انگشت ردیف‌های جمع یکسان است ({cluster[0].summary_hash[:12]}…)؛ "
                            "همان جمع‌ها در چند برگه تکرار شده‌اند."
                        )
                    ],
                    reason="جمع تکراری نباید به‌عنوان مشاهده مستقل شمرده شود.",
                )
            )
            grouped.update(member.identity for member in cluster)

    remaining = [profile for profile in eligible if profile.identity not in grouped]
    near: list[DuplicateGroup] = []
    for profile in sorted(remaining, key=lambda item: item.identity):
        best: DuplicateGroup | None = None
        best_score = 0.0
        for group in near:
            if group.members[0].column_count != profile.column_count:
                continue
            score = max(
                (
                    _similarity(profile, member)
                    for member in group.members
                    if _periods_compatible(profile, member)
                ),
                default=0.0,
            )
            if score > best_score:
                best, best_score = group, score
        if best is not None and best_score >= 0.8:
            best.members.append(profile)
            best.similarity = min(best.similarity, best_score)
            best.evidence.append(
                f"هم‌پوشانی ردیف‌ها با گروه {best.group_id} برابر {best_score:.0%} است."
            )
        else:
            near.append(
                DuplicateGroup(
                    group_id=f"N{len(near) + 1}",
                    kind="table",
                    identical=False,
                    similarity=1.0,
                    members=[profile],
                    reason="ردیف‌های مشترک زیاد است؛ بازنمایی تکراری احتمالی، بدون حذف خودکار.",
                )
            )
    for group in near:
        if len(group.members) < 2:
            continue
        group.group_id = f"D{len(groups) + 1}"
        group.evidence.insert(
            0,
            f"هم‌پوشانی ردیف‌ها {group.similarity:.0%} است؛ محتوا کاملاً یکسان نیست.",
        )
        groups.append(group)
    return groups


def _record_duplicates(profile: BlockProfile) -> list[RecordDuplicate]:
    grouped: dict[tuple[str, ...], list[dict[str, Any]]] = {}
    for record in profile.records:
        grouped.setdefault(record["key"], []).append(record)
    duplicates: list[RecordDuplicate] = []
    for key, rows in grouped.items():
        if len(rows) < 2:
            continue
        measure_sets = {tuple(row["measures"]) for row in rows}
        kind = "identical" if len(measure_sets) == 1 else "conflicting"
        measures = [
            f"{header}={value}"
            for header, value in zip(profile.measure_headers, rows[0]["measures"])
            if value
        ]
        duplicates.append(
            RecordDuplicate(
                source=profile.source,
                sheet=profile.sheet,
                block=profile.block,
                key_headers=list(profile.key_headers),
                key_values=list(key),
                row_numbers=[int(row["row"]) for row in rows],
                measure_values=measures,
                kind=kind,
                evidence=[
                    (
                        "همان کلید و همان مقدارها چند بار تکرار شده‌اند."
                        if kind == "identical"
                        else "کلید یکسان با مقدارهای متفاوت؛ نیازمند بررسی، بدون حذف خودکار."
                    )
                ],
            )
        )
        if len(duplicates) >= 50:
            break
    return duplicates


# --------------------------------------------------------------------------- #
# Parent/child categories and overlapping groups
# --------------------------------------------------------------------------- #
def _block_context(block: Any) -> tuple[
    dict[tuple[int, int], Any], dict[int, str], dict[int, str], int, int
]:
    bounds = dict(getattr(block, "boundaries", {}) or {})
    first_data = int(bounds.get("first_data_row") or 0)
    last_data = int(bounds.get("last_data_row") or 0)
    cells = {
        (int(getattr(cell, "row", 0)), int(getattr(cell, "column", 0))): cell
        for cell in (getattr(block, "cells", []) or [])
    }
    roles: dict[int, str] = {}
    labels: dict[int, str] = {}
    for record in getattr(block, "row_roles", []) or []:
        if not isinstance(record, dict):
            continue
        row = int(record.get("row") or 0)
        if row:
            roles[row] = str(record.get("role") or "")
            labels[row] = str(record.get("label") or "")
    return cells, roles, labels, first_data, last_data


def _row_text(
    cells: dict[tuple[int, int], Any],
    labels: dict[int, str],
    row: int,
    label_column: Any | None,
) -> str:
    if label_column is not None:
        value = _cell_value(cells, row, int(getattr(label_column, "index", 0) or 0))
        text = str(value or "").strip()
        if text and _as_decimal(value) is None:
            return text
    return labels.get(row) or f"ردیف {row}"


def _category_relations(profile: BlockProfile, block: Any) -> list[CategoryRelation]:
    cells, roles, labels, first_data, last_data = _block_context(block)
    columns = list(getattr(block, "columns", []) or [])
    active = [column for column in columns if _header_key(column)] or columns
    label_column = next(
        (
            column
            for column in active
            if _role_of(column)
            in {*_CATEGORY_ROLES, *_KEY_ROLES, DataRole.TEXT.value}
        ),
        None,
    )
    measure_columns = [column for column in active if _role_of(column) in _MEASURE_ROLES]
    if not measure_columns:
        measure_columns = [
            column
            for column in active
            if str(getattr(column, "data_type", "")) == ValueKind.NUMBER.value
            and _role_of(column) not in _KEY_ROLES
        ]
    if not measure_columns:
        return []

    data_rows = [
        row for row in range(first_data, last_data + 1) if roles.get(row) in _DATA_ROLES
    ]
    if not data_rows:
        data_rows = [
            row
            for row in range(first_data, last_data + 1)
            if any(
                _as_decimal(_cell_value(cells, row, int(getattr(column, "index", 0) or 0)))
                is not None
                for column in measure_columns
            )
        ]
    summary_rows = sorted(profile.summary_row_numbers)
    relations: list[CategoryRelation] = []

    def matched_columns(
        row: int, children: list[int], strict: bool = False
    ) -> list[tuple[Any, Decimal, Decimal]]:
        matched: list[tuple[Any, Decimal, Decimal]] = []
        considered = 0
        for column in measure_columns:
            index = int(getattr(column, "index", 0) or 0)
            parent = _as_decimal(_cell_value(cells, row, index))
            if parent is None:
                continue
            considered += 1
            values = [_as_decimal(_cell_value(cells, child, index)) for child in children]
            if any(value is None for value in values):
                if strict:
                    return []
                continue
            total = sum((value for value in values if value is not None), ZERO)
            if _close(parent, total):
                matched.append((column, parent, total))
            elif strict:
                return []
        if strict and considered and len(matched) != considered:
            return []
        return matched

    def is_parent(row: int, children: list[int]) -> bool:
        matched = matched_columns(row, children)
        if not matched:
            return False
        for column, parent, _total in matched:
            index = int(getattr(column, "index", 0) or 0)
            values = [_as_decimal(_cell_value(cells, child, index)) for child in children]
            numbers = [value for value in values if value is not None]
            if numbers and parent <= max(numbers):
                return False
        return True

    def top_level(rows: list[int]) -> list[int]:
        """Collapse nested parent rows so a grand total is not summed twice.

        ``[parent, child, child, parent, child, child]`` becomes
        ``[parent, parent]`` when each parent equals the sum of the rows that
        follow it.  A row that merely equals one following row is not a parent.
        """
        result: list[int] = []
        index = 0
        while index < len(rows):
            row = rows[index]
            span = 0
            for size in range(2, min(20, len(rows) - index) + 1):
                if is_parent(row, rows[index + 1 : index + 1 + size]):
                    span = size
            result.append(row)
            index += 1 + span
        return result

    previous_summary = first_data - 1
    previous_total = first_data - 1
    for summary_row in summary_rows:
        role = roles.get(summary_row)
        start = previous_total if role == ROW_TOTAL else max(previous_summary, previous_total)
        children = top_level([row for row in data_rows if start < row < summary_row])
        previous_summary = summary_row
        if role == ROW_TOTAL:
            previous_total = summary_row
        if len(children) < 2:
            continue
        matched = matched_columns(summary_row, children)
        if not matched:
            continue
        kind = RelationKind.GRAND_TOTAL_OF if role == ROW_TOTAL else RelationKind.PARENT_OF
        column_label = "، ".join(_display_header(item[0]) for item in matched)
        parent_label = _row_text(cells, labels, summary_row, label_column)
        child_labels = [_row_text(cells, labels, row, label_column) for row in children]
        letter = str(getattr(matched[0][0], "letter", "") or "")
        relations.append(
            CategoryRelation(
                relation_id=f"R{len(relations) + 1}",
                kind=kind,
                source=profile.source,
                sheet=profile.sheet,
                block=profile.block,
                parent_label=parent_label,
                child_labels=child_labels,
                column=column_label,
                parent_value=_decimal_text(matched[0][1]),
                child_sum=_decimal_text(matched[0][2]),
                difference=_decimal_text(matched[0][1] - matched[0][2]),
                confirmed=True,
                evidence=[
                    (
                        "مقدار ردیف والد با مجموع ردیف‌های جزئی برابر است "
                        f"({column_label}: {_decimal_text(matched[0][1])} = "
                        f"{_decimal_text(matched[0][2])})."
                    )
                ],
                addresses=[f"{letter}{summary_row}", *[f"{letter}{row}" for row in children]],
                parent_row=summary_row,
                child_rows=list(children),
            )
        )
        if len(relations) >= 60:
            return relations

    # A parent row that is not classified as a summary row can still be a
    # category parent: its value equals the sum of the following rows.  This is
    # reported as probable, never as proof, and nothing is removed.
    data_set = set(data_rows)
    used_parents = {relation.parent_row for relation in relations}
    for row in data_rows:
        if row in used_parents:
            continue
        run: list[int] = []
        cursor = row + 1
        while cursor in data_set and len(run) < 20:
            run.append(cursor)
            cursor += 1
        if len(run) < 2:
            continue
        matched: list[tuple[Any, Decimal, Decimal]] = []
        span = 0
        for size in range(2, len(run) + 1):
            if not is_parent(row, run[:size]):
                continue
            candidate = matched_columns(row, run[:size], strict=True)
            if candidate:
                matched, span = candidate, size
        if not matched:
            continue
        run = run[:span]
        parent_label = _row_text(cells, labels, row, label_column)
        child_labels = [_row_text(cells, labels, child, label_column) for child in run]
        parent_tokens = _label_tokens(parent_label)
        if not parent_tokens:
            continue
        column_label = "، ".join(_display_header(item[0]) for item in matched)
        letter = str(getattr(matched[0][0], "letter", "") or "")
        relations.append(
            CategoryRelation(
                relation_id=f"R{len(relations) + 1}",
                kind=RelationKind.PARENT_OF,
                source=profile.source,
                sheet=profile.sheet,
                block=profile.block,
                parent_label=parent_label,
                child_labels=child_labels,
                column=column_label,
                parent_value=_decimal_text(matched[0][1]),
                child_sum=_decimal_text(matched[0][2]),
                difference=_decimal_text(matched[0][1] - matched[0][2]),
                confirmed=False,
                evidence=[
                    (
                        "مقدار این ردیف با مجموع ردیف‌های بعدی برابر است؛ رابطه والد/فرزند "
                        "احتمالی است و باید بازبینی شود (حذف خودکار انجام نشد)."
                    )
                ],
                addresses=[f"{letter}{row}", *[f"{letter}{child}" for child in run]],
                parent_row=row,
                child_rows=list(run),
            )
        )
        if len(relations) >= 60:
            break
    return relations


def _renumber_relations(relations: list[CategoryRelation]) -> list[CategoryRelation]:
    parent_index = 0
    overlap_index = 0
    for relation in relations:
        if relation.kind is RelationKind.OVERLAP:
            overlap_index += 1
            relation.relation_id = f"O{overlap_index}"
        else:
            parent_index += 1
            relation.relation_id = f"R{parent_index}"
    return relations


def _periods_compatible(left: BlockProfile, right: BlockProfile) -> bool:
    """Two blocks in different periods cannot double-count the same records."""
    left_tokens = left.period_tokens[:1] or ([left.title] if left.title else [])
    right_tokens = right.period_tokens[:1] or ([right.title] if right.title else [])
    left_periods = parse_periods(left_tokens)
    right_periods = parse_periods(right_tokens)
    if not left_periods or not right_periods:
        return True
    relation = compare_period_sets(left_periods, right_periods).relation
    return relation in {"equivalent", "overlapping", "within", "contains"}


def _value_overlap(
    left: BlockProfile, right: BlockProfile
) -> tuple[float, str, list[str], int, int] | None:
    candidates: list[tuple[str, set[str], set[str]]] = []
    for left_header, left_values in left.key_values.items():
        for right_header, right_values in right.key_values.items():
            if _normalize_value(left_header) == _normalize_value(right_header):
                candidates.append(
                    (f"شناسه «{left_header}»", set(left_values), set(right_values))
                )
    for left_header, left_values in left.category_values.items():
        for right_header, right_values in right.category_values.items():
            if _normalize_value(left_header) == _normalize_value(right_header):
                candidates.append(
                    (f"دسته «{left_header}»", set(left_values), set(right_values))
                )
    if not candidates:
        left_keys = set().union(*(set(values) for values in left.key_values.values())) if left.key_values else set()
        right_keys = set().union(*(set(values) for values in right.key_values.values())) if right.key_values else set()
        if len(left_keys) >= 2 and len(right_keys) >= 2:
            candidates.append(("مجموع شناسه‌ها", left_keys, right_keys))
        else:
            left_categories = set().union(
                *(set(values) for values in left.category_values.values())
            ) if left.category_values else set()
            right_categories = set().union(
                *(set(values) for values in right.category_values.values())
            ) if right.category_values else set()
            if len(left_categories) >= 2 and len(right_categories) >= 2:
                candidates.append(("مجموع دسته‌ها", left_categories, right_categories))

    best: tuple[float, str, list[str], int, int] | None = None
    for basis, left_values, right_values in candidates:
        if len(left_values) < 2 or len(right_values) < 2:
            continue
        union = left_values | right_values
        if not union:
            continue
        shared = sorted(left_values & right_values)
        ratio = len(shared) / len(union)
        if best is None or ratio > best[0]:
            best = (ratio, basis, shared[:_MAX_SHARED_VALUES], len(left_values), len(right_values))
    return best


def _overlap_relations(
    profiles: list[BlockProfile], groups: list[DuplicateGroup]
) -> list[CategoryRelation]:
    membership = {
        member.identity: group.group_id
        for group in groups
        for member in group.members
    }
    relations: list[CategoryRelation] = []
    ordered = sorted(profiles, key=lambda item: item.identity)
    for index, left in enumerate(ordered):
        for right in ordered[index + 1 :]:
            if left.identity == right.identity:
                continue
            if not _periods_compatible(left, right):
                continue
            left_group = membership.get(left.identity)
            if left_group is not None and left_group == membership.get(right.identity):
                continue
            overlap = _value_overlap(left, right)
            if overlap is None:
                continue
            ratio, basis, shared, left_size, right_size = overlap
            if ratio <= 0:
                continue
            relations.append(
                CategoryRelation(
                    relation_id=f"O{len(relations) + 1}",
                    kind=RelationKind.OVERLAP,
                    source=left.source,
                    sheet=left.sheet,
                    block=left.block,
                    parent_label=left.label,
                    child_labels=[right.label],
                    column=basis,
                    confirmed=True,
                    evidence=[
                        (
                            f"{len(shared)} مقدار مشترک در {basis} وجود دارد؛ "
                            f"هم‌پوشانی {ratio:.0%} از {left_size} و {right_size} مقدار. "
                            f"نمونه: {'، '.join(shared)}"
                        )
                    ],
                    addresses=[left.address, right.address],
                )
            )
            if len(relations) >= _MAX_RELATIONS:
                return relations
    return relations


# --------------------------------------------------------------------------- #
# Worksheet classification
# --------------------------------------------------------------------------- #
def _group_primary(group: DuplicateGroup) -> BlockProfile:
    """One member of a duplicate group is the representation kept as source."""
    best = group.members[0]
    best_score = (not best.summary_only, len(best.summary_signatures), best.row_count)
    for member in group.members[1:]:
        score = (not member.summary_only, len(member.summary_signatures), member.row_count)
        if score > best_score:
            best, best_score = member, score
    return best


def _classify_sheets(
    analysis: Any,
    profiles: list[BlockProfile],
    groups: list[DuplicateGroup],
    source: str = "",
) -> list[SheetProfile]:
    extraction = getattr(analysis, "extraction", None)
    if extraction is None:
        return []
    source_name = source or str(getattr(extraction, "source_filename", "") or "")
    membership = {member.identity: group for group in groups for member in group.members}
    intelligence = getattr(analysis, "intelligence", None)
    sheets: list[SheetProfile] = []
    for sheet_model in getattr(extraction, "sheets", []) or []:
        name = str(getattr(sheet_model, "name", "") or "")
        blocks = [
            profile
            for profile in profiles
            if profile.sheet == name and profile.source == source_name
        ]
        intel = intelligence.sheet(name) if intelligence is not None else None
        title = str(getattr(intel, "title", "") or getattr(sheet_model, "title", "") or "")
        intel_refs = list(getattr(intel, "referenced_sheets", []) or [])
        formula_cells = max(
            sum(block.formula_cells for block in blocks),
            int(getattr(intel, "formula_cells", 0) or 0),
        )
        cross_sheet = max(
            sum(block.cross_sheet_formula_cells for block in blocks),
            int(getattr(intel, "cross_sheet_formula_cells", 0) or 0),
        )
        references = sorted(
            {*intel_refs, *(ref for block in blocks for ref in block.referenced_sheets)}
        )
        summary_rows = sum(len(block.summary_row_numbers) for block in blocks)
        data_blocks = sum(1 for block in blocks if block.row_signatures)
        summary_only = bool(blocks) and all(block.summary_only for block in blocks)
        labels = [
            name,
            title,
            *(block.title for block in blocks),
            *(
                header
                for block in blocks
                for header in [
                    *block.measure_headers,
                    *block.key_headers,
                    *block.category_headers,
                ]
            ),
        ]
        aggregation_label = any(
            label
            and (
                is_total_label(label)
                or summary_label_role(label)
                or any(word in _normalize_value(label) for word in _SUMMARY_TITLE_WORDS)
            )
            for label in labels
        )
        formula_based = (
            (cross_sheet >= 2 and cross_sheet / max(formula_cells, 1) >= 0.5)
            or (cross_sheet >= 1 and not blocks)
            or (
                bool(getattr(intel, "derived", False))
                and (summary_only or aggregation_label)
            )
        )
        sheet_label_summary = any(
            text and any(word in _normalize_value(text) for word in _SUMMARY_TITLE_WORDS)
            for text in (name, title)
        ) or is_total_label(name) or bool(summary_label_role(name))
        primaries = {id(group): _group_primary(group) for group in groups}
        duplicate_blocks = 0
        external: set[str] = set()
        duplicate_notes: list[str] = []
        for block in blocks:
            group = membership.get(block.identity)
            if group is None:
                continue
            primary = primaries.get(id(group))
            if primary is not None and primary.identity == block.identity:
                continue
            others = [
                member
                for member in group.members
                if member.identity != block.identity
                and (member.source != block.source or member.sheet != block.sheet)
            ]
            if others:
                duplicate_blocks += 1
                external.update(f"{member.source}/{member.sheet}" for member in others)
                duplicate_notes.append(
                    f"بلوک {block.block} با "
                    f"{'، '.join(f'{member.source}/{member.sheet}' for member in others)} "
                    + (
                        "یکسان است."
                        if group.identical
                        else f"در گروه {group.group_id} بازنمایی/جمع تکراری دارد."
                    )
                )
        evidence: list[str] = []
        if not blocks and not formula_based:
            sheets.append(
                SheetProfile(
                    source=source_name,
                    sheet=name,
                    title=title,
                    role=SheetRole.UNKNOWN,
                    confidence=0.3,
                    evidence=["هیچ بلوک داده‌ای در این برگه شناسایی نشد."],
                    referenced_sheets=references,
                )
            )
            continue
        if formula_based:
            if summary_rows or summary_only or aggregation_label:
                role = SheetRole.SUMMARY
                confidence = 0.9
                evidence.append(
                    f"{cross_sheet} از {formula_cells} سلول فرمول به برگه‌های دیگر ارجاع می‌دهد"
                    + (f" ({'، '.join(references)})" if references else "")
                    + "؛ برگه بازنمایی جمع‌شدهٔ داده‌های دیگر است."
                )
                if summary_rows:
                    evidence.append(f"{summary_rows} ردیف جمع/مقایسه/درصد در برگه وجود دارد.")
                if aggregation_label:
                    evidence.append("عنوان یا سرستون برگه واژه جمع/خلاصه دارد.")
                if not blocks:
                    evidence.append(
                        "هیچ بلوک داده مستقلی در برگه شناسایی نشد؛ فقط فرمول جمع/ارجاع دارد."
                    )
            else:
                role = SheetRole.DERIVED
                confidence = 0.85
                evidence.append(
                    f"{cross_sheet} از {formula_cells} سلول فرمول به برگه‌های دیگر ارجاع می‌دهد"
                    + (f" ({'، '.join(references)})" if references else "")
                    + "؛ برگه از داده‌های دیگر ساخته شده و منبع مستقل نیست."
                )
        elif summary_only:
            role = SheetRole.SUMMARY
            confidence = 0.85
            evidence.append(
                "برگه فقط ردیف‌های جمع دارد و ردیف داده مستقل ندارد؛ "
                "افزودن آن به جزئیات شمارش دوباره است."
            )
        elif duplicate_blocks and duplicate_blocks / len(blocks) >= 0.5:
            role = SheetRole.DUPLICATE
            confidence = 0.9
            evidence.extend(duplicate_notes)
        elif sheet_label_summary:
            role = SheetRole.SUMMARY
            confidence = 0.55
            evidence.append(
                "نام یا عنوان برگه واژه خلاصه/جمع دارد؛ شاهد فرمول یا تکرار محتوا "
                "ثبت نشد و این طبقه‌بندی احتمالی است."
            )
        elif data_blocks:
            role = SheetRole.PRIMARY
            confidence = 0.8
            evidence.append(
                f"برگه {data_blocks} بلوک داده مستقل دارد و فرمول ارجاع‌دهنده به "
                "برگه‌های دیگر در آن دیده نشد."
            )
        else:
            role = SheetRole.UNKNOWN
            confidence = 0.4
            evidence.append("بلوک‌های برگه ردیف داده کافی برای طبقه‌بندی ندارند.")
        if references and not formula_based:
            evidence.append(f"ارجاع فرمولی به {'، '.join(references)} ثبت شده است.")
        if summary_rows and role is not SheetRole.SUMMARY:
            evidence.append(
                f"{summary_rows} ردیف جمع در برگه هست؛ این ردیف‌ها از تجمیع کنار گذاشته می‌شوند."
            )
        sheets.append(
            SheetProfile(
                source=source_name,
                sheet=name,
                title=title,
                role=role,
                confidence=confidence,
                evidence=evidence,
                referenced_sheets=references,
                formula_cells=formula_cells,
                cross_sheet_formula_cells=cross_sheet,
                data_blocks=data_blocks,
                duplicate_of=sorted(external),
                summary_rows=summary_rows,
            )
        )
    return sheets


# --------------------------------------------------------------------------- #
# Safe aggregation checks
# --------------------------------------------------------------------------- #
def _profile_for_field(field: Any, profiles: list[BlockProfile]) -> BlockProfile | None:
    for profile in profiles:
        if (
            profile.source == getattr(field, "source", "")
            and profile.sheet == getattr(field, "sheet", "")
            and getattr(field, "block", "") in {profile.block, profile.address}
        ):
            return profile
    return None


def _sheet_profile(
    field: Any, sheets: dict[tuple[str, str], SheetProfile]
) -> SheetProfile | None:
    return sheets.get((getattr(field, "source", ""), getattr(field, "sheet", "")))


def _indicator_check(match: Any) -> AggregationCheck:
    if match.status.value == "matched":
        return AggregationCheck(
            "indicator_equivalence",
            CheckStatus.PASS,
            "برچسب نرمال‌شده و نقش معنایی دو فیلد یکسان است.",
        )
    if match.status.value == "probable":
        return AggregationCheck(
            "indicator_equivalence",
            CheckStatus.REVIEW,
            "برچسب‌ها هم‌معنا هستند اما برابری کامل اثبات نشد (دوره/واحد/برچسب ناقص).",
        )
    return AggregationCheck(
        "indicator_equivalence",
        CheckStatus.REVIEW,
        "چند کاندید نزدیک وجود دارد و شاخص یکی از آن‌ها باید تأیید شود.",
    )


def _period_check(match: Any) -> AggregationCheck:
    if getattr(match, "period_status", "") == "conflicting":
        return AggregationCheck(
            "period_compatibility",
            CheckStatus.REVIEW,
            "یک فیلد دو دوره ناسازگار دارد؛ دوره معتبر باید تأیید شود.",
        )
    relation = getattr(match, "period_relation", "") or ""
    comparable = bool(getattr(match, "comparable", False))
    if comparable and relation == "equivalent":
        return AggregationCheck(
            "period_compatibility",
            CheckStatus.PASS,
            "دوره گزارش‌دهی هر دو فیلد یکسان است.",
        )
    if relation in {"within", "contains", "overlapping"}:
        return AggregationCheck(
            "period_compatibility",
            CheckStatus.BLOCK,
            "دوره‌ها هم‌پوشانی دارند؛ جمع آن‌ها بخشی از داده را دو بار می‌شمارد.",
        )
    if relation == "incompatible":
        return AggregationCheck(
            "period_compatibility",
            CheckStatus.BLOCK,
            "دوره‌ها یکسان نیستند؛ جمع آن‌ها فقط با تفکیک دوره معنا دارد.",
        )
    return AggregationCheck(
        "period_compatibility",
        CheckStatus.REVIEW,
        "دوره گزارش‌دهی کامل یا قابل مقایسه اثبات نشد.",
    )


def _unit_check(match: Any) -> AggregationCheck:
    state = getattr(match, "unit_state", "") or ""
    if state in {"equal", "both_missing"}:
        return AggregationCheck(
            "unit_compatibility",
            CheckStatus.PASS,
            "واحد اندازه‌گیری یکسان است یا در هر دو ثبت نشده و تفاوتی ندارد.",
        )
    if state in {"scalable", "same_dimension"}:
        return AggregationCheck(
            "unit_compatibility",
            CheckStatus.REVIEW,
            "واحد هم‌جنس است اما مقیاس/واحد یکسان نیست؛ پیش از جمع باید تبدیل ثبت شود.",
        )
    return AggregationCheck(
        "unit_compatibility",
        CheckStatus.BLOCK,
        "واحدها ناسازگار یا ترکیبی‌اند و جمع آن‌ها معنا ندارد.",
    )


def _category_check(match: Any, relations: list[CategoryRelation]) -> AggregationCheck:
    members = list(getattr(match, "members", []) or [])
    for index, member in enumerate(members):
        tokens = _label_tokens(
            getattr(member, "normalized_label", "") or getattr(member, "header", "")
        )
        for other in members[index + 1 :]:
            other_tokens = _label_tokens(
                getattr(other, "normalized_label", "") or getattr(other, "header", "")
            )
            if (tokens & _AGGREGATION_TOKENS) != (other_tokens & _AGGREGATION_TOKENS):
                return AggregationCheck(
                    "category_relationship",
                    CheckStatus.BLOCK,
                    "یکی از دو فیلد جمع/میانگین/درصد است و دیگری مقدار پایه؛ "
                    "جمع آن‌ها شمارش دوباره است.",
                )
    for relation in relations:
        if not relation.confirmed:
            continue
        parent_tokens = _label_tokens(relation.parent_label)
        child_tokens: set[str] = set()
        for child in relation.child_labels:
            child_tokens |= _label_tokens(child)
        for index, member in enumerate(members):
            tokens = _label_tokens(getattr(member, "header", ""))
            if not tokens & parent_tokens:
                continue
            for other in members[index + 1 :]:
                other_tokens = _label_tokens(getattr(other, "header", ""))
                if other_tokens & child_tokens:
                    return AggregationCheck(
                        "category_relationship",
                        CheckStatus.BLOCK,
                        "یکی از فیلدها جمع والد و دیگری جزء فرزند است؛ "
                        "جمع آن‌ها شمارش دوباره است.",
                    )
    return AggregationCheck(
        "category_relationship",
        CheckStatus.PASS,
        "هیچ رابطه والد/فرزند یا جمع/جزء بین دو فیلد شناسایی نشد.",
    )


def _independence_check(
    left: Any,
    right: Any,
    profiles: list[BlockProfile],
    sheets: dict[tuple[str, str], SheetProfile],
    groups: list[DuplicateGroup],
) -> AggregationCheck:
    left_profile = _profile_for_field(left, profiles)
    right_profile = _profile_for_field(right, profiles)
    if left_profile is None or right_profile is None:
        return AggregationCheck(
            "source_independence",
            CheckStatus.REVIEW,
            "بلوک متناظر با فیلد پیدا نشد؛ استقلال منبع اثبات نشد.",
        )
    if left_profile.identity == right_profile.identity:
        return AggregationCheck(
            "source_independence",
            CheckStatus.BLOCK,
            "هر دو فیلد در یک بلوک‌اند؛ جمع آن‌ها همان داده را دو بار می‌شمارد.",
        )
    membership = {
        member.identity: group.group_id
        for group in groups
        for member in group.members
    }
    left_group = membership.get(left_profile.identity)
    if left_group is not None and left_group == membership.get(right_profile.identity):
        return AggregationCheck(
            "source_independence",
            CheckStatus.BLOCK,
            f"هر دو بلوک عضو یک گروه محتوای تکراری‌اند ({left_group}).",
        )
    left_sheet = _sheet_profile(left, sheets)
    right_sheet = _sheet_profile(right, sheets)
    if left_sheet is None or right_sheet is None:
        return AggregationCheck(
            "source_independence",
            CheckStatus.REVIEW,
            "طبقه‌بندی برگه در دسترس نیست؛ استقلال منبع اثبات نشد.",
        )
    if left_profile.source == right_profile.source:
        for source_sheet, target_sheet in (
            (left_sheet, right_sheet),
            (right_sheet, left_sheet),
        ):
            if source_sheet.role in {SheetRole.SUMMARY, SheetRole.DERIVED}:
                if target_sheet.sheet in source_sheet.referenced_sheets:
                    return AggregationCheck(
                        "source_independence",
                        CheckStatus.BLOCK,
                        f"برگه «{source_sheet.sheet}» ({source_sheet.role_label}) به "
                        f"برگه «{target_sheet.sheet}» ارجاع می‌دهد و منبع مستقل نیست.",
                    )
                return AggregationCheck(
                    "source_independence",
                    CheckStatus.REVIEW,
                    f"برگه «{source_sheet.sheet}» {source_sheet.role_label} است؛ "
                    "استقلال آن از برگه دیگر اثبات نشد.",
                )
            if source_sheet.role is SheetRole.DUPLICATE:
                return AggregationCheck(
                    "source_independence",
                    CheckStatus.BLOCK,
                    f"برگه «{source_sheet.sheet}» تکراری است.",
                )
        if left_sheet.role is SheetRole.UNKNOWN or right_sheet.role is SheetRole.UNKNOWN:
            return AggregationCheck(
                "source_independence",
                CheckStatus.REVIEW,
                "نقش یکی از برگه‌ها نامشخص است.",
            )
        return AggregationCheck(
            "source_independence",
            CheckStatus.PASS,
            "دو برگه مستقل با داده اولیه و بدون ارجاع فرمولی متقابل‌اند.",
        )
    if left_sheet.role is not SheetRole.PRIMARY or right_sheet.role is not SheetRole.PRIMARY:
        return AggregationCheck(
            "source_independence",
            CheckStatus.REVIEW,
            "یکی از فایل‌ها برگه خلاصه/مشتق/تکراری دارد؛ استقلال کامل اثبات نشد.",
        )
    return AggregationCheck(
        "source_independence",
        CheckStatus.PASS,
        "فایل‌های متفاوت با محتوای غیرتکراری‌اند.",
    )


def _coverage_check(
    left: Any,
    right: Any,
    profiles: list[BlockProfile],
    *,
    period_comparable: bool = True,
) -> AggregationCheck:
    left_profile = _profile_for_field(left, profiles)
    right_profile = _profile_for_field(right, profiles)
    if left_profile is None or right_profile is None:
        return AggregationCheck(
            "coverage_non_overlap",
            CheckStatus.REVIEW,
            "پوشش رکوردها قابل بررسی نیست.",
        )
    if not period_comparable:
        return AggregationCheck(
            "coverage_non_overlap",
            CheckStatus.REVIEW,
            "دوره دو فیلد یکسان نیست؛ پوشش رکوردها قابل مقایسه نیست.",
        )
    overlap = _value_overlap(left_profile, right_profile)
    if overlap is None:
        return AggregationCheck(
            "coverage_non_overlap",
            CheckStatus.REVIEW,
            "شناسه/دسته کافی برای اثبات جدا بودن رکوردها وجود ندارد.",
        )
    ratio, basis, shared, left_size, right_size = overlap
    if ratio == 0:
        return AggregationCheck(
            "coverage_non_overlap",
            CheckStatus.PASS,
            f"هیچ مقدار مشترکی در {basis} نیست ({left_size} و {right_size} مقدار).",
        )
    return AggregationCheck(
        "coverage_non_overlap",
        CheckStatus.BLOCK,
        f"{len(shared)} مقدار مشترک در {basis} وجود دارد؛ هم‌پوشانی {ratio:.0%} "
        "یعنی بخشی از رکوردها دو بار شمرده می‌شود.",
    )


# --------------------------------------------------------------------------- #
# Decisions and exclusions
# --------------------------------------------------------------------------- #
def _member_label(field: Any) -> str:
    return (
        f"{getattr(field, 'source', '')}/{getattr(field, 'sheet', '')}/"
        f"{getattr(field, 'block', '')}/«{getattr(field, 'header', '')}»"
    )


def _build_decisions(
    matching: Any,
    profiles: list[BlockProfile],
    sheets: dict[tuple[str, str], SheetProfile],
    groups: list[DuplicateGroup],
    relations: list[CategoryRelation],
) -> list[AggregationDecision]:
    decisions: list[AggregationDecision] = []
    items: list[Any] = []
    if matching is not None:
        items.extend(getattr(matching, "matches", []) or [])
        items.extend(getattr(matching, "probable", []) or [])
        items.extend(getattr(matching, "ambiguities", []) or [])
    for match in items:
        if getattr(match, "kind", "") != "measure":
            continue
        members = list(getattr(match, "members", []) or [])
        if getattr(match, "status", None) is not None and match.status.value == "ambiguous":
            members = [*members, *(candidate.field for candidate in match.candidates)]
        if len(members) < 2:
            continue
        indicator = _indicator_check(match)
        period = _period_check(match)
        unit = _unit_check(match)
        category = _category_check(match, relations)
        metric = (
            getattr(match, "normalized_label", "")
            or "، ".join(dict.fromkeys(getattr(member, "header", "") for member in members))
        )
        period_comparable = bool(getattr(match, "comparable", False))
        for index, left in enumerate(members):
            for right in members[index + 1 :]:
                independence = _independence_check(left, right, profiles, sheets, groups)
                coverage = _coverage_check(
                    left, right, profiles, period_comparable=period_comparable
                )
                checks = [indicator, period, unit, category, independence, coverage]
                if any(check.status is CheckStatus.BLOCK for check in checks):
                    status = AggregationStatus.BLOCKED
                elif any(check.status is CheckStatus.REVIEW for check in checks):
                    status = AggregationStatus.REVIEW
                else:
                    status = AggregationStatus.ALLOWED
                reasons = [
                    f"{check.label}: {check.explanation}"
                    for check in checks
                    if check.status is not CheckStatus.PASS
                ]
                decisions.append(
                    AggregationDecision(
                        decision_id=f"A{len(decisions) + 1}",
                        metric=metric or "—",
                        kind=str(getattr(match, "kind", "")),
                        members=[_member_label(left), _member_label(right)],
                        status=status,
                        checks=checks,
                        reasons=reasons,
                        evidence=list(getattr(match, "evidence", []) or [])[:6],
                        period_relation=str(getattr(match, "period_relation", "") or ""),
                        unit_state=str(getattr(match, "unit_state", "") or ""),
                    )
                )
                if len(decisions) >= _MAX_DECISIONS:
                    return decisions

    for group in groups:
        if len(group.members) < 2:
            continue
        summary_group = group.kind == "summary"
        checks = [
            AggregationCheck(
                "indicator_equivalence",
                CheckStatus.PASS,
                (
                    "اثر انگشت ردیف‌های جمع یکسان است؛ شاخص هم یکی است."
                    if summary_group
                    else "محتوای بلوک‌ها یکسان است؛ بنابراین شاخص هم یکی است."
                ),
            ),
            AggregationCheck(
                "period_compatibility",
                CheckStatus.PASS,
                (
                    "جمع‌های یکسان به همان دوره تعلق دارند."
                    if summary_group
                    else "چون محتوا یکسان است، دوره هم همان دوره است."
                ),
            ),
            AggregationCheck(
                "unit_compatibility",
                CheckStatus.PASS,
                "واحد در محتوای یکسان تفاوتی ندارد.",
            ),
            AggregationCheck(
                "category_relationship",
                CheckStatus.PASS,
                "رابطه والد/فرزند بین دو بازنمایی یکسان معنا ندارد.",
            ),
            AggregationCheck(
                "source_independence",
                CheckStatus.BLOCK,
                (
                    "همان ردیف‌های جمع در چند برگه/فایل تکرار شده‌اند؛ منبع مستقل نیست."
                    if summary_group
                    else "همان محتوا در چند برگه/فایل تکرار شده است؛ منبع مستقل نیست."
                ),
            ),
            AggregationCheck(
                "coverage_non_overlap",
                CheckStatus.BLOCK,
                "همان رکوردها پوشش داده شده‌اند؛ جمع آن‌ها شمارش دوباره است.",
            ),
        ]
        decisions.append(
            AggregationDecision(
                decision_id=f"A{len(decisions) + 1}",
                metric=group.members[0].title or group.members[0].label,
                kind="duplicate_block",
                members=[member.label for member in group.members],
                status=AggregationStatus.BLOCKED,
                checks=checks,
                reasons=[check.explanation for check in checks if check.status is CheckStatus.BLOCK],
                evidence=list(group.evidence),
            )
        )
        if len(decisions) >= _MAX_DECISIONS:
            break
    return decisions


def _summary_row_kind(block: Any, row: int) -> str:
    for record in getattr(block, "row_roles", []) or []:
        if not isinstance(record, dict) or int(record.get("row") or 0) != row:
            continue
        role = str(record.get("role") or "")
        return {
            ROW_SUBTOTAL: "subtotal_row",
            ROW_TOTAL: "grand_total_row",
            ROW_DIFFERENCE: "difference_row",
            ROW_PERCENTAGE: "percentage_row",
        }.get(role, "summary_row")
    return "summary_row"


def _exclusions(
    profiles: list[BlockProfile],
    sheets: list[SheetProfile],
    groups: list[DuplicateGroup],
    record_duplicates: list[RecordDuplicate],
    relations: list[CategoryRelation],
    decisions: list[AggregationDecision],
    blocks_by_identity: dict[str, Any],
) -> list[ExclusionRecord]:
    exclusions: list[ExclusionRecord] = []
    summary_rows_by_identity = {
        profile.identity: set(profile.summary_row_numbers) for profile in profiles
    }
    for sheet in sheets:
        if sheet.role in {SheetRole.PRIMARY, SheetRole.UNKNOWN}:
            continue
        exclusions.append(
            ExclusionRecord(
                source=sheet.source,
                sheet=sheet.sheet,
                block="",
                address="",
                kind=f"sheet_{sheet.role.value}",
                reason=(
                    f"برگه «{sheet.sheet}» {sheet.role_label} است و به‌عنوان منبع مستقل "
                    "در تجمیع شمرده نمی‌شود."
                ),
                evidence=list(sheet.evidence),
            )
        )
    for group in groups:
        if len(group.members) < 2:
            continue
        primary = _group_primary(group)
        for member in group.members:
            if member.identity == primary.identity:
                continue
            exclusions.append(
                ExclusionRecord(
                    source=member.source,
                    sheet=member.sheet,
                    block=member.block,
                    address=member.address,
                    kind="duplicate_table",
                    reason=group.reason,
                    evidence=[
                        f"گروه {group.group_id}؛ بازنمایی مبنا: {primary.label}",
                        *group.evidence,
                    ],
                )
            )
    for item in record_duplicates:
        exclusions.append(
            ExclusionRecord(
                source=item.source,
                sheet=item.sheet,
                block=item.block,
                address=f"ردیف‌های {item.row_numbers}",
                kind="record_duplicate" if item.kind == "identical" else "conflicting_record",
                reason=(
                    "همان کلید و همان مقدار چند بار تکرار شده است؛ فقط یک رکورد شمرده می‌شود."
                    if item.kind == "identical"
                    else "کلید یکسان با مقدار متفاوت؛ تا بازبینی کاربر تجمیع نمی‌شود."
                ),
                evidence=list(item.evidence),
            )
        )
    for relation in relations:
        if relation.kind is RelationKind.OVERLAP:
            exclusions.append(
                ExclusionRecord(
                    source=relation.source,
                    sheet=relation.sheet,
                    block=relation.block,
                    address="، ".join(relation.addresses),
                    kind="overlapping_group",
                    reason=(
                        "دو گروه/برگه بخشی از رکوردهای مشترک دارند؛ جمع آن‌ها آن بخش را "
                        "دو بار می‌شمارد."
                    ),
                    evidence=list(relation.evidence),
                )
            )
            continue
        if not relation.confirmed:
            continue
        identity = f"{relation.source}::{relation.sheet}::{relation.block}"
        if relation.parent_row in summary_rows_by_identity.get(identity, set()):
            continue
        exclusions.append(
            ExclusionRecord(
                source=relation.source,
                sheet=relation.sheet,
                block=relation.block,
                address=relation.addresses[0] if relation.addresses else "",
                kind=(
                    "grand_total_row"
                    if relation.kind is RelationKind.GRAND_TOTAL_OF
                    else "parent_category"
                ),
                reason=(
                    "این ردیف جمع ردیف‌های جزئی است و نباید دوباره با آن‌ها جمع شود."
                ),
                evidence=list(relation.evidence),
            )
        )
    for profile in profiles:
        block = blocks_by_identity.get(profile.identity)
        for row in profile.summary_row_numbers:
            kind = _summary_row_kind(block, row) if block is not None else "summary_row"
            exclusions.append(
                ExclusionRecord(
                    source=profile.source,
                    sheet=profile.sheet,
                    block=profile.block,
                    address=f"ردیف {row}",
                    kind=kind,
                    reason=(
                        "ردیف جمع/میانگین/درصد/اختلاف است و با ردیف‌های داده جمع نمی‌شود."
                    ),
                    evidence=["row_classification این ردیف را در نقش‌های خلاصه قرار داده است."],
                )
            )
    for decision in decisions:
        if decision.status is not AggregationStatus.REVIEW:
            continue
        if not any("کاندید" in reason for reason in decision.reasons):
            continue
        exclusions.append(
            ExclusionRecord(
                source="",
                sheet="",
                block="",
                address="",
                kind="ambiguous_match",
                reason="تطبیق مبهم است؛ تا تأیید کاربر هیچ تجمیعی انجام نمی‌شود.",
                evidence=[decision.metric, *decision.members],
            )
        )
    unique: dict[tuple[str, str, str, str, str], ExclusionRecord] = {}
    for item in exclusions:
        key = (item.kind, item.source, item.sheet, item.block, item.address)
        unique.setdefault(key, item)
    return list(unique.values())[:_MAX_EXCLUSIONS]


def _warnings(
    report: DuplicateControlReport,
    profiles: list[BlockProfile],
    profile_total: int | None = None,
) -> list[str]:
    warnings: list[str] = []
    roles = report.role_counts()
    non_primary = roles.get(SheetRole.SUMMARY.value, 0) + roles.get(
        SheetRole.DERIVED.value, 0
    ) + roles.get(SheetRole.DUPLICATE.value, 0)
    if non_primary:
        warnings.append(
            f"{non_primary} برگه خلاصه/مشتق/تکراری شناسایی شد و به‌عنوان منبع مستقل "
            "در تجمیع شمرده نشد."
        )
    if report.duplicate_groups:
        warnings.append(
            f"{len(report.duplicate_groups)} گروه محتوای تکراری شناسایی شد؛ "
            "هیچ‌کدام حذف نشد و جمع آن‌ها مسدود است."
        )
    if report.record_duplicates:
        warnings.append(
            f"{len(report.record_duplicates)} کلید تکراری در سطح رکورد ثبت شد؛ "
            "مقادیر بازبینی شوند و به‌طور خودکار حذف نشدند."
        )
    if report.blocked_decisions:
        warnings.append(
            f"{len(report.blocked_decisions)} تصمیم تجمیع به دلیل شمارش دوباره مسدود شد."
        )
    if report.review_decisions:
        warnings.append(
            f"{len(report.review_decisions)} تصمیم تجمیع نیازمند بررسی است و "
            "به‌طور خودکار مجاز نشد."
        )
    if any(profile.truncated for profile in profiles):
        warnings.append(
            "مدل استخراج برخی بلوک‌ها به سقف سلول رسیده است؛ اثر انگشت تکرار بر "
            "همان سلول‌های موجود ساخته شد."
        )
    total = profile_total if profile_total is not None else len(profiles)
    if total > _MAX_PROFILES:
        warnings.append(
            f"تعداد بلوک‌ها ({total}) از سقف تحلیل تکرار ({_MAX_PROFILES}) "
            "بیشتر است؛ نتایج تکرار برای بلوک‌های اول محدود شد."
        )
    return warnings


# --------------------------------------------------------------------------- #
# Public entry points
# --------------------------------------------------------------------------- #
def _blocks_by_identity(analysis: Any, source: str) -> dict[str, Any]:
    extraction = getattr(analysis, "extraction", None)
    blocks: dict[str, Any] = {}
    for sheet in getattr(extraction, "sheets", []) or []:
        for block in getattr(sheet, "blocks", []) or []:
            identity = (
                f"{source}::{getattr(block, 'sheet', '')}::"
                f"{getattr(block, 'block_id', '')}"
            )
            blocks[identity] = block
    return blocks


def analyze_analysis(analysis: Any) -> DuplicateControlReport:
    """Build the duplicate/derived/aggregation report for one analysed file."""
    extraction = getattr(analysis, "extraction", None)
    source = str(getattr(extraction, "source_filename", "") or "")
    profiles = profiles_from_analysis(analysis, source)
    profile_total = len(profiles)
    if profile_total > _MAX_PROFILES:
        profiles = profiles[:_MAX_PROFILES]
    blocks = _blocks_by_identity(analysis, source)
    groups = _duplicate_groups(profiles)
    sheets = _classify_sheets(analysis, profiles, groups, source)
    record_duplicates: list[RecordDuplicate] = []
    for profile in profiles:
        record_duplicates.extend(_record_duplicates(profile))
    relations: list[CategoryRelation] = []
    for profile in profiles:
        block = blocks.get(profile.identity)
        if block is not None:
            relations.extend(_category_relations(profile, block))
    relations.extend(_overlap_relations(profiles, groups))
    relations = _renumber_relations(relations[:_MAX_RELATIONS])
    sheet_map = {(sheet.source, sheet.sheet): sheet for sheet in sheets}
    decisions = _build_decisions(
        getattr(analysis, "semantic_matching", None),
        profiles,
        sheet_map,
        groups,
        relations,
    )
    exclusions = _exclusions(
        profiles,
        sheets,
        groups,
        record_duplicates,
        relations,
        decisions,
        blocks,
    )
    report = DuplicateControlReport(
        source_count=1 if extraction is not None else 0,
        sheets=sheets,
        blocks=profiles,
        duplicate_groups=groups,
        record_duplicates=record_duplicates[:_MAX_RECORD_GROUPS],
        category_relations=relations,
        decisions=decisions,
        exclusions=exclusions,
    )
    report.warnings = _warnings(report, profiles, profile_total)
    return report


def analyze_sources(
    settings: Settings,
    analyses: Iterable[tuple[str, Any]],
    *,
    matching: Any = None,
) -> DuplicateControlReport:
    """Build one duplicate/aggregation report across several files."""
    items = [(str(name), analysis) for name, analysis in analyses]
    all_profiles: list[BlockProfile] = []
    blocks: dict[str, Any] = {}
    for name, analysis in items:
        extraction = getattr(analysis, "extraction", None)
        source = name or str(getattr(extraction, "source_filename", "") or "")
        all_profiles.extend(profiles_from_analysis(analysis, source))
        blocks.update(_blocks_by_identity(analysis, source))
    profile_total = len(all_profiles)
    if profile_total > _MAX_PROFILES:
        all_profiles = all_profiles[:_MAX_PROFILES]
    groups = _duplicate_groups(all_profiles)
    sheets: list[SheetProfile] = []
    for name, analysis in items:
        extraction = getattr(analysis, "extraction", None)
        source = name or str(getattr(extraction, "source_filename", "") or "")
        sheets.extend(_classify_sheets(analysis, all_profiles, groups, source))
    record_duplicates: list[RecordDuplicate] = []
    for profile in all_profiles:
        record_duplicates.extend(_record_duplicates(profile))
    relations: list[CategoryRelation] = []
    for profile in all_profiles:
        block = blocks.get(profile.identity)
        if block is not None:
            relations.extend(_category_relations(profile, block))
    relations.extend(_overlap_relations(all_profiles, groups))
    relations = _renumber_relations(relations[:_MAX_RELATIONS])
    if matching is None and items:
        from local_office_ai.excel.semantic_matching import match_across_sources

        matching = match_across_sources(settings, items)
    sheet_map = {(sheet.source, sheet.sheet): sheet for sheet in sheets}
    decisions = _build_decisions(matching, all_profiles, sheet_map, groups, relations)
    exclusions = _exclusions(
        all_profiles,
        sheets,
        groups,
        record_duplicates,
        relations,
        decisions,
        blocks,
    )
    report = DuplicateControlReport(
        source_count=len(items),
        sheets=sheets,
        blocks=all_profiles,
        duplicate_groups=groups,
        record_duplicates=record_duplicates[:_MAX_RECORD_GROUPS],
        category_relations=relations,
        decisions=decisions,
        exclusions=exclusions,
    )
    report.warnings = _warnings(report, all_profiles, profile_total)
    return report


class DuplicateControlAnalyzer:
    """Settings-bound facade used by the workbook and folder workflows."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def analyze(self, analysis: Any) -> DuplicateControlReport:
        return analyze_analysis(analysis)

    def analyze_sources(
        self,
        analyses: Iterable[tuple[str, Any]],
        *,
        matching: Any = None,
    ) -> DuplicateControlReport:
        return analyze_sources(self.settings, analyses, matching=matching)
