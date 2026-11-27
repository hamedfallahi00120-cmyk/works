"""Shared data model for the Excel intelligence pipeline.

The Excel reader, the analyzers, the AI layer and the report writer all speak in
these types.  Money is carried as ``Decimal`` so no binary float ever touches a
financial figure.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from typing import Any, Iterable


class Severity(str, Enum):
    INFO = "info"
    WARNING = "warning"
    CRITICAL = "critical"


class TableKind(str, Enum):
    LOAN = "loan"
    OBLIGATION = "obligation"
    TIME_SERIES = "time_series"
    KPI = "kpi"
    GENERIC = "generic"


class ValueKind(str, Enum):
    NUMBER = "number"
    DATE = "date"
    TEXT = "text"
    BOOLEAN = "boolean"
    EMPTY = "empty"
    ERROR = "error"


class AuditStatus(str, Enum):
    VERIFIED = "verified"
    MISMATCH = "mismatch"
    NOT_VERIFIABLE = "not_verifiable"
    NO_DECLARED_VALUE = "no_declared_value"


class MatchStatus(str, Enum):
    MATCH = "match"
    PROBABLE = "probable"
    NO_MATCH = "no_match"
    OUTSIDE_LOAN_TABLE = "outside_loan_table"
    NO_OBLIGATION = "no_obligation"


class ColumnRole(str, Enum):
    INSTITUTION = "institution"
    PRINCIPAL = "principal"
    INSTALLMENT = "installment"
    INSTALLMENT_COUNT = "installment_count"
    PAID = "paid"
    DATE = "date"
    DESCRIPTION = "description"
    AMOUNT = "amount"
    COUNT = "count"
    ID = "id"
    UNKNOWN = "unknown"


class DataRole(str, Enum):
    """Statistical role of a column, independent of business vocabulary."""

    IDENTIFIER = "identifier"
    TEMPORAL_KEY = "temporal_key"
    NUMERIC_MEASURE = "numeric_measure"
    CATEGORICAL = "categorical"
    GROUP = "group"
    TEXT = "text"
    CALCULATED = "calculated"
    UNKNOWN = "unknown"


@dataclass
class Provenance:
    """Where a value came from, down to the cell."""

    sheet: str
    address: str
    row: int
    column: int
    raw: Any = None
    formula: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "address": self.address,
            "row": self.row,
            "column": self.column,
            "raw": _jsonable(self.raw),
            "formula": self.formula,
        }


@dataclass
class NormalizedValue:
    """A raw cell value plus its standardized interpretation.

    ``raw`` is always kept; ``value`` holds the standardized form.  Doubtful
    conversions add an entry to ``warnings`` instead of being applied silently.
    """

    raw: Any
    value: Any
    kind: ValueKind
    method: str = "identity"
    warnings: list[str] = field(default_factory=list)
    provenance: Provenance | None = None

    @property
    def ok(self) -> bool:
        return self.kind is not ValueKind.ERROR and self.value is not None

    def to_dict(self) -> dict[str, Any]:
        return {
            "raw": _jsonable(self.raw),
            "value": _jsonable(self.value),
            "kind": self.kind.value,
            "method": self.method,
            "warnings": list(self.warnings),
            "provenance": self.provenance.to_dict() if self.provenance else None,
        }


@dataclass
class ColumnProfile:
    index: int
    letter: str
    header: str
    header_address: str
    data_type: ValueKind
    role: ColumnRole = ColumnRole.UNKNOWN
    numeric_ratio: float = 0.0
    date_ratio: float = 0.0
    text_ratio: float = 0.0
    empty_ratio: float = 0.0
    formula_ratio: float = 0.0
    non_empty: int = 0
    distinct: int = 0
    samples: list[str] = field(default_factory=list)
    confidence: float = 0.0
    # Hierarchical header information.  These fields stay empty for a flat
    # single-row header, so every existing consumer keeps working unchanged.
    original_header: str = ""
    parent_header: str = ""
    sub_header: str = ""
    period: str = ""
    source_address: str = ""
    statistical_role: DataRole = DataRole.UNKNOWN
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "letter": self.letter,
            "header": self.header,
            "header_address": self.header_address,
            "data_type": self.data_type.value,
            "role": self.role.value,
            "numeric_ratio": round(self.numeric_ratio, 4),
            "date_ratio": round(self.date_ratio, 4),
            "text_ratio": round(self.text_ratio, 4),
            "empty_ratio": round(self.empty_ratio, 4),
            "formula_ratio": round(self.formula_ratio, 4),
            "non_empty": self.non_empty,
            "distinct": self.distinct,
            "samples": list(self.samples),
            "confidence": round(self.confidence, 4),
            "original_header": self.original_header,
            "parent_header": self.parent_header,
            "sub_header": self.sub_header,
            "period": self.period,
            "source_address": self.source_address,
            "statistical_role": self.statistical_role.value,
            "notes": list(self.notes),
        }


@dataclass
class TableCandidate:
    """A rectangular table detected inside a sheet."""

    sheet: str
    header_row: int
    first_data_row: int
    last_data_row: int
    first_col: int
    last_col: int
    columns: list[ColumnProfile] = field(default_factory=list)
    kind: TableKind = TableKind.GENERIC
    confidence: float = 0.0
    reasons: list[str] = field(default_factory=list)
    header_values: list[str] = field(default_factory=list)
    title: str = ""
    # Extra structure information.  Empty lists keep flat tables unchanged.
    header_rows: list[int] = field(default_factory=list)
    header_group_count: int = 0
    data_rows_with_values: int = 0
    notes: list[str] = field(default_factory=list)
    # The block may continue below ``last_data_row`` with a grand total and
    # comparison/percentage rows.  Those rows are structural, never data.
    block_end_row: int = 0
    subtotal_rows: list[int] = field(default_factory=list)
    grand_total_rows: list[int] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return max(0, self.last_data_row - self.first_data_row + 1)

    @property
    def block_end(self) -> int:
        """Last row owned by this block, including its summary rows."""
        return max(self.block_end_row or self.last_data_row, self.last_data_row)

    @property
    def address(self) -> str:
        from openpyxl.utils import get_column_letter

        return (
            f"{get_column_letter(self.first_col)}{self.header_row}:"
            f"{get_column_letter(self.last_col)}{self.last_data_row}"
        )

    @property
    def block_address(self) -> str:
        from openpyxl.utils import get_column_letter

        return (
            f"{get_column_letter(self.first_col)}{self.header_row}:"
            f"{get_column_letter(self.last_col)}{self.block_end}"
        )

    def column_by_role(self, role: ColumnRole) -> ColumnProfile | None:
        for column in self.columns:
            if column.role is role:
                return column
        return None

    def columns_by_role(self, role: ColumnRole) -> list[ColumnProfile]:
        return [column for column in self.columns if column.role is role]

    def has_any_role(self, roles: Iterable[ColumnRole]) -> bool:
        wanted = set(roles)
        return any(column.role in wanted for column in self.columns)

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "kind": self.kind.value,
            "address": self.address,
            "header_row": self.header_row,
            "first_data_row": self.first_data_row,
            "last_data_row": self.last_data_row,
            "first_col": self.first_col,
            "last_col": self.last_col,
            "row_count": self.row_count,
            "confidence": round(self.confidence, 4),
            "reasons": list(self.reasons),
            "title": self.title,
            "header_rows": list(self.header_rows),
            "header_group_count": self.header_group_count,
            "data_rows_with_values": self.data_rows_with_values,
            "notes": list(self.notes),
            "block_end_row": self.block_end,
            "subtotal_rows": list(self.subtotal_rows),
            "grand_total_rows": list(self.grand_total_rows),
            "columns": [column.to_dict() for column in self.columns],
        }


@dataclass
class LoanRecord:
    institution_raw: str
    institution_key: str
    principal: Decimal | None
    installment: Decimal | None
    installment_count: Decimal | None
    last_installment_date: date | None
    contract_repayment: Decimal | None
    provenance: Provenance
    fields: dict[str, NormalizedValue] = field(default_factory=dict)
    issues: list[str] = field(default_factory=list)
    paid_installments: Decimal | None = None

    @property
    def remaining_installments(self) -> Decimal | None:
        if self.installment_count is None or self.paid_installments is None:
            return None
        return self.installment_count - self.paid_installments

    @property
    def remaining_amount(self) -> Decimal | None:
        remaining = self.remaining_installments
        if remaining is None or self.installment is None:
            return None
        return remaining * self.installment

    def to_dict(self) -> dict[str, Any]:
        return {
            "institution_raw": self.institution_raw,
            "institution_key": self.institution_key,
            "principal": _jsonable(self.principal),
            "installment": _jsonable(self.installment),
            "installment_count": _jsonable(self.installment_count),
            "last_installment_date": _jsonable(self.last_installment_date),
            "contract_repayment": _jsonable(self.contract_repayment),
            "paid_installments": _jsonable(self.paid_installments),
            "remaining_installments": _jsonable(self.remaining_installments),
            "remaining_amount": _jsonable(self.remaining_amount),
            "provenance": self.provenance.to_dict(),
            "issues": list(self.issues),
        }


@dataclass
class ObligationRecord:
    institution_raw: str
    institution_key: str
    amount: Decimal
    provenance: Provenance
    label: str = ""
    issues: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "institution_raw": self.institution_raw,
            "institution_key": self.institution_key,
            "amount": _jsonable(self.amount),
            "label": self.label,
            "provenance": self.provenance.to_dict(),
            "issues": list(self.issues),
        }


@dataclass
class Calculation:
    name: str
    label: str
    value: Decimal | int | float | None
    unit: str = ""
    method: str = ""
    inputs: list[str] = field(default_factory=list)
    sources: list[Provenance] = field(default_factory=list)
    verification_value: Decimal | None = None
    verification_method: str = ""
    verified: bool | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "label": self.label,
            "value": _jsonable(self.value),
            "unit": self.unit,
            "method": self.method,
            "inputs": list(self.inputs),
            "sources": [source.to_dict() for source in self.sources],
            "verification_value": _jsonable(self.verification_value),
            "verification_method": self.verification_method,
            "verified": self.verified,
        }


@dataclass
class AuditResult:
    measure: str
    label: str
    declared: Decimal | None
    computed: Decimal | None
    difference: Decimal | None
    status: AuditStatus
    declared_source: Provenance | None = None
    method: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "measure": self.measure,
            "label": self.label,
            "declared": _jsonable(self.declared),
            "computed": _jsonable(self.computed),
            "difference": _jsonable(self.difference),
            "status": self.status.value,
            "declared_source": self.declared_source.to_dict() if self.declared_source else None,
            "method": self.method,
        }


@dataclass
class ReconciliationItem:
    institution: str
    loan_total: Decimal
    obligation_total: Decimal
    difference: Decimal | None
    status: MatchStatus
    loan_name: str = ""
    obligation_name: str = ""
    similarity: float = 1.0
    note: str = ""
    loan_sources: list[Provenance] = field(default_factory=list)
    obligation_sources: list[Provenance] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "institution": self.institution,
            "loan_total": _jsonable(self.loan_total),
            "obligation_total": _jsonable(self.obligation_total),
            "difference": _jsonable(self.difference),
            "status": self.status.value,
            "loan_name": self.loan_name,
            "obligation_name": self.obligation_name,
            "similarity": round(self.similarity, 4),
            "note": self.note,
            "loan_sources": [source.to_dict() for source in self.loan_sources],
            "obligation_sources": [source.to_dict() for source in self.obligation_sources],
        }


@dataclass
class Anomaly:
    severity: Severity
    category: str
    message: str
    sheet: str = ""
    address: str = ""
    detail: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "severity": self.severity.value,
            "category": self.category,
            "message": self.message,
            "sheet": self.sheet,
            "address": self.address,
            "detail": self.detail,
        }


@dataclass
class SheetMeta:
    title: str
    state: str
    declared_dimensions: str
    used_dimensions: str
    max_row: int
    max_col: int
    merged_ranges: int = 0
    hidden_rows: int = 0
    hidden_cols: int = 0
    protected: bool = False
    formula_count: int = 0
    error_cells: list[str] = field(default_factory=list)
    table_count: int = 0
    suspicious: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "state": self.state,
            "declared_dimensions": self.declared_dimensions,
            "used_dimensions": self.used_dimensions,
            "max_row": self.max_row,
            "max_col": self.max_col,
            "merged_ranges": self.merged_ranges,
            "hidden_rows": self.hidden_rows,
            "hidden_cols": self.hidden_cols,
            "protected": self.protected,
            "formula_count": self.formula_count,
            "error_cells": list(self.error_cells),
            "table_count": self.table_count,
            "suspicious": list(self.suspicious),
        }


@dataclass
class WorkbookMeta:
    path: str
    sha256: str
    size_bytes: int
    sheets: list[SheetMeta] = field(default_factory=list)
    defined_names: list[str] = field(default_factory=list)
    has_macros: bool = False
    external_links: int = 0
    total_formulas: int = 0
    total_error_cells: int = 0
    load_seconds: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "sha256": self.sha256,
            "size_bytes": self.size_bytes,
            "sheets": [sheet.to_dict() for sheet in self.sheets],
            "defined_names": list(self.defined_names),
            "has_macros": self.has_macros,
            "external_links": self.external_links,
            "total_formulas": self.total_formulas,
            "total_error_cells": self.total_error_cells,
            "load_seconds": round(self.load_seconds, 6),
        }


@dataclass
class DeclaredTotal:
    """A total written in the workbook, with the cell it came from."""

    measure: str
    label: str
    value: Decimal
    provenance: Provenance
    mapping: str = "label-next-cell"
    confidence: float = 0.6

    def to_dict(self) -> dict[str, Any]:
        return {
            "measure": self.measure,
            "label": self.label,
            "value": _jsonable(self.value),
            "provenance": self.provenance.to_dict(),
            "mapping": self.mapping,
            "confidence": round(self.confidence, 4),
        }


@dataclass
class AnalysisResult:
    meta: WorkbookMeta
    tables: list[TableCandidate] = field(default_factory=list)
    primary_table: TableCandidate | None = None
    loans: list[LoanRecord] = field(default_factory=list)
    obligations: list[ObligationRecord] = field(default_factory=list)
    calculations: list[Calculation] = field(default_factory=list)
    audits: list[AuditResult] = field(default_factory=list)
    reconciliation: list[ReconciliationItem] = field(default_factory=list)
    anomalies: list[Anomaly] = field(default_factory=list)
    declared_totals: list[DeclaredTotal] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    structure_confidence: float = 0.0
    normalized_rows: list[dict[str, Any]] = field(default_factory=list)
    column_profiles: list[ColumnProfile] = field(default_factory=list)
    ai: dict[str, Any] = field(default_factory=dict)
    validation: list[Any] = field(default_factory=list)
    # Generic statistical profile; typed as Any to avoid a models <-> analyzer import cycle.
    generic: Any = None
    # Per-sheet profiles and workbook-level structure intelligence.  Typed as Any
    # for the same reason; both objects implement ``to_dict``.
    sheets: list[Any] = field(default_factory=list)
    intelligence: Any = None
    # Structured, traceable extraction model (workbook -> sheets -> blocks ->
    # columns/cells/period evidence).  Typed as Any to avoid a models <-> excel
    # import cycle; the object implements ``to_dict``.
    extraction: Any = None
    # Explicit semantic roles per extracted field and recorded data-quality
    # findings.  Both are produced before statistical conclusions are drawn.
    semantics: Any = None
    quality: Any = None
    # Deterministic reconciliation of declared totals, subtotals, groups,
    # monthly cells and cross-period figures.  Every record carries the
    # expected/calculated values, tolerance, status and source addresses.
    statistical_reconciliation: Any = None
    # Controlled semantic matching of equivalent fields across worksheets and
    # files: normalized labels, units, periods, identifier/value evidence and
    # matched/probable/ambiguous/rejected decisions.  It never aggregates.
    semantic_matching: Any = None
    # Duplicate/derived-sheet control and safe-aggregation decisions across
    # worksheets and files.  Every inclusion, exclusion and blocked decision is
    # recorded; nothing is aggregated, repaired or deleted here.
    duplicate_control: Any = None
    # Period matrix built from the periodic sheets (month x period x metric) and
    # the grounded answers to the user's Persian request.
    integration: Any = None
    request_intent: Any = None
    request_answers: list[Any] = field(default_factory=list)

    def calculation(self, name: str) -> Calculation | None:
        for item in self.calculations:
            if item.name == name:
                return item
        return None

    def value(self, name: str) -> Decimal | int | float | None:
        item = self.calculation(name)
        return None if item is None else item.value

    def to_dict(self) -> dict[str, Any]:
        return {
            "meta": self.meta.to_dict(),
            "tables": [table.to_dict() for table in self.tables],
            "primary_table": self.primary_table.to_dict() if self.primary_table else None,
            "loans": [loan.to_dict() for loan in self.loans],
            "obligations": [item.to_dict() for item in self.obligations],
            "calculations": [item.to_dict() for item in self.calculations],
            "audits": [item.to_dict() for item in self.audits],
            "reconciliation": [item.to_dict() for item in self.reconciliation],
            "anomalies": [item.to_dict() for item in self.anomalies],
            "declared_totals": [item.to_dict() for item in self.declared_totals],
            "warnings": list(self.warnings),
            "structure_confidence": round(self.structure_confidence, 4),
            "ai": dict(self.ai),
            "generic": self.generic.to_dict() if hasattr(self.generic, "to_dict") else self.generic,
            "sheets": [
                item.to_dict() if hasattr(item, "to_dict") else item for item in self.sheets
            ],
            "intelligence": (
                self.intelligence.to_dict()
                if hasattr(self.intelligence, "to_dict")
                else self.intelligence
            ),
            "extraction": (
                self.extraction.to_dict()
                if hasattr(self.extraction, "to_dict")
                else self.extraction
            ),
            "semantics": (
                self.semantics.to_dict()
                if hasattr(self.semantics, "to_dict")
                else self.semantics
            ),
            "quality": (
                self.quality.to_dict()
                if hasattr(self.quality, "to_dict")
                else self.quality
            ),
            "statistical_reconciliation": (
                self.statistical_reconciliation.to_dict()
                if hasattr(self.statistical_reconciliation, "to_dict")
                else self.statistical_reconciliation
            ),
            "semantic_matching": (
                self.semantic_matching.to_dict()
                if hasattr(self.semantic_matching, "to_dict")
                else self.semantic_matching
            ),
            "duplicate_control": (
                self.duplicate_control.to_dict()
                if hasattr(self.duplicate_control, "to_dict")
                else self.duplicate_control
            ),
            "validation": [
                item.to_dict() if hasattr(item, "to_dict") else item for item in self.validation
            ],
            "request_intent": (
                self.request_intent.to_dict()
                if hasattr(self.request_intent, "to_dict")
                else self.request_intent
            ),
            "request_answers": [
                item.to_dict() if hasattr(item, "to_dict") else item
                for item in self.request_answers
            ],
            "integration": (
                self.integration.to_dict()
                if hasattr(self.integration, "to_dict")
                else self.integration
            ),
        }


@dataclass
class RunResult:
    output_path: str
    status: str
    input_sha256: str = ""
    input_unchanged: bool = False
    analysis: AnalysisResult | None = None
    performance: dict[str, Any] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)
    report: Any = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": self.output_path,
            "status": self.status,
            "input_sha256": self.input_sha256,
            "input_unchanged": self.input_unchanged,
            "analysis": self.analysis.to_dict() if self.analysis else None,
            "performance": dict(self.performance),
            "errors": list(self.errors),
            "report": (
                self.report.to_dict()
                if hasattr(self.report, "to_dict")
                else self.report
            ),
        }


def _jsonable(value: Any) -> Any:
    """Convert Decimals/dates to JSON-friendly primitives without losing money."""
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, dict):
        return {key: _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    if isinstance(value, Enum):
        return value.value
    return value


def to_excel_number(value: Any) -> float | int | None:
    """Convert a Decimal/str/number to a plain Excel number (last step only)."""
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, Decimal):
        return float(value) if value != value.to_integral_value() else int(value)
    if isinstance(value, (int, float)):
        return value
    return None
