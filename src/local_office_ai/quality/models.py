"""One quality vocabulary for every processing stage.

The engine already produces specialised reports: the semantic classifier, the
data-quality analyzer, the statistical reconciler, duplicate control and the
output validators.  This module does not duplicate those checks.  It gives them
one shared status model, one stage vocabulary and one evidence container so a
caller can decide whether the delivered report is complete, qualified or
blocked.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, ClassVar, Iterable


class GateStatus(str, Enum):
    """The three outcomes every quality stage must use."""

    PASS = "PASS"
    WARNING = "WARNING"
    FAIL = "FAIL"

    @property
    def label_fa(self) -> str:
        return GATE_STATUS_FA[self]


GATE_STATUS_FA = {
    GateStatus.PASS: "قبول",
    GateStatus.WARNING: "هشدار",
    GateStatus.FAIL: "رد",
}


def status_of(value: Any) -> GateStatus | None:
    """Read a status produced by a legacy enum or a plain string."""

    if value is None:
        return None
    if isinstance(value, GateStatus):
        return value
    text = getattr(value, "value", value)
    normalized = str(text).strip().upper()
    mapping = {
        "PASS": GateStatus.PASS,
        "PASSED": GateStatus.PASS,
        "ALLOWED": GateStatus.PASS,
        "VERIFIED": GateStatus.PASS,
        "WARNING": GateStatus.WARNING,
        "REVIEW": GateStatus.WARNING,
        "SKIPPED": GateStatus.WARNING,
        "NOT_CHECKED": GateStatus.WARNING,
        "NOT_VERIFIABLE": GateStatus.WARNING,
        "FAIL": GateStatus.FAIL,
        "FAILED": GateStatus.FAIL,
        "BLOCK": GateStatus.FAIL,
        "BLOCKED": GateStatus.FAIL,
        "MISMATCH": GateStatus.FAIL,
    }
    return mapping.get(normalized)


def worst_status(statuses: Iterable[GateStatus | str | None]) -> GateStatus:
    """Aggregate statuses with FAIL > WARNING > PASS.  Empty means WARNING."""

    resolved = [status_of(item) for item in statuses]
    resolved = [item for item in resolved if item is not None]
    if GateStatus.FAIL in resolved:
        return GateStatus.FAIL
    if GateStatus.WARNING in resolved:
        return GateStatus.WARNING
    if not resolved:
        # Nothing was checked, so nothing is confirmed.
        return GateStatus.WARNING
    return GateStatus.PASS


class GateStage(str, Enum):
    """The thirteen stages named by the quality-gate contract."""

    SOURCE_AVAILABILITY = "source_availability"
    FILE_READABILITY = "file_readability"
    EXTRACTION_COMPLETENESS = "extraction_completeness"
    BLOCK_DETECTION = "block_detection"
    PERIOD_IDENTIFICATION = "period_identification"
    SEMANTIC_CLASSIFICATION = "semantic_classification"
    DATA_QUALITY = "data_quality"
    STATISTICAL_VALIDITY = "statistical_validity"
    DUPLICATE_DETECTION = "duplicate_detection"
    CROSS_SOURCE_RECONCILIATION = "cross_source_reconciliation"
    PROVENANCE_COMPLETENESS = "provenance_completeness"
    OUTPUT_FORMAT = "output_format"
    DOCUMENT_INTEGRITY = "document_integrity"

    @property
    def label_fa(self) -> str:
        return STAGE_LABELS_FA[self]


STAGE_LABELS_FA = {
    GateStage.SOURCE_AVAILABILITY: "دسترس‌پذیری منبع",
    GateStage.FILE_READABILITY: "خوانایی فایل",
    GateStage.EXTRACTION_COMPLETENESS: "کامل بودن استخراج",
    GateStage.BLOCK_DETECTION: "تشخیص بلوک داده",
    GateStage.PERIOD_IDENTIFICATION: "شناسایی دوره",
    GateStage.SEMANTIC_CLASSIFICATION: "طبقه‌بندی معنایی",
    GateStage.DATA_QUALITY: "کیفیت داده",
    GateStage.STATISTICAL_VALIDITY: "اعتبار آماری",
    GateStage.DUPLICATE_DETECTION: "تشخیص تکرار و هم‌پوشانی",
    GateStage.CROSS_SOURCE_RECONCILIATION: "تطبیق بین منابع",
    GateStage.PROVENANCE_COMPLETENESS: "کامل بودن ردیابی منبع",
    GateStage.OUTPUT_FORMAT: "درستی قالب خروجی",
    GateStage.DOCUMENT_INTEGRITY: "یکپارچگی سند نهایی",
}

STAGE_ORDER: tuple[GateStage, ...] = tuple(GateStage)

# A failure in these stages means the application cannot claim a valid report.
CRITICAL_STAGES: frozenset[GateStage] = frozenset(
    {
        GateStage.SOURCE_AVAILABILITY,
        GateStage.FILE_READABILITY,
        GateStage.EXTRACTION_COMPLETENESS,
        GateStage.SEMANTIC_CLASSIFICATION,
        GateStage.PROVENANCE_COMPLETENESS,
        GateStage.OUTPUT_FORMAT,
        GateStage.DOCUMENT_INTEGRITY,
    }
)


@dataclass
class GateFinding:
    """One concrete observation belonging to a stage."""

    code: str
    status: GateStatus
    message: str
    location: str = ""
    affected: int = 0
    blocking: bool = False
    evidence: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "status": self.status.value,
            "status_label": self.status.label_fa,
            "message": self.message,
            "location": self.location,
            "affected": self.affected,
            "blocking": self.blocking,
            "evidence": dict(self.evidence),
        }


@dataclass
class GateStageResult:
    """The verdict and the evidence of exactly one stage."""

    stage: GateStage
    status: GateStatus
    summary: str
    checks: list[GateFinding] = field(default_factory=list)
    evidence: dict[str, Any] = field(default_factory=dict)
    evaluated: bool = True

    @property
    def blocking(self) -> bool:
        if self.status is not GateStatus.FAIL:
            return False
        if any(check.blocking for check in self.checks):
            return True
        return self.stage in CRITICAL_STAGES

    @property
    def label_fa(self) -> str:
        return self.stage.label_fa

    def to_dict(self) -> dict[str, Any]:
        return {
            "stage": self.stage.value,
            "stage_label": self.label_fa,
            "status": self.status.value,
            "status_label": self.status.label_fa,
            "summary": self.summary,
            "blocking": self.blocking,
            "evaluated": self.evaluated,
            "checks": [check.to_dict() for check in self.checks],
            "evidence": dict(self.evidence),
        }


@dataclass
class ProvenanceSource:
    """A source cell or range that a derived value points back to."""

    filename: str
    sheet: str = ""
    block: str = ""
    cell: str = ""
    label: str = ""
    original_value: str = ""
    normalized_value: str = ""
    period: str = ""
    unit: str = ""
    formula: str = ""
    validation_status: str = ""
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "filename": self.filename,
            "sheet": self.sheet,
            "block": self.block,
            "cell": self.cell,
            "label": self.label,
            "original_value": self.original_value,
            "normalized_value": self.normalized_value,
            "period": self.period,
            "unit": self.unit,
            "formula": self.formula,
            "validation_status": self.validation_status,
            "note": self.note,
        }


@dataclass
class ProvenanceRecord:
    """The full trail of one important figure, including its source values."""

    key: str
    filename: str
    sheet: str = ""
    block: str = ""
    cell: str = ""
    original_label: str = ""
    normalized_label: str = ""
    original_value: str = ""
    normalized_value: str = ""
    period: str = ""
    unit: str = ""
    operation: str = ""
    formula: str = ""
    validation_status: str = "PASS"
    sources: list[ProvenanceSource] = field(default_factory=list)
    related_records: list[str] = field(default_factory=list)
    derived: bool = False
    important: bool = True
    notes: list[str] = field(default_factory=list)

    REQUIRED_FIELDS: ClassVar[tuple[str, ...]] = (
        "filename",
        "sheet",
        "cell",
        "normalized_value",
    )

    def missing(self, field_name: str) -> bool:
        value = getattr(self, field_name, "")
        return value is None or str(value).strip() == ""

    def missing_fields(self) -> list[str]:
        missing = [name for name in self.REQUIRED_FIELDS if self.missing(name)]
        if self.derived and not self.sources:
            missing.append("sources")
        return missing

    @property
    def complete(self) -> bool:
        return not self.missing_fields()

    def optional_coverage(self) -> float:
        fields = (
            "original_label",
            "normalized_label",
            "period",
            "unit",
            "operation",
            "formula",
        )
        filled = sum(1 for name in fields if not self.missing(name))
        return filled / len(fields)

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "filename": self.filename,
            "sheet": self.sheet,
            "block": self.block,
            "cell": self.cell,
            "original_label": self.original_label,
            "normalized_label": self.normalized_label,
            "original_value": self.original_value,
            "normalized_value": self.normalized_value,
            "period": self.period,
            "unit": self.unit,
            "operation": self.operation,
            "formula": self.formula,
            "validation_status": self.validation_status,
            "sources": [source.to_dict() for source in self.sources],
            "related_records": list(self.related_records),
            "derived": self.derived,
            "important": self.important,
            "complete": self.complete,
            "missing_fields": self.missing_fields(),
            "optional_coverage": round(self.optional_coverage(), 3),
            "notes": list(self.notes),
        }


@dataclass
class ProvenanceLedger:
    """Machine-readable evidence behind every important numeric result."""

    filename: str = ""
    source_path: str = ""
    records: list[ProvenanceRecord] = field(default_factory=list)

    def add(self, record: ProvenanceRecord) -> ProvenanceRecord:
        for index, existing in enumerate(self.records):
            if existing.key == record.key:
                self.records[index] = record
                return record
        self.records.append(record)
        return record

    def record(self, key: str) -> ProvenanceRecord | None:
        for item in self.records:
            if item.key == key:
                return item
        return None

    @property
    def important_records(self) -> list[ProvenanceRecord]:
        return [item for item in self.records if item.important]

    @property
    def incomplete_records(self) -> list[ProvenanceRecord]:
        return [item for item in self.important_records if not item.complete]

    @property
    def derived_without_sources(self) -> list[ProvenanceRecord]:
        return [
            item
            for item in self.important_records
            if item.derived and not item.sources
        ]

    def missing_fields(self) -> dict[str, list[str]]:
        missing: dict[str, list[str]] = {}
        for item in self.important_records:
            for field_name in item.missing_fields():
                missing.setdefault(field_name, []).append(item.key)
        return missing

    @property
    def status(self) -> GateStatus:
        if not self.important_records:
            return GateStatus.PASS
        if self.derived_without_sources:
            return GateStatus.FAIL
        if any(item.missing(name) for item in self.important_records for name in ProvenanceRecord.REQUIRED_FIELDS):
            return GateStatus.FAIL
        return GateStatus.PASS

    def counts(self) -> dict[str, int]:
        important = self.important_records
        return {
            "records": len(self.records),
            "important": len(important),
            "derived": sum(1 for item in important if item.derived),
            "with_sources": sum(1 for item in important if item.sources),
            "complete": sum(1 for item in important if item.complete),
            "incomplete": len(self.incomplete_records),
        }

    def coverage(self) -> float:
        important = self.important_records
        if not important:
            return 1.0
        return sum(item.optional_coverage() for item in important) / len(important)

    def to_dict(self) -> dict[str, Any]:
        counts = self.counts()
        return {
            "filename": self.filename,
            "source_path": self.source_path,
            "status": self.status.value,
            "status_label": self.status.label_fa,
            "counts": counts,
            "coverage": round(self.coverage(), 3),
            "missing_fields": {key: list(value) for key, value in self.missing_fields().items()},
            "derived_without_sources": [item.key for item in self.derived_without_sources],
            "records": [item.to_dict() for item in self.records],
        }


@dataclass
class QualityGateReport:
    """The single verdict the application uses to describe a delivery."""

    stages: list[GateStageResult] = field(default_factory=list)
    provenance: ProvenanceLedger | None = None
    source: str = ""
    output_path: str = ""
    generated_at: str = ""
    finalized: bool = False
    metadata: dict[str, Any] = field(default_factory=dict)

    def stage(self, stage: GateStage) -> GateStageResult | None:
        for item in self.stages:
            if item.stage is stage:
                return item
        return None

    def status_for(self, stage: GateStage) -> GateStatus | None:
        item = self.stage(stage)
        return item.status if item is not None else None

    @property
    def status(self) -> GateStatus:
        return worst_status(item.status for item in self.stages)

    @property
    def status_label(self) -> str:
        return self.status.label_fa

    @property
    def is_complete_success(self) -> bool:
        return bool(self.stages) and self.status is GateStatus.PASS

    @property
    def blocking_failures(self) -> list[GateStageResult]:
        return [item for item in self.stages if item.blocking]

    @property
    def failures(self) -> list[GateStageResult]:
        return [item for item in self.stages if item.status is GateStatus.FAIL]

    @property
    def warnings(self) -> list[GateStageResult]:
        return [item for item in self.stages if item.status is GateStatus.WARNING]

    @property
    def blocked(self) -> bool:
        return bool(self.blocking_failures)

    @property
    def can_claim_valid_report(self) -> bool:
        return bool(self.stages) and not self.blocked

    @property
    def evaluated_stages(self) -> list[GateStage]:
        return [item.stage for item in self.stages if item.evaluated]

    @property
    def missing_stages(self) -> list[GateStage]:
        evaluated = set(self.evaluated_stages)
        return [stage for stage in STAGE_ORDER if stage not in evaluated]

    def counts(self) -> dict[str, int]:
        counts = {status.value: 0 for status in GateStatus}
        for item in self.stages:
            counts[item.status.value] += 1
        counts["stages"] = len(self.stages)
        counts["blocking"] = len(self.blocking_failures)
        return counts

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status.value,
            "status_label": self.status_label,
            "is_complete_success": self.is_complete_success,
            "can_claim_valid_report": self.can_claim_valid_report,
            "blocked": self.blocked,
            "finalized": self.finalized,
            "source": self.source,
            "output_path": self.output_path,
            "generated_at": self.generated_at,
            "counts": self.counts(),
            "missing_stages": [stage.value for stage in self.missing_stages],
            "blocking_stages": [item.stage.value for item in self.blocking_failures],
            "stages": [item.to_dict() for item in self.stages],
            "provenance": self.provenance.to_dict() if self.provenance else None,
            "metadata": dict(self.metadata),
        }
