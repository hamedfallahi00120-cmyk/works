"""Run the complete multi-source acceptance scenario and audit its evidence.

This module is the honest answer to "did the application really deliver a valid
report?".  It processes the supplied files through the normal engine, records
all fifteen audit items the acceptance contract requires, compares optional
verification checkpoints with independently extracted values, and refuses to
call the run successful when a source was skipped, a block was ignored, a
required period is missing, a critical check failed, the output format is wrong,
the DOCX is invalid, or the document contains an ungrounded number.

Checkpoints are **verification inputs only**.  They are never inserted into the
report and never replace an extracted value.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

from local_office_ai.app.folder_workflow import FolderWorkflow
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.settings import Settings
from local_office_ai.excel.indicator_periods import (
    IndicatorPeriodReport,
    build_indicator_periods,
)
from local_office_ai.excel.normalization import normalize_digits, normalize_text
from local_office_ai.quality import GateStatus
from local_office_ai.reporting.docx_validation import validate_docx_output
from local_office_ai.reporting.formats import OutputFormat
from local_office_ai.reporting.template import inspect_docx_template

LOGGER = get_logger(__name__)

STATUS_ACCEPTED = "accepted"
STATUS_PARTIAL = "partial"
STATUS_FAILED = "failed"
STATUS_UNVERIFIED = "unverified"

STATUS_FA = {
    STATUS_ACCEPTED: "پذیرفته شد",
    STATUS_PARTIAL: "جزئی — پذیرش کامل نیست",
    STATUS_FAILED: "ناموفق — گزارش معتبر نیست",
    STATUS_UNVERIFIED: "تأیید نشده — فایل‌های منبع موجود نیستند",
}

CHECK_PASS = "PASS"
CHECK_WARNING = "WARNING"
CHECK_FAIL = "FAIL"
CHECK_UNVERIFIED = "UNVERIFIED"


@dataclass
class AcceptanceInputFile:
    path: str
    role: str
    exists: bool
    kind: str = ""
    size_bytes: int = 0
    sha256: str = ""
    error: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "role": self.role,
            "exists": self.exists,
            "kind": self.kind,
            "size_bytes": self.size_bytes,
            "sha256": self.sha256,
            "error": self.error,
        }


@dataclass
class AcceptanceCheckpoint:
    label: str
    previous: Decimal | None = None
    current: Decimal | None = None
    reported_percent: Decimal | None = None
    aliases: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.previous = _decimal(self.previous)
        self.current = _decimal(self.current)
        self.reported_percent = _decimal(self.reported_percent)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "AcceptanceCheckpoint":
        return cls(
            label=str(payload.get("label") or ""),
            previous=_decimal(payload.get("previous")),
            current=_decimal(payload.get("current")),
            reported_percent=_decimal(
                payload.get("reported_percent", payload.get("percent"))
            ),
            aliases=[str(item) for item in payload.get("aliases", []) or []],
        )

    @property
    def arithmetic_percent(self) -> Decimal | None:
        if self.previous in (None, Decimal("0")) or self.current is None:
            return None
        return ((self.current - self.previous) / abs(self.previous) * 100).quantize(
            Decimal("0.01")
        )

    @property
    def arithmetic_consistent(self) -> bool | None:
        if self.reported_percent is None:
            return None
        calculated = self.arithmetic_percent
        if calculated is None:
            return False
        return abs(calculated - self.reported_percent) <= Decimal("0.01")

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "aliases": list(self.aliases),
            "previous": _text(self.previous),
            "current": _text(self.current),
            "reported_percent": _text(self.reported_percent),
            "arithmetic_percent": _text(self.arithmetic_percent),
            "arithmetic_consistent": self.arithmetic_consistent,
        }


@dataclass
class CheckpointResult:
    label: str
    status: str
    matched_indicator: str = ""
    expected_previous: Decimal | None = None
    expected_current: Decimal | None = None
    expected_percent: Decimal | None = None
    extracted_previous: Decimal | None = None
    extracted_current: Decimal | None = None
    extracted_percent: Decimal | None = None
    previous_address: str = ""
    current_address: str = ""
    note: str = ""

    @property
    def matched(self) -> bool:
        return self.status == CHECK_PASS

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "status": self.status,
            "matched_indicator": self.matched_indicator,
            "expected_previous": _text(self.expected_previous),
            "expected_current": _text(self.expected_current),
            "expected_percent": _text(self.expected_percent),
            "extracted_previous": _text(self.extracted_previous),
            "extracted_current": _text(self.extracted_current),
            "extracted_percent": _text(self.extracted_percent),
            "previous_address": self.previous_address,
            "current_address": self.current_address,
            "note": self.note,
        }


@dataclass
class FalseSuccessCheck:
    code: str
    status: str
    message: str
    evidence: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "status": self.status,
            "message": self.message,
            "evidence": dict(self.evidence),
        }


@dataclass
class AcceptanceReport:
    status: str = STATUS_UNVERIFIED
    status_label: str = ""
    request: str = ""
    inputs: list[AcceptanceInputFile] = field(default_factory=list)
    files: list[dict[str, Any]] = field(default_factory=list)
    worksheets: list[dict[str, Any]] = field(default_factory=list)
    blocks: list[dict[str, Any]] = field(default_factory=list)
    periods: list[dict[str, Any]] = field(default_factory=list)
    semantics: list[dict[str, Any]] = field(default_factory=list)
    comparisons: list[dict[str, Any]] = field(default_factory=list)
    cross_source_matches: list[dict[str, Any]] = field(default_factory=list)
    matching: dict[str, Any] = field(default_factory=dict)
    ambiguous: list[dict[str, Any]] = field(default_factory=list)
    duplicates: dict[str, Any] = field(default_factory=dict)
    reconciliation: dict[str, Any] = field(default_factory=dict)
    checks: list[FalseSuccessCheck] = field(default_factory=list)
    provenance: dict[str, Any] = field(default_factory=dict)
    outputs: dict[str, Any] = field(default_factory=dict)
    artifacts: dict[str, Any] = field(default_factory=dict)
    docx: dict[str, Any] = field(default_factory=dict)
    accidents: dict[str, Any] = field(default_factory=dict)
    checkpoints: list[CheckpointResult] = field(default_factory=list)
    limitations: list[str] = field(default_factory=list)
    unverified_stages: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    generated_at: str = ""
    result: Any = None

    def check(self, code: str) -> FalseSuccessCheck | None:
        for item in self.checks:
            if item.code == code:
                return item
        return None

    @property
    def failed_checks(self) -> list[FalseSuccessCheck]:
        return [item for item in self.checks if item.status == CHECK_FAIL]

    @property
    def warning_checks(self) -> list[FalseSuccessCheck]:
        return [item for item in self.checks if item.status == CHECK_WARNING]

    @property
    def unverified_checks(self) -> list[FalseSuccessCheck]:
        return [item for item in self.checks if item.status == CHECK_UNVERIFIED]

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "status_label": self.status_label or STATUS_FA.get(self.status, self.status),
            "request": self.request,
            "generated_at": self.generated_at,
            "inputs": [item.to_dict() for item in self.inputs],
            "files": list(self.files),
            "worksheets": list(self.worksheets),
            "blocks": list(self.blocks),
            "periods": list(self.periods),
            "semantics": list(self.semantics),
            "comparisons": list(self.comparisons),
            "cross_source_matches": list(self.cross_source_matches),
            "matching": dict(self.matching),
            "ambiguous": list(self.ambiguous),
            "duplicates": dict(self.duplicates),
            "reconciliation": dict(self.reconciliation),
            "checks": [item.to_dict() for item in self.checks],
            "provenance": dict(self.provenance),
            "outputs": dict(self.outputs),
            "artifacts": dict(self.artifacts),
            "docx": dict(self.docx),
            "accidents": dict(self.accidents),
            "checkpoints": [item.to_dict() for item in self.checkpoints],
            "limitations": list(self.limitations),
            "unverified_stages": list(self.unverified_stages),
            "warnings": list(self.warnings),
            "errors": list(self.errors),
        }

    def to_markdown(self) -> str:
        return render_markdown(self)

    def write(self, output_dir: Path | str) -> tuple[Path, Path]:
        target = Path(output_dir)
        target.mkdir(parents=True, exist_ok=True)
        base = target / "acceptance_report"
        json_path = base.with_suffix(".json")
        md_path = base.with_suffix(".md")
        self.artifacts = {"json": str(json_path), "markdown": str(md_path)}
        json_path.write_text(
            json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )
        md_path.write_text(self.to_markdown(), encoding="utf-8")
        return json_path, md_path


class AcceptanceRunner:
    """Process the real inputs and evaluate the acceptance contract."""

    def __init__(
        self,
        settings: Settings,
        *,
        progress: Callable[[str], None] | None = None,
    ) -> None:
        self.settings = settings
        self.progress = progress or (lambda message: LOGGER.info(message))

    def run(
        self,
        inputs: Sequence[Path | str],
        *,
        request: str,
        output: Path | str | None = None,
        output_dir: Path | str | None = None,
        template: Path | str | None = None,
        checkpoints: Iterable[AcceptanceCheckpoint] = (),
        use_ai: bool = False,
        output_format: str = "docx",
    ) -> AcceptanceReport:
        report = AcceptanceReport(
            request=request,
            generated_at=__import__("datetime").datetime.now().isoformat(timespec="seconds"),
        )
        detected = [_inspect_input(Path(item).expanduser()) for item in inputs]
        report.inputs = detected
        excel_inputs = [item for item in detected if item.kind == "excel" and item.exists]
        missing_excel = [
            item for item in detected if item.kind == "excel" and not item.exists
        ]
        reference_inputs = [item for item in detected if item.kind == "docx"]
        missing_reference = [item for item in reference_inputs if not item.exists]

        if not excel_inputs:
            report.status = STATUS_UNVERIFIED
            report.status_label = STATUS_FA[STATUS_UNVERIFIED]
            report.unverified_stages = [
                "خواندن فایل‌های Excel",
                "بازرسی برگه‌ها",
                "تشخیص بلوک داده",
                "شناسایی دوره",
                "تطبیق شاخص",
                "کنترل تکرار",
                "تطبیق آماری",
                "ردیابی منبع",
                "تولید DOCX",
                "اعتبارسنجی DOCX",
            ]
            report.errors.append(
                "هیچ فایل Excel منبع موجود نیست؛ هیچ موفقیت سرتاسری ادعا نمی‌شود."
            )
            report.limitations.extend(
                f"فایل منبع موجود نیست: {item.path}" for item in missing_excel
            )
            report.limitations.extend(
                f"فایل مرجع Word موجود نیست: {item.path}" for item in missing_reference
            )
            report.checkpoints = _unverified_checkpoints(checkpoints)
            return report

        if missing_excel:
            report.warnings.append(
                "برخی فایل‌های منبع موجود نیستند؛ نتیجه حداکثر «جزئی» است."
            )
            report.limitations.extend(
                f"فایل منبع موجود نیست: {item.path}" for item in missing_excel
            )

        if missing_reference:
            report.warnings.append(
                "فایل مرجع Word اعلام‌شده موجود نیست؛ گزارش بدون قالب مرجع ساخته می‌شود."
            )
            report.limitations.extend(
                f"فایل مرجع Word موجود نیست: {item.path}" for item in missing_reference
            )

        if missing_excel:
            report.unverified_stages.extend(
                [
                    "پردازش همه فایل‌های Excel اعلام‌شده (فایل موجود نیست)",
                    "تطبیق کامل شاخص‌ها میان همه منابع",
                    "کنترل تکرار و تجمیع بین همه منابع",
                    "تطبیق آماری و ردیابی منبع برای فایل‌های غایب",
                ]
            )
        if missing_reference:
            report.unverified_stages.append(
                "استفاده از قالب Word مرجع (اختیاری؛ گزارش بدون قالب ساخته شد)"
            )

        if reference_inputs:
            for item in reference_inputs:
                if item.exists:
                    try:
                        inspection = inspect_docx_template(item.path)
                        report.limitations.append(
                            f"فایل مرجع Word «{Path(item.path).name}» فقط از نظر ساختار بررسی شد "
                            f"(بخش {inspection.section_count}، سبک {len(inspection.styles)}، "
                            f"جدول {len(inspection.tables)}) و داده تاریخی آن وارد گزارش نشد."
                        )
                    except Exception as exc:  # noqa: BLE001 - reference is optional
                        report.warnings.append(
                            f"بازرسی فایل مرجع Word ناموفق بود: {exc}"
                        )

        source_dir = Path(excel_inputs[0].path).parent
        output_path = Path(output) if output else None
        effective_template = Path(template) if template else next(
            (Path(item.path) for item in reference_inputs if item.exists), None
        )
        workflow = FolderWorkflow(self.settings, progress=self.progress)
        result = None
        try:
            result = workflow.run(
                source_dir,
                request=request,
                use_ai=use_ai,
                output=output_path,
                output_format=output_format,
                sources=[item.path for item in excel_inputs],
                template=effective_template,
            )
            report.result = result
        except Exception as exc:  # noqa: BLE001 - acceptance must report, not crash
            report.status = STATUS_FAILED
            report.errors.append(f"{type(exc).__name__}: {exc}")
            report.checks.append(
                FalseSuccessCheck(
                    "workflow_failed",
                    CHECK_FAIL,
                    f"اجرای کامل گردش‌کار ناموفق بود: {type(exc).__name__}: {exc}",
                )
            )
            report.checkpoints = _unverified_checkpoints(checkpoints)
            return report

        self._collect(report, result)
        self._checkpoints(report, result, checkpoints)
        self._false_success(report, result, output_format)
        self._final_status(report)
        if output_dir is not None:
            report.write(output_dir)
        return report

    # ------------------------------------------------------------------ #
    def _collect(self, report: AcceptanceReport, result: Any) -> None:
        for item in getattr(result, "files", []) or []:
            analysis = getattr(item, "analysis", None)
            blocks = _analysis_blocks(analysis)
            sheets = _analysis_sheets(analysis)
            report.files.append(
                {
                    "name": getattr(item, "name", ""),
                    "path": getattr(item, "path", ""),
                    "status": getattr(item, "status", ""),
                    "error": getattr(item, "error", ""),
                    "sheets": len(sheets),
                    "blocks": len(blocks),
                    "rows": getattr(item, "rows", 0),
                    "columns": getattr(item, "columns", 0),
                }
            )
            for sheet in sheets:
                report.worksheets.append(
                    {
                        "file": getattr(item, "name", ""),
                        "sheet": getattr(sheet, "name", ""),
                        "title": getattr(sheet, "title", ""),
                        "blocks": len(getattr(sheet, "blocks", []) or []),
                        "periods": [
                            str(period.get("label") or period.get("token") or "")
                            for period in getattr(sheet, "periods", []) or []
                            if isinstance(period, dict)
                        ],
                        "period_status": getattr(sheet, "period_status", ""),
                    }
                )
            semantics = getattr(analysis, "semantics", None) if analysis is not None else None
            for semantic_block in getattr(semantics, "blocks", []) or []:
                for column in getattr(semantic_block, "columns", []) or []:
                    report.semantics.append(
                        {
                            "file": getattr(item, "name", ""),
                            "sheet": getattr(semantic_block, "sheet", ""),
                            "block": getattr(semantic_block, "block_id", ""),
                            "header": getattr(column, "header", ""),
                            "role": str(
                                getattr(getattr(column, "role", ""), "value", "")
                            ),
                            "role_label": getattr(column, "role_label", ""),
                            "statistical_role": getattr(column, "statistical_role", ""),
                            "unit": getattr(column, "unit", ""),
                            "period": getattr(column, "period", ""),
                            "confidence": round(float(getattr(column, "confidence", 0.0) or 0.0), 4),
                        }
                    )
            for block in blocks:
                report.blocks.append(
                    {
                        "file": getattr(item, "name", ""),
                        "sheet": getattr(block, "sheet", ""),
                        "block": getattr(block, "block_id", ""),
                        "title": getattr(block, "title", ""),
                        "rows": (getattr(block, "boundaries", {}) or {}).get("row_count", 0),
                        "columns": (getattr(block, "boundaries", {}) or {}).get("column_count", 0),
                        "cells_truncated": getattr(block, "cells_truncated", False),
                        "periods": [
                            str(period.get("label") or period.get("token") or "")
                            for period in getattr(block, "periods", []) or []
                            if isinstance(period, dict)
                        ],
                        "period_status": getattr(block, "period_status", ""),
                    }
                )
            integration = getattr(analysis, "integration", None) if analysis is not None else None
            block_periods = sorted(
                {
                    str(period.get("label") or period.get("token") or "")
                    for block in blocks
                    for period in getattr(block, "periods", []) or []
                    if isinstance(period, dict)
                }
                - {""}
            )
            file_periods = list(getattr(integration, "periods", []) or []) or block_periods
            report.periods.append(
                {
                    "file": getattr(item, "name", ""),
                    "periods": file_periods,
                    "months": list(getattr(integration, "months", []) or []),
                    "status": "detected" if file_periods else "not_detected",
                }
            )

        self._match_indicators_across_sources(report, result)

        matching = getattr(result, "matching", None)
        if matching is not None:
            report.matching = matching.to_dict() if hasattr(matching, "to_dict") else {}
            counts = matching.counts() if hasattr(matching, "counts") else {}
            counts["row_indicators_matched"] = sum(
                1 for item in report.cross_source_matches if item["status"] == "matched"
            )
            counts["row_indicators_ambiguous"] = sum(
                1 for item in report.cross_source_matches if item["status"] == "ambiguous"
            )
            report.matching["counts"] = counts
            report.ambiguous.extend(
                [
                    {
                        "label": getattr(item, "normalized_label", ""),
                        "status": str(
                            getattr(
                                getattr(item, "status", ""),
                                "value",
                                getattr(item, "status", ""),
                            )
                        ),
                        "members": [
                            getattr(member, "display_label", "")
                            for member in getattr(item, "members", []) or []
                        ],
                        "reason": getattr(item, "reason", ""),
                    }
                    for item in getattr(matching, "decisions", []) or []
                    if str(getattr(getattr(item, "status", ""), "value", "")) == "ambiguous"
                ]
            )

        duplicate = getattr(result, "duplicate_control", None)
        if duplicate is not None:
            payload = duplicate.to_dict() if hasattr(duplicate, "to_dict") else {}
            report.duplicates = {
                "counts": payload.get("counts", {}),
                "groups": payload.get("duplicate_groups", []),
                "record_duplicates": payload.get("record_duplicates", []),
                "decisions": [
                    {
                        "kind": getattr(item, "kind", ""),
                        "status": str(
                            getattr(getattr(item, "status", ""), "value", getattr(item, "status", ""))
                        ),
                        "headers": sorted(getattr(item, "headers", set()) or []),
                        "sources": sorted(getattr(item, "sources", set()) or []),
                        "reason": getattr(item, "reason", ""),
                    }
                    for item in getattr(duplicate, "decisions", []) or []
                ],
                "exclusions": [
                    item.to_dict() if hasattr(item, "to_dict") else item
                    for item in getattr(duplicate, "exclusions", []) or []
                ],
                "warnings": list(getattr(duplicate, "warnings", []) or []),
            }

        reconciliation_rows = []
        for item in getattr(result, "files", []) or []:
            analysis = getattr(item, "analysis", None)
            statistical = getattr(analysis, "statistical_reconciliation", None) if analysis else None
            for record in getattr(statistical, "reconciliations", []) or []:
                reconciliation_rows.append(
                    {
                        "file": getattr(item, "name", ""),
                        "kind": "reconciliation",
                        "label": getattr(record, "label", ""),
                        "expected": _text(getattr(record, "expected", None)),
                        "calculated": _text(getattr(record, "calculated", None)),
                        "difference": _text(getattr(record, "difference", None)),
                        "status": str(
                            getattr(getattr(record, "status", ""), "value", getattr(record, "status", ""))
                        ),
                    }
                )
            for record in getattr(statistical, "comparisons", []) or []:
                reconciliation_rows.append(
                    {
                        "file": getattr(item, "name", ""),
                        "kind": "comparison",
                        "label": getattr(record, "label", ""),
                        "expected": _text(getattr(record, "expected", None)),
                        "calculated": _text(getattr(record, "calculated", None)),
                        "difference": _text(getattr(record, "difference", None)),
                        "status": str(
                            getattr(getattr(record, "status", ""), "value", getattr(record, "status", ""))
                        ),
                    }
                )
        report.reconciliation = {
            "results": reconciliation_rows,
            "failed": sum(1 for item in reconciliation_rows if item["status"].upper() == "FAIL"),
            "warnings": sum(
                1 for item in reconciliation_rows if item["status"].upper() == "WARNING"
            ),
        }

        indicators = getattr(result, "indicator_periods", None)
        if indicators is not None:
            report.matching["indicator_comparisons"] = len(
                getattr(indicators, "comparisons", []) or []
            )
            report.matching["indicator_periods"] = list(
                getattr(indicators, "periods", []) or []
            )
            report.comparisons = [
                item.to_dict() if hasattr(item, "to_dict") else dict(item)
                for item in getattr(indicators, "comparisons", []) or []
            ]

        accidents = getattr(result, "accident_semantics", None)
        if accidents is not None:
            report.accidents = (
                accidents.to_dict() if hasattr(accidents, "to_dict") else {}
            )

        gate = getattr(result, "quality_gate", None)
        if gate is not None:
            report.provenance = {
                "status": _status_value(
                    getattr(getattr(gate, "provenance", None), "status", getattr(gate, "status", ""))
                ),
                "gate_status": _status_value(getattr(gate, "status", "")),
                "counts": (
                    gate.provenance.counts()
                    if getattr(gate, "provenance", None) is not None
                    and hasattr(gate.provenance, "counts")
                    else {}
                ),
                "coverage": (
                    round(gate.provenance.coverage(), 4)
                    if getattr(gate, "provenance", None) is not None
                    and hasattr(gate.provenance, "coverage")
                    else None
                ),
                "missing_fields": (
                    gate.provenance.missing_fields()
                    if getattr(gate, "provenance", None) is not None
                    and hasattr(gate.provenance, "missing_fields")
                    else {}
                ),
                "samples": [
                    {
                        "key": record.key,
                        "label": record.original_label or record.normalized_label,
                        "filename": record.filename,
                        "sheet": record.sheet,
                        "cell": record.cell,
                        "period": record.period,
                        "value": record.normalized_value,
                        "operation": record.operation,
                        "derived": record.derived,
                        "sources": [
                            f"{source.sheet}!{source.cell}" for source in record.sources
                        ],
                    }
                    for record in (
                        getattr(gate.provenance, "important_records", [])
                        if getattr(gate, "provenance", None) is not None
                        else []
                    )[:25]
                ],
            }
            for stage in getattr(gate, "stages", []) or []:
                status = _status_value(getattr(stage, "status", ""))
                if status.upper() != "PASS":
                    report.warnings.append(
                        f"{getattr(getattr(stage, 'stage', ''), 'label_fa', '')}: {status} — "
                        f"{getattr(stage, 'summary', '')}"
                    )

        template_info = getattr(getattr(result, "report", None), "template", None)
        report.outputs = {
            "path": str(getattr(result, "output_path", "") or ""),
            "notes": list(getattr(getattr(result, "report", None), "notes", []) or []),
            "template": (
                str(getattr(template_info, "path", "") or "")
                if template_info is not None
                else ""
            ),
            "format": (
                getattr(getattr(result, "report", None), "output_format", None).value
                if getattr(result, "report", None) is not None
                and getattr(getattr(result, "report", None), "output_format", None) is not None
                else ""
            ),
            "writer": getattr(getattr(result, "report", None), "writer", ""),
            "mode": getattr(getattr(result, "report", None), "mode", ""),
            "tables": getattr(getattr(result, "report", None), "tables_written", 0),
            "sections": list(getattr(getattr(result, "report", None), "sections_written", ()) or ()),
        }
        report.warnings.extend(getattr(result, "warnings", []) or [])
        report.limitations.extend(getattr(result, "warnings", []) or [])

        output_path = Path(report.outputs.get("path") or "")
        validation = getattr(getattr(result, "report", None), "validation", None)
        docx_details: dict[str, Any] = {
            "exists": output_path.is_file(),
            "size_bytes": output_path.stat().st_size if output_path.is_file() else 0,
        }
        if validation is not None:
            docx_details["writer_validation"] = (
                validation.to_dict() if hasattr(validation, "to_dict") else validation
            )
            docx_details["valid"] = bool(getattr(validation, "valid", False))
        if output_path.is_file() and output_path.suffix.lower() == ".docx":
            try:
                reopened = validate_docx_output(
                    output_path,
                    None,
                    require_rtl=True,
                    require_page_numbers=True,
                )
                docx_details["reopen"] = reopened.to_dict()
                docx_details["reopen_valid"] = reopened.valid
                docx_details.setdefault("valid", reopened.valid)
            except Exception as exc:  # noqa: BLE001 - invalid package is a finding
                docx_details["reopen"] = {"valid": False, "errors": [str(exc)]}
                docx_details["reopen_valid"] = False
                docx_details["valid"] = False
        report.docx = docx_details
        if not output_path.is_file():
            report.outputs["path"] = report.outputs.get("path") or str(output_path)

    def _match_indicators_across_sources(
        self, report: AcceptanceReport, result: Any
    ) -> None:
        grouped: dict[str, dict[str, Any]] = {}
        for outcome in getattr(result, "files", []) or []:
            analysis = getattr(outcome, "analysis", None)
            if analysis is None:
                continue
            source = getattr(outcome, "name", "")
            report_ind = build_indicator_periods(analysis, source=source)
            for observation in report_ind.observations:
                key = normalize_text(normalize_digits(observation.indicator))
                if not key:
                    continue
                entry = grouped.setdefault(
                    key,
                    {
                        "indicator": observation.indicator,
                        "unit": observation.unit,
                        "units": set(),
                        "role": observation.role,
                        "sources": {},
                    },
                )
                if observation.unit:
                    entry["units"].add(observation.unit)
                source_entry = entry["sources"].setdefault(
                    source, {"periods": [], "values": [], "addresses": []}
                )
                if observation.period not in source_entry["periods"]:
                    source_entry["periods"].append(observation.period)
                source_entry["values"].append(
                    {
                        "period": observation.period,
                        "value": _text(observation.value),
                        "address": f"{observation.sheet}!{observation.address}",
                    }
                )
                source_entry["addresses"].append(
                    f"{observation.sheet}!{observation.address}"
                )
        matches: list[dict[str, Any]] = []
        for entry in grouped.values():
            if len(entry["sources"]) < 2:
                continue
            units = sorted(entry["units"])
            ambiguous = len(units) > 1
            matches.append(
                {
                    "indicator": entry["indicator"],
                    "unit": entry["unit"],
                    "role": entry["role"],
                    "sources": sorted(entry["sources"]),
                    "periods": {
                        source: sorted(values["periods"])
                        for source, values in sorted(entry["sources"].items())
                    },
                    "values": {
                        source: values["values"]
                        for source, values in sorted(entry["sources"].items())
                    },
                    "status": "ambiguous" if ambiguous else "matched",
                    "comparable": False,
                    "reason": (
                        "برچسب یکسان است اما واحدها متفاوت‌اند "
                        f"({'، '.join(units)}); برابری شاخص اثبات نمی‌شود."
                        if ambiguous
                        else (
                            "برچسب و واحد یکسان است؛ مقایسه عددی بین منابع فقط با "
                            "دوره مشترک و کنترل تکرار مجاز است."
                        )
                    ),
                }
            )
            if ambiguous:
                report.ambiguous.append(
                    {
                        "label": entry["indicator"],
                        "status": "ambiguous",
                        "members": sorted(entry["sources"]),
                        "reason": (
                            "برچسب یکسان با واحدهای متفاوت: " + "، ".join(units) +
                            "؛ تطبیق قطعی نیست."
                        ),
                    }
                )
        report.cross_source_matches = sorted(
            matches, key=lambda item: (item["role"] != "total", item["indicator"])
        )

    # ------------------------------------------------------------------ #
    def _checkpoints(
        self,
        report: AcceptanceReport,
        result: Any,
        checkpoints: Iterable[AcceptanceCheckpoint],
    ) -> None:
        requested = list(checkpoints)
        if not requested:
            report.limitations.append(
                "هیچ checkpoint تاریخی ارائه نشد؛ تأیید عددی بیرونی انجام نشد و "
                "اتکا فقط بر ردیابی سلول منبع است."
            )
            return
        indicators = getattr(result, "indicator_periods", None)
        if indicators is None:
            report.checkpoints = _unverified_checkpoints(requested)
            report.warnings.append(
                "مقایسه شاخص‌های ردیفی ساخته نشد؛ checkpoint ها قابل بررسی نیستند."
            )
            return
        for checkpoint in requested:
            candidates = [checkpoint.label, *checkpoint.aliases]
            exact_candidates: list[Any] = []
            for candidate in candidates:
                if not candidate:
                    continue
                exact_candidates.extend(indicators.exact_comparisons(candidate))
            deduped: list[Any] = []
            for item in exact_candidates:
                if all(item is not seen for seen in deduped):
                    deduped.append(item)
            exact_candidates = deduped
            comparison = exact_candidates[0] if len(exact_candidates) == 1 else None
            cross_source = False
            if len(exact_candidates) > 1:
                for candidate in candidates:
                    if not candidate:
                        continue
                    found = indicators.cross_source_comparison(candidate)
                    if found is not None:
                        comparison = found
                        cross_source = True
                        break
                if comparison is None:
                    report.checkpoints.append(
                        CheckpointResult(
                            label=checkpoint.label,
                            status=CHECK_UNVERIFIED,
                            expected_previous=checkpoint.previous,
                            expected_current=checkpoint.current,
                            expected_percent=checkpoint.reported_percent,
                            note=(
                                f"این برچسب در {len(exact_candidates)} فایل/منبع تکرار شده "
                                "است و یک جفت دوره یکتا بین دو فایل وجود ندارد؛ "
                                "برای تأیید، منبع یا برچسب یکتا لازم است."
                            ),
                        )
                    )
                    continue
            if comparison is None:
                for candidate in candidates:
                    if not candidate:
                        continue
                    found = indicators.comparison(candidate)
                    if found is not None:
                        comparison = found
                        break
            item = CheckpointResult(
                label=checkpoint.label,
                status=CHECK_UNVERIFIED,
                expected_previous=checkpoint.previous,
                expected_current=checkpoint.current,
                expected_percent=checkpoint.reported_percent,
                note="برچسب checkpoint با هیچ شاخص استخراج‌شده تطبیق داده نشد؛ مقدار حدس زده نشد.",
            )
            if comparison is None:
                report.checkpoints.append(item)
                continue
            comparison_labels = {
                normalize_text(normalize_digits(comparison.indicator)),
                normalize_text(normalize_digits(comparison.row_label)),
            }
            exact_label = any(
                normalize_text(normalize_digits(candidate)) in comparison_labels
                for candidate in candidates
                if candidate
            )
            item.matched_indicator = comparison.indicator
            item.extracted_previous = comparison.previous
            item.extracted_current = comparison.current
            item.extracted_percent = comparison.percent
            item.previous_address = comparison.previous_address
            item.current_address = comparison.current_address
            mismatches = []
            if checkpoint.previous is not None and comparison.previous != checkpoint.previous:
                mismatches.append(
                    f"مقدار دوره قبل: انتظار {checkpoint.previous}، استخراج {comparison.previous}"
                )
            if checkpoint.current is not None and comparison.current != checkpoint.current:
                mismatches.append(
                    f"مقدار دوره جاری: انتظار {checkpoint.current}، استخراج {comparison.current}"
                )
            if (
                checkpoint.reported_percent is not None
                and comparison.percent is not None
                and abs(checkpoint.reported_percent - comparison.percent) > Decimal("0.01")
            ):
                mismatches.append(
                    "درصد تغییر: انتظار "
                    f"{checkpoint.reported_percent}، محاسبه {comparison.percent}"
                )
            if mismatches:
                item.status = CHECK_FAIL
                item.note = "اختلاف واقعی ثبت شد و برای بررسی گزارش می‌شود: " + "؛ ".join(mismatches)
            elif not exact_label:
                item.status = CHECK_UNVERIFIED
                item.note = (
                    f"برچسب checkpoint دقیقاً با «{comparison.indicator}» یکی نیست؛ "
                    "این تطبیق نزدیک برای تأیید کافی نیست. برچسب دقیق یا alias بدهید."
                )
            else:
                item.status = CHECK_PASS
                item.note = "مقدار استخراج‌شده با checkpoint هم‌خوان است."
                if cross_source:
                    item.note += (
                        f" (مقایسه بین منابع: {comparison.source}؛ "
                        f"{comparison.previous_period}→{comparison.current_period})"
                    )
                    limitation = (
                        "تأیید checkpoint از دو فایل انجام شد؛ برابری دامنه، تعریف شاخص "
                        "و پوشش دوره بین دو فایل باید توسط کاربر تأیید شود."
                    )
                    if limitation not in report.limitations:
                        report.limitations.append(limitation)
            report.checkpoints.append(item)

    # ------------------------------------------------------------------ #
    def _false_success(
        self,
        report: AcceptanceReport,
        result: Any,
        requested_format: str,
    ) -> None:
        add = report.checks.append
        missing_inputs = [item for item in report.inputs if not item.exists]
        missing_excel = [item for item in missing_inputs if item.kind == "excel"]
        if missing_inputs:
            add(
                FalseSuccessCheck(
                    "source_file_skipped",
                    CHECK_FAIL,
                    "برخی فایل‌های منبع موجود نیستند؛ پردازش کامل و پذیرش سرتاسری ممکن نیست.",
                    {
                        "missing": [item.path for item in missing_inputs],
                        "missing_excel": [item.path for item in missing_excel],
                    },
                )
            )
        else:
            add(
                FalseSuccessCheck(
                    "source_file_skipped",
                    CHECK_PASS,
                    "همه فایل‌های منبع اعلام‌شده موجود بودند.",
                )
            )

        failed_files = [
            item for item in getattr(result, "files", []) or []
            if getattr(item, "status", "") != "تحلیل شد"
        ]
        if failed_files:
            add(
                FalseSuccessCheck(
                    "file_processing_failed",
                    CHECK_FAIL,
                    "دست‌کم یک فایل منبع تحلیل نشد؛ گزارش نمی‌تواند موفق تلقی شود.",
                    {
                        "files": [
                            {
                                "name": getattr(item, "name", ""),
                                "status": getattr(item, "status", ""),
                                "error": getattr(item, "error", ""),
                            }
                            for item in failed_files
                        ]
                    },
                )
            )
        else:
            add(
                FalseSuccessCheck(
                    "file_processing_failed",
                    CHECK_PASS,
                    "همه فایل‌های منبع با موفقیت تحلیل شدند.",
                )
            )

        ignored = [
            item for item in report.blocks if getattr(item, "get", lambda *_: None)("cells_truncated")
        ]
        blocked = list(
            getattr(getattr(result, "indicator_periods", None), "blocked_blocks", []) or []
        )
        if ignored or blocked:
            add(
                FalseSuccessCheck(
                    "data_block_ignored",
                    CHECK_WARNING,
                    "بخشی از سلول‌ها یا یک بلوک قابل انتساب به شاخص نبود؛ در محدودیت‌ها ثبت شد.",
                    {"truncated_blocks": ignored, "blocked_blocks": blocked},
                )
            )
        elif not report.blocks:
            add(
                FalseSuccessCheck(
                    "data_block_ignored",
                    CHECK_FAIL,
                    "هیچ بلوک داده‌ای استخراج نشد؛ گزارش تحلیلی معتبر نیست.",
                )
            )
        else:
            add(
                FalseSuccessCheck(
                    "data_block_ignored",
                    CHECK_PASS,
                    f"{len(report.blocks)} بلوک داده استخراج شد.",
                )
            )

        gate = getattr(result, "quality_gate", None)
        period_stage = None
        if gate is not None:
            for stage in getattr(gate, "stages", []) or []:
                name = getattr(getattr(stage, "stage", ""), "name", "")
                if name == "PERIOD_IDENTIFICATION":
                    period_stage = stage
                    break
        period_status = _status_value(getattr(period_stage, "status", "")) if period_stage else ""
        comparisons = getattr(
            getattr(result, "indicator_periods", None), "comparisons", []
        ) or []
        periods = sorted(
            {
                period
                for item in report.periods
                for period in item.get("periods", []) or []
            }
        )
        if period_status.upper() == "FAIL" or (
            not periods and not getattr(getattr(result, "indicator_periods", None), "periods", [])
        ):
            add(
                FalseSuccessCheck(
                    "period_not_identified",
                    CHECK_FAIL,
                    "دوره گزارش‌دهی به‌طور قابل اتکا شناسایی نشد؛ مقایسه دوره‌ای معتبر نیست.",
                    {"period_stage": period_status, "periods": periods},
                )
            )
        elif not comparisons:
            add(
                FalseSuccessCheck(
                    "period_not_identified",
                    CHECK_WARNING,
                    "دوره‌ها شناسایی شدند اما مقایسه دو دوره برای هیچ شاخصی ساخته نشد.",
                    {"periods": periods},
                )
            )
        else:
            add(
                FalseSuccessCheck(
                    "period_not_identified",
                    CHECK_PASS,
                    f"{len(comparisons)} مقایسه شاخص در دو دوره ساخته شد.",
                    {"periods": periods},
                )
            )

        if gate is None:
            add(
                FalseSuccessCheck(
                    "critical_validation_failed",
                    CHECK_FAIL,
                    "دروازه کیفیت اجرا نشد؛ اعتبار گزارش اثبات نشده است.",
                )
            )
        elif getattr(gate, "blocked", False):
            add(
                FalseSuccessCheck(
                    "critical_validation_failed",
                    CHECK_FAIL,
                    "دروازه کیفیت یک شکست مسدودکننده ثبت کرده است.",
                    {
                        "blocking": [
                            getattr(getattr(item, "stage", ""), "label_fa", "")
                            for item in getattr(gate, "blocking_failures", []) or []
                        ]
                    },
                )
            )
        elif _status_value(getattr(gate, "status", "")).upper() == "FAIL":
            add(
                FalseSuccessCheck(
                    "critical_validation_failed",
                    CHECK_FAIL,
                    "دست‌کم یک مرحله اعتبارسنجی وضعیت «رد» دارد؛ پذیرش کامل ممکن نیست.",
                    {
                        "failed_stages": [
                            getattr(getattr(item, "stage", ""), "label_fa", "")
                            for item in getattr(gate, "stages", []) or []
                            if _status_value(getattr(item, "status", "")).upper() == "FAIL"
                        ]
                    },
                )
            )
        elif _status_value(getattr(gate, "status", "")).upper() == "WARNING":
            add(
                FalseSuccessCheck(
                    "critical_validation_failed",
                    CHECK_WARNING,
                    "اعتبارسنجی با هشدار تمام شد؛ موفقیت کامل گزارش نمی‌شود.",
                )
            )
        else:
            add(
                FalseSuccessCheck(
                    "critical_validation_failed",
                    CHECK_PASS,
                    "همه مراحل اعتبارسنجی قبول شدند.",
                )
            )

        generation = getattr(result, "report", None)
        output_format = getattr(generation, "output_format", None)
        output_path = str(getattr(generation, "output_path", "") or "")
        expected = OutputFormat.DOCX
        format_ok = (
            generation is not None
            and output_format is expected
            and requested_format.lower() == expected.value
            and Path(output_path).suffix.lower() == expected.extension
        )
        add(
            FalseSuccessCheck(
                "wrong_output_format",
                CHECK_PASS if format_ok else CHECK_FAIL,
                (
                    f"قالب خروجی {getattr(output_format, 'value', '')} با پسوند "
                    f"{Path(output_path).suffix} درست است."
                    if format_ok
                    else (
                        "پذیرش سرتاسری فقط خروجی DOCX را می‌پذیرد؛ قالب یا پسوند "
                        "خروجی با این قرارداد هم‌خوان نیست."
                    )
                ),
                {
                    "requested": requested_format,
                    "actual": getattr(output_format, "value", None),
                    "path": output_path,
                },
            )
        )

        validation = report.docx.get("writer_validation") or {}
        errors = list(validation.get("errors", []) or [])
        reopen = report.docx.get("reopen") or {}
        errors.extend(item for item in reopen.get("errors", []) or [] if item not in errors)
        if not validation:
            errors.append(
                "اعتبارسنجی محتوای نویسنده در دسترس نیست؛ ادعای عددی سند تأیید نشده است."
            )
        valid = bool(report.docx.get("valid")) and not errors
        add(
            FalseSuccessCheck(
                "invalid_docx",
                CHECK_PASS if valid else CHECK_FAIL,
                (
                    "فایل Word از نظر بسته، ساختار، RTL و شماره صفحه معتبر است."
                    if valid
                    else "فایل Word معتبر نیست: " + "؛ ".join(errors[:5])
                ),
                {"errors": errors},
            )
        )

        unsupported = list(validation.get("unsupported_numbers", []) or [])
        add(
            FalseSuccessCheck(
                "unverified_numeric_claims",
                CHECK_PASS if not unsupported else CHECK_FAIL,
                (
                    "همه اعداد سند در محتوای اعتبارسنجی‌شده گزارش موجودند."
                    if not unsupported
                    else "اعدادی در سند هست که در محتوای معتبر گزارش نیستند: "
                    + "، ".join(unsupported[:10])
                ),
                {"unsupported_numbers": unsupported},
            )
        )

        provenance = report.provenance or {}
        coverage = provenance.get("coverage")
        missing = provenance.get("missing_fields") or {}
        provenance_ok = provenance.get("status", "").upper() == "PASS" and not missing
        add(
            FalseSuccessCheck(
                "provenance_incomplete",
                CHECK_PASS if provenance_ok else CHECK_FAIL,
                (
                    "هر رقم مهم به سلول یا محدوده منبع متصل است."
                    if provenance_ok
                    else "ردیابی منبع کامل نیست؛ پوشش یا فیلدهای الزامی ناقص است."
                ),
                {"coverage": coverage, "missing_fields": missing},
            )
        )

        important = int((provenance.get("counts") or {}).get("important", 0) or 0)
        add(
            FalseSuccessCheck(
                "no_verified_figures",
                CHECK_PASS if important else CHECK_FAIL,
                (
                    f"{important} رقم مهم با شاهد منبع در گزارش ثبت شد."
                    if important
                    else "هیچ رقم مهمی با شاهد منبع ثبت نشد؛ گزارش تحلیلی قابل ارزیابی نیست."
                ),
                {"important_records": important},
            )
        )

        checkpoint_failures = [item for item in report.checkpoints if item.status == CHECK_FAIL]
        checkpoint_unverified = [
            item for item in report.checkpoints if item.status == CHECK_UNVERIFIED
        ]
        if checkpoint_failures:
            add(
                FalseSuccessCheck(
                    "checkpoint_mismatch",
                    CHECK_FAIL,
                    "دست‌کم یک checkpoint با مقدار مستقل استخراج‌شده اختلاف دارد؛ اختلاف گزارش شده است.",
                    {"labels": [item.label for item in checkpoint_failures]},
                )
            )
        elif checkpoint_unverified:
            add(
                FalseSuccessCheck(
                    "checkpoint_mismatch",
                    CHECK_UNVERIFIED,
                    "برچسب برخی checkpoint ها به شاخص منبع متصل نشد؛ تأیید عددی انجام نشد.",
                    {"labels": [item.label for item in checkpoint_unverified]},
                )
            )
        elif report.checkpoints:
            add(
                FalseSuccessCheck(
                    "checkpoint_mismatch",
                    CHECK_PASS,
                    "همه checkpoint ها با مقدار مستقل استخراج‌شده هم‌خوان‌اند.",
                )
            )

        if report.ambiguous:
            add(
                FalseSuccessCheck(
                    "ambiguous_matches_present",
                    CHECK_WARNING,
                    f"{len(report.ambiguous)} تطبیق مبهم بین منابع ثبت شد و به‌عنوان محدودیت گزارش می‌شود.",
                )
            )

    # ------------------------------------------------------------------ #
    def _final_status(self, report: AcceptanceReport) -> None:
        failures = report.failed_checks
        hard_failures = [item for item in failures if item.code != "source_file_skipped"]
        unverified = report.unverified_checks
        warning = report.warning_checks
        missing_inputs = any(
            item.kind == "excel" and not item.exists for item in report.inputs
        )
        if hard_failures:
            report.status = STATUS_FAILED
        elif failures or missing_inputs:
            report.status = STATUS_PARTIAL
        elif not report.docx.get("valid") or unverified:
            report.status = STATUS_UNVERIFIED if not report.docx.get("valid") else STATUS_PARTIAL
        elif warning:
            report.status = STATUS_PARTIAL
        else:
            report.status = STATUS_ACCEPTED
        report.status_label = STATUS_FA[report.status]
        if report.status != STATUS_ACCEPTED:
            report.limitations.append(
                "این اجرا «موفق کامل» اعلام نمی‌شود؛ دلیل در بخش بررسی‌های جلوگیری از موفقیت کاذب آمده است."
            )


# ---------------------------------------------------------------------- #
def load_checkpoints(path: Path | str) -> list[AcceptanceCheckpoint]:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    items = payload.get("checkpoints", payload if isinstance(payload, list) else [])
    return [AcceptanceCheckpoint.from_dict(item) for item in items]


def _unverified_checkpoints(
    checkpoints: Iterable[AcceptanceCheckpoint],
) -> list[CheckpointResult]:
    return [
        CheckpointResult(
            label=item.label,
            status=CHECK_UNVERIFIED,
            expected_previous=item.previous,
            expected_current=item.current,
            expected_percent=item.reported_percent,
            note="فایل منبع موجود نیست؛ مقدار استخراجی برای مقایسه وجود ندارد.",
        )
        for item in checkpoints
    ]


def _inspect_input(path: Path) -> AcceptanceInputFile:
    suffix = path.suffix.lower()
    kind = "excel" if suffix in {".xlsx", ".xlsm"} else "docx" if suffix == ".docx" else "other"
    if not path.is_file():
        return AcceptanceInputFile(
            path=str(path), role="data" if kind == "excel" else "reference",
            exists=False, kind=kind, error="فایل موجود نیست.",
        )
    item = AcceptanceInputFile(
        path=str(path),
        role="data" if kind == "excel" else "reference",
        exists=True,
        kind=kind,
        size_bytes=path.stat().st_size,
    )
    try:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        item.sha256 = digest.hexdigest()
    except OSError as exc:
        item.error = str(exc)
    return item


def _analysis_sheets(analysis: Any) -> list[Any]:
    extraction = getattr(analysis, "extraction", None)
    return list(getattr(extraction, "sheets", []) or []) if extraction is not None else []


def _analysis_blocks(analysis: Any) -> list[Any]:
    return [block for sheet in _analysis_sheets(analysis) for block in getattr(sheet, "blocks", []) or []]


def _period_status_fa(value: str) -> str:
    return {
        "detected": "شناسایی شد",
        "not_detected": "شناسایی نشد",
        "conflicting": "متناقض",
        "partial": "ناقص",
    }.get(str(value or ""), str(value or "—"))


def _status_value(value: Any) -> str:
    return str(getattr(value, "value", value) or "")


def _decimal(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    if isinstance(value, Decimal):
        return value
    text = normalize_digits(str(value)).replace(",", "").replace("٪", "").strip()
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return format(value, "f")
    return str(value)


def _md_table(headers: Sequence[str], rows: Iterable[Sequence[Any]]) -> str:
    lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
    for row in rows:
        cells = [str(cell if cell is not None else "—").replace("|", "/") for cell in row]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def render_markdown(report: AcceptanceReport) -> str:
    """The structured fifteen-part acceptance report in Persian."""

    lines: list[str] = []
    lines.append("# گزارش پذیرش سرتاسری")
    lines.append("")
    lines.append(f"**وضعیت نهایی: {report.status_label or STATUS_FA.get(report.status, report.status)}**")
    lines.append("")
    lines.append(f"- درخواست: {report.request or '—'}")
    lines.append(f"- زمان اجرا: {report.generated_at or '—'}")
    lines.append("")

    lines.append("## ۱. فایل‌های ورودی شناسایی‌شده")
    lines.append(
        _md_table(
            ["فایل", "نوع", "وضعیت", "حجم (بایت)", "SHA-256", "خطا"],
            [
                [
                    Path(item.path).name,
                    item.kind,
                    "موجود" if item.exists else "موجود نیست",
                    item.size_bytes,
                    item.sha256[:16] + ("…" if item.sha256 else ""),
                    item.error or "—",
                ]
                for item in report.inputs
            ],
        )
    )
    lines.append("")

    lines.append("## ۲. فایل‌های منبع پردازش‌شده")
    if report.files:
        lines.append(
            _md_table(
                ["فایل", "وضعیت", "برگه", "بلوک", "ردیف", "ستون", "خطا"],
                [
                    [
                        item["name"], item["status"], item["sheets"], item["blocks"],
                        item["rows"], item["columns"], item["error"] or "—",
                    ]
                    for item in report.files
                ],
            )
        )
    else:
        lines.append("هیچ فایلی پردازش نشد.")
    lines.append("")

    lines.append("## ۳. برگه‌های بازرسی‌شده")
    if report.worksheets:
        lines.append(
            _md_table(
                ["فایل", "برگه", "عنوان", "بلوک", "دوره‌ها", "وضعیت دوره"],
                [
                    [
                        item["file"], item["sheet"], item["title"], item["blocks"],
                        "، ".join(item["periods"]) or "—",
                        _period_status_fa(item["period_status"]),
                    ]
                    for item in report.worksheets
                ],
            )
        )
    else:
        lines.append("برگه‌ای بازرسی نشد.")
    lines.append("")

    lines.append("## ۴. بلوک‌های داده شناسایی‌شده")
    lines.append(f"تعداد: {len(report.blocks)}")
    if report.blocks:
        lines.append(
            _md_table(
                ["فایل", "برگه", "بلوک", "عنوان", "ردیف", "ستون", "دوره‌ها", "برش سلول"],
                [
                    [
                        item["file"], item["sheet"], item["block"], item["title"],
                        item["rows"], item["columns"], "، ".join(item["periods"]) or "—",
                        "بله" if item["cells_truncated"] else "خیر",
                    ]
                    for item in report.blocks
                ],
            )
        )
    lines.append("")

    lines.append("## ۵. بلوک‌های استخراج‌شده موفق")
    truncated = [item for item in report.blocks if item.get("cells_truncated")]
    blocked = list(
        getattr(getattr(report.result, "indicator_periods", None), "blocked_blocks", []) or []
    )
    lines.append(f"{len(report.blocks)} بلوک در مدل استخراج ثبت شد.")
    if truncated or blocked:
        lines.append(
            f"- برش سلول: {len(truncated)} بلوک | بلوک دوره‌ای بدون برچسب شاخص: {len(blocked)} "
            "(به‌عنوان محدودیت یا بررسی هشدار ثبت شده است)."
        )
    lines.append("")

    lines.append("## ۶. دوره‌های شناسایی‌شده")
    if report.periods:
        lines.append(
            _md_table(
                ["فایل", "دوره‌ها", "ماه‌ها", "وضعیت"],
                [
                    [
                        item["file"], "، ".join(item["periods"]) or "—",
                        "، ".join(item["months"]) or "—",
                        "شناسایی شد" if item["status"] == "detected" else "شناسایی نشد",
                    ]
                    for item in report.periods
                ],
            )
        )
    else:
        lines.append("دوره‌ای شناسایی نشد.")
    lines.append("")

    lines.append("## ۷. شاخص‌های تطبیق‌داده‌شده و نقش‌های معنایی")
    counts = (report.matching or {}).get("counts", {})
    lines.append(
        f"- تطبیق قطعی: {counts.get('matched', 0)} | احتمالی: {counts.get('probable', 0)} | "
        f"مبهم ستونی: {counts.get('ambiguous', 0)} | رد‌شده: {counts.get('rejected', 0)} | "
        f"شاخص ردیفی مشترک بین منابع: {counts.get('row_indicators_matched', 0)} | "
        f"شاخص ردیفی مبهم: {counts.get('row_indicators_ambiguous', 0)}"
    )
    lines.append(f"- مقایسه شاخص ردیفی بین دو دوره: {report.matching.get('indicator_comparisons', 0)}")
    if len(report.files) == 1:
        lines.append("- فقط یک فایل منبع تحلیل شد؛ تطبیق بین‌منبعی موضوعیت ندارد.")

    if report.semantics:
        lines.append("")
        lines.append("### نقش‌های معنایی شناسایی‌شده")
        lines.append(
            _md_table(
                ["فایل", "برگه", "بلوک", "ستون", "نقش", "نقش آماری", "واحد", "دوره"],
                [
                    [
                        item["file"], item["sheet"], item["block"], item["header"],
                        item["role_label"] or item["role"], item["statistical_role"] or "—",
                        item["unit"] or "—", item["period"] or "—",
                    ]
                    for item in report.semantics[:40]
                ],
            )
        )
        if len(report.semantics) > 40:
            lines.append(f"- {len(report.semantics) - 40} ردیف دیگر در JSON حسابرسی ثبت شده است.")

    matched = list((report.matching or {}).get("matches", [])) + list(
        (report.matching or {}).get("probable", [])
    )
    if matched:
        lines.append("")
        lines.append("### تطبیق ستون‌های معنایی بین منابع")
        rows = []
        for item in matched[:40]:
            members = []
            for member in item.get("members", []) or []:
                members.append(
                    f"{member.get('source', '')}/{member.get('sheet', '')}/"
                    f"{member.get('block', '')}/«{member.get('header', '')}»"
                )
            rows.append(
                [
                    item.get("status_label", item.get("status", "")),
                    item.get("normalized_label", ""),
                    "؛ ".join(members),
                    item.get("period_relation", "—"),
                    "بله" if item.get("comparable") else "خیر",
                    item.get("reason", "")[:120],
                ]
            )
        lines.append(
            _md_table(
                ["وضعیت", "برچسب", "اعضا", "نسبت دوره", "قابل مقایسه", "دلیل"],
                rows,
            )
        )

    cross_source = list(report.cross_source_matches)
    if cross_source:
        lines.append("")
        lines.append("### شاخص‌های ردیفی مشترک بین منابع")
        lines.append(
            _md_table(
                ["شاخص", "نقش", "منابع", "دوره‌ها", "نشانی نمونه", "وضعیت", "دلیل"],
                [
                    [
                        item["indicator"], "جمع" if item["role"] == "total" else "داده",
                        "، ".join(item["sources"]),
                        "؛ ".join(
                            f"{source}: {'، '.join(periods)}"
                            for source, periods in item["periods"].items()
                        ),
                        "؛ ".join(
                            f"{source}: {values[0].get('address', '')}"
                            for source, values in item["values"].items()
                            if values
                        ),
                        item["status"],
                        item["reason"],
                    ]
                    for item in cross_source[:40]
                ],
            )
        )
        if len(cross_source) > 40:
            lines.append(f"- {len(cross_source) - 40} شاخص مشترک دیگر در JSON ثبت شده است.")
    if report.comparisons:
        lines.append("")
        lines.append("### مقایسه شاخص‌ها در دو دوره")
        lines.append(
            _md_table(
                ["شاخص", "دوره قبل", "دوره جاری", "تغییر", "درصد", "وضعیت", "منبع"],
                [
                    [
                        item["indicator"],
                        f"{item['previous_period']}: {item['previous'] or '—'}",
                        f"{item['current_period']}: {item['current'] or '—'}",
                        item["difference"] or "—",
                        item["percent"] or "—",
                        item["status_label"],
                        f"{item.get('source', '')}/{item['sheet']}!"
                        f"{item['previous_address']} ← {item['current_address']}",
                    ]
                    for item in report.comparisons[:40]
                ],
            )
        )
        if len(report.comparisons) > 40:
            lines.append(f"- {len(report.comparisons) - 40} مقایسه دیگر در JSON حسابرسی ثبت شده است.")
    lines.append("")

    lines.append("## ۸. تطبیق‌های مبهم")
    if report.ambiguous:
        lines.append(
            _md_table(
                ["برچسب", "اعضا", "دلیل"],
                [
                    [item["label"], "؛ ".join(item["members"]), item["reason"]]
                    for item in report.ambiguous
                ],
            )
        )
    else:
        lines.append("تطبیق مبهمی ثبت نشد.")
    lines.append("")

    lines.append("## ۹. داده تکراری یا هم‌پوشان")
    dup_counts = (report.duplicates or {}).get("counts", {})
    lines.append(
        f"- گروه تکرار: {dup_counts.get('duplicate_groups', 0)} | تصمیم مسدود: "
        f"{dup_counts.get('blocked', 0)} | تصمیم نیازمند بررسی: {dup_counts.get('review', 0)} | "
        f"موارد خارج‌شده: {dup_counts.get('exclusions', 0)}"
    )
    groups = (report.duplicates or {}).get("groups", [])
    if groups:
        lines.append("")
        lines.append("### گروه‌های داده تکراری / هم‌پوشان")
        lines.append(
            _md_table(
                ["شناسه", "نوع", "یکسان", "شباهت", "اعضا", "دلیل"],
                [
                    [
                        item.get("group_id", ""), item.get("kind", ""),
                        "بله" if item.get("identical") else "خیر",
                        item.get("similarity", ""),
                        "؛ ".join(item.get("member_labels", []) or []) or "—",
                        item.get("reason", ""),
                    ]
                    for item in groups[:30]
                ],
            )
        )
    decisions = list((report.duplicates or {}).get("decisions", []))
    if decisions:
        lines.append("")
        lines.append("### تصمیم‌های تجمیع بین منابع")
        lines.append(
            _md_table(
                ["نوع", "وضعیت", "ستون‌ها", "منابع", "دلیل"],
                [
                    [
                        item["kind"], item["status"], "، ".join(item["headers"]) or "—",
                        "، ".join(item["sources"]) or "—", item["reason"],
                    ]
                    for item in decisions[:40]
                ],
            )
        )
    exclusions = (report.duplicates or {}).get("exclusions", [])
    if exclusions:
        lines.append("")
        lines.append("### موارد خارج‌شده از تجمیع")
        lines.append(
            _md_table(
                ["فایل", "برگه", "بلوک", "نوع", "نشانی", "دلیل"],
                [
                    [
                        item.get("source", ""), item.get("sheet", ""), item.get("block", ""),
                        item.get("kind", ""), item.get("address", ""), item.get("reason", ""),
                    ]
                    for item in exclusions[:40]
                ],
            )
        )
    record_duplicates = (report.duplicates or {}).get("record_duplicates", [])
    if record_duplicates:
        lines.append("")
        lines.append(f"- تکرار در سطح رکورد: {len(record_duplicates)} مورد (حذف نشده، فقط ثبت شده است).")

    accidents = report.accidents or {}
    if accidents.get("tables"):
        lines.append("")
        lines.append("### تفکیک داده‌های تصادف و تلفات")
        lines.append(
            f"- جدول تصادف: {accidents.get('tables', 0)} | "
            f"شاخص‌های شناسایی‌شده: {len(accidents.get('metrics', []))}"
        )
        if accidents.get("missing"):
            lines.append(
                "- آمار موجود نیست و حدس زده نشد: " + "، ".join(accidents.get("missing", []))
            )
        metrics = list(accidents.get("metrics", []))
        if metrics:
            lines.append(
                _md_table(
                    ["برگه", "شاخص", "نوع", "نشانی"],
                    [
                        [
                            item.get("sheet", ""), item.get("label", ""),
                            item.get("kind_label", item.get("kind", "")), item.get("address", ""),
                        ]
                        for item in metrics[:30]
                    ],
                )
            )
        findings = accidents.get("findings", [])
        if findings:
            lines.append(
                _md_table(
                    ["کد", "وضعیت", "پیام"],
                    [
                        [item.get("code", ""), item.get("status", ""), item.get("message", "")]
                        for item in findings[:30]
                    ],
                )
            )
    lines.append("")

    lines.append("## ۱۰. نتایج تطبیق آماری")
    reconciliation = report.reconciliation or {}
    lines.append(
        f"- نتایج ثبت‌شده: {len(reconciliation.get('results', []))} | رد: "
        f"{reconciliation.get('failed', 0)} | هشدار: {reconciliation.get('warnings', 0)}"
    )
    results = list(reconciliation.get("results", []))
    if results:
        lines.append(
            _md_table(
                ["فایل", "نوع", "شاخص", "انتظار", "محاسبه", "اختلاف", "وضعیت"],
                [
                    [
                        item["file"], item["kind"], item["label"], item["expected"],
                        item["calculated"], item["difference"], item["status"],
                    ]
                    for item in results[:40]
                ],
            )
        )
        if len(results) > 40:
            lines.append(f"- {len(results) - 40} نتیجه دیگر در JSON حسابرسی ثبت شده است.")
    else:
        lines.append("- هیچ رابطه عددی قابل تطبیقی ثبت نشد.")
    lines.append("")

    lines.append("## ۱۱. بررسی‌های ناموفق و هشدارها")
    lines.append(
        _md_table(
            ["کد", "وضعیت", "پیام"],
            [[item.code, item.status, item.message] for item in report.checks],
        )
    )
    if report.warnings:
        lines.append("")
        lines.append("هشدارهای ثبت‌شده:")
        for warning in dict.fromkeys(report.warnings):
            lines.append(f"- {warning}")
    lines.append("")

    lines.append("## ۱۲. پوشش ردیابی منبع")
    provenance = report.provenance or {}
    coverage = provenance.get("coverage")
    coverage_text = f"{float(coverage) * 100:.1f}٪" if isinstance(coverage, (int, float)) else "—"
    counts = provenance.get("counts", {})
    lines.append(
        f"- وضعیت: {provenance.get('status', '—')} | پوشش فیلدهای اختیاری: {coverage_text}"
    )
    lines.append(
        f"- رکوردهای مهم: {counts.get('important', 0)} | دارای منبع: "
        f"{counts.get('with_sources', 0)} | کامل: {counts.get('complete', 0)}"
    )
    missing = provenance.get("missing_fields") or {}
    if missing:
        lines.append(f"- فیلدهای ناقص: {missing}")
    samples = provenance.get("samples") or []
    if samples:
        lines.append(
            _md_table(
                ["شاخص", "فایل", "برگه", "سلول", "دوره", "مقدار", "عملیات", "منبع"],
                [
                    [
                        item.get("label", ""), item.get("filename", ""),
                        item.get("sheet", ""), item.get("cell", ""),
                        item.get("period", "") or "—", item.get("value", "") or "—",
                        ("مشتق" if item.get("derived") else "منبع") +
                        (f" / {item.get('operation')}" if item.get("operation") else ""),
                        "؛ ".join(item.get("sources", []) or []) or "—",
                    ]
                    for item in samples
                ],
            )
        )
        if int(counts.get("important", 0) or 0) > len(samples):
            lines.append(
                f"- {int(counts.get('important', 0)) - len(samples)} رکورد دیگر در JSON حسابرسی ثبت شده است."
            )
    lines.append("")

    lines.append("## ۱۳. فایل‌های خروجی تولیدشده")
    outputs = report.outputs or {}
    lines.append(f"- مسیر: {outputs.get('path', '—')}")
    lines.append(
        f"- قالب: {outputs.get('format', '—')} | نویسنده: {outputs.get('writer', '—')} | "
        f"حالت: {outputs.get('mode', '—')} | جدول‌ها: {outputs.get('tables', 0)}"
    )
    artifacts = report.artifacts or {}
    if artifacts.get("json"):
        lines.append(f"- گزارش حسابرسی JSON: {artifacts.get('json')}")
    if artifacts.get("markdown"):
        lines.append(f"- گزارش حسابرسی Markdown: {artifacts.get('markdown')}")
    if outputs.get("template"):
        lines.append(f"- قالب Word مرجع (فقط چیدمان): {outputs.get('template')}")
    for note in outputs.get("notes", []) or []:
        lines.append(f"- یادداشت قالب: {note}")
    lines.append("")

    lines.append("## ۱۴. اعتبارسنجی DOCX")
    docx = report.docx or {}
    validation = docx.get("writer_validation") or {}
    reopen = docx.get("reopen") or {}
    lines.append(
        f"- معتبر: {'بله' if docx.get('valid') else 'خیر'} | حجم: {docx.get('size_bytes', 0)} بایت | "
        f"بازخوانی مستقل: {'موفق' if docx.get('reopen_valid') else 'ناموفق'}"
    )
    lines.append(
        f"- بخش‌های RTL: {reopen.get('sections_with_rtl', '—')} | جدول RTL: "
        f"{reopen.get('rtl_tables', '—')} | شماره صفحه: "
        f"{'بله' if reopen.get('has_page_number_field') else 'خیر'}"
    )
    if reopen.get("heading_titles"):
        lines.append(f"- سرفصل‌ها: {'، '.join(reopen.get('heading_titles', []))}")
    if validation.get("unsupported_numbers"):
        lines.append(
            f"- اعداد بدون منبع در سند: {validation.get('unsupported_numbers')}"
        )
    lines.append("")

    lines.append("## ۱۵. محدودیت‌های باقی‌مانده")
    limitations = list(dict.fromkeys(report.limitations))
    if limitations:
        for item in limitations:
            lines.append(f"- {item}")
    else:
        lines.append("- محدودیت ثبت‌شده‌ای وجود ندارد.")
    if report.unverified_stages:
        lines.append("")
        lines.append("مراحل تأییدنشده:")
        for item in report.unverified_stages:
            lines.append(f"- {item}")
    if report.checkpoints:
        lines.append("")
        lines.append("### بررسی checkpoint های تاریخی (فقط برای تأیید، نه ورودی گزارش)")
        lines.append(
            _md_table(
                ["برچسب", "وضعیت", "شاخص منبع", "دوره قبل", "دوره جاری", "درصد", "منبع"],
                [
                    [
                        item.label, item.status, item.matched_indicator or "—",
                        _text(item.extracted_previous) or _text(item.expected_previous) or "—",
                        _text(item.extracted_current) or _text(item.expected_current) or "—",
                        _text(item.extracted_percent) or _text(item.expected_percent) or "—",
                        f"{item.previous_address} ← {item.current_address}"
                        if item.previous_address or item.current_address else "—",
                    ]
                    for item in report.checkpoints
                ],
            )
        )
    if report.accidents and report.accidents.get("tables"):
        lines.append("")
        lines.append("### آمار تصادف و تلفات")
        missing = report.accidents.get("missing") or []
        lines.append(f"- وضعیت: {report.accidents.get('status', '—')}")
        if missing:
            lines.append(f"- شاخص‌های موجود نبوده: {'، '.join(missing)}")
    lines.append("")
    return "\n".join(lines)
