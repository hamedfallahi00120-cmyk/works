"""Central quality gate over every processing stage.

The gate evaluates evidence that already exists.  It never repairs, deletes or
recomputes a value.  A stage may report ``PASS``, ``WARNING`` or ``FAIL``; a
failure in a critical stage blocks the claim that a valid final report was
produced, while a warning prevents the run from being described as a complete
success.
"""

from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any

from local_office_ai.quality.models import (
    STAGE_ORDER,
    GateFinding,
    GateStage,
    GateStageResult,
    GateStatus,
    ProvenanceLedger,
    QualityGateReport,
    status_of,
    worst_status,
)
from local_office_ai.quality.provenance import build_provenance_ledger


def _text(value: Any) -> str:
    return "" if value is None else str(value)


def _meta(analysis: Any) -> Any:
    return getattr(analysis, "meta", None)


def _source_path(analysis: Any, source_path: str | Path | None) -> Path | None:
    if source_path:
        return Path(source_path)
    meta = _meta(analysis)
    raw = _text(getattr(meta, "path", ""))
    return Path(raw) if raw else None


def _count(items: Any) -> int:
    try:
        return len(items or [])
    except TypeError:
        return 0


def _format_code(value: Any) -> str:
    return str(getattr(value, "value", value) or "").strip().lower()


def _format_extension(value: Any) -> str:
    extension = _text(getattr(value, "extension", ""))
    if extension:
        return extension
    code = _format_code(value)
    return f".{code}" if code in {"xlsx", "docx", "pdf"} else ""


def _stage(
    stage: GateStage,
    status: GateStatus,
    summary: str,
    *,
    checks: list[GateFinding] | None = None,
    evidence: dict[str, Any] | None = None,
) -> GateStageResult:
    return GateStageResult(
        stage=stage,
        status=status,
        summary=summary,
        checks=list(checks or []),
        evidence=dict(evidence or {}),
    )


def _finding(
    code: str,
    status: GateStatus,
    message: str,
    *,
    location: str = "",
    affected: int = 0,
    blocking: bool = False,
    evidence: dict[str, Any] | None = None,
) -> GateFinding:
    return GateFinding(
        code=code,
        status=status,
        message=message,
        location=location,
        affected=affected,
        blocking=blocking,
        evidence=dict(evidence or {}),
    )


class QualityGate:
    """Evaluate all thirteen stages against existing stage evidence."""

    def evaluate(
        self,
        analysis: Any,
        *,
        source_path: str | Path | None = None,
        specification: Any = None,
        generation: Any = None,
        provenance: Any = None,
    ) -> QualityGateReport:
        ledger = provenance or build_provenance_ledger(analysis, source_path=source_path)
        source = _source_path(analysis, source_path)
        stages = [
            self._source_availability(analysis, source),
            self._file_readability(analysis),
            self._extraction_completeness(analysis),
            self._block_detection(analysis),
            self._period_identification(analysis),
            self._semantic_classification(analysis),
            self._data_quality(analysis),
            self._statistical_validity(analysis),
            self._duplicate_detection(analysis),
            self._cross_source_reconciliation(analysis),
            self._provenance_completeness(analysis, ledger),
        ]
        if specification is not None or generation is not None:
            stages.append(self._output_format(specification, generation))
            stages.append(self._document_integrity(generation))
        return QualityGateReport(
            stages=stages,
            provenance=ledger,
            source=str(source) if source else "",
            output_path=(
                str(getattr(generation, "output_path", "") or "")
                if generation is not None
                else str(getattr(specification, "output_path", "") or "")
            ),
            generated_at=datetime.now().isoformat(timespec="seconds"),
            finalized=generation is not None,
        )

    def evaluate_folder(
        self,
        result: Any,
        *,
        specification: Any = None,
        generation: Any = None,
    ) -> QualityGateReport:
        """Evaluate a multi-file folder run with the same thirteen stages."""

        outcomes = [
            item
            for item in getattr(result, "files", []) or []
            if getattr(item, "analysis", None) is not None
        ]
        named_reports = [
            (
                str(getattr(item, "name", "") or getattr(item, "path", "")),
                self.evaluate(
                    item.analysis,
                    source_path=getattr(item, "path", None),
                    provenance=getattr(item.analysis, "provenance", None),
                ),
            )
            for item in outcomes
        ]

        merged = ProvenanceLedger(
            filename=Path(str(getattr(result, "folder", ""))).name,
            source_path=str(getattr(result, "folder", "") or ""),
        )
        for index, (_, sub_report) in enumerate(named_reports, start=1):
            if sub_report.provenance is None:
                continue
            key_map = {
                record.key: f"{index}.{record.key}"
                for record in sub_report.provenance.records
            }
            for record in sub_report.provenance.records:
                merged.records.append(
                    replace(
                        record,
                        key=key_map[record.key],
                        related_records=[
                            key_map.get(item, item) for item in record.related_records
                        ],
                    )
                )

        from local_office_ai.quality.provenance import add_indicator_provenance

        add_indicator_provenance(merged, getattr(result, "indicator_periods", None))

        stages: list[GateStageResult] = []
        for stage in STAGE_ORDER[:-2]:
            if stage is GateStage.CROSS_SOURCE_RECONCILIATION:
                stages.append(
                    self._cross_source_stage(getattr(result, "duplicate_control", None))
                )
            elif stage is GateStage.PROVENANCE_COMPLETENESS:
                stages.append(
                    self._provenance_stage(
                        merged,
                        has_primary=any(
                            getattr(item.analysis, "primary_table", None) is not None
                            for item in outcomes
                        ),
                    )
                )
            else:
                stages.append(self._merge_stage(stage, named_reports))

        failed_files = [
            item
            for item in getattr(result, "files", []) or []
            if getattr(item, "status", "") != "تحلیل شد"
        ]
        if failed_files:
            source_stage = stages[0]
            findings = list(source_stage.checks)
            findings.extend(
                _finding(
                    "folder-file-failed",
                    GateStatus.WARNING,
                    f"{getattr(item, 'name', '')}: {getattr(item, 'error', 'تحلیل نشد')}",
                    location=str(getattr(item, "path", "") or ""),
                )
                for item in failed_files
            )
            stages[0] = replace(
                source_stage,
                status=worst_status([source_stage.status, GateStatus.WARNING]),
                summary=f"{len(outcomes)} فایل تحلیل شد و {len(failed_files)} فایل ناموفق بود.",
                checks=findings,
            )

        output_path = str(
            getattr(generation, "output_path", "") or getattr(result, "output_path", "") or ""
        )
        if specification is None and generation is None and output_path:
            from local_office_ai.reporting import (
                OutputFormat,
                ReportGenerationResult,
                ReportSpecification,
            )

            specification = ReportSpecification(OutputFormat.XLSX, Path(output_path))
            generation = ReportGenerationResult(
                output_path=Path(output_path),
                output_format=OutputFormat.XLSX,
                writer=OutputFormat.XLSX.writer_name,
                content_validated=True,
            )
        stages.append(self._output_format(specification, generation))
        stages.append(self._document_integrity(generation))

        report = QualityGateReport(
            stages=stages,
            provenance=merged,
            source=str(getattr(result, "folder", "") or ""),
            output_path=output_path,
            generated_at=datetime.now().isoformat(timespec="seconds"),
            finalized=generation is not None,
        )
        result.quality_gate = report
        return report

    def _merge_stage(
        self,
        stage: GateStage,
        named_reports: list[tuple[str, QualityGateReport]],
    ) -> GateStageResult:
        if not named_reports:
            return _stage(
                stage,
                GateStatus.WARNING,
                "هیچ فایل قابل ارزیابی برای این مرحله وجود ندارد.",
            )
        results = [
            (name, report.stage(stage))
            for name, report in named_reports
            if report.stage(stage) is not None
        ]
        status = worst_status(item.status for _, item in results)
        checks: list[GateFinding] = []
        for name, item in results:
            if item.status is GateStatus.PASS:
                continue
            for check in item.checks:
                checks.append(
                    _finding(
                        check.code,
                        check.status,
                        f"{name}: {check.message}",
                        location=check.location,
                        affected=check.affected,
                        blocking=check.blocking,
                        evidence=check.evidence,
                    )
                )
        return _stage(
            stage,
            status,
            f"{len(results)} فایل در این مرحله ارزیابی شد.",
            checks=checks,
            evidence={
                "files": len(results),
                "statuses": {
                    status_value.value: sum(
                        1 for _, item in results if item.status is status_value
                    )
                    for status_value in GateStatus
                },
            },
        )

    # ------------------------------------------------------------------ #
    # 1. Source availability
    # ------------------------------------------------------------------ #
    def _source_availability(self, analysis: Any, source: Path | None) -> GateStageResult:
        meta = _meta(analysis)
        if source is None:
            return _stage(
                GateStage.SOURCE_AVAILABILITY,
                GateStatus.FAIL,
                "مسیر فایل منبع ثبت نشده است.",
                checks=[
                    _finding(
                        "source-path",
                        GateStatus.FAIL,
                        "مسیر فایل منبع در مدل تحلیل وجود ندارد.",
                        blocking=True,
                    )
                ],
                evidence={"path": ""},
            )
        if not source.exists():
            return _stage(
                GateStage.SOURCE_AVAILABILITY,
                GateStatus.FAIL,
                f"فایل منبع پیدا نشد: {source}",
                checks=[
                    _finding(
                        "source-missing",
                        GateStatus.FAIL,
                        "فایل منبع در زمان ارزیابی وجود ندارد.",
                        location=str(source),
                        blocking=True,
                    )
                ],
                evidence={"path": str(source)},
            )
        if not source.is_file():
            return _stage(
                GateStage.SOURCE_AVAILABILITY,
                GateStatus.FAIL,
                "مسیر منبع یک فایل نیست.",
                checks=[
                    _finding(
                        "source-not-file",
                        GateStatus.FAIL,
                        "مسیر منبع به یک فایل معمولی اشاره نمی‌کند.",
                        location=str(source),
                        blocking=True,
                    )
                ],
                evidence={"path": str(source)},
            )
        size = int(getattr(meta, "size_bytes", 0) or 0)
        sha = _text(getattr(meta, "sha256", ""))
        if size <= 0:
            return _stage(
                GateStage.SOURCE_AVAILABILITY,
                GateStatus.FAIL,
                "فایل منبع خالی است.",
                checks=[
                    _finding(
                        "source-empty",
                        GateStatus.FAIL,
                        "حجم فایل منبع صفر است.",
                        location=str(source),
                        blocking=True,
                    )
                ],
                evidence={"path": str(source), "size_bytes": size},
            )
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if not sha:
            status = GateStatus.WARNING
            checks.append(
                _finding(
                    "source-hash",
                    GateStatus.WARNING,
                    "اثر انگشت SHA-256 فایل منبع ثبت نشده است.",
                    location=str(source),
                )
            )
        return _stage(
            GateStage.SOURCE_AVAILABILITY,
            status,
            "فایل منبع موجود و قابل دسترسی است.",
            checks=checks,
            evidence={"path": str(source), "size_bytes": size, "sha256": sha},
        )

    # ------------------------------------------------------------------ #
    # 2. File readability
    # ------------------------------------------------------------------ #
    def _file_readability(self, analysis: Any) -> GateStageResult:
        meta = _meta(analysis)
        sheets = list(getattr(meta, "sheets", []) or []) if meta is not None else []
        if meta is None:
            return _stage(
                GateStage.FILE_READABILITY,
                GateStatus.FAIL,
                "فراداده فایل خوانده نشد.",
                checks=[
                    _finding(
                        "metadata",
                        GateStatus.FAIL,
                        "مدل فایل (فراداده) وجود ندارد.",
                        blocking=True,
                    )
                ],
            )
        if not sheets:
            return _stage(
                GateStage.FILE_READABILITY,
                GateStatus.FAIL,
                "هیچ برگه‌ای خوانده نشد.",
                checks=[
                    _finding(
                        "sheets",
                        GateStatus.FAIL,
                        "فایل باز شد اما هیچ برگه قابل خواندن ندارد.",
                        blocking=True,
                    )
                ],
            )
        error_cells = int(getattr(meta, "total_error_cells", 0) or 0)
        warnings = list(getattr(analysis, "warnings", []) or [])
        read_warnings = [
            item
            for item in warnings
            if any(
                marker in item
                for marker in ("قفل موقت", "ماکرو", "پیوند خارجی", "خوانده می‌شوند")
            )
        ]
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if error_cells:
            status = GateStatus.WARNING
            checks.append(
                _finding(
                    "error-cells",
                    GateStatus.WARNING,
                    f"{error_cells} سلول خطای فرمول در فایل خوانده شد.",
                    affected=error_cells,
                )
            )
        if read_warnings:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "read-warning",
                    GateStatus.WARNING,
                    "هنگام خواندن برگه‌ها هشدار ثبت شد.",
                    affected=len(read_warnings),
                    evidence={"warnings": read_warnings[:20]},
                )
            )
        return _stage(
            GateStage.FILE_READABILITY,
            status,
            f"{len(sheets)} برگه خوانده شد.",
            checks=checks,
            evidence={
                "sheets": len(sheets),
                "error_cells": error_cells,
                "read_warnings": read_warnings[:20],
            },
        )

    # ------------------------------------------------------------------ #
    # 3. Extraction completeness
    # ------------------------------------------------------------------ #
    def _extraction_completeness(self, analysis: Any) -> GateStageResult:
        extraction = getattr(analysis, "extraction", None)
        if extraction is None:
            return _stage(
                GateStage.EXTRACTION_COMPLETENESS,
                GateStatus.FAIL,
                "مدل استخراج ساخته نشد.",
                checks=[
                    _finding(
                        "extraction-missing",
                        GateStatus.FAIL,
                        "هیچ مدل استخراجی برای جدول‌ها وجود ندارد.",
                        blocking=True,
                    )
                ],
            )
        sheets = list(getattr(extraction, "sheets", []) or [])
        meta_sheets = list(getattr(_meta(analysis), "sheets", []) or [])
        if meta_sheets and not sheets:
            return _stage(
                GateStage.EXTRACTION_COMPLETENESS,
                GateStatus.FAIL,
                "مدل استخراج هیچ برگه‌ای ندارد.",
                checks=[
                    _finding(
                        "extraction-sheets",
                        GateStatus.FAIL,
                        "برگه‌های خوانده‌شده در مدل استخراج حاضر نیستند.",
                        blocking=True,
                    )
                ],
                evidence={"workbook_sheets": len(meta_sheets), "extracted_sheets": 0},
            )

        blocks = [block for sheet in sheets for block in getattr(sheet, "blocks", [])]
        truncated = [block for block in blocks if getattr(block, "cells_truncated", False)]
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if not blocks:
            status = GateStatus.WARNING
            checks.append(
                _finding(
                    "no-block",
                    GateStatus.WARNING,
                    "هیچ بلوک داده‌ای استخراج نشد؛ گزارش فقط ساختار فایل را نشان می‌دهد.",
                )
            )
        if truncated:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "cells-truncated",
                    GateStatus.WARNING,
                    f"{len(truncated)} بلوک به سقف ذخیره سلول رسیده است.",
                    affected=len(truncated),
                    evidence={"blocks": [getattr(item, "block_id", "") for item in truncated]},
                )
            )
        missing_sheets = sorted(
            {getattr(item, "title", "") for item in meta_sheets}
            - {getattr(item, "name", "") for item in sheets}
        )
        if missing_sheets:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "missing-sheet",
                    GateStatus.WARNING,
                    f"{len(missing_sheets)} برگه در مدل استخراج غایب است.",
                    affected=len(missing_sheets),
                    evidence={"sheets": missing_sheets},
                )
            )
        return _stage(
            GateStage.EXTRACTION_COMPLETENESS,
            status,
            f"{len(sheets)} برگه و {len(blocks)} بلوک استخراج شد.",
            checks=checks,
            evidence={
                "sheets": len(sheets),
                "blocks": len(blocks),
                "truncated_blocks": len(truncated),
                "missing_sheets": missing_sheets,
            },
        )

    # ------------------------------------------------------------------ #
    # 4. Block detection
    # ------------------------------------------------------------------ #
    def _block_detection(self, analysis: Any) -> GateStageResult:
        extraction = getattr(analysis, "extraction", None)
        if extraction is None:
            return _stage(
                GateStage.BLOCK_DETECTION,
                GateStatus.FAIL,
                "برای تشخیص بلوک، مدل استخراج لازم است.",
                checks=[
                    _finding(
                        "detection-unavailable",
                        GateStatus.FAIL,
                        "مدل استخراج موجود نیست.",
                        blocking=True,
                    )
                ],
            )
        intelligence = getattr(analysis, "intelligence", None)
        if intelligence is None:
            return _stage(
                GateStage.BLOCK_DETECTION,
                GateStatus.FAIL,
                "هوش ساختاری فایل ساخته نشد.",
                checks=[
                    _finding(
                        "intelligence-missing",
                        GateStatus.FAIL,
                        "خروجی تحلیل ساختار برگه‌ها وجود ندارد.",
                        blocking=True,
                    )
                ],
            )
        blocks = [
            block
            for sheet in getattr(extraction, "sheets", []) or []
            for block in getattr(sheet, "blocks", []) or []
        ]
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if not blocks:
            status = GateStatus.WARNING
            checks.append(
                _finding(
                    "blocks-none",
                    GateStatus.WARNING,
                    "بلوک داده‌ای تشخیص داده نشد.",
                )
            )
        incomplete = [
            block
            for block in blocks
            if not getattr(block, "boundaries", None)
            or not getattr(block, "header_rows", None)
            or not getattr(block, "columns", None)
        ]
        if incomplete:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "blocks-incomplete",
                    GateStatus.WARNING,
                    f"{len(incomplete)} بلوک مرز، سرستون یا ستون کامل ندارد.",
                    affected=len(incomplete),
                )
            )
        primary = getattr(analysis, "primary_table", None)
        return _stage(
            GateStage.BLOCK_DETECTION,
            status,
            f"{len(blocks)} بلوک با مرز مشخص تشخیص داده شد.",
            checks=checks,
            evidence={
                "blocks": len(blocks),
                "incomplete": len(incomplete),
                "primary": _text(getattr(primary, "address", "")) if primary else "",
            },
        )

    # ------------------------------------------------------------------ #
    # 5. Period identification
    # ------------------------------------------------------------------ #
    def _period_identification(self, analysis: Any) -> GateStageResult:
        extraction = getattr(analysis, "extraction", None)
        if extraction is None:
            return _stage(
                GateStage.PERIOD_IDENTIFICATION,
                GateStatus.WARNING,
                "شواهد دوره بدون مدل استخراج قابل ارزیابی نیست.",
            )
        statuses: list[tuple[str, str]] = [
            ("workbook", _text(getattr(extraction, "period_status", "")))
        ]
        for sheet in getattr(extraction, "sheets", []) or []:
            statuses.append((_text(getattr(sheet, "name", "")), _text(getattr(sheet, "period_status", ""))))
            for block in getattr(sheet, "blocks", []) or []:
                statuses.append(
                    (
                        _text(getattr(block, "block_id", "")),
                        _text(getattr(block, "period_status", "")),
                    )
                )
        conflicts = [(name, value) for name, value in statuses if value == "conflicting"]
        detected = [name for name, value in statuses if value == "detected"]
        integration = getattr(analysis, "integration", None)
        if conflicts:
            return _stage(
                GateStage.PERIOD_IDENTIFICATION,
                GateStatus.FAIL,
                f"{len(conflicts)} منبع دوره متناقض دارد.",
                checks=[
                    _finding(
                        "period-conflict",
                        GateStatus.FAIL,
                        "شواهد دوره با یکدیگر هم‌خوان نیستند و هیچ منبعی بر دیگری ترجیح داده نشد.",
                        affected=len(conflicts),
                        evidence={"sources": [name for name, _ in conflicts][:20]},
                    )
                ],
                evidence={"conflicts": [name for name, _ in conflicts], "detected": detected},
            )
        if detected:
            checks: list[GateFinding] = []
            status = GateStatus.PASS
            if integration is not None and int(getattr(integration, "mismatched_cells", 0) or 0) > 0:
                status = GateStatus.WARNING
                checks.append(
                    _finding(
                        "period-verification",
                        GateStatus.WARNING,
                        f"{int(getattr(integration, 'mismatched_cells', 0) or 0)} سلول دوره با جمع اعلامی هم‌خوان نیست.",
                        affected=int(getattr(integration, "mismatched_cells", 0) or 0),
                    )
                )
            return _stage(
                GateStage.PERIOD_IDENTIFICATION,
                status,
                f"{len(detected)} منبع دوره شناسایی شد.",
                checks=checks,
                evidence={"periods": list(getattr(extraction, "periods", []) or []), "detected": detected},
            )
        required = integration is not None and bool(getattr(integration, "series", None))
        if required:
            return _stage(
                GateStage.PERIOD_IDENTIFICATION,
                GateStatus.WARNING,
                "دوره ماهانه لازم است اما شواهد کافی برای شناسایی آن پیدا نشد.",
                checks=[
                    _finding(
                        "period-missing",
                        GateStatus.WARNING,
                        "ماتریس دوره وجود دارد ولی برچسب دوره معتبر شناسایی نشد.",
                    )
                ],
                evidence={"integration_periods": list(getattr(integration, "periods", []) or [])},
            )
        return _stage(
            GateStage.PERIOD_IDENTIFICATION,
            GateStatus.PASS,
            "این فایل به دوره گزارش نیاز ندارد؛ دوره‌ای جعل نشد.",
            evidence={"periods": [], "required": False},
        )

    # ------------------------------------------------------------------ #
    # 6. Semantic classification
    # ------------------------------------------------------------------ #
    def _semantic_classification(self, analysis: Any) -> GateStageResult:
        semantics = getattr(analysis, "semantics", None)
        if semantics is None:
            return _stage(
                GateStage.SEMANTIC_CLASSIFICATION,
                GateStatus.FAIL,
                "طبقه‌بندی معنایی انجام نشد.",
                checks=[
                    _finding(
                        "semantics-missing",
                        GateStatus.FAIL,
                        "هیچ نقش معنایی برای فیلدها ثبت نشده است.",
                        blocking=True,
                    )
                ],
            )
        blocks = list(getattr(semantics, "blocks", []) or [])
        fields = [field for block in blocks for field in getattr(block, "columns", []) or []]
        unknown = [field for field in fields if _text(getattr(getattr(field, "role", ""), "value", getattr(field, "role", ""))) == "unknown"]
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if not blocks:
            status = GateStatus.WARNING
            checks.append(
                _finding(
                    "semantics-empty",
                    GateStatus.WARNING,
                    "بلوکی برای طبقه‌بندی معنایی وجود نداشت.",
                )
            )
        elif fields and len(unknown) == len(fields):
            status = GateStatus.WARNING
            checks.append(
                _finding(
                    "semantics-all-unknown",
                    GateStatus.WARNING,
                    "همه فیلدها نقش «نامشخص» گرفتند.",
                    affected=len(unknown),
                )
            )
        return _stage(
            GateStage.SEMANTIC_CLASSIFICATION,
            status,
            f"{len(fields)} فیلد طبقه‌بندی شد.",
            checks=checks,
            evidence={
                "blocks": len(blocks),
                "fields": len(fields),
                "unknown": len(unknown),
                "roles": semantics.role_counts() if hasattr(semantics, "role_counts") else {},
            },
        )

    # ------------------------------------------------------------------ #
    # 7. Data quality
    # ------------------------------------------------------------------ #
    def _data_quality(self, analysis: Any) -> GateStageResult:
        quality = getattr(analysis, "quality", None)
        if quality is None:
            return _stage(
                GateStage.DATA_QUALITY,
                GateStatus.FAIL,
                "بررسی کیفیت داده اجرا نشد.",
                checks=[
                    _finding(
                        "quality-missing",
                        GateStatus.FAIL,
                        "هیچ یافته کیفیت داده‌ای ثبت نشده است.",
                    )
                ],
            )
        status = status_of(getattr(quality, "status", None)) or GateStatus.WARNING
        findings = list(getattr(quality, "findings", []) or [])
        checks = [
            _finding(
                item.code,
                status_of(getattr(item, "status", None)) or GateStatus.WARNING,
                item.message,
                location="|".join(
                    part
                    for part in (
                        _text(getattr(item, "sheet", "")),
                        _text(getattr(item, "block", "")),
                        _text(getattr(item, "address", "")),
                    )
                    if part
                ),
                affected=int(getattr(item, "count", 0) or 0),
            )
            for item in findings
            if status_of(getattr(item, "status", None)) is not GateStatus.PASS
        ]
        if not findings:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "quality-empty",
                    GateStatus.WARNING,
                    "هیچ بررسی کیفیت داده‌ای ثبت نشد.",
                )
            )
        legacy = list(getattr(analysis, "validation", []) or [])
        legacy_failures = [
            item
            for item in legacy
            if status_of(getattr(item, "status", None)) is GateStatus.FAIL
        ]
        # The legacy checklist is evidence for this stage. Its structural checks
        # duplicate the dedicated extraction/block stages and keep their own
        # thresholds, so they do not override this stage's semantic quality status.
        return _stage(
            GateStage.DATA_QUALITY,
            status,
            f"کیفیت داده: {getattr(getattr(quality, 'status', ''), 'value', getattr(quality, 'status', ''))}",
            checks=checks,
            evidence={
                "counts": quality.counts() if hasattr(quality, "counts") else {},
                "findings": len(findings),
                "legacy_checks": len(legacy),
                "legacy_failures": [
                    str(getattr(item, "code", "") or "") for item in legacy_failures
                ],
            },
        )

    # ------------------------------------------------------------------ #
    # 8. Statistical validity
    # ------------------------------------------------------------------ #
    def _statistical_validity(self, analysis: Any) -> GateStageResult:
        report = getattr(analysis, "statistical_reconciliation", None)
        if report is None:
            return _stage(
                GateStage.STATISTICAL_VALIDITY,
                GateStatus.WARNING,
                "تطبیق آماری اجرا نشد.",
                checks=[
                    _finding(
                        "statistics-missing",
                        GateStatus.WARNING,
                        "هیچ شواهد تطبیق آماری وجود ندارد.",
                    )
                ],
            )
        reconciliations = list(getattr(report, "reconciliations", []) or [])
        relevant = [
            item
            for item in reconciliations
            if getattr(item, "expected", None) is not None
        ]
        items = (
            relevant
            + list(getattr(report, "comparisons", []) or [])
            + list(getattr(report, "averages", []) or [])
        )
        statuses = [status_of(getattr(item, "status", None)) for item in items]
        warnings = list(getattr(report, "warnings", []) or [])
        audits = [
            item
            for item in getattr(analysis, "audits", []) or []
            if status_of(getattr(item, "status", None)) is GateStatus.FAIL
        ]
        checks: list[GateFinding] = []
        for item in items:
            item_status = status_of(getattr(item, "status", None)) or GateStatus.WARNING
            if item_status is GateStatus.PASS:
                continue
            checks.append(
                _finding(
                    _text(getattr(item, "kind", "reconciliation")),
                    item_status,
                    _text(getattr(item, "explanation", "تطبیق آماری نیازمند بررسی است."))
                    or "تطبیق آماری نیازمند بررسی است.",
                    location=_text(getattr(item, "sheet", "")),
                )
            )
        for audit in audits:
            checks.append(
                _finding(
                    "declared-total-audit",
                    GateStatus.FAIL,
                    f"جمع اعلامی «{_text(getattr(audit, 'label', ''))}» با مقدار محاسبه‌شده برابر نیست.",
                    location=_text(getattr(getattr(audit, "declared_source", None), "address", "")),
                )
            )
        for warning in warnings:
            checks.append(
                _finding(
                    "reconciliation-warning",
                    GateStatus.WARNING,
                    _text(warning),
                )
            )
        if not items and not warnings and not audits:
            return _stage(
                GateStage.STATISTICAL_VALIDITY,
                GateStatus.PASS,
                "رابطه آماری قابل تأییدی در فایل اعلام نشده است؛ مقداری ساخته نشد.",
                evidence={"records": 0},
            )
        status = worst_status(statuses + ([GateStatus.WARNING] if warnings else []) + ([GateStatus.FAIL] if audits else []))
        summary = (
            f"{len(items)} رابطه آماری بررسی شد."
            if items
            else "تطبیق آماری هشدار ثبت کرد."
        )
        return _stage(
            GateStage.STATISTICAL_VALIDITY,
            status,
            summary,
            checks=checks,
            evidence={
                "records": len(items),
                "reconciliations": len(reconciliations),
                "verified_reconciliations": len(relevant),
                "warnings": warnings[:20],
                "audit_mismatches": len(audits),
                "counts": report.counts() if hasattr(report, "counts") else {},
            },
        )

    # ------------------------------------------------------------------ #
    # 9. Duplicate and overlap detection
    # ------------------------------------------------------------------ #
    def _duplicate_detection(self, analysis: Any) -> GateStageResult:
        report = getattr(analysis, "duplicate_control", None)
        if report is None:
            return _stage(
                GateStage.DUPLICATE_DETECTION,
                GateStatus.FAIL,
                "کنترل تکرار و هم‌پوشانی اجرا نشد.",
                checks=[
                    _finding(
                        "duplicate-control-missing",
                        GateStatus.FAIL,
                        "هیچ گزارش کنترل تکرار وجود ندارد.",
                    )
                ],
            )
        groups = list(getattr(report, "duplicate_groups", []) or [])
        relations = list(getattr(report, "category_relations", []) or [])
        blocked = list(getattr(report, "blocked_decisions", []) or [])
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if groups:
            status = GateStatus.WARNING
            checks.append(
                _finding(
                    "duplicate-groups",
                    GateStatus.WARNING,
                    f"{len(groups)} گروه محتوای تکراری/خلاصه کپی‌شده شناسایی شد.",
                    affected=len(groups),
                )
            )
        if blocked:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "blocked-aggregation",
                    GateStatus.WARNING,
                    f"{len(blocked)} تصمیم تجمیع مسدود شده است.",
                    affected=len(blocked),
                )
            )
        return _stage(
            GateStage.DUPLICATE_DETECTION,
            status,
            f"{len(groups)} گروه تکراری و {len(relations)} رابطه دسته ثبت شد.",
            checks=checks,
            evidence=report.counts() if hasattr(report, "counts") else {},
        )

    # ------------------------------------------------------------------ #
    # 10. Cross-source reconciliation
    # ------------------------------------------------------------------ #
    def _cross_source_reconciliation(self, analysis: Any) -> GateStageResult:
        return self._cross_source_stage(
            getattr(analysis, "duplicate_control", None)
        )

    def _cross_source_stage(self, duplicate_report: Any) -> GateStageResult:
        decisions = list(getattr(duplicate_report, "decisions", []) or []) if duplicate_report is not None else []
        cross = [item for item in decisions if getattr(item, "cross_source", False)]
        if not cross:
            return _stage(
                GateStage.CROSS_SOURCE_RECONCILIATION,
                GateStatus.PASS,
                "این تحلیل تک‌منبعی است؛ تجمیع بین‌فایلی انجام نشد.",
                evidence={"cross_source_decisions": 0},
            )
        statuses = [status_of(getattr(item, "status", None)) for item in cross]
        checks: list[GateFinding] = []
        for item in cross:
            item_status = status_of(getattr(item, "status", None)) or GateStatus.WARNING
            if item_status is GateStatus.PASS:
                continue
            checks.append(
                _finding(
                    "cross-source-decision",
                    item_status,
                    _text(getattr(item, "explanation", "تصمیم تجمیع بین منابع نیازمند بررسی است.")),
                    location="، ".join(sorted(str(source) for source in getattr(item, "sources", set()))),
                )
            )
        status = worst_status(statuses)
        if status is GateStatus.FAIL:
            # A blocked cross-file total is refused rather than silently summed.
            checks.append(
                _finding(
                    "cross-source-blocked",
                    GateStatus.FAIL,
                    "تجمیع بین منابع مسدود شده است؛ گزارش جمع بین‌فایلی را ادعا نمی‌کند.",
                    affected=len(cross),
                )
            )
        return _stage(
            GateStage.CROSS_SOURCE_RECONCILIATION,
            status,
            f"{len(cross)} تصمیم تجمیع بین منابع بررسی شد.",
            checks=checks,
            evidence={
                "cross_source_decisions": len(cross),
                "counts": {
                    status_value.value: sum(
                        1
                        for item in cross
                        if status_of(getattr(item, "status", None)) is status_value
                    )
                    for status_value in GateStatus
                },
            },
        )

    # ------------------------------------------------------------------ #
    # 11. Provenance completeness
    # ------------------------------------------------------------------ #
    def _provenance_completeness(self, analysis: Any, ledger: Any) -> GateStageResult:
        return self._provenance_stage(
            ledger,
            has_primary=getattr(analysis, "primary_table", None) is not None,
        )

    def _provenance_stage(self, ledger: Any, *, has_primary: bool) -> GateStageResult:
        important = list(getattr(ledger, "important_records", []) or [])
        if not important:
            if has_primary:
                return _stage(
                    GateStage.PROVENANCE_COMPLETENESS,
                    GateStatus.WARNING,
                    "جدول داده وجود دارد اما هیچ رکورد ردیابی برای مقدار مهم ساخته نشد.",
                    checks=[
                        _finding(
                            "provenance-empty",
                            GateStatus.WARNING,
                            "هیچ مقدار مهمی برای ردیابی ثبت نشد.",
                        )
                    ],
                    evidence=ledger.to_dict() if hasattr(ledger, "to_dict") else {},
                )
            return _stage(
                GateStage.PROVENANCE_COMPLETENESS,
                GateStatus.PASS,
                "مقدار عددی مهمی برای ردیابی وجود ندارد.",
                evidence=ledger.to_dict() if hasattr(ledger, "to_dict") else {},
            )
        status = ledger.status
        missing = ledger.missing_fields() if hasattr(ledger, "missing_fields") else {}
        checks: list[GateFinding] = []
        if status is GateStatus.FAIL:
            for field_name, keys in missing.items():
                checks.append(
                    _finding(
                        f"provenance-missing-{field_name}",
                        GateStatus.FAIL,
                        f"{len(keys)} مقدار مهم فیلد «{field_name}» را ندارد.",
                        affected=len(keys),
                        blocking=True,
                        evidence={"records": keys[:20]},
                    )
                )
            if getattr(ledger, "derived_without_sources", None):
                checks.append(
                    _finding(
                        "provenance-derived-sources",
                        GateStatus.FAIL,
                        "مقدار مشتق بدون منبع سلولی ثبت شده است.",
                        affected=len(ledger.derived_without_sources),
                        blocking=True,
                        evidence={"records": [item.key for item in ledger.derived_without_sources][:20]},
                    )
                )
        return _stage(
            GateStage.PROVENANCE_COMPLETENESS,
            status,
            f"{len(important)} مقدار مهم با شواهد منبع ثبت شد.",
            checks=checks,
            evidence=ledger.counts() if hasattr(ledger, "counts") else {},
        )

    # ------------------------------------------------------------------ #
    # 12. Output-format correctness
    # ------------------------------------------------------------------ #
    def _output_format(self, specification: Any, generation: Any) -> GateStageResult:
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if specification is None and generation is None:
            return _stage(
                GateStage.OUTPUT_FORMAT,
                GateStatus.WARNING,
                "قرارداد قالب خروجی برای ارزیابی موجود نیست.",
            )
        output_format = getattr(specification, "output_format", None) or getattr(
            generation, "output_format", None
        )
        path = Path(
            str(
                getattr(generation, "output_path", "")
                or getattr(specification, "output_path", "")
            )
        )
        if output_format is None:
            return _stage(
                GateStage.OUTPUT_FORMAT,
                GateStatus.FAIL,
                "قالب خروجی تعیین نشد.",
                checks=[
                    _finding(
                        "format-missing",
                        GateStatus.FAIL,
                        "هیچ قالب خروجی معتبری انتخاب نشده است.",
                        blocking=True,
                    )
                ],
            )
        extension = _format_extension(output_format)
        if not getattr(output_format, "supports_generation", True):
            status = GateStatus.FAIL
            checks.append(
                _finding(
                    "format-unsupported",
                    GateStatus.FAIL,
                    f"تولید خروجی {_text(getattr(output_format, 'label_fa', ''))} پشتیبانی نمی‌شود.",
                    blocking=True,
                )
            )
        if extension and path.suffix.lower() != extension.lower():
            status = GateStatus.FAIL
            checks.append(
                _finding(
                    "format-suffix",
                    GateStatus.FAIL,
                    f"پسوند خروجی ({path.suffix}) با قالب انتخاب‌شده ({extension}) هم‌خوان نیست.",
                    location=str(path),
                    blocking=True,
                )
            )
        if generation is not None:
            writer = _text(getattr(generation, "writer", ""))
            expected_writer = _text(getattr(output_format, "writer_name", ""))
            if writer and expected_writer and writer != expected_writer:
                status = GateStatus.FAIL
                checks.append(
                    _finding(
                        "format-writer",
                        GateStatus.FAIL,
                        f"نویسنده اجراشده ({writer}) با نویسنده قالب ({expected_writer}) یکی نیست.",
                        blocking=True,
                    )
                )
            if getattr(generation, "content_validated", True) is False:
                status = GateStatus.FAIL
                checks.append(
                    _finding(
                        "format-content",
                        GateStatus.FAIL,
                        "محتوای گزارش پیش از نوشتن اعتبارسنجی نشد.",
                        blocking=True,
                    )
                )
        if specification is not None:
            template_path = getattr(specification, "template_path", None)
            if template_path is not None and not Path(template_path).is_file():
                status = GateStatus.FAIL
                checks.append(
                    _finding(
                        "format-template",
                        GateStatus.FAIL,
                        "قالب Word انتخاب‌شده یک فایل خواندنی نیست.",
                        location=str(template_path),
                        blocking=True,
                    )
                )
        return _stage(
            GateStage.OUTPUT_FORMAT,
            status,
            f"قالب خروجی {_text(getattr(output_format, 'label_fa', getattr(output_format, 'value', '')))} معتبر است.",
            checks=checks,
            evidence={
                "output_format": _text(getattr(output_format, "value", output_format)),
                "path": str(path),
                "writer": _text(getattr(generation, "writer", "")) if generation else "",
            },
        )

    # ------------------------------------------------------------------ #
    # 13. Final document integrity
    # ------------------------------------------------------------------ #
    def _document_integrity(self, generation: Any) -> GateStageResult:
        if generation is None:
            return _stage(
                GateStage.DOCUMENT_INTEGRITY,
                GateStatus.WARNING,
                "سند نهایی هنوز تولید نشده است.",
            )
        path = Path(str(getattr(generation, "output_path", "") or ""))
        checks: list[GateFinding] = []
        status = GateStatus.PASS
        if not path.exists() or not path.is_file() or path.stat().st_size <= 0:
            return _stage(
                GateStage.DOCUMENT_INTEGRITY,
                GateStatus.FAIL,
                "فایل خروجی ساخته نشد یا خالی است.",
                checks=[
                    _finding(
                        "output-missing",
                        GateStatus.FAIL,
                        "فایل خروجی نهایی موجود نیست.",
                        location=str(path),
                        blocking=True,
                    )
                ],
                evidence={"path": str(path)},
            )
        output_format = _format_code(getattr(generation, "output_format", None))
        details: dict[str, Any] = {
            "path": str(path),
            "size_bytes": path.stat().st_size,
            "output_format": output_format,
        }
        if output_format == "xlsx":
            try:
                from openpyxl import load_workbook

                workbook = load_workbook(path, read_only=True)
                sheet_names = list(workbook.sheetnames)
                workbook.close()
            except Exception as exc:  # noqa: BLE001 - an invalid package is a finding
                return _stage(
                    GateStage.DOCUMENT_INTEGRITY,
                    GateStatus.FAIL,
                    "فایل Excel قابل بازخوانی نیست.",
                    checks=[
                        _finding(
                            "output-unreadable",
                            GateStatus.FAIL,
                            f"بازخوانی فایل Excel ناموفق بود: {exc}",
                            location=str(path),
                            blocking=True,
                        )
                    ],
                    evidence=details,
                )
            details["sheets"] = sheet_names
            if not sheet_names:
                status = GateStatus.FAIL
                checks.append(
                    _finding(
                        "output-empty-workbook",
                        GateStatus.FAIL,
                        "کارپوشه خروجی هیچ برگه‌ای ندارد.",
                        location=str(path),
                        blocking=True,
                    )
                )
        elif output_format == "docx":
            validation = getattr(generation, "validation", None)
            if validation is None:
                from local_office_ai.reporting.docx_validation import validate_docx_output

                validation = validate_docx_output(path)
            details["validation"] = validation.to_dict() if hasattr(validation, "to_dict") else validation
            if not getattr(validation, "valid", False):
                status = GateStatus.FAIL
                checks.append(
                    _finding(
                        "output-invalid-docx",
                        GateStatus.FAIL,
                        "اعتبارسنجی بسته Word ناموفق بود.",
                        location=str(path),
                        blocking=True,
                        evidence={"errors": list(getattr(validation, "errors", []) or [])[:20]},
                    )
                )
            elif list(getattr(validation, "warnings", []) or []):
                status = worst_status([status, GateStatus.WARNING])
                checks.append(
                    _finding(
                        "output-docx-warning",
                        GateStatus.WARNING,
                        "اعتبارسنجی Word هشدار ثبت کرد.",
                        location=str(path),
                        evidence={"warnings": list(getattr(validation, "warnings", []) or [])[:20]},
                    )
                )
        else:
            status = GateStatus.FAIL
            checks.append(
                _finding(
                    "output-format-unknown",
                    GateStatus.FAIL,
                    "قالب فایل خروجی برای بازخوانی پشتیبانی نمی‌شود.",
                    location=str(path),
                    blocking=True,
                )
            )
        generation_warnings = list(getattr(generation, "warnings", []) or [])
        unresolved = list(getattr(generation, "unresolved_placeholders", []) or [])
        if generation_warnings:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "output-warnings",
                    GateStatus.WARNING,
                    f"{len(generation_warnings)} هشدار در ساخت خروجی ثبت شد.",
                    affected=len(generation_warnings),
                    evidence={"warnings": generation_warnings[:20]},
                )
            )
        if unresolved:
            status = worst_status([status, GateStatus.WARNING])
            checks.append(
                _finding(
                    "output-placeholders",
                    GateStatus.WARNING,
                    f"{len(unresolved)} جای‌نگهدار در قالب حل نشد.",
                    affected=len(unresolved),
                    evidence={"placeholders": unresolved[:20]},
                )
            )
        return _stage(
            GateStage.DOCUMENT_INTEGRITY,
            status,
            "سند نهایی بازخوانی و اعتبارسنجی شد.",
            checks=checks,
            evidence=details,
        )
