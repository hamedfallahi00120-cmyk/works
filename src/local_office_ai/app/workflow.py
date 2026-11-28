"""End-to-end workflow: read → inspect → detect → calculate → reconcile → AI → report.

Each stage is timed, each failure is attributed to a stage, and the source file
is verified byte-identical at the end of the run.
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from local_office_ai.ai.analyzer import AIAnalyzer, AIOutcome
from local_office_ai.ai.request_analyzer import RequestAnalyzer
from local_office_ai.core.exceptions import ProcessingAborted
from local_office_ai.core.logging_config import get_logger
from local_office_ai.core.models import (
    AnalysisResult,
    Anomaly,
    AuditStatus,
    ColumnRole,
    DataRole,
    RunResult,
    Severity,
    TableCandidate,
    TableKind,
    ValueKind,
)
from local_office_ai.core.settings import Settings
from local_office_ai.excel.anomaly_detection import AnomalyDetector
from local_office_ai.excel.calculations import FinancialCalculator
from local_office_ai.excel.data_extractor import DataExtractor
from local_office_ai.excel.data_quality import DataQualityAnalyzer
from local_office_ai.excel.duplicate_control import DuplicateControlAnalyzer
from local_office_ai.excel.statistical_reconciliation import StatisticalReconciler
from local_office_ai.excel.extraction_model import ExtractionModelBuilder
from local_office_ai.excel.reader import WorkbookReader
from local_office_ai.excel.reconciliation import Reconciler
from local_office_ai.excel.row_classification import (
    classify_block_rows,
    role_counts,
    summary_rows_from_roles,
)
from local_office_ai.excel.semantic_classification import SemanticClassifier
from local_office_ai.excel.semantic_matching import SemanticMatcher
from local_office_ai.excel.sheet_analyzer import SheetGrid
from local_office_ai.excel.statistical_analysis import StatisticalAnalyzer
from local_office_ai.excel.structure_intelligence import StructureIntelligenceAnalyzer
from local_office_ai.excel.data_qa import answer_request
from local_office_ai.excel.period_integration import PeriodIntegrator
from local_office_ai.excel.table_detector import TableDetector, is_report_sheet
from local_office_ai.excel.validation import DataValidator
from local_office_ai.excel.workbook_inspector import inspect_workbook
from local_office_ai.monitoring.performance_report import build_performance_report
from local_office_ai.monitoring.resource_monitor import ResourceMonitor
from local_office_ai.quality import (
    GateStatus,
    QualityGate,
    build_provenance_ledger,
)
from local_office_ai.reporting import (
    OutputFormat,
    ReportGenerator,
    ReportSpecification,
    resolve_output_intent,
)

LOGGER = get_logger(__name__)


def _decimal_text(value: Any) -> str | None:
    return None if value is None else str(value)


@dataclass
class WorkflowOptions:
    source: Path
    model: str | None = None
    use_ai: bool = True
    task: str = "general_analysis"
    question: str | None = None
    goal: str | None = None
    request: str | None = None
    output_path: Path | None = None
    output_format: str | None = None
    template: Path | None = None
    keep_template_figures: bool = False
    confirm_structure: Callable[[AnalysisResult], bool] | None = None
    report_profile: str = "clean"


@dataclass
class StageReport:
    name: str
    status: str
    seconds: float = 0.0
    error: str = ""


class ExcelIntelligenceWorkflow:
    def __init__(
        self,
        settings: Settings,
        *,
        monitor: ResourceMonitor | None = None,
        progress: Callable[[str], None] | None = None,
        abort_event: Any = None,
    ) -> None:
        self.settings = settings
        self.monitor = monitor or (
            ResourceMonitor(settings) if settings.enable_resource_monitoring else None
        )
        self.progress = progress or (lambda message: LOGGER.info(message))
        self.abort_event = abort_event
        self.stages: list[StageReport] = []
        self.timings: dict[str, float] = {}

    # ------------------------------------------------------------------ #
    def run(self, options: WorkflowOptions) -> RunResult:
        source = Path(options.source).expanduser().resolve()
        started_at = datetime.now()
        run_started = time.perf_counter()
        if self.monitor is not None:
            self.monitor.start()
        errors: list[str] = []
        budget = float(self.settings.max_processing_seconds or 0)
        self._deadline = (run_started + budget) if budget > 0 else None
        profile = options.report_profile or self.settings.report_profile or "clean"

        bundle = None
        try:
            with self._stage("خواندن فایل") as stage:
                reader = WorkbookReader(self.settings)
                bundle = reader.load(source)
                stage.detail = f"{bundle.size_bytes} بایت"
            self._check_abort()

            with self._stage("بازرسی ساختار"):
                meta = inspect_workbook(bundle, self.settings)
            self._check_abort()

            with self._stage("تشخیص جدول و استخراج"):
                analysis = self._analyze(
                    bundle,
                    meta,
                    goal=options.goal or options.question,
                    request=options.request or options.goal or options.question,
                )
            self._check_abort()

            if options.confirm_structure is not None and not options.confirm_structure(analysis):
                raise ProcessingAborted("کاربر ادامه پردازش را تأیید نکرد.")

            with self._stage("تحلیل مدل"):
                ai = AIAnalyzer(self.settings).analyze(
                    analysis,
                    model=options.model,
                    task=options.task,
                    question=options.request or options.question,
                    use_ai=options.use_ai,
                )
                analysis.ai = ai.to_dict()
                if ai.error:
                    analysis.warnings.append(f"تحلیل مدل انجام نشد: {ai.error}")
            self._check_abort()

            output_path = options.output_path or self.settings.output_path_for(source)
            output_path = Path(output_path).resolve()
            if output_path == source:
                raise ProcessingAborted("مسیر خروجی با فایل ورودی یکسان است؛ خروجی در مسیر دیگری ذخیره شود.")

            request_text = options.request or options.goal or options.question or ""
            intent = resolve_output_intent(
                explicit_format=options.output_format,
                output_path=options.output_path,
                request_text=request_text,
                default=self._default_output_format(),
            )
            if options.output_path is None:
                output_path = self.settings.output_path_for(
                    source,
                    extension=intent.output_format.extension,
                )
            if options.template:
                template_path = self._resolve_template_path(options.template)
            elif intent.output_format is OutputFormat.DOCX:
                template_path = self._resolve_template_path(None)
            else:
                template_path = None
            specification = ReportSpecification(
                output_format=intent.output_format,
                output_path=output_path,
                profile=profile,
                template_path=template_path,
                preserve_template_figures=options.keep_template_figures,
                request=request_text,
                evidence=intent.evidence,
            )
            specification.validate()
            output_path = specification.output_path.resolve()
            if output_path == source:
                raise ProcessingAborted("مسیر خروجی با فایل ورودی یکسان است؛ خروجی در مسیر دیگری ذخیره شود.")

            finished_at = datetime.now()
            total_seconds = time.perf_counter() - run_started
            performance = None
            if profile == "audit":
                performance = build_performance_report(
                    self.monitor or ResourceMonitor(self.settings),
                    self.settings,
                    started_at=started_at,
                    finished_at=finished_at,
                    total_seconds=total_seconds,
                    timings=self.timings,
                    ai_metadata=self._ai_metadata(ai),
                    input_sha256=bundle.sha256,
                    output_path=str(output_path),
                )

            with self._stage(f"ساخت خروجی {specification.output_format.label_fa}"):
                generation = ReportGenerator(self.settings).generate(
                    specification,
                    analysis=analysis,
                    ai=ai,
                    performance=performance,
                    monitor=self.monitor,
                )
            output_path = Path(generation.output_path)

            # ---------------------------------------------------------------- #
            # The final gate adds output-format correctness and document
            # integrity to the evidence already evaluated before delivery.
            # ---------------------------------------------------------------- #
            quality_gate = QualityGate().evaluate(
                analysis,
                source_path=source,
                specification=specification,
                generation=generation,
                provenance=analysis.provenance,
            )
            analysis.quality_gate = quality_gate
            if quality_gate.blocked:
                blocking = "، ".join(item.stage.label_fa for item in quality_gate.blocking_failures)
                message = (
                    "دروازه کیفیت: گزارش نهایی معتبر نیست؛ "
                    f"مراحل مسدودکننده: {blocking}."
                )
                errors.append(message)
                analysis.warnings.append(message)
            elif quality_gate.status is GateStatus.FAIL:
                analysis.warnings.append(
                    "دروازه کیفیت: تحلیل با «رد» تکمیل شد؛ گزارش کامل و بی‌قید نیست."
                )
            elif quality_gate.status is GateStatus.WARNING:
                analysis.warnings.append(
                    "دروازه کیفیت: تحلیل با هشدار تکمیل شد؛ موفقیت کامل گزارش نمی‌شود."
                )

            input_unchanged = bundle.verify_unchanged()
            if not input_unchanged:
                errors.append("هشدار: فایل ورودی پس از پردازش تغییر کرده است!")
                analysis.warnings.append(errors[-1])
                LOGGER.error("input file changed during processing: %s", source)

            if quality_gate.blocked:
                status = "failed"
            elif quality_gate.status is GateStatus.PASS and not errors and input_unchanged:
                status = "completed"
            else:
                status = "completed_with_warnings"
            self._write_sidecar(analysis, performance, output_path, source)
            self.progress(f"خروجی ساخته شد: {output_path}")
            return RunResult(
                output_path=str(output_path),
                status=status,
                input_sha256=bundle.sha256,
                input_unchanged=input_unchanged,
                analysis=analysis,
                performance=performance.to_dict() if performance is not None else {},
                errors=errors,
                report=generation,
                quality_gate=quality_gate,
            )
        finally:
            if self.monitor is not None:
                elapsed = self.monitor.stop()
                self.timings.setdefault("کل پردازش (پایش)", elapsed)
            if bundle is not None:
                bundle.close()

    # ------------------------------------------------------------------ #
    def analyze_only(
        self,
        source: Path | str,
        *,
        goal: str | None = None,
        request: str | None = None,
    ) -> AnalysisResult:
        """Detect, extract and profile without writing an output workbook."""
        reader = WorkbookReader(self.settings)
        bundle = reader.load(Path(source).expanduser().resolve())
        try:
            meta = inspect_workbook(bundle, self.settings)
            return self._analyze(bundle, meta, goal=goal, request=request or goal)
        finally:
            bundle.close()

    # ------------------------------------------------------------------ #
    def _analyze(
        self,
        bundle,
        meta,
        *,
        goal: str | None = None,
        request: str | None = None,
    ) -> AnalysisResult:
        detector = TableDetector(self.settings)
        extractor = DataExtractor(self.settings)
        calculator = FinancialCalculator(self.settings)

        grids: dict[str, SheetGrid] = {}
        tables: list[TableCandidate] = []
        key_value_tables: list[TableCandidate] = []
        sheet_structures: dict[str, Any] = {}
        detection_warnings: list[str] = []
        for sheet_meta in meta.sheets:
            values_ws = bundle.values[sheet_meta.title]
            formulas_ws = bundle.formulas[sheet_meta.title] if sheet_meta.title in bundle.formulas.sheetnames else None
            grid = SheetGrid(sheet_meta.title, values_ws, formulas_ws, sheet_meta.max_row, sheet_meta.max_col)
            grids[sheet_meta.title] = grid
            structure = detector.detect(grid)
            sheet_structures[sheet_meta.title] = structure
            tables.extend(structure.tables)
            key_value_tables.extend(structure.key_value_tables)
            if structure.tables:
                sheet_meta.table_count = len(structure.tables)
            if sheet_meta.max_row >= self.settings.max_rows_scanned:
                detection_warnings.append(
                    f"برگه «{sheet_meta.title}» به سقف {self.settings.max_rows_scanned} ردیف رسید؛ ادامه برگه خوانده نشد."
                )

        derived = self._derived_sheets(bundle, meta, {table.sheet for table in tables})
        primary = self.select_primary(tables, meta, exclude=derived)
        analysis = AnalysisResult(meta=meta, tables=tables + key_value_tables, primary_table=primary)
        analysis.warnings.extend(bundle.warnings)
        analysis.warnings.extend(detection_warnings)

        # ---------------------------------------------------------------- #
        # Semantic roles and data-quality findings are established before any
        # statistical conclusion.  A numeric identifier, postal code, personnel
        # number or year is a key, never a measure, and a questionable value is
        # recorded with its location instead of being repaired.
        # ---------------------------------------------------------------- #
        analysis.semantics = SemanticClassifier(self.settings).classify(
            source_path=bundle.path,
            sheet_names=[sheet.title for sheet in meta.sheets],
            grids=grids,
            structures=sheet_structures,
        )
        analysis.quality = DataQualityAnalyzer(self.settings).analyze(analysis, grids)

        if primary is None:
            analysis.warnings.append("هیچ جدول داده قابل اتکایی شناسایی نشد؛ فقط ساختار فایل گزارش شد.")
            analysis.structure_confidence = 0.0
        else:
            grid = grids[primary.sheet]
            analysis.column_profiles = list(primary.columns)
            analysis.structure_confidence = primary.confidence
            analysis.normalized_rows = extractor.extract_normalized_rows(grid, primary)
            analysis.generic = StatisticalAnalyzer(self.settings).analyze(grid, primary, goal=goal)

        # ---------------------------------------------------------------- #
        # Every sheet is inspected, not only the primary one.  Statistical
        # profiles are built per sheet with the same type-aware engine.
        # ---------------------------------------------------------------- #
        known_sheets = {sheet.title for sheet in meta.sheets}
        structure_engine = StructureIntelligenceAnalyzer(self.settings)
        sheet_intelligence = []
        for sheet_meta in meta.sheets:
            structure = sheet_structures.get(sheet_meta.title)
            if structure is None:
                continue
            sheet_tables = [table for table in structure.tables if not is_report_sheet(table.sheet)]
            sheet_primary = self.select_best(sheet_tables)
            sheet_generic = None
            if sheet_primary is not None:
                if primary is not None and sheet_primary is primary:
                    sheet_generic = analysis.generic
                else:
                    sheet_generic = StatisticalAnalyzer(self.settings).analyze(
                        grids[sheet_meta.title], sheet_primary, goal=goal
                    )
            all_sheet_tables = structure.tables + structure.key_value_tables
            intel = structure_engine.inspect_sheet(
                grids[sheet_meta.title],
                sheet_meta,
                all_sheet_tables,
                sheet_primary,
                generic=sheet_generic,
                known_sheets=known_sheets,
                source_name=Path(bundle.path).name,
            )
            intel.block_analyses = self._build_block_analyses(
                grids[sheet_meta.title],
                all_sheet_tables,
                sheet_primary=sheet_primary,
                sheet_generic=sheet_generic,
                global_primary=primary,
                global_generic=analysis.generic,
                goal=goal,
            )
            intel.block_relationships = self._block_relationships(all_sheet_tables)
            if len(intel.block_analyses) > 1:
                intel.warnings.append(
                    f"{len(intel.block_analyses)} بلوک مستقل در این برگه شناسایی و جدا تحلیل شد؛ "
                    "هیچ بلوکی فقط به دلیل کوچک‌تر بودن کنار گذاشته نشد."
                )
            sheet_intelligence.append(intel)
        analysis.sheets = sheet_intelligence
        analysis.intelligence = structure_engine.finalize(str(meta.path), sheet_intelligence)
        # Matching needs the finalized sheet-level period evidence (a month
        # title combined with the sheet year), not only the raw block titles.
        analysis.semantic_matching = SemanticMatcher(self.settings).match_analysis(
            analysis, source_name=Path(bundle.path).name
        )
        analysis.extraction = ExtractionModelBuilder(self.settings).build(
            source_path=bundle.path,
            sheet_names=[sheet.title for sheet in meta.sheets],
            grids=grids,
            structures=sheet_structures,
            intelligence=analysis.intelligence,
        )
        # Repeated content, copied summaries, formula-derived sheets and
        # overlapping categories are recorded before any cross-source total.
        analysis.duplicate_control = DuplicateControlAnalyzer(self.settings).analyze(analysis)
        for warning in analysis.intelligence.warnings:
            analysis.warnings.append(warning)

        # ---------------------------------------------------------------- #
        # Merge every periodic sheet into one month x period x metric matrix and
        # answer the Persian request from the numbers Python recomputed.
        # ---------------------------------------------------------------- #
        formulas = {}
        try:
            formulas = {
                name: bundle.formulas[name] for name in bundle.formulas.sheetnames
            }
        except (AttributeError, KeyError):  # pragma: no cover - defensive
            formulas = {}
        integration = PeriodIntegrator(self.settings).build(
            sheet_intelligence, grids, formulas=formulas
        )
        analysis.integration = integration if integration.series else None
        intent = RequestAnalyzer().parse(
            request or "",
            available_metrics=list(
                dict.fromkeys(
                    integration.metric_names
                    + [profile.header for profile in (analysis.column_profiles or [])]
                )
            ),
        )
        analysis.request_intent = intent
        analysis.request_answers = answer_request(intent, analysis, analysis.integration)

        # ---------------------------------------------------------------- #
        # Reconciliation runs after the structure, semantics and period matrix
        # exist.  It validates relationships that are explicitly present in the
        # workbook and never invents a total, a denominator or a period match.
        # ---------------------------------------------------------------- #
        analysis.statistical_reconciliation = StatisticalReconciler(self.settings).analyze(
            analysis, grids
        )

        if primary is not None:

            # Loan reconciliation is an optional domain module, never a prerequisite.
            if primary.kind is TableKind.LOAN:
                analysis.loans = extractor.extract_loans(grid, primary)
                analysis.declared_totals = extractor.find_declared_totals(grid, primary)
                obligations, notes = extractor.collect_obligations(
                    tables + key_value_tables, grids, primary.sheet
                )
                analysis.obligations = obligations
                analysis.warnings.extend(notes)

            calculations = calculator.compute(
                analysis.loans, analysis.obligations, analysis.declared_totals
            )
            if primary.kind is not TableKind.LOAN:
                calculations.extend(calculator.summarize_table(grid, primary))
            analysis.calculations = calculations
            if primary.kind is TableKind.LOAN or analysis.declared_totals:
                analysis.audits = calculator.build_audits(calculations, analysis.declared_totals)
            else:
                analysis.audits = []
            analysis.reconciliation = Reconciler(self.settings).reconcile(
                analysis.loans, analysis.obligations
            )
            analysis.anomalies = AnomalyDetector(self.settings).detect(
                meta,
                analysis.tables,
                primary,
                analysis.loans,
                analysis.obligations,
                analysis.calculations,
                analysis.audits,
                analysis.reconciliation,
                normalized_rows=analysis.normalized_rows,
            )
            intelligence = analysis.intelligence
            if intelligence is not None:
                percent_mismatches = 0
                for audit in intelligence.all_audits:
                    if audit.status is not AuditStatus.MISMATCH:
                        continue
                    if audit.kind not in {"total", "difference"}:
                        percent_mismatches += 1
                        continue
                    analysis.anomalies.append(
                        Anomaly(
                            severity=Severity.WARNING,
                            category="declared-total-mismatch",
                            message=(
                                f"مغایرت جمع/مقایسه در «{audit.sheet}» برای «{audit.metric}»: "
                                f"ثبت‌شده {audit.declared} در برابر محاسبه {audit.computed}"
                            ),
                            sheet=audit.sheet,
                            address=audit.address,
                            detail=audit.cause or "اختلاف بدون اصلاح خودکار ثبت شد.",
                        )
                    )
                if percent_mismatches:
                    analysis.anomalies.append(
                        Anomaly(
                            severity=Severity.INFO,
                            category="placeholder-percent-row",
                            message=(
                                f"{percent_mismatches} سلول در ردیف «درصد» فایل مقدار ثابت دارد و با "
                                "نسبت محاسبه‌شده هم‌خوان نیست"
                            ),
                            detail=(
                                "ردیف درصد نمایشی است؛ مقدار محاسبه‌شده Python مبنا قرار گرفت و "
                                "این ردیف در نتیجه‌گیری استفاده نشد."
                            ),
                        )
                    )
            for warning in analysis.generic.warnings if analysis.generic else []:
                analysis.warnings.append(warning)
            loan_tables = [table for table in tables if table.kind is TableKind.LOAN]
            if len(loan_tables) > 1:
                analysis.warnings.append(
                    f"{len(loan_tables)} جدول وام شناسایی شد؛ جدول «{primary.sheet}!{primary.address}» "
                    "به‌عنوان مرجع انتخاب شد."
                )

        analysis.validation = DataValidator(self.settings).validate(analysis, grids)
        # Provenance is built from the values Python already calculated, then the
        # pre-delivery gate evaluates stages 1-11.  The final gate adds output
        # format and document integrity after the writer has run.
        analysis.provenance = build_provenance_ledger(analysis, source_path=bundle.path)
        analysis.quality_gate = QualityGate().evaluate(
            analysis,
            source_path=bundle.path,
            provenance=analysis.provenance,
        )
        return analysis

    # ------------------------------------------------------------------ #
    @staticmethod
    def _derived_sheets(bundle, meta, sheets_with_tables: set[str]) -> set[str]:
        """Sheets whose data cells are mostly formulas referencing other sheets.

        Such a sheet is an aggregate of the others; using it as the primary
        table would mix two periods into one column set.
        """

        derived: set[str] = set()
        for sheet_meta in meta.sheets:
            if sheet_meta.title not in sheets_with_tables:
                continue
            try:
                worksheet = bundle.formulas[sheet_meta.title]
            except (KeyError, AttributeError, TypeError):
                continue
            formulas = 0
            external = 0
            for row in worksheet.iter_rows(
                min_row=1,
                max_row=min(sheet_meta.max_row, 200),
                max_col=min(sheet_meta.max_col, 60),
            ):
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.startswith("="):
                        formulas += 1
                        if "!" in value:
                            external += 1
            if formulas >= 5 and external / formulas >= 0.5:
                derived.add(sheet_meta.title)
        return derived

    @staticmethod
    def select_primary(
        tables: list[TableCandidate],
        meta,
        exclude: set[str] | None = None,
    ) -> TableCandidate | None:
        visible = {sheet.title for sheet in meta.sheets if sheet.state == "visible"}
        candidates = [table for table in tables if table.sheet in visible] or list(tables)
        if exclude:
            without_derived = [table for table in candidates if table.sheet not in exclude]
            if without_derived:
                candidates = without_derived
        usable = [
            table
            for table in candidates
            if table.row_count >= 2
            and len(table.columns) >= 2
            and not is_report_sheet(table.sheet)
        ]
        if not usable:
            return None
        return ExcelIntelligenceWorkflow.select_best(usable)

    @staticmethod
    def select_best(tables: list[TableCandidate]) -> TableCandidate | None:
        usable = [
            table
            for table in tables
            if table.row_count >= 2
            and len(table.columns) >= 2
            and not is_report_sheet(table.sheet)
        ]
        if not usable:
            return None

        def score(table: TableCandidate) -> float:
            numeric_columns = sum(1 for column in table.columns if column.data_type is ValueKind.NUMBER)
            date_columns = sum(1 for column in table.columns if column.data_type is ValueKind.DATE)
            known_roles = sum(1 for column in table.columns if column.role is not ColumnRole.UNKNOWN)
            value = (
                table.confidence * 2.0
                + min(table.row_count, 40) * 0.06
                + min(numeric_columns, 8) * 0.2
                + min(date_columns, 3) * 0.15
                + min(table.last_col - table.first_col, 20) * 0.02
                + min(known_roles, 6) * 0.15
            )
            if table.kind is TableKind.TIME_SERIES:
                value += 0.2
            if table.kind is TableKind.OBLIGATION:
                value -= 0.3
            if table.kind is TableKind.KPI:
                value -= 0.4
            return value

        return max(usable, key=score)

    # ------------------------------------------------------------------ #
    @staticmethod
    def _analyzable_tables(tables: list[TableCandidate]) -> list[TableCandidate]:
        result: list[TableCandidate] = []
        seen: set[int] = set()
        for table in tables:
            if is_report_sheet(table.sheet):
                continue
            if table.row_count < 1 or len(table.columns) < 1:
                continue
            if id(table) in seen:
                continue
            seen.add(id(table))
            result.append(table)
        return result

    def _build_block_analyses(
        self,
        grid: SheetGrid,
        tables: list[TableCandidate],
        *,
        sheet_primary: TableCandidate | None,
        sheet_generic: Any,
        global_primary: TableCandidate | None,
        global_generic: Any,
        goal: str | None,
    ) -> list[dict[str, Any]]:
        """Analyze every valid block on the sheet, not only the best one."""
        analyses: list[dict[str, Any]] = []
        for index, table in enumerate(self._analyzable_tables(tables), start=1):
            if table is sheet_primary:
                generic = sheet_generic
            elif global_primary is not None and table is global_primary:
                generic = global_generic
            else:
                generic = StatisticalAnalyzer(self.settings).analyze(grid, table, goal=goal)
            headers = list(table.header_rows or [table.header_row])
            roles = classify_block_rows(
                grid, table, headers, block_end_row=table.block_end
            )
            analyses.append(
                {
                    "block_id": f"{table.sheet}#B{index}",
                    "address": table.address,
                    "block_address": table.block_address,
                    "kind": table.kind.value,
                    "title": table.title,
                    "header_row": table.header_row,
                    "first_data_row": table.first_data_row,
                    "last_data_row": table.last_data_row,
                    "block_end_row": table.block_end,
                    "row_count": table.row_count,
                    "column_count": len(table.columns),
                    "confidence": round(table.confidence, 4),
                    "is_primary": table is global_primary,
                    "is_sheet_primary": table is sheet_primary,
                    "measures": [
                        column.header
                        for column in table.columns
                        if column.statistical_role
                        in {DataRole.NUMERIC_MEASURE, DataRole.CALCULATED}
                    ],
                    "row_role_counts": role_counts(roles),
                    "summary_rows": summary_rows_from_roles(roles),
                    "statistics": self._block_statistics(generic),
                    "warnings": list(getattr(generic, "warnings", []) or []),
                }
            )
        return analyses

    @staticmethod
    def _block_statistics(generic: Any) -> list[dict[str, Any]]:
        if generic is None:
            return []
        statistics: list[dict[str, Any]] = []
        for column in getattr(generic, "columns", []) or []:
            if not getattr(column, "is_numeric", False):
                continue
            if getattr(column, "is_identifier", False) or getattr(column, "is_temporal_key", False):
                continue
            statistics.append(
                {
                    "header": column.header,
                    "letter": column.letter,
                    "count": column.count,
                    "missing": column.missing,
                    "distinct": column.distinct,
                    "total": _decimal_text(column.total),
                    "mean": _decimal_text(column.mean),
                    "median": _decimal_text(column.median),
                    "minimum": _decimal_text(column.minimum),
                    "maximum": _decimal_text(column.maximum),
                }
            )
        return statistics

    @staticmethod
    def _block_relationships(tables: list[TableCandidate]) -> list[dict[str, Any]]:
        """Link blocks only when their own headers or periods provide evidence."""
        candidates = ExcelIntelligenceWorkflow._analyzable_tables(tables)
        relationships: list[dict[str, Any]] = []
        for left_index, left in enumerate(candidates):
            for right_index in range(left_index + 1, len(candidates)):
                right = candidates[right_index]
                if left.sheet != right.sheet:
                    continue
                left_signature = tuple(
                    column.parent_header or column.header
                    for column in left.columns
                    if column.role not in {ColumnRole.ID, ColumnRole.DATE}
                )
                right_signature = tuple(
                    column.parent_header or column.header
                    for column in right.columns
                    if column.role not in {ColumnRole.ID, ColumnRole.DATE}
                )
                adjacent = abs(right.header_row - left.block_end) <= 3
                if left_signature and left_signature == right_signature:
                    relationships.append(
                        {
                            "kind": "shared_header_signature",
                            "blocks": [
                                f"{left.sheet}#B{left_index + 1}",
                                f"{right.sheet}#B{right_index + 1}",
                            ],
                            "adjacent": adjacent,
                            "note": (
                                "این دو بلوک سرستون یکسان دارند؛ احتمال ادامه یا تکرار همان جدول است، "
                                "اما داده‌هایشان جدا نگه داشته شد."
                            ),
                        }
                    )
                    continue
                left_periods = {column.period for column in left.columns if column.period}
                right_periods = {column.period for column in right.columns if column.period}
                if left_periods and left_periods == right_periods:
                    relationships.append(
                        {
                            "kind": "shared_period",
                            "blocks": [
                                f"{left.sheet}#B{left_index + 1}",
                                f"{right.sheet}#B{right_index + 1}",
                            ],
                            "adjacent": adjacent,
                            "note": (
                                "این دو بلوک دوره یکسان دارند؛ رابطه فقط ثبت شد و جمع خودکار انجام نشد."
                            ),
                        }
                    )
        return relationships

    # ------------------------------------------------------------------ #
    def _stage(self, name: str):
        return _StageContext(self, name)

    def _check_abort(self) -> None:
        if self.abort_event is not None and self.abort_event.is_set():
            raise ProcessingAborted("پردازش به‌درخواست کاربر متوقف شد.")
        deadline = getattr(self, "_deadline", None)
        if deadline is not None and time.perf_counter() > deadline:
            raise ProcessingAborted(
                "سقف زمان پردازش تمام شد "
                f"({self.settings.max_processing_seconds:.0f} ثانیه)؛ پردازش متوقف شد."
            )

    @staticmethod
    def _ai_metadata(ai: AIOutcome) -> dict[str, Any]:
        metadata = dict(ai.metadata)
        metadata.setdefault("mode", ai.mode)
        metadata.setdefault("model", ai.model or "بدون مدل")
        metadata["error"] = ai.error
        metadata["grounded"] = ai.response.is_grounded
        metadata["unsupported_numbers"] = ai.response.unsupported_numbers
        return metadata

    def _default_output_format(self) -> OutputFormat:
        return OutputFormat.from_token(self.settings.report_format) or OutputFormat.XLSX

    def _resolve_template_path(self, value: Path | str | None) -> Path | None:
        candidate = value or self.settings.docx_template or None
        if candidate in (None, ""):
            return None
        path = Path(candidate).expanduser()
        if not path.is_absolute():
            path = (self.settings.project_root / path).resolve()
        return path

    def _write_sidecar(self, analysis: AnalysisResult, performance, output_path: Path, source: Path) -> None:
        if not self.settings.keep_intermediate_json:
            return
        import json

        self.settings.processed_dir.mkdir(parents=True, exist_ok=True)
        sidecar = self.settings.processed_dir / f"{output_path.stem}.json"
        payload = {
            "source": str(source),
            "source_sha256": analysis.meta.sha256,
            "output": str(output_path),
            "analysis": analysis.to_dict(),
            "raw_cells": analysis.normalized_rows,
            "performance": performance.to_dict() if performance is not None else {},
        }
        try:
            sidecar.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2, default=_json_default),
                encoding="utf-8",
            )
        except OSError as exc:
            LOGGER.warning("could not write sidecar json: %s", exc)


def _json_default(value: Any) -> str:
    """Serialize Decimals and dates that appear inside the preserved raw cells."""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


class _StageContext:
    def __init__(self, workflow: ExcelIntelligenceWorkflow, name: str) -> None:
        self.workflow = workflow
        self.name = name
        self.detail = ""
        self.report = StageReport(name=name, status="running")

    def __enter__(self) -> "_StageContext":
        self.workflow.progress(f"شروع مرحله: {self.name}")
        self.started = time.perf_counter()
        monitor = self.workflow.monitor
        self._scope = monitor.phase_scope(self.name) if monitor is not None else None
        if self._scope is not None:
            self._scope.__enter__()
        self.workflow.stages.append(self.report)
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        if self._scope is not None:
            self._scope.__exit__(exc_type, exc, tb)
        elapsed = time.perf_counter() - self.started
        self.report.seconds = elapsed
        self.workflow.timings[self.name] = elapsed
        if exc_type is None:
            self.report.status = "ok"
            self.workflow.progress(f"پایان مرحله: {self.name} ({elapsed:.2f} ثانیه)")
        else:
            self.report.status = "failed"
            self.report.error = f"{exc_type.__name__}: {exc}"
            self.workflow.progress(f"خطا در مرحله: {self.name} — {exc}")
            LOGGER.exception("stage failed: %s", self.name)
        return False
