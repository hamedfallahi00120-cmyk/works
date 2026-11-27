# Repository guide (Local Office AI)

Python package that reads, profiles and statistically analyzes administrative/organizational
Excel workbooks (Persian/English), computes every figure with `Decimal`, and writes a
formatted Persian report workbook. Loan/obligation reconciliation is an optional domain
module for workbooks that actually contain those columns, not the product's core.
Optional local LLM (Ollama/Qwen) writes narrative only; optional ComfyUI image studio.

## Verified handoff state and mandatory next-stage protocol

The read-only Phase 0 audit is already complete and is recorded in
`PHASE0_BASELINE_AUDIT.md`. **Do not repeat Phase 0, create a replacement audit, or
restart the project from scratch.** The sanitized project handover is
`MASTER_PROJECT_HANDOVER_PUBLIC.txt`; the live evidence ledger is
`docs/quality_evidence.md`; the concise start instructions are
`docs/NEXT_AGENT_START_HERE.md`.

The repository is private. The preparation baseline is on
`fix/multisource-extraction-integrity`. Prompt 01A was merged to `main` as PR #4
(`f572c47`). The dedicated Prompt 01B branch is
`feature/prompt-01b-multiblock-extraction`. Never work directly on `main`.

The preparation baseline is **Phase-1a preparation only**: isolated sandbox evidence,
four focused tests (T2/T3/T4/T9), the quality ledger, the Phase 0 audit, and the
sanitized handover. Prompt 01A is built on top of that baseline and is now merged.

Prompt 01A is implemented and merged:
- T2 (G3/R3) is fixed: totals keywords match on word boundaries, so `کلانتری ۱۲`,
  `تکلیف`, `جمعیت` and `کلاسه پرونده` are ordinary labels while `جمع کل` still totals.
- T3 (G5/R4) is fixed: `سال`, `ماه` and `روز` columns whose content agrees
  (integer years, months 1..12, days 1..31) become `DataRole.TEMPORAL_KEY`, never measures.
- T9 (G11/R8) remains open and stays `xfail(strict=True)`: output suffix is not validated.

T2 and T3 no longer carry xfail markers; they must keep passing normally. Do not
reintroduce substring totals matching or let temporal keys back into numeric measures.

Prompt 01B was merged to `main` as PR #5 (`113cc1f`). Prompt 02A is implemented on
`feature/prompt-02a-semantic-quality`; the branch is local and is not merged or pushed to
`main`. See the section below.

### Prompt 01A extraction model

- `excel/period_detection.py` — period tokens from five evidence sources
  (`column_header`, `sheet_name`, `block_title`, `row_label`, `filename`). Every token
  keeps its evidence text, address, source and confidence; `status` is `detected`,
  `not_detected` or `conflicting`. No period is invented when evidence is insufficient.
  Scalar-source disagreements are reported, never resolved by preferring one source.
  A scalar month (`فروردین` sheet) and a single scalar year (filename) are joined into
  `فروردین 1403`; multiple scalar years stay separate because the pairing would be a guess.
  Row-label evidence requires a numeric value on the same row, so a section title or note
  is never mistaken for a period label. A bare year in a `سال`/`year` temporal-key column
  is row-label evidence too; it stays a `TEMPORAL_KEY`, never a measure.
- `excel/extraction_model.py` — `WorkbookModel → SheetModel → BlockModel → ColumnModel/
  CellRecord/MergedRangeRecord`. It projects the existing detectors (it does not
  re-detect): block boundaries, header rows, parent/sub/period headers, business and
  semantic roles, merged ranges with anchor values, original cell values, formulas,
  cell addresses, period evidence and confidences. Attached as
  `AnalysisResult.extraction` and serialized into the processed JSON.
- `SheetIntelligence` now also carries `filename_years`, `periods`, `period_evidence`,
  `period_sources`, `period_status` and `period_confidence`; filename years take part in
  `metadata_conflicts` (`filename_vs_sheet_name_year`, `filename_vs_title_year`).
- `DataRole.GROUP` separates low-cardinality group columns (`گروه`, `دسته`, `نوع` ...)
  from free text; `DATA_ROLE_FA` maps it to «گروه‌بندی».
- The audit profile writes the new «مدل استخراج» sheet; the clean profile never does.
- `TableDetector` now stops a block at a lone long text row after data begins when the next
  non-empty row is a header (or the region ends), so a second table's title is not swallowed
  into the first block's data range. A sheet with two stacked tables yields two blocks with
  their own boundaries, titles and periods.

Known limits: row-label period evidence is scanned for the first 300 data rows per block
(500 for sheet intelligence), and cells are capped at 4000 per block (`cells_truncated`
flags the cap). Block boundaries, headers and columns stay complete even when cells are
truncated.

### Prompt 01B multi-block extraction and totals

- `excel/row_classification.py` — the shared classifier for `data`, `category`,
  `group_heading`, `subtotal`, `total`, `difference`, `percentage`, `note`, `footer`,
  `title`, `header` and `blank`. It classifies only; it never aggregates.
- `sheet_analyzer.py` — `summary_label_role()` checks subtotal before grand total;
  `SUBTOTAL_KEYWORDS`, `GRAND_TOTAL_KEYWORDS`, `DIFFERENCE_KEYWORDS`, `PERCENT_KEYWORDS`
  and `NOTE_KEYWORDS` are explicit. Totals phrases use word boundaries, so `کلانتری`,
  `تکلیف`, `جمعیت`, `کلاسه پرونده` and `کلاسی` are not totals. The `جمع کل` phrase does
  not match inside `جمع کلاسی`; the standalone `جمع` still marks that label as a generic
  total.
- `table_detector.py` — a totals row with real data below it is a subtotal and stays in
  the table; a terminal totals row is the grand total. `block_end_row` keeps comparison,
  percentage and note rows attached. Detection resumes after the block, so a following
  table is not swallowed.
- `core/models.py` — `TableCandidate.block_end`, `block_end_row`, `subtotal_rows`,
  `grand_total_rows` and `block_address`, all serialized.
- `structure_intelligence.py` / `app/workflow.py` — `block_analyses` and
  `block_relationships`; every analyzable block gets its own statistics, row roles,
  summary rows and boundaries. Blocks are linked only when headers or periods agree, and
  are never summed across automatically.
- `extraction_model.py` — each `BlockModel` carries row roles, summary rows, block end,
  subtotal/grand-total rows and merged ranges through the block end.
- `data_extractor.py` — subtotal rows are excluded from normalized data, declared totals
  and loan/obligation records.
- The audit profile writes «بلوک‌های استخراج»; the clean profile never does.
- `tests/test_multiblock_extraction.py` (16 tests) covers the eight required Prompt 01B
  cases plus multiple subtotals, notes around the grand total, a section title after a
  grand total, row-role distinctions including a category row with a numeric code,
  subtotal exclusion and workflow traceability for every block.
  `tests/test_output_quality.py` guards the audit-only block sheet.

### Prompt 02A semantic classification and data quality

- `excel/semantic_vocabulary.py` — leaf vocabulary and value-pattern evidence for
  identifier, measure, temporal, unit, scale, currency, negative-allowed and
  unique-identifier decisions. A measure word in the header blocks the fixed-width and
  digit-text code rules, so a ten-digit monetary amount stays a measure.
- `excel/semantic_classification.py` — `SemanticRole` (`identifier`, `temporal_key`,
  `measure`, `category`, `group`, `subtotal`, `grand_total`, `descriptive`, plus
  `calculated` and `unknown`), `SemanticField`, `SemanticBlock`, `WorkbookSemantics` and
  `SemanticClassifier`. Every field keeps its role, confidence, business/statistical role,
  unit and evidence. Subtotal and grand-total rows come from `row_classification`.
- `excel/data_quality.py` — `QualityStatus` (`PASS`/`WARNING`/`FAIL`), `QualityFinding`,
  `QualityReport` and `DataQualityAnalyzer`. The eight required checks are
  `missing_values`, `invalid_numeric`, `unexpected_text`, `negative_values`,
  `duplicate_identifiers`, `inconsistent_units`, `conflicting_periods` and
  `suspicious_values`. Quality runs on every analyzable block; findings keep the sheet,
  block address, cell, column, count and samples, and values are never repaired.
  Severity is requirement-based: complete measure `PASS`, some blanks `WARNING`,
  conflicting period or broken unique identifier `FAIL`. `conflicting_periods` compares
  scalar sources with each other and with column-header/row-label periods. A measure-word
  header with all-text values is `unexpected_text`; `DESCRIPTIVE_HEADER_WORDS` keeps
  «توضیحات مبلغ» descriptive.
- `sheet_analyzer.py` — `assign_statistical_roles` uses the shared identifier evidence and
  checks temporal keys before identifiers; formula ratios never override a key role.
- `core/models.py` — `AnalysisResult.semantics` and `AnalysisResult.quality`, serialized
  into the processed JSON.
- `app/workflow.py` — semantics and quality are built after table detection and before
  statistics.
- The audit profile writes «نقش‌های معنایی» and «کیفیت داده (تفصیلی)» (with a block
  column); the clean profile never does. The clean «اعتبارسنجی داده» sheet shows a compact
  quality summary and only non-`PASS` findings.
- `tests/test_semantic_quality.py` (23 tests) covers numeric identifiers versus measures,
  date/year/month keys versus measures, missing values, invalid numeric content,
  unexpected text, negative values, unit inconsistencies, conflicting periods,
  duplicate identifiers and record numbers, administrative/vehicle/region and plural code
  headers, English and Arabic-Indic digits, all-text measure columns, category headers
  that mention a measure word, metadata evidence and per-block quality with real openpyxl
  workbooks. `tests/test_output_quality.py` guards the two audit-only sheets.

### Instructions for the next AI agent

1. Fetch the repository and start from the branch that contains Prompt 02A
   (see `docs/quality_evidence.md` for the verified state).
2. Verify the checked-out branch, HEAD SHA and clean working tree. Do not switch to
   `main` for implementation.
3. Read `docs/NEXT_AGENT_START_HERE.md`, `MASTER_PROJECT_HANDOVER_PUBLIC.txt`,
   `PHASE0_BASELINE_AUDIT.md`, this guide and `docs/quality_evidence.md`.
4. Treat the Phase 0 audit and its recorded evidence as the accepted project baseline.
   Do not repeat the full audit or ask the user to resend these repository documents.
5. Prompt 01A (**Data Model, Workbook Structure, and Metadata Extraction**), Prompt 01B
   (**Multi-Block Extraction, Table Boundaries, and Total Detection**) and Prompt 02A
   (**Semantic Classification and Data Quality Validation**) are implemented. Do not
   repeat them or substitute audit Phase 1a for them. The next engineering action starts
   only after the owner accepts Prompt 02A.
6. Follow the supplied shared execution contract: one dedicated branch per stage,
   focused tests, actual evidence, no merge/push to `main`, no rewrite of project
   architecture, Python as the numeric source of truth, and Qwen for narrative only.
7. The three original administrative files and the private reference workbook are not
   in GitHub. Do not fabricate them, commit them, or claim real-file acceptance. The
   owner performs those tests locally on the Windows target.
8. Keep the Windows target `F:\Local-Office-AI\venv`, system Python, `F:\work`,
   `legacy/` and all original input documents untouched. Never recreate or upgrade
   the owner's Windows environment.

Any reported test count must come from the actual command run in the current environment.
The 36 tests marked `requires_sample` depend on a non-public workbook and are excluded
from the CI gate with `-m "not requires_sample"`; they are not counted as successful.

## Layout
- `src/local_office_ai/core/` — settings (JSON override), logging, exceptions, dataclasses.
- `src/local_office_ai/excel/` — read → inspect → detect tables → normalize → profile →
  descriptive stats → frequencies → groups/crosstabs → trend/correlation → outliers →
  quality → optional loan calculations/reconciliation → report writer.
  `semantic_vocabulary.py`, `semantic_classification.py` and `data_quality.py` classify
  field roles and record quality findings before statistics.
  `statistical_analysis.py` is the general engine; `generic_report.py` holds the generic
  report sheets as a mixin. `validation.py` holds the 9-check quality checklist.
- `src/local_office_ai/ai/` — Ollama client, context builder, prompts, response validator,
  analyzer (model + deterministic fallback), ComfyUI image studio.
- `src/local_office_ai/monitoring/` — psutil sampler + Windows GPU counters (`typeperf`).
- `src/local_office_ai/app/` — workflow stages and coordinator.
- `src/local_office_ai/ui/` — argparse CLI, optional Tkinter wizard, and `event_bridge`
  (workers post `UiEvent`s; only `root.after` touches widgets).
- `src/local_office_ai/ui/desktop/` — primary modern PySide6/RTL interface: pages, RTL
  theme, Persian font discovery, QThread workers, results/preview, an optional
  hardware panel and an optional desktop `floating_monitor.py` widget (frameless,
  always-on-top, tray, persisted geometry, optional Windows autostart). Entry point
  is `run_ui.py` / `Run_Local_Office_AI_Modern.bat`; `--floating-only` starts just
  the widget.
- `tests/` — pytest suite with real in-process HTTP servers for Ollama/ComfyUI.
- `legacy/` — untouched V4.2 script and its global-`py` launcher. Never make these primary.
- `data/input/` — sample workbook (NOT present in this repository; tests needing it fail).
  `reference/v5_sample_report/` — design reference only (also not present).
- `Deploy_Update.bat` / `DEPLOY.md` — code-only deployment manifest. It must copy
  `run_ui.py` and `Run_Local_Office_AI_Modern.bat` and must never copy `venv`,
  `data\input`, `data\output`, `data\processed`, `logs`, `reports` or `reference`.
  `tests/test_deploy_manifest.py` parses both lists.

## Commands
```bash
python -m pip install -r requirements.txt -r requirements-documents.txt -r requirements-dev.txt
python -m pytest tests -o addopts="" -q      # collection varies by optional dependencies; record actual output
python -m pytest tests -o addopts="" -q -m "not requires_sample"   # CI gate: reference workbook is not shipped
python run_engine.py --file <xlsx> --no-ai --yes
python run_engine.py --file <xlsx> --request "مقایسه 1404 با 1403 و بیشترین ماه" --no-ai
python run_engine.py --folder data/input --request "جمع شاخص‌ها را مقایسه کن" --no-ai
python run_engine.py --file <xlsx> --audit --no-ai     # full traceability output
python run_engine.py --check-structure --file <xlsx>
```

## Report profiles
- Default is `report_profile="clean"`: only sheets the request needs (`پاسخ درخواست`,
  `داشبورد`, `خلاصه اجرایی`, monthly/comparison/statistics sheets when the data supports
  them, `شناسنامه داده`, `اعتبارسنجی داده`, `موارد نیازمند بررسی`). No monitoring, raw-data,
  provenance or AI-execution sheets ever reach the delivered workbook.
- `--audit` (or `report_profile="audit"`) adds the technical/traceability sheets for
  developer debugging; `tests/conftest.py` opts into it so regression tests keep coverage.
- `excel/report_writer.py::_plan` is the single source of truth for which sheet is written.
- Resource monitoring is optional: `WorkflowOptions.monitor` stays `None` in normal runs.
- The AI/provenance sheet is audit-only. With a live model the narrative is rendered inside
  `خلاصه اجرایی` (and `پاسخ درخواست`), so a clean workbook still has exactly the requested
  sheets. `test_clean_profile_keeps_model_narrative_out_of_a_technical_sheet` guards it.
- The modern PySide6 UI mirrors the CLI: a `درخواست (فارسی)` field feeds `request=`,
  `AnalysisWorker` runs `Coordinator`, and the `تحلیل پوشه` button runs selected files
  through `FolderWorkflow` in a QThread. The legacy Tk bridge remains as a rollback path.
- Folder reports carry `درخواست کاربر`, `تفسیر درخواست` and a Decimal-built `پاسخ درخواست`
  row in `خلاصه پوشه`; `FolderResult.request` preserves the raw text.

## Request-aware analysis
- `ai/request_analyzer.py` parses Persian requests into scope, target metrics, months, years.
- `excel/data_qa.py::answer_request` produces grounded answers: scope, metadata/label
  conflicts, ranking per metric, period comparison, quality, incomplete periods. Every
  number comes from `PeriodIntegration`; the model only phrases it.
- `excel/period_integration.py` builds month × period matrix from the year band of each
  sheet, excludes derived sheets, and marks incomplete months (day coverage below the
  calendar length) so totals/ranking never mix partial months. A month that is incomplete
  but equally covered in both compared periods is included and noted.
- Context building keeps `period_analysis` and `request_answers` through every trimming
  stage (they are normally what the request needs) and reports `main_period` so the model
  ranks the right year.
- `app/folder_workflow.py` analyzes a folder, validates each Excel package, optionally
  filters to an explicit selected subset, dedupes by content hash and writes a cross-file
  comparison; source files are never modified.

## Year-label reality check
Monthly workbooks often disagree: sheet name says 1404, sheet title says 1405, the year
band says 1404/1405, and a `هفت ماهه` sheet sums the first column of every sheet while
labelling it 1403. The engine reports the conflict (`metadata_conflicts`) and never picks
a winner; comparison uses only sheets that carry both requested years. `آمار روزانه 1404.xlsx`
is such a file: 4 sheets labelled 1403/1404, 4 labelled 1404/1405, heaviest overlap is
مرداد..آبان.
```

## Rules
- Python is the only authority for arithmetic. Never let model output alter figures.
- A numeric identifier (postal code, personnel number, vehicle/region code, record number,
  year/month key) is never a measure. Use `semantic_vocabulary.identifier_evidence`; do not
  add a header-only ID check that ignores value patterns.
- Data-quality findings carry `PASS`/`WARNING`/`FAIL` and a sheet/cell location. Never
  delete, repair or coerce a questionable value. The clean profile lists non-`PASS`
  findings only; the audit profile keeps the full «کیفیت داده (تفصیلی)» sheet.
- General statistics must stay type-aware and non-destructive: missing values and IQR
  outliers are counted/labelled, never silently deleted.
- Do not make loan/financial columns a prerequisite for selecting the primary table or
  building the report. Skip loan sheets when no loan records exist.
- Unmeasurable metrics are `None`/"در دسترس نبود", never `0`.
- Never write to the input workbook; reports go to `data/output/` via atomic temp file,
  the machine-readable JSON (with raw cells) goes to `data/processed/`.
- No sheet names, column letters, or sample figures hard-coded in engine logic.
- Table detection must stay layout-agnostic; new heuristics need a test in
  `tests/test_sheet_structure.py`, `tests/test_workflow_generic.py` or
  `tests/test_structure_intelligence.py`.
- Multi-row headers keep parent and child identity (`parent_header`, `sub_header`, `period`).
  A merged parent must extend the detected span to every child column, including the final pair.
- Row roles (total/difference/percentage/note/footer) stay outside descriptive observations.
  Cross-sheet totals exclude incomplete periods and list them in `excluded_incomplete`.
- Empty cells, real zeros, invalid values and formula errors are counted separately. Never
  coerce a blank to zero. Identifier/temporal keys never enter means, variance or correlation.
- AI context is built by Python in chunks under `max_ai_context_chars`/`max_ai_prompt_chars`.
  Before every model call the prompt is also checked against `ollama_num_ctx` with a reserved
  output budget; oversized prompts are trimmed instead of relying on an HTTP 400 retry.
  Model failures must degrade to the rule-based summary without crashing; report provenance
  separately for Python, real model, rule-based fallback and failed chunks.
- Persian text in user-facing strings; code/comments/docstrings in English.
- Tkinter is imported lazily only (headless CI has no `libtk8.6.so`). Worker threads must
  never call Tk; use `UiBridge` (`Queue` + confirmation `Event`) and `root.after`.
- In the PySide6 UI, workers never touch widgets: only signals cross to the main thread.
  Keep the hardware panel collapsed by default and isolated from the report writer.
- The floating widget must reuse `HardwareMonitor` through `HardwareWorker`; never add a
  second sensor implementation. It samples only while visible, persists geometry/opacity
  in `ui_settings.json`, and keeps `None` as «در دسترس نیست». Windows autostart is
  opt-in, writes only `HKCU\...\Run`, and is reported as unsupported elsewhere.
- `Deploy_Update.bat` and `DEPLOY.md` must keep the same code-only manifest: entry points
  in, user data (`venv`, `data\input`, `data\output`, `data\processed`, `logs`, `reports`,
  `reference`) out. `tests/test_deploy_manifest.py` enforces it.
- A missing monthly grid is reported as «دوره ماهانه شناسایی نشد», never as «0 ماه».
- Scale words (هزار/میلیون/میلیارد/همت) multiply the parsed `Decimal`; never drop them.
- `config/synonyms.json` is a matching aid only: never rewrite `institution_raw` or cell text.
- AI number comparison uses `NumberComparisonPolicy` with zero tolerance. Do not reintroduce
  a relative tolerance for financial figures.
- Semantic column roles must consider header and real content; conflicting headers produce
  a reason string, not a silent role.
- Trend/correlation language must stay descriptive: correlation is not causation.
- Generic report sheets are conditional; never emit empty loan audit/reconciliation sheets.
- Delivery criteria are output quality and correctness, not internal telemetry: no test,
  performance, GPU, RAM or "AI execution" sheet in the delivered workbook.
- Only day-number date columns may drive calendar completeness; row numbers must never be
  read as days (`labels_are_days` in `structure_intelligence.py`).
- A few mismatching declared totals (under 10% of audited totals) stay a warning, not a
  failed checklist, because Python figures remain authoritative.

## Desktop UI structure (RTL refinement pass)
- `ui/desktop/theme.py` — `Palette` + `palette_for(theme)`, `build_stylesheet(...)`, `apply_palette`.
  Every colour, radius and control height for both themes lives here; pages never hard-code colours.
- `ui/desktop/components.py` — shared primitives: `Card`, `SectionHeader`, `Badge`, `MetricTile`,
  `CallbackButton`, `IconButton`, `NavigationButton`, `EmptyState`, `BarChart`, plus the layout
  helpers `form_row()` and `horizontal_line()`.
- `ui/desktop/chrome.py` — window chrome: `Sidebar` (sectioned nav from `NAV_SECTIONS`, collapsible
  to a 68 px icon rail), `TopBar`, `HardwarePanel`, `StatusBar`.
- `ui/desktop/source.py` — `SourceListCard` (multi-file/folder picker + file table) and
  `RequestComposer` (Persian request box, suggestion chips, primary action). Home, file and
  folder pages reuse these two instead of re-implementing source selection.
- `ui/desktop/results.py` — results card with tabs (خلاصه/یافته‌ها/جداول/نمودار/هشدارها).
- `ui/desktop/formatting.py` — `ltr()`, `rtl()`, number/percent helpers, `align_table_header()`.
- Rules: never restore `QGraphicsDropShadowEffect`/`QGraphicsOpacityEffect` on page or card
  widgets — on the Windows target they painted solid black rectangles over the content, and page
  cross-fades were removed for the same reason (the symptom was the black, seemingly empty boxes
  in the 1366×768 screenshots). Prefer real layouts over fixed widths/heights so 1366×768…1920×1080
  all stay tidy, and keep sidebar icon+label rows built from a layout, not a rich-text label.
- Qt's own Persian translation (`qtbase_fa`) is installed in `desktop/app.py::_install_translations`;
  without it every `QFileDialog`/`QMessageBox` label is English ("New folder", "OK", "Cancel").
  Keep the translator referenced, otherwise Qt garbage-collects it and the labels revert.
- `RequestComposer.set_output_note()` states where the report is written (file/folder pages), and
  `MainWindow._sync_compact()` hides the home hero subtitle below 820 px so the primary action
  stays reachable at 1366×768.
- Typography defaults live in `ui_settings.py` / `config/ui_settings.json`: body 13 pt, headings
  19 pt, family auto-detected by `fonts.py` (bundled font folders first, then installed Persian
  system fonts; Vazirmatn was detected on the reference run). The chosen family is applied to
  the QApplication font and re-applied to QSS on every theme change.
- Offscreen preview for UI work: `QT_QPA_PLATFORM=offscreen python run_ui.py` is not enough for
  screenshots; grab widgets with `window.grab().save(...)` from a small harness script.

## Environment notes
- Target host is `F:\Local-Office-AI` with `venv\Scripts\python.exe` (Python 3.14).
  Never use global `py`, create a new venv, reinstall Python, or download Ollama models.
- Original V4.2 files live in `legacy/` and must not be modified.
- Windows GPU paths and the Tk window cannot be verified on Linux; keep both "unavailable"
  branches honest and state them in the delivery report.
