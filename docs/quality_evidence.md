# Quality evidence ledger — Local Office AI

This ledger records the verified handoff state for the next engineering stage.
The preparation baseline was Phase-1a preparation only. **Prompt 01A (Data Model, Workbook
Structure, and Metadata Extraction) is implemented and merged to `main` as PR #4.**
Prompt 01B (Multi-Block Extraction, Table Boundaries, and Total Detection) is merged as
PR #5, Prompt 02A (Semantic Classification and Data Quality Validation) as PR #6 and
Prompt 02B (Statistical Reconciliation and Calculation Integrity) as PR #7. **Prompt 03A
(Semantic Source Matching and Period Alignment) is implemented on
`feature/prompt-03a-semantic-source-matching`.** Evidence for each stage is recorded at
the end of this file.

## Verified repository state

| Item | Value |
| :--- | :--- |
| Repository | `hamedfallahi00120-cmyk/works` (private) |
| Phase-0 base | `main` at `6df625659c97c0b7da1f78f33fa5fc37864a7580` |
| Prompt 01A merge | PR #4; `main` at `f572c47` |
| Prompt 01B merge | PR #5; `main` at `113cc1f` |
| Prompt 02A merge | PR #6; `main` at `d173465` |
| Prompt 02B merge | PR #7; `main` at `891b662` |
| Preparation branch | `fix/multisource-extraction-integrity` |
| Preparation commits | `5b20cce` (tests + ledger), `06dae62` (sanitized handover + audit), followed by CI/documentation-preflight commits |
| Phase 0 | Completed; see `PHASE0_BASELINE_AUDIT.md` |
| Prompt 01A | Implemented and merged as PR #4 (evidence at end) |
| Prompt 01B | Implemented and merged as PR #5 (evidence at end) |
| Prompt 02A | Implemented and merged as PR #6 (evidence at end) |
| Prompt 02B | Implemented and merged as PR #7 (evidence at end) |
| Prompt 03A | Implemented on `feature/prompt-03a-semantic-source-matching`; PR #8 open (evidence at end) |
| Production source fixes in the preparation baseline | None; Prompts 01A, 01B, 02A, 02B and 03A add production source changes (below) |

## Evidence labels

- `[SBX]`: isolated Linux sandbox evidence, Python 3.13.15. Not Windows target validation.
- `[GHA-WIN]`: GitHub-hosted Windows Actions runner. This is CI evidence, not the user's local installation.
- `[WIN-TARGET]`: the actual user environment at `F:\Local-Office-AI`, Python 3.14.7. Not accessed or validated from this session.

## Phase-1a preparation evidence (before original Prompt 01A)

### Isolated sandbox baseline [SBX]

The sandbox was created at `/tmp/loa-sandbox-venv` using Python 3.13.15. Dependencies were installed there only. The Windows virtual environment, system Python, and user project dependencies were not touched.

Recorded before the four targeted tests were added:

```text
python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 166 passed, 2 skipped, 36 deselected (exit 0)
```

The original report also recorded `--collect-only -> 202`. These are historical sandbox figures; do not use them as a current cross-platform test-count promise. The full collection differs with optional PySide6 availability and the four later test additions.

### Targeted tests

| ID | Test | Guard | Defect state before implementation | Current CI policy |
| :--- | :--- | :--- | :--- | :--- |
| T2 | `test_total_label_needs_a_word_boundary` | G3 / R3 | **Fixed in Prompt 01A:** totals keywords match on word boundaries | Normal passing test |
| T3 | `test_year_month_day_columns_are_not_measures` | G5 / R4 | **Fixed in Prompt 01A:** year/month/day keys are temporal, not measures | Normal passing test |
| T4 | `test_code_columns_stay_identifiers` | G5 / R5 | Passes; R5 was not reproducible and must not be treated as a confirmed defect | Normal passing test |
| T9 | `test_output_suffix_is_validated` | G11 / R8 | Known defect: XLSX content may be written under a `.docx` suffix | `xfail(strict=True)` |

The strict expected-failure markers are test-baseline bookkeeping, not defect fixes. They keep the baseline CI green while preserving executable regression tests. When a later prompt fixes one of these defects, that test must pass normally and its xfail marker must be removed; strict xfail intentionally makes an unexpected pass fail CI.

### First GitHub Actions run [GHA-WIN] — historical failure

Run `45984936491`, commit `06dae62`, Windows Server 2025, Python `3.14.10`:

```text
3 failed, 178 passed, 1 skipped, 38 deselected
```

The three failures were exactly T2, T3, and T9. Dependency installation and openpyxl import succeeded. Artifact upload also failed because the repository's GitHub Actions artifact storage quota was exhausted. This prompted two preparation-only CI changes:
- pin CI to the target Python `3.14.7`;
- make test-result artifact upload best-effort (`continue-on-error: true`), because artifact storage is diagnostic and must not decide test correctness.

**A post-change CI run is required before calling the preparation branch green.** Its result is recorded in the update section below.

## Phase 0 audit corrections

1. R5 was not reproducible: the duplicate-role demotion path is unreachable for the profiled named columns because the parent header is already populated. T4 remains a passing guard, not a fabricated red test.
2. ZIP drift is **7 files**, not 6; `src/` is byte-identical. The differing files include `.gitignore`.
3. Test counts are environment- and revision-dependent. The audit's 209 is a count of test functions, not the same thing as collected pytest items. The initial audit did not have pytest/PySide6 dependencies installed, so the next agent must use actual current pytest output rather than assume a static count.

## Explicitly not validated

- The user's actual Windows environment `F:\Local-Office-AI` and its existing `venv`.
- Ollama/Qwen, LibreHardwareMonitor, PySide6 desktop runtime and Tesseract in the user's installation.
- The three original administrative files and real-file acceptance. They are intentionally absent from GitHub and must remain local/read-only.
- The private source workbook used by `requires_sample` tests. The 36 marked tests are excluded from the CI gate, not silently treated as passed.

## Preparation CI update

- CI configuration and strict-xfail bookkeeping have been committed to the preparation branch.
- **Run 11 passed:** [GitHub Actions run 45987467733](https://github.com/hamedfallahi00120-cmyk/works/actions/runs/45987467733), validated commit `7721c7232c5b1016400e1676f4b6c82045a93160`.
- Runner: Windows Server 2025, CPython **3.14.7**, pytest 9.2.0.
- Targeted gate result: **178 passed, 1 skipped, 38 deselected, 3 xfailed** in 68.34s; exit code 0.
- Sample-dependent collection: **38/220** tests carry `requires_sample`; they were collected but intentionally not executed because their private reference workbook is absent.
- Artifact upload was attempted but GitHub reported that the Actions artifact storage quota is full. Because this diagnostic upload is configured as best-effort, the job and workflow completed successfully. The test result above is from the runner log, not the artifact.
- The three `xfail(strict=True)` outcomes are the known defects listed above, not fixes. The CI gate is green while accurately retaining those unresolved issues.
- The next-stage branch `feature/prompt-01a-data-model-metadata` was created and verified at the same prepared baseline.
- **Prompt 01A readiness: repository and handoff prepared.** The next AI may start the exact original Prompt 01A on the dedicated branch without repeating Phase 0. This CI run validates the code/tests/configuration at the stated commit; it does not validate the owner's local Windows installation, Ollama/Qwen, LHM, or the real administrative files.


## Prompt 01A implementation evidence

Branch: `feature/prompt-01a-data-model-metadata`. Base: the verified preparation baseline;
`e7b491a200fee77c3bb394341c15d4efbf8ac408` remains an ancestor. `main` was not modified.

### What changed

- `src/local_office_ai/excel/period_detection.py` (new): period tokens from column
  headers, worksheet names, block titles, row labels and filenames. Each token records
  its evidence text, cell address (when one exists), source kind and confidence.
  `status` is `detected`, `not_detected` or `conflicting`; scalar-source disagreements
  are reported, not resolved. Month names match as words (`دی` no longer matches `ردیف`).
  A scalar month and a single scalar year combine into one `month_year` token; two scalar
  years stay separate. Row-label evidence requires a numeric value on the same row, and a
  bare year in a `سال`/`year` temporal-key column is row-label evidence too.
- `src/local_office_ai/excel/extraction_model.py` (new): `WorkbookModel → SheetModel →
  BlockModel → ColumnModel / CellRecord / MergedRangeRecord`. Projects the existing
  detectors into block boundaries, header rows, parent/sub/period headers, business and
  semantic roles, merged ranges with anchor values, original cell values, formulas, cell
  addresses, periods and confidence. Attached as `AnalysisResult.extraction`; serialized
  by `AnalysisResult.to_dict()`.
- `structure_intelligence.py`: `SheetIntelligence` gains `filename_years`, `periods`,
  `period_evidence`, `period_sources`, `period_status`, `period_confidence`; filename years
  participate in `metadata_conflicts` (`filename_vs_sheet_name_year`,
  `filename_vs_title_year`); month may fall back to the filename when sheet name and title
  have none, with a warning.
- `sheet_analyzer.py`: totals keywords match on word boundaries (T2 fix); year/month/day
  headers whose content agrees become `DataRole.TEMPORAL_KEY` (T3 fix); low-cardinality
  group headers become `DataRole.GROUP`.
- `table_detector.py`: a lone long text row after data starts ends the block, so a second
  table's title is not absorbed into the first block; merged covered cells are now recorded
  as `CellRecord` entries with `merged_anchor`/`merged_range`.
- `core/models.py`: `DataRole.GROUP`; `AnalysisResult.extraction`.
- `generic_report.py` / `report_writer.py`: audit-only «مدل استخراج» sheet; clean profile
  unchanged.
- Tests: `tests/test_extraction_model.py` (15 tests) covers the seven required cases plus
  text and numeric-year row labels, cross-source month/year, role families, two stacked
  blocks, no-period-without-evidence, filename-vs-sheet-name conflict and scalar-year
  merging.
  T2 and T3 xfail markers removed; T9 remains `xfail(strict=True)`.

### Test evidence [SBX]

Sandbox: Linux, CPython 3.13.15. Not Windows-target validation.

```text
python -m pytest tests/test_extraction_model.py -o addopts="" -q
-> 15 passed in 0.38s (exit 0)

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 184 passed, 2 skipped, 36 deselected, 1 xfailed in 19.20s (exit 0)
```

The single `xfailed` outcome is T9, the still-open output-suffix defect. The 36 deselected
tests are the `requires_sample` set; their private reference workbook is absent and they are
not counted as passing.

### End-to-end report evidence [SBX]

A synthetic two-level-header workbook (`merged.xlsx`, block `A3:D7`) was analyzed with
`run_engine.py --no-ai --yes`:

```text
clean profile: داشبورد، خلاصه اجرایی، شناسنامه داده، آمار توصیفی، روند و همبستگی،
               اعتبارسنجی داده، موارد نیازمند بررسی
audit profile: ... + مدل استخراج (block داده#B1, A3:D7, header rows 3/4,
               periods 1403/1404 from column_header, merged C3:D3)
```

A second synthetic workbook with two stacked tables (`two_blocks.xlsx`) was run through
`Coordinator.run(..., use_ai=False)`:

```text
clean profile: 8 delivery sheets, no مدل استخراج
audit profile: 23 sheets incl. مدل استخراج
               داده#B1 A2:C5 1403 block_title (A1)
               داده#B2 A8:C11 1404 block_title (A7)
```

### Not validated

- The owner's Windows environment `F:\Local-Office-AI` (Python 3.14) and its `venv`.
- The three original administrative files and the private reference workbook. No
  real-file acceptance is claimed; the owner runs that locally.
- Ollama/Qwen narrative, ComfyUI, LibreHardwareMonitor, Tkinter and PySide6 runtime.

## Prompt 01B evidence — multi-block extraction and total classification

Full write-up: `docs/prompt_01b_technical_report.md`.

### What changed

- `src/local_office_ai/excel/row_classification.py` (new): one shared, context-aware row
  classifier. It distinguishes `data`, `category`, `group_heading`, `subtotal`, `total`,
  `difference`, `percentage`, `note`, `footer`, `title`, `header` and `blank`. A totals
  label is matched with word boundaries; position inside the block decides subtotal vs
  grand total. It only classifies; it never aggregates.
- `sheet_analyzer.py`: explicit `SUBTOTAL_KEYWORDS`, `GRAND_TOTAL_KEYWORDS`,
  `DIFFERENCE_KEYWORDS`, `PERCENT_KEYWORDS` and `NOTE_KEYWORDS`;
  `summary_label_role()` returns `subtotal` before `total`. Multi-word totals phrases use
  word boundaries on both ends, so the `جمع کل` phrase does not match inside `جمع کلاسی`.
  The standalone `جمع` still marks that label as a generic total; `کلاسی` alone is not a
  total.
- `table_detector.py`: a totals row with real data below it is a subtotal and stays inside
  the table; a terminal totals row is the grand total. `block_end_row` keeps comparison,
  percentage and note rows attached to the block. Table detection resumes after the block,
  so a following table is not swallowed. Formula-only cells count as numeric evidence.
- `core/models.py`: `TableCandidate.block_end_row`, `subtotal_rows`, `grand_total_rows`,
  `block_end` and `block_address`; all are serialized.
- `structure_intelligence.py`: `SheetIntelligence.block_analyses` and
  `block_relationships`; row roles now come from the block that owns each row, so a
  secondary table's headers and subtotals are not mistaken for notes. `_column_totals`,
  `_coverage`, `_audits` and period comparison skip structural rows.
- `app/workflow.py`: every analyzable block on a sheet gets its own statistics, row roles,
  summary rows and boundaries. Blocks are linked only when their headers or periods agree;
  no automatic cross-block sum is performed.
- `extraction_model.py`: each `BlockModel` now carries row roles, summary rows, block end,
  subtotal rows and grand-total rows, plus merged ranges through the block end.
- `data_extractor.py`: subtotal rows are excluded from normalized data, declared totals and
  loan/obligation records.
- `generic_report.py` / `report_writer.py`: audit-only «بلوک‌های استخراج» sheet with every
  block's address, boundaries, measures, summary rows and block relationships. The clean
  profile never writes it.
- Tests: `tests/test_multiblock_extraction.py` (16 tests) covers two independent tables,
  different header structures, a mid-table subtotal, multiple subtotals with data after
  each, a terminal grand total, notes before and after a grand total, a section title
  after a grand total, `کل` inside ordinary words, blank rows between blocks, repeated
  headers, different periods, row-role distinctions including a category row with a
  numeric code, subtotal exclusion from declared totals and from statistics, and workflow
  traceability for every block.
  `tests/test_output_quality.py` now asserts the audit-only block sheet is present in
  audit output and absent from clean output.

### Test evidence [SBX]

Sandbox: Linux, CPython 3.13.15. Not Windows-target validation.

```text
python -m pytest tests/test_multiblock_extraction.py -o addopts="" -q
-> 16 passed in 0.46s (exit 0)

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 196 passed, 6 skipped, 36 deselected, 1 xfailed in 20.75s (exit 0)
```

The single `xfailed` outcome is still T9, the open output-suffix defect. The 36 deselected
tests are the `requires_sample` set; their private reference workbook is absent and they are
not counted as passing.

### Not validated

- The owner's Windows environment and the three original administrative files.
- Real-file acceptance for multi-block workbooks; the evidence uses synthetic `.xlsx`
  files created by the test suite.

## Prompt 02A — Semantic classification and data quality (branch evidence)

Branch `feature/prompt-02a-semantic-quality`, based on `main` at `113cc1f` (the
merge of Prompt 01B). No cross-source aggregation was added.

### Implementation

- `excel/semantic_vocabulary.py` (new): shared identifier, measure, temporal,
  unit, scale, currency, negative-allowed and unique-identifier vocabularies plus
  value-pattern evidence. A measure word in the header blocks the fixed-width and
  digit-text code rules, so wide monetary amounts stay measures.
- `excel/semantic_classification.py` (new): `SemanticRole` for `identifier`,
  `temporal_key`, `measure`, `category`, `group`, `subtotal`, `grand_total` and
  `descriptive` (plus `calculated` and `unknown`), `SemanticField`,
  `SemanticBlock`, `WorkbookSemantics` and `SemanticClassifier`. Each field keeps
  its role, confidence, business/statistical role, unit and evidence. Subtotal
  and grand-total rows come from the shared `row_classification` module.
- `excel/data_quality.py` (new): `QualityStatus` (`PASS`/`WARNING`/`FAIL`),
  `QualityFinding`, `QualityReport` and `DataQualityAnalyzer`. The eight required
  checks are `missing_values`, `invalid_numeric`, `unexpected_text`,
  `negative_values`, `duplicate_identifiers`, `inconsistent_units`,
  `conflicting_periods` and `suspicious_values`. Quality runs on every analyzable
  block, and each finding keeps its sheet, block address, cell, column, count and
  samples; values are never repaired. `conflicting_periods` compares scalar
  sources with each other and with column-header/row-label periods.
  `duplicate_identifiers` treats record/case numbers as strict. A measure-word
  header with all-text values is reported as `unexpected_text`, while
  `DESCRIPTIVE_HEADER_WORDS` keeps a header such as «توضیحات مبلغ» descriptive.
- `sheet_analyzer.py`: `assign_statistical_roles` now uses the shared identifier
  evidence and checks temporal keys before identifiers. Formula ratios no longer
  override identifier or temporal roles.
- `core/models.py`: `AnalysisResult.semantics` and `AnalysisResult.quality`,
  serialized in `to_dict()`.
- `app/workflow.py`: semantics and quality are built after table detection and
  before statistics.
- `generic_report.py` / `report_writer.py`: audit-only «نقش‌های معنایی» and
  «کیفیت داده (تفصیلی)» sheets (the latter includes a block column); the clean
  «اعتبارسنجی داده» sheet gains a compact quality summary and lists non-`PASS`
  findings only.
- `tests/test_semantic_quality.py` (new, 23 tests) covers the eight required
  cases with real openpyxl workbooks, plus administrative/vehicle/region/record and
  plural identifiers, English and Arabic-Indic digits, all-text measure columns,
  category headers that mention a measure word, metadata evidence, cross-source period
  conflicts, quality on every block, role coverage, row roles, evidence and serialization.
  `tests/test_output_quality.py` guards the audit-only sheets.

### Test evidence [SBX]

Sandbox: Linux, CPython 3.13.15. Not Windows-target validation. Runtime,
document and development dependencies were installed from the repository
requirements files before this run.

```text
python -m pytest tests/test_semantic_quality.py -o addopts="" -q
-> 23 passed in 1.35s (exit 0)

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 223 passed, 2 skipped, 36 deselected, 1 xfailed in 25.58s (exit 0)
```

The single `xfailed` outcome is still T9, the open output-suffix defect. The 36
deselected tests are the `requires_sample` set; their private reference workbook
is absent and they are not counted as passing.

### Not validated

- The owner's Windows environment and the three original administrative files.
- Real-file acceptance for the new semantic and quality checks; the evidence uses
  synthetic `.xlsx` files created by the test suite.

## Prompt 02B — Statistical reconciliation and calculation integrity (branch evidence)

Branch `feature/prompt-02b-statistical-reconciliation`, based on `main` at `d173465`
(the merge of Prompt 02A). No aggregation was added to the request-answering path;
the new engine validates relationships that already exist in the workbook.

### Implementation

- `excel/statistical_reconciliation.py` (new): deterministic `Decimal` arithmetic for
  component-to-total, subtotal-to-grand-total, group-to-subgroup, monthly-to-period and
  cross-period checks. `AggregationEvidence` rejects forbidden roles (grand total,
  subtotal, percentage, difference), duplicate categories, overlapping category tags
  and mixed units before a sum is calculated; note/footer rows are excluded by the
  shared row classifier. A missing component is `WARNING` with `calculated=None`,
  never zero.
- `compute_change` records absolute change, percentage change, direction and baseline
  status. A zero or missing baseline yields `percent=None` with `WARNING`; no division
  by zero and no fabricated percentage. `compute_ratio` and `compute_average` refuse a
  missing denominator; a ratio column with no detectable weight column, or with more
  than one plausible weight column, yields `WARNING` and no average, and a simple mean
  is never reported as the weighted result.
- `Tolerance.for_workbook` combines an absolute rounding tolerance from the component
  count with a small relative tolerance for large monetary totals; `PASS`, `WARNING`
  and `FAIL` are explicit. Every `ReconciliationResult` keeps expected, calculated,
  difference, tolerance, status, explanation, sheet, block, measure and source
  addresses; `verified` is true only for a justified relationship with a real expected
  value.
- A percentage/rate column is excluded from component sums (a percentage is not a
  count). A grand total with prior subtotals also includes trailing data rows after the
  last subtotal. Incomplete months are excluded from monthly reconciliation with a
  `WARNING`; missing months and missing declared totals are reported rather than
  skipped. An unusable period matrix or two periods without a common complete month
  adds a report warning.
- No relationship at all is not a `PASS`: the report status becomes `WARNING` and the
  clean/audit views show the warning text. The clean view also lists non-`PASS`
  averages, so a missing weight is never silent.
- `app/workflow.py`: `AnalysisResult.statistical_reconciliation` is built after
  semantics, quality and the period matrix. `core/models.py` serializes it into the
  processed JSON.
- `generic_report.py` / `report_writer.py`: the audit profile writes «تطبیق آماری و
  صحت محاسبات»; the clean «اعتبارسنجی داده» sheet gains a compact reconciliation
  summary plus non-`PASS` reconciliation, comparison and average findings and the
  report warnings.
- `ai/context_builder.py`: reconciliation facts and warnings are included in the AI
  context and all reconciliation values are added to `allowed_numbers`.
  `ai/analyzer.py` treats any statistical `FAIL` (reconciliation, comparison or
  average) as a real mismatch. Qwen may explain the recorded figures but cannot
  introduce a new number without failing the existing number validator.
- `tests/test_statistical_reconciliation.py` (new, 36 tests) covers the nine required
  cases with real openpyxl workbooks plus direct engine calls: correct and incorrect
  component totals, subtotal/grand-total including trailing data after the last
  subtotal, nested and overlapping categories, zero baseline, missing denominators,
  weighted versus unweighted averages, rounding tolerances (absolute and relative),
  increase/decrease interpretation, mixed units from a «واحد» column, percentage
  totals that must not be reconciled as counts, ambiguous weight columns,
  monthly-to-period reconciliation, incomplete/empty months, missing declared totals,
  cross-period comparison, serialization, clean/audit output profiles, report warnings
  and the AI number allow-list. `tests/test_output_quality.py` guards the new audit-only
  sheet.

### Test evidence [SBX]

Sandbox: Linux, CPython 3.13.15, pytest 9.2.0. Not Windows-target validation.

```text
python -m pytest tests/test_statistical_reconciliation.py -o addopts="" -q
-> 36 passed in 3.36s (exit 0)

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 259 passed, 2 skipped, 36 deselected, 1 xfailed in 29.21s (exit 0)
```

The single `xfailed` outcome is still T9, the open output-suffix defect. The 36
deselected tests are the `requires_sample` set; their private reference workbook is
absent and they are not counted as passing. The two skips are the PySide6 GUI tests
(`libEGL.so.1` is unavailable in this sandbox) and are not counted as passing.

### Not validated

- The owner's Windows environment and the three original administrative files.
- Real-file acceptance for the new reconciliation checks; the evidence uses synthetic
  `.xlsx` files created by the test suite.
- Flat parent/child categories and overlapping administrative classifications are not
  inferred from labels alone. `reconcile_hierarchy` and
  `AggregationComponent.categories` cover those structures when the caller supplies
  them; automatic block reconciliation handles declared totals, subtotals, group
  headings, units and percentages.
- Category-level comparisons are exposed through `compare_categories` for aligned
  category maps; the workbook engine does not guess category pairs from arbitrary
  layouts. Group-to-subgroup comparisons are exercised directly and through
  group-heading blocks.
- A ratio column with no weight column, or with more than one plausible weight column,
  is reported as `WARNING` with no average; the engine does not pick a denominator by
  position.

## Prompt 03A — Semantic Source Matching and Period Alignment

Status: implemented on `feature/prompt-03a-semantic-source-matching`; branch pushed and
PR #8 open, not merged to `main`. Base: `main` at `891b662` (PR #7 merged).

### What was added

- `excel/period_alignment.py` (new): `NormalizedPeriod` and `parse_period` /
  `parse_periods`. It recognizes Jalali and Gregorian years, months, days, numeric
  dates (`1403/04`, `1403/04/15`, `2024-07-05`), named months (`تیر 1403`,
  `July 2024`), consecutive multi-month periods (`مرداد و شهریور 1403`,
  `January and February 2024`), quarters, halves, seasons, explicit ranges
  (`از ... تا`, `..`, ` - `), compact ranges (`1403/01/01-1403/03/31`,
  `1403-1404`) and embedded dates in headers (`تعداد پرونده 1403/04`).
  Persian/Arabic-Indic digits are normalized first. A month name paired with a year
  from the other calendar (`تیر 2024`, `July 1403`) is not a period. An exact bare
  year is accepted inside the converter band (1178–1633 Jalali, 1700–2200
  Gregorian); a bare year embedded in a short title is limited to 1300–1500 Jalali /
  1900–2100 Gregorian. A long free-text cell and a code such as `کد 1234` are not
  periods; `ردیف 12` is not a period. A code number next to a real year is ignored,
  and multiple plausible years yield no period. A multi-month phrase without an
  ordinal (`سه ماهه 1403`) or several non-consecutive month names (`مرداد و مهر
  1403`) yields no period, never a bare year. The matcher consumes the finalized `SheetIntelligence.periods`, so a month
  block title under a year-only sheet name becomes `مرداد 1404`; if only a bare year
  remains for a month title, the period is `not_detected` rather than equated with
  another month or invented. `compare_period_sets` returns an explained relationship
  (`equivalent`, `contains`, `within`, `overlapping`, `incompatible`, `unknown`)
  with the real overlap dates.
- `excel/semantic_matching.py` (new): controlled cross-source matching. Every
  `SourceField` keeps source file, sheet, block, column letter, original header,
  normalized label, canonical tokens, role, unit (dimension/canonical/scale), periods,
  table title, row labels, samples, shared identifier columns and identifier values,
  and evidence. Label evidence uses canonical tokens (spelling, ZWNJ, Persian plural
  forms, a small cross-language administrative vocabulary, configured synonyms) and
  separates content words from measure words, so `تعداد پرونده` matches `Case Count`
  but `تعداد پرونده ورودی` is not `تعداد پرونده خروجی`. Aggregation words are
  identity-bearing: `مبلغ` is not matched to `مبلغ کل` or `میانگین مبلغ`. Units are
  compared by dimension and scale; `ریال` versus `نفر` is rejected, `میلیون ریال`
  versus `ریال` is the same indicator but not directly comparable. Periods are
  compared as date ranges, never as strings. Row-label overlap, shared identifier
  values and table-title similarity are recorded as supporting evidence, so a
  reordered table with an extra descriptive column can still align on the same rows.
- Identity and value comparability are separate: a field pair is `matched` when the
  normalized label, role and unit identity agree; `comparable` is true only when units
  are equal (or both absent) and periods are equivalent. Every field and decision also
  carries an explicit period status (`detected`, `not_detected`, `conflicting` or
  `partial`), so an absent period is reported rather than invented. The same indicator in two
  different periods stays mapped with `comparable=False`; a missing period yields
  `probable`, not `matched`. Two or more plausible candidates produce one `ambiguous`
  record that keeps every candidate. A category or identifier pair whose names differ
  but whose sampled values overlap at least 80% is `probable` with `comparable=False`,
  never a definitive match. Similar names that describe different indicators are
  `rejected` with the reason. The engine never aggregates.
- `semantic_classification.py`: `SemanticField` now carries the column period and
  sample values; `SemanticBlock` carries row labels taken from the primary
  identifier/category column (descriptive columns such as `شرح` are skipped).
  `semantic_vocabulary.identifier_evidence` no longer treats a weak identifier word
  as an identifier when a measure word is present, so `تعداد پرونده` stays a count
  measure while `شماره پرونده` stays an identifier.
- `app/workflow.py` / `core/models.py`: `AnalysisResult.semantic_matching` is built
  after semantics, quality and structure intelligence (so sheet-level period evidence is
  available) and serialized into the processed JSON.
- `app/folder_workflow.py`: `FolderResult.matching` is built with
  `match_across_sources` over every analyzed file; the folder report gains
  «تطبیق شاخص‌ها بین منابع» and a summary line.
- `report_writer.py` / `generic_report.py`: the audit profile writes
  «تطبیق معنایی منابع» with matched/probable/ambiguous/rejected decisions, Persian
  period/unit labels, period status, comparability and reasons; the clean profile never
  writes it.
- `tests/test_semantic_matching.py` (new, 23 tests) covers the seven required cases
  with real openpyxl workbooks: different names, similar names for different
  indicators, different period formats, Persian/Gregorian date references, different
  unit dimensions, unit scale, missing periods, ambiguity, category-value overlap
  (now a probable match), shared identifier values, aggregation-label rejection, a configured synonym mapping equivalent categories,
  reordered tables with an extra descriptive column, conflicting periods in one field,
  month titles under a year-only sheet name (including two month blocks in one sheet),
  multi-month named periods, multi-period sets that require full coverage, the period
  parser, the audit-only report sheet, serialization and cross-file folder matching.
- `tests/test_output_quality.py` now guards «تطبیق معنایی منابع» as audit-only.

### Test evidence [SBX]

Sandbox: Linux, CPython 3.13.15, pytest 9.2.0. Not Windows-target validation.

```text
python -m pytest tests/test_semantic_matching.py -o addopts="" -q
-> 23 passed in 2.11s (exit 0)

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 278 passed, 6 skipped, 36 deselected, 1 xfailed in 31.66s (exit 0)
```

The single `xfailed` outcome remains T9 (output suffix validation). The 36 deselected
tests are the `requires_sample` set; the private reference workbook is absent. The six
skips are optional GUI/monitoring tests whose system libraries are unavailable in this
sandbox. None is counted as passing.

### Not validated

- The owner's Windows environment, the three original administrative files and real
  cross-file acceptance; all evidence uses synthetic `.xlsx` files.
- The matching vocabulary is deliberately small. A domain synonym not present in
  `config/synonyms.json` or the built-in administrative vocabulary will be reported as
  `ambiguous`/`rejected` rather than guessed.
- A Jalali month and a Gregorian month with the same name do not have identical date
  ranges, so they are `overlapping`, not `equivalent`; only identical ranges are
  comparable. An explicit day or range that converts to the same dates is equivalent.
- Matching is pairwise. Very large folders (hundreds of files with many fields) have no
  field cap yet and can be slow; the report records decisions, not an exhaustive index.
- No unrestricted aggregation is performed in this stage. `comparable=True` only means
  the two values may be compared; later stages must still apply the declared-total and
  reconciliation rules from Prompt 02B.
