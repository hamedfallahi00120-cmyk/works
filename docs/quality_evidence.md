# Quality evidence ledger — Local Office AI

This ledger records the verified handoff state for the next engineering stage.
The preparation baseline was Phase-1a preparation only. **Prompt 01A (Data Model, Workbook
Structure, and Metadata Extraction) is implemented and merged to `main` as PR #4**, and
its evidence is recorded at the end of this file. Prompt 01B (Multi-Block Extraction,
Table Boundaries, and Total Detection) is implemented on
`feature/prompt-01b-multiblock-extraction`; its evidence is recorded at the end of this
file.

## Verified repository state

| Item | Value |
| :--- | :--- |
| Repository | `hamedfallahi00120-cmyk/works` (private) |
| Phase-0 base | `main` at `6df625659c97c0b7da1f78f33fa5fc37864a7580` |
| Prompt 01A merge | PR #4; `main` at `f572c47` |
| Preparation branch | `fix/multisource-extraction-integrity` |
| Preparation commits | `5b20cce` (tests + ledger), `06dae62` (sanitized handover + audit), followed by CI/documentation-preflight commits |
| Phase 0 | Completed; see `PHASE0_BASELINE_AUDIT.md` |
| Original Prompt 01A | Implemented and merged as PR #4 (evidence at end) |
| Prompt 01B | Implemented on `feature/prompt-01b-multiblock-extraction` (evidence at end) |
| Production source fixes in the preparation baseline | None; Prompt 01A and Prompt 01B add production source changes (below) |

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
- The private source workbook used by `requires_sample` tests. The 38 marked tests are excluded from the CI gate, not silently treated as passed.

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
  `conflicting_periods` and `suspicious_values`. Every check has an explicit
  status, location, count and samples; values are never repaired.
- `sheet_analyzer.py`: `assign_statistical_roles` now uses the shared identifier
  evidence and checks temporal keys before identifiers. Formula ratios no longer
  override identifier or temporal roles.
- `core/models.py`: `AnalysisResult.semantics` and `AnalysisResult.quality`,
  serialized in `to_dict()`.
- `app/workflow.py`: semantics and quality are built after table detection and
  before statistics.
- `generic_report.py` / `report_writer.py`: audit-only «نقش‌های معنایی» and
  «کیفیت داده (تفصیلی)» sheets; the clean «اعتبارسنجی داده» sheet gains a compact
  quality summary and lists non-`PASS` findings only.
- `tests/test_semantic_quality.py` (new, 13 tests) covers the eight required
  cases with real openpyxl workbooks, plus role coverage, row roles, evidence and
  serialization. `tests/test_output_quality.py` guards the audit-only sheets.

### Test evidence [SBX]

Sandbox: Linux, CPython 3.13.15. Not Windows-target validation.

```text
python -m pytest tests/test_semantic_quality.py -o addopts="" -q
-> 13 passed in 0.58s (exit 0)

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 209 passed, 6 skipped, 36 deselected, 1 xfailed in 23.12s (exit 0)
```

The single `xfailed` outcome is still T9, the open output-suffix defect. The 36
deselected tests are the `requires_sample` set; their private reference workbook
is absent and they are not counted as passing.

### Not validated

- The owner's Windows environment and the three original administrative files.
- Real-file acceptance for the new semantic and quality checks; the evidence uses
  synthetic `.xlsx` files created by the test suite.
