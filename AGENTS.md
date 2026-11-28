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
(`f572c47`). Prompt 01B was merged as PR #5 (`113cc1f`), Prompt 02A as PR #6
(`d173465`) and Prompt 02B as PR #7 (`891b662`). The dedicated Prompt 03A branch is
`feature/prompt-03a-semantic-source-matching`. Never work directly on `main`.

The preparation baseline is **Phase-1a preparation only**: isolated sandbox evidence,
four focused tests (T2/T3/T4/T9), the quality ledger, the Phase 0 audit, and the
sanitized handover. Prompt 01A is built on top of that baseline and is now merged.

Prompt 01A is implemented and merged:
- T2 (G3/R3) is fixed: totals keywords match on word boundaries, so `کلانتری ۱۲`,
  `تکلیف`, `جمعیت` and `کلاسه پرونده` are ordinary labels while `جمع کل` still totals.
- T3 (G5/R4) is fixed: `سال`, `ماه` and `روز` columns whose content agrees
  (integer years, months 1..12, days 1..31) become `DataRole.TEMPORAL_KEY`, never measures.
- T9 (G11/R8) is fixed: a `.docx` output is a real Word package. Prompt 04B builds a
  standalone Word report when no template is supplied, and never writes Excel content
  under a Word name.

T2 and T3 no longer carry xfail markers; they must keep passing normally. Do not
reintroduce substring totals matching or let temporal keys back into numeric measures.

Prompt 01B was merged to `main` as PR #5 (`113cc1f`). Prompt 02A was merged to `main`
as PR #6 (`d173465`). Prompt 02B was merged to `main` as PR #7 (`891b662`). Prompt 03A
is implemented on `feature/prompt-03a-semantic-source-matching`; the branch is pushed to
origin and PR #8 is open, but it is not merged to `main`. Prompt 03B was merged to
`main` as PR #9 (`18fb3d4`), Prompt 04A was merged as PR #10 (`4741a98`), and
Prompt 04B was merged as PR #11 (`8f77c25`). **Prompt 05A is implemented on
`feature/prompt-05a-quality-gate-provenance`; PR #12 is open and it is not merged to
`main` yet.**
See the sections below.

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

### Prompt 02B statistical reconciliation and calculation integrity

- `excel/statistical_reconciliation.py` — deterministic `Decimal` engine for
  component-to-total, subtotal-to-grand-total, group-to-subgroup, monthly-to-period and
  cross-period checks. `AggregationEvidence` rejects grand-total, subtotal, percentage
  and difference roles, overlapping category tags and mixed units; note/footer rows are
  excluded by the shared row classifier. A missing component stays `None` and yields
  `WARNING`, never zero. `compute_change` records absolute change, percentage change,
  direction and a zero/missing baseline without dividing by zero. `compute_ratio` and
  `compute_average` refuse a missing denominator; a ratio column with no detectable
  weight column yields `WARNING` and no simple mean, and a weighted average is used
  whenever a weight column exists. `Tolerance.for_workbook` combines rounding and
  relative tolerance, and every result records expected, calculated, difference,
  tolerance, `PASS`/`WARNING`/`FAIL`, explanation, sheet, block, measure and source
  addresses.
- `app/workflow.py` — reconciliation is built after semantics, quality and the period
  matrix. `AnalysisResult.statistical_reconciliation` is serialized into the processed
  JSON.
- The audit profile writes «تطبیق آماری و صحت محاسبات»; the clean «اعتبارسنجی داده»
  sheet shows a compact summary plus non-`PASS` reconciliation, comparison and average
  findings and report warnings.
- `ai/context_builder.py` includes reconciliation facts and warnings and adds every
  approved value to `allowed_numbers`; `ai/analyzer.py` treats any statistical `FAIL`
  (reconciliation, comparison or average) as a real mismatch. Qwen explains recorded
  results and may not recalculate them.
- `tests/test_statistical_reconciliation.py` (36 tests) covers the nine required cases
  with real openpyxl workbooks: correct/incorrect component totals, subtotal/grand total
  (including data after the last subtotal), nested and overlapping categories, zero
  baseline, missing denominators, weighted versus unweighted averages, rounding
  tolerances (absolute and relative), increase/decrease interpretation, mixed units from
  a «واحد» column, percentage totals that must not be reconciled as counts, ambiguous
  weight columns, monthly-to-period reconciliation, incomplete/empty months, missing
  declared totals, cross-period comparison, serialization, clean/audit output profiles,
  warnings and the AI allow-list.
- Known limits: automatic block reconciliation relies on declared total/subtotal rows
  and the shared row classifier. Prompt 03B detects a probable parent-category row and
  excludes it from the automatic component sum, downgrading a matching result to
  `WARNING` for review. Overlapping administrative classifications without explicit
  tags still cannot be proven safe from labels alone; use `reconcile_hierarchy` with
  `HierarchyNode` or `AggregationComponent.categories` for those structures.
  `compare_categories` is a deterministic API for aligned category maps; the engine does
  not infer category pairs from every layout. A ratio column without a detectable weight
  column, or with more than one possible weight column, yields `WARNING` and no average
  because the denominator is not certain. An unusable period matrix, or two periods with
  no common complete month, adds a report warning instead of a silent omission.

### Prompt 03A semantic source matching and period alignment

- `excel/period_alignment.py` — `NormalizedPeriod`, `parse_period`, `parse_periods` and
  `compare_period_sets`. It parses Jalali/Gregorian years, months, days, numeric dates,
  named months, consecutive multi-month periods (`مرداد و شهریور 1403`,
  `January and February 2024`), quarters, halves, seasons, ranges (including compact
  `1403/01/01-1403/03/31` and `1403-1404`) and dates embedded in headers
  (`تعداد پرونده 1403/04`). Persian/Arabic-Indic digits are normalized first. A
  month name paired with a year from the other calendar (`تیر 2024`, `July 1403`) is
  not a period. An exact bare year counts inside the converter band (1178–1633
  Jalali, 1700–2200 Gregorian); a bare year embedded in a title is limited to
  1300–1500 Jalali / 1900–2100 Gregorian. A long free-text cell, `ردیف 12` and
  `کد 1234` are not periods; a code number next to a real year is ignored, and
  multiple plausible years yield no period. A multi-month phrase without an ordinal
  (`سه ماهه 1403`) or non-consecutive month names (`مرداد و مهر 1403`) yields no
  period, never a bare year. Comparison returns
  an explained relation (`equivalent`, `contains`, `within`, `overlapping`,
  `incompatible`, `unknown`) with real overlap dates. No period is invented.
- `excel/semantic_matching.py` — controlled cross-source matching. `SourceField` keeps
  source file, sheet, block, column letter, original header, normalized label, canonical
  tokens, role, unit (dimension/canonical/scale), periods, title, row labels, samples,
  shared identifier columns and identifier values, and evidence. Label evidence
  separates content words from measure words and uses a small cross-language
  administrative vocabulary plus configured synonyms. Aggregation words are
  identity-bearing: `مبلغ` is not matched to `مبلغ کل` or `میانگین مبلغ`. Units are
  compared by dimension and scale; periods by date range. A pair is `matched` when
  label/role/unit identity agree; `comparable` is true only when units are equal (or
  both absent) and periods are equivalent. Each field and decision carries an explicit
  period status (`detected`, `not_detected`, `conflicting`, `partial`). The same
  indicator in different periods stays mapped with `comparable=False`; a missing period
  is `probable`, never `matched`; a category/identifier pair with at least 80% sampled
  value overlap is `probable` with `comparable=False`; two close candidates are
  `ambiguous` with every candidate kept; similar names for different indicators are
  `rejected` with the reason. Row-label overlap, shared identifier values and
  table-title similarity are supporting evidence, so reordered tables with an extra
  descriptive column can still align. It never aggregates.
- `semantic_classification.py` — `SemanticField` carries the column period and samples;
  `SemanticBlock` carries row labels from the primary identifier/category column
  (descriptive columns such as `شرح` are skipped).
  `semantic_vocabulary.identifier_evidence` does not treat a weak identifier word as an
  identifier when a measure word is present, so `تعداد پرونده` is a count measure while
  `شماره پرونده` remains an identifier.
- `app/workflow.py` / `core/models.py` — `AnalysisResult.semantic_matching` is built after
  semantics, quality and structure intelligence (so sheet-level period evidence is
  available) and serialized into the processed JSON.
- `app/folder_workflow.py` — `FolderResult.matching` is built with
  `match_across_sources` over all analyzed files; the folder report writes
  «تطبیق شاخص‌ها بین منابع» and a summary line.
- `report_writer.py` / `generic_report.py` — the audit profile writes
  «تطبیق معنایی منابع» with Persian period/unit labels and period status; the clean
  profile never writes it.
- `tests/test_semantic_matching.py` (23 tests) covers the seven required cases with real
  openpyxl workbooks, category-value overlap (probable, not definitive), shared
  identifier values, aggregation-label rejection, a configured synonym mapping equivalent categories,
  reordered tables with an extra descriptive column, conflicting periods, month titles under a year-only sheet name
  (including two month blocks in one sheet), multi-month named periods, multi-period
  sets that require full coverage, the period parser, serialization, audit-only output
  and cross-file folder matching.
  `tests/test_output_quality.py` guards the audit-only sheet.

Known limits: the matching vocabulary is deliberately small; an unknown domain synonym
is reported as `ambiguous`/`rejected` rather than guessed. A Jalali month and a Gregorian
month with the same name are `overlapping`, not `equivalent`, because their date ranges
differ; only identical ranges are comparable. Matching is pairwise and has no field cap
yet, so very large folders can be slow. `comparable=True` only permits a later comparison;
it does not aggregate, and Prompt 02B reconciliation rules still apply.

### Prompt 03B duplicate detection, derived sheets and safe aggregation

- `excel/duplicate_control.py` — deterministic `Decimal`/hash engine that records
  repeated content, copied summaries, formula-derived sheets, parent/child categories,
  subtotal/grand-total rows and overlapping administrative groups before any cross-source
  total. It never deletes, repairs or aggregates a value.
- `BlockProfile` — per-block fingerprints (`table_hash`, `content_hash`, `summary_hash`),
  trimmed row signatures, key/category values, record keys, summary rows, formula and
  cross-sheet-reference counts and the block's primary period. Positional `ردیف`/`row`
  columns are not record identifiers. A summary row enters the summary fingerprint only
  when at least one measure cell has a value, so an empty formula total is not mistaken
  for a copied summary.
- `SheetRole` — `primary`, `summary`, `derived`, `duplicate`, `unknown`, each with
  confidence and evidence. Formula evidence comes from whole-sheet formula scans
  (`SheetIntelligence.formula_cells` / `cross_sheet_formula_cells`), so a small summary
  sheet that never becomes a table is still classified. A sheet-name/title summary word
  is a probable `summary` (0.55) only when no formula/duplicate evidence exists.
- `DuplicateGroup` — exact content groups and summary-fingerprint groups are split by
  reporting period, so two months with equal totals are not called duplicates. Near
  duplicates require ≥75% header-token overlap and ≥80% row overlap. `safe_to_aggregate`
  is always `False`.
- `CategoryRelation` — `parent_of` (probable, `confirmed=False`), `grand_total_of`
  (confirmed when the declared total equals the collapsed child rows) and `overlap`
  (shared identifier/category values, only for period-compatible blocks). Nested parents
  are collapsed before a grand total so a parent is never added to its children.
- `AggregationDecision` — one decision per field pair with six explicit checks:
  indicator equivalence, period compatibility, unit compatibility, category
  relationship, source independence and non-overlapping coverage. `ALLOWED` requires all
  six `PASS`; any `BLOCK` makes it `BLOCKED`; the rest is `REVIEW`. `comparable=False`
  keeps coverage at `REVIEW`, and a pair inside one block is never independent.
- `DuplicateControlReport` — sheets, blocks, groups, record duplicates, relations,
  decisions, exclusions and warnings; `cross_source_status(metric, sources)` only
  considers decisions whose two members come from different files. Serialized as
  `AnalysisResult.duplicate_control` into the processed JSON.
- `app/folder_workflow.py` — `FolderResult.duplicate_control` is built after semantic
  matching. The folder report writes «کنترل تکرار و تجمیع بین فایل‌ها» and the
  «مقایسه بین فایل‌ها» sheet suppresses «جمع فایل‌ها»/share with a status when the
  contributors' cross-source status is not `ALLOWED`; a single contributing file is
  always safe because nothing is added across files. «پاسخ درخواست» omits the share for
  an unsafe cross-file sum but keeps the most/least comparison.
- `report_writer.py` / `generic_report.py` — audit-only «کنترل تکرار و تجمیع» with sheet
  roles, duplicate groups, record duplicates, relations, decisions and every exclusion;
  the clean profile never writes it. The clean «اعتبارسنجی داده» shows a compact role /
  decision summary, the duplicate-control warnings and up to 40 blocked/review decisions.
- `structure_intelligence.py` — `formula_cells` and `cross_sheet_formula_cells` on every
  `SheetIntelligence`, computed from the whole sheet even when no table is detected.
- `statistical_reconciliation.py` — a probable/confirmed parent-category row recorded by
  duplicate control is excluded from the automatic component sum, so a parent is never
  added to its children. A matching result with a probable relation is `WARNING`, not
  `PASS`, and its explanation names the excluded parent rows.
- `tests/test_duplicate_control.py` (18 tests) uses real openpyxl workbooks and covers
  exact duplicate tables in different worksheets, copied summaries across workbooks,
  formula-based summary sheets, parent/child categories, subtotal and grand-total
  duplication, overlapping administrative classifications, independent sources with the
  same metric, similar-looking but independent datasets, a derived sheet without summary
  labels, copied total rows with different detail, record-level duplicates (kept, not
  removed), identical values in different periods, a mixed sheet with one duplicate and
  one unique block (stays primary), a cross-sheet total formula in a data sheet (stays
  primary), an unknown sheet, parent rows excluded from automatic reconciliation,
  selected-file decision scope, serialization and the audit/clean output profiles.
  `tests/test_output_quality.py` guards the audit-only sheet.
- Known limits: duplicate/overlap detection is fingerprint- and period-based; a real
  duplicate whose values were edited slightly is `REVIEW`, not proof. Without cached
  formula values (a file never opened by Excel) a total row has no fingerprint, so it is
  not grouped. A formula-derived sheet is recognised when cross-sheet formulas cover at
  least half of its data rows; a data sheet with only a total formula stays primary.
  Formula references are matched by sheet name; external-workbook links are not followed. Row-number columns are excluded from identifiers, so a table with no real
  key gets `REVIEW` coverage rather than a guessed overlap. The folder comparison refuses
  a cross-file total only when its decision is not `ALLOWED`; it does not rewrite any
  per-file figure.

### Prompt 04A output-format routing and DOCX template architecture

- `reporting/formats.py` — `OutputFormat` (`XLSX`, `DOCX`, `PDF`), `OutputIntent`,
  `detect_requested_formats`, `resolve_output_intent`, `ensure_extension_matches` and
  `ensure_xlsx_delivery`. Format is resolved once from explicit `--format`, the output
  suffix, the Persian/English request text and the configured default. Contradictory
  signals raise `ReportFormatError`; a bare «فایل اکسل» describes the input, so only
  «گزارش/خروجی اکسل» counts as an output request. PDF is recognised but generation is
  explicitly unsupported.
- `reporting/specification.py` — `ReportSpecification` (format, path, profile, optional
  template, `preserve_template_figures`), `ReportContent`/`ReportValue`/`ReportSection`,
  and `ReportGenerationResult`. An explicit DOCX template must be readable, and the
  output path may not equal the template.
  `ReportContent.from_analysis` builds values only from the validated `AnalysisResult`
  (calculations, descriptive statistics, request answers, warnings, anomalies) and keeps
  each value's source and address. It never reads a template.
- `reporting/template.py` — `inspect_docx_template` records page size, orientation,
  margins, header/footer distances, gutter, section/paragraph bidi, RTL runs, fonts,
  paragraph styles, heading hierarchy, tables (style, header row, `bidiVisual`), headers
  and footers per section/variant, numbering definitions and every `{{...}}` placeholder
  with its location. `ReportTemplate.is_data_source` is always `False`; static numbers
  found in the template are reported as formatting artifacts, never as report data.
- `reporting/docx_writer.py` — `DocxTemplateRenderer` copies the template and replaces
  placeholders in body paragraphs, table cells and header/footer paragraphs. It preserves
  sections, margins, styles, heading levels, table styles, numbering and RTL settings;
  a placeholder split across runs is merged into the first run so that run's formatting
  is kept. Unknown placeholders stay untouched and are reported. Output is saved through
  an atomic temp file and checked as a valid DOCX package.
- `reporting/generator.py` — `ReportGenerator.plan` selects exactly one writer
  (`ExcelReportWriter` or `DocxReportWriter`) and `generate` routes to it.
  `RunResult.report` carries the `ReportGenerationResult`. `ReportTemplate.inspect(...)`
  is a convenience alias for `inspect_docx_template`.
- `app/workflow.py` / `coordinator.py` / `ui/cli.py` — `--format`, optional `--template`,
  `--keep-template-figures` and `--inspect-template`; the workflow resolves the intent
  before analysis. `settings.report_format`
  and `settings.docx_template` provide defaults; the configured template applies only to
  Word output, so it cannot break an Excel run, while an explicit `--template` with a
  non-Word format is rejected. `app/folder_workflow.py` and
  `app/document_set_workflow.py` are Excel-only and reject DOCX/PDF requests with an
  explicit error; the CLI also rejects `--folder --template`.
- `excel/report_writer.py`, `excel/mixed_report_writer.py` and
  `FolderReportWriter.build_folder` validate `.xlsx` at the writer boundary, so a
  mismatched suffix cannot reach a workbook writer.
- `tests/test_report_output_formats.py` (23 tests) uses real openpyxl workbooks and real
  python-docx templates: format tokens/suffixes, explicit/suffix/request/default routing,
  conflicts, extension validation, PDF rejection, template inspection (layout, margins,
  styles, headings, tables, headers/footers, numbering, RTL), placeholder detection across
  body/table/header/footer, style/layout preservation after rendering, unknown
  placeholders, missing/corrupt/unsupported templates, template/data separation, a
  configured Word template that must not break Excel output, a folder run that rejects
  `--template`, zero-valued request answers, template preservation when an output name
  would collide with the old temp name, and CLI DOCX generation. Input-file extensions
  (`report.docx`, «فایل docx») are not mistaken for output requests. T9 now passes
  normally; its strict xfail marker was removed.
- Known limits of this stage: it defines the contract and a placeholder renderer; Prompt
  04B adds the native layout writer. A template is optional for Word output. A placeholder
  that crosses a run containing a tab, line break or field is left untouched and warned.
  Placeholders inside hyperlinks, text boxes, footnotes or comments are not rewritten;
  body paragraphs, table cells and headers/footers are covered. Folder and multi-source
  reports remain XLSX-only. PDF is recognised but not generated.

### Prompt 04B DOCX report writer and Persian RTL rendering

- `reporting/docx_rtl.py` — real OOXML RTL primitives: section `w:bidi`/`w:rtlGutter`,
  paragraph `w:bidi`, run `w:rtl`/`w:bCs`, complex-script fonts (`w:rFonts/@w:cs`) and
  `w:szCs`, RTL tables (`w:bidiVisual`), repeated header rows (`w:tblHeader`), header
  shading, built-in `PAGE`/`NUMPAGES` fields and header/footer iteration. Alignment alone
  is never treated as RTL.
- `reporting/docx_report_writer.py` — `DocxReportWriter` with two modes. Standalone builds
  a complete A4 report: title page, styled Persian headings, executive summary, RTL
  tables with captions, source/statistical notes, page breaks, header/footer with real
  page numbers, warnings, conclusions and recommendations. Template mode preserves the
  template's sections, margins, styles, headings, tables, numbering and RTL settings,
  replaces placeholders, strips static historical figures unless
  `strip_static_figures=False`, inserts the body at `{{content}}`/`{{body}}`/`[[CONTENT]]`
  or appends it, adds generated content under a heading the template already contains
  instead of skipping the section, and fills a table whose header row matches a template
  table instead of duplicating it. Every generated file is reopened and validated before
  it is returned; an invalid file is never delivered.
- `reporting/docx_validation.py` — `validate_docx_output`/`ensure_valid_docx`: confirms
  the DOCX ZIP members, parses every XML part, reopens with python-docx, checks required
  headings, the required table header rows and count, RTL sections/runs/tables, header/footer, a
  real PAGE field, unresolved placeholders and that every numeric claim outside Word
  fields is in `ReportContent.approved_numbers()`. Word field numbers are excluded, benign
  template numbering is optional, and a preserved template figure can be allowed
  explicitly.
- `reporting/specification.py` — `ReportContent.from_analysis` now also emits metric,
  descriptive-statistics, request-answer, quality and reconciliation tables plus identity,
  issue and answer values, all with source addresses. `placeholder_values()` and
  `approved_numbers()` are the only bridge a template or narrative may use. AI narrative
  is admitted only when every number it quotes is approved; a dropped narrative or finding
  is reported as a warning.
- `reporting/formats.py` / `generator.py` — `OutputFormat.DOCX.writer_name` is
  `DocxReportWriter`; `requires_template` is `False`. `ReportSpecification` accepts an
  optional template and `preserve_template_figures`; `ReportGenerator` passes
  `strip_static_figures=not preserve_template_figures`.
- `app/workflow.py` / `coordinator.py` / `ui/cli.py` — `--keep-template-figures` keeps
  template figures only when explicitly requested; a Word request without a template now
  builds a standalone report instead of failing.
- `tests/test_docx_report_writer.py` (25 tests) uses real openpyxl workbooks and real
  python-docx templates: valid DOCX packages, Persian paragraph RTL (bidi/rtl/cs/szCs and
  `w:lang/@w:bidi`), RTL tables with repeated shaded headers, template margins/styles/
  headings/tables/header/footer preservation, placeholder replacement, historical-figure
  removal and explicit preservation (including formatted runs), no-anchor append,
  generated content under an existing heading, reuse of a matching template table,
  numeric grounding (including tampered output and ungrounded Qwen text), manual table/
  header validation, missing
  templates, invalid output paths (wrong suffix, PDF, output=template, directory, nested
  creation), reopen/validate, generation-result serialization and CLI end-to-end Word runs
  with and without a template; an anchor with surrounding template text keeps that text.
- Known limits: the writer targets Word's built-in styles and a small explicit typography
  set; it does not embed custom fonts (the Windows target must have B Nazanin or Word
  substitutes it). Template placeholders inside hyperlinks, text boxes, footnotes and
  comments remain outside the replacement pass (Prompt 04A limit). An existing template
  heading is not written twice: generated content is added under it, while a matching
  template table is filled by replacing its data rows. The match is exact on normalized
  header cells, so a template table with different headers receives a new table instead;
  a table whose header row had merged cells may not be reused. Folder and multi-source
  reports remain XLSX-only. PDF is recognised but not generated. The Windows rendering
  itself cannot be verified on Linux; tests assert the OOXML properties Word reads.

### Prompt 05A quality gate and data provenance

- `quality/models.py` — one `PASS` / `WARNING` / `FAIL` model for all thirteen stages:
  source availability, file readability, extraction completeness, block detection,
  period identification, semantic classification, data quality, statistical validity,
  duplicate/overlap detection, cross-source reconciliation, provenance completeness,
  output-format correctness and final document integrity. `GateFinding`,
  `GateStageResult` and `QualityGateReport` carry status, checks, evidence, blocking
  failures and the provenance ledger. A warning never sets
  `QualityGateReport.is_complete_success`; a blocking `FAIL` sets `blocked=True` and
  `can_claim_valid_report=False`.
- `quality/provenance.py` — `ProvenanceLedger` / `ProvenanceRecord` /
  `ProvenanceSource` built from values Python already calculated. A record keeps the
  source filename, sheet, block, cell or range, original and normalized labels, original
  and normalized values, period, unit, operation, formula/transformation, source cells
  with their raw values and formulas, related records and validation status. Declared
  totals are source records; calculations, generic statistics, period totals, request
  answers, reconciliations, comparisons and averages are derived records. A derived
  record must point to at least one source cell/range; a missing required field or a
  derived value without a source is a blocking provenance failure. Loan defaults are not
  treated as figures of a non-loan workbook.
- `quality/gate.py` — `QualityGate.evaluate` reuses the existing `DataQualityAnalyzer`,
  `StatisticalReconciliation`, `DuplicateControlAnalyzer`, extraction model, semantics,
  period evidence and `ProvenanceLedger`; it does not create a second validator.
  `QualityGate.evaluate_folder` applies the same thirteen stages to multi-file reports,
  merges per-file ledgers and refuses unsafe cross-file totals. Stages 1–11 are evaluated
  before the writer reads the analysis; stages 12–13 after generation.
- `core/models.py` — `AnalysisResult.provenance`, `AnalysisResult.quality_gate`,
  `RunResult.quality_gate`; all serialize into the processed JSON. `FolderResult` carries
  the folder gate. `RunResult.status` is `failed` for a blocking gate failure, `completed`
  only for a clean `PASS` with unchanged input, and `completed_with_warnings` for a
  warning or non-blocking failure. Do not present a warning as complete success.
- `app/workflow.py`, `app/folder_workflow.py`, `app/coordinator.py` — attach the gate,
  record blocking failures in `errors`, and print the gate status, stage counts and the
  valid-report/complete-success claim in the console summary. The desktop result page,
  folder result page, banner and composer hint render a blocked/failed run as
  «ناموفق — گزارش معتبر نیست», never as an ordinary warning or completed report.
- `excel/report_writer.py`, `reporting/specification.py` — the audit profile writes
  «دروازه کیفیت» and «ردیابی منبع (شواهد)»; the clean profile never writes them and
  shows only a compact gate summary in «اعتبارسنجی داده». `ReportContent` carries the
  gate and provenance summary for the Word writer. A template never becomes a data
  source.
- `tests/test_quality_gate.py` (9 tests) covers complete provenance, missing cell
  references, derived values with source references, failed extraction, statistical
  validation failure, warning propagation, blocking critical failures, status
  consistency and merged folder/cross-source evidence. `tests/test_workflow_generic.py`
  and `tests/test_output_quality.py` now expect `completed_with_warnings` where a
  no-table/vague file or an unsafe cross-file total is not complete success.
- Known limits: a reconciliation that produced no numeric value is reported by the
  reconciliation engine, not the provenance ledger; a formula-only total with no cached
  value is `WARNING` (unverifiable), not a mismatch. A report with no data table is a
  warning/non-blocking failure, not a valid data analysis. Folder reports remain
  XLSX-only and their gate is attached after the folder workbook is written.

### Instructions for the next AI agent

1. Fetch the repository and start from the branch that contains the latest accepted
   prompt (see `docs/quality_evidence.md` for the verified state).
2. Verify the checked-out branch, HEAD SHA and clean working tree. Do not switch to
   `main` for implementation.
3. Read `docs/NEXT_AGENT_START_HERE.md`, `MASTER_PROJECT_HANDOVER_PUBLIC.txt`,
   `PHASE0_BASELINE_AUDIT.md`, this guide and `docs/quality_evidence.md`.
4. Treat the Phase 0 audit and its recorded evidence as the accepted project baseline.
   Do not repeat the full audit or ask the user to resend these repository documents.
5. Prompt 01A (**Data Model, Workbook Structure, and Metadata Extraction**), Prompt 01B
   (**Multi-Block Extraction, Table Boundaries, and Total Detection**), Prompt 02A
   (**Semantic Classification and Data Quality Validation**), Prompt 02B
   (**Statistical Reconciliation and Calculation Integrity**), Prompt 03A
   (**Semantic Source Matching and Period Alignment**), Prompt 03B
   (**Duplicate Detection, Derived Sheets, and Safe Aggregation**), Prompt 04A
   (**Output Format Routing and Word Template Architecture**), Prompt 04B
   (**DOCX Report Writer and Persian RTL Rendering**) and Prompt 05A
   (**Quality Gate Architecture and Data Provenance**) are implemented. Do not repeat
   them or substitute audit Phase 1a for them. The next engineering action starts only
   after the owner accepts the latest prompt.
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
  `statistical_reconciliation.py` validates declared totals, groups and periods without
  summing unrelated rows, and exposes explicit hierarchy/category comparison APIs for
  structures that cannot be inferred from a flat table.
  `statistical_analysis.py` is the general engine; `generic_report.py` holds the generic
  report sheets as a mixin. `validation.py` holds the 9-check quality checklist.
- `src/local_office_ai/ai/` — Ollama client, context builder, prompts, response validator,
  analyzer (model + deterministic fallback), ComfyUI image studio.
- `src/local_office_ai/quality/` — the central gate and evidence layer:
  `models.py` (`PASS`/`WARNING`/`FAIL`, thirteen stages, gate report, provenance
  records), `provenance.py` (ledger built from existing values and cells) and
  `gate.py` (`QualityGate.evaluate` / `evaluate_folder`, reusing the existing analyzers).
- `src/local_office_ai/reporting/` — output-format decision layer and DOCX template
  contract: `formats.py` (format resolution/validation), `specification.py`
  (`ReportSpecification`, validated `ReportContent`), `template.py` (DOCX inspection),
  `docx_writer.py` (template-mode placeholder rendering), `docx_report_writer.py`
  (standalone/template Word writer), `docx_rtl.py` (OOXML RTL primitives),
  `docx_validation.py` (reopen/validate a delivered DOCX) and `generator.py` (writer
  routing).
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
python run_engine.py --inspect-template templates/report.docx   # DOCX structure/styles/placeholders
python run_engine.py --file <xlsx> --format docx --template templates/report.docx --no-ai --yes
python run_engine.py --file <xlsx> --output report.docx --template templates/report.docx --no-ai --yes
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

## Output formats
- `OutputFormat` distinguishes `XLSX`, `DOCX` and `PDF`. Selection order is explicit
  `--format`, output suffix, request text, then `settings.report_format`. Contradictory
  signals are rejected instead of guessed.
- Word output is template mode: `--template` is required and the output must end in
  `.docx`. A Word request never falls back to an Excel workbook.
- PDF is recognised but generation is not implemented; it fails with a clear error.
- Folder and multi-source reports remain XLSX-only. Their workflows reject DOCX/PDF.
- Template figures are presentation artifacts. Report numbers come only from
  `ReportContent.from_analysis` / `AnalysisResult`.

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
- Resolve the output format before selecting a writer. Never write an Excel workbook under
  a `.docx`/`.pdf` name, never silently downgrade a Word request to Excel, and never use a
  template's historical figures as report data.
- A numeric identifier (postal code, personnel number, vehicle/region code, record number,
  year/month key) is never a measure. Use `semantic_vocabulary.identifier_evidence`; do not
  add a header-only ID check that ignores value patterns.
- Data-quality findings carry `PASS`/`WARNING`/`FAIL` and a sheet/cell location. Never
  delete, repair or coerce a questionable value. The clean profile lists non-`PASS`
  findings only; the audit profile keeps the full «کیفیت داده (تفصیلی)» sheet.
- Never sum rows just because they are numeric. A reconciliation needs a justified
  relationship: distinct non-overlapping components, one unit, no total/subtotal/
  percentage/difference row as a component, and a real expected value. A missing
  component stays `None` (`WARNING`), never zero. A zero or missing baseline yields
  `percent=None`, never a fabricated percentage. Use the weighted average when a ratio
  column and a weight column exist; never present the simple mean as the weighted result.
  Every result records expected, calculated, difference, tolerance, status, explanation
  and source addresses. A reconciliation `FAIL` is a real mismatch for the AI validator.
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
