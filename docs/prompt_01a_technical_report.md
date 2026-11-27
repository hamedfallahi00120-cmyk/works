# Prompt 01A — Technical Report

**Scope:** Data Model, Workbook Structure, and Metadata Extraction.
**Branch:** `feature/prompt-01a-data-model-metadata`. **Base:** verified preparation
baseline (`e7b491a200fee77c3bb394341c15d4efbf8ac408` is an ancestor). `main` is unchanged.

## 1. Problem

The engine could inspect a workbook, detect a primary table and build
`SheetIntelligence`, but the extracted structure was not a first-class, traceable object.
Periods were inferred from column headers only, filename evidence was ignored, and
year/month/day keys could still be read as numeric measures. Later stages had to
rediscover block boundaries, merged parents and period evidence.

## 2. Approach

Prompt 01A adds a projection layer on top of the existing detectors. It does not replace
`TableDetector`, `SheetAnalyzer` or `StructureIntelligence` and does not introduce a second
detection pipeline. The new layer only records what the existing detectors already found,
plus period evidence and cell-level provenance.

```
WorkbookReader -> WorkbookInspector -> TableDetector -> SheetAnalyzer/SheetIntelligence
                                                   \-> ExtractionModelBuilder
                                                          WorkbookModel
                                                            SheetModel
                                                              BlockModel
                                                                ColumnModel / CellRecord / MergedRangeRecord
```

`ExtractionModelBuilder.build(...)` receives the grids, detected structures and finalized
intelligence from `ExcelIntelligenceWorkflow`. It is attached as
`AnalysisResult.extraction` and serialized by `AnalysisResult.to_dict()` into the processed
JSON sidecar. The delivered workbook is untouched unless the audit profile is active.

## 3. Data model

| Level | Preserved |
| :--- | :--- |
| `WorkbookModel` | source filename, source path, sheet models, workbook periods, evidence, status, confidence, warnings |
| `SheetModel` | sheet name, title, merged ranges, blocks, periods/evidence/conflicts, warnings |
| `BlockModel` | block id, kind, boundaries (A1 range + row/column indexes), header rows, columns, merged ranges, cell records, periods/evidence/conflicts, structure confidence, detection reasons |
| `ColumnModel` | index/letter, full header, header cell address, header rows, original/parent/sub header, period, business role, semantic role, role confidence, data type, counts, samples, notes |
| `CellRecord` | address, row, column, letter, original value, formula, data type, number format, merged anchor, merged range |
| `MergedRangeRecord` | range address, bounds, anchor address, anchor value, width, height |

Cell values are kept exactly as read; they are not normalized or coerced. A merged child
cell keeps `merged_anchor`/`merged_range` so a later stage can resolve the parent value.

## 4. Period detection

Five evidence sources are collected, each with its own confidence:

| Source | Confidence | Address recorded |
| :--- | ---: | :--- |
| Column header (parent/sub/period) | 0.95 | header cell or merged range |
| Block title | 0.85 | title cell address when found |
| Row label (first two columns, first 500 data rows) | 0.80 | row-label cell |
| Worksheet name | 0.70 | sheet name |
| Source filename | 0.60 | filename |

Rules:

- A year is a four-digit Jalali/Gregorian year (`13xx`, `14xx`, `19xx`, `20xx`) on digit
  boundaries. `1403` never matches inside `14031`.
- Month names match as words, so `دی` does not match `ردیف`. `فروردین‌ماه` still counts.
- One year plus one or more months produces `month_year` tokens such as `فروردین 1403`.
  Multiple years produce separate year tokens plus any bare month token; the code never
  pairs a month with a year it cannot justify.
- Disagreement between scalar sources (sheet name, title, filename) for the same dimension
  is reported as a conflict and sets `status="conflicting"`. No source is preferred.
- If no source yields a period, `status="not_detected"` and the token list stays empty.
- `SheetIntelligence` now carries `filename_years`, `periods`, `period_evidence`,
  `period_sources`, `period_status` and `period_confidence`; filename years also take part
  in `metadata_conflicts`.

## 5. Semantic roles

`DataRole` now distinguishes `IDENTIFIER`, `TEMPORAL_KEY`, `NUMERIC_MEASURE`,
`CATEGORICAL`, `GROUP`, `TEXT`, `CALCULATED` and `UNKNOWN`.

- T2 (G3/R3) fixed: totals keywords match on word boundaries. `جمع کل` and `total` still
  total; `کلانتری ۱۲`, `تکلیف`, `جمعیت` and `کلاسه پرونده` do not.
- T3 (G5/R4) fixed: a header containing `سال`, `ماه` or `روز` (or English `year`, `month`,
  `day`) becomes `TEMPORAL_KEY` only when its content agrees — integer years, months 1..12,
  or days 1..31. Such a column never enters means, variance or correlation.
- Identifier headers and date headers keep their existing identifier/date roles.
- Low-cardinality text columns whose header says `گروه`, `دسته`, `نوع`, `بخش`, `category`,
  `group`, `type` ... become `GROUP`; other low-cardinality text stays `CATEGORICAL`.

## 6. Files changed

- New: `src/local_office_ai/excel/period_detection.py`
- New: `src/local_office_ai/excel/extraction_model.py`
- New: `tests/test_extraction_model.py`
- `src/local_office_ai/excel/structure_intelligence.py` — period fields, title address,
  filename conflicts, row-label filtering.
- `src/local_office_ai/excel/sheet_analyzer.py` — total boundaries, temporal keys, groups.
- `src/local_office_ai/core/models.py` — `DataRole.GROUP`, `AnalysisResult.extraction`.
- `src/local_office_ai/app/workflow.py` — build and attach the extraction model.
- `src/local_office_ai/excel/generic_report.py`, `report_writer.py` — audit-only
  «مدل استخراج» sheet.
- `tests/test_sheet_structure.py`, `tests/test_structure_intelligence.py` — T2/T3 xfail
  markers removed.
- `AGENTS.md`, `docs/quality_evidence.md`, `docs/NEXT_AGENT_START_HERE.md` — status and
  evidence.

## 7. Tests

`tests/test_extraction_model.py` covers the seven required cases plus three guards:

1. period from column headers (`D1`, `E1`);
2. period from worksheet name (`فروردین 1403`);
3. period from block title (merged `A1:C1`, evidence address `A1`);
4. period from filename (`آمار روزانه 1404.xlsx`);
5. date/identifier columns are not measures;
6. original cell references, raw values and formulas (`=B2*2`) are preserved;
7. multi-row header with merged parent (`C3:D3`), child periods and header addresses;
8. no period is invented without evidence (`not_detected`, empty lists);
9. filename vs sheet-name period conflict is reported, not resolved;
10. a filename year merges into the sheet's month-year token instead of duplicating it.

## 8. Evidence (sandbox, CPython 3.13.15)

```text
python -m pytest tests/test_extraction_model.py -o addopts="" -q
-> 10 passed in 0.33s (exit 0)

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 179 passed, 2 skipped, 36 deselected, 1 xfailed in 18.98s (exit 0)
```

The `xfailed` case is T9 (output suffix), intentionally still open. The 36 deselected tests
are `requires_sample` and are not counted as passing.

End-to-end `run_engine.py --no-ai --yes` on a synthetic two-level-header workbook:

- clean profile: no extraction sheet; unchanged delivery sheets.
- audit profile: «مدل استخراج» lists block `داده#B1`, `A3:D7`, header rows 3/4, periods
  `1403`/`1404` from `column_header`, merged range `C3:D3`.

## 9. Limitations and out of scope

- Sandbox evidence only. The owner's Windows 3.14 environment, Ollama/Qwen, ComfyUI,
  LibreHardwareMonitor, Tkinter/PySide6 runtime and real administrative files were not
  exercised.
- Row-label period evidence is limited to the first 500 data rows; cell records are capped
  at 4000 per block (`cells_truncated`). Boundaries, headers and columns remain complete.
- No statistical reconciliation engine, no UI redesign, no Prompt 01B work.
