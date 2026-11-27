# Prompt 01B — Technical Report

**Scope:** Multi-Block Extraction, Table Boundaries, and Total Detection.
**Branch:** `feature/prompt-01b-multiblock-extraction`. **Base:** `main` after Prompt 01A
merge (`f572c47`). `main` is unchanged.

## 1. Problem

The engine could detect more than one table, but downstream analysis behaved as if only
the single best table mattered. A worksheet with stacked tables, repeated headers or
mid-table subtotals could lose blocks, swallow the next table's title, or treat a subtotal
as the end of the data. Total detection also relied on permissive text matching, so a
label containing `کل` inside an ordinary word could be misclassified.

## 2. Approach

The existing `TableDetector`, `SheetAnalyzer`, `StructureIntelligence` and
`ExtractionModelBuilder` remain the detection pipeline. Prompt 01B changes their
classification and boundary behavior and adds a shared row-role layer:

```
TableDetector
  -> block boundaries, subtotal rows, grand-total rows
  -> SheetIntelligence.row_roles / summary_rows
  -> SheetIntelligence.block_analyses / block_relationships
  -> ExtractionModelBuilder BlockModel
  -> audit-only «بلوک‌های استخراج» sheet
```

Structural classification is separate from numerical aggregation. A block can be
discovered, bounded and traced without being summed into another block.

## 3. Total and row classification

`excel/row_classification.py` is the shared classifier. It distinguishes:

| Role | Meaning |
| :--- | :--- |
| `data` | ordinary data row with a numeric value |
| `data_missing` | data row whose measures are blank |
| `category` | text-only row with several filled cells |
| `group_heading` | text-only row with one filled cell |
| `subtotal` | totals row with data rows below it |
| `total` | terminal grand-total row |
| `difference` | comparison/difference row |
| `percentage` | percentage/ratio row |
| `note` | explanatory row |
| `footer` | note after the last summary row |
| `title` / `header` / `subheader` / `blank` | structural rows |

`sheet_analyzer.py` exposes explicit `SUBTOTAL_KEYWORDS`, `GRAND_TOTAL_KEYWORDS`,
`DIFFERENCE_KEYWORDS`, `PERCENT_KEYWORDS` and `NOTE_KEYWORDS`. Totals matching uses word
boundaries, so `کلانتری`, `تکلیف`, `جمعیت`, `کلاسه پرونده` and `کلاسی` are ordinary
labels. The `جمع کل` phrase does not match inside `جمع کلاسی`; the standalone `جمع` still
marks that label as a generic total.

## 4. Table boundaries

- A totals row with real data below it is a **subtotal** and stays inside the table.
- A terminal totals row is the **grand total**.
- `block_end_row` keeps the grand total, comparison rows, percentage rows and keyword
  notes attached to the block.
- Detection resumes after the block end, so a following table is not swallowed.
- Formula-only cells count as numeric evidence.
- Repeated headers and lone section titles still split blocks.
- A subtotal never triggers an automatic sum; it is a structural row, not data.

## 5. Multi-block analysis and traceability

`SheetIntelligence` now carries:

- `block_analyses` — one entry per analyzable block with address, full block address,
  kind, title, header row, first/last data row, block end, row/column counts, measures,
  summary rows, row-role counts, confidence, primary flags and warnings.
- `block_relationships` — links blocks only when their headers or periods provide
  evidence. No cross-block sum is performed.

Each `BlockModel` in `AnalysisResult.extraction` carries row roles, summary rows, block
end, subtotal/grand-total rows and merged ranges through the block end. The processed JSON
sidecar includes the same structure.

The audit profile writes the «بلوک‌های استخراج» sheet. The clean profile never writes it.

## 6. Aggregation guard

Subtotal and other structural rows are excluded from:

- normalized data rows and declared totals (`data_extractor.py`);
- loan/obligation records;
- descriptive statistics and group comparisons (`statistical_analysis.py`);
- coverage, column totals, period comparison and declared-total audits
  (`structure_intelligence.py`);
- period integration (`period_integration.py`).

`data_missing` rows remain in blank-cell counts; only structural rows are skipped.

## 7. Tests

`tests/test_multiblock_extraction.py` contains 16 tests covering the eight required
Prompt 01B cases plus multiple subtotals, notes before/after a grand total, a section title
after a grand total, row-role distinctions including a category row with a numeric
code, subtotal exclusion and workflow traceability for every block.

`tests/test_output_quality.py` asserts the audit-only block sheet is present in audit
output and absent from clean output.

Verified sandbox results:

```text
python -m pytest tests/test_multiblock_extraction.py -o addopts="" -q
-> 16 passed in 0.46s

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 196 passed, 6 skipped, 36 deselected, 1 xfailed in 20.75s
```

The single `xfailed` outcome is T9, the open output-suffix defect.

## 8. Limits

- The evidence uses synthetic `.xlsx` files; the owner's real administrative files and
  Windows target were not available in the sandbox.
- A table that has no numeric column (or too few numeric cells to meet the detector's
  numeric-ratio threshold) is not added to the statistical table list. Such text-only
  blocks remain outside the numeric analysis path.
- The single `xfailed` T9 output-suffix defect is unchanged.
