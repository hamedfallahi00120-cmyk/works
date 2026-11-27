# Prompt 02A — Semantic Classification and Data Quality Validation

Branch: `feature/prompt-02a-semantic-quality`, based on `main` at `113cc1f`
(the merge of Prompt 01B). Scope: classify every extracted field into an explicit
semantic role and record data-quality findings before statistical conclusions.
No cross-source aggregation was added.

## Why this stage exists

Prompt 01B produced correct blocks, headers and total rows, but the statistical
role of a column was still a short enum derived mostly from its header. That is
not enough for administrative workbooks: a ten-digit postal code, a personnel
number or a year can look exactly like a measure, and a questionable cell must
be reported without being silently repaired.

## Semantic vocabulary

`excel/semantic_classification.py` defines the required vocabulary:

| Role | Persian label | Meaning |
| :--- | :--- | :--- |
| `identifier` | شناسه | Administrative, postal, personnel, vehicle, region, record or row key |
| `temporal_key` | کلید زمانی | Date, year, month or day key |
| `measure` | متغیر اندازه‌گیری | Numeric quantity that may enter statistics |
| `category` | دسته‌بندی | Low-cardinality descriptive field |
| `group` | گروه‌بندی | Low-cardinality grouping key |
| `subtotal` | جمع میانی | Row role, never a column measure |
| `grand_total` | جمع کل | Row role, never a column measure |
| `descriptive` | فیلد توصیفی | Free text that is not a measure or a key |
| `calculated` | فیلد محاسباتی | Formula-driven measure, kept for traceability |
| `unknown` | نامشخص | Insufficient evidence; never silently promoted |

`SemanticClassifier` projects the existing detectors; it does not re-detect
blocks or headers. Each `SemanticField` records:

- the semantic role and a confidence;
- the original business role (`ColumnRole`) and statistical role (`DataRole`);
- the unit when the header or an adjacent «واحد» column provides one;
- the evidence list behind the decision (header words, numeric/text ratio,
  distinct count, parent/sub header, period, block title, source filename).

Subtotal and grand-total rows are reported on the `SemanticBlock` through the
shared `row_classification` module, so they are never folded into a column
measure.

### Identifier protection

`excel/semantic_vocabulary.py` is a leaf module with the shared word lists and
value heuristics. A column becomes an identifier when:

- its header contains a strong identifier word (`شناسه`, `کد ملی`, `کد پرسنلی`,
  `شماره ملی`, `سریال`, `بارکد`, `ردیف`, `یکتا`, ...);
- a weaker key word (`کد`, `شماره`, `پلاک`, `پرسنلی`, ...) is combined with
  numeric or code-like values;
- the values have leading zeros;
- the values are fixed-width numeric codes of five or more digits and the header
  has no measure word;
- all-digit text samples have code length and the column does not parse as
  numbers.

A measure word in the header (`مبلغ`, `تعداد`, `مجموع`, ...) prevents the
fixed-width and digit-text rules from firing, so a ten-digit monetary amount is
not turned into a code. Temporal evidence is checked before identifier evidence:
`سال`, `ماه` and `روز` columns whose content agrees stay `TEMPORAL_KEY`.

`sheet_analyzer.assign_statistical_roles` now applies the same vocabulary. The
old `ColumnRole.ID` check remains as a fallback, and a formula ratio no longer
overrides an identifier or temporal key.

## Data-quality checks

`excel/data_quality.py` runs on the primary table and returns a `QualityReport`.
Every finding has `PASS`, `WARNING` or `FAIL`, a Persian label, a scope, a sheet,
a cell address, a column, a count, samples and evidence. The eight required
checks always appear, even when they pass.

| Code | What is recorded | Default severity |
| :--- | :--- | :--- |
| `missing_values` | Measure cells that are empty or formula-less blanks | `WARNING`; `FAIL` at 50% or more |
| `invalid_numeric` | Excel errors and mixed digit/letter values in numeric columns | `WARNING`; `FAIL` at 50% or more |
| `unexpected_text` | Pure text in a column expected to be numeric | `WARNING` |
| `negative_values` | Negative values in measures whose header does not allow them | `WARNING` |
| `duplicate_identifiers` | Repeated values in a column that should be unique | `FAIL` for unique identifiers; `WARNING` otherwise |
| `inconsistent_units` | Mixed scale words, mixed currencies, or header/cell mismatch | `WARNING` |
| `conflicting_periods` | Scalar period sources that disagree | `FAIL` |
| `suspicious_values` | IQR outliers, extreme magnitudes, or a constant measure | `WARNING` |

`allows_negative()` exempts balances, differences, variances, profit/loss and
similar headers. `requires_unique_identifier()` marks national IDs, personnel
IDs, serials, barcodes and row IDs as strict; a repeated region or category code
is not reported as a broken unique identifier unless the column is almost
entirely unique.

The analyzer never edits the workbook. A cell with `#DIV/0!`, `12abc` or
`نامشخص` stays exactly as it was; only the finding is recorded. Conflicting
periods reuse `period_detection` (sheet name, filename, block title, column
headers and row labels) and never pick a winner.

## Integration

- `AnalysisResult.semantics` and `AnalysisResult.quality` are new optional
  fields; both are serialized into the processed JSON through `to_dict()`.
- `app/workflow.py` builds semantics and quality immediately after table
  detection and before statistics, so later stages already see the roles and
  findings.
- The audit profile writes two new sheets:
  - «نقش‌های معنایی» — every field's role, confidence, unit, evidence and row
    roles;
  - «کیفیت داده (تفصیلی)» — every finding with its status, location and samples.
- The clean profile never writes those sheets. Instead, the existing
  «اعتبارسنجی داده» sheet gains a compact quality summary and lists only
  non-`PASS` findings.

## Tests

`tests/test_semantic_quality.py` (13 tests) covers the eight required cases with
real `.xlsx` files created by openpyxl:

1. numeric identifiers versus numeric measures, including a fixed-width code
   without an identifier keyword and a ten-digit monetary amount that stays a
   measure;
2. date, year and month fields versus numeric measures;
3. missing values with their exact cells;
4. invalid numeric content, including an Excel error, a mixed value and pure
   text, plus a re-open check that the source cells were not repaired;
5. negative values, with an exemption for a balance column;
6. mixed scale words (`میلیون` versus `میلیارد`);
7. conflicting period labels in the filename and sheet name;
8. duplicate national identifiers.

Additional tests cover the required role vocabulary, subtotal/grand-total row
roles, identifier evidence, the presence of every required check status, and
JSON serialization. `tests/test_output_quality.py` now asserts the two new
audit sheets are present in audit output and absent from clean output.

## Known limits

- Quality checks run on the primary table only. Secondary blocks are classified
  but not quality-audited, to avoid broad cross-source aggregation in this
  stage.
- A row with a label and no numeric value at all is still treated as a group
  heading by the shared row classifier, so a wholly empty measure row is not
  counted as missing. Partial missing values inside a data row are counted.
- Unit detection recognises scale words and currency words; it does not convert
  between units.
- `suspicious_values` uses a float IQR only for anomaly flags. Authoritative
  arithmetic remains `Decimal` in the analysis modules.

## Not validated

The owner's Windows environment, the three original administrative files and the
private reference workbook were not available. Tests use synthetic workbooks.
