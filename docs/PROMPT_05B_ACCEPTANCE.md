# Prompt 05B — End-to-End Acceptance and False-Success Prevention

## Honest status first

The two named workbooks (`5ماهه تصادفات.xlsx`, `شهريار 5 ماهه 1405.xlsx`) and the old
Word reference (`دیییییی.docx`) are **not present in this repository or sandbox**.
Therefore:

- No end-to-end acceptance against the owner's real files is claimed.
- The extracted values, periods, labels and DOCX content for those files are **not
  verified**.
- No historical checkpoint value is hard-coded in the engine. The sandbox used synthetic
  workbooks to exercise the pipeline and the checkpoint mechanism; those fixtures are
  not evidence about the real files.

The owner must run the command below on the Windows target. Until that run is recorded,
stages marked “unverified” stay unverified.

## Exact acceptance command

```bat
cd /d F:\Local-Office-AI
venv\Scripts\python.exe run_engine.py --acceptance ^
  --acceptance-input "F:\work\5ماهه تصادفات.xlsx" ^
  --acceptance-input "F:\work\شهريار 5 ماهه 1405.xlsx" ^
  --acceptance-input "F:\work\دیییییی.docx" ^
  --checkpoints "F:\work\checkpoints.json" ^
  --output "F:\work\گزارش پذیرش.docx" ^
  --acceptance-dir "F:\work\acceptance-audit" ^
  --no-ai --quiet
```

`checkpoints.json` is optional and only verifies historical control figures:

```json
{"checkpoints": [
  {"label": "تخلفات منجر به تصادف", "previous": "30947", "current": "18870", "reported_percent": "-39.02"}
]}
```

Exit codes: `0` accepted, `3` partial, `4` failed, `5` unverified (no Excel input
available). The Word reference is read and inspected, but it is never a data source and
its historical numbers are stripped from generated content.

## What the acceptance runner verifies

1. Every declared input exists and its type is identified.
2. Every Excel source opens and is analyzed; a skipped or unreadable file prevents
   acceptance.
3. Every worksheet and detected block is reported, including cell caps.
4. Periods are reported per file from block evidence when the monthly grid is not
   applicable.
5. Semantic matching decisions and row-oriented period comparisons are counted; row
   comparisons stay inside one source file, and equivalent row indicators across files
   are listed as matches without pretending they are numerically comparable.
6. Duplicate/overlap groups, blocked aggregation decisions and exclusions are reported.
7. Statistical reconciliation results and failed/warning checks are listed.
8. Provenance coverage is measured on required records, not optional metadata alone.
9. The generated file is the requested format, reopened and validated (ZIP, XML,
   python-docx, RTL, PAGE field, headers, unresolved placeholders, grounded numbers).
10. Historical checkpoints are compared only with independently extracted values; a
    mismatch is reported, never repaired, and never inserted into the report. Only an
    exact normalized label (or an explicitly supplied alias) can confirm a checkpoint;
    a near label or a label repeated across files without a unique cross-file period
    pair is `unverified`. When a checkpoint label exists in two different files, the
    runner may compare the earliest period in one file with the latest period in the
    other **for verification only**; that comparison is never inserted into the report,
    and the audit records that equality of scope/definition between the two files still
    needs the owner's confirmation.

`accepted` requires: all sources processed, all blocks extracted, periods identified,
all gate stages PASS, correct format, valid DOCX, all numbers grounded, provenance
complete and all checkpoints matched (when supplied). Any failure yields `partial` or
`failed`; missing Excel inputs yield `unverified`.

## Sandbox evidence (synthetic, not real-file acceptance)

Environment: Linux, CPython 3.13.15, pytest 9.2.0, openpyxl, python-docx.

```text
python -m pytest tests/test_end_to_end_acceptance.py -o addopts="" -q
-> 24 passed in 6.02s

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 378 passed, 6 skipped, 36 deselected in 54.64s
```

Synthetic workbook (`شاخص | 1404 | 1405`) produced:

```text
periods: 1404, 1405
rows compared: 5 (4 indicators + 1 declared total)
provenance: 40/40 important records with sources, 0 incomplete
gate: 13/13 stages PASS
DOCX: 6 tables, reopened, RTL and PAGE validated
acceptance: accepted (with two matching checkpoints)
```

Changing a checkpoint's current value to `999` produced `failed`; the report still
contained the extracted `900`, demonstrating that the checkpoint cannot overwrite data.

A month-labelled sheet whose title said `1405` while its columns said `1404/1405`
produced `failed` with `period_not_identified` and `critical_validation_failed`
instead of a false success.

With the three real input paths absent, the command produced:

```text
وضعیت پذیرش: تأیید نشده — فایل‌های منبع موجود نیستند
exit=5
no report file written
```

## Fifteen-section audit produced by the runner

The generated `acceptance_report.md` contains: input files detected; files processed;
worksheets inspected; blocks detected; blocks extracted; periods identified; indicators
matched; ambiguous matches; duplicates/overlaps; statistical reconciliation; failed and
warning checks; provenance coverage; generated outputs; DOCX validation; remaining
limitations. It is accompanied by `acceptance_report.json` with the same evidence and
machine-readable records.

## False-success protections added in this stage

- A folder/multi-source Word report is only claimed after stages 12–13 run on the actual
  generated `.docx`; a PDF request is rejected.
- `PERIOD_IDENTIFICATION` fails when source evidence conflicts; no source is preferred.
- A row-oriented period table is excluded from the monthly metric grid, so no fabricated
  metric named `1404`/`1405` and no false “incomplete month” warning appears.
- Declared total rows are compared as recorded totals, never summed with components.
- Accident counts, fatal accidents, fatalities, injury accidents and injured people stay
  separate; missing casualty statistics are named, never inferred.
- A checkpoint mismatch is a failed check; the extracted value is kept and the mismatch
  is written to the audit.
- A missing input is `unverified`, not “completed with warnings”.

## Unverified on this sandbox

- The real workbooks and Word reference; their labels, periods and metric definitions.
- Whether the historical checkpoints actually correspond to the real files.
- Windows Word rendering and fonts; only OOXML structure is asserted.
- Ollama/Qwen narrative (tests run with `--no-ai`).
- The owner's `F:\Local-Office-AI\venv` and Python 3.14 environment.
