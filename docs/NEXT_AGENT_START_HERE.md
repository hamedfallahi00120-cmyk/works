# Next Agent Start Here — Prompt 05A

## Repository and starting point

- Repository: `hamedfallahi00120-cmyk/works` (private).
- Accepted Phase-0 base on `main`: `6df625659c97c0b7da1f78f33fa5fc37864a7580`.
- Prompt 01A was merged to `main` as PR #4; Prompt 01B as PR #5 (`113cc1f`);
  Prompt 02A as PR #6 (`d173465`); Prompt 02B as PR #7 (`891b662`); Prompt 03B as
  PR #9 (`18fb3d4`); Prompt 04A as PR #10 (`4741a98`).
- Prompt 03A is implemented on `feature/prompt-03a-semantic-source-matching`; PR #8
  is open and not merged to `main`.
- Prompt 04B was merged to `main` as PR #11 (`8f77c25`).
- Prompt 05A is implemented on `feature/prompt-05a-quality-gate-provenance`, created
  from `main` at `8f77c25`.
- The three original administrative source files and the private sample workbook are
  intentionally not stored in GitHub.

The Phase-0 read-only audit has already been completed. Its report is
`PHASE0_BASELINE_AUDIT.md`. Do not repeat Phase 0 or restart the repository audit.

## Exact status at handoff

Prompt 01A — Data Model, Workbook Structure, and Metadata Extraction — is implemented
and merged. Prompt 01B, Prompt 02A, Prompt 02B, Prompt 03B, Prompt 04A and Prompt 04B
are merged. Prompt 03A remains on its branch and PR #8.

**Prompt 05A — Quality Gate Architecture and Data Provenance — is implemented on
`feature/prompt-05a-quality-gate-provenance`.** It adds:

- `quality/models.py` — one `PASS` / `WARNING` / `FAIL` model over the thirteen required
  stages, `GateFinding`, `GateStageResult`, `QualityGateReport`, and the machine-readable
  `ProvenanceRecord` / `ProvenanceSource` / `ProvenanceLedger` evidence.
- `quality/gate.py` — `QualityGate.evaluate` and `QualityGate.evaluate_folder` reuse the
  existing extraction, semantics, data-quality, reconciliation, duplicate-control,
  period, matching and DOCX-validation components. There is no second validator.
- `quality/provenance.py` — every important figure keeps filename, sheet, block,
  cell/range, original and normalized labels/values, period, unit, operation, formula,
  source cells with raw values and formulas, related records and validation status.
  Derived values without a source cell/range are a blocking failure.
- `app/workflow.py`, `app/folder_workflow.py`, `core/models.py` — provenance is built
  before the writer, stages 1–11 are evaluated before generation, stages 12–13 after it,
  and `AnalysisResult`, `RunResult` and `FolderResult` serialize the gate.
- `RunResult.status` is `failed` for a blocking gate failure, `completed` only for a clean
  `PASS` with unchanged input, and `completed_with_warnings` for a warning or non-blocking
  failure. A warning is never presented as complete success.
- The desktop result/folder pages, banners and composer hints render a failed run as
  «ناموفق — گزارش معتبر نیست», never as a warning or a completed report.
- `excel/report_writer.py` — the audit profile writes «دروازه کیفیت» and
  «ردیابی منبع (شواهد)»; the clean profile only shows the compact gate summary in
  «اعتبارسنجی داده». `ReportContent` carries the gate and provenance summary for Word.

Recorded sandbox evidence (Linux, CPython 3.13.15, pytest 9.2.0):

```text
python -m pytest tests/test_quality_gate.py -o addopts="" -q
-> 9 passed in 2.06s

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 358 passed, 2 skipped, 36 deselected in 51.24s
```

The 36 deselected tests are `requires_sample`; their private workbook is absent and they
are not counted as passing. The 2 skips are optional GUI/monitoring tests whose system
libraries are unavailable in this sandbox.

## Start procedure

1. Fetch the private repository and switch to
   `feature/prompt-05a-quality-gate-provenance`.
2. Verify the current branch, HEAD SHA and clean working tree. Record the exact output.
3. Read, in this order:
   - `MASTER_PROJECT_HANDOVER_PUBLIC.txt`
   - `PHASE0_BASELINE_AUDIT.md`
   - `AGENTS.md`
   - `docs/quality_evidence.md`
   - this file
4. Use the Phase-0 audit as the accepted baseline. Do not demand that the owner resend
   these files, repeat the audit, or perform another exploratory phase.
5. Prompts 01A, 01B, 02A, 02B, 03A, 03B, 04A, 04B and 05A are implemented. Do not
   repeat them. Wait for the owner to accept Prompt 05A before starting the next stage.
6. Keep Python as the only arithmetic authority, Qwen for narrative only, and every
   report number traceable to validated `ReportContent`.
7. Windows Word rendering and the real administrative files remain owner-side checks on
   `F:\Local-Office-AI`. Do not recreate that environment or fabricate those files.

## Non-negotiable rules for the next stage

- One dedicated branch per stage, focused tests, actual evidence, no direct work on
  `main`, no merge/push without owner instruction.
- A `.docx` output must always be a real Word package; never write Excel or plain text
  under that suffix.
- Applying right alignment alone is not RTL. Keep `w:bidi`, `w:rtl`, `w:cs`, `w:szCs`
  and `w:bidiVisual`; `tests/test_docx_report_writer.py` guards them.
- A template is presentation only. Its historical figures are removed by default and
  never become report data; `--keep-template-figures` is the only exception.
- Reopen and validate every generated DOCX. Unsupported numbers block delivery.
