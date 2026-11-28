# Next Agent Start Here — Prompt 04B

## Repository and starting point

- Repository: `hamedfallahi00120-cmyk/works` (private).
- Accepted Phase-0 base on `main`: `6df625659c97c0b7da1f78f33fa5fc37864a7580`.
- Prompt 01A was merged to `main` as PR #4; Prompt 01B as PR #5 (`113cc1f`);
  Prompt 02A as PR #6 (`d173465`); Prompt 02B as PR #7 (`891b662`); Prompt 03B as
  PR #9 (`18fb3d4`); Prompt 04A as PR #10 (`4741a98`).
- Prompt 03A is implemented on `feature/prompt-03a-semantic-source-matching`; PR #8
  is open and not merged to `main`.
- Prompt 04B is implemented on `feature/prompt-04b-docx-report-writer`, created from
  `main` at `4741a98`; it is not merged to `main` yet.
- The three original administrative source files and the private sample workbook are
  intentionally not stored in GitHub.

The Phase-0 read-only audit has already been completed. Its report is
`PHASE0_BASELINE_AUDIT.md`. Do not repeat Phase 0 or restart the repository audit.

## Exact status at handoff

Prompt 01A — Data Model, Workbook Structure, and Metadata Extraction — is implemented
and merged. Prompt 01B, Prompt 02A, Prompt 02B, Prompt 03B and Prompt 04A are merged.
Prompt 03A remains on its branch and PR #8.

**Prompt 04B — DOCX Report Writer and Persian RTL Rendering — is implemented on
`feature/prompt-04b-docx-report-writer`.** It adds:

- `reporting/docx_rtl.py` — real OOXML RTL: section `w:bidi`/`w:rtlGutter`, paragraph
  `w:bidi`, run `w:rtl`/`w:bCs`, complex-script fonts and `w:szCs`, RTL tables with
  `w:bidiVisual`, repeated shaded header rows and `PAGE`/`NUMPAGES` fields.
- `reporting/docx_report_writer.py` — `DocxReportWriter`. Standalone mode builds a
  complete A4 Persian report: title page, headings, executive summary, RTL tables,
  captions, source/statistical notes, warnings, conclusions and recommendations, with
  header/footer and page numbers. Template mode preserves layout/styles, replaces
  placeholders, removes static historical figures (unless explicitly preserved),
  inserts the analytical body at `{{content}}` or appends it, adds generated content under
  headings the template already contains, reuses a matching template table instead of
  duplicating it, and skips only a table that is already represented.
- `reporting/docx_validation.py` — reopens the delivered file and verifies every XML
  part, required headings, the expected table headers/count, RTL sections/runs/tables, headers,
  footers, real `PAGE` fields, unresolved placeholders and every numeric claim outside
  Word fields against `ReportContent.approved_numbers()`. An invalid file is never
  delivered.
- `reporting/specification.py` — `ReportContent.from_analysis` now emits identity,
  metrics, descriptive statistics, request answers, quality and reconciliation tables,
  plus `placeholder_values()` and `approved_numbers()`. Model narrative is admitted only
  when every number it quotes is approved.
- `reporting/formats.py` / `generator.py` — DOCX writer is `DocxReportWriter` and a
  template is optional. `ReportSpecification.preserve_template_figures` maps to
  `strip_static_figures`.
- `app/workflow.py` / `coordinator.py` / `ui/cli.py` — a Word request without a template
  builds a standalone report; `--keep-template-figures` keeps template figures only when
  explicitly requested.

Recorded sandbox evidence (Linux, CPython 3.13.15, pytest 9.2.0):

```text
python -m pytest tests/test_docx_report_writer.py -o addopts="" -q
-> 25 passed in 9.07s

python -m pytest tests/test_report_output_formats.py -o addopts="" -q
-> 23 passed in 1.34s

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 345 passed, 6 skipped, 36 deselected in 48.55s
```

The 36 deselected tests are `requires_sample`; their private workbook is absent and they
are not counted as passing. The 6 skips are optional GUI/monitoring tests whose system
libraries are unavailable in this sandbox and are not counted as passing.

## Start procedure

1. Fetch the private repository and switch to
   `feature/prompt-04b-docx-report-writer`.
2. Verify the current branch, HEAD SHA and clean working tree. Record the exact output.
3. Read, in this order:
   - `MASTER_PROJECT_HANDOVER_PUBLIC.txt`
   - `PHASE0_BASELINE_AUDIT.md`
   - `AGENTS.md`
   - `docs/quality_evidence.md`
   - this file
4. Use the Phase-0 audit as the accepted baseline. Do not demand that the owner resend
   these files, repeat the audit, or perform another exploratory phase.
5. Prompts 01A, 01B, 02A, 02B, 03A, 03B, 04A and 04B are implemented. Do not repeat
   them. Wait for the owner to accept Prompt 04B before starting the next stage.
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
