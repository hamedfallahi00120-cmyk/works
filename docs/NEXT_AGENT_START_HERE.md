# Next Agent Start Here — Prompt 03A

## Repository and starting point

- Repository: `hamedfallahi00120-cmyk/works` (private).
- Accepted Phase-0 base on `main`: `6df625659c97c0b7da1f78f33fa5fc37864a7580`.
- Prompt 01A was merged to `main` as PR #4; `main` was at `f572c47`.
- Prompt 01B was merged to `main` as PR #5; `main` was at `113cc1f`.
- Prompt 02A was merged to `main` as PR #6; `main` was at `d173465`.
- Prompt 02B was merged to `main` as PR #7; `main` is at `891b662`.
- Dedicated Prompt 03A branch: `feature/prompt-03a-semantic-source-matching`.
- The three original administrative source files and the private sample workbook are
  intentionally not stored in GitHub.

The Phase-0 read-only audit has already been completed. Its report is
`PHASE0_BASELINE_AUDIT.md`. Do not repeat Phase 0 or restart the repository audit.

## Exact status at handoff

Prompt 01A — Data Model, Workbook Structure, and Metadata Extraction — is implemented
and merged. See the Prompt 01A section in `docs/quality_evidence.md` and
`docs/prompt_01a_technical_report.md`. Prompt 01B has its own
`docs/prompt_01b_technical_report.md`.

Prompt 02A was merged to `main` as PR #6. Prompt 02B — Statistical Reconciliation and
Calculation Integrity — was merged to `main` as PR #7. **Prompt 03A — Semantic Source
Matching and Period Alignment — is implemented on
`feature/prompt-03a-semantic-source-matching`.** It adds `period_alignment.py` (Jalali
and Gregorian periods, ranges and compatibility) and `semantic_matching.py` (controlled
field matching with label, role, unit, period, identifier and value evidence). Identity
and value comparability are separate decisions; missing or conflicting periods and
ambiguous candidates are reported, never guessed. The audit profile writes
«تطبیق معنایی منابع», and the folder report writes «تطبیق شاخص‌ها بین منابع». The
branch is local and is not merged or pushed to `main`. See `docs/quality_evidence.md`
for the implementation record.

Recorded sandbox evidence (Linux, CPython 3.13.15, pytest 9.2.0):

```text
python -m pytest tests/test_semantic_matching.py -o addopts="" -q
-> 20 passed in 1.90s

python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 275 passed, 6 skipped, 36 deselected, 1 xfailed in 30.81s
```

The `xfailed` result is T9. The 36 deselected tests are `requires_sample`; their private
workbook is absent and they are not counted as passing. The six skips are optional
GUI/monitoring tests whose system libraries are unavailable in this sandbox and are not
counted as passing.

## Start procedure

1. Fetch the private repository and switch to
   `feature/prompt-03a-semantic-source-matching`.
2. Verify the current branch, HEAD SHA and clean working tree. Record the exact output.
3. Read, in this order:
   - `MASTER_PROJECT_HANDOVER_PUBLIC.txt`
   - `PHASE0_BASELINE_AUDIT.md`
   - `AGENTS.md`
   - `docs/quality_evidence.md`
   - this file
4. Use the Phase-0 audit as the accepted baseline. Do not demand that the owner resend
   these files, repeat the audit, or perform another exploratory phase.
5. Prompt 01A, Prompt 01B, Prompt 02A, Prompt 02B and Prompt 03A are complete. Do not
   repeat them or substitute audit Phase 1a for them. The next engineering action starts
   only after the owner accepts Prompt 03A.
6. Work only on a dedicated branch for the next stage. Never commit directly to `main`.
7. Keep later prompts out of scope until the owner authorizes them. Do not merge or push
   to `main`.

## Test and evidence policy

Run the focused test gate and record the actual output:

```bash
python -m pytest tests -o addopts="" -q -ra -m "not requires_sample"
```

Do not rely on historical test counts; collection changes with Python, PySide6 and
optional dependencies. Read the latest GitHub Actions run for the exact CI result.

T2 and T3 are fixed by Prompt 01A and must keep passing normally. T9 is still
`xfail(strict=True)` because the output-suffix defect remains open. T4 is a normal passing
guard; the audit did not reproduce R5.

## CI execution policy

The repository workflow is intentionally PR-only:
- One Windows/Python 3.14.7 test run starts when a PR is opened.
- A new run starts only when that PR's commit changes (`synchronize`) or the PR is reopened.
- Documentation-only changes outside the CI path list do not start the test suite.
- There is **no `push` trigger on `main`** and no test run is subscribed to the PR
  merge/close event. Therefore merging an already tested PR does not repeat the suite.
- One invocation of pytest is used. Tests marked `requires_sample` are excluded; there is
  no second collection run or artifact upload.
- Do not interpret this as “never test changed code”: each new source/test/dependency/
  workflow snapshot in a PR must be validated once before merge.

## Fixed project constraints

- Python is the source of truth for extraction, numerical calculations and validation.
- Qwen is for semantic interpretation and narrative only; it must not invent or calculate
  source figures.
- Preserve the existing architecture, UI, RTL behaviour and working functionality; do not
  rewrite from scratch.
- Source documents are read-only.
- Do not recreate, modify or upgrade the owner's Windows venv or system Python.
- Preserve `F:\work`, `legacy/`, user data and existing input documents.
- Keep temporary and project artifacts within the previously approved locations.
- The user's local Windows target is `F:\Local-Office-AI`, Python 3.14.7. GitHub Actions
  is not the same as that installation.
- Do not claim Ollama/Qwen, LibreHardwareMonitor, Tesseract, desktop UI or real-file
  acceptance was verified unless it was actually tested on the target Windows machine.
- Never add the original administrative documents, private sample workbook, real figures
  or personal paths to GitHub.

## Known unresolved items

- T9 / G11 / R8: output suffix validation is missing (still `xfail(strict=True)`).
- T2 / G3 / R3 and T3 / G5 / R4 are fixed in Prompt 01A; keep their tests passing.
- T4 passes; R5 was found unreachable in the audited code path.
- The three real acceptance files and the private sample workbook are not in the
  repository; real-file validation remains local to the owner's Windows environment.
- Prompt 01A and Prompt 01B evidence is sandbox-only; Windows-target and real-file
  acceptance are still owner-run.
