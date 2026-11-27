# Next Agent Start Here — Prompt 01A

## Repository and starting point

- Repository: `hamedfallahi00120-cmyk/works` (private).
- Accepted Phase-0 base on `main`: `6df625659c97c0b7da1f78f33fa5fc37864a7580`.
- Prepared baseline branch: `fix/multisource-extraction-integrity`.
- Dedicated working branch for the next stage: `feature/prompt-01a-data-model-metadata`.
- The three original administrative source files and the private sample workbook are intentionally not stored in GitHub.

The Phase-0 read-only audit has already been completed. Its report is
`PHASE0_BASELINE_AUDIT.md`. Do not repeat Phase 0 or restart the repository audit.

## Exact status at handoff

The preparation work on the baseline branch consists of:
- The completed Phase-0 audit and sanitized project handover.
- Isolated Linux sandbox baseline evidence.
- Four targeted regression tests: T2, T3, T4 and T9.
- `docs/quality_evidence.md`.
- CI configuration aligned to Python 3.14.7, with artifact upload treated as best-effort.
- Strict expected-failure classification for the three known unresolved defects.

This is **Phase-1a preparation only**. It is not the agreed Prompt 01A.
The original Prompt 01A — **Data Model, Workbook Structure, and Metadata Extraction** —
has not yet been executed. Prompt 01B has not started.

## Start procedure

1. Fetch the private repository and switch to `feature/prompt-01a-data-model-metadata`.
2. Verify the current branch, HEAD SHA and clean working tree. Record the exact output.
3. Read, in this order:
   - `MASTER_PROJECT_HANDOVER_PUBLIC.txt`
   - `PHASE0_BASELINE_AUDIT.md`
   - `AGENTS.md`
   - `docs/quality_evidence.md`
   - this file
4. Use the Phase-0 audit as the accepted baseline. Do not demand that the owner resend these files, repeat the audit, or perform another exploratory phase.
5. The owner will provide the Shared Execution Contract and the exact original Prompt 01A. Start implementing that exact scope; do not reinterpret it as audit Phase 1a.
6. Work only on the dedicated Prompt 01A branch or a clearly named child branch. Never commit directly to `main`.
7. Keep Prompt 01B and all later prompts out of scope. Do not merge or push to `main`.

## Test and evidence policy

Run the focused test gate and record the actual output:

```bash
python -m pytest tests -o addopts="" -q -ra -m "not requires_sample"
```

Do not rely on historical test counts; collection changes with Python, PySide6 and optional dependencies. Read the latest GitHub Actions run for the exact CI result.

T2, T3 and T9 are temporarily marked `xfail(strict=True)` because their defects are known and intentionally not fixed in this preparation stage. This is not a claim that they pass. If Prompt 01A genuinely fixes any of them, remove that test's xfail marker and require a normal pass. T4 is a normal passing guard; the audit did not reproduce R5.

## CI execution policy

The repository workflow is intentionally PR-only:
- One Windows/Python 3.14.7 test run starts when a PR is opened.
- A new run starts only when that PR's commit changes (`synchronize`) or the PR is reopened.
- Documentation-only changes outside the CI path list do not start the test suite.
- There is **no `push` trigger on `main`** and no test run is subscribed to the PR merge/close event. Therefore merging an already tested PR does not repeat the suite.
- One invocation of pytest is used. Tests marked `requires_sample` are excluded; there is no second collection run or artifact upload.
- Do not interpret this as “never test changed code”: each new source/test/dependency/workflow snapshot in a PR must be validated once before merge.

## Fixed project constraints

- Python is the source of truth for extraction, numerical calculations and validation.
- Qwen is for semantic interpretation and narrative only; it must not invent or calculate source figures.
- Preserve the existing architecture, UI, RTL behaviour and working functionality; do not rewrite from scratch.
- Source documents are read-only.
- Do not recreate, modify or upgrade the owner's Windows venv or system Python.
- Preserve `F:\work`, `legacy/`, user data and existing input documents.
- Keep temporary and project artifacts within the previously approved locations.
- The user's local Windows target is `F:\Local-Office-AI`, Python 3.14.7. GitHub Actions is not the same as that installation.
- Do not claim Ollama/Qwen, LibreHardwareMonitor, Tesseract, desktop UI or real-file acceptance was verified unless it was actually tested on the target Windows machine.
- Never add the original administrative documents, private sample workbook, real figures or personal paths to GitHub.

## Known unresolved items

- T2 / G3 / R3: total-label false positives.
- T3 / G5 / R4: year/month/day columns can be misclassified as numeric measures.
- T9 / G11 / R8: output suffix validation is missing.
- T4 passes; R5 was found unreachable in the audited code path.
- The three real acceptance files and the private sample workbook are not in the repository; real-file validation remains local to the owner's Windows environment.

**The next engineering action is Prompt 01A, and only Prompt 01A.**
