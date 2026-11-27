# Quality evidence ledger — Local Office AI

This ledger records the verified handoff state for the next engineering stage.
**The work on this branch is Phase-1a preparation only; it is NOT completion of the agreed Prompt 01A.**
The original Prompt 01A (Data Model, Workbook Structure, and Metadata Extraction) has not started.

## Verified repository state

| Item | Value |
| :--- | :--- |
| Repository | `hamedfallahi00120-cmyk/works` (private) |
| Stable base | `main` at `6df625659c97c0b7da1f78f33fa5fc37864a7580` |
| Preparation branch | `fix/multisource-extraction-integrity` |
| Preparation commits | `5b20cce` (tests + ledger), `06dae62` (sanitized handover + audit), followed by CI/documentation-preflight commits |
| Phase 0 | Completed; see `PHASE0_BASELINE_AUDIT.md` |
| Original Prompt 01A | **Not started** |
| Prompt 01B | **Not started** |
| Production source fixes in this preparation | None |

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
| T2 | `test_total_label_needs_a_word_boundary` | G3 / R3 | Known defect: substring matching classifies labels such as `کلانتری ۱۲` as totals | `xfail(strict=True)` |
| T3 | `test_year_month_day_columns_are_not_measures` | G5 / R4 | Known defect: `سال/ماه/روز` can be treated as numeric measures | `xfail(strict=True)` |
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

