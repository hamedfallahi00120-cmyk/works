# Quality evidence ledger — Local Office AI

Running evidence for the ten-stage remediation. Every row is a real command result,
not a claim. `[SBX]` = executed in the isolated Linux sandbox; `[WIN]` = must be run on
the Windows target `F:\Local-Office-AI` and is not yet executed.

| Stage | Branch | Scope | Evidence |
| :--- | :--- | :--- | :--- |
| 00 | `main` | read-only baseline audit | `PHASE0_BASELINE_AUDIT.md`; commit `6df6256` |
| 01A | `fix/multisource-extraction-integrity` | sandbox env + real baseline + defect tests | this file, §01A |

## Environment contract

- Sandbox venv: `/tmp/loa-sandbox-venv` (Python 3.13.15), isolated; **never** the Windows
  `venv`, never system Python. Installed from `requirements.txt` + `requirements-documents.txt`
  + `requirements-dev.txt`.
- Windows target (`F:\Local-Office-AI`, Python 3.14.7, Ollama `qwen3.5:9b`,
  LibreHardwareMonitor) is **not reachable here**. Nothing below is a Windows runtime result.
- Real-file acceptance (`5ماهه تصادفات.xlsx`, `شهريار 5 ماهه 1405.xlsx`, `دیییییی.docx`)
  is reserved for the Windows target and is not executed.

## Stage 01A — sandbox baseline and red defect tests

### Baseline (unchanged code, before adding tests) `[SBX]`

```text
python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 166 passed, 2 skipped, 36 deselected in 20.76s   (exit 0)
```

`--collect-only` -> **202 tests collected**. This contradicts the docs: `README.md`/`AGENTS.md`
say 216, `docs/MODERN_UI.md` says ۲۰۷. Real collected count is 202. (G12/C1.)

### New tests (added in 01A)

| ID | Test | Guards | Status on 01A code | Evidence |
| :--- | :--- | :--- | :--- | :--- |
| T2 | `test_total_label_needs_a_word_boundary` | G3 / R3 | **FAIL (red)** — `is_total_label("کلانتری ۱۲") is True` | `test_sheet_structure.py:169` |
| T3 | `test_year_month_day_columns_are_not_measures` | G5 / R4 | **FAIL (red)** — `سال` role is `NUMERIC_MEASURE`, expected `TEMPORAL_KEY` | `test_structure_intelligence.py:267` |
| T4 | `test_code_columns_stay_identifiers` | G5 / R5 | **PASS (green guard)** — R5 not reproducible, see note | `test_structure_intelligence.py` |
| T9 | `test_output_suffix_is_validated` | G11 / R8 | **FAIL (red)** — xlsx written under `report.docx`, exit 0 | `test_cli.py:153` |

### Full gate after adding tests `[SBX]`

```text
python -m pytest tests -o addopts="" -q -m "not requires_sample"
-> 3 failed, 167 passed, 2 skipped, 36 deselected in 18.38s   (exit 1, expected)
```

The 3 failures are exactly T2, T3, T9 and are the intended red baseline; T4 passes.
01B must turn T2/T3/T9 green without breaking the other 167.

### Corrections to the Phase-0 audit (verified against code at `6df6256`)

1. **R5 is unreachable.** `sheet_analyzer.py` skips duplicate-role demotion whenever
   `parent_header`/`sub_header`/`period` is set; `parent_header` is always populated for a
   named column (`profile_columns_from_band`). Two code columns therefore both stay
   `IDENTIFIER`. T4 is kept as a regression guard, not a red test.
2. **ZIP drift is 7 files, not 6.** Byte-diff of `Local-Office-AI_UI_UX_Refined.zip`
   (prefix stripped): 123 identical, 7 differ — `.gitignore`, `AGENTS.md`, `README.md`,
   `pytest.ini`, `requirements.txt`, `tests/conftest.py`, `tests/test_packaging.py`.
   `src/` is identical.
3. **Collected test count is 202**, not 216/207.

### Pending (Windows target, `[WIN]`)

- Re-run `-m "not requires_sample"` in `F:\Local-Office-AI\venv` and compare.
- Acceptance on the three real files (`--audit`, `--check-structure`), input `sha256`
  unchanged.
