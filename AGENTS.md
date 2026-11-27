# Repository guide (Local Office AI)

Python package that reads, profiles and statistically analyzes administrative/organizational
Excel workbooks (Persian/English), computes every figure with `Decimal`, and writes a
formatted Persian report workbook. Loan/obligation reconciliation is an optional domain
module for workbooks that actually contain those columns, not the product's core.
Optional local LLM (Ollama/Qwen) writes narrative only; optional ComfyUI image studio.

## Handoff protocol (mandatory for any AI agent working here)

This repository is a published baseline extracted from `Local-Office-AI_UI_UX_Refined.zip`
on branch `upload/local-office-ai-ui-ux-refined`. The sample workbook
`data/input/Book2_AI_V4_2_Monitored.xlsx` is absent, so 38 tests fail with
"reference workbook is missing"; that is a missing asset, not a regression.

Any agent working on this repository must:

1. Inspect the entire repository first (tree, docs, tests, git history).
2. Identify the real application architecture and entry points (`run_engine.py`,
   `run_ui.py`, `src/local_office_ai/`) before touching code.
3. Review dependencies (`requirements.txt`, `requirements-documents.txt`) and configuration
   (`config/*.json`, `pytest.ini`).
4. Run the available tests before modifying code and record the real output.
5. Produce an initial technical audit of the current state.
6. Identify bugs, architectural weaknesses, incomplete features and security risks.
7. Propose a prioritized implementation plan.
8. Work on its own dedicated branch.
9. Never work directly on `main`.
10. Never remove existing functionality without documented justification and approval.
11. Never invent test results or claim successful execution without evidence.
12. Run the relevant tests after every change.
13. Submit changes through Pull Requests.
14. Include a clear summary, changed files, test results and remaining risks in each PR.

## Layout
- `src/local_office_ai/core/` — settings (JSON override), logging, exceptions, dataclasses.
- `src/local_office_ai/excel/` — read → inspect → detect tables → normalize → profile →
  descriptive stats → frequencies → groups/crosstabs → trend/correlation → outliers →
  quality → optional loan calculations/reconciliation → report writer.
  `statistical_analysis.py` is the general engine; `generic_report.py` holds the generic
  report sheets as a mixin. `validation.py` holds the 9-check quality checklist.
- `src/local_office_ai/ai/` — Ollama client, context builder, prompts, response validator,
  analyzer (model + deterministic fallback), ComfyUI image studio.
- `src/local_office_ai/monitoring/` — psutil sampler + Windows GPU counters (`typeperf`).
- `src/local_office_ai/app/` — workflow stages and coordinator.
- `src/local_office_ai/ui/` — argparse CLI, optional Tkinter wizard, and `event_bridge`
  (workers post `UiEvent`s; only `root.after` touches widgets).
- `src/local_office_ai/ui/desktop/` — primary modern PySide6/RTL interface: pages, RTL
  theme, Persian font discovery, QThread workers, results/preview, an optional
  hardware panel and an optional desktop `floating_monitor.py` widget (frameless,
  always-on-top, tray, persisted geometry, optional Windows autostart). Entry point
  is `run_ui.py` / `Run_Local_Office_AI_Modern.bat`; `--floating-only` starts just
  the widget.
- `tests/` — pytest suite with real in-process HTTP servers for Ollama/ComfyUI.
- `legacy/` — untouched V4.2 script and its global-`py` launcher. Never make these primary.
- `data/input/` — sample workbook (NOT present in this repository; tests needing it fail).
  `reference/v5_sample_report/` — design reference only (also not present).
- `Deploy_Update.bat` / `DEPLOY.md` — code-only deployment manifest. It must copy
  `run_ui.py` and `Run_Local_Office_AI_Modern.bat` and must never copy `venv`,
  `data\input`, `data\output`, `data\processed`, `logs`, `reports` or `reference`.
  `tests/test_deploy_manifest.py` parses both lists.

## Commands
```bash
python -m pip install -r requirements.txt -r requirements-documents.txt -r requirements-dev.txt
python -m pytest tests -o addopts="" -q      # 216 tests collected (202 when PySide6 is unavailable)
python -m pytest tests -o addopts="" -q -m "not requires_sample"   # CI gate: reference workbook is not shipped
python run_engine.py --file <xlsx> --no-ai --yes
python run_engine.py --file <xlsx> --request "مقایسه 1404 با 1403 و بیشترین ماه" --no-ai
python run_engine.py --folder data/input --request "جمع شاخص‌ها را مقایسه کن" --no-ai
python run_engine.py --file <xlsx> --audit --no-ai     # full traceability output
python run_engine.py --check-structure --file <xlsx>
```

## Report profiles
- Default is `report_profile="clean"`: only sheets the request needs (`پاسخ درخواست`,
  `داشبورد`, `خلاصه اجرایی`, monthly/comparison/statistics sheets when the data supports
  them, `شناسنامه داده`, `اعتبارسنجی داده`, `موارد نیازمند بررسی`). No monitoring, raw-data,
  provenance or AI-execution sheets ever reach the delivered workbook.
- `--audit` (or `report_profile="audit"`) adds the technical/traceability sheets for
  developer debugging; `tests/conftest.py` opts into it so regression tests keep coverage.
- `excel/report_writer.py::_plan` is the single source of truth for which sheet is written.
- Resource monitoring is optional: `WorkflowOptions.monitor` stays `None` in normal runs.
- The AI/provenance sheet is audit-only. With a live model the narrative is rendered inside
  `خلاصه اجرایی` (and `پاسخ درخواست`), so a clean workbook still has exactly the requested
  sheets. `test_clean_profile_keeps_model_narrative_out_of_a_technical_sheet` guards it.
- The modern PySide6 UI mirrors the CLI: a `درخواست (فارسی)` field feeds `request=`,
  `AnalysisWorker` runs `Coordinator`, and the `تحلیل پوشه` button runs selected files
  through `FolderWorkflow` in a QThread. The legacy Tk bridge remains as a rollback path.
- Folder reports carry `درخواست کاربر`, `تفسیر درخواست` and a Decimal-built `پاسخ درخواست`
  row in `خلاصه پوشه`; `FolderResult.request` preserves the raw text.

## Request-aware analysis
- `ai/request_analyzer.py` parses Persian requests into scope, target metrics, months, years.
- `excel/data_qa.py::answer_request` produces grounded answers: scope, metadata/label
  conflicts, ranking per metric, period comparison, quality, incomplete periods. Every
  number comes from `PeriodIntegration`; the model only phrases it.
- `excel/period_integration.py` builds month × period matrix from the year band of each
  sheet, excludes derived sheets, and marks incomplete months (day coverage below the
  calendar length) so totals/ranking never mix partial months. A month that is incomplete
  but equally covered in both compared periods is included and noted.
- Context building keeps `period_analysis` and `request_answers` through every trimming
  stage (they are normally what the request needs) and reports `main_period` so the model
  ranks the right year.
- `app/folder_workflow.py` analyzes a folder, validates each Excel package, optionally
  filters to an explicit selected subset, dedupes by content hash and writes a cross-file
  comparison; source files are never modified.

## Year-label reality check
Monthly workbooks often disagree: sheet name says 1404, sheet title says 1405, the year
band says 1404/1405, and a `هفت ماهه` sheet sums the first column of every sheet while
labelling it 1403. The engine reports the conflict (`metadata_conflicts`) and never picks
a winner; comparison uses only sheets that carry both requested years. `آمار روزانه 1404.xlsx`
is such a file: 4 sheets labelled 1403/1404, 4 labelled 1404/1405, heaviest overlap is
مرداد..آبان.
```

## Rules
- Python is the only authority for arithmetic. Never let model output alter figures.
- General statistics must stay type-aware and non-destructive: missing values and IQR
  outliers are counted/labelled, never silently deleted.
- Do not make loan/financial columns a prerequisite for selecting the primary table or
  building the report. Skip loan sheets when no loan records exist.
- Unmeasurable metrics are `None`/"در دسترس نبود", never `0`.
- Never write to the input workbook; reports go to `data/output/` via atomic temp file,
  the machine-readable JSON (with raw cells) goes to `data/processed/`.
- No sheet names, column letters, or sample figures hard-coded in engine logic.
- Table detection must stay layout-agnostic; new heuristics need a test in
  `tests/test_sheet_structure.py`, `tests/test_workflow_generic.py` or
  `tests/test_structure_intelligence.py`.
- Multi-row headers keep parent and child identity (`parent_header`, `sub_header`, `period`).
  A merged parent must extend the detected span to every child column, including the final pair.
- Row roles (total/difference/percentage/note/footer) stay outside descriptive observations.
  Cross-sheet totals exclude incomplete periods and list them in `excluded_incomplete`.
- Empty cells, real zeros, invalid values and formula errors are counted separately. Never
  coerce a blank to zero. Identifier/temporal keys never enter means, variance or correlation.
- AI context is built by Python in chunks under `max_ai_context_chars`/`max_ai_prompt_chars`.
  Before every model call the prompt is also checked against `ollama_num_ctx` with a reserved
  output budget; oversized prompts are trimmed instead of relying on an HTTP 400 retry.
  Model failures must degrade to the rule-based summary without crashing; report provenance
  separately for Python, real model, rule-based fallback and failed chunks.
- Persian text in user-facing strings; code/comments/docstrings in English.
- Tkinter is imported lazily only (headless CI has no `libtk8.6.so`). Worker threads must
  never call Tk; use `UiBridge` (`Queue` + confirmation `Event`) and `root.after`.
- In the PySide6 UI, workers never touch widgets: only signals cross to the main thread.
  Keep the hardware panel collapsed by default and isolated from the report writer.
- The floating widget must reuse `HardwareMonitor` through `HardwareWorker`; never add a
  second sensor implementation. It samples only while visible, persists geometry/opacity
  in `ui_settings.json`, and keeps `None` as «در دسترس نیست». Windows autostart is
  opt-in, writes only `HKCU\...\Run`, and is reported as unsupported elsewhere.
- `Deploy_Update.bat` and `DEPLOY.md` must keep the same code-only manifest: entry points
  in, user data (`venv`, `data\input`, `data\output`, `data\processed`, `logs`, `reports`,
  `reference`) out. `tests/test_deploy_manifest.py` enforces it.
- A missing monthly grid is reported as «دوره ماهانه شناسایی نشد», never as «0 ماه».
- Scale words (هزار/میلیون/میلیارد/همت) multiply the parsed `Decimal`; never drop them.
- `config/synonyms.json` is a matching aid only: never rewrite `institution_raw` or cell text.
- AI number comparison uses `NumberComparisonPolicy` with zero tolerance. Do not reintroduce
  a relative tolerance for financial figures.
- Semantic column roles must consider header and real content; conflicting headers produce
  a reason string, not a silent role.
- Trend/correlation language must stay descriptive: correlation is not causation.
- Generic report sheets are conditional; never emit empty loan audit/reconciliation sheets.
- Delivery criteria are output quality and correctness, not internal telemetry: no test,
  performance, GPU, RAM or "AI execution" sheet in the delivered workbook.
- Only day-number date columns may drive calendar completeness; row numbers must never be
  read as days (`labels_are_days` in `structure_intelligence.py`).
- A few mismatching declared totals (under 10% of audited totals) stay a warning, not a
  failed checklist, because Python figures remain authoritative.

## Desktop UI structure (RTL refinement pass)
- `ui/desktop/theme.py` — `Palette` + `palette_for(theme)`, `build_stylesheet(...)`, `apply_palette`.
  Every colour, radius and control height for both themes lives here; pages never hard-code colours.
- `ui/desktop/components.py` — shared primitives: `Card`, `SectionHeader`, `Badge`, `MetricTile`,
  `CallbackButton`, `IconButton`, `NavigationButton`, `EmptyState`, `BarChart`, plus the layout
  helpers `form_row()` and `horizontal_line()`.
- `ui/desktop/chrome.py` — window chrome: `Sidebar` (sectioned nav from `NAV_SECTIONS`, collapsible
  to a 68 px icon rail), `TopBar`, `HardwarePanel`, `StatusBar`.
- `ui/desktop/source.py` — `SourceListCard` (multi-file/folder picker + file table) and
  `RequestComposer` (Persian request box, suggestion chips, primary action). Home, file and
  folder pages reuse these two instead of re-implementing source selection.
- `ui/desktop/results.py` — results card with tabs (خلاصه/یافته‌ها/جداول/نمودار/هشدارها).
- `ui/desktop/formatting.py` — `ltr()`, `rtl()`, number/percent helpers, `align_table_header()`.
- Rules: never restore `QGraphicsDropShadowEffect`/`QGraphicsOpacityEffect` on page or card
  widgets — on the Windows target they painted solid black rectangles over the content, and page
  cross-fades were removed for the same reason (the symptom was the black, seemingly empty boxes
  in the 1366×768 screenshots). Prefer real layouts over fixed widths/heights so 1366×768…1920×1080
  all stay tidy, and keep sidebar icon+label rows built from a layout, not a rich-text label.
- Qt's own Persian translation (`qtbase_fa`) is installed in `desktop/app.py::_install_translations`;
  without it every `QFileDialog`/`QMessageBox` label is English ("New folder", "OK", "Cancel").
  Keep the translator referenced, otherwise Qt garbage-collects it and the labels revert.
- `RequestComposer.set_output_note()` states where the report is written (file/folder pages), and
  `MainWindow._sync_compact()` hides the home hero subtitle below 820 px so the primary action
  stays reachable at 1366×768.
- Typography defaults live in `ui_settings.py` / `config/ui_settings.json`: body 13 pt, headings
  19 pt, family auto-detected by `fonts.py` (bundled font folders first, then installed Persian
  system fonts; Vazirmatn was detected on the reference run). The chosen family is applied to
  the QApplication font and re-applied to QSS on every theme change.
- Offscreen preview for UI work: `QT_QPA_PLATFORM=offscreen python run_ui.py` is not enough for
  screenshots; grab widgets with `window.grab().save(...)` from a small harness script.

## Environment notes
- Target host is `F:\Local-Office-AI` with `venv\Scripts\python.exe` (Python 3.14).
  Never use global `py`, create a new venv, reinstall Python, or download Ollama models.
- Original V4.2 files live in `legacy/` and must not be modified.
- Windows GPU paths and the Tk window cannot be verified on Linux; keep both "unavailable"
  branches honest and state them in the delivery report.
