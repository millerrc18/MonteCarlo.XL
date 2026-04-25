# @RISK-Parity Roadmap

This roadmap tracks the 12 robustness and @RISK-like initiatives identified after the first Excel workflow stabilization pass. It is intentionally honest about what is shipped, what has foundations in the codebase, and what remains unfinished.

An item is not considered truly closed until it has survived a live Excel stress-test pass. Code-complete slices can move the product forward, but the roadmap should keep them open until workbook-level verification is done.

## Status Legend

| Status | Meaning |
| --- | --- |
| Shipped | Usable in the current add-in build. |
| Foundation | Some engine/UI pieces exist, but the workflow is not complete enough to call done. |
| Planned | Not implemented yet beyond notes or architecture. |

## Recommended Build Order

1. Reliability and trust: model preflight, Excel state safety, input/output manager.
2. Analyst workflow: settings, distribution wizard, correlation workflow.
3. Communication: report builder, scenario/stress analysis, function swap.
4. Advanced analytics: distribution fitting, goal seek, performance modes, optimization/time-series later.

## The 12 Initiatives

| # | Initiative | Current Status | Why It Matters | Definition Of Done |
| --- | --- | --- | --- | --- |
| 1 | Model preflight / validation panel | Foundation | Users need to know whether a workbook is safe to simulate before pressing Run. | A preflight view checks inputs, outputs, formulas, broken references, numeric output cells, invalid distributions, iteration settings, correlation validity, workbook calculation mode, and any unsupported workbook state. It blocks or warns before run with actionable fixes. |
| 2 | Input / output manager | Foundation | Large workbooks need a single command center for assumptions and forecast cells. | A table view lists all inputs and outputs with cell, label, distribution, parameters, mean/representative value, jump-to-cell, edit, duplicate, remove, and highlight controls. |
| 3 | Better simulation settings | Foundation | @RISK-like tools expose simulation defaults clearly instead of hiding behavior in code. | Settings include default iterations, random seed behavior, sampling method, convergence auto-stop, recalc mode, export defaults, default percentiles, and global vs workbook-specific persistence. |
| 4 | Distribution palette / wizard | Foundation | Analysts should not need distribution theory to choose a reasonable assumption. | A guided picker suggests distributions from plain-English prompts such as percentage, count, positive skew, min/mode/max, waiting time, and extreme value. It shows parameters, previews, and examples inline. |
| 5 | Distribution fitting | Foundation | Historical data should be convertible into a simulation input without manual parameter hunting. | User selects an Excel range, fits candidate distributions, ranks by goodness-of-fit statistics, previews overlays, and applies the selected distribution to an input cell. |
| 6 | Correlation workflow polish | Foundation | Correlation is critical for realistic risk models. | The matrix editor is reachable from Setup and ribbon, validates PSD matrices, can auto-fix, imports/exports Excel ranges, warns on risky values, and persists cleanly in workbook config. |
| 7 | Report builder | Foundation | Simulation outputs need to be shared with stakeholders. | A report wizard lets users choose outputs and sections, then creates a formatted Excel report with histogram, CDF, tornado, summary stats, assumptions, target analysis, metadata, and optional raw-data appendix. |
| 8 | Scenario / stress analysis | Foundation | Users need to understand not just the distribution, but why good or bad cases happen. | Results support target probabilities, conditional analysis for best/worst/target-hit cases, stress ranges, and baseline-vs-stressed comparisons. |
| 9 | Goal seek under uncertainty | Foundation | Many decisions ask for the input level needed to reach a probability target. | User selects a decision cell, target output condition, desired probability/statistic, bounds, and simulation settings. The add-in iterates simulations and reports the required decision value. |
| 10 | Excel state safety | Foundation | An Excel add-in must leave workbooks stable after success, failure, and cancel. | All simulation/export paths reliably restore calculation mode, screen updating, events, status bar, selection, active sheet, and input cell values. Failures log phase-specific diagnostics. |
| 11 | Function swap / model sharing | Foundation | Workbooks should be shareable with people who do not have the add-in. | User can replace `MC.*` formulas with static expected values, preserve a restore map, restore formulas later, and export a non-add-in workbook copy. |
| 12 | Performance modes | Foundation | Users need predictable speed/accuracy tradeoffs. | UI exposes preview/full/deep run presets, Monte Carlo vs Latin Hypercube where supported, Excel recalc vs engine modes, iteration/sec, ETA, raw-data memory warnings, and benchmark diagnostics. |

## Current Implementation Notes

- Core Excel add-in workflow is implemented with Excel-DNA, WPF task pane, ribbon, packed 64-bit XLL deployment, and startup diagnostics.
- Native ARM64 Excel support is not shipped in a supported form. The current Excel-DNA host stack still depends on x86/x64 XLL loaders, but the repository now also contains an active Office.js dual-host foundation:
  - `src/MonteCarlo.Shared` for shared formula, parser, config, and bridge contracts
  - `src/MonteCarlo.Engine.Wasm` for browser-hosted `[JSExport]` engine methods
  - `src/MonteCarlo.OfficeAddin` for a React plus Fluent UI task pane, Office custom functions, workbook scanning, simulation execution, benchmark calls, and native Excel chart export
- Summary export now creates a foundation report with workbook/run metadata, histogram, CDF, optional tornado chart, target analysis, scenario analysis, input assumptions, correlation assumptions, and unique worksheets by default. Multi-output models can now export either the selected output or bundle all outputs into one summary worksheet, and the sheet now applies landscape, fit-to-width, repeat-header, and page-break settings so it prints more cleanly to PDF.
- Light/dark/system theme switching is implemented and persisted.
- Settings now persist export worksheet behavior, default iterations, random vs fixed seed defaults, prompt-at-run seed mode, sampling method, convergence auto-stop, Excel calculation behavior, screen/event suspension, whether Model Check warnings pause a run, and default percentile lists used by summary exports.
- Settings can now be edited in two scopes: Windows-user defaults or workbook-specific overrides stored in the active workbook. Runs, Goal Seek, exports, and Model Check warning pauses resolve against the effective merged settings, and the Run view plus summary export metadata now show which settings were applied.
- The task pane supports 15 distributions; worksheet UDFs exist for all except Discrete.
- The add-input flow includes a guided distribution helper with category, complexity, and search filters plus plain-English starting points for all supported distributions.
- The add-input helper can open a dedicated Excel range picker, fit candidate distributions from the selected numeric data, and apply the selected fit to the parameter editor.
- Rank correlation engine and matrix editor exist. The editor is reachable from Setup and ribbon, imports numeric n by n ranges, exports and re-imports labeled templates, warns about high/fragile matrices, shows exact conflicting pairs plus nearest-PSD adjustment guidance when the matrix is invalid, clearly shows the independent-input state, persists workbook correlation config, and passes correlations into simulation runs.
- Results now include a Scenario Analysis card for worst-tail, best-tail, at-or-below-target, and above-target filters. It shows conditional input summaries so users can see which assumptions changed most in the selected cases.
- Summary export now includes scenario-analysis sections comparing worst, best, and target-hit input means against all runs when a target threshold is entered before export.
- The Results ribbon now includes a Stress Analysis workflow for configured inputs: choose one or more fixed-value, additive-shift, or range-scale rules, rerun the model with the same comparison seed for baseline and stressed runs, add a comparison sheet with stressed assumptions, output impact ranking, and primary-output histogram/CDF comparison charts, and optionally export baseline and stressed raw-data sheets for the primary output.
- Goal Seek now has a ribbon workflow for monotonic decision cells: select a deterministic decision cell, choose an output, metric, bounds, run budget, and tolerance, then MonteCarlo.XL reruns simulations, restores the decision cell, and writes a solver-history report sheet with confidence guidance. The ribbon dialog now supports mean, percentile, P(output > target), and P(output <= target) targets.
- Excel state capture/restore is centralized for simulation runs, goal seek trials, summary/raw exports, workbook writes, highlight refresh, hidden-sheet cleanup, and cell-selection status messages. Restore failures are logged with phase-specific diagnostics.
- The Support ribbon includes a `Recover Excel` command that restores automatic calculation, events, screen updating, alerts, and the status bar after an interrupted run or external automation failure, then logs before/after Excel state snapshots.
- The ribbon now includes workbook-sharing commands to save a separate shareable value-only copy, replace `MC.*` formulas with current values in the active workbook, and later restore them from a workbook custom-XML map.
- Setup, Settings, and the ribbon now expose named run presets for Preview, Standard, Full, and Deep runs. The Run view labels the current run scale, shows live iteration/sec throughput next to elapsed and remaining time, raw-data export warns before writing large datasets or blocks exports that exceed Excel's row limit, and the Support ribbon includes a benchmark command that writes workbook recalc, synthetic engine-throughput, and synthetic export-cost diagnostics.
- Model Check now validates the setup profile before run, blocks missing/invalid inputs and outputs, detects duplicate/conflicting cells, validates distribution parameters, checks correlation matrix shape/validity, warns about very small/large runs, adds Excel-aware checks for active workbook context, calculation/events/screen updating state, workbook protection, missing sheets/cells, protected input cells, formula errors, non-numeric outputs, and static output cells, and can pause runs on warning-only reports when enabled in Settings.
- The Setup view now includes a Model Manager section with searchable, sortable input/output tables, inline label edits, jump-to-cell, highlight, duplicate, remove, refresh-highlights, and bulk-clear actions.
- Stop/cancel is available through task pane, keyboard shortcut, and ribbon callback.

## Platform Track: Native ARM / Office.js Dual Host

Status: Foundation

Why it matters:

- Native Windows ARM Excel cannot load the existing x64 `.xll`.
- The product needs a credible path for Surface Pro and other ARM work deployments.

Shipped foundation:

- Shared formula catalog, parser, and deterministic normal-mode evaluator now back both host strategies.
- Shared bridge contracts let the browser host request validation, sampling, result analysis, goal-seek metric evaluation, and benchmarks from the same engine logic.
- The Office host can be built locally, publishes the `.NET` WebAssembly bridge, bundles a sideloadable manifest, and produces a task-pane UI with charts plus summary-sheet export.

Definition of done:

- A native ARM Excel user can sideload/install MonteCarlo.XL, see the ribbon/task pane, add inputs and outputs, run simulations, export reports, persist workbook config, reopen the workbook, and trust the workflow without needing the x64 `.xll`.

Open work:

- Manually validate the Office host on real ARM hardware with native desktop Excel.
- Expand the current command surface from foundation buttons to richer parity with the Excel-DNA ribbon.
- Carry over the remaining workflow features: workbook-scoped settings editing, richer goal seek, shareable copy, formula replace/restore, stress analysis, and deeper report options.
- Harden workbook persistence compatibility across Custom XML and hidden-sheet fallback.
- Move from localhost sideload development hosting to a deliberate deployment story for work machines.

## Unfinished Work Register

### 1. Model Preflight / Validation Panel

Open work:

- Expand Excel-interoperability checks for unsupported workbook states beyond the configured input/output cells.
- Add manual verification scenarios for success, failure, and warning-only runs in Excel.
- Add per-workbook warning-pause overrides if global Settings is too coarse for teams.
- Expand formula dependency checks beyond configured input/output cells if performance remains acceptable on large workbooks.

### 2. Input / Output Manager

Open work:

- Expand inline editing beyond labels and into parameter cells; complex distribution parameter edits still reopen the existing editor.
- Add output formula/result previews once Excel interop can safely evaluate them.
- Add manual verification for jump-to-cell, highlight refresh, inline label edits, and table behavior across multiple sheets and workbooks.

### 3. Better Simulation Settings

Open work:

- Add manual Excel verification covering prompt seed choices plus workbook-vs-global execution settings.

### 4. Distribution Palette / Wizard

Open work:

- Expand the current category/search helper into a deeper multi-question wizard with branching prompts.
- Add richer category metadata and filtering as more model archetypes emerge from manual testing.
- Link each suggestion to the full `docs/DISTRIBUTION_GUIDE.md` explanation.
- Add parameter guardrails that explain invalid parameter combinations before preview creation fails.

### 5. Distribution Fitting

Open work:

- Expand candidate fitting for GEV and other advanced distributions.
- Add goodness-of-fit details beyond the current KS-style score.
- Add histogram/fit overlay previews and confidence warnings before applying a fit.

### 6. Correlation Workflow Polish

Open work:

- Add manual Excel verification for import/export, persistence after reopening, and a correlated sample smoke test.

### 7. Report Builder

Open work:

- Live-stress-test the new report builder dialog in Excel across selected-output and all-output exports before calling this item done.
- Consider a richer task-pane report designer later; the current implementation now has a dedicated export dialog with section toggles and multi-output scope selection, while the exporter honors those section choices for metadata, stats, percentiles, target analysis, sensitivity, scenario analysis, charts, and assumptions.

### 8. Scenario / Stress Analysis

Open work:

- Add fuller conditional summaries such as target-linked stress thresholds and scenario exports.
- Extend the stress workflow beyond configured/tagged inputs and into richer task-pane comparisons instead of the current ribbon-to-report flow.
- Manually verify worst/best/target-hit filtering in Excel against a known workbook.

### 9. Goal Seek Under Uncertainty

Open work:

- Expand the current probability-above-target ribbon dialog into a richer task-pane workflow.
- Add progress/cancel UI for repeated simulations.
- Add manual Excel verification with an example workbook and known monotonic decision target.

### 10. Excel State Safety

Open work:

- Expand state scopes to any future COM automation paths as they are added.
- Add an automated Excel smoke harness for success, failure, and cancel paths if CI can run desktop Excel.
- Manually test simulation success, simulation failure, cancel, summary export, raw export, highlight refresh, and cell selection in a live workbook.

### 11. Function Swap / Model Sharing

Open work:

- Add a catalog preview before replacement so users can review affected sheets and cells.
- Add conflict handling when cells have changed after replacement but before restore.
- Manually verify replace/restore across multiple worksheets, protected sheets, and workbooks without the add-in installed.

### 12. Performance Modes

Open work:

- Expose explicit Excel recalc-vs-engine execution modes when both are independently selectable.

## First Three Recommended Tickets

1. Build the preflight validation model and surface it before Run.
2. Finish correlation workflow polish: import/export, warnings, and report inclusion.
3. Add the input/output manager table with jump-to-cell and edit/remove controls.
