# @RISK-Parity Roadmap

This roadmap tracks the 12 robustness and @RISK-like initiatives identified after the first Excel workflow stabilization pass. It is intentionally honest about what is shipped, what has foundations in the codebase, and what remains unfinished.

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
| 4 | Distribution palette / wizard | Planned | Analysts should not need distribution theory to choose a reasonable assumption. | A guided picker suggests distributions from plain-English prompts such as percentage, count, positive skew, min/mode/max, waiting time, and extreme value. It shows parameters, previews, and examples inline. |
| 5 | Distribution fitting | Planned | Historical data should be convertible into a simulation input without manual parameter hunting. | User selects an Excel range, fits candidate distributions, ranks by goodness-of-fit statistics, previews overlays, and applies the selected distribution to an input cell. |
| 6 | Correlation workflow polish | Foundation | Correlation is critical for realistic risk models. | The matrix editor is reachable from Setup and ribbon, validates PSD matrices, can auto-fix, imports/exports Excel ranges, warns on risky values, and persists cleanly in workbook config. |
| 7 | Report builder | Foundation | Simulation outputs need to be shared with stakeholders. | A report wizard lets users choose outputs and sections, then creates a formatted Excel report with histogram, CDF, tornado, summary stats, assumptions, target analysis, metadata, and optional raw-data appendix. |
| 8 | Scenario / stress analysis | Foundation | Users need to understand not just the distribution, but why good or bad cases happen. | Results support target probabilities, conditional analysis for best/worst/target-hit cases, stress ranges, and baseline-vs-stressed comparisons. |
| 9 | Goal seek under uncertainty | Planned | Many decisions ask for the input level needed to reach a probability target. | User selects a decision cell, target output condition, desired probability/statistic, bounds, and simulation settings. The add-in iterates simulations and reports the required decision value. |
| 10 | Excel state safety | Foundation | An Excel add-in must leave workbooks stable after success, failure, and cancel. | All simulation/export paths reliably restore calculation mode, screen updating, events, status bar, selection, active sheet, and input cell values. Failures log phase-specific diagnostics. |
| 11 | Function swap / model sharing | Planned | Workbooks should be shareable with people who do not have the add-in. | User can replace `MC.*` formulas with static expected values, preserve a restore map, restore formulas later, and export a non-add-in workbook copy. |
| 12 | Performance modes | Foundation | Users need predictable speed/accuracy tradeoffs. | UI exposes preview/full/deep run presets, Monte Carlo vs Latin Hypercube where supported, Excel recalc vs engine modes, iteration/sec, ETA, raw-data memory warnings, and benchmark diagnostics. |

## Current Implementation Notes

- Core Excel add-in workflow is implemented with Excel-DNA, WPF task pane, ribbon, packed 64-bit XLL deployment, and startup diagnostics.
- Summary export now includes charts and can create unique worksheets by default.
- Light/dark/system theme switching is implemented and persisted.
- The task pane supports 15 distributions; worksheet UDFs exist for all except Discrete.
- Rank correlation engine and matrix editor exist. This roadmap pass wires the existing correlation editor into Setup navigation and adds a ribbon entry. Import/export and stronger warnings remain unfinished.
- Model Check now validates the setup profile before run, blocks missing/invalid inputs and outputs, detects duplicate/conflicting cells, validates distribution parameters, checks correlation matrix shape/validity, and warns about very small/large runs.
- Stop/cancel is available through task pane, keyboard shortcut, and ribbon callback.

## Unfinished Work Register

### 1. Model Preflight / Validation Panel

Open work:

- Add Excel-interoperability checks for broken formulas, non-numeric output cells, worksheet/workbook protection, workbook calculation state, and unsupported workbook state.
- Add active workbook/sheet context to the preflight report.
- Add manual verification scenarios for success, failure, and warning-only runs in Excel.
- Decide whether warnings should optionally pause a run or only appear in the Model Check view.

### 2. Input / Output Manager

Open work:

- Replace scattered cards with an optional sortable table view.
- Add edit-in-place and jump-to-cell controls.
- Show distribution summaries and mean/P5/P95 previews.
- Add bulk remove and refresh-highlight actions.

### 3. Better Simulation Settings

Open work:

- Persist default iteration count globally.
- Persist seed mode: random, fixed, or prompt.
- Expose convergence auto-stop and recalc mode.
- Add default percentile configuration.
- Decide which settings are global versus workbook-specific.

### 4. Distribution Palette / Wizard

Open work:

- Add distribution category metadata.
- Add guided questions and suggestions.
- Add inline use-case help from `docs/DISTRIBUTION_GUIDE.md`.
- Add parameter guardrails and examples.

### 5. Distribution Fitting

Open work:

- Add range selection for historical data.
- Implement fitting services for candidate distributions.
- Add goodness-of-fit metrics and ranking.
- Add overlay preview and apply-to-input flow.

### 6. Correlation Workflow Polish

Open work:

- Import/export a matrix from/to an Excel range.
- Add better warnings for near-singular matrices and very high absolute correlations.
- Add a clear "no correlations configured" state.
- Include correlation assumptions in exported reports.

### 7. Report Builder

Open work:

- Define report section options.
- Support multiple outputs in one report.
- Add workbook/report metadata.
- Offer PDF-friendly formatting after Excel report creation.

### 8. Scenario / Stress Analysis

Open work:

- Add best/worst/target-hit filtering.
- Compute conditional input summaries.
- Add stressed-run configuration.
- Compare baseline and stressed output distributions.

### 9. Goal Seek Under Uncertainty

Open work:

- Define decision-cell editing workflow.
- Add search algorithms and stop conditions.
- Add run budget controls.
- Present convergence and confidence of the found decision value.

### 10. Excel State Safety

Open work:

- Centralize Excel state capture/restore.
- Use it around simulation, export, and cell-selection flows.
- Add diagnostics for each phase.
- Manually test success, failure, and cancel paths.

### 11. Function Swap / Model Sharing

Open work:

- Detect and catalog `MC.*` formulas.
- Replace with deterministic values.
- Store restore metadata safely in workbook custom XML.
- Restore formulas and validate workbook compatibility.

### 12. Performance Modes

Open work:

- Add explicit run presets and mode labels.
- Show iteration/sec and ETA in the run view.
- Warn before raw-data export for large simulations.
- Benchmark recalc mode separately from engine mode.

## First Three Recommended Tickets

1. Build the preflight validation model and surface it before Run.
2. Finish correlation workflow polish: import/export, warnings, and report inclusion.
3. Add the input/output manager table with jump-to-cell and edit/remove controls.
