# MonteCarlo.XL Quickstart

This guide gets a local build loaded into Excel and runs a first simulation from the included smoke workbook. The production path is still the packed x64 `.xll`; an experimental Office.js plus WebAssembly host now exists separately for ARM-oriented testing.

## 1. Build And Install The Add-In

From Windows PowerShell in the repository root:

```powershell
dotnet build src/MonteCarlo.Addin/MonteCarlo.Addin.csproj -c Release
.\scripts\install-local-addin.ps1 -Configuration Release
```

The install script detects Excel bitness, copies the correct packed `.xll` into your user `XLSTART` folder, and unblocks the file. On ARM64 Windows, it now stops by default because the current add-in is not a native ARM64 Excel build. On the local development machine, Excel is 64-bit, so the packed add-in is:

```text
src/MonteCarlo.Addin/bin/Release/net8.0-windows/publish/MonteCarlo.Addin-AddIn64-packed.xll
```

Restart Excel after installing. The ribbon should show a `MonteCarlo.XL` tab.

If you are working on a Windows ARM device and intentionally want to try an unsupported setup, rerun the install script with `-AllowUnsupportedArm`. See [ARM64 support status](ARM64_SUPPORT.md) before doing that.

## 2. Open The Smoke Workbook

The sample smoke workbook is checked in at:

```text
samples/MonteCarlo.XL.Smoke.xlsx
```

If you need to regenerate it:

```powershell
.\samples\create-smoke-workbook.ps1
```

Open the workbook in Excel. It contains a simple model with uncertain inputs using `MC.Normal`, `MC.Triangular`, and `MC.PERT` formulas.

## 3. Run Your First Simulation

1. Open the `MonteCarlo.XL` ribbon tab.
2. Click `Task Pane`.
3. Add output cell `B9` on the `Smoke Model` sheet.
4. Confirm the detected input formulas or add inputs manually. If you are adding an input from scratch, use the distribution helper's category, complexity, and search filters for a plain-English starting point, or use `Select Range` to fit from historical Excel data.
5. For correlated inputs, click `Correlations`, enter pairwise Spearman correlations, export a labeled template, or select the top-left cell of a labeled/numeric matrix and use `Import Range`.
6. Click `Model Check` to validate the setup. It checks the saved configuration plus Excel-specific issues such as missing sheets, protected input cells, formula errors, non-numeric outputs, and manual calculation mode.
7. Choose the `Preview` run preset for a quick 1,000-iteration run.
8. Click `Run Simulation`.

During the run, the task pane shows the run scale, completed iterations, iteration/sec speed, elapsed time, and estimated remaining time. The results view should show summary statistics and charts for the selected output.

Use `Target Analysis` to enter a threshold and see the probability above or below it. Use `Scenario Analysis` to compare the input assumptions behind worst-tail, best-tail, and target-hit outcomes.

## 4. Export Results

After a simulation finishes:

1. Select the output you want to export in the results view.
2. Use `Export Summary` for workbook/run metadata, statistics, percentiles, target analysis, scenario analysis, input assumptions, correlation assumptions, and histogram/CDF/sensitivity charts. If the simulation has multiple outputs, Excel now asks whether to export only the selected output or combine all outputs into one report sheet. The exported summary sheet is now formatted for cleaner printing or PDF export with landscape, fit-to-width, repeat-header, and report page-break layout. If a target threshold is entered first, the report also includes target-hit scenario comparisons.
3. Use `Export Raw Data` for iteration-level input and output samples.

By default, each export creates a new worksheet so prior summaries are preserved. To change that behavior, open `Settings` and turn off `Create a new worksheet for each export`.

Raw-data export can create very large worksheets. MonteCarlo.XL warns before large exports and blocks exports that exceed Excel's worksheet row limit.

Settings also lets you choose defaults for new workbook setups: run preset or custom iteration count, random versus fixed versus prompt-at-run seed behavior, sampling method, convergence auto-stop, Excel calculation behavior during runs, whether to suspend screen updating and events, whether warning-only Model Check reports pause before a run, and the percentile list used in summary exports.

Use the Settings scope toggle if you want those defaults to apply only to the active workbook. Workbook overrides are saved inside the workbook and take precedence over your Windows-user defaults for that workbook only.

## 5. Goal Seek Under Uncertainty

Use `Goal Seek` when you want to find the deterministic decision value needed to reach a probability target.

1. Select a deterministic decision cell in Excel. Do not select a Monte Carlo input cell.
2. Click `Goal Seek` on the `MonteCarlo.XL` ribbon.
3. Choose the output, metric, lower and upper decision bounds, any target or percentile inputs required by that metric, the desired metric value, iterations per trial, and solver tolerance.
4. Run the workflow. MonteCarlo.XL tests decision values across the bounds, restores the original decision cell value, and adds a `MC Goal Seek` report sheet with the solver history.

The ribbon dialog now supports `P(output > target)`, `P(output <= target)`, mean, and percentile metrics. A richer task-pane workflow, progress UI, and confidence guidance remain on the roadmap.

## 6. Stress Analysis

Use `Stress Analysis` when you want a baseline-vs-stressed comparison instead of only looking at tail cases.

1. Open the `MonteCarlo.XL` ribbon and click `Stress Analysis`.
2. Choose the primary output to highlight in the report.
3. Pick one or more configured inputs and apply a stress rule:
   `Fixed value`, `Add shift`, or `Range scale`.
4. Run the workflow. MonteCarlo.XL executes the current model twice using the same comparison seed, once baseline and once stressed, then adds an `MC Stress Analysis` sheet.

The comparison sheet includes the stressed assumptions, an output impact ranking, and histogram/CDF comparison charts for the primary output.

## 7. Add Monte Carlo Formulas To Your Own Workbook

Use `MC.*` formulas where your deterministic model has uncertain assumptions. For example:

```excel
=MC.Normal(100000, 15000)
=MC.Triangular(80, 100, 140)
=MC.PERT(50000, 75000, 125000)
=MC.Binomial(20, 0.35)
```

When Excel is not simulating, these formulas return an expected or representative value so the workbook still calculates normally. During a simulation, MonteCarlo.XL samples those cells and recalculates the selected output cells repeatedly.

## 8. Share A Workbook Without The Add-In

If someone needs to open the workbook without MonteCarlo.XL:

1. Open the `MonteCarlo.XL` ribbon tab.
2. Click `Save Shareable Copy`.
3. Choose a destination path.
4. Send the value-only copy.

The active workbook stays formula-backed. The older `Replace MC Formulas` and `Restore MC Formulas` commands are still available when you intentionally want to flatten and later restore the current workbook.

## 9. Troubleshooting

If the ribbon does not appear, the task pane does not open, or a simulation/export fails, check:

```text
%LOCALAPPDATA%\MonteCarlo.XL\Logs\startup.log
```

Common fixes:

- Close all Excel instances before reinstalling the `.xll`.
- Use the 64-bit packed `.xll` with 64-bit Excel.
- On Windows ARM, expect the install script to stop unless you explicitly pass `-AllowUnsupportedArm`.
- Put the `.xll` in a trusted location or unblock it if Excel blocks loading.
- Rebuild after code changes before copying the packed add-in to `XLSTART`.
- Use `MonteCarlo.XL` > `Recover Excel` if Excel appears stuck in manual calculation, disabled events, or a stale status-bar message after an interrupted run.
- Use `MonteCarlo.XL` > `Benchmark` to add a sheet comparing active-workbook recalculation time with synthetic simulation-engine throughput plus synthetic summary/raw export timing.

See [Local Excel Debug Path](LOCAL_EXCEL_DEBUG.md) for more detailed developer notes.

## 10. Experimental ARM / Office.js Host

If you are testing native Windows ARM Excel, do **not** use the `install-local-addin.ps1` flow above as your main path. The current ARM track is the new Office.js host:

```powershell
cd src\MonteCarlo.OfficeAddin
npm install
npm run dev
```

Then sideload `src\MonteCarlo.OfficeAddin\manifest.xml` into desktop Excel and let it connect to `https://localhost:3000`.

What works today:

- the task pane builds and loads from the Office host foundation
- `MC.*` custom functions are bundled for the Office host
- the shared WebAssembly bridge can validate profiles, generate input samples, analyze results, and benchmark
- the Office task pane can scan formulas, add selected inputs and outputs, run a simulation, and export a summary sheet with native Excel charts

What still needs real ARM verification:

- manual Surface Pro acceptance
- deeper ribbon parity
- broader persistence and reopen testing
- richer task-pane workflows such as the full settings, sharing, and goal-seek experience

For a Surface-specific install checklist, including the shared-folder sideload flow and the `localhost` loopback fix, see [SURFACE_INSTALL.md](SURFACE_INSTALL.md).
