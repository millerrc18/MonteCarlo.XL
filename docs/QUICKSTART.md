# MonteCarlo.XL Quickstart

This guide gets a local build loaded into Excel and runs a first simulation from the included smoke workbook.

## 1. Build And Install The Add-In

From Windows PowerShell in the repository root:

```powershell
dotnet build src/MonteCarlo.Addin/MonteCarlo.Addin.csproj -c Release
.\scripts\install-local-addin.ps1 -Configuration Release
```

The install script detects Excel bitness, copies the correct packed `.xll` into your user `XLSTART` folder, and unblocks the file. On the local development machine, Excel is 64-bit, so the packed add-in is:

```text
src/MonteCarlo.Addin/bin/Release/net8.0-windows/publish/MonteCarlo.Addin-AddIn64-packed.xll
```

Restart Excel after installing. The ribbon should show a `MonteCarlo.XL` tab.

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
4. Confirm the detected input formulas or add inputs manually. If you are adding an input from scratch, use the distribution helper's category, complexity, and search filters for a plain-English starting point, or select historical data in Excel and use `Fit Range`.
5. For correlated inputs, click `Correlations`, enter pairwise Spearman correlations, or select the top-left cell of an n by n numeric matrix and use `Import Range`.
6. Click `Model Check` to validate the setup. It checks the saved configuration plus Excel-specific issues such as missing sheets, protected input cells, formula errors, non-numeric outputs, and manual calculation mode.
7. Choose the `Preview` run preset for a quick 1,000-iteration run.
8. Click `Run Simulation`.

During the run, the task pane shows the run scale, completed iterations, iteration/sec speed, elapsed time, and estimated remaining time. The results view should show summary statistics and charts for the selected output.

Use `Target Analysis` to enter a threshold and see the probability above or below it. Use `Scenario Analysis` to compare the input assumptions behind worst-tail, best-tail, and target-hit outcomes.

## 4. Export Results

After a simulation finishes:

1. Select the output you want to export in the results view.
2. Use `Export Summary` for statistics, percentiles, scenario analysis, input assumptions, correlation assumptions, and summary charts.
3. Use `Export Raw Data` for iteration-level input and output samples.

By default, each export creates a new worksheet so prior summaries are preserved. To change that behavior, open `Settings` and turn off `Create a new worksheet for each export`.

Raw-data export can create very large worksheets. MonteCarlo.XL warns before large exports and blocks exports that exceed Excel's worksheet row limit.

Settings also lets you choose defaults for new workbook setups: run preset or custom iteration count, random versus fixed seed, sampling method, convergence auto-stop, whether warning-only Model Check reports pause before a run, and the percentile list used in summary exports.

## 5. Add Monte Carlo Formulas To Your Own Workbook

Use `MC.*` formulas where your deterministic model has uncertain assumptions. For example:

```excel
=MC.Normal(100000, 15000)
=MC.Triangular(80, 100, 140)
=MC.PERT(50000, 75000, 125000)
=MC.Binomial(20, 0.35)
```

When Excel is not simulating, these formulas return an expected or representative value so the workbook still calculates normally. During a simulation, MonteCarlo.XL samples those cells and recalculates the selected output cells repeatedly.

## 6. Share A Workbook Without The Add-In

If someone needs to open the workbook without MonteCarlo.XL:

1. Save a normal backup copy first.
2. Open the `MonteCarlo.XL` ribbon tab.
3. Click `Replace MC Formulas` to replace `MC.*` formulas with their current values.
4. Send the value-only workbook.
5. Click `Restore MC Formulas` later to restore the formulas from the workbook restore map.

## 7. Troubleshooting

If the ribbon does not appear, the task pane does not open, or a simulation/export fails, check:

```text
%LOCALAPPDATA%\MonteCarlo.XL\Logs\startup.log
```

Common fixes:

- Close all Excel instances before reinstalling the `.xll`.
- Use the 64-bit packed `.xll` with 64-bit Excel.
- Put the `.xll` in a trusted location or unblock it if Excel blocks loading.
- Rebuild after code changes before copying the packed add-in to `XLSTART`.
- Use `MonteCarlo.XL` > `Recover Excel` if Excel appears stuck in manual calculation, disabled events, or a stale status-bar message after an interrupted run.

See [Local Excel Debug Path](LOCAL_EXCEL_DEBUG.md) for more detailed developer notes.
