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
4. Confirm the detected input formulas or add inputs manually. If you are adding an input from scratch, use the distribution helper for a plain-English starting point, or select historical data in Excel and use `Fit Range`.
5. For correlated inputs, click `Correlations`, enter pairwise Spearman correlations, or select the top-left cell of an n by n numeric matrix and use `Import Range`.
6. Click `Model Check` to validate the setup.
7. Set iterations to `1,000` for a quick run.
8. Click `Run Simulation`.

The results view should show summary statistics and charts for the selected output.

## 4. Export Results

After a simulation finishes:

1. Select the output you want to export in the results view.
2. Use `Export Summary` for statistics, percentiles, assumptions, and summary charts.
3. Use `Export Raw Data` for iteration-level input and output samples.

By default, each export creates a new worksheet so prior summaries are preserved. To change that behavior, open `Settings` and turn off `Create a new worksheet for each export`.

Settings also lets you choose defaults for new workbook setups: iteration count, random versus fixed seed, sampling method, convergence auto-stop, and the percentile list used in summary exports.

## 5. Add Monte Carlo Formulas To Your Own Workbook

Use `MC.*` formulas where your deterministic model has uncertain assumptions. For example:

```excel
=MC.Normal(100000, 15000)
=MC.Triangular(80, 100, 140)
=MC.PERT(50000, 75000, 125000)
=MC.Binomial(20, 0.35)
```

When Excel is not simulating, these formulas return an expected or representative value so the workbook still calculates normally. During a simulation, MonteCarlo.XL samples those cells and recalculates the selected output cells repeatedly.

## 6. Troubleshooting

If the ribbon does not appear, the task pane does not open, or a simulation/export fails, check:

```text
%LOCALAPPDATA%\MonteCarlo.XL\Logs\startup.log
```

Common fixes:

- Close all Excel instances before reinstalling the `.xll`.
- Use the 64-bit packed `.xll` with 64-bit Excel.
- Put the `.xll` in a trusted location or unblock it if Excel blocks loading.
- Rebuild after code changes before copying the packed add-in to `XLSTART`.

See [Local Excel Debug Path](LOCAL_EXCEL_DEBUG.md) for more detailed developer notes.
