# MonteCarlo.XL

MonteCarlo.XL is a Windows desktop Excel add-in for Monte Carlo simulation, built with Excel-DNA, C#/.NET 8, WPF, and a pure simulation engine. The product goal is an @RISK-style workflow inside existing Excel workbooks: define uncertain inputs, choose output cells, run simulations, and inspect modern histogram/CDF/tornado-style results.

## Quickstart

Start here if you want to try the add-in in Excel:

- [Quickstart guide](docs/QUICKSTART.md) - install the packed XLL, open the smoke workbook, run a first simulation, and export results.
- [Distribution guide](docs/DISTRIBUTION_GUIDE.md) - plain-English explanations of every supported distribution, with formulas, parameters, and real modeling examples.
- [Local Excel debug path](docs/LOCAL_EXCEL_DEBUG.md) - developer-focused build, install, smoke test, and diagnostics notes.

## Current Direction

- Excel remains the primary product surface.
- Excel-DNA is the add-in host, not VSTO.
- `MonteCarlo.Engine` stays pure and reusable, with no Excel or UI dependency.
- A standalone WPF EXE can be added later as a demo/development harness, after the Excel workflow runs end-to-end.

## Supported Distributions

The task pane supports Normal, Triangular, PERT, Lognormal, Uniform, Discrete, Beta, Weibull, Exponential, Poisson, Gamma, Logistic, GEV, Binomial, and Geometric distributions. Worksheet UDFs are available for all except Discrete; see the [distribution guide](docs/DISTRIBUTION_GUIDE.md) for details and usage examples.

When not running a simulation, `MC.*` worksheet functions return an expected or representative value so the workbook stays usable as a normal deterministic model. During simulation, MonteCarlo.XL replaces those input cells with sampled values and records the selected output cells.

## Build

From Windows PowerShell or WSL using Windows `dotnet.exe`:

```powershell
dotnet build MonteCarlo.XL.sln
dotnet test tests/MonteCarlo.Engine.Tests
```

On this machine from WSL:

```bash
"/mnt/c/Program Files/dotnet/dotnet.exe" build MonteCarlo.XL.sln --no-restore
"/mnt/c/Program Files/dotnet/dotnet.exe" test MonteCarlo.XL.sln --no-build
```

For 64-bit Excel, load:

```text
src/MonteCarlo.Addin/bin/Debug/net8.0-windows/publish/MonteCarlo.Addin-AddIn64-packed.xll
```

See `docs/LOCAL_EXCEL_DEBUG.md` for the local Excel install/debug path.

## User Settings

Open `MonteCarlo.XL > Settings` in the ribbon/task pane to choose Light, Dark, or System theme. The settings page also includes a results-export option. By default, each export creates a new worksheet so earlier summaries remain available; turn that off if you prefer to replace the prior export sheet.
