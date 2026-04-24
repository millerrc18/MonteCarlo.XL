# MonteCarlo.XL

MonteCarlo.XL is an Excel Monte Carlo simulation tool built around an @RISK-style workflow inside existing workbooks: define uncertain inputs, choose output cells, run simulations, and inspect histogram/CDF/tornado-style results. Today the production host is an Excel-DNA `.xll` for desktop Excel on Windows x64, and the repository now also contains an Office.js plus .NET WebAssembly host foundation aimed at native Windows ARM Excel.

## Quickstart

Start here if you want to try the add-in in Excel:

- [Quickstart guide](docs/QUICKSTART.md) - install the packed XLL, open the smoke workbook, run a first simulation, and export results.
- [Distribution guide](docs/DISTRIBUTION_GUIDE.md) - plain-English explanations of every supported distribution, with formulas, parameters, and real modeling examples.
- [@RISK-parity roadmap](docs/RISK_PARITY_ROADMAP.md) - 12 robustness and analyst-workflow initiatives with status and unfinished work.
- [ARM64 support status](docs/ARM64_SUPPORT.md) - what works today, what does not, and the real paths to native Windows ARM support.
- [Surface Pro install guide](docs/SURFACE_INSTALL.md) - the current experimental sideload checklist for Windows ARM tablets.
- [Local Excel debug path](docs/LOCAL_EXCEL_DEBUG.md) - developer-focused build, install, smoke test, and diagnostics notes.

## Current Direction

- Excel remains the primary product surface.
- Excel-DNA is the add-in host, not VSTO.
- `MonteCarlo.Engine` stays pure and reusable, with no Excel or UI dependency.
- ARM is now being pursued through a dual-host path: shared simulation/formula logic plus an Office.js task-pane and custom-functions host for native ARM Excel.
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

For the experimental Office.js and ARM-oriented host:

```bash
cd src/MonteCarlo.OfficeAddin
cmd.exe /C npm install
cmd.exe /C npm run build
```

That build publishes the .NET WebAssembly bridge into `src/MonteCarlo.OfficeAddin/public/wasm`, bundles the Office task pane into `src/MonteCarlo.OfficeAddin/dist`, and leaves the sideload manifest at `src/MonteCarlo.OfficeAddin/manifest.xml`.

## Work Install

MonteCarlo.XL is meant for desktop Excel on Windows. The current packaged add-in is the 64-bit `.xll`, so the best-fit work setup is:

- Excel for Windows desktop
- 64-bit Excel
- permission to load custom `.xll` add-ins from a trusted local folder or `XLSTART`

It is not a fit for Excel for web or Excel for Mac. On Windows ARM, there is still no native ARM64 Excel-DNA build in this project today.

The local install script now stops on ARM64 systems by default so it does not copy an add-in into a native ARM64 Excel setup that cannot load it. See [ARM64 support status](docs/ARM64_SUPPORT.md) for the real upgrade paths.

The repo now also contains an **experimental Office.js dual-host foundation** for ARM work:

- `src/MonteCarlo.Shared` shares formula catalog, parser, normal-mode evaluation, workbook-config constants, and simulation analysis contracts.
- `src/MonteCarlo.Engine.Wasm` exposes the engine through `[JSExport]` methods in a browser build.
- `src/MonteCarlo.OfficeAddin` contains a React plus Fluent UI task pane, custom functions, workbook scanning, simulation execution, Excel chart export, and an Office manifest for sideload testing.

That Office host builds today, but it is still a foundation project rather than a fully validated replacement for the x64 `.xll`. The production recommendation for workbooks you need to trust remains the x64 Excel-DNA add-in until the ARM path is manually validated on real hardware.

The most common work blocker is corporate policy rather than Excel itself. Some organizations block unsigned/custom add-ins, restrict trusted locations, or lock down `XLSTART`. If your company allows local add-ins, load the packed `.xll` and confirm the `MonteCarlo.XL` ribbon appears. If Excel blocks it, you will usually need a trusted location or IT approval.

## Surface / ARM Trial

If you are trying MonteCarlo.XL on a **Surface Pro running Windows on ARM**, do **not** start with the Excel-DNA `.xll` installer. The current ARM path is the experimental Office.js host.

High-level flow:

1. Open PowerShell on the Surface
2. Run:

```powershell
cd "C:\path\to\MonteCarlo.XL\src\MonteCarlo.OfficeAddin"
npm install
npm run dev
```

3. Open `https://localhost:3000/taskpane.html` in Edge and trust the local HTTPS certificate if prompted
4. Sideload `src\MonteCarlo.OfficeAddin\manifest.xml` into desktop Excel through a trusted shared-folder catalog

Use the full checklist in [docs/SURFACE_INSTALL.md](docs/SURFACE_INSTALL.md). That guide also covers:

- the `localhost` loopback fix if Excel blocks the add-in
- the shared-folder catalog setup in Excel
- first-run smoke validation on the Surface
- the current limits of the ARM path

## User Settings

Open `MonteCarlo.XL > Settings` in the ribbon/task pane to choose Light, Dark, or System theme. The settings page also includes a results-export option. By default, each export creates a new worksheet so earlier summaries remain available; turn that off if you prefer to replace the prior export sheet.

Simulation and export defaults can now be edited either as Windows-user defaults or as workbook-specific overrides stored inside the active workbook. That includes random, fixed, or prompt-at-run seed behavior plus Excel execution defaults such as calculation mode, screen updating, and events during runs. The Run view and exported summary metadata show the effective settings source used for the run.

`Export Summary` creates an Excel report sheet with workbook/run metadata, statistics, percentiles, target analysis, scenario analysis, assumptions, correlations, and summary charts. Enter a target threshold in the Results view before exporting if you want the report to include probability-above and probability-below calculations for that target.
