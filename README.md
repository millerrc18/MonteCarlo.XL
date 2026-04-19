# MonteCarlo.XL

MonteCarlo.XL is a Windows desktop Excel add-in for Monte Carlo simulation, built with Excel-DNA, C#/.NET 8, WPF, and a pure simulation engine. The product goal is an @RISK-style workflow inside existing Excel workbooks: define uncertain inputs, choose output cells, run simulations, and inspect modern histogram/CDF/tornado-style results.

## Current Direction

- Excel remains the primary product surface.
- Excel-DNA is the add-in host, not VSTO.
- `MonteCarlo.Engine` stays pure and reusable, with no Excel or UI dependency.
- A standalone WPF EXE can be added later as a demo/development harness, after the Excel workflow runs end-to-end.

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
