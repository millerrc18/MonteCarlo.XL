# MonteCarlo.XL — Task Log

> This log is the shared memory between the **architect** (Claude chat) and the **developer** (Claude Code). The developer writes entries after each task. The architect reads them to plan the next task.

---

## Project Initialized
**Status**: COMPLETE
**Date**: 2026-03-19

### What Was Done
- Created GitHub repo: millerrc18/MonteCarlo.XL
- Added ROADMAP.md with full architecture, phased plan, and design system
- Added CLAUDE.md with developer workflow instructions
- Set up task management system (tasks/ directory structure)

### Notes
- Tech stack: VSTO (C#), WPF, Math.NET Numerics, LiveCharts2, SkiaSharp
- Windows-only (desktop Excel)
- Open question: VSTO vs ExcelDna — spike needed in Task 001

---

## TASK-001 — Solution Scaffolding + VSTO vs ExcelDna Spike
**Status**: COMPLETE
**Date**: 2026-03-19
**Branch**: claude/work-on-current-task-99CnH

### What Was Done
- Created the full solution structure with 5 projects (Engine, Charts, UI, Addin, Engine.Tests)
- Created `MonteCarlo.XL.sln` solution file with solution folders (src/, tests/)
- Set up all project references matching the required dependency graph
- Added all required NuGet package references
- Created `.gitignore` for Visual Studio / .NET projects
- Created placeholder classes: `IDistribution` interface (Engine), `ChartPlaceholder` (Charts), `MainViewModel` (UI), `AddIn` + `MonteCarloRibbon` (Addin), `IDistributionTests` (Tests)

### VSTO vs ExcelDna Decision: **ExcelDna**

Chose ExcelDna over VSTO for these reasons:

1. **Modern .NET**: ExcelDna supports .NET 8+, giving us access to latest C# features, better performance, and a unified target framework across the solution. VSTO locks us to .NET Framework 4.8.
2. **Simpler deployment**: ExcelDna produces a single `.xll` file — no VSTO runtime installation needed on target machines. This eliminates a major deployment friction point.
3. **Native UDF support**: ExcelDna has first-class support for Excel UDFs (`=MC.Normal()` etc.), which is planned for Phase 3. With VSTO, UDF registration requires additional COM automation plumbing.
4. **Custom ribbon**: ExcelDna supports ribbon XML customization via `ExcelRibbon` base class — equivalent to VSTO's ribbon support.
5. **WPF task pane**: ExcelDna can host WPF content in custom task panes via `ExcelDna.Integration.CustomUI` — confirmed by the ExcelDna documentation and community examples.
6. **Lighter ceremony**: No Visual Studio Office development workload required. Projects can be created and built with standard `dotnet` CLI.

**Framework targets chosen:**
- `MonteCarlo.Engine` → `net8.0` (pure computation, no UI dependency)
- `MonteCarlo.Charts` → `net8.0-windows` (WPF)
- `MonteCarlo.UI` → `net8.0-windows` (WPF)
- `MonteCarlo.Addin` → `net8.0-windows` (ExcelDna + WPF host)
- `MonteCarlo.Engine.Tests` → `net8.0`

### Files Created/Modified
- `MonteCarlo.XL.sln` — solution file with all 5 projects in src/tests solution folders
- `.gitignore` — Visual Studio + .NET + ExcelDna ignores
- `src/MonteCarlo.Engine/MonteCarlo.Engine.csproj` — class library, MathNet.Numerics
- `src/MonteCarlo.Engine/Distributions/IDistribution.cs` — core distribution interface
- `src/MonteCarlo.Charts/MonteCarlo.Charts.csproj` — WPF lib, LiveCharts2 + SkiaSharp
- `src/MonteCarlo.Charts/ChartPlaceholder.cs` — placeholder
- `src/MonteCarlo.UI/MonteCarlo.UI.csproj` — WPF lib, CommunityToolkit.Mvvm
- `src/MonteCarlo.UI/ViewModels/MainViewModel.cs` — placeholder MVVM view model
- `src/MonteCarlo.Addin/MonteCarlo.Addin.csproj` — ExcelDna.AddIn
- `src/MonteCarlo.Addin/AddIn.cs` — IExcelAddIn entry point (AutoOpen/AutoClose)
- `src/MonteCarlo.Addin/MonteCarloRibbon.cs` — Custom ribbon tab with Run/Stop/TaskPane/Settings buttons
- `tests/MonteCarlo.Engine.Tests/MonteCarlo.Engine.Tests.csproj` — xUnit + FluentAssertions + Moq
- `tests/MonteCarlo.Engine.Tests/Distributions/IDistributionTests.cs` — placeholder tests verifying interface contract

### Key Decisions Made During Implementation
- Used `net8.0` for Engine (not .NET Standard 2.0) since ExcelDna allows modern .NET — simpler than maintaining netstandard2.0 compatibility
- Created `IDistribution` as the real Engine placeholder (not a throwaway class) since it's the foundational interface for TASK-002
- Included a proper `MonteCarloRibbon` with XML and stub callbacks — more useful than a bare placeholder for verifying ExcelDna ribbon support

### Issues / Notes for Architect
- **No `dotnet` SDK in this environment** — could not run `dotnet build` or `dotnet test`. All project files were hand-written. They follow standard MSBuild/SDK-style conventions and should build cleanly on any machine with .NET 8 SDK installed.
- **Excel smoke test not possible** — no Excel or .NET SDK available in this environment. The add-in structure follows ExcelDna conventions and should load correctly when built and opened in Excel.
- **LiveCharts2 version**: Used `2.0.0-rc3.3` which is the latest stable RC. LiveCharts2 hasn't had a non-RC 2.x release yet, so this is the correct package to use.

### Test Results
- Cannot run tests (no dotnet SDK). Two placeholder tests written that verify the `IDistribution` interface exists and defines the required members.

---

## TASK-002 — Distribution Module Implementation
**Status**: COMPLETE
**Date**: 2026-03-19
**Branch**: claude/work-on-current-task-99CnH

### What Was Done
- Updated `IDistribution` interface with full contract: added `Variance`, `StdDev`, `Minimum`, `Maximum`, `Sample(int count)`, `ParameterSummary()`, and capitalized `PDF`/`CDF` method names
- Implemented all 6 Phase 1 distributions:
  - **NormalDistribution** — wraps MathNet `Normal`
  - **TriangularDistribution** — wraps MathNet `Triangular`
  - **PERTDistribution** — implemented as scaled `Beta(alpha, beta)` with configurable lambda (default 4)
  - **LognormalDistribution** — wraps MathNet `LogNormal`, parameters are μ/σ of underlying normal
  - **UniformDistribution** — wraps MathNet `ContinuousUniform`
  - **DiscreteDistribution** — uses MathNet `Categorical` with sorted values for consistent CDF/Percentile
- Implemented `DistributionFactory` with case-insensitive name lookup and parameter dictionary
- Wrote comprehensive test suite: **143 tests, all passing**

### Files Created/Modified
- `src/MonteCarlo.Engine/Distributions/IDistribution.cs` — updated with full contract
- `src/MonteCarlo.Engine/Distributions/NormalDistribution.cs` — new
- `src/MonteCarlo.Engine/Distributions/TriangularDistribution.cs` — new
- `src/MonteCarlo.Engine/Distributions/PERTDistribution.cs` — new
- `src/MonteCarlo.Engine/Distributions/LognormalDistribution.cs` — new
- `src/MonteCarlo.Engine/Distributions/UniformDistribution.cs` — new
- `src/MonteCarlo.Engine/Distributions/DiscreteDistribution.cs` — new
- `src/MonteCarlo.Engine/Distributions/DistributionFactory.cs` — new
- `tests/MonteCarlo.Engine.Tests/Distributions/IDistributionTests.cs` — updated
- `tests/MonteCarlo.Engine.Tests/Distributions/NormalDistributionTests.cs` — new
- `tests/MonteCarlo.Engine.Tests/Distributions/TriangularDistributionTests.cs` — new
- `tests/MonteCarlo.Engine.Tests/Distributions/PERTDistributionTests.cs` — new
- `tests/MonteCarlo.Engine.Tests/Distributions/LognormalDistributionTests.cs` — new
- `tests/MonteCarlo.Engine.Tests/Distributions/UniformDistributionTests.cs` — new
- `tests/MonteCarlo.Engine.Tests/Distributions/DiscreteDistributionTests.cs` — new
- `tests/MonteCarlo.Engine.Tests/Distributions/DistributionFactoryTests.cs` — new

### Key Decisions Made During Implementation
- **PERT variance**: Computed as `range² × Beta.Variance` rather than the approximate formula, since we have access to the exact Beta distribution object
- **DiscreteDistribution sorts values**: Input values/probabilities are sorted by value at construction time so CDF and Percentile behave consistently regardless of input order
- **DistributionFactory Discrete encoding**: Uses indexed `value_0`/`prob_0`, `value_1`/`prob_1` keys since `Dictionary<string, double>` can't directly represent arrays — pragmatic trade-off for the factory pattern
- **Method capitalization**: Used `PDF`/`CDF` (uppercase) per the task spec, changing from the TASK-001 placeholder's `Pdf`/`Cdf`

### Issues / Notes for Architect
- **WPF projects don't build on Linux**: Expected — the `net8.0-windows` projects (Charts, UI, Addin) require the Windows Desktop SDK. Engine and Engine.Tests build and test cleanly on Linux.
- **`dotnet build` at solution level shows 4 errors** from WPF projects but Engine builds with 0 warnings. On Windows with the .NET Desktop workload, the full solution should build cleanly.
- All distributions accept an optional `int? seed` for reproducible RNG — confirmed by reproducibility tests.

### Test Results
- **143 tests, 143 passed, 0 failed** in ~1 second
- Test coverage per distribution: construction validation, statistical convergence (100k samples within 1% of mean / 2% of stddev), quantile round-trip (CDF∘Percentile ≈ identity), PDF numerical integration ≈ 1.0, CDF boundary conditions, seed reproducibility
- PERT-specific: mean formula validation, symmetric case, custom lambda, lower variance than Triangular
- Discrete-specific: frequency convergence, step-function CDF, unsorted input handling
- Factory: all 6 types, case-insensitive, seed passthrough, error cases
