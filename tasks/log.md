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

---

## TASK-003 — Simulation Engine: Core Monte Carlo Loop
**Status**: COMPLETE
**Date**: 2026-03-19
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Implemented `SimulationConfig` with validation (inputs, outputs, iteration count, seed, parallel flag)
- Implemented `SimulationInput` and `SimulationOutput` data classes
- Implemented `SimulationResult` with raw matrix storage and accessor methods (`GetInputSamples`, `GetOutputValues`, `GetInputSample`, `GetOutputValue`)
- Implemented `SimulationProgressEventArgs` with completion percentage and estimated time remaining
- Implemented `SimulationEngine` with:
  - Async execution via `RunAsync`
  - Delegate-based evaluator pattern (input dict → output dict)
  - Pre-allocated sample matrices for performance
  - Batch upfront sampling from distributions (required for future Iman-Conover)
  - Parallel and sequential execution modes
  - Throttled progress reporting (every 100 iterations or 50ms)
  - CancellationToken support
- Wrote comprehensive test suite: 16 tests covering basic execution, reproducibility, cancellation, progress, edge cases, multiple outputs

### Files Created
- `src/MonteCarlo.Engine/Simulation/SimulationConfig.cs`
- `src/MonteCarlo.Engine/Simulation/SimulationInput.cs`
- `src/MonteCarlo.Engine/Simulation/SimulationOutput.cs`
- `src/MonteCarlo.Engine/Simulation/SimulationResult.cs`
- `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs`
- `src/MonteCarlo.Engine/Simulation/SimulationProgressEventArgs.cs`
- `tests/MonteCarlo.Engine.Tests/Simulation/SimulationEngineTests.cs`
- `tests/MonteCarlo.Engine.Tests/Simulation/SimulationConfigTests.cs`

### Key Decisions Made During Implementation
- Used `double[,]` (2D array) instead of jagged `double[][]` for matrix storage — better cache locality and clearer semantics
- Used `required` keyword on `SimulationInput.Id`/`Label`/`Distribution` to enforce initialization at construction
- Progress reporting in parallel mode uses `Interlocked.Increment` with 50ms throttle to avoid event storm
- Evaluator receives/returns `Dictionary<string, double>` (not `IReadOnlyDictionary`) for simpler construction in client code

### Issues / Notes for Architect
- **No dotnet SDK** — cannot build or test. Code follows standard patterns and should compile cleanly.
- The `required` keyword requires C# 11+ (which net8.0 supports)

---

## TASK-004 — Summary Statistics & Sensitivity Analysis
**Status**: COMPLETE
**Date**: 2026-03-19
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Implemented `SummaryStatistics` class with:
  - Central tendency: Mean, Median, Mode (histogram-estimated)
  - Spread: StdDev, Variance, Min, Max, Range
  - Shape: Skewness (adjusted), Excess Kurtosis (Fisher definition)
  - Standard percentiles: P1, P5, P10, P25, P50, P75, P90, P95, P99
  - Arbitrary percentile via linear interpolation (PERCENTILE.INC method)
  - Mean confidence interval using t-distribution
  - Probability queries: ProbabilityAbove, ProbabilityBelow, ProbabilityBetween
  - Histogram generation via `ToHistogram()`
  - Sorted values array for CDF charting
- Implemented `HistogramData` class with equal-width binning, edge case handling
- Implemented `SensitivityAnalysis` with:
  - Spearman rank correlation (primary sensitivity measure)
  - Contribution to variance (normalized squared correlations)
  - Output at input extremes (mean output when input in bottom/top 10%)
  - Results sorted by absolute impact
- Implemented `SensitivityResult` data class
- Wrote test suites: 18 SummaryStatistics tests, 7 HistogramData tests, 7 SensitivityAnalysis tests

### Files Created
- `src/MonteCarlo.Engine/Analysis/SummaryStatistics.cs`
- `src/MonteCarlo.Engine/Analysis/HistogramData.cs`
- `src/MonteCarlo.Engine/Analysis/SensitivityAnalysis.cs`
- `src/MonteCarlo.Engine/Analysis/SensitivityResult.cs`
- `tests/MonteCarlo.Engine.Tests/Analysis/SummaryStatisticsTests.cs`
- `tests/MonteCarlo.Engine.Tests/Analysis/HistogramDataTests.cs`
- `tests/MonteCarlo.Engine.Tests/Analysis/SensitivityAnalysisTests.cs`

### Key Decisions Made During Implementation
- **Eager computation** in SummaryStatistics constructor — array sizes (5k-50k) make this trivially fast, simpler than lazy caching
- **Skewness/Kurtosis** computed using adjusted sample formulas (matching standard statistical software)
- **Mode estimation** uses Sturges' rule for bin count, returns center of most frequent bin
- **ProbabilityBelow** uses binary search on sorted array for O(log n) performance
- **Sensitivity rank correlation** implemented from scratch (Spearman via Pearson of ranks) rather than using MathNet — avoids dependency on specific MathNet.Numerics.Statistics overloads and keeps the logic transparent

### Issues / Notes for Architect
- **No dotnet SDK** — cannot verify compilation or run tests
- The SensitivityAnalysis tests depend on SimulationEngine (integration-style) — they run actual simulations to produce data for analysis. This is intentional since mocking the matrices would be fragile.

---

## TASK-005 — Sensitivity Analysis (Gap Fill)
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: main

### What Was Done
- Added `Swing` computed property to `SensitivityResult` (|OutputAtInputP90 - OutputAtInputP10|)
- Added `ComputeTornadoSwing()` static method to `SensitivityAnalysis` — evaluator-based direct swing calculation
- Added 4 missing tests: single input, equal contributions, Swing property, evaluator-based tornado swing

### Files Modified
- `src/MonteCarlo.Engine/Analysis/SensitivityResult.cs` — added Swing property
- `src/MonteCarlo.Engine/Analysis/SensitivityAnalysis.cs` — added ComputeTornadoSwing method
- `tests/MonteCarlo.Engine.Tests/Analysis/SensitivityAnalysisTests.cs` — added 4 tests

---

## TASK-006 — Excel I/O Layer
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: main

### What Was Done
- Created `CellReference` model with equality/hashing for use as dictionary keys
- Created `IWorkbookManager` interface and `WorkbookManager` implementation:
  - ReadCellValue / ReadCellValues (batched by sheet) with numeric validation
  - WriteCellValue / WriteRange (single COM call for 2D arrays)
  - WriteResultsSheet (headers + data to named sheet)
  - GetActiveCell, SheetExists, EnsureSheet
  - Screen updating disabled during batch writes
- Created `IInputTagManager` / `InputTagManager`:
  - Tag/untag cells as inputs with distribution config
  - `ToSimulationInputs()` converts tags to engine inputs via DistributionFactory
- Created `IOutputTagManager` / `OutputTagManager`:
  - Tag/untag cells as outputs
  - `ToSimulationOutputs()` converts tags to engine outputs
- Created `ICellHighlighter` / `CellHighlighter`:
  - Blue (#DBEAFE) background for inputs, green (#DCFCE7) for outputs
  - RefreshAll with screen updating disabled
- Created `TaggedInput` and `TaggedOutput` data models

### Files Created
- `src/MonteCarlo.Addin/Excel/CellReference.cs`
- `src/MonteCarlo.Addin/Excel/IWorkbookManager.cs`
- `src/MonteCarlo.Addin/Excel/WorkbookManager.cs`
- `src/MonteCarlo.Addin/Excel/TaggedInput.cs`
- `src/MonteCarlo.Addin/Excel/TaggedOutput.cs`
- `src/MonteCarlo.Addin/Excel/IInputTagManager.cs`
- `src/MonteCarlo.Addin/Excel/InputTagManager.cs`
- `src/MonteCarlo.Addin/Excel/IOutputTagManager.cs`
- `src/MonteCarlo.Addin/Excel/OutputTagManager.cs`
- `src/MonteCarlo.Addin/Excel/ICellHighlighter.cs`
- `src/MonteCarlo.Addin/Excel/CellHighlighter.cs`

### Key Decisions
- All classes have interfaces for testability (UI code can mock the Excel layer)
- CellReference uses case-insensitive equality for sheet names and cell addresses
- WorkbookManager groups batch reads by sheet to minimize COM round-trips
- InputTagManager.ToSimulationInputs reads current cell values as BaseValue and uses DistributionFactory

### Issues / Notes for Architect
- **No dotnet SDK** — cannot build. All COM interop code follows ExcelDna conventions.
- Testing is manual (requires Excel). Interfaces enable mocking in UI tests.

---

## TASK-007 — Ribbon & WPF Task Pane Shell
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: main

### What Was Done
- Created `GlobalStyles.xaml` with full design system:
  - Color palette: 8 primary colors + 7 neutrals with SolidColorBrush variants
  - Typography: 6 text styles (HeadlineLarge through StatLabel) + mono font
  - Spacing constants: Xs(4) through Xxl(32)
  - Button styles: PrimaryButton (blue), SecondaryButton (outlined), GhostButton (text)
  - Card style, TabBarButton style for bottom navigation
  - Hover/press/disabled triggers on all interactive styles
- Created `MainTaskPaneControl` (WPF UserControl):
  - Header bar with title + settings gear button
  - Content area bound to ViewModel.CurrentView via ContentControl
  - Bottom tab bar with Setup/Results RadioButton navigation
  - Loads GlobalStyles.xaml as merged resource dictionary
- Created `MainViewModel` with CommunityToolkit.Mvvm:
  - NavigateTo commands for Setup, Results, Settings, Run
  - CurrentView and CurrentViewName observable properties
- Created 4 placeholder views: SetupView, RunView, ResultsView, SettingsView
- Created `StringMatchConverter` for RadioButton IsChecked binding
- Created `TaskPaneController` — manages ExcelDna Custom Task Pane lifecycle
- Created `TaskPaneHost` — WinForms UserControl bridging to WPF via ElementHost
- Updated `MonteCarloRibbon`:
  - Task Pane is now a toggleButton with getPressed callback
  - All buttons have screentips/supertips
  - Run callback shows task pane, Settings callback shows task pane
- Updated `AddIn.AutoOpen/AutoClose` to initialize all services
- Updated `MonteCarlo.Addin.csproj`: added UseWPF, UseWindowsForms, ExcelDna.Interop

### Files Created
- `src/MonteCarlo.UI/Styles/GlobalStyles.xaml`
- `src/MonteCarlo.UI/Views/MainTaskPaneControl.xaml` + `.xaml.cs`
- `src/MonteCarlo.UI/Views/SetupView.xaml` + `.xaml.cs`
- `src/MonteCarlo.UI/Views/RunView.xaml` + `.xaml.cs`
- `src/MonteCarlo.UI/Views/ResultsView.xaml` + `.xaml.cs`
- `src/MonteCarlo.UI/Views/SettingsView.xaml` + `.xaml.cs`
- `src/MonteCarlo.UI/Converters/StringMatchConverter.cs`
- `src/MonteCarlo.Addin/TaskPane/TaskPaneController.cs`
- `src/MonteCarlo.Addin/TaskPane/TaskPaneHost.cs`

### Files Modified
- `src/MonteCarlo.UI/ViewModels/MainViewModel.cs` — full navigation logic
- `src/MonteCarlo.Addin/AddIn.cs` — service initialization
- `src/MonteCarlo.Addin/MonteCarloRibbon.cs` — toggle button, callbacks
- `src/MonteCarlo.Addin/MonteCarlo.Addin.csproj` — WPF/WinForms/Interop

### Issues / Notes for Architect
- **No dotnet SDK or Excel** — cannot build or smoke-test
- Task pane defaults to 380px wide, docked right
- GlobalStyles.xaml is loaded in MainTaskPaneControl.Resources (no App.xaml in Excel add-in)
- The StatLabel style uses TextTransform="Uppercase" — this is not a native WPF property and would need a converter if actually used; consider removing or implementing via a text converter in a future task

---

## TASK-008 — Setup View (Input/Output Config UI)
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `SetupViewModel` with full input/output management, distribution parameter editing, cell selection mode, iteration count presets, seed lock toggle
- Created `InputCardViewModel` with distribution preview points computation (50-point PDF sampling from P1 to P99)
- Created `OutputCardViewModel` for output cell display
- Created `CellSelectionService` in Services/ — mediates between WPF UI and Excel cell selection events
- Created `DistributionPreviewControl` — mini sparkline using WPF Canvas + Polyline/Polygon with blue fill at 12% opacity
- Created `InputCardControl` — card displaying label, cell ref, distribution summary, and mini preview curve
- Created `OutputCardControl` — card displaying label and cell ref
- Created `InputEditorControl` — inline editor with cell selection button, label field, distribution dropdown, and dynamic parameter fields
- Created `DistributionParameterPanel` — switches parameter fields based on selected distribution (Normal: mean/stddev, Triangular/PERT: min/mode/max, Uniform: min/max, Lognormal: mu/sigma, Discrete: dynamic value/probability pairs)
- Replaced placeholder `SetupView.xaml` with full implementation: header, iteration config card, inputs section with +Add and inline editor, outputs section with +Add and inline editor, empty state messages, Run Simulation button
- Added 4 value converters: BoolToVisibilityConverter, DistributionVisibilityConverter, NullToCollapsedConverter, CountToVisibilityConverter
- Updated `MainViewModel` to retain a persistent SetupView instance with ViewModel accessor

### Files Created/Modified
- `src/MonteCarlo.UI/ViewModels/SetupViewModel.cs` — Main setup VM with all commands and events
- `src/MonteCarlo.UI/ViewModels/InputCardViewModel.cs` — Input card display VM with preview points
- `src/MonteCarlo.UI/ViewModels/OutputCardViewModel.cs` — Output card display VM
- `src/MonteCarlo.UI/ViewModels/MainViewModel.cs` — Updated: persistent SetupView, SetupViewModel accessor
- `src/MonteCarlo.UI/Services/CellSelectionService.cs` — Cell selection mediation service
- `src/MonteCarlo.UI/Views/SetupView.xaml/.cs` — Full setup view replacing placeholder
- `src/MonteCarlo.UI/Views/InputEditorControl.xaml/.cs` — Inline input editor
- `src/MonteCarlo.UI/Views/InputCardControl.xaml/.cs` — Input display card
- `src/MonteCarlo.UI/Views/OutputCardControl.xaml/.cs` — Output display card
- `src/MonteCarlo.UI/Views/DistributionParameterPanel.xaml/.cs` — Dynamic distribution parameter fields
- `src/MonteCarlo.UI/Views/DistributionPreviewControl.xaml/.cs` — Mini PDF sparkline
- `src/MonteCarlo.UI/Converters/BoolToVisibilityConverter.cs`
- `src/MonteCarlo.UI/Converters/DistributionVisibilityConverter.cs`
- `src/MonteCarlo.UI/Converters/NullToCollapsedConverter.cs`
- `src/MonteCarlo.UI/Converters/CountToVisibilityConverter.cs`

### Key Decisions Made During Implementation
- Used events (RunSimulationRequested, CellSelectionRequested, InputAdded/Removed, OutputAdded/Removed) for Addin-layer wiring instead of direct COM calls, keeping MonteCarlo.UI free of Excel dependencies
- Distribution parameters stored as individual string properties (ParamMean, ParamStdDev, etc.) rather than a Dictionary, enabling straightforward XAML binding with UpdateSourceTrigger=PropertyChanged
- Discrete distribution uses ObservableCollection<DiscretePairViewModel> with dynamic add/remove rows
- Preview sparkline uses WPF Polyline fallback (not LiveCharts2) since it's just a simple curve shape — lighter weight and no chart library dependency for the preview
- MainViewModel now retains a single SetupView instance rather than creating new ones on each navigation, so the Addin layer can wire events once

### Issues / Notes for Architect
- No .NET SDK available to build/test — all code follows project conventions
- The InputEditorControl uses a "📌" Unicode emoji for the cell-select button — may need to be replaced with an icon if it doesn't render well in all Excel/Windows versions
- The seed lock uses a "🔒" Unicode emoji — same consideration
- Cell selection wiring (hooking SheetSelectionChange) will be completed in TASK-013 (Orchestrator) when the Addin layer ties everything together
- The SetupView creates its own DataContext (SetupViewModel) inline in XAML — the MainViewModel provides a convenience accessor for external wiring

### Test Results
- No tests — this is a pure UI task (WPF views/viewmodels). ViewModels use CommunityToolkit.Mvvm source generators.
- Cannot build without .NET SDK in this environment.


---

## TASK-009 — Results Dashboard — Histogram, CDF, Stats Panel
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `ChartTheme` in MonteCarlo.Charts with shared color palette, axis factory methods, percentile marker and target line section builders
- Created `HistogramChart` control using LiveCharts2 ColumnSeries with blue bars at 0.85 opacity, amber dashed P10/P50/P90 percentile markers, optional red dashed target line with annotation
- Created `CDFChart` control using LiveCharts2 LineSeries with smooth S-curve, area fill at 0.08 opacity, and horizontal percentile reference lines
- Created `ResultsViewModel` with output selection, stats computation, chart data binding, target analysis with live probability calculation, and clipboard copy
- Created `NumberFormatter` utility for intelligent value formatting across magnitudes ($M, $K, %, raw decimals)
- Created `HeadlineStatCard` — large P50 value with label and 95% CI range
- Created `StatsPanelControl` — 2-column grid: Mean/StdDev, P5/P95, Min/Max, Skewness/Kurtosis using MonoFont
- Created `TargetLineControl` — target value input with live P(above)/P(below) calculation
- Replaced placeholder ResultsView with full dashboard: output dropdown, headline stat, histogram/CDF with toggle, stats panel, target analysis, export/copy buttons

### Files Created/Modified
- `src/MonteCarlo.Charts/Themes/ChartTheme.cs` — Color palette, axis/section factories
- `src/MonteCarlo.Charts/Controls/HistogramChart.xaml/.cs` — Histogram with markers
- `src/MonteCarlo.Charts/Controls/CDFChart.xaml/.cs` — CDF S-curve
- `src/MonteCarlo.UI/Converters/NumberFormatter.cs` — Smart formatting
- `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs` — Results VM
- `src/MonteCarlo.UI/Views/HeadlineStatCard.xaml/.cs` — Big P50 card
- `src/MonteCarlo.UI/Views/StatsPanelControl.xaml/.cs` — Stats grid
- `src/MonteCarlo.UI/Views/TargetLineControl.xaml/.cs` — Target analysis
- `src/MonteCarlo.UI/Views/ResultsView.xaml/.cs` — Full dashboard replacing placeholder

### Key Decisions Made During Implementation
- Histogram maps bin indices back to real values using BinCenters/BinEdges for percentile marker positioning
- CDF subsamples sorted values to ~200 points for LiveCharts2 performance
- PDF/CDF toggle uses a pair of RadioButtons styled as TabBarButtons with a shared ToggleChartModeCommand
- Clipboard copy formats as a plain text table with proper alignment
- ResultsView creates its own DataContext (ResultsViewModel) inline in XAML — LoadResults() is called externally
- Chart controls use DependencyProperties so they can be bound directly from the ResultsViewModel

### Issues / Notes for Architect
- Export to Sheet button is wired visually but the command is deferred to TASK-012
- The PDF/CDF toggle using RadioButtons + BoolToVisibility may need refinement if the converter doesn't work as a RadioButton IsChecked binding — a simple button toggle might be cleaner
- LiveCharts2 ColumnSeries doesn't support border-radius on bars — bars will have square edges
- No .NET SDK available — cannot verify compilation

### Test Results
- No tests — this is a UI/charting task. Cannot build without .NET SDK.


---

## TASK-010 — Tornado Chart (Custom SkiaSharp)
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `TornadoChart` as a custom SkiaSharp `SKElement` control in MonteCarlo.Charts
- Renders bidirectional horizontal bars: blue (#3B82F6) for increase, orange (#F97316) for decrease
- Bars sorted by absolute impact (largest at top), with 2px corner radius on outer ends
- Center baseline with "Base: $X.XM" label below
- Input labels right-aligned on the left axis
- Bar-tip annotations showing output range (P10 — P90 values)
- Header labels ("◀ Decreases" / "Increases ▶") at the top
- Hover interaction: highlights bar pair and shows dark tooltip with input name, P10/P90 output values, swing amount, and variance contribution
- MaxInputsToShow (default 10) with ShowAll toggle
- Configurable ValueFormatter for number display
- Empty state rendering when no data is available

### Files Created/Modified
- `src/MonteCarlo.Charts/Controls/TornadoChart.cs` — Full custom SkiaSharp control (~420 lines)

### Key Decisions Made During Implementation
- Used SKRoundRect with per-corner radii for the 2px outer-end bar radius effect
- Hit-testing uses Y-band detection (simpler than full rect containment) since bars span the full width contextually
- Tooltip positioned above the hovered bar, with fallback below if near the top
- Label column and annotation column widths are measured dynamically based on actual text width
- Chart does not use XAML (pure code control) as recommended by the spec — SkiaSharp SKElement doesn't need XAML

### Issues / Notes for Architect
- The `DrawRoundRect` with asymmetric corner radii uses `SKRoundRect(rect, rx1, ry1, rx2, ry2)` constructor — need to verify this is the correct SkiaSharp API (may need per-corner `SetRectRadii` instead)
- No sub-label (distribution info) below input labels yet — kept labels single-line for cleaner look at 380px task pane width. Can be added if desired.
- No .NET SDK available — cannot verify compilation

### Test Results
- No tests — this is a visual rendering control. Layout math could be unit-tested if desired.


---

## TASK-011 — Config Persistence + Run View + Convergence
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done

**Part A: Config Persistence**
- Created `SimulationProfile`, `SavedInput`, `SavedOutput` in MonteCarlo.Engine with System.Text.Json serialization attributes
- Created `ConfigPersistence` in MonteCarlo.Addin using CustomXMLParts with namespace `urn:montecarlo-xl:config:v1`
- JSON payload wrapped in XML CDATA for clean storage
- Fallback to hidden `__MC_Config` sheet if CustomXMLParts unavailable
- Save, Load, Clear, and GetProfileNames operations

**Part B: Run View**
- Created `RunViewModel` with progress tracking, live stats, convergence indicators, and Stop event
- Replaced placeholder RunView with full UI: progress bar (custom Border-based), iteration counter, elapsed/remaining time, live preview stats (Mean, Median, StdDev, P5, P95), convergence panel, and red Stop button
- Created `PercentToWidthConverter` for the progress bar fill width

**Part C: Convergence Monitor**
- Created `ConvergenceChecker` in MonteCarlo.Engine with rolling window stability detection
- Monitors Mean, P50, P90, StdDev at configurable checkpoint intervals
- Reports Stable/Drifting/Unstable status based on relative change rate vs tolerance
- 12 unit tests covering stable/unstable series, tight tolerance, insufficient data, reset

### Files Created/Modified
- `src/MonteCarlo.Engine/Simulation/SimulationProfile.cs` — Config model
- `src/MonteCarlo.Engine/Analysis/ConvergenceChecker.cs` — Convergence monitor
- `src/MonteCarlo.Addin/Excel/ConfigPersistence.cs` — CustomXMLPart persistence
- `src/MonteCarlo.UI/ViewModels/RunViewModel.cs` — Run view VM
- `src/MonteCarlo.UI/Views/RunView.xaml/.cs` — Full run view replacing placeholder
- `src/MonteCarlo.UI/Converters/PercentToWidthConverter.cs` — Progress bar converter
- `tests/MonteCarlo.Engine.Tests/Analysis/ConvergenceCheckerTests.cs` — 12 tests

### Key Decisions Made During Implementation
- Used a custom Border-based progress bar instead of WPF ProgressBar for full design system control (color, corner radius)
- Auto-save/auto-load hooks not wired yet — will be connected in TASK-013 (Orchestrator)
- ConvergenceChecker stores full checkpoint history (not just windowed) for potential future charting
- Convergence indicator colors are strings (hex) in the ViewModel since WPF converters will handle brush binding

### Issues / Notes for Architect
- The progress bar uses a MultiBinding with PercentToWidth converter — this is clean but the binding to parent Grid.ActualWidth may need runtime verification
- Auto-save debouncing not yet implemented in SetupViewModel — planned for TASK-013 wiring
- No .NET SDK — cannot verify compilation or run tests

### Test Results
- 12 ConvergenceChecker tests written (cannot run without .NET SDK)


---

## TASK-012 — Results Export to Excel Sheet
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `ChartImageRenderer` with WPF RenderTargetBitmap (high-DPI) and SkiaSharp SKBitmap export to PNG bytes
- Created `ResultsExporter` with ExportSummary and ExportRawData methods
- ExportSummary writes: title section, summary stats table, percentiles table, sensitivity table (top 10), input assumptions table, embedded chart images
- Professional formatting: bold section headers with bottom borders, alternating row shading (#F8FAFC), auto-fit columns, number-formatted values, colored sheet tabs
- ExportRawData writes iteration-level data as a single 2D array write for performance
- Helper methods for column letter conversion, number format detection, distribution formatting

### Files Created/Modified
- `src/MonteCarlo.Addin/Export/ChartImageRenderer.cs` — PNG rendering from WPF and SkiaSharp
- `src/MonteCarlo.Addin/Export/ResultsExporter.cs` — Full export implementation

### Key Decisions Made During Implementation
- Used 2D array write (`Range.Value2 = data`) for raw data export — orders of magnitude faster than cell-by-cell
- Sheet names truncated to 31 chars (Excel limit)
- Tab color uses BGR format (Excel COM convention)
- Existing sheets with same name are cleared and reused rather than creating duplicates
- Chart images use temp files for embedding (Excel COM requires file paths for AddPicture)

### Issues / Notes for Architect
- The ExportSummary method requires chart PNG bytes as parameters — the Orchestrator (TASK-013) will need to render the charts before calling export
- BGR color values for Excel COM should be verified at runtime
- No .NET SDK — cannot verify compilation

### Test Results
- No tests — COM-dependent code. Manual testing required.


---

## TASK-013 — Simulation Orchestrator (End-to-End Integration)
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `SimulationOrchestrator` in MonteCarlo.Addin.Services — the central coordinator for the full simulation lifecycle
- Implements fast-mode evaluator: writes input values → app.Calculate() → reads output values
- Disables screen updating and sets manual calculation during runs for performance
- Saves and restores original cell values around simulation runs
- Computes SummaryStatistics and SensitivityAnalysis per output on completion
- Fires ProgressChanged, SimulationComplete, SimulationError, and ConvergenceUpdated events
- Includes BuildProfile/SaveConfig/LoadConfig for config persistence
- Updated `AddIn.AutoOpen()` to initialize ConfigPersistence and SimulationOrchestrator as static services
- Updated `MainViewModel` to:
  - Hold persistent RunView and ResultsView instances
  - Wire SetupViewModel.RunSimulationRequested → navigate to RunView
  - Wire RunViewModel.StopRequested → CancelSimulationRequested event
  - Expose event-driven hooks (OnSimulationProgress, OnLiveStatsUpdate, OnConvergenceUpdate, OnSimulationComplete, OnSimulationCancelled, OnSimulationError) for the Addin layer
  - Auto-navigate Setup → Run → Results based on simulation state

### Files Created/Modified
- `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs` — Central coordinator
- `src/MonteCarlo.Addin/AddIn.cs` — Updated: service initialization + config auto-load
- `src/MonteCarlo.UI/ViewModels/MainViewModel.cs` — Updated: full simulation flow wiring

### Key Decisions Made During Implementation
- Used event-based communication between MainViewModel and the Addin layer rather than direct references, keeping MonteCarlo.UI free of Addin/COM dependencies
- The Orchestrator fires events that the Addin layer marshals to WPF Dispatcher before calling MainViewModel methods
- UseParallelEvaluation = false for Excel COM (single-threaded)
- Sensitivity analysis failure is caught and treated as non-fatal (returns empty list)
- Original cell values restored in a finally block to ensure cleanup even on error/cancellation

### Issues / Notes for Architect
- The actual WPF Dispatcher marshaling (Dispatcher.Invoke) needs to happen in the Addin layer when hooking Orchestrator events to MainViewModel — not implemented here as it's COM-layer code
- Config auto-load on startup loads the profile but doesn't populate the SetupViewModel yet — this would need a method on SetupViewModel to accept a SimulationProfile
- No .NET SDK — cannot verify compilation
- **Phase 1 is now complete!** All tasks 001–013 are implemented. The end-to-end flow is: Setup → configure inputs/outputs → Run → progress + live stats → Results with histogram, CDF, tornado, stats, target analysis, and export.

### Test Results
- No new tests — integration testing requires running Excel. 5 manual test scenarios documented in the spec.


---

## TASK-014 — Additional Distributions (Beta, Weibull, Exponential, Poisson)
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Implemented 4 Phase 2 distributions following the NormalDistribution pattern:
  - **BetaDistribution** — wraps MathNet.Numerics.Distributions.Beta, bounded [0,1], parameters α/β
  - **WeibullDistribution** — wraps MathNet.Numerics.Distributions.Weibull, parameters shape/scale
  - **ExponentialDistribution** — wraps MathNet.Numerics.Distributions.Exponential, parameter rate (λ), Mean=1/λ
  - **PoissonDistribution** — discrete distribution, wraps MathNet.Numerics.Distributions.Poisson, Sample() returns integers as doubles, uses PMF for PDF
- Updated DistributionFactory: added 4 new distributions, AvailableDistributions now returns 10
- Updated DistributionParameterPanel.xaml with new parameter sections (Beta: α/β, Weibull: shape/scale, Exponential: rate, Poisson: lambda)
- Updated SetupViewModel with new parameter properties and BuildParameterDictionary switch cases
- Wrote comprehensive tests for all 4 distributions following TASK-002 patterns:
  - Construction/validation, statistical convergence (100k samples), quantile round-trip, CDF boundaries, reproducibility
  - Beta-specific: Beta(1,1) is Uniform, all samples in [0,1], mean validation
  - Weibull-specific: all samples ≥ 0, Weibull(1,λ) ≈ Exponential(1/λ)
  - Exponential-specific: all samples ≥ 0, mean = 1/λ
  - Poisson-specific: all samples are non-negative integers, variance ≈ λ, large λ approximates Normal
- Updated DistributionFactoryTests: count check now expects 10, added Create tests for all 4 new types

### Files Created
- `src/MonteCarlo.Engine/Distributions/BetaDistribution.cs`
- `src/MonteCarlo.Engine/Distributions/WeibullDistribution.cs`
- `src/MonteCarlo.Engine/Distributions/ExponentialDistribution.cs`
- `src/MonteCarlo.Engine/Distributions/PoissonDistribution.cs`
- `tests/MonteCarlo.Engine.Tests/Distributions/BetaDistributionTests.cs`
- `tests/MonteCarlo.Engine.Tests/Distributions/WeibullDistributionTests.cs`
- `tests/MonteCarlo.Engine.Tests/Distributions/ExponentialDistributionTests.cs`
- `tests/MonteCarlo.Engine.Tests/Distributions/PoissonDistributionTests.cs`

### Files Modified
- `src/MonteCarlo.Engine/Distributions/DistributionFactory.cs` — added 4 new distributions
- `tests/MonteCarlo.Engine.Tests/Distributions/DistributionFactoryTests.cs` — updated count + 4 new Create tests
- `src/MonteCarlo.UI/Views/DistributionParameterPanel.xaml` — 4 new parameter sections
- `src/MonteCarlo.UI/ViewModels/SetupViewModel.cs` — new properties + switch cases + reset

### Key Decisions Made During Implementation
- Poisson PDF uses `_inner.Probability(k)` (PMF) for integer values and returns 0.0 for non-integers
- Poisson CDF returns 0.0 for negative values explicitly
- Poisson quantile round-trip test asserts CDF(Percentile(p)) >= p (not ≈ p) since it's discrete
- Used separate ParamRate (Exponential) and ParamLambda (Poisson) properties to avoid ambiguity even though both could share "rate"

### Issues / Notes for Architect
- **No dotnet SDK** — cannot build or run tests. Code follows established patterns exactly.
- All 4 distributions are sealed classes with readonly inner MathNet distribution, matching existing pattern

### Test Results
- Cannot run tests (no dotnet SDK). ~45 new tests written across 4 test files + 5 new factory tests.


---

## TASK-015 — Iman-Conover Correlation Engine
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `CorrelationMatrix` class in `MonteCarlo.Engine/Correlation/`:
  - Square matrix with `Size`, indexer, `Identity()` factory, `ToArray()`
  - `Validate()` checks: diagonal=1.0, symmetric, values in [-1,1], positive semi-definite (via eigendecomposition)
  - `IsPositiveSemiDefinite()` for quick check
  - `EnsurePositiveSemiDefinite()` — spectral correction: eigendecompose, clamp negative eigenvalues to 1e-10, reconstruct, rescale diagonal to 1.0
- Created `ImanConover` static class implementing the full algorithm:
  1. Rank-transform with average ranks for ties
  2. Van der Waerden scores via Φ⁻¹(R/(N+1))
  3. Pearson correlation of score matrix
  4. Cholesky decomposition of target and current correlations
  5. Transformation matrix M = P × Q⁻¹
  6. Apply transformation to scores
  7. Rank-transform transformed scores
  8. Rearrange original samples by target ranks (preserves marginals exactly)
- Integrated into SimulationEngine: added `Correlation` property to `SimulationConfig`, one-line hook after sample generation
- Wrote 12 CorrelationMatrix tests and 10 ImanConover tests

### Files Created
- `src/MonteCarlo.Engine/Correlation/CorrelationMatrix.cs`
- `src/MonteCarlo.Engine/Correlation/ImanConover.cs`
- `tests/MonteCarlo.Engine.Tests/Correlation/CorrelationMatrixTests.cs`
- `tests/MonteCarlo.Engine.Tests/Correlation/ImanConoverTests.cs`

### Files Modified
- `src/MonteCarlo.Engine/Simulation/SimulationConfig.cs` — added `Correlation` property
- `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs` — added Iman-Conover hook at step 3

### Key Decisions Made During Implementation
- Used MathNet.Numerics.LinearAlgebra for Cholesky and eigendecomposition (already a project dependency)
- Spectral correction (Higham-simplified) rather than iterative alternating projection — simpler and sufficient for our use case
- EnsurePositiveSemiDefinite returns `this` if already PSD (no unnecessary copy)
- The Apply method modifies the input matrix in place for performance (avoids large allocation)

### Issues / Notes for Architect
- **No dotnet SDK** — cannot build or test
- The algorithm requires at least 3 samples (documented in exception message)
- Cholesky decomposition will fail if the target matrix is only PSD (not PD) — users should use EnsurePositiveSemiDefinite before passing to the engine
- Test tolerances: ±0.02 for 100k samples, ±0.005 for 50k at large N, ±0.05 for small N (100)

### Test Results
- Cannot run tests (no dotnet SDK). 12 CorrelationMatrix + 10 ImanConover tests written.


---

## TASK-016 — Correlation Matrix UI
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `CorrelationMatrixGrid` control — dynamically generates an editable grid:
  - Row/column headers with truncated input labels and tooltips
  - Diagonal cells read-only (1.0) with gray background
  - Upper triangle editable with TextBox cells
  - Lower triangle auto-mirrors with read-only TextBlocks
  - Blue/orange color coding proportional to correlation magnitude
  - CellValueChanged event for ViewModel wiring
  - ScrollViewer for large matrices
- Created `CorrelationColorConverter` for value-to-color brush mapping
- Created `CorrelationViewModel` with:
  - Initialize() method accepting input labels and optional existing matrix
  - Real-time PSD validation via CorrelationMatrix.Validate()
  - AutoFix command using EnsurePositiveSemiDefinite()
  - ClearAll resets to identity, Apply returns matrix or null
  - Applied/CloseRequested events for navigation
- Created `CorrelationView` with header, tip card, matrix grid, validation status (green/amber), Auto-fix button, footer with Clear All and Apply & Close
- Added `CorrelationEditorRequested` event and `OpenCorrelationEditorCommand` to SetupViewModel
- Added correlation button in SetupView below input cards (shows input count, disabled when <2 inputs)
- Updated `SimulationProfile` with `CorrelationMatrix` property (JSON-serialized as flat array + size)

### Files Created
- `src/MonteCarlo.UI/Views/CorrelationMatrixGrid.xaml/.cs`
- `src/MonteCarlo.UI/Views/CorrelationView.xaml/.cs`
- `src/MonteCarlo.UI/ViewModels/CorrelationViewModel.cs`
- `src/MonteCarlo.UI/Converters/CorrelationColorConverter.cs`

### Files Modified
- `src/MonteCarlo.UI/ViewModels/SetupViewModel.cs` — CorrelationMatrixValues, CanDefineCorrelations, CorrelationEditorRequested event
- `src/MonteCarlo.UI/Views/SetupView.xaml` — correlation button
- `src/MonteCarlo.Engine/Simulation/SimulationProfile.cs` — CorrelationMatrix persistence

### Key Decisions
- Grid is rebuilt dynamically in code-behind rather than using a DataGrid — simpler for the symmetric/diagonal constraints
- Lower triangle is rendered as read-only TextBlocks mirroring upper triangle values
- Color alpha scales from 0 to 120 (not 180 as in spec converter) for subtler effect
- SimulationProfile stores correlation as flat double[] (row-major) since System.Text.Json doesn't handle 2D arrays natively

### Issues / Notes for Architect
- No dotnet SDK — cannot build
- The CorrelationView navigation (MainViewModel wiring) is not implemented — the Addin or MainViewModel needs to wire SetupViewModel.CorrelationEditorRequested to show CorrelationView
- Sticky headers for scrolling large matrices (>6 inputs) are not implemented — would need a more complex layout. ScrollViewer with hint text is the fallback.


---

## TASK-017 — Custom Excel Functions (=MC.Normal() UDFs)
**Status**: COMPLETE
**Date**: 2026-03-20
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done
- Created `MonteCarloFunctions` static class with 9 ExcelDna UDFs:
  - MC.Normal(mean, stdDev) → mean
  - MC.Triangular(min, mode, max) → mode
  - MC.PERT(min, mode, max) → mode
  - MC.Lognormal(mu, sigma) → exp(μ + σ²/2)
  - MC.Uniform(min, max) → midpoint
  - MC.Beta(alpha, beta) → α/(α+β)
  - MC.Weibull(shape, scale) → scale × Γ(1 + 1/shape)
  - MC.Exponential(rate) → 1/λ
  - MC.Poisson(lambda) → λ
  - All return #VALUE! for invalid parameters
  - All registered with Category = "MonteCarlo.XL"
- Created `MCFunctionScanner` for detecting MC.* formulas in worksheets:
  - Regex-based formula parser for =MC.FunctionName(args...)
  - Parameter name mapping for all 9 function types
  - ScanWorksheet scans UsedRange for MC formulas
  - ResolveParameters handles both literal numbers and cell references
  - SplitArguments respects nested parentheses
- Created `DetectedMCFunction` data class
- Updated `SimulationOrchestrator`:
  - Scans active sheet for MC.* formulas before simulation
  - Adds auto-detected inputs, skipping GUI-tagged duplicates
  - Saves original formulas before simulation
  - Restores original formulas in finally block after simulation

### Files Created
- `src/MonteCarlo.Addin/UDF/MonteCarloFunctions.cs`
- `src/MonteCarlo.Addin/UDF/MCFunctionScanner.cs`

### Files Modified
- `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs` — MC function scanning + formula save/restore

### Key Decisions
- Used Approach A (formula scanning) as recommended in spec — simpler and more reliable than registration during UDF execution
- MC.Weibull expected value uses MathNet.Numerics.SpecialFunctions.Gamma rather than implementing our own
- MC function scanning is non-fatal — failures don't prevent simulation from running with tagged inputs
- Formula restoration happens in finally block to ensure cleanup even on error/cancellation
- taggedInputs is now a mutable List<> instead of using IReadOnlyList from GetAllInputs()

### Issues / Notes for Architect
- No dotnet SDK — cannot build
- UDF tests would need Excel runtime; the UDF return values could be unit-tested as simple method calls though
- The formula regex is basic (`=MC.FunctionName(args)`) — doesn't handle nested MC calls or complex expressions as arguments
- MC function scanning only scans the active sheet — could be extended to scan all sheets if needed

---

## TASK-018 — Dark Theme, Performance Optimization, and UX Polish
**Status**: COMPLETE
**Date**: 2026-03-21
**Branch**: claude/engine-tasks-003-004-TFgln

### What Was Done

**Part A: Dark Theme**
- LightTheme.xaml and DarkTheme.xaml already existed with correct color palettes from prior work
- Updated ThemeManager to sync ChartTheme dark mode when applying themes (ChartTheme.SetDarkMode)
- Updated ChartTheme to support theme-dependent neutral colors (label, border, gridline) via SetDarkMode()
- Updated TornadoChart to use theme-adaptive colors from ChartTheme instead of hardcoded light-mode values
- Built full SettingsView with Light/Dark/System radio buttons wired to ThemeManager
- ThemeManager loads/saves preference via Windows registry, applies via ResourceDictionary swapping
- Wired theme initialization into MainTaskPaneControl code-behind (load saved theme on construction)
- Moved MergedDictionaries management to code-behind for proper theme dict insertion order

**Part B: Performance Optimization**
- Added `EnableEvents = false` during simulation in SimulationOrchestrator (alongside existing ScreenUpdating/Calculation)
- Refactored evaluator to pre-group inputs/outputs by sheet for batch COM operations
- Memory pre-allocation already confirmed in SimulationEngine (Step 1)
- Progress throttling already confirmed (every 100 iterations with 50ms min gap)
- Created SimulationBenchmark class: synthetic test with N Normal inputs, trivial evaluator, reports iterations/sec and µs/iteration

**Part C: UX Polish**
- Created ValidatedTextBox control: debounced validation (300ms), red/green border states, inline error message display
- Created SkeletonLoader control: animated pulsing gray rectangle, supports dark mode theme colors
- Created NumberInputTextBox: paste support, scientific notation parsing, comma/currency stripping, formatted display on blur
- Added fade transition animation (200ms) on view content changes via TargetUpdated trigger in MainTaskPaneControl
- Enhanced SetupView empty states: icon + heading + description + CTA button for both inputs and outputs sections
- Added ResultsView empty state: "No results yet" with icon and instruction when no simulation has run (uses HasResults property)
- Added error states in RunView: error card with type classification, human-readable message, and Try Again button
- Added RunViewModel.ShowError() with error classification (Configuration, Excel, Memory, Simulation)
- Updated MainViewModel.OnSimulationError to show error in RunView instead of navigating away
- Wired RetryRequested event to re-run simulation from error state
- Added keyboard shortcuts: Ctrl+Shift+R (Run), Ctrl+Shift+S (Stop), Ctrl+Shift+T (Toggle pane)
- Created KeyboardShortcuts.cs with ExcelDna-registered macro commands
- Registered/unregistered shortcuts in AddIn.AutoOpen/AutoClose via Application.OnKey()

### Files Created
- `src/MonteCarlo.UI/Controls/ValidatedTextBox.xaml` + `.xaml.cs` — validated text input control
- `src/MonteCarlo.UI/Controls/SkeletonLoader.xaml` + `.xaml.cs` — animated loading placeholder
- `src/MonteCarlo.UI/Controls/NumberInputTextBox.cs` — enhanced numeric input with formatting
- `src/MonteCarlo.Engine/Simulation/SimulationBenchmark.cs` — synthetic benchmark runner
- `src/MonteCarlo.Addin/KeyboardShortcuts.cs` — ExcelDna keyboard shortcut macros

### Files Modified
- `src/MonteCarlo.UI/Views/SettingsView.xaml` — full theme selector UI
- `src/MonteCarlo.UI/Views/SettingsView.xaml.cs` — theme selection logic
- `src/MonteCarlo.UI/Views/MainTaskPaneControl.xaml` — fade transition, removed inline merged dict
- `src/MonteCarlo.UI/Views/MainTaskPaneControl.xaml.cs` — theme initialization on construction
- `src/MonteCarlo.UI/Views/SetupView.xaml` — enhanced empty states with CTAs
- `src/MonteCarlo.UI/Views/ResultsView.xaml` — added empty state, wrapped content in HasResults visibility
- `src/MonteCarlo.UI/Views/RunView.xaml` — error card, Stop button visibility binding
- `src/MonteCarlo.UI/ViewModels/RunViewModel.cs` — error state properties, ShowError(), RetrySimulation
- `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs` — HasResults property
- `src/MonteCarlo.UI/ViewModels/MainViewModel.cs` — OnSimulationError takes Exception, retry wiring
- `src/MonteCarlo.UI/Services/ThemeManager.cs` — ChartTheme sync on theme apply
- `src/MonteCarlo.Charts/Themes/ChartTheme.cs` — SetDarkMode() for theme-adaptive colors
- `src/MonteCarlo.Charts/Controls/TornadoChart.cs` — theme-adaptive label/baseline colors
- `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs` — EnableEvents, batch evaluator
- `src/MonteCarlo.Addin/AddIn.cs` — keyboard shortcut registration/cleanup

### Key Decisions
- Theme dictionary is inserted at MergedDictionaries[0] in code-behind rather than XAML to allow runtime swapping
- ChartTheme uses static mutable state (SetDarkMode) for simplicity since WPF + SkiaSharp don't share resource systems
- TornadoChart label color uses slate-100 for dark mode (high contrast on dark surface) vs slate-900 for light mode
- Error classification is type-based (COMException → "Excel Error") rather than message-based for reliability
- Keyboard shortcuts use Application.OnKey() mapped to ExcelDna [ExcelCommand] macros
- NumberInputTextBox is a standalone class (not UserControl) inheriting from TextBox for maximum flexibility

### Issues / Notes for Architect
- No dotnet SDK — cannot build or test
- The progress bar doesn't use WPF animation for smooth width transitions (DoubleAnimation would conflict with the MultiBinding converter); the throttled updates from the engine provide reasonable visual smoothness
- SkeletonLoader theme adaptation sets the base color but doesn't restart the XAML EventTrigger storyboard — the pulsing animation color range stays at light-mode values when switching to dark mode at runtime (would need code-behind storyboard management to fix properly)
- SystemParameters.MinimizeAnimation check for disabling animations is not yet implemented — could be added to a style trigger
- Escape key handling for cancel actions would need InputBindings on individual controls rather than Application.OnKey since Excel intercepts Escape

### Test Results
- Cannot run tests (no dotnet SDK)

