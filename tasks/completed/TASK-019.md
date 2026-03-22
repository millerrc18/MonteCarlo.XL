# TASK-019: Comprehensive Code Review & Bug Sweep

## Status: COMPLETE
## Date: 2026-03-22

---

## Summary

Performed a methodical, file-by-file review of all 111 source files (~14,500 lines) in the MonteCarlo.XL codebase. Fixed all identified bugs. The codebase compiles with 0 errors and all 327 tests pass.

---

## Bugs Found & Fixed

### XAML Resource Loading

**1. CorrelationView.xaml — Missing MergedDictionaries**
- **File**: `src/MonteCarlo.UI/Views/CorrelationView.xaml`
- **Issue**: CorrelationView uses `StaticResource` references to styles defined in GlobalStyles.xaml (HeadlineSmall, GhostButton, PrimaryButton, SecondaryButton, CaptionText) but had no MergedDictionaries to bring those resources into scope. When instantiated standalone (not as a child of a view that already has the dictionaries), all StaticResource lookups would fail at runtime.
- **Fix**: Added MergedDictionaries with LightTheme.xaml and GlobalStyles.xaml to the UserControl.Resources section, matching the pattern used by all other standalone views.

### Thread Safety

**2. SimulationEngine.cs — Race condition on lastProgressReport**
- **File**: `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs` (line 88-110)
- **Issue**: In the `Parallel.For` execution path, the `lastProgressReport` variable (a `TimeSpan` struct) was read and written from multiple threads without synchronization. Multiple threads could simultaneously read the same stale value and all decide to fire progress events, or worse, cause torn reads of the 8-byte struct on 32-bit platforms.
- **Fix**: Replaced `TimeSpan lastProgressReport` with `long lastProgressReportTicks` and used `Interlocked.Read` / `Interlocked.Exchange` for all accesses, ensuring atomic thread-safe operations.

---

## Verified Correct (No Bugs Found)

### Build (Part A)
- Initial build: 0 errors, 1 expected warning (CS0067 ConvergenceUpdated)
- All 327 unit tests pass

### XAML Resources (Part B)
- `sys:Double` namespace in GlobalStyles.xaml correctly uses `System.Runtime` (not `mscorlib`)
- All 5 main standalone views (MainTaskPaneControl, SetupView, RunView, ResultsView, SettingsView) have proper MergedDictionaries
- TaskPaneHost.cs correctly creates a WPF Application with merged resource dictionaries for the ExcelDna hosting scenario
- All sub-controls (InputEditorControl, InputCardControl, OutputCardControl, HeadlineStatCard, StatsPanelControl, DistributionParameterPanel, DistributionPreviewControl, ValidatedTextBox, SkeletonLoader, CorrelationMatrixGrid) are only instantiated via XAML in parent views — they correctly inherit resources
- DynamicResource vs StaticResource usage is correct: theme-sensitive brushes use DynamicResource, invariant colors/styles use StaticResource

### Namespace & Type Conflicts (Part C)
- GlobalUsings.cs correctly aliases Application, Range, File for the Addin project
- No ambiguous type references found in any Addin source files
- Dependency boundaries verified clean:
  - MonteCarlo.Engine: only MathNet.Numerics (pure, no UI/Excel)
  - MonteCarlo.Charts: Engine + LiveCharts + SkiaSharp (no Excel)
  - MonteCarlo.UI: Engine + Charts + CommunityToolkit.Mvvm (no Excel)
  - MonteCarlo.Addin: Engine + UI + ExcelDna

### API Mismatches (Part D)
- **MathNet.Numerics**: All distribution API calls correct. PoissonDistribution uses manual CDF search (no InverseCumulativeDistribution). WeibullDistribution uses analytical percentile formula.
- **SkiaSharp**: SKRoundRect uses 2-param radius constructor (correct). All SKPaint/SKCanvas usage valid.
- **LiveCharts2**: ColumnSeries, LineSeries, Axis, RectangularSection, SolidColorPaint, DashEffect — all match rc3.3 API.
- **ExcelDna**: IExcelAddIn, CustomTaskPaneFactory, ExcelFunction, ExcelCommand — all correct.
- **CommunityToolkit.Mvvm**: ObservableObject, [ObservableProperty], [RelayCommand] — all match v8.3.2.
- **SimulationResult API**: GetOutputValues/GetInputSamples always called with string IDs (verified all call sites).

### ViewModel & Bindings (Part E)
- All [ObservableProperty] naming conventions correct (private _camelCase field → PascalCase generated property)
- All [RelayCommand] naming correct (MethodName → MethodNameCommand)
- All XAML DataContext assignments verified with parameterless constructors
- All 7 converter classes exist and implement IValueConverter/IMultiValueConverter correctly

### Event Wiring (Part F)
- Simulation pipeline from ribbon → task pane → MainViewModel is structurally present
- SimulationOrchestrator correctly saves/restores Excel state, handles cancellation

### Edge Cases (Part G)
- COM interop objects are null-checked where accessed
- TaskPaneController uses null-forgiving operator (!) after EnsureCreated() which guarantees non-null — acceptable pattern
- Parallel execution: SimulationEngine uses Interlocked for thread-safe counter
- SummaryStatistics handles edge cases (empty arrays, single values) with proper validation
- Division by zero protected in statistics calculations

---

## [UNRESOLVED] — Architectural Gaps (Not Bugs)

1. **Orchestrator not wired to UI events**: The Addin layer never subscribes to `MainViewModel.RunSimulationRequested` / `CancelSimulationRequested`. The ribbon's Run button only shows the task pane — it doesn't trigger a simulation. This is an incomplete feature integration, not a code bug.

2. **CorrelationEditorRequested event has no subscribers**: `SetupViewModel.CorrelationEditorRequested` is raised when the user clicks "Define Correlations" but nothing in the Addin layer subscribes to open a CorrelationView. The CorrelationView and CorrelationViewModel exist but are not connected.

3. **Ribbon Stop button is a TODO**: `MonteCarloRibbon.OnStopSimulation` has a `// TODO` comment and doesn't call the orchestrator.

These are architectural integration gaps that would require new wiring code, not bug fixes.

---

## Build Results

```
Build succeeded.
    1 Warning(s) [CS0067 — expected, per task spec]
    0 Error(s)
```

## Test Results

```
Passed!  - Failed: 0, Passed: 327, Skipped: 0, Total: 327
```
