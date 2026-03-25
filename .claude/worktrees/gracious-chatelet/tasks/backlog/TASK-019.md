# TASK-019: Comprehensive Code Review & Bug Sweep

## Context

Read `ROADMAP.md` for full project context. MonteCarlo.XL is an Excel add-in built with ExcelDna + .NET 8 + WPF. The codebase was written without a compiler (no dotnet SDK in the authoring environment), so it has never been fully verified. A first pass of build fixes has been done, but the first runtime test revealed XAML resource loading failures and likely more latent bugs.

**This is the most important task in the project.** The codebase has ~14,500 lines across 111 source files. Every file needs scrutiny.

## Build Environment

You MUST use this to build:
```bash
dotnet build -p:EnableWindowsTargeting=true
```
The `EnableWindowsTargeting=true` flag is required because this is a Windows-targeting project (WPF/WinForms) and you are running on Linux.

For tests:
```bash
dotnet test tests/MonteCarlo.Engine.Tests/
```

If you get stale BAML errors, clean first:
```bash
dotnet clean -p:EnableWindowsTargeting=true
dotnet build -p:EnableWindowsTargeting=true --no-incremental
```

## Objective

Do a methodical, file-by-file review of the entire codebase. Fix every bug you find. The goal is **zero build errors** and **all tests passing**, plus elimination of every runtime bug that can be identified through static analysis.

**Do NOT skip files. Do NOT skim. Read every file.**

---

## Part A: Build Verification (do this FIRST)

1. Run `dotnet build -p:EnableWindowsTargeting=true` and fix every error and warning (except the single `CS0067` warning about `ConvergenceUpdated` — leave that).
2. Run `dotnet test tests/MonteCarlo.Engine.Tests/` and fix any failures.
3. Iterate until perfectly clean.

---

## Part B: XAML Resource Loading (CRITICAL)

**Background:** This add-in runs inside Excel via ExcelDna. There is NO `App.xaml` — no WPF Application startup. WPF controls are hosted in a WinForms `ElementHost` inside an Excel Custom Task Pane. This means:

- There is no global Application.Resources dictionary
- `StaticResource` lookups fail if the resource isn't in the control's own `Resources` or a parent's `MergedDictionaries`
- Views that are instantiated in code (not added to a visual tree) cannot inherit resources from parents

**What to check:**

1. **Every `.xaml` file that uses `StaticResource`** must either:
   - Be a child of a view that already has the resources merged, OR
   - Have its own `MergedDictionaries` with LightTheme.xaml + GlobalStyles.xaml

2. **Views instantiated standalone in code** (check ViewModels and code-behind for `new SomeView()` patterns) MUST have their own MergedDictionaries. Currently these views are created standalone:
   - `MainTaskPaneControl` (created by `TaskPaneHost`)
   - `SetupView` (created by `MainViewModel` field initializer)
   - `RunView` (created by `MainViewModel` field initializer)
   - `ResultsView` (created by `MainViewModel` field initializer)
   - `SettingsView` (created by `MainViewModel.NavigateToSettings()`)
   - Check for ANY other `new XxxView()` or `new XxxControl()` patterns

3. **Sub-controls inside these views**: Controls like `InputCardControl`, `InputEditorControl`, `OutputCardControl`, `HeadlineStatCard`, `StatsPanelControl`, `DistributionParameterPanel`, `DistributionPreviewControl`, `CorrelationView`, `CorrelationMatrixGrid`, `TargetLineControl`, `SkeletonLoader`, `ValidatedTextBox` — these are typically instantiated via XAML inside a parent view, so they inherit resources. BUT verify this is actually true. If any of them are created in code-behind or a ViewModel, they need their own MergedDictionaries.

4. **The MergedDictionaries pattern** for standalone views:
```xml
<UserControl.Resources>
    <ResourceDictionary>
        <ResourceDictionary.MergedDictionaries>
            <ResourceDictionary Source="/MonteCarlo.UI;component/Styles/LightTheme.xaml"/>
            <ResourceDictionary Source="/MonteCarlo.UI;component/Styles/GlobalStyles.xaml"/>
        </ResourceDictionary.MergedDictionaries>
        <!-- any local resources like converters go here -->
    </ResourceDictionary>
</UserControl.Resources>
```

5. **`sys:` namespace in GlobalStyles.xaml**: The file uses `sys:Double` with namespace `xmlns:sys="clr-namespace:System;assembly=mscorlib"`. On .NET 8, the correct assembly is `System.Runtime`, not `mscorlib`. Verify this is correct. If it says `mscorlib`, change it to `System.Runtime`.

6. **DynamicResource vs StaticResource**: Theme-sensitive brushes (BackgroundBrush, SurfaceBrush, TextPrimaryBrush, etc.) from LightTheme/DarkTheme MUST use `DynamicResource`. Data color brushes (Blue500Brush etc.) and styles (HeadlineSmall, PrimaryButton, etc.) from GlobalStyles can use `StaticResource` — but only if the dictionaries are in scope.

---

## Part C: Namespace & Type Conflicts

The Addin project mixes Excel Interop, WinForms, and WPF. Check for:

1. **Ambiguous types**: `Application`, `Range`, `File`, `Size`, `Color`, `Font`, etc. A `GlobalUsings.cs` exists to alias some of these — verify it covers all cases.
2. **Every `.cs` file in `src/MonteCarlo.Addin/`** — read each one and check for unresolved or ambiguous type references.
3. **Cross-project references**: Verify that `MonteCarlo.UI` does NOT reference `MonteCarlo.Addin` or Excel Interop. Verify `MonteCarlo.Engine` does NOT reference WPF, WinForms, or Excel. These dependency boundaries are critical.

---

## Part D: API Mismatches

CC (the authoring agent) frequently hallucinated API signatures. Check every external API call:

1. **Math.NET Numerics**: Every call to `InverseCumulativeDistribution`, `CumulativeDistribution`, `Density`, `Sample`, etc. Verify the method actually exists on the specific distribution type being used. Math.NET's `Poisson` (discrete) doesn't have `InverseCumulativeDistribution`. Neither does `Weibull` in some versions.

2. **SkiaSharp**: Check all `SKPaint`, `SKCanvas`, `SKRoundRect`, `SKRect`, `SKTypeface` usage. Verify constructor signatures, especially `SKRoundRect` (does not support per-corner radii).

3. **LiveCharts2**: Verify `LineSeries`, `ColumnSeries`, `Axis`, etc. API calls match the `2.0.0-rc3.3` version.

4. **ExcelDna**: Verify `CustomTaskPaneFactory.CreateCustomTaskPane`, `ExcelDnaUtil.Application`, `ExcelFunction` attribute usage, `IExcelAddIn` lifecycle.

5. **CommunityToolkit.Mvvm**: Verify `[ObservableProperty]`, `[RelayCommand]`, `ObservableObject` base class usage matches v8.3.2.

6. **SimulationResult API**: The methods `GetOutputValues(string outputId)` and `GetInputSamples(string inputId)` take string IDs, not int indices. Search the entire codebase for calls to these methods and verify they pass strings.

---

## Part E: ViewModel & Binding Review

1. **Every `[ObservableProperty]` field** — verify naming convention. CommunityToolkit generates a PascalCase property from a camelCase field. E.g., `private string _title` generates `Title`. Verify XAML bindings match generated property names.

2. **Every `[RelayCommand]` method** — verify XAML `Command="{Binding XxxCommand}"` matches. The generated command name is `MethodName` + `Command`. E.g., `void RunSimulation()` generates `RunSimulationCommand`.

3. **Every `DataContext` assignment** — verify ViewModels are instantiated correctly and have parameterless constructors (since they're created via XAML `<vm:SetupViewModel/>`).

4. **Converter classes** — verify every converter referenced in XAML (`x:Key="BoolVis"` etc.) actually exists and implements `IValueConverter` correctly.

---

## Part F: Event Wiring & Orchestration

Review the simulation execution pipeline:

1. `MonteCarloRibbon.OnRunSimulation` → `AddIn.TaskPane.Show()` — but this doesn't actually start a simulation. Verify the full pipeline from button click → orchestrator → engine → results.
2. `SimulationOrchestrator` — verify it correctly:
   - Reads input values from Excel
   - Builds `SimulationConfig`
   - Calls `SimulationEngine.RunAsync()`
   - Fires `SimulationProgress` events
   - Fires `SimulationComplete` with results
   - Handles cancellation
3. `MainViewModel` — verify it subscribes to orchestrator events and routes results to `ResultsViewModel`.
4. Thread marshaling — all UI updates from the simulation (which runs on background threads) MUST be dispatched to the UI thread. Check for `Dispatcher.Invoke` or `Dispatcher.BeginInvoke` usage wherever events flow from Engine → UI.

---

## Part G: Edge Cases & Defensive Checks

1. **Null checks**: COM interop objects (`ExcelDnaUtil.Application`, `ActiveWorkbook`, `ActiveSheet`) can be null. Verify every COM access is null-checked.
2. **COM release**: Verify there are no COM object leaks. Excel Interop objects should be released or allowed to be GC'd promptly.
3. **Thread safety**: The simulation engine runs in parallel. Verify shared state is properly synchronized.
4. **Division by zero**: In statistics calculations (variance, skewness, kurtosis), verify edge cases with 0 or 1 data point.
5. **Empty collections**: Verify behavior when there are 0 inputs, 0 outputs, or 0 iterations.

---

## Deliverables

1. Fix every bug found — commit with clear messages explaining each fix.
2. After all fixes: `dotnet build -p:EnableWindowsTargeting=true` → 0 errors
3. After all fixes: `dotnet test tests/MonteCarlo.Engine.Tests/` → all pass
4. Create a summary file `tasks/completed/TASK-019.md` listing every bug found and fixed, organized by category (XAML, namespace, API mismatch, etc.)

## Commit Strategy

Make **granular commits** — one commit per logical fix or category. Do NOT lump everything into one giant commit. Example:
- `fix(xaml): add MergedDictionaries to CorrelationView and sub-controls`
- `fix(interop): resolve remaining ambiguous type references in Addin`
- `fix(api): correct LiveCharts2 series configuration`
- `fix(viewmodel): fix binding name mismatches in SetupView`

## Important Notes

- Do NOT add new features. This is a bug-fix-only task.
- Do NOT refactor working code for style preferences. Only change code that is actually broken.
- Do NOT delete test files or reduce test coverage.
- If you find a bug you're unsure how to fix, document it in the summary with `[UNRESOLVED]` tag.
- The project targets `net8.0` (Engine, Tests) and `net8.0-windows` (Charts, UI, Addin). Do not change target frameworks.
