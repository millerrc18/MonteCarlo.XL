# @RISK Parity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the 9 highest-impact feature gaps between MonteCarlo.XL and @RISK, organized from broken-basics through power-user features.

**Architecture:** The engine (`MonteCarlo.Engine`) is pure C# with no Excel dependency, tested independently. The add-in (`MonteCarlo.Addin`) bridges Excel COM and the engine. The UI (`MonteCarlo.UI` + `MonteCarlo.Charts`) is WPF hosted via ElementHost. Changes follow existing patterns: distributions implement `IDistribution` and register in `DistributionFactory`; charts use LiveCharts2 with SkiaSharp; viewmodels use CommunityToolkit.Mvvm.

**Tech Stack:** .NET 8, ExcelDna 1.8, LiveCharts2 (rc3.3), SkiaSharp 2.88.8, MathNet.Numerics, CommunityToolkit.Mvvm

**Testing:** PowerShell COM automation drives Excel for integration tests. Unit tests via `dotnet test` on the Engine project. Visual verification via PowerShell screenshots of the Excel window.

---

## File Map

### Tier 1: Fix What's Broken
| Task | Create | Modify |
|------|--------|--------|
| 1. Fix export | | `TaskPaneIntegration.cs`, `SimulationOrchestrator.cs`, `ResultsViewModel.cs` |
| 2. Live charts | | `SimulationOrchestrator.cs`, `RunViewModel.cs`, `MainViewModel.cs`, `RunView.xaml` |
| 3. Clipboard copy | | `ResultsViewModel.cs` |

### Tier 2: What @RISK Users Expect
| Task | Create | Modify |
|------|--------|--------|
| 4. Latin Hypercube | `LatinHypercubeSampler.cs` | `SimulationEngine.cs`, `SimulationConfig.cs` |
| 5. Convergence auto-stop | | `SimulationEngine.cs`, `SimulationConfig.cs`, `SimulationOrchestrator.cs`, `SettingsView.xaml` |
| 6. Target probability UI | | `ResultsView.xaml`, `ResultsViewModel.cs` |

### Tier 3: Polish and Depth
| Task | Create | Modify |
|------|--------|--------|
| 7. New distributions (5) | `GammaDistribution.cs`, `LogisticDistribution.cs`, `GEVDistribution.cs`, `BinomialDistribution.cs`, `GeometricDistribution.cs` | `DistributionFactory.cs`, `MonteCarloFunctions.cs` |
| 8. PDF overlay on histogram | | `HistogramChart.xaml.cs`, `ChartTheme.cs` |
| 9. Sensitivity scatter plot | `ScatterChart.xaml`, `ScatterChart.xaml.cs` | `ResultsView.xaml`, `ResultsViewModel.cs` |

---

## Task 1: Fix Export to Worksheet

The export button does nothing because (a) sensitivity data is always passed as `null`, and (b) exceptions in the export path are silently swallowed by the dispatcher exception handler.

**Files:**
- Modify: `src/MonteCarlo.Addin/TaskPane/TaskPaneIntegration.cs` (ExportCurrentResults method, ~line 458-520)
- Modify: `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs` (expose SensitivityResults)
- Modify: `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs` (store sensitivity in result)

- [ ] **Step 1: Add SensitivityResults storage to ResultsViewModel**

In `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs`, add a property to store sensitivity results that the orchestrator computed:

```csharp
// Add after the existing SimulationResult property
public Dictionary<string, IReadOnlyList<SensitivityResult>>? SensitivityResults { get; private set; }
```

Update `LoadResults()` to accept and store sensitivity:

```csharp
public void LoadResults(SimulationResult result, 
    Dictionary<string, IReadOnlyList<SensitivityResult>>? sensitivity = null)
{
    SensitivityResults = sensitivity;
    // ... rest of existing LoadResults code unchanged
}
```

- [ ] **Step 2: Pass sensitivity through the completion chain**

In `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs`, the `SimulationCompleteEventArgs` already contains sensitivity data. Verify that `OnSimulationComplete` in `TaskPaneIntegration.cs` forwards it.

In `src/MonteCarlo.Addin/TaskPane/TaskPaneIntegration.cs`, update `OnSimulationComplete`:

```csharp
private void OnSimulationComplete(object? sender, SimulationCompleteEventArgs e)
{
    Dispatch(() => _viewModel?.OnSimulationComplete(e.Result, e.Sensitivity));
}
```

In `src/MonteCarlo.UI/ViewModels/MainViewModel.cs`, update `OnSimulationComplete` to forward sensitivity:

```csharp
public void OnSimulationComplete(SimulationResult result,
    Dictionary<string, IReadOnlyList<SensitivityResult>>? sensitivity = null)
{
    ResultsViewModel.LoadResults(result, sensitivity);
    // ... rest unchanged
}
```

- [ ] **Step 3: Fix ExportCurrentResults to pass sensitivity**

In `src/MonteCarlo.Addin/TaskPane/TaskPaneIntegration.cs`, update `ExportCurrentResults()` to read sensitivity from the view model:

```csharp
// Replace: sensitivity: null
// With:
var selectedId = viewModel.ResultsViewModel.SelectedOutputId;
IReadOnlyList<SensitivityResult>? sensitivity = null;
if (selectedId != null && viewModel.ResultsViewModel.SensitivityResults != null)
{
    viewModel.ResultsViewModel.SensitivityResults.TryGetValue(selectedId, out var sens);
    sensitivity = sens;
}
```

Pass `sensitivity` to `exporter.ExportSummary(...)`.

- [ ] **Step 4: Add try-catch logging around export**

Wrap the export call in `ExportCurrentResults()` with explicit logging so failures don't vanish:

```csharp
try
{
    if (exportRawData)
        exporter.ExportRawData(result, outputIndex);
    else
        exporter.ExportSummary(result, stats, sensitivity, profile, outputIndex);
    
    StartupDiagnostics.Log($"Export completed: {(exportRawData ? "raw data" : "summary")}");
}
catch (Exception ex)
{
    StartupDiagnostics.LogException("Export failed", ex);
    Dispatch(() => _viewModel?.OnSimulationError(ex));
}
```

- [ ] **Step 5: Build and test**

Run: `dotnet build src/MonteCarlo.Addin/MonteCarlo.Addin.csproj -c Release`
Expected: 0 errors

Deploy, open Excel, run simulation, click Export. Verify a new worksheet is created with stats, percentiles, and sensitivity table.

- [ ] **Step 6: Commit**

```bash
git add src/MonteCarlo.Addin/TaskPane/TaskPaneIntegration.cs src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs src/MonteCarlo.UI/ViewModels/MainViewModel.cs
git commit -m "fix: wire sensitivity data through export path"
```

---

## Task 2: Live Chart Updates During Simulation

@RISK shows the distribution building in real-time. Currently charts only render after completion. The engine already fires `ProgressChanged` every 100 iterations with the current iteration data. We need to compute incremental statistics and push them to the RunView.

**Files:**
- Modify: `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs`
- Modify: `src/MonteCarlo.Addin/TaskPane/TaskPaneIntegration.cs`
- Modify: `src/MonteCarlo.UI/ViewModels/RunViewModel.cs`
- Modify: `src/MonteCarlo.UI/ViewModels/MainViewModel.cs`
- Modify: `src/MonteCarlo.Engine/Simulation/SimulationProgressEventArgs.cs`

- [ ] **Step 1: Extend SimulationProgressEventArgs with interim stats**

In `src/MonteCarlo.Engine/Simulation/SimulationProgressEventArgs.cs`, add:

```csharp
/// <summary>Interim histogram for the first output, updated periodically.</summary>
public HistogramData? InterimHistogram { get; init; }

/// <summary>Interim CDF sorted values for the first output.</summary>
public double[]? InterimSortedValues { get; init; }
```

- [ ] **Step 2: Compute interim stats in SimulationEngine progress reporting**

In `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs`, inside the progress reporting block (every 100 iterations), compute interim statistics for the first output:

```csharp
// After existing progress computation, before firing event:
HistogramData? interimHist = null;
double[]? interimSorted = null;
if (outputCount > 0 && (i + 1) >= 200) // Need enough data for meaningful histogram
{
    var slice = new double[i + 1];
    for (int r = 0; r <= i; r++)
        slice[r] = outputMatrix[r, 0];
    Array.Sort(slice);
    interimSorted = slice;
    interimHist = new HistogramData(slice, Math.Min(50, (int)Math.Sqrt(i + 1)));
}
```

Include these in the event args:
```csharp
InterimHistogram = interimHist,
InterimSortedValues = interimSorted,
```

- [ ] **Step 3: Forward interim data to RunViewModel**

In `src/MonteCarlo.UI/ViewModels/RunViewModel.cs`, add properties and update method:

```csharp
[ObservableProperty] private HistogramData? _liveHistogramData;
[ObservableProperty] private double[]? _liveSortedValues;

public void UpdateLiveChart(HistogramData? histogram, double[]? sortedValues)
{
    LiveHistogramData = histogram;
    LiveSortedValues = sortedValues;
}
```

- [ ] **Step 4: Add live chart to RunView.xaml**

In `src/MonteCarlo.UI/Views/RunView.xaml`, add a histogram chart below the progress bar (inside the existing Grid):

```xml
<!-- Live histogram preview -->
<charts:HistogramChart HistogramData="{Binding LiveHistogramData}"
                       Height="160"
                       Margin="0,12,0,0"
                       Visibility="{Binding LiveHistogramData, Converter={StaticResource NullToCollapsed}}"/>
```

Add the charts namespace if not present:
```xml
xmlns:charts="clr-namespace:MonteCarlo.Charts.Controls;assembly=MonteCarlo.Charts"
```

- [ ] **Step 5: Wire progress events through MainViewModel**

In `src/MonteCarlo.UI/ViewModels/MainViewModel.cs`, update `OnSimulationProgress` to pass interim chart data:

```csharp
public void OnSimulationProgress(SimulationProgressEventArgs e)
{
    RunViewModel.UpdateProgress(e.CompletedIterations, e.TotalIterations, e.Elapsed, e.EstimatedRemaining);
    RunViewModel.UpdateLiveChart(e.InterimHistogram, e.InterimSortedValues);
}
```

- [ ] **Step 6: Throttle chart updates to avoid UI flooding**

In `SimulationEngine.cs`, only compute interim stats every 500 iterations (not every 100) to avoid excessive allocation:

```csharp
if (outputCount > 0 && (i + 1) >= 200 && (i + 1) % 500 == 0)
```

- [ ] **Step 7: Build and test**

Run: `dotnet build src/MonteCarlo.Addin/MonteCarlo.Addin.csproj -c Release`
Deploy, run simulation. Verify histogram appears in the run view and updates as iterations progress.

- [ ] **Step 8: Commit**

```bash
git add -A
git commit -m "feat: live histogram preview during simulation"
```

---

## Task 3: Clipboard Copy of Statistics

**Files:**
- Modify: `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs`

- [ ] **Step 1: Check existing CopyStatsToClipboardCommand**

`ResultsViewModel` already has a `CopyStatsToClipboardCommand` referenced in `TaskPaneIntegration.cs`. Find whether it exists as a stub or is missing entirely.

- [ ] **Step 2: Implement CopyStatsToClipboardCommand**

In `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs`, add or implement:

```csharp
[RelayCommand]
private void CopyStatsToClipboard()
{
    if (CurrentStats == null) return;

    var s = CurrentStats;
    var sb = new System.Text.StringBuilder();
    sb.AppendLine($"Output\t{SelectedOutputId}");
    sb.AppendLine($"Mean\t{s.Mean:G6}");
    sb.AppendLine($"Median\t{s.Median:G6}");
    sb.AppendLine($"Std Dev\t{s.StdDev:G6}");
    sb.AppendLine($"Variance\t{s.Variance:G6}");
    sb.AppendLine($"Skewness\t{s.Skewness:G6}");
    sb.AppendLine($"Kurtosis\t{s.Kurtosis:G6}");
    sb.AppendLine($"Min\t{s.Min:G6}");
    sb.AppendLine($"Max\t{s.Max:G6}");
    sb.AppendLine($"P5\t{s.P5:G6}");
    sb.AppendLine($"P10\t{s.P10:G6}");
    sb.AppendLine($"P25\t{s.P25:G6}");
    sb.AppendLine($"P50\t{s.P50:G6}");
    sb.AppendLine($"P75\t{s.P75:G6}");
    sb.AppendLine($"P90\t{s.P90:G6}");
    sb.AppendLine($"P95\t{s.P95:G6}");
    sb.AppendLine($"Iterations\t{s.Count}");

    System.Windows.Clipboard.SetText(sb.ToString());
}
```

- [ ] **Step 3: Add copy button to ResultsView.xaml**

In `src/MonteCarlo.UI/Views/ResultsView.xaml`, add a copy button near the export button:

```xml
<Button Content="Copy Stats"
        Style="{StaticResource SecondaryButton}"
        Command="{Binding CopyStatsToClipboardCommand}"
        Margin="0,0,8,0"/>
```

- [ ] **Step 4: Build and test**

Build, deploy, run simulation, click Copy Stats. Paste into Notepad — verify tab-separated stats.

- [ ] **Step 5: Commit**

```bash
git add src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs src/MonteCarlo.UI/Views/ResultsView.xaml
git commit -m "feat: copy summary statistics to clipboard"
```

---

## Task 4: Latin Hypercube Sampling

LHS stratifies each input's distribution into N equal-probability intervals and samples once from each, then shuffles. This gives much better coverage than simple Monte Carlo with the same iteration count.

**Files:**
- Create: `src/MonteCarlo.Engine/Sampling/LatinHypercubeSampler.cs`
- Modify: `src/MonteCarlo.Engine/Simulation/SimulationConfig.cs`
- Modify: `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs`
- Test: `tests/MonteCarlo.Engine.Tests/Sampling/LatinHypercubeSamplerTests.cs`

- [ ] **Step 1: Write failing test for LHS sampler**

Create `tests/MonteCarlo.Engine.Tests/Sampling/LatinHypercubeSamplerTests.cs`:

```csharp
using MonteCarlo.Engine.Sampling;

namespace MonteCarlo.Engine.Tests.Sampling;

public class LatinHypercubeSamplerTests
{
    [Fact]
    public void GenerateSamples_ReturnsCorrectDimensions()
    {
        var sampler = new LatinHypercubeSampler(seed: 42);
        double[,] samples = sampler.Generate(iterations: 100, dimensions: 3);
        Assert.Equal(100, samples.GetLength(0));
        Assert.Equal(3, samples.GetLength(1));
    }

    [Fact]
    public void GenerateSamples_EachDimensionCoversAllStrata()
    {
        var sampler = new LatinHypercubeSampler(seed: 42);
        int n = 1000;
        double[,] samples = sampler.Generate(n, dimensions: 1);

        // Each stratum [i/n, (i+1)/n] should have exactly one sample
        var strata = new bool[n];
        for (int i = 0; i < n; i++)
        {
            int stratum = (int)(samples[i, 0] * n);
            if (stratum == n) stratum = n - 1; // edge case for 1.0
            Assert.False(strata[stratum], $"Stratum {stratum} sampled twice");
            strata[stratum] = true;
        }
        Assert.All(strata, s => Assert.True(s));
    }

    [Fact]
    public void GenerateSamples_ValuesAreUniformOnUnitInterval()
    {
        var sampler = new LatinHypercubeSampler(seed: 42);
        double[,] samples = sampler.Generate(10_000, dimensions: 1);

        double mean = 0;
        for (int i = 0; i < 10_000; i++) mean += samples[i, 0];
        mean /= 10_000;

        Assert.InRange(mean, 0.49, 0.51); // Should be close to 0.5
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test tests/MonteCarlo.Engine.Tests --filter "LatinHypercube" -v n`
Expected: FAIL — `LatinHypercubeSampler` does not exist.

- [ ] **Step 3: Implement LatinHypercubeSampler**

Create `src/MonteCarlo.Engine/Sampling/LatinHypercubeSampler.cs`:

```csharp
namespace MonteCarlo.Engine.Sampling;

/// <summary>
/// Generates Latin Hypercube samples on the unit interval [0,1].
/// Each dimension is stratified into N equal-probability intervals
/// with exactly one sample per interval, then columns are shuffled independently.
/// </summary>
public class LatinHypercubeSampler
{
    private readonly Random _rng;

    public LatinHypercubeSampler(int? seed = null)
    {
        _rng = seed.HasValue ? new Random(seed.Value) : new Random();
    }

    /// <summary>
    /// Generates an [iterations x dimensions] matrix of samples on [0, 1].
    /// Each column has exactly one sample per stratum.
    /// </summary>
    public double[,] Generate(int iterations, int dimensions)
    {
        var samples = new double[iterations, dimensions];

        for (int d = 0; d < dimensions; d++)
        {
            // Create stratified samples for this dimension
            var perm = CreatePermutation(iterations);
            for (int i = 0; i < iterations; i++)
            {
                // Sample uniformly within stratum perm[i]
                double lower = (double)perm[i] / iterations;
                double upper = (double)(perm[i] + 1) / iterations;
                samples[i, d] = lower + _rng.NextDouble() * (upper - lower);
            }
        }

        return samples;
    }

    private int[] CreatePermutation(int n)
    {
        var perm = new int[n];
        for (int i = 0; i < n; i++) perm[i] = i;
        // Fisher-Yates shuffle
        for (int i = n - 1; i > 0; i--)
        {
            int j = _rng.Next(i + 1);
            (perm[i], perm[j]) = (perm[j], perm[i]);
        }
        return perm;
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test tests/MonteCarlo.Engine.Tests --filter "LatinHypercube" -v n`
Expected: 3 PASS

- [ ] **Step 5: Add SamplingMethod enum to SimulationConfig**

In `src/MonteCarlo.Engine/Simulation/SimulationConfig.cs`, add:

```csharp
public enum SamplingMethod { MonteCarlo, LatinHypercube }
```

Add property to `SimulationConfig`:
```csharp
public SamplingMethod Sampling { get; set; } = SamplingMethod.LatinHypercube;
```

Default to LHS (matching @RISK's default).

- [ ] **Step 6: Integrate LHS into SimulationEngine**

In `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs`, replace the input sampling block (where it pre-generates all samples) with a branch:

```csharp
if (config.Sampling == SamplingMethod.LatinHypercube)
{
    var lhs = new LatinHypercubeSampler(config.RandomSeed);
    var unitSamples = lhs.Generate(iterations, inputCount);
    
    // Transform unit samples through each distribution's inverse CDF
    for (int i = 0; i < iterations; i++)
        for (int j = 0; j < inputCount; j++)
            inputMatrix[i, j] = config.Inputs[j].Distribution.InverseCDF(unitSamples[i, j]);
}
else
{
    // Existing simple Monte Carlo sampling
    for (int i = 0; i < iterations; i++)
        for (int j = 0; j < inputCount; j++)
            inputMatrix[i, j] = config.Inputs[j].Distribution.Sample();
}
```

**Note:** This requires `IDistribution` to have an `InverseCDF(double p)` method. If it doesn't exist, add it to the interface and implement it for each distribution using `MathNet.Numerics` inverse CDF functions. Each existing distribution already wraps a MathNet distribution that provides `InverseCumulativeDistribution(p)`.

- [ ] **Step 7: Add InverseCDF to IDistribution if missing**

Check `src/MonteCarlo.Engine/Distributions/IDistribution.cs`. If `InverseCDF` is missing, add:

```csharp
double InverseCDF(double probability);
```

Implement in each distribution. Example for `NormalDistribution`:
```csharp
public double InverseCDF(double p) => _dist.InverseCumulativeDistribution(p);
```

Most distributions wrap MathNet distributions that already have this method.

- [ ] **Step 8: Build and run all tests**

Run: `dotnet test tests/MonteCarlo.Engine.Tests -v n`
Expected: All pass

- [ ] **Step 9: Commit**

```bash
git add -A
git commit -m "feat: Latin Hypercube Sampling with stratified coverage"
```

---

## Task 5: Convergence Auto-Stop

`ConvergenceChecker` already exists and computes Stable/Drifting/Unstable status. Wire it into the engine loop so simulation stops early when all statistics converge.

**Files:**
- Modify: `src/MonteCarlo.Engine/Simulation/SimulationConfig.cs`
- Modify: `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs`
- Modify: `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs`

- [ ] **Step 1: Add auto-stop config**

In `src/MonteCarlo.Engine/Simulation/SimulationConfig.cs`:

```csharp
/// <summary>Stop early when all tracked statistics converge (within 0.5% change rate).</summary>
public bool AutoStopOnConvergence { get; set; } = false;

/// <summary>Minimum iterations before convergence checks begin.</summary>
public int ConvergenceMinIterations { get; set; } = 500;
```

- [ ] **Step 2: Add convergence checking to the engine loop**

In `src/MonteCarlo.Engine/Simulation/SimulationEngine.cs`, add a convergence event:

```csharp
public event EventHandler<ConvergenceEventArgs>? ConvergenceChecked;
```

Create `ConvergenceEventArgs.cs` in the same directory:
```csharp
public class ConvergenceEventArgs : EventArgs
{
    public required IReadOnlyList<ConvergenceIndicator> Indicators { get; init; }
    public required bool AllConverged { get; init; }
}
```

In the sequential loop, after progress reporting (every 500 iterations past the minimum), check convergence:

```csharp
if (config.AutoStopOnConvergence && (i + 1) >= config.ConvergenceMinIterations && (i + 1) % 500 == 0)
{
    // Compute interim stats for convergence check
    var slice = new double[i + 1];
    for (int r = 0; r <= i; r++) slice[r] = outputMatrix[r, 0];
    var interimStats = new SummaryStatistics(slice);
    
    convergenceChecker.RecordCheckpoint(i + 1, interimStats);
    var indicators = convergenceChecker.CheckAll();
    bool allConverged = indicators.All(ind => ind.Status == ConvergenceStatus.Stable);
    
    ConvergenceChecked?.Invoke(this, new ConvergenceEventArgs 
    { 
        Indicators = indicators, 
        AllConverged = allConverged 
    });
    
    if (allConverged)
    {
        // Trim matrices to actual iteration count
        actualIterations = i + 1;
        break;
    }
}
```

After the loop, resize the result matrices if stopped early.

- [ ] **Step 3: Wire convergence events in SimulationOrchestrator**

In `src/MonteCarlo.Addin/Services/SimulationOrchestrator.cs`, subscribe to the engine's `ConvergenceChecked` event and forward to the UI:

```csharp
engine.ConvergenceChecked += (_, e) =>
{
    ConvergenceUpdated?.Invoke(this, new ConvergenceUpdatedEventArgs { Indicators = e.Indicators });
    if (e.AllConverged)
        StartupDiagnostics.Log($"Simulation auto-stopped: all statistics converged");
};
```

- [ ] **Step 4: Build and test**

Run: `dotnet build src/MonteCarlo.Addin/MonteCarlo.Addin.csproj -c Release`
Deploy, set iterations to 50,000 with auto-stop enabled. Run a simple model. Verify it stops before 50k when results stabilize.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: convergence auto-stop when statistics stabilize"
```

---

## Task 6: Target/Threshold Probability UI

`ResultsViewModel` already has `OnTargetValueTextChanged()`, `ProbabilityAboveTarget`, and `ProbabilityBelowTarget` properties. The UI just needs the input field and display.

**Files:**
- Modify: `src/MonteCarlo.UI/Views/ResultsView.xaml`

- [ ] **Step 1: Check existing target properties in ResultsViewModel**

Verify `TargetValueText`, `ProbabilityAboveTarget`, `ProbabilityBelowTarget`, and `TargetAnnotation` properties exist and work. These are already implemented.

- [ ] **Step 2: Add target input section to ResultsView.xaml**

In `src/MonteCarlo.UI/Views/ResultsView.xaml`, after the chart toggle radio buttons and before the end of the chart card Border, add:

```xml
<!-- Target analysis -->
<Border Style="{StaticResource Card}" Margin="0,12,0,0" Padding="12,8">
    <StackPanel>
        <TextBlock Text="Target Analysis" 
                   FontWeight="SemiBold" FontSize="13"
                   Foreground="{StaticResource TextPrimary}" Margin="0,0,0,6"/>
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Target value:" VerticalAlignment="Center" Margin="0,0,8,0"
                       Foreground="{StaticResource TextSecondary}"/>
            <TextBox Grid.Column="1" 
                     Text="{Binding TargetValueText, UpdateSourceTrigger=PropertyChanged}"
                     Style="{StaticResource DefaultTextBox}"/>
        </Grid>
        <StackPanel Margin="0,8,0,0"
                    Visibility="{Binding ProbabilityAboveTarget, Converter={StaticResource NullToCollapsed}}">
            <TextBlock Foreground="{StaticResource TextSecondary}">
                <Run Text="P(X > target) = "/>
                <Run Text="{Binding ProbabilityAboveTarget, StringFormat=P2, Mode=OneWay}" 
                     FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}"/>
            </TextBlock>
            <TextBlock Foreground="{StaticResource TextSecondary}">
                <Run Text="P(X ≤ target) = "/>
                <Run Text="{Binding ProbabilityBelowTarget, StringFormat=P2, Mode=OneWay}"
                     FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}"/>
            </TextBlock>
        </StackPanel>
    </StackPanel>
</Border>
```

- [ ] **Step 3: Add target line to charts**

In `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs`, when target is set and valid, store `TargetValueNumeric` (already exists). The `HistogramChart` already binds `TargetValue` and `TargetLabel` — verify these render the red dashed line via `ChartTheme.CreateTargetLine()`.

- [ ] **Step 4: Build and test**

Deploy, run simulation, enter a target value. Verify probability percentages appear and the target line shows on the histogram.

- [ ] **Step 5: Commit**

```bash
git add src/MonteCarlo.UI/Views/ResultsView.xaml
git commit -m "feat: target probability analysis UI"
```

---

## Task 7: Five New Distributions

Add Gamma, Logistic, GEV, Binomial, and Geometric. Follow the existing pattern: distribution class + factory registration + Excel UDF.

**Files:**
- Create: `src/MonteCarlo.Engine/Distributions/GammaDistribution.cs`
- Create: `src/MonteCarlo.Engine/Distributions/LogisticDistribution.cs`
- Create: `src/MonteCarlo.Engine/Distributions/GEVDistribution.cs`
- Create: `src/MonteCarlo.Engine/Distributions/BinomialDistribution.cs`
- Create: `src/MonteCarlo.Engine/Distributions/GeometricDistribution.cs`
- Modify: `src/MonteCarlo.Engine/Distributions/DistributionFactory.cs`
- Modify: `src/MonteCarlo.Addin/UDF/MonteCarloFunctions.cs`
- Test: `tests/MonteCarlo.Engine.Tests/Distributions/NewDistributionTests.cs`

- [ ] **Step 1: Write failing tests for all five distributions**

Create `tests/MonteCarlo.Engine.Tests/Distributions/NewDistributionTests.cs`:

```csharp
using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.Engine.Tests.Distributions;

public class NewDistributionTests
{
    [Fact]
    public void Gamma_SampleMeanMatchesExpected()
    {
        var dist = new GammaDistribution(shape: 2.0, rate: 0.5, seed: 42);
        var samples = Enumerable.Range(0, 50_000).Select(_ => dist.Sample()).ToArray();
        Assert.InRange(samples.Average(), 3.8, 4.2); // E[X] = shape/rate = 4.0
    }

    [Fact]
    public void Logistic_SampleMeanMatchesExpected()
    {
        var dist = new LogisticDistribution(mu: 5.0, s: 2.0, seed: 42);
        var samples = Enumerable.Range(0, 50_000).Select(_ => dist.Sample()).ToArray();
        Assert.InRange(samples.Average(), 4.8, 5.2); // E[X] = mu = 5.0
    }

    [Fact]
    public void Binomial_SampleMeanMatchesExpected()
    {
        var dist = new BinomialDistribution(n: 20, p: 0.3, seed: 42);
        var samples = Enumerable.Range(0, 50_000).Select(_ => dist.Sample()).ToArray();
        Assert.InRange(samples.Average(), 5.8, 6.2); // E[X] = n*p = 6.0
    }

    [Fact]
    public void Geometric_SampleMeanMatchesExpected()
    {
        var dist = new GeometricDistribution(p: 0.25, seed: 42);
        var samples = Enumerable.Range(0, 50_000).Select(_ => dist.Sample()).ToArray();
        Assert.InRange(samples.Average(), 3.8, 4.2); // E[X] = 1/p = 4.0
    }

    [Fact]
    public void GEV_TypeI_SampleMeanMatchesExpected()
    {
        // Type I (Gumbel): xi=0, mu=0, sigma=1 => E[X] = mu + sigma*gamma ≈ 0.5772
        var dist = new GEVDistribution(mu: 0, sigma: 1, xi: 0, seed: 42);
        var samples = Enumerable.Range(0, 50_000).Select(_ => dist.Sample()).ToArray();
        Assert.InRange(samples.Average(), 0.5, 0.65);
    }

    [Fact]
    public void DistributionFactory_CreatesAllNewDistributions()
    {
        foreach (var name in new[] { "gamma", "logistic", "gev", "binomial", "geometric" })
        {
            Assert.Contains(name, DistributionFactory.AvailableDistributions);
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test tests/MonteCarlo.Engine.Tests --filter "NewDistribution" -v n`
Expected: FAIL

- [ ] **Step 3: Implement all five distribution classes**

Each follows this pattern (using MathNet.Numerics):

**GammaDistribution.cs:**
```csharp
using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

public class GammaDistribution : IDistribution
{
    private readonly Gamma _dist;
    public string Name => "Gamma";
    public double ExpectedValue { get; }

    public GammaDistribution(double shape, double rate, int? seed = null)
    {
        _dist = new Gamma(shape, rate, seed.HasValue ? new Random(seed.Value) : new Random());
        ExpectedValue = shape / rate;
    }

    public double Sample() => _dist.Sample();
    public double InverseCDF(double p) => _dist.InverseCumulativeDistribution(p);
}
```

**LogisticDistribution.cs** — uses `MathNet.Numerics` does not have Logistic; implement manually:
```csharp
namespace MonteCarlo.Engine.Distributions;

public class LogisticDistribution : IDistribution
{
    private readonly Random _rng;
    private readonly double _mu, _s;
    public string Name => "Logistic";
    public double ExpectedValue => _mu;

    public LogisticDistribution(double mu, double s, int? seed = null)
    {
        _mu = mu;
        _s = s;
        _rng = seed.HasValue ? new Random(seed.Value) : new Random();
    }

    public double Sample()
    {
        double u = _rng.NextDouble();
        return InverseCDF(u);
    }

    public double InverseCDF(double p) => _mu + _s * Math.Log(p / (1.0 - p));
}
```

**GEVDistribution.cs** — Generalized Extreme Value:
```csharp
namespace MonteCarlo.Engine.Distributions;

public class GEVDistribution : IDistribution
{
    private readonly Random _rng;
    private readonly double _mu, _sigma, _xi;
    public string Name => "GEV";
    public double ExpectedValue { get; }

    public GEVDistribution(double mu, double sigma, double xi, int? seed = null)
    {
        _mu = mu; _sigma = sigma; _xi = xi;
        _rng = seed.HasValue ? new Random(seed.Value) : new Random();

        // E[X] = mu + sigma * (Gamma(1-xi) - 1) / xi  for xi != 0 and xi < 1
        if (Math.Abs(xi) < 1e-10)
            ExpectedValue = mu + sigma * 0.5772156649; // Euler-Mascheroni
        else if (xi < 1)
            ExpectedValue = mu + sigma * (MathNet.Numerics.SpecialFunctions.Gamma(1 - xi) - 1) / xi;
        else
            ExpectedValue = double.PositiveInfinity;
    }

    public double Sample() => InverseCDF(_rng.NextDouble());

    public double InverseCDF(double p)
    {
        double t = -Math.Log(p);
        if (Math.Abs(_xi) < 1e-10)
            return _mu - _sigma * Math.Log(t);
        return _mu + _sigma * (Math.Pow(t, -_xi) - 1.0) / _xi;
    }
}
```

**BinomialDistribution.cs** and **GeometricDistribution.cs**: similar pattern using `MathNet.Numerics.Distributions.Binomial` and manual geometric implementation.

- [ ] **Step 4: Register in DistributionFactory**

Add cases to the switch in `DistributionFactory.Create()`:
```csharp
"gamma" => new GammaDistribution(GetRequired(p, "shape"), GetRequired(p, "rate"), seed),
"logistic" => new LogisticDistribution(GetRequired(p, "mu"), GetRequired(p, "s"), seed),
"gev" => new GEVDistribution(GetRequired(p, "mu"), GetRequired(p, "sigma"), GetRequired(p, "xi"), seed),
"binomial" => new BinomialDistribution((int)GetRequired(p, "n"), GetRequired(p, "p"), seed),
"geometric" => new GeometricDistribution(GetRequired(p, "p"), seed),
```

Add names to `_available` array.

- [ ] **Step 5: Add Excel UDFs**

In `src/MonteCarlo.Addin/UDF/MonteCarloFunctions.cs`, add UDFs following the existing pattern:

```csharp
[ExcelFunction(Name = "MC.Gamma", Description = "Gamma distribution")]
public static object MCGamma(double shape, double rate)
{
    if (shape <= 0 || rate <= 0) return ExcelError.ExcelErrorValue;
    return shape / rate; // Expected value
}

[ExcelFunction(Name = "MC.Logistic", Description = "Logistic distribution")]
public static object MCLogistic(double mu, double s)
{
    if (s <= 0) return ExcelError.ExcelErrorValue;
    return mu;
}

[ExcelFunction(Name = "MC.Binomial", Description = "Binomial distribution")]
public static object MCBinomial(int n, double p)
{
    if (n <= 0 || p < 0 || p > 1) return ExcelError.ExcelErrorValue;
    return (double)(n * p);
}

[ExcelFunction(Name = "MC.Geometric", Description = "Geometric distribution")]
public static object MCGeometric(double p)
{
    if (p <= 0 || p > 1) return ExcelError.ExcelErrorValue;
    return 1.0 / p;
}

[ExcelFunction(Name = "MC.GEV", Description = "Generalized Extreme Value distribution")]
public static object MCGEV(double mu, double sigma, double xi)
{
    if (sigma <= 0) return ExcelError.ExcelErrorValue;
    if (Math.Abs(xi) < 1e-10) return mu + sigma * 0.5772;
    if (xi < 1) return mu + sigma * (MathNet.Numerics.SpecialFunctions.Gamma(1 - xi) - 1) / xi;
    return ExcelError.ExcelErrorValue;
}
```

- [ ] **Step 6: Run all tests**

Run: `dotnet test tests/MonteCarlo.Engine.Tests -v n`
Expected: All pass

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat: add Gamma, Logistic, GEV, Binomial, Geometric distributions"
```

---

## Task 8: PDF Overlay on Histogram

Show the theoretical probability density function (smooth curve) overlaid on the histogram bars. This requires computing the PDF from the fitted distribution at each bin center.

**Files:**
- Modify: `src/MonteCarlo.Charts/Controls/HistogramChart.xaml.cs`
- Modify: `src/MonteCarlo.Engine/Analysis/HistogramData.cs`

- [ ] **Step 1: Add PDF curve data to HistogramData**

In `src/MonteCarlo.Engine/Analysis/HistogramData.cs`, add a static method that generates a smooth PDF curve from kernel density estimation:

```csharp
/// <summary>
/// Generates a smooth kernel density estimate for overlay on the histogram.
/// Returns (x, y) pairs where y is scaled to match relative frequency density.
/// </summary>
public (double[] X, double[] Y) ComputeKDE(int points = 200)
{
    double min = BinEdges[0];
    double max = BinEdges[^1];
    double range = max - min;
    if (range <= 0) return (Array.Empty<double>(), Array.Empty<double>());

    double h = 1.06 * StdDev * Math.Pow(Count, -0.2); // Silverman's rule
    if (h <= 0) h = range / 50;

    var xs = new double[points];
    var ys = new double[points];
    double step = range / (points - 1);

    for (int i = 0; i < points; i++)
    {
        xs[i] = min + i * step;
        double sum = 0;
        for (int j = 0; j < Count; j++)
        {
            double z = (xs[i] - _sortedValues[j]) / h;
            sum += Math.Exp(-0.5 * z * z);
        }
        ys[i] = sum / (Count * h * Math.Sqrt(2 * Math.PI));
    }

    // Scale to match histogram relative frequency density
    double yScale = BinWidth;
    for (int i = 0; i < points; i++) ys[i] *= yScale;

    return (xs, ys);
}
```

Store `_sortedValues`, `Count`, and `StdDev` as fields from the constructor.

- [ ] **Step 2: Add PDF overlay series to HistogramChart**

In `src/MonteCarlo.Charts/Controls/HistogramChart.xaml.cs`, in the `RebuildChart()` method, after adding the bar series, add a line series for the PDF:

```csharp
// PDF overlay
var (kdeX, kdeY) = histogramData.ComputeKDE();
if (kdeX.Length > 0)
{
    var pdfPoints = new List<LiveChartsCore.Defaults.ObservablePoint>();
    for (int i = 0; i < kdeX.Length; i++)
        pdfPoints.Add(new LiveChartsCore.Defaults.ObservablePoint(kdeX[i], kdeY[i]));

    Series.Add(new LineSeries<LiveChartsCore.Defaults.ObservablePoint>
    {
        Values = pdfPoints,
        Stroke = new SolidColorPaint(ChartTheme.Violet500) { StrokeThickness = 2 },
        Fill = null,
        GeometrySize = 0,
        LineSmoothness = 0.7,
        ScalesYAt = 1 // Use secondary Y axis (hidden, auto-scaled)
    });
}
```

Add a hidden secondary Y axis for the PDF scale:
```csharp
YAxes.Add(new Axis { ShowSeparatorLines = false, IsVisible = false });
```

- [ ] **Step 3: Build and test**

Deploy, run simulation, view histogram. Verify smooth purple PDF curve overlays the blue bars.

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "feat: KDE probability density overlay on histogram"
```

---

## Task 9: Sensitivity Scatter Plot

Add a scatter plot showing input vs. output values for the selected sensitivity input. Users click a row in the tornado chart or a dropdown to see the scatter.

**Files:**
- Create: `src/MonteCarlo.Charts/Controls/ScatterChart.xaml`
- Create: `src/MonteCarlo.Charts/Controls/ScatterChart.xaml.cs`
- Modify: `src/MonteCarlo.UI/Views/ResultsView.xaml`
- Modify: `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs`

- [ ] **Step 1: Create ScatterChart control**

Create `src/MonteCarlo.Charts/Controls/ScatterChart.xaml`:
```xml
<UserControl x:Class="MonteCarlo.Charts.Controls.ScatterChart"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             xmlns:lvc="clr-namespace:LiveChartsCore.SkiaSharpView.WPF;assembly=LiveChartsCore.SkiaSharpView.WPF">
    <Grid Background="Transparent">
        <lvc:CartesianChart x:Name="Chart"
                            Series="{Binding Series, RelativeSource={RelativeSource AncestorType=UserControl}}"
                            XAxes="{Binding XAxes, RelativeSource={RelativeSource AncestorType=UserControl}}"
                            YAxes="{Binding YAxes, RelativeSource={RelativeSource AncestorType=UserControl}}"
                            TooltipPosition="Hidden"/>
    </Grid>
</UserControl>
```

Create `src/MonteCarlo.Charts/Controls/ScatterChart.xaml.cs`:
```csharp
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using MonteCarlo.Charts.Themes;

namespace MonteCarlo.Charts.Controls;

public partial class ScatterChart : UserControl
{
    public static readonly DependencyProperty InputValuesProperty =
        DependencyProperty.Register(nameof(InputValues), typeof(double[]), typeof(ScatterChart),
            new PropertyMetadata(null, OnDataChanged));

    public static readonly DependencyProperty OutputValuesProperty =
        DependencyProperty.Register(nameof(OutputValues), typeof(double[]), typeof(ScatterChart),
            new PropertyMetadata(null, OnDataChanged));

    public static readonly DependencyProperty InputLabelProperty =
        DependencyProperty.Register(nameof(InputLabel), typeof(string), typeof(ScatterChart),
            new PropertyMetadata("Input", OnDataChanged));

    public double[]? InputValues { get => (double[]?)GetValue(InputValuesProperty); set => SetValue(InputValuesProperty, value); }
    public double[]? OutputValues { get => (double[]?)GetValue(OutputValuesProperty); set => SetValue(OutputValuesProperty, value); }
    public string InputLabel { get => (string)GetValue(InputLabelProperty); set => SetValue(InputLabelProperty, value); }

    public ObservableCollection<ISeries> Series { get; } = new();
    public ObservableCollection<Axis> XAxes { get; } = new();
    public ObservableCollection<Axis> YAxes { get; } = new();

    public ScatterChart()
    {
        _ = typeof(LiveChartsCore.SkiaSharpView.WPF.CartesianChart).Assembly;
        InitializeComponent();
    }

    private static void OnDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e) => ((ScatterChart)d).Rebuild();

    private void Rebuild()
    {
        Series.Clear(); XAxes.Clear(); YAxes.Clear();
        var ix = InputValues; var ox = OutputValues;
        if (ix == null || ox == null || ix.Length == 0) return;

        int n = Math.Min(ix.Length, ox.Length);
        int step = Math.Max(1, n / 500); // Subsample to 500 points max
        var points = new List<LiveChartsCore.Defaults.ObservablePoint>();
        for (int i = 0; i < n; i += step)
            points.Add(new LiveChartsCore.Defaults.ObservablePoint(ix[i], ox[i]));

        Series.Add(new ScatterSeries<LiveChartsCore.Defaults.ObservablePoint>
        {
            Values = points,
            Stroke = null,
            Fill = new SolidColorPaint(ChartTheme.Blue500_85),
            GeometrySize = 4
        });

        XAxes.Add(ChartTheme.CreateXAxis(name: InputLabel));
        YAxes.Add(ChartTheme.CreateYAxis(name: "Output"));
    }
}
```

- [ ] **Step 2: Add scatter plot selection to ResultsViewModel**

In `src/MonteCarlo.UI/ViewModels/ResultsViewModel.cs`:

```csharp
[ObservableProperty] private string? _selectedScatterInputId;
[ObservableProperty] private double[]? _scatterInputValues;
[ObservableProperty] private double[]? _scatterOutputValues;
[ObservableProperty] private string? _scatterInputLabel;

partial void OnSelectedScatterInputIdChanged(string? value)
{
    if (value == null || _result == null) { ScatterInputValues = null; return; }

    // Find input index
    int inputIdx = -1;
    for (int i = 0; i < _result.InputIds.Length; i++)
        if (_result.InputIds[i] == value) { inputIdx = i; break; }
    if (inputIdx < 0) { ScatterInputValues = null; return; }

    // Find output index
    int outputIdx = Array.IndexOf(_result.OutputIds, SelectedOutputId);
    if (outputIdx < 0) { ScatterInputValues = null; return; }

    int n = _result.Iterations;
    var ix = new double[n]; var ox = new double[n];
    for (int i = 0; i < n; i++)
    {
        ix[i] = _result.InputMatrix[i, inputIdx];
        ox[i] = _result.OutputMatrix[i, outputIdx];
    }
    ScatterInputValues = ix;
    ScatterOutputValues = ox;
    ScatterInputLabel = value;
}
```

- [ ] **Step 3: Add scatter chart and input selector to ResultsView.xaml**

After the tornado chart section:

```xml
<!-- Sensitivity scatter plot -->
<Border Style="{StaticResource Card}" Margin="0,0,0,12" Padding="8"
        Visibility="{Binding ScatterInputValues, Converter={StaticResource NullToCollapsed}}">
    <StackPanel>
        <Grid Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Scatter: " VerticalAlignment="Center" FontWeight="SemiBold"
                       Foreground="{StaticResource TextPrimary}"/>
            <ComboBox Grid.Column="1"
                      ItemsSource="{Binding AvailableInputs}"
                      SelectedItem="{Binding SelectedScatterInputId}"
                      Style="{StaticResource DefaultComboBox}"/>
        </Grid>
        <charts:ScatterChart InputValues="{Binding ScatterInputValues}"
                             OutputValues="{Binding ScatterOutputValues}"
                             InputLabel="{Binding ScatterInputLabel}"
                             Height="200"/>
    </StackPanel>
</Border>
```

- [ ] **Step 4: Populate AvailableInputs on results load**

In `ResultsViewModel.LoadResults()`, populate:
```csharp
AvailableInputs = new ObservableCollection<string>(result.InputIds);
SelectedScatterInputId = AvailableInputs.FirstOrDefault();
```

- [ ] **Step 5: Build and test**

Deploy, run simulation, scroll to scatter section. Select different inputs from dropdown. Verify scatter plot updates.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat: sensitivity scatter plot for input vs output"
```

---

## Testing Strategy

### Automated (run after each task)

```bash
# Unit tests
dotnet test tests/MonteCarlo.Engine.Tests -v n

# Build verification
dotnet build src/MonteCarlo.Addin/MonteCarlo.Addin.csproj -c Release
```

### Excel Integration (run after deploy)

PowerShell COM automation script to verify end-to-end:

```powershell
# Deploy
$src = "...\publish\MonteCarlo.Addin-AddIn64-packed.xll"
$dst = "$env:APPDATA\Microsoft\Excel\XLSTART\MonteCarlo.XL.xll"
Copy-Item $src $dst -Force

# Screenshot verification
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Screen]::PrimaryScreen | Select Bounds
# Take screenshot of Excel window after simulation
```

### Visual Verification Checklist

After each tier, verify in Excel:
- [ ] Tier 1: Export creates worksheet with stats + sensitivity table. Clipboard paste works. Live histogram appears during run.
- [ ] Tier 2: Simulation stops early when converged. Target probability shows P(X > target). Histogram uses LHS (smoother with fewer iterations).
- [ ] Tier 3: MC.Gamma/MC.Logistic/MC.GEV/MC.Binomial/MC.Geometric UDFs work. PDF curve overlays histogram. Scatter plot shows input-output relationship.
