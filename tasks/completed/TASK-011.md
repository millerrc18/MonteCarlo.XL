# TASK-011: Config Persistence, Run View, and Convergence Monitor

## Context

Read `ROADMAP.md` for full project context. This task covers three related features that round out Phase 1: saving/loading simulation configs inside the workbook, the progress UI during simulation, and convergence monitoring.

## Dependencies

- TASK-003 (SimulationEngine with progress events)
- TASK-004 (SummaryStatistics)
- TASK-006 (WorkbookManager — for CustomXMLPart access)
- TASK-007 (Task pane shell, GlobalStyles)
- TASK-008 (SetupViewModel — triggers simulation)

## Part A: Config Persistence

### Objective

Save and load the complete simulation configuration (inputs, outputs, distributions, parameters, iteration count) inside the Excel workbook so it travels with the file.

### Storage: CustomXMLPart

Excel supports `CustomXMLParts` — arbitrary XML data stored inside the .xlsx file. This is the cleanest approach: invisible to the user, survives save/close/reopen, and travels when the file is shared.

```csharp
public class ConfigPersistence
{
    private const string CustomXmlNamespace = "urn:montecarlo-xl:config:v1";

    /// Save the current simulation config into the active workbook.
    public void Save(SimulationProfile profile);

    /// Load the simulation config from the active workbook. Returns null if none exists.
    public SimulationProfile? Load();

    /// Delete any saved config from the workbook.
    public void Clear();

    /// List all saved profiles (for Phase 3 multi-profile support).
    public List<string> GetProfileNames();
}
```

### SimulationProfile

```csharp
public class SimulationProfile
{
    public string Name { get; set; } = "Default";
    public int IterationCount { get; set; } = 5000;
    public int? RandomSeed { get; set; }

    public List<SavedInput> Inputs { get; set; } = new();
    public List<SavedOutput> Outputs { get; set; } = new();
    // Phase 3: public double[,]? CorrelationMatrix { get; set; }
}

public class SavedInput
{
    public string SheetName { get; set; }
    public string CellAddress { get; set; }
    public string Label { get; set; }
    public string DistributionName { get; set; }
    public Dictionary<string, double> Parameters { get; set; }
}

public class SavedOutput
{
    public string SheetName { get; set; }
    public string CellAddress { get; set; }
    public string Label { get; set; }
}
```

### Serialization

Serialize `SimulationProfile` to JSON, then wrap it in a minimal XML envelope for the CustomXMLPart:

```xml
<MonteCarloConfig xmlns="urn:montecarlo-xl:config:v1">
  <Profile name="Default">
    <![CDATA[
      { "iterationCount": 5000, "inputs": [...], "outputs": [...] }
    ]]>
  </Profile>
</MonteCarloConfig>
```

Use `System.Text.Json` for JSON serialization.

### Auto-Save / Auto-Load

- **Auto-load:** When the add-in starts (or when a workbook is activated), check for a CustomXMLPart with our namespace. If found, load the config and populate the SetupView.
- **Auto-save:** Whenever the user modifies the config (add/remove input, change distribution, change iterations), save automatically. Debounce to avoid excessive writes — save at most once per second.

### Fallback: Hidden Sheet

If CustomXMLParts prove unreliable (some versions of Excel or third-party tools strip them), implement a fallback that stores JSON in a cell on a hidden sheet named `__MC_Config`. Check for the hidden sheet on load if no CustomXMLPart is found.

---

## Part B: Run View

### Objective

Build the UI that displays during a simulation run — progress bar, iteration counter, elapsed time, estimated time remaining, and a live-updating stats preview.

### Layout

```
┌───────────────────────────────────┐
│ RUNNING SIMULATION                │
├───────────────────────────────────┤
│                                   │
│ ████████████████░░░░░░░░  67%    │  ← Progress bar (animated)
│                                   │
│ Iterations:  3,350 / 5,000       │
│ Elapsed:     2.4s                │
│ Remaining:   ~1.2s               │
│                                   │
├───────────────────────────────────┤
│ LIVE PREVIEW         (updates    │
│                      every 500)  │
│ ┌───────────────────────────────┐ │
│ │ Mean:    $2,430,000           │ │  ← Running stats
│ │ Median:  $2,380,000           │ │
│ │ Std Dev: $815,000             │ │
│ │ P5:      $1,050,000           │ │
│ │ P95:     $3,820,000           │ │
│ └───────────────────────────────┘ │
│                                   │
│ ┌───────────────────────────────┐ │
│ │ Convergence:                  │ │
│ │ Mean    ━━━━━━━━━━━━ stable ✓ │ │  ← Converging indicator
│ │ P50     ━━━━━━━━━━━━ stable ✓ │ │
│ │ P90     ━━━━━━━━━━━ drifting  │ │
│ └───────────────────────────────┘ │
│                                   │
│ [       ■ Stop Simulation       ] │
└───────────────────────────────────┘
```

### RunViewModel

```csharp
public partial class RunViewModel : ObservableObject
{
    [ObservableProperty] private double _progressPercent;
    [ObservableProperty] private int _completedIterations;
    [ObservableProperty] private int _totalIterations;
    [ObservableProperty] private string _elapsedTime;
    [ObservableProperty] private string _estimatedRemaining;

    // Live stats (updated periodically during the run)
    [ObservableProperty] private string _liveMean;
    [ObservableProperty] private string _liveMedian;
    [ObservableProperty] private string _liveStdDev;
    [ObservableProperty] private string _liveP5;
    [ObservableProperty] private string _liveP95;

    // Convergence indicators
    [ObservableProperty] private ObservableCollection<ConvergenceIndicator> _convergenceIndicators;

    [RelayCommand] private void StopSimulation();
}

public class ConvergenceIndicator
{
    public string StatName { get; set; }           // "Mean", "P50", "P90"
    public bool IsStable { get; set; }
    public double CurrentValue { get; set; }
    public double ChangeRate { get; set; }         // % change over last window
}
```

### Live Stats Update

During the simulation, update live stats every 500 iterations (not every iteration — too expensive):
1. Take the output values collected so far
2. Compute SummaryStatistics on the partial results
3. Update the RunViewModel properties

This happens on the SimulationEngine's `ProgressChanged` event, dispatched to the WPF UI thread.

### Progress Bar

WPF `ProgressBar` styled to match the design system:
- Height: 8px, rounded corners (4px radius)
- Background: `#E2E8F0`
- Fill: `#3B82F6` (blue-500)
- Animated fill transition (smooth, not jumpy)

---

## Part C: Convergence Monitor

### Objective

Detect when simulation statistics have stabilized, helping the user know if they've run enough iterations.

### ConvergenceChecker

```csharp
public class ConvergenceChecker
{
    /// Check convergence of a running statistic.
    /// Takes the stat values at regular checkpoints (e.g., every 500 iterations).
    /// Returns true if the stat has stabilized.
    public bool IsConverged(double[] checkpointValues, double tolerance = 0.005);

    /// Check multiple stats at once, returning a status for each.
    public List<ConvergenceIndicator> CheckAll(
        double[] outputValues,
        int currentIteration,
        int checkpointInterval = 500
    );
}
```

### Convergence Logic

A statistic is "converged" when its value changes by less than `tolerance` (default 0.5%) over the last `windowSize` checkpoints (default: last 3 checkpoints = last 1,500 iterations):

```
changeRate = |currentValue - valueAtWindowStart| / |valueAtWindowStart|
isConverged = changeRate < tolerance
```

Check these stats: Mean, P50, P90, StdDev.

Display in the Run view:
- ✓ (green) if converged
- "drifting" (amber) if not converged but improving (change rate decreasing)
- "unstable" (red) if change rate is increasing (rare, usually means the model is weird)

---

## File Structure

```
MonteCarlo.Addin/
└── Excel/
    └── ConfigPersistence.cs

MonteCarlo.Engine/
├── Simulation/
│   └── SimulationProfile.cs         # Config model (SavedInput, SavedOutput)
└── Analysis/
    └── ConvergenceChecker.cs

MonteCarlo.UI/
└── Views/
    ├── RunView.xaml/.cs
    └── ConvergenceIndicatorControl.xaml/.cs
└── ViewModels/
    └── RunViewModel.cs

MonteCarlo.Engine.Tests/
└── Analysis/
    └── ConvergenceCheckerTests.cs
```

## Tests

### ConvergenceCheckerTests

1. **Stable series converges** — feed checkpoint values [100.5, 100.2, 100.1, 100.08, 100.05] → should report converged
2. **Unstable series does not converge** — [100, 102, 98, 105, 95] → not converged
3. **Eventually converges** — long series that starts volatile then stabilizes
4. **Tolerance parameter works** — tight tolerance (0.001) requires more stability than default (0.005)
5. **Single checkpoint** — not enough data, should report "not enough data" rather than converged

### ConfigPersistence

Manual testing only (requires Excel COM). Verify:
- Save → close workbook → reopen → load returns same config
- Config survives "Save As" to a new file
- No config returns null (not an error)

## Commit Strategy

```
feat(engine): add SimulationProfile and SavedInput/SavedOutput models
feat(addin): implement ConfigPersistence using CustomXMLParts
feat(engine): add ConvergenceChecker with rolling window stability detection
feat(ui): add RunView with progress bar, live stats, and convergence indicators
test(engine): add ConvergenceChecker tests
```

## Done When

- [ ] Config saves to CustomXMLPart and loads on workbook open
- [ ] Auto-save triggers on config changes (debounced)
- [ ] RunView shows animated progress bar with iteration count and timing
- [ ] Live stats update every 500 iterations during a run
- [ ] Convergence indicators show stable/drifting status for key stats
- [ ] Stop button cancels the simulation cleanly
- [ ] `dotnet build` clean, `dotnet test` green
