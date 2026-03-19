# TASK-013: Simulation Orchestrator — End-to-End Integration

## Context

Read `ROADMAP.md` for full project context. Tasks 002–012 build individual components: engine, statistics, sensitivity, Excel I/O, UI views, charts, config persistence, and export. This task wires everything together into a working end-to-end flow. This is the **integration task** — it connects the dots.

## Dependencies

- All of TASK-002 through TASK-012 (or at minimum: 002, 003, 004, 006, 007, 008, 009, 011)
- TASK-005, 010, 012 are needed for tornado/export but the core flow works without them

## Objective

Build the `SimulationOrchestrator` service that coordinates the full simulation lifecycle: read inputs from Excel → run simulation → compute statistics → display results → enable export. Also wire up the MainViewModel to drive the flow through the UI.

## Design

### SimulationOrchestrator

This is the central coordinator — it owns the workflow from "Run" button click to results displayed.

```csharp
public class SimulationOrchestrator
{
    private readonly IWorkbookManager _workbook;
    private readonly InputTagManager _inputManager;
    private readonly OutputTagManager _outputManager;
    private readonly SimulationEngine _engine;
    private readonly ConfigPersistence _configPersistence;

    private CancellationTokenSource? _cts;
    private SimulationResult? _lastResult;

    public event EventHandler<SimulationProgressEventArgs>? ProgressChanged;
    public event EventHandler<SimulationCompleteEventArgs>? SimulationComplete;
    public event EventHandler<SimulationErrorEventArgs>? SimulationError;

    /// Full simulation workflow.
    public async Task RunSimulationAsync()
    {
        // 1. Read current config from InputTagManager + OutputTagManager
        // 2. Build SimulationConfig
        //    - For each tagged input: read BaseValue from Excel, attach Distribution
        //    - For each tagged output: record cell reference
        // 3. Build the evaluator function:
        //    Fast mode: evaluator reads formula results after writing input values
        //    (See Evaluator section below)
        // 4. Start the engine on a background thread
        // 5. Forward progress events to UI (marshal to WPF dispatcher)
        // 6. On completion: compute SummaryStatistics for each output
        // 7. Compute SensitivityResult for each output
        // 8. Fire SimulationComplete with all results
    }

    /// Cancel a running simulation.
    public void CancelSimulation()
    {
        _cts?.Cancel();
    }
}

public class SimulationCompleteEventArgs : EventArgs
{
    public SimulationResult Result { get; }
    public Dictionary<string, SummaryStatistics> StatsByOutput { get; }
    public Dictionary<string, SensitivityResult> SensitivityByOutput { get; }
    public TimeSpan TotalElapsed { get; }
}
```

### Evaluator Function — Fast Mode

In Phase 1 "fast mode," the output cells contain formulas that depend on the input cells. The evaluator:

1. Writes input values to their Excel cells (batch write via WorkbookManager)
2. Triggers `Application.Calculate()` to recalculate the workbook
3. Reads output cell values (batch read)
4. Returns the output dictionary

```csharp
private Func<Dictionary<string, double>, Dictionary<string, double>> BuildEvaluator()
{
    return inputs =>
    {
        // Write all input values to Excel
        foreach (var (id, value) in inputs)
        {
            var cell = _inputManager.GetCellForInput(id);
            _workbook.WriteCellValue(cell.SheetName, cell.CellAddress, value);
        }

        // Recalculate
        var app = (Application)ExcelDnaUtil.Application;
        app.Calculate();

        // Read all output values
        var outputs = new Dictionary<string, double>();
        foreach (var output in _outputManager.GetAllOutputs())
        {
            outputs[output.Cell.FullReference] =
                _workbook.ReadCellValue(output.Cell.SheetName, output.Cell.CellAddress);
        }

        return outputs;
    };
}
```

**IMPORTANT:** This evaluator is called on every iteration and touches Excel COM, which is single-threaded. The simulation loop must be sequential (not parallel) when using this evaluator. Performance will be roughly 100–500 iterations/second depending on workbook complexity.

**Performance optimization:** Disable screen updating and set calculation to manual at the start, restore at the end:

```csharp
var app = (Application)ExcelDnaUtil.Application;
app.ScreenUpdating = false;
app.Calculation = XlCalculation.xlCalculationManual;
try
{
    await _engine.RunAsync(config, evaluator, _cts.Token);
}
finally
{
    app.Calculation = XlCalculation.xlCalculationAutomatic;
    app.ScreenUpdating = true;
    // Restore original input values
    RestoreOriginalValues(originalValues);
}
```

**Restore original values:** Before starting, save the original cell values for all inputs. After simulation completes (or is cancelled), write the original values back so the spreadsheet returns to its pre-simulation state.

### MainViewModel Updates

Wire the orchestrator into the navigation flow:

```csharp
public partial class MainViewModel : ObservableObject
{
    private readonly SimulationOrchestrator _orchestrator;

    [RelayCommand]
    private async Task RunSimulation()
    {
        // Navigate to RunView
        CurrentView = _runView;
        _runViewModel.Reset();

        // Wire progress events
        _orchestrator.ProgressChanged += (s, e) =>
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                _runViewModel.UpdateProgress(e);
            });
        };

        try
        {
            await _orchestrator.RunSimulationAsync();
        }
        catch (OperationCanceledException)
        {
            // User cancelled — navigate back to setup
            CurrentView = _setupView;
            return;
        }
        catch (Exception ex)
        {
            // Show error — stay on run view with error message
            _runViewModel.ShowError(ex.Message);
            return;
        }

        // Navigate to ResultsView with results
        _resultsViewModel.LoadResults(/* SimulationCompleteEventArgs */);
        CurrentView = _resultsView;
    }
}
```

### Dependency Injection / Service Setup

Since we're in an ExcelDna add-in (not an ASP.NET app), DI is simpler. Set up services in `AddIn.AutoOpen()`:

```csharp
public class AddIn : IExcelAddIn
{
    public static IWorkbookManager WorkbookManager { get; private set; }
    public static InputTagManager InputManager { get; private set; }
    public static OutputTagManager OutputManager { get; private set; }
    public static ConfigPersistence ConfigPersistence { get; private set; }
    public static SimulationOrchestrator Orchestrator { get; private set; }

    public void AutoOpen()
    {
        WorkbookManager = new WorkbookManager();
        InputManager = new InputTagManager();
        OutputManager = new OutputTagManager();
        ConfigPersistence = new ConfigPersistence();
        Orchestrator = new SimulationOrchestrator(
            WorkbookManager, InputManager, OutputManager,
            new SimulationEngine(), ConfigPersistence);

        // Load saved config if present
        var profile = ConfigPersistence.Load();
        if (profile != null)
            InputManager.LoadFromProfile(profile);
    }
}
```

If you prefer a proper DI container, use `Microsoft.Extensions.DependencyInjection` — it's lightweight and familiar. But static service properties on the AddIn class are simpler for this scale.

## Integration Test Scenarios

These are manual test scripts to run with the sample workbook:

### Scenario 1: Basic End-to-End
1. Open `SimpleFinancialModel.xlsx`
2. Open the task pane
3. Add 3 input cells (B4, B5, B6) with Normal distributions
4. Add 1 output cell (D10)
5. Set iterations to 1,000
6. Click Run
7. **Verify:** Progress bar advances, iteration counter increments, time estimates update
8. **Verify:** Results view appears with histogram, stats, percentile markers
9. **Verify:** Stats are reasonable (mean near the expected value)

### Scenario 2: Cancellation
1. Set iterations to 50,000
2. Click Run
3. Click Stop after ~2 seconds
4. **Verify:** Simulation stops, navigates back to Setup

### Scenario 3: Config Persistence
1. Configure a simulation (3 inputs, 1 output)
2. Save and close the workbook
3. Reopen the workbook
4. Open the task pane
5. **Verify:** All inputs, outputs, and distributions are restored

### Scenario 4: Multiple Outputs
1. Add 2 output cells
2. Run simulation
3. **Verify:** Can switch between outputs in the results view dropdown
4. **Verify:** Each output has independent stats and histogram

### Scenario 5: Cell Value Restoration
1. Note the original values in input cells
2. Run simulation
3. After completion, check input cells
4. **Verify:** Original values are restored

## File Structure

```
MonteCarlo.Addin/
├── AddIn.cs                         # Updated: service initialization
├── Services/
│   └── SimulationOrchestrator.cs    # Central coordinator
└── Excel/
    └── (existing files updated as needed)

MonteCarlo.UI/
└── ViewModels/
    └── MainViewModel.cs             # Updated: RunSimulation command wired to orchestrator
```

## Commit Strategy

```
feat(addin): add SimulationOrchestrator — end-to-end simulation coordination
feat(addin): implement fast-mode evaluator with Excel recalculation
feat(addin): add original value save/restore around simulation runs
feat(addin): wire service initialization in AddIn.AutoOpen()
feat(ui): connect MainViewModel Run command to orchestrator
feat(addin): add config auto-load on workbook open
```

## Done When

- [ ] Full flow works: Setup → configure inputs/outputs → Run → see progress → see results
- [ ] Fast-mode evaluator correctly reads/writes Excel cells and triggers recalc
- [ ] Screen updating disabled during simulation for performance
- [ ] Original cell values restored after simulation completes or is cancelled
- [ ] Cancellation works cleanly (no orphaned threads, UI returns to setup)
- [ ] Errors display gracefully in the RunView (not unhandled exceptions)
- [ ] Config auto-loads when a workbook with saved config is opened
- [ ] Progress events marshal correctly to the WPF UI thread
- [ ] All 5 integration test scenarios pass manually
- [ ] `dotnet build` clean
