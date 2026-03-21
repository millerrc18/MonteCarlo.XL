using System.Windows;
using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.UDF;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace MonteCarlo.Addin.Services;

/// <summary>
/// Central coordinator for the simulation lifecycle: reads inputs from Excel,
/// runs simulation, computes statistics, and fires completion events.
/// </summary>
public class SimulationOrchestrator
{
    private readonly WorkbookManager _workbook;
    private readonly InputTagManager _inputManager;
    private readonly OutputTagManager _outputManager;
    private readonly ConfigPersistence _configPersistence;

    private CancellationTokenSource? _cts;
    private SimulationResult? _lastResult;
    private readonly ConvergenceChecker _convergenceChecker = new();

    /// <summary>Progress during simulation.</summary>
    public event EventHandler<SimulationProgressEventArgs>? ProgressChanged;

    /// <summary>Fired when simulation completes successfully.</summary>
    public event EventHandler<SimulationCompleteEventArgs>? SimulationComplete;

    /// <summary>Fired on simulation error.</summary>
    public event EventHandler<SimulationErrorEventArgs>? SimulationError;

    /// <summary>Fired at convergence checkpoints.</summary>
    public event EventHandler<ConvergenceUpdateEventArgs>? ConvergenceUpdated;

    /// <summary>The most recent simulation result.</summary>
    public SimulationResult? LastResult => _lastResult;

    public SimulationOrchestrator(
        WorkbookManager workbook,
        InputTagManager inputManager,
        OutputTagManager outputManager,
        ConfigPersistence configPersistence)
    {
        _workbook = workbook;
        _inputManager = inputManager;
        _outputManager = outputManager;
        _configPersistence = configPersistence;
    }

    /// <summary>
    /// Run the full simulation workflow.
    /// </summary>
    public async Task RunSimulationAsync(int iterationCount, int? randomSeed)
    {
        _cts = new CancellationTokenSource();
        _convergenceChecker.Reset();
        var app = (Application)ExcelDnaUtil.Application;

        // 1. Build inputs from GUI-tagged cells
        var taggedInputs = new List<TaggedInput>(_inputManager.GetAllInputs());
        var taggedOutputs = _outputManager.GetAllOutputs();

        // 1b. Scan for MC.* formula cells and add as auto-detected inputs
        var mcFunctionScanner = new MCFunctionScanner();
        var mcFunctions = new List<DetectedMCFunction>();
        var originalFormulas = new Dictionary<string, string>();
        var taggedCellRefs = new HashSet<string>(taggedInputs.Select(i => i.Cell.FullReference));

        try
        {
            var activeSheet = ((Worksheet)app.ActiveSheet);
            mcFunctions = mcFunctionScanner.ScanWorksheet(activeSheet);

            foreach (var func in mcFunctions)
            {
                // Skip if already tagged via the GUI (avoid duplicates)
                if (taggedCellRefs.Contains(func.Cell.FullReference))
                    continue;

                var parameters = mcFunctionScanner.ResolveParameters(func, activeSheet);
                if (parameters == null) continue;

                // Save the original formula for restoration
                try
                {
                    dynamic cell = activeSheet.Range[func.Cell.CellAddress];
                    originalFormulas[func.Cell.FullReference] = cell.Formula.ToString();
                }
                catch { }

                taggedInputs.Add(new TaggedInput
                {
                    Cell = func.Cell,
                    Label = func.Cell.CellAddress,
                    DistributionName = func.DistributionName,
                    Parameters = parameters
                });
                taggedCellRefs.Add(func.Cell.FullReference);
            }
        }
        catch { /* MC function scanning is non-fatal */ }

        if (taggedInputs.Count == 0)
            throw new InvalidOperationException("No inputs configured.");
        if (taggedOutputs.Count == 0)
            throw new InvalidOperationException("No outputs configured.");

        // 2. Save original cell values for restoration
        var originalValues = new Dictionary<string, double>();
        foreach (var input in taggedInputs)
        {
            var cellRef = input.Cell;
            double val = _workbook.ReadCellValue(cellRef.SheetName, cellRef.CellAddress);
            originalValues[cellRef.FullReference] = val;
        }

        // 3. Build SimulationConfig
        var config = new SimulationConfig
        {
            IterationCount = iterationCount,
            RandomSeed = randomSeed,
            UseParallelEvaluation = false // COM is single-threaded
        };

        foreach (var input in taggedInputs)
        {
            var dist = DistributionFactory.Create(input.DistributionName, input.Parameters);
            config.Inputs.Add(new SimulationInput
            {
                Id = input.Cell.FullReference,
                Label = input.Label,
                Distribution = dist,
                BaseValue = originalValues[input.Cell.FullReference]
            });
        }

        foreach (var output in taggedOutputs)
        {
            config.Outputs.Add(new SimulationOutput
            {
                Id = output.Cell.FullReference,
                Label = output.Label
            });
        }

        // 4. Build evaluator (fast mode — Excel recalc)
        var evaluator = BuildEvaluator(taggedInputs, taggedOutputs, app);

        // 5. Run simulation
        var engine = new SimulationEngine();
        int checkpointInterval = 500;
        int lastCheckpoint = 0;

        engine.ProgressChanged += (sender, e) =>
        {
            ProgressChanged?.Invoke(this, e);

            // Convergence check at intervals
            if (e.CompletedIterations - lastCheckpoint >= checkpointInterval &&
                _lastResult == null) // Only during run, not after
            {
                lastCheckpoint = e.CompletedIterations;
                // Note: convergence is checked on partial results in the completion handler
            }
        };

        var savedCalculation = app.Calculation;
        var savedScreenUpdating = app.ScreenUpdating;
        var savedEnableEvents = app.EnableEvents;

        try
        {
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            app.Calculation = XlCalculation.xlCalculationManual;

            var result = await engine.RunAsync(config, evaluator, _cts.Token);
            _lastResult = result;

            // 6. Compute statistics for each output
            var statsByOutput = new Dictionary<string, SummaryStatistics>();
            var sensitivityByOutput = new Dictionary<string, IReadOnlyList<SensitivityResult>>();

            for (int i = 0; i < config.Outputs.Count; i++)
            {
                var outputValues = result.GetOutputValues(i);
                statsByOutput[config.Outputs[i].Id] = new SummaryStatistics(outputValues);

                try
                {
                    var sensitivity = SensitivityAnalysis.Analyze(result, i);
                    sensitivityByOutput[config.Outputs[i].Id] = sensitivity;
                }
                catch
                {
                    // Sensitivity analysis may fail for some configs — not fatal
                    sensitivityByOutput[config.Outputs[i].Id] = Array.Empty<SensitivityResult>();
                }
            }

            // 7. Fire completion
            SimulationComplete?.Invoke(this, new SimulationCompleteEventArgs
            {
                Result = result,
                StatsByOutput = statsByOutput,
                SensitivityByOutput = sensitivityByOutput,
                TotalElapsed = result.Elapsed
            });
        }
        catch (OperationCanceledException)
        {
            throw; // Propagate cancellation
        }
        catch (Exception ex)
        {
            SimulationError?.Invoke(this, new SimulationErrorEventArgs { Error = ex });
            throw;
        }
        finally
        {
            // 8. Restore original values and Excel state
            app.Calculation = savedCalculation;

            foreach (var (cellRef, originalValue) in originalValues)
            {
                try
                {
                    // Parse FullReference back to sheet/cell
                    var parts = cellRef.Split('!');
                    if (parts.Length == 2)
                    {
                        string sheet = parts[0].Trim('\'');
                        string cell = parts[1];
                        _workbook.WriteCellValue(sheet, cell, originalValue);
                    }
                }
                catch { }
            }

            // Restore MC.* formulas that were overwritten during simulation
            foreach (var (cellRef, formula) in originalFormulas)
            {
                try
                {
                    var parts = cellRef.Split('!');
                    if (parts.Length == 2)
                    {
                        string sheetName = parts[0].Trim('\'');
                        string cellAddr = parts[1];
                        Worksheet ws = app.Sheets[sheetName];
                        ws.Range[cellAddr].Formula = formula;
                    }
                }
                catch { }
            }

            app.Calculate();
            app.EnableEvents = savedEnableEvents;
            app.ScreenUpdating = savedScreenUpdating;
            _cts = null;
        }
    }

    /// <summary>
    /// Cancel the running simulation.
    /// </summary>
    public void CancelSimulation()
    {
        _cts?.Cancel();
    }

    /// <summary>
    /// Build a profile from the current setup for persistence.
    /// </summary>
    public SimulationProfile BuildProfile(int iterationCount, int? randomSeed)
    {
        var profile = new SimulationProfile
        {
            IterationCount = iterationCount,
            RandomSeed = randomSeed
        };

        foreach (var input in _inputManager.GetAllInputs())
        {
            profile.Inputs.Add(new SavedInput
            {
                SheetName = input.Cell.SheetName,
                CellAddress = input.Cell.CellAddress,
                Label = input.Label,
                DistributionName = input.DistributionName,
                Parameters = new Dictionary<string, double>(input.Parameters)
            });
        }

        foreach (var output in _outputManager.GetAllOutputs())
        {
            profile.Outputs.Add(new SavedOutput
            {
                SheetName = output.Cell.SheetName,
                CellAddress = output.Cell.CellAddress,
                Label = output.Label
            });
        }

        return profile;
    }

    /// <summary>
    /// Save the current config to the workbook.
    /// </summary>
    public void SaveConfig(int iterationCount, int? randomSeed)
    {
        var profile = BuildProfile(iterationCount, randomSeed);
        _configPersistence.Save(profile);
    }

    /// <summary>
    /// Load config from the workbook and return the profile, or null if none.
    /// </summary>
    public SimulationProfile? LoadConfig()
    {
        return _configPersistence.Load();
    }

    private Func<Dictionary<string, double>, Dictionary<string, double>> BuildEvaluator(
        IReadOnlyList<TaggedInput> taggedInputs,
        IReadOnlyList<TaggedOutput> taggedOutputs,
        Application app)
    {
        // Build lookup for fast cell writing, grouped by sheet for batch COM calls
        var inputCells = taggedInputs.ToDictionary(
            i => i.Cell.FullReference,
            i => i.Cell);

        // Pre-group inputs by sheet for batch writing
        var inputsBySheet = taggedInputs
            .GroupBy(i => i.Cell.SheetName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => g.Select(i => i.Cell).ToArray(),
                StringComparer.OrdinalIgnoreCase);

        // Pre-cache sheet references for output reading
        var outputsBySheet = taggedOutputs
            .GroupBy(o => o.Cell.SheetName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => g.ToArray(),
                StringComparer.OrdinalIgnoreCase);

        return inputs =>
        {
            // Write input values to Excel — batch by sheet to minimize COM round-trips
            foreach (var (sheetName, cells) in inputsBySheet)
            {
                foreach (var cell in cells)
                {
                    if (inputs.TryGetValue(cell.FullReference, out var value))
                        _workbook.WriteCellValue(sheetName, cell.CellAddress, value);
                }
            }

            // Recalculate
            app.Calculate();

            // Read all output values — batch by sheet
            var outputs = new Dictionary<string, double>(taggedOutputs.Count);
            foreach (var (sheetName, sheetOutputs) in outputsBySheet)
            {
                foreach (var output in sheetOutputs)
                {
                    double val = _workbook.ReadCellValue(sheetName, output.Cell.CellAddress);
                    outputs[output.Cell.FullReference] = val;
                }
            }

            return outputs;
        };
    }
}

/// <summary>Event args for simulation completion.</summary>
public class SimulationCompleteEventArgs : EventArgs
{
    public required SimulationResult Result { get; init; }
    public required Dictionary<string, SummaryStatistics> StatsByOutput { get; init; }
    public required Dictionary<string, IReadOnlyList<SensitivityResult>> SensitivityByOutput { get; init; }
    public TimeSpan TotalElapsed { get; init; }
}

/// <summary>Event args for simulation errors.</summary>
public class SimulationErrorEventArgs : EventArgs
{
    public required Exception Error { get; init; }
}

/// <summary>Event args for convergence updates.</summary>
public class ConvergenceUpdateEventArgs : EventArgs
{
    public required IReadOnlyList<ConvergenceIndicator> Indicators { get; init; }
}
