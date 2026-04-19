using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.Export;
using MonteCarlo.Addin.Services;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Addin.UDF;
using MonteCarlo.UI.Services;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.Addin.TaskPane;

/// <summary>
/// Bridges the WPF task pane view models to Excel-specific add-in services.
/// </summary>
internal sealed class TaskPaneIntegration : IDisposable
{
    private readonly TaskPaneController _taskPane;
    private readonly InputTagManager _inputManager;
    private readonly OutputTagManager _outputManager;
    private readonly CellHighlighter _highlighter;
    private readonly SimulationOrchestrator _orchestrator;

    private MainViewModel? _viewModel;
    private Application? _excelApp;
    private AppEvents_SheetSelectionChangeEventHandler? _selectionHandler;
    private string? _selectionMode;
    private bool _disposed;
    private bool _isRunning;
    private bool _errorRaisedDuringRun;
    private bool _profileLoaded;

    public TaskPaneIntegration(
        TaskPaneController taskPane,
        InputTagManager inputManager,
        OutputTagManager outputManager,
        CellHighlighter highlighter,
        SimulationOrchestrator orchestrator)
    {
        _taskPane = taskPane;
        _inputManager = inputManager;
        _outputManager = outputManager;
        _highlighter = highlighter;
        _orchestrator = orchestrator;

        _taskPane.Created += OnTaskPaneCreated;
        _orchestrator.ProgressChanged += OnProgressChanged;
        _orchestrator.SimulationComplete += OnSimulationComplete;
        _orchestrator.SimulationError += OnSimulationError;
        _orchestrator.ConvergenceUpdated += OnConvergenceUpdated;

        EnsureWired();
    }

    /// <summary>
    /// Ensures the task pane view model is connected after the pane is created.
    /// </summary>
    public void EnsureWired()
    {
        if (_disposed || _viewModel != null)
            return;

        var viewModel = _taskPane.ViewModel;
        if (viewModel == null)
            return;

        _viewModel = viewModel;
        viewModel.RunSimulationRequested += OnRunSimulationRequested;
        viewModel.CancelSimulationRequested += OnCancelSimulationRequested;
        viewModel.SetupViewModel.CellSelectionRequested += OnCellSelectionRequested;
        viewModel.SetupViewModel.InputAdded += OnInputAdded;
        viewModel.SetupViewModel.OutputAdded += OnOutputAdded;
        viewModel.SetupViewModel.InputRemoved += OnInputRemoved;
        viewModel.SetupViewModel.OutputRemoved += OnOutputRemoved;

        LoadSavedProfile(viewModel);
    }

    /// <summary>
    /// Starts a simulation from the ribbon or keyboard shortcut.
    /// </summary>
    public void RequestRunFromRibbon()
    {
        EnsureWired();
        Dispatch(() => _viewModel?.RequestSimulationRun());
    }

    /// <summary>
    /// Opens the setup view from host UI such as the Excel ribbon.
    /// </summary>
    public void ShowSetup()
    {
        EnsureWired();
        Dispatch(() => _viewModel?.ShowSetupView());
    }

    /// <summary>
    /// Opens the results view from host UI such as the Excel ribbon.
    /// </summary>
    public void ShowResults()
    {
        EnsureWired();
        Dispatch(() => _viewModel?.ShowResultsView());
    }

    /// <summary>
    /// Opens the settings view from host UI such as the Excel ribbon.
    /// </summary>
    public void ShowSettings()
    {
        EnsureWired();
        Dispatch(() => _viewModel?.ShowSettingsView());
    }

    /// <summary>
    /// Opens setup and starts the add-input flow.
    /// </summary>
    public void BeginAddInput()
    {
        EnsureWired();
        Dispatch(() =>
        {
            _viewModel?.ShowSetupView();
            _viewModel?.SetupViewModel.BeginAddInput();
        });
    }

    /// <summary>
    /// Opens setup and starts the add-output flow.
    /// </summary>
    public void BeginAddOutput()
    {
        EnsureWired();
        Dispatch(() =>
        {
            _viewModel?.ShowSetupView();
            _viewModel?.SetupViewModel.BeginAddOutput();
        });
    }

    /// <summary>
    /// Sets the simulation iteration count from a ribbon preset.
    /// </summary>
    public void SetIterationCount(int iterationCount)
    {
        EnsureWired();
        Dispatch(() =>
        {
            if (_viewModel == null)
                return;

            _viewModel.SetupViewModel.IterationCount = iterationCount;
            _viewModel.ShowSetupView();
            SaveCurrentProfile();
        });
    }

    /// <summary>
    /// Copies current results statistics to the clipboard, if results are available.
    /// </summary>
    public void CopyStatsToClipboard()
    {
        EnsureWired();
        Dispatch(() =>
        {
            if (_viewModel?.ResultsViewModel.CopyStatsToClipboardCommand.CanExecute(null) == true)
            {
                _viewModel.ShowResultsView();
                _viewModel.ResultsViewModel.CopyStatsToClipboardCommand.Execute(null);
            }
        });
    }

    /// <summary>
    /// Exports the current results summary to a worksheet, if results are available.
    /// </summary>
    public void ExportCurrentSummary()
    {
        ExportCurrentResults(exportRawData: false);
    }

    /// <summary>
    /// Exports the current results raw data to a worksheet, if results are available.
    /// </summary>
    public void ExportCurrentRawData()
    {
        ExportCurrentResults(exportRawData: true);
    }

    private void OnTaskPaneCreated(object? sender, EventArgs e)
    {
        EnsureWired();
    }

    private async void OnRunSimulationRequested(object? sender, EventArgs e)
    {
        EnsureWired();
        var viewModel = _viewModel;
        if (viewModel == null || _isRunning)
            return;

        _isRunning = true;
        _errorRaisedDuringRun = false;

        try
        {
            SyncManagersFromSetup(viewModel.SetupViewModel, clearExistingHighlights: true);
            SaveCurrentProfile();
            await _orchestrator.RunSimulationAsync(
                viewModel.SetupViewModel.IterationCount,
                viewModel.SetupViewModel.RandomSeed);
        }
        catch (OperationCanceledException)
        {
            Dispatch(() => viewModel.OnSimulationCancelled());
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Simulation failed before the orchestrator error event.", ex);
            if (!_errorRaisedDuringRun)
                Dispatch(() => viewModel.OnSimulationError(ex));
        }
        finally
        {
            _isRunning = false;
            _errorRaisedDuringRun = false;
        }
    }

    private void OnCancelSimulationRequested(object? sender, EventArgs e)
    {
        _orchestrator.CancelSimulation();
    }

    private void OnProgressChanged(object? sender, MonteCarlo.Engine.Simulation.SimulationProgressEventArgs e)
    {
        Dispatch(() => _viewModel?.OnSimulationProgress(e));
    }

    private void OnSimulationComplete(object? sender, SimulationCompleteEventArgs e)
    {
        Dispatch(() => _viewModel?.OnSimulationComplete(e.Result, e.SensitivityByOutput));
    }

    private void OnSimulationError(object? sender, SimulationErrorEventArgs e)
    {
        _errorRaisedDuringRun = true;
        StartupDiagnostics.LogException("Simulation orchestrator error.", e.Error);
        Dispatch(() => _viewModel?.OnSimulationError(e.Error));
    }

    private void OnConvergenceUpdated(object? sender, ConvergenceUpdateEventArgs e)
    {
        Dispatch(() => _viewModel?.OnConvergenceUpdate(e.Indicators));
    }

    private void OnCellSelectionRequested(object? sender, CellSelectionRequestedEventArgs e)
    {
        StartCellSelection(e.Mode);
    }

    private void OnInputAdded(object? sender, InputAddedEventArgs e)
    {
        var cell = ToCellReference(e.Input);
        _inputManager.TagInput(
            cell,
            e.Input.Label,
            e.Input.DistributionName,
            new Dictionary<string, double>(e.Input.Parameters));
        _highlighter.HighlightInput(cell);
        SaveCurrentProfile();
    }

    private void OnOutputAdded(object? sender, OutputAddedEventArgs e)
    {
        var cell = ToCellReference(e.Output);
        _outputManager.TagOutput(cell, e.Output.Label);
        _highlighter.HighlightOutput(cell);
        SaveCurrentProfile();
    }

    private void OnInputRemoved(object? sender, InputRemovedEventArgs e)
    {
        var cell = ToCellReference(e.Input);
        _inputManager.UntagInput(cell);
        _highlighter.ClearHighlight(cell);
        SaveCurrentProfile();
    }

    private void OnOutputRemoved(object? sender, OutputRemovedEventArgs e)
    {
        var cell = ToCellReference(e.Output);
        _outputManager.UntagOutput(cell);
        _highlighter.ClearHighlight(cell);
        SaveCurrentProfile();
    }

    private void StartCellSelection(string mode)
    {
        StopCellSelection();

        try
        {
            _excelApp = (Application)ExcelDnaUtil.Application;
            _selectionMode = mode;
            _selectionHandler = OnSheetSelectionChange;
            _excelApp.SheetSelectionChange += _selectionHandler;
            _excelApp.StatusBar = $"MonteCarlo.XL: click a worksheet cell to use as a simulation {mode}.";
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to start cell selection.", ex);
            StopCellSelection();
            Dispatch(() => _viewModel?.OnSimulationError(ex));
        }
    }

    private void StopCellSelection()
    {
        try
        {
            if (_excelApp != null && _selectionHandler != null)
                _excelApp.SheetSelectionChange -= _selectionHandler;

            if (_excelApp != null)
                _excelApp.StatusBar = false;
        }
        catch
        {
            // Excel may already be shutting down.
        }
        finally
        {
            _selectionMode = null;
            _selectionHandler = null;
            _excelApp = null;
        }
    }

    private void OnSheetSelectionChange(object sheetObject, Range target)
    {
        if (_selectionMode == null || _viewModel == null)
            return;

        try
        {
            var cell = (Range)target.Cells[1, 1];
            var worksheet = sheetObject as Worksheet ?? (Worksheet)cell.Worksheet;
            var address = cell.Address[RowAbsolute: false, ColumnAbsolute: false];
            var label = GetSuggestedLabel(worksheet, cell) ?? address;

            // Auto-detect MC.* distribution formula so the setup form pre-populates.
            string? detectedDist = null;
            Dictionary<string, double>? detectedParams = null;
            try
            {
                if (cell.HasFormula)
                {
                    string formula = cell.Formula?.ToString() ?? string.Empty;
                    if (formula.StartsWith("=MC.", StringComparison.OrdinalIgnoreCase))
                    {
                        var scanner = new MCFunctionScanner();
                        var detected = new DetectedMCFunction
                        {
                            Cell = new Addin.Excel.CellReference { SheetName = worksheet.Name, CellAddress = address },
                            DistributionName = "",
                            ParameterNames = Array.Empty<string>(),
                            Formula = formula
                        };
                        // Reparse via scanner to fill distribution/param names
                        foreach (var fn in scanner.ScanWorksheet(worksheet))
                        {
                            if (string.Equals(fn.Cell.CellAddress, address, StringComparison.OrdinalIgnoreCase))
                            {
                                detected = fn;
                                break;
                            }
                        }
                        if (!string.IsNullOrEmpty(detected.DistributionName))
                        {
                            detectedDist = detected.DistributionName;
                            detectedParams = scanner.ResolveParameters(detected, worksheet);
                        }
                    }
                }
            }
            catch
            {
                // Formula detection is best-effort; fall back to manual entry.
            }

            Dispatch(() => _viewModel?.SetupViewModel.OnCellSelected(
                worksheet.Name, address, label, detectedDist, detectedParams));
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to complete cell selection.", ex);
            Dispatch(() => _viewModel?.OnSimulationError(ex));
        }
        finally
        {
            StopCellSelection();
        }
    }

    private void LoadSavedProfile(MainViewModel viewModel)
    {
        if (_profileLoaded)
            return;

        _profileLoaded = true;

        try
        {
            var profile = _orchestrator.LoadConfig();
            if (profile == null)
                return;

            var setup = viewModel.SetupViewModel;
            setup.Inputs.Clear();
            setup.Outputs.Clear();
            setup.IterationCount = profile.IterationCount;
            setup.RandomSeed = profile.RandomSeed;
            setup.IsSeedLocked = profile.RandomSeed.HasValue;
            setup.CorrelationMatrixValues = profile.CorrelationMatrix;

            foreach (var input in profile.Inputs)
            {
                var parameters = new Dictionary<string, double>(input.Parameters);
                var distribution = DistributionFactory.Create(input.DistributionName, parameters);
                setup.Inputs.Add(new InputCardViewModel(
                    input.SheetName,
                    input.CellAddress,
                    input.Label,
                    input.DistributionName,
                    parameters,
                    distribution));
            }

            foreach (var output in profile.Outputs)
            {
                setup.Outputs.Add(new OutputCardViewModel(
                    output.SheetName,
                    output.CellAddress,
                    output.Label));
            }

            SyncManagersFromSetup(setup, clearExistingHighlights: false);
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to load saved workbook profile.", ex);
        }
    }

    private void SyncManagersFromSetup(SetupViewModel setup, bool clearExistingHighlights)
    {
        if (clearExistingHighlights)
        {
            foreach (var input in _inputManager.GetAllInputs())
                _highlighter.ClearHighlight(input.Cell);

            foreach (var output in _outputManager.GetAllOutputs())
                _highlighter.ClearHighlight(output.Cell);
        }

        _inputManager.Clear();
        _outputManager.Clear();

        foreach (var input in setup.Inputs)
        {
            _inputManager.TagInput(
                ToCellReference(input),
                input.Label,
                input.DistributionName,
                new Dictionary<string, double>(input.Parameters));
        }

        foreach (var output in setup.Outputs)
            _outputManager.TagOutput(ToCellReference(output), output.Label);

        _highlighter.RefreshAll(_inputManager, _outputManager);
    }

    private void SaveCurrentProfile()
    {
        try
        {
            var setup = _viewModel?.SetupViewModel;
            if (setup == null)
                return;

            _orchestrator.SaveConfig(setup.IterationCount, setup.RandomSeed);
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to save workbook profile.", ex);
        }
    }

    private void ExportCurrentResults(bool exportRawData)
    {
        EnsureWired();
        Dispatch(() =>
        {
            try
            {
                var viewModel = _viewModel;
                var resultsViewModel = viewModel?.ResultsViewModel;
                var result = resultsViewModel?.SimulationResult;
                var selectedOutputId = resultsViewModel?.SelectedOutputId;

                if (viewModel == null ||
                    resultsViewModel == null ||
                    result == null ||
                    string.IsNullOrWhiteSpace(selectedOutputId))
                {
                    return;
                }

                var outputIndex = -1;
                for (var i = 0; i < result.Config.Outputs.Count; i++)
                {
                    if (result.Config.Outputs[i].Id == selectedOutputId)
                    {
                        outputIndex = i;
                        break;
                    }
                }

                if (outputIndex < 0)
                    return;

                var exporter = new ResultsExporter();
                var userSettings = new UserSettingsService().Load();
                if (exportRawData)
                {
                    exporter.ExportRawData(
                        result,
                        outputIndex,
                        userSettings.CreateNewWorksheetForExports);
                    return;
                }

                var stats = resultsViewModel.CurrentStats;
                if (stats == null)
                    return;

                var profile = _orchestrator.LoadConfig() ??
                              _orchestrator.BuildProfile(
                                  viewModel.SetupViewModel.IterationCount,
                                  viewModel.SetupViewModel.RandomSeed);

                // Look up sensitivity data for the selected output
                IReadOnlyList<SensitivityResult>? sensitivity = null;
                if (resultsViewModel.SensitivityResults != null && selectedOutputId != null)
                    resultsViewModel.SensitivityResults.TryGetValue(selectedOutputId, out sensitivity);

                exporter.ExportSummary(
                    result,
                    stats,
                    sensitivity,
                    profile,
                    outputIndex,
                    createNewSheet: userSettings.CreateNewWorksheetForExports);

                StartupDiagnostics.Log($"Export summary completed for output '{selectedOutputId}'.");
            }
            catch (Exception ex)
            {
                StartupDiagnostics.LogException("Failed to export simulation results.", ex);
                _viewModel?.OnSimulationError(ex);
            }
        });
    }

    private void Dispatch(System.Action action)
    {
        var dispatcher = _taskPane.Dispatcher ?? System.Windows.Application.Current?.Dispatcher;
        if (dispatcher == null || dispatcher.CheckAccess())
        {
            action();
            return;
        }

        dispatcher.BeginInvoke(action);
    }

    private static CellReference ToCellReference(InputCardViewModel input) => new()
    {
        SheetName = input.SheetName,
        CellAddress = input.CellAddress
    };

    private static CellReference ToCellReference(OutputCardViewModel output) => new()
    {
        SheetName = output.SheetName,
        CellAddress = output.CellAddress
    };

    private static string? GetSuggestedLabel(Worksheet worksheet, Range cell)
    {
        var left = TryReadCellText(worksheet, cell.Row, cell.Column - 1);
        if (!string.IsNullOrWhiteSpace(left))
            return left;

        var above = TryReadCellText(worksheet, cell.Row - 1, cell.Column);
        if (!string.IsNullOrWhiteSpace(above))
            return above;

        return null;
    }

    private static string? TryReadCellText(Worksheet worksheet, int row, int column)
    {
        if (row < 1 || column < 1)
            return null;

        try
        {
            var range = (Range)worksheet.Cells[row, column];
            return range.Value2?.ToString();
        }
        catch
        {
            return null;
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        StopCellSelection();

        _taskPane.Created -= OnTaskPaneCreated;
        _orchestrator.ProgressChanged -= OnProgressChanged;
        _orchestrator.SimulationComplete -= OnSimulationComplete;
        _orchestrator.SimulationError -= OnSimulationError;
        _orchestrator.ConvergenceUpdated -= OnConvergenceUpdated;

        if (_viewModel != null)
        {
            _viewModel.RunSimulationRequested -= OnRunSimulationRequested;
            _viewModel.CancelSimulationRequested -= OnCancelSimulationRequested;
            _viewModel.SetupViewModel.CellSelectionRequested -= OnCellSelectionRequested;
            _viewModel.SetupViewModel.InputAdded -= OnInputAdded;
            _viewModel.SetupViewModel.OutputAdded -= OnOutputAdded;
            _viewModel.SetupViewModel.InputRemoved -= OnInputRemoved;
            _viewModel.SetupViewModel.OutputRemoved -= OnOutputRemoved;
        }

        _viewModel = null;
    }
}
