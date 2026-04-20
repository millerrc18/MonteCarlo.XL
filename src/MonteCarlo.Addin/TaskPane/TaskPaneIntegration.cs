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
    private const int ExcelMaxRows = 1_048_576;
    private const long RawDataExportWarningCellThreshold = 500_000;

    private readonly TaskPaneController _taskPane;
    private readonly InputTagManager _inputManager;
    private readonly OutputTagManager _outputManager;
    private readonly CellHighlighter _highlighter;
    private readonly SimulationOrchestrator _orchestrator;

    private MainViewModel? _viewModel;
    private Application? _excelApp;
    private ExcelStateScope? _selectionState;
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
        viewModel.PreflightReportProvider = profile =>
            ExcelModelPreflightValidator.Validate(profile, TryGetExcelApplication());
        viewModel.PauseOnPreflightWarningsProvider = () =>
            new UserSettingsService().Load().PauseOnPreflightWarnings;
        viewModel.RunSimulationRequested += OnRunSimulationRequested;
        viewModel.CancelSimulationRequested += OnCancelSimulationRequested;
        viewModel.CorrelationImportRequested += OnCorrelationImportRequested;
        viewModel.CorrelationExportRequested += OnCorrelationExportRequested;
        viewModel.CorrelationMatrixChanged += OnCorrelationMatrixChanged;
        viewModel.SetupViewModel.CellSelectionRequested += OnCellSelectionRequested;
        viewModel.SetupViewModel.InputAdded += OnInputAdded;
        viewModel.SetupViewModel.OutputAdded += OnOutputAdded;
        viewModel.SetupViewModel.InputRemoved += OnInputRemoved;
        viewModel.SetupViewModel.OutputRemoved += OnOutputRemoved;
        viewModel.SetupViewModel.CellActionRequested += OnCellActionRequested;
        viewModel.SetupViewModel.RefreshHighlightsRequested += OnRefreshHighlightsRequested;
        viewModel.SetupViewModel.DistributionFitRequested += OnDistributionFitRequested;

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
    /// Opens the input correlation editor from host UI such as the Excel ribbon.
    /// </summary>
    public void ShowCorrelations()
    {
        EnsureWired();
        Dispatch(() => _viewModel?.ShowCorrelationView());
    }

    /// <summary>
    /// Opens the model preflight validation view.
    /// </summary>
    public void ShowPreflight()
    {
        EnsureWired();
        Dispatch(() => _viewModel?.ShowPreflightView());
    }

    private Application? TryGetExcelApplication()
    {
        try
        {
            _excelApp ??= (Application)ExcelDnaUtil.Application;
            return _excelApp;
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to get Excel application for model preflight.", ex);
            return null;
        }
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
            var userSettings = new UserSettingsService().Load();
            var runSeed = viewModel.SetupViewModel.IsSeedLocked
                ? viewModel.SetupViewModel.RandomSeed
                : userSettings.SeedMode == SeedMode.Fixed
                    ? userSettings.FixedRandomSeed
                    : null;

            SyncManagersFromSetup(viewModel.SetupViewModel, clearExistingHighlights: true);
            SaveCurrentProfile();
            await _orchestrator.RunSimulationAsync(
                viewModel.SetupViewModel.IterationCount,
                runSeed,
                userSettings.SamplingMethod,
                userSettings.AutoStopOnConvergence,
                viewModel.SetupViewModel.CorrelationMatrixValues);
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

    private void OnCorrelationMatrixChanged(object? sender, EventArgs e)
    {
        SaveCurrentProfile();
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

    private void OnCellActionRequested(object? sender, SetupCellActionEventArgs e)
    {
        var cell = new Addin.Excel.CellReference
        {
            SheetName = e.SheetName,
            CellAddress = e.CellAddress
        };

        try
        {
            if (e.Action == SetupCellAction.Jump)
            {
                SelectCell(cell);
                return;
            }

            if (e.Role == SetupCellRole.Input)
                _highlighter.HighlightInput(cell);
            else
                _highlighter.HighlightOutput(cell);
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to complete setup cell action.", ex);
            Dispatch(() => _viewModel?.OnSimulationError(ex));
        }
    }

    private void OnRefreshHighlightsRequested(object? sender, EventArgs e)
    {
        try
        {
            if (_viewModel != null)
                SyncManagersFromSetup(_viewModel.SetupViewModel, clearExistingHighlights: false);
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to refresh input/output highlights.", ex);
            Dispatch(() => _viewModel?.OnSimulationError(ex));
        }
    }

    private void OnDistributionFitRequested()
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            if (app.Selection is not Range selection)
                throw new InvalidOperationException("Select a numeric Excel range before fitting a distribution.");

            var values = ReadNumericValues(selection);
            if (values.Length < 3)
                throw new InvalidOperationException("Select at least three numeric cells before fitting a distribution.");

            var worksheet = (Worksheet)selection.Worksheet;
            var address = $"{worksheet.Name}!{selection.Address[RowAbsolute: false, ColumnAbsolute: false]}";
            var results = DistributionFitService.Fit(values);

            Dispatch(() => _viewModel?.SetupViewModel.LoadDistributionFitResults(results, address));
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Distribution fitting failed.", ex);
            Dispatch(() => _viewModel?.SetupViewModel.ShowDistributionFitError(ex.Message));
        }
    }

    private void OnCorrelationImportRequested(CorrelationViewModel editor)
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            if (app.Selection is not Range selection)
                throw new InvalidOperationException("Select the top-left cell of a numeric correlation matrix in Excel.");

            var inputCount = editor.InputLabels.Count;
            if (inputCount < 2)
                throw new InvalidOperationException("At least two inputs are required before importing correlations.");

            var source = GetTopLeftRange(selection).Resize[inputCount, inputCount];
            var matrix = ReadCorrelationMatrix(source, inputCount);
            var address = GetRangeAddress(source);

            Dispatch(() => editor.LoadImportedMatrix(matrix, address));
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Correlation matrix import failed.", ex);
            Dispatch(() => editor.ShowRangeError(ex.Message));
        }
    }

    private void OnCorrelationExportRequested(CorrelationViewModel editor)
    {
        try
        {
            if (editor.MatrixValues == null)
                throw new InvalidOperationException("No correlation matrix is available to export.");

            var app = (Application)ExcelDnaUtil.Application;
            if (app.Selection is not Range selection)
                throw new InvalidOperationException("Select the top-left destination cell for the correlation matrix.");

            var target = GetTopLeftRange(selection).Resize[editor.MatrixValues.GetLength(0), editor.MatrixValues.GetLength(1)];
            WriteCorrelationMatrix(target, editor.MatrixValues);
            var address = GetRangeAddress(target);

            Dispatch(() => editor.MarkExported(address));
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Correlation matrix export failed.", ex);
            Dispatch(() => editor.ShowRangeError(ex.Message));
        }
    }

    private static double[] ReadNumericValues(Range range)
    {
        var values = new List<double>();
        foreach (Range cell in range.Cells)
        {
            object? rawValue = cell.Value2;
            if (TryConvertToDouble(rawValue, out double value) && double.IsFinite(value))
                values.Add(value);
        }

        return values.ToArray();
    }

    private static double[,] ReadCorrelationMatrix(Range range, int size)
    {
        var matrix = new double[size, size];
        for (var row = 0; row < size; row++)
        {
            for (var col = 0; col < size; col++)
            {
                var cell = (Range)range.Cells[row + 1, col + 1];
                object? rawValue = cell.Value2;
                if (!TryConvertToDouble(rawValue, out double value) || !double.IsFinite(value))
                {
                    var address = cell.Address[RowAbsolute: false, ColumnAbsolute: false];
                    throw new InvalidOperationException($"Cell {address} does not contain a numeric correlation value.");
                }

                if (value < -1 || value > 1)
                {
                    var address = cell.Address[RowAbsolute: false, ColumnAbsolute: false];
                    throw new InvalidOperationException($"Cell {address} is outside the valid correlation range of -1 to 1.");
                }

                matrix[row, col] = value;
            }
        }

        return matrix;
    }

    private static void WriteCorrelationMatrix(Range target, double[,] matrix)
    {
        target.Value2 = matrix;
        target.NumberFormat = "0.00";
        target.Columns.AutoFit();
    }

    private static Range GetTopLeftRange(Range range) => (Range)range.Cells[1, 1];

    private static string GetRangeAddress(Range range)
    {
        var worksheet = (Worksheet)range.Worksheet;
        return $"{worksheet.Name}!{range.Address[RowAbsolute: false, ColumnAbsolute: false]}";
    }

    private static bool TryConvertToDouble(object? rawValue, out double value)
    {
        switch (rawValue)
        {
            case double doubleValue:
                value = doubleValue;
                return true;
            case int intValue:
                value = intValue;
                return true;
            case decimal decimalValue:
                value = (double)decimalValue;
                return true;
            case string stringValue:
                return double.TryParse(
                    stringValue,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out value);
            default:
                value = 0;
                return false;
        }
    }

    private void StartCellSelection(string mode)
    {
        StopCellSelection();

        try
        {
            _excelApp = (Application)ExcelDnaUtil.Application;
            _selectionState = ExcelStateScope.Capture(_excelApp, "Cell selection");
            _selectionMode = mode;
            _selectionHandler = OnSheetSelectionChange;
            _excelApp.SheetSelectionChange += _selectionHandler;
            _selectionState.Apply(
                statusBar: $"MonteCarlo.XL: click a worksheet cell to use as a simulation {mode}.");
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to start cell selection.", ex);
            StopCellSelection();
            Dispatch(() => _viewModel?.OnSimulationError(ex));
        }
    }

    private void SelectCell(Addin.Excel.CellReference cell)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook is available.");
        var sheet = (Worksheet)workbook.Sheets[cell.SheetName];
        var range = (Range)sheet.Range[cell.CellAddress];

        sheet.Activate();
        app.Goto(range, true);
        app.StatusBar = $"MonteCarlo.XL: selected {cell.SheetName}!{cell.CellAddress}.";
    }

    private void StopCellSelection()
    {
        try
        {
            if (_excelApp != null && _selectionHandler != null)
                _excelApp.SheetSelectionChange -= _selectionHandler;

            _selectionState?.Dispose();
        }
        catch
        {
            // Excel may already be shutting down.
        }
        finally
        {
            _selectionMode = null;
            _selectionState = null;
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
            var setup = viewModel.SetupViewModel;
            if (profile == null)
            {
                ApplyUserDefaults(setup);
                return;
            }

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

    private static void ApplyUserDefaults(SetupViewModel setup)
    {
        var settings = new UserSettingsService().Load();
        setup.IterationCount = settings.DefaultIterationCount;

        if (settings.SeedMode == SeedMode.Fixed)
        {
            setup.RandomSeed = settings.FixedRandomSeed;
            setup.IsSeedLocked = true;
        }
        else
        {
            setup.RandomSeed = null;
            setup.IsSeedLocked = false;
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

            _orchestrator.SaveConfig(setup.IterationCount, setup.RandomSeed, setup.CorrelationMatrixValues);
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
                    if (!ConfirmRawDataExport(result))
                        return;

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
                                  viewModel.SetupViewModel.RandomSeed,
                                  viewModel.SetupViewModel.CorrelationMatrixValues);

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
                    createNewSheet: userSettings.CreateNewWorksheetForExports,
                    percentiles: userSettings.GetDefaultPercentileFractions());

                StartupDiagnostics.Log($"Export summary completed for output '{selectedOutputId}'.");
            }
            catch (Exception ex)
            {
                StartupDiagnostics.LogException("Failed to export simulation results.", ex);
                _viewModel?.OnSimulationError(ex);
            }
        });
    }

    private static bool ConfirmRawDataExport(MonteCarlo.Engine.Simulation.SimulationResult result)
    {
        var rowCount = (long)result.IterationCount + 1;
        var columnCount = result.Config.Inputs.Count + 2; // iteration + inputs + selected output

        if (rowCount > ExcelMaxRows)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Raw data export needs {rowCount:N0} rows, but Excel worksheets can hold only {ExcelMaxRows:N0} rows.\r\n\r\n" +
                "Use fewer iterations or export a summary instead.",
                "Raw Data Export Too Large",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
            return false;
        }

        var exportedCells = rowCount * columnCount;
        if (exportedCells < RawDataExportWarningCellThreshold)
            return true;

        var confirm = System.Windows.Forms.MessageBox.Show(
            $"Raw data export will write about {exportedCells:N0} cells " +
            $"({result.IterationCount:N0} iterations by {columnCount:N0} columns).\r\n\r\n" +
            "This can take a while and make the workbook much larger. Continue?",
            "Large Raw Data Export",
            System.Windows.Forms.MessageBoxButtons.YesNo,
            System.Windows.Forms.MessageBoxIcon.Warning);

        return confirm == System.Windows.Forms.DialogResult.Yes;
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
            _viewModel.CorrelationImportRequested -= OnCorrelationImportRequested;
            _viewModel.CorrelationExportRequested -= OnCorrelationExportRequested;
            _viewModel.CorrelationMatrixChanged -= OnCorrelationMatrixChanged;
            _viewModel.SetupViewModel.CellSelectionRequested -= OnCellSelectionRequested;
            _viewModel.SetupViewModel.InputAdded -= OnInputAdded;
            _viewModel.SetupViewModel.OutputAdded -= OnOutputAdded;
            _viewModel.SetupViewModel.InputRemoved -= OnInputRemoved;
            _viewModel.SetupViewModel.OutputRemoved -= OnOutputRemoved;
            _viewModel.SetupViewModel.CellActionRequested -= OnCellActionRequested;
            _viewModel.SetupViewModel.RefreshHighlightsRequested -= OnRefreshHighlightsRequested;
            _viewModel.SetupViewModel.DistributionFitRequested -= OnDistributionFitRequested;
        }

        _viewModel = null;
    }
}
