using System.Windows;
using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.Export;
using MonteCarlo.Addin.TaskPane;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.ViewModels;
using MonteCarlo.UI.Views;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Services;

/// <summary>
/// Wires all events between the UI layer (ViewModels) and the Addin layer
/// (Orchestrator, Excel COM, TagManagers, Highlighter, Config).
/// This class lives in the Addin project because it's the only layer that
/// can reference both UI and Excel interop. All cross-layer communication
/// from UI → Addin uses events (UI fires events, Addin subscribes).
/// </summary>
internal class SimulationPipelineWiring : IDisposable
{
    private readonly MainViewModel _mainViewModel;
    private readonly SimulationOrchestrator _orchestrator;
    private readonly InputTagManager _inputManager;
    private readonly OutputTagManager _outputManager;
    private readonly CellHighlighter _highlighter;
    private readonly ConfigPersistence _configPersistence;
    private readonly WorkbookManager _workbook;

    private bool _disposed;
    private Worksheet? _hookedSheet;
    private System.Timers.Timer? _autoSaveDebounce;

    public SimulationPipelineWiring(
        MainViewModel mainViewModel,
        SimulationOrchestrator orchestrator,
        InputTagManager inputManager,
        OutputTagManager outputManager,
        CellHighlighter highlighter,
        ConfigPersistence configPersistence,
        WorkbookManager workbook)
    {
        _mainViewModel = mainViewModel;
        _orchestrator = orchestrator;
        _inputManager = inputManager;
        _outputManager = outputManager;
        _highlighter = highlighter;
        _configPersistence = configPersistence;
        _workbook = workbook;

        WireSimulationPipeline();
        WireCellSelection();
        WireInputOutputTagging();
        WireCorrelationEditor();
        WireExport();
    }

    #region Simulation Pipeline

    private void WireSimulationPipeline()
    {
        // UI → Addin: Run simulation
        _mainViewModel.RunSimulationRequested += OnRunSimulationRequested;

        // UI → Addin: Cancel simulation
        _mainViewModel.CancelSimulationRequested += OnCancelSimulationRequested;

        // Addin → UI: Progress updates (background thread → UI thread)
        _orchestrator.ProgressChanged += OnOrchestratorProgressChanged;

        // Addin → UI: Simulation complete (background thread → UI thread)
        _orchestrator.SimulationComplete += OnOrchestratorSimulationComplete;

        // Addin → UI: Simulation error (background thread → UI thread)
        _orchestrator.SimulationError += OnOrchestratorSimulationError;
    }

    private async void OnRunSimulationRequested(object? sender, EventArgs e)
    {
        var setup = _mainViewModel.SetupViewModel;

        // Sync SetupViewModel inputs/outputs → TagManagers before running
        SyncTagManagersFromSetup(setup);

        // Save config before running
        try
        {
            _orchestrator.SaveConfig(setup.IterationCount, setup.RandomSeed);
        }
        catch { /* Non-fatal */ }

        try
        {
            await _orchestrator.RunSimulationAsync(setup.IterationCount, setup.RandomSeed);
        }
        catch (OperationCanceledException)
        {
            DispatchToUI(() => _mainViewModel.OnSimulationCancelled());
        }
        catch (Exception ex)
        {
            DispatchToUI(() => _mainViewModel.OnSimulationError(ex));
        }
    }

    private void OnCancelSimulationRequested(object? sender, EventArgs e)
    {
        _orchestrator.CancelSimulation();
    }

    private void OnOrchestratorProgressChanged(object? sender, SimulationProgressEventArgs e)
    {
        DispatchToUI(() => _mainViewModel.OnSimulationProgress(e));
    }

    private void OnOrchestratorSimulationComplete(object? sender, SimulationCompleteEventArgs e)
    {
        DispatchToUI(() =>
        {
            _mainViewModel.OnSimulationComplete(e.Result, e.StatsByOutput, e.SensitivityByOutput);
        });
    }

    private void OnOrchestratorSimulationError(object? sender, SimulationErrorEventArgs e)
    {
        DispatchToUI(() => _mainViewModel.OnSimulationError(e.Error));
    }

    #endregion

    #region Cell Selection

    private void WireCellSelection()
    {
        _mainViewModel.SetupViewModel.CellSelectionRequested += OnCellSelectionRequested;
    }

    private void OnCellSelectionRequested(object? sender, CellSelectionRequestedEventArgs e)
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;

            // Unhook previous sheet if any
            UnhookSheetSelection();

            // Hook the active sheet's SelectionChange event
            var activeSheet = app.ActiveSheet as Worksheet;
            if (activeSheet == null) return;

            _hookedSheet = activeSheet;
            activeSheet.SelectionChange += OnExcelSelectionChange;
        }
        catch { /* Non-fatal */ }
    }

    private void OnExcelSelectionChange(Range target)
    {
        try
        {
            var setup = _mainViewModel.SetupViewModel;

            string sheetName = target.Worksheet.Name;
            string cellAddress = target.Address.Replace("$", "");

            // Try to get a suggested label from the cell to the left or above
            string? suggestedLabel = TryGetCellLabel(target);

            DispatchToUI(() => setup.OnCellSelected(sheetName, cellAddress, suggestedLabel));

            // Unhook after selection (one-shot)
            UnhookSheetSelection();
        }
        catch { /* Non-fatal */ }
    }

    private void UnhookSheetSelection()
    {
        if (_hookedSheet != null)
        {
            try { _hookedSheet.SelectionChange -= OnExcelSelectionChange; }
            catch { }
            _hookedSheet = null;
        }
    }

    private static string? TryGetCellLabel(Range cell)
    {
        try
        {
            // Check the cell to the left
            var leftCell = cell.Offset[0, -1];
            var leftValue = leftCell?.Value2;
            if (leftValue is string label && !string.IsNullOrWhiteSpace(label))
                return label.Trim();
        }
        catch { }

        try
        {
            // Check the cell above
            var aboveCell = cell.Offset[-1, 0];
            var aboveValue = aboveCell?.Value2;
            if (aboveValue is string label && !string.IsNullOrWhiteSpace(label))
                return label.Trim();
        }
        catch { }

        return null;
    }

    #endregion

    #region Input/Output Tagging & Highlighting

    private void WireInputOutputTagging()
    {
        var setup = _mainViewModel.SetupViewModel;

        setup.InputAdded += OnInputAdded;
        setup.InputRemoved += OnInputRemoved;
        setup.OutputAdded += OnOutputAdded;
        setup.OutputRemoved += OnOutputRemoved;
    }

    private void OnInputAdded(object? sender, InputAddedEventArgs e)
    {
        var input = e.Input;
        var cellRef = new CellReference { SheetName = input.SheetName, CellAddress = input.CellAddress };

        _inputManager.TagInput(cellRef, input.Label, input.DistributionName, input.Parameters);

        try { _highlighter.HighlightInput(cellRef); }
        catch { }

        ScheduleAutoSave();
    }

    private void OnInputRemoved(object? sender, InputRemovedEventArgs e)
    {
        var input = e.Input;
        var cellRef = new CellReference { SheetName = input.SheetName, CellAddress = input.CellAddress };

        _inputManager.UntagInput(cellRef);

        try { _highlighter.ClearHighlight(cellRef); }
        catch { }

        ScheduleAutoSave();
    }

    private void OnOutputAdded(object? sender, OutputAddedEventArgs e)
    {
        var output = e.Output;
        var cellRef = new CellReference { SheetName = output.SheetName, CellAddress = output.CellAddress };

        _outputManager.TagOutput(cellRef, output.Label);

        try { _highlighter.HighlightOutput(cellRef); }
        catch { }

        ScheduleAutoSave();
    }

    private void OnOutputRemoved(object? sender, OutputRemovedEventArgs e)
    {
        var output = e.Output;
        var cellRef = new CellReference { SheetName = output.SheetName, CellAddress = output.CellAddress };

        _outputManager.UntagOutput(cellRef);

        try { _highlighter.ClearHighlight(cellRef); }
        catch { }

        ScheduleAutoSave();
    }

    /// <summary>
    /// Syncs the SetupViewModel inputs/outputs to the TagManagers
    /// so the orchestrator can read them when running the simulation.
    /// </summary>
    private void SyncTagManagersFromSetup(SetupViewModel setup)
    {
        // Clear and rebuild from current ViewModel state
        _inputManager.Clear();
        _outputManager.Clear();

        foreach (var input in setup.Inputs)
        {
            var cellRef = new CellReference { SheetName = input.SheetName, CellAddress = input.CellAddress };
            _inputManager.TagInput(cellRef, input.Label, input.DistributionName, input.Parameters);
        }

        foreach (var output in setup.Outputs)
        {
            var cellRef = new CellReference { SheetName = output.SheetName, CellAddress = output.CellAddress };
            _outputManager.TagOutput(cellRef, output.Label);
        }
    }

    #endregion

    #region Correlation Editor

    private void WireCorrelationEditor()
    {
        _mainViewModel.SetupViewModel.CorrelationEditorRequested += OnCorrelationEditorRequested;
    }

    private void OnCorrelationEditorRequested()
    {
        DispatchToUI(() =>
        {
            var setup = _mainViewModel.SetupViewModel;
            var labels = setup.Inputs.Select(i => i.Label).ToList();

            if (labels.Count < 2) return;

            var correlationVM = new CorrelationViewModel();
            correlationVM.Initialize(labels, setup.CorrelationMatrixValues);

            correlationVM.Applied += matrix =>
            {
                setup.CorrelationMatrixValues = matrix;
            };

            correlationVM.CloseRequested += () =>
            {
                _mainViewModel.CurrentView = _mainViewModel.SetupView;
                _mainViewModel.CurrentViewName = "Setup";
            };

            var correlationView = new CorrelationView();
            correlationView.DataContext = correlationVM;

            _mainViewModel.CurrentView = correlationView;
            _mainViewModel.CurrentViewName = "Correlations";
        });
    }

    #endregion

    #region Export

    private void WireExport()
    {
        _mainViewModel.ResultsViewModel.ExportRequested += OnExportRequested;
    }

    private void OnExportRequested(object? sender, EventArgs e)
    {
        var resultsVM = _mainViewModel.ResultsViewModel;
        var result = resultsVM.SimulationResult;
        var stats = resultsVM.CurrentStats;
        var sensitivity = resultsVM.SensitivityResults;

        if (result == null || stats == null) return;

        // Find the output index
        int outputIndex = -1;
        for (int i = 0; i < result.Config.Outputs.Count; i++)
        {
            if (result.Config.Outputs[i].Id == resultsVM.SelectedOutputId)
            {
                outputIndex = i;
                break;
            }
        }
        if (outputIndex < 0) return;

        try
        {
            var exporter = new ResultsExporter();
            var setup = _mainViewModel.SetupViewModel;
            var profile = _orchestrator.BuildProfile(setup.IterationCount, setup.RandomSeed);

            // Render chart images
            byte[]? histogramImage = null;
            byte[]? tornadoImage = null;

            try
            {
                // Render histogram as a WPF control
                var histChart = new MonteCarlo.Charts.Controls.HistogramChart
                {
                    HistogramData = resultsVM.HistogramData,
                    P10 = resultsVM.P10Value,
                    P50 = resultsVM.P50Value,
                    P90 = resultsVM.P90Value,
                    TargetValue = resultsVM.TargetValueNumeric,
                    TargetLabel = resultsVM.TargetAnnotation
                };
                histogramImage = ChartImageRenderer.RenderWpfControl(histChart, 500, 280);
            }
            catch { }

            try
            {
                // Render tornado chart via SkiaSharp
                if (sensitivity != null && sensitivity.Count > 0)
                {
                    var tornadoControl = new MonteCarlo.Charts.Controls.TornadoChart
                    {
                        Results = sensitivity,
                        BaseOutputValue = resultsVM.BaseOutputValue,
                        MaxInputsToShow = 10
                    };
                    tornadoImage = ChartImageRenderer.RenderWpfControl(tornadoControl, 500, 300);
                }
            }
            catch { }

            exporter.ExportSummary(result, stats, sensitivity, profile, outputIndex,
                histogramImage, tornadoImage);
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"Export failed:\n{ex.Message}",
                "MonteCarlo.XL Export Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
    }

    #endregion

    #region Config Auto-Load

    /// <summary>
    /// Loads a saved profile into the SetupViewModel and re-tags/highlights cells.
    /// Called from AddIn after the task pane is first created.
    /// </summary>
    public void LoadSavedProfile()
    {
        try
        {
            var profile = _configPersistence.Load();
            if (profile == null) return;

            DispatchToUI(() =>
            {
                _mainViewModel.SetupViewModel.LoadFromProfile(profile);

                // Re-tag and highlight cells in Excel
                foreach (var input in _mainViewModel.SetupViewModel.Inputs)
                {
                    var cellRef = new CellReference { SheetName = input.SheetName, CellAddress = input.CellAddress };
                    _inputManager.TagInput(cellRef, input.Label, input.DistributionName, input.Parameters);
                    try { _highlighter.HighlightInput(cellRef); } catch { }
                }

                foreach (var output in _mainViewModel.SetupViewModel.Outputs)
                {
                    var cellRef = new CellReference { SheetName = output.SheetName, CellAddress = output.CellAddress };
                    _outputManager.TagOutput(cellRef, output.Label);
                    try { _highlighter.HighlightOutput(cellRef); } catch { }
                }
            });
        }
        catch { /* Non-fatal */ }
    }

    #endregion

    #region Auto-Save (Debounced)

    private void ScheduleAutoSave()
    {
        _autoSaveDebounce?.Stop();
        _autoSaveDebounce?.Dispose();

        _autoSaveDebounce = new System.Timers.Timer(2000); // 2 second debounce
        _autoSaveDebounce.AutoReset = false;
        _autoSaveDebounce.Elapsed += (_, _) =>
        {
            try
            {
                var setup = _mainViewModel.SetupViewModel;
                _orchestrator.SaveConfig(setup.IterationCount, setup.RandomSeed);
            }
            catch { }
        };
        _autoSaveDebounce.Start();
    }

    #endregion

    #region Thread Marshaling

    private static void DispatchToUI(System.Action action)
    {
        var app = System.Windows.Application.Current;
        if (app != null)
        {
            app.Dispatcher.BeginInvoke(action);
        }
        else
        {
            // Fallback: run directly (may fail if on wrong thread)
            action();
        }
    }

    #endregion

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        UnhookSheetSelection();
        _autoSaveDebounce?.Stop();
        _autoSaveDebounce?.Dispose();

        // Unsubscribe events
        _mainViewModel.RunSimulationRequested -= OnRunSimulationRequested;
        _mainViewModel.CancelSimulationRequested -= OnCancelSimulationRequested;
        _orchestrator.ProgressChanged -= OnOrchestratorProgressChanged;
        _orchestrator.SimulationComplete -= OnOrchestratorSimulationComplete;
        _orchestrator.SimulationError -= OnOrchestratorSimulationError;

        _mainViewModel.SetupViewModel.CellSelectionRequested -= OnCellSelectionRequested;
        _mainViewModel.SetupViewModel.InputAdded -= OnInputAdded;
        _mainViewModel.SetupViewModel.InputRemoved -= OnInputRemoved;
        _mainViewModel.SetupViewModel.OutputAdded -= OnOutputAdded;
        _mainViewModel.SetupViewModel.OutputRemoved -= OnOutputRemoved;
        _mainViewModel.SetupViewModel.CorrelationEditorRequested -= OnCorrelationEditorRequested;
        _mainViewModel.ResultsViewModel.ExportRequested -= OnExportRequested;
    }
}
