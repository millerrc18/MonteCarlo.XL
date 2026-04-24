using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.Engine.Validation;
using MonteCarlo.UI.Views;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// Main view model for the task pane. Manages navigation between views
/// and coordinates the simulation lifecycle with the Addin layer.
/// </summary>
public partial class MainViewModel : ObservableObject
{
    [ObservableProperty]
    private string _title = "MonteCarlo.XL";

    [ObservableProperty]
    private object? _currentView;

    [ObservableProperty]
    private string _currentViewName = "Setup";

    private SetupView? _setupView;
    private RunView? _runView;
    private ResultsView? _resultsView;
    private PreflightView? _preflightView;
    private SettingsView? _settingsView;

    /// <summary>The persistent SetupView instance.</summary>
    public SetupView SetupView => _setupView ??= new SetupView();

    /// <summary>The persistent RunView instance.</summary>
    public RunView RunView => _runView ??= new RunView();

    /// <summary>The persistent ResultsView instance.</summary>
    public ResultsView ResultsView => _resultsView ??= new ResultsView();

    /// <summary>The persistent PreflightView instance.</summary>
    public PreflightView PreflightView => _preflightView ??= CreatePreflightView();

    /// <summary>The persistent SettingsView instance.</summary>
    public SettingsView SettingsView => _settingsView ??= new SettingsView();

    /// <summary>Convenience accessor for the SetupViewModel.</summary>
    public SetupViewModel SetupViewModel => SetupView.ViewModel;

    /// <summary>Convenience accessor for the RunViewModel.</summary>
    public RunViewModel RunViewModel => RunView.ViewModel;

    /// <summary>Convenience accessor for the ResultsViewModel.</summary>
    public ResultsViewModel ResultsViewModel => ResultsView.ViewModel;

    /// <summary>
    /// Event raised when the user initiates a simulation run.
    /// The Addin layer subscribes to this and calls the SimulationOrchestrator.
    /// </summary>
    public event EventHandler? RunSimulationRequested;

    /// <summary>
    /// Event raised when the user cancels a running simulation.
    /// </summary>
    public event EventHandler? CancelSimulationRequested;

    /// <summary>
    /// Optional host-supplied preflight provider. Excel add-in hosts can append workbook/cell checks.
    /// </summary>
    public Func<SimulationProfile?, PreflightReport>? PreflightReportProvider { get; set; }

    /// <summary>
    /// Optional host-supplied setting that controls whether warning-only reports pause a run.
    /// </summary>
    public Func<bool>? PauseOnPreflightWarningsProvider { get; set; }

    /// <summary>
    /// Optional host-supplied summary builder for the run screen.
    /// </summary>
    public Func<RunSettingsSummary>? RunSettingsProvider { get; set; }

    /// <summary>
    /// Event raised when the correlation editor requests an Excel range import.
    /// </summary>
    public event Action<CorrelationViewModel>? CorrelationImportRequested;

    /// <summary>
    /// Event raised when the correlation editor requests an Excel range export.
    /// </summary>
    public event Action<CorrelationViewModel>? CorrelationExportRequested;

    /// <summary>
    /// Event raised when a correlation matrix is applied to the setup.
    /// </summary>
    public event EventHandler? CorrelationMatrixChanged;

    public MainViewModel()
    {
        // Wire SetupViewModel's run event to navigate to RunView
        SetupViewModel.RunSimulationRequested += (_, _) =>
            RequestSimulationRun();
        SetupViewModel.CorrelationEditorRequested += () =>
            NavigateToCorrelation();
        SetupViewModel.PreflightRequested += () =>
            ShowPreflightView();

        // Wire RunViewModel's stop event
        RunViewModel.StopRequested += (_, _) =>
        {
            CancelSimulationRequested?.Invoke(this, EventArgs.Empty);
        };

        // Wire RunViewModel's retry event
        RunViewModel.RetryRequested += (_, _) =>
            RequestSimulationRun();

        NavigateToSetup();
    }

    /// <summary>
    /// Starts the UI side of a simulation run and notifies the host add-in.
    /// </summary>
    public void RequestSimulationRun()
    {
        var report = BuildPreflightReport();
        if (report.HasErrors || (report.WarningCount > 0 && ShouldPauseOnPreflightWarnings()))
        {
            NavigateToPreflight(report);
            return;
        }

        StartSimulationRun();
    }

    private void StartSimulationRun()
    {
        NavigateToRun();
        RunViewModel.Reset(SetupViewModel.IterationCount, RunSettingsProvider?.Invoke());
        RunSimulationRequested?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Public navigation entry point for host integrations such as the Excel ribbon.
    /// </summary>
    public void ShowSetupView() => NavigateToSetup();

    /// <summary>
    /// Public navigation entry point for host integrations such as the Excel ribbon.
    /// </summary>
    public void ShowResultsView() => NavigateToResults();

    /// <summary>
    /// Public navigation entry point for host integrations such as the Excel ribbon.
    /// </summary>
    public void ShowSettingsView() => NavigateToSettings();

    /// <summary>
    /// Public navigation entry point for host integrations such as the Excel ribbon.
    /// </summary>
    public void ShowCorrelationView() => NavigateToCorrelation();

    /// <summary>
    /// Public navigation entry point for host integrations such as the Excel ribbon.
    /// </summary>
    public void ShowPreflightView() =>
        NavigateToPreflight(BuildPreflightReport());

    /// <summary>
    /// Called by the Addin layer when simulation progress is reported.
    /// </summary>
    public void OnSimulationProgress(SimulationProgressEventArgs e)
    {
        RunViewModel.UpdateProgress(e.CompletedIterations, e.TotalIterations, e.Elapsed, e.EstimatedRemaining);
        RunViewModel.UpdateLiveChart(e.InterimHistogram);
    }

    /// <summary>
    /// Called by the Addin layer with live stats during the simulation.
    /// </summary>
    public void OnLiveStatsUpdate(SummaryStatistics stats)
    {
        RunViewModel.UpdateLiveStats(stats);
    }

    /// <summary>
    /// Called by the Addin layer with convergence data.
    /// </summary>
    public void OnConvergenceUpdate(IReadOnlyList<ConvergenceIndicator> indicators)
    {
        RunViewModel.UpdateConvergence(indicators);
    }

    /// <summary>
    /// Called by the Addin layer when simulation completes successfully.
    /// </summary>
    public void OnSimulationComplete(SimulationResult result, Dictionary<string, IReadOnlyList<SensitivityResult>>? sensitivityResults = null)
    {
        RunViewModel.IsRunning = false;
        ResultsViewModel.LoadResults(result, sensitivityResults);
        NavigateToResults();
    }

    /// <summary>
    /// Called by the Addin layer when simulation is cancelled.
    /// </summary>
    public void OnSimulationCancelled()
    {
        RunViewModel.IsRunning = false;
        NavigateToSetup();
    }

    /// <summary>
    /// Called by the Addin layer when simulation fails.
    /// </summary>
    public void OnSimulationError(Exception ex)
    {
        RunViewModel.ShowError(ex);
    }

    [RelayCommand]
    private void NavigateToSetup()
    {
        CurrentView = SetupView;
        CurrentViewName = "Setup";
    }

    [RelayCommand]
    private void NavigateToResults()
    {
        CurrentView = ResultsView;
        CurrentViewName = "Results";
    }

    [RelayCommand]
    private void NavigateToSettings()
    {
        CurrentView = SettingsView;
        CurrentViewName = "Settings";
    }

    [RelayCommand]
    private void NavigateToRun()
    {
        CurrentView = RunView;
        CurrentViewName = "Run";
    }

    private void NavigateToPreflight(PreflightReport report)
    {
        var viewModel = (PreflightViewModel)PreflightView.DataContext;
        viewModel.Load(report);

        CurrentView = PreflightView;
        CurrentViewName = "Model Check";
    }

    private PreflightReport BuildPreflightReport()
    {
        var profile = SetupViewModel.BuildSimulationProfile();
        return PreflightReportProvider?.Invoke(profile) ?? ModelPreflightValidator.Validate(profile);
    }

    private bool ShouldPauseOnPreflightWarnings() =>
        PauseOnPreflightWarningsProvider?.Invoke() ?? false;

    private void NavigateToCorrelation()
    {
        var viewModel = new CorrelationViewModel();
        viewModel.Initialize(
            SetupViewModel.Inputs.Select(input => input.Label).ToList(),
            SetupViewModel.CorrelationMatrixValues);
        viewModel.Applied += matrix =>
        {
            SetupViewModel.CorrelationMatrixValues = matrix;
            CorrelationMatrixChanged?.Invoke(this, EventArgs.Empty);
        };
        viewModel.ImportRequested += () => CorrelationImportRequested?.Invoke(viewModel);
        viewModel.ExportRequested += () => CorrelationExportRequested?.Invoke(viewModel);
        viewModel.CloseRequested += NavigateToSetup;

        CurrentView = new CorrelationView { DataContext = viewModel };
        CurrentViewName = "Correlations";
    }

    private PreflightView CreatePreflightView()
    {
        var viewModel = new PreflightViewModel();
        viewModel.BackToSetupRequested += NavigateToSetup;
        viewModel.RunAnywayRequested += StartSimulationRun;

        return new PreflightView { DataContext = viewModel };
    }
}
