using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
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

    /// <summary>The persistent SetupView instance.</summary>
    public SetupView SetupView => _setupView ??= new SetupView();

    /// <summary>The persistent RunView instance.</summary>
    public RunView RunView => _runView ??= new RunView();

    /// <summary>The persistent ResultsView instance.</summary>
    public ResultsView ResultsView => _resultsView ??= new ResultsView();

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

    public MainViewModel()
    {
        // Wire SetupViewModel's run event to navigate to RunView
        SetupViewModel.RunSimulationRequested += (_, _) =>
            RequestSimulationRun();
        SetupViewModel.CorrelationEditorRequested += () =>
            NavigateToCorrelation();

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
        NavigateToRun();
        RunViewModel.Reset(SetupViewModel.IterationCount);
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
        CurrentView = new SettingsView();
        CurrentViewName = "Settings";
    }

    [RelayCommand]
    private void NavigateToRun()
    {
        CurrentView = RunView;
        CurrentViewName = "Run";
    }

    private void NavigateToCorrelation()
    {
        var viewModel = new CorrelationViewModel();
        viewModel.Initialize(
            SetupViewModel.Inputs.Select(input => input.Label).ToList(),
            SetupViewModel.CorrelationMatrixValues);
        viewModel.Applied += matrix => SetupViewModel.CorrelationMatrixValues = matrix;
        viewModel.CloseRequested += NavigateToSetup;

        CurrentView = new CorrelationView { DataContext = viewModel };
        CurrentViewName = "Correlations";
    }
}
