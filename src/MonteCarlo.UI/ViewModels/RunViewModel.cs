using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.UI.Converters;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// View model for the Run view — displays progress, live stats, and convergence during simulation.
/// </summary>
public partial class RunViewModel : ObservableObject
{
    [ObservableProperty] private double _progressPercent;
    [ObservableProperty] private int _completedIterations;
    [ObservableProperty] private int _totalIterations;
    [ObservableProperty] private string _elapsedTime = "0.0s";
    [ObservableProperty] private string _estimatedRemaining = "—";
    [ObservableProperty] private string _iterationRate = "—";
    [ObservableProperty] private string _runModeSummary = "—";
    [ObservableProperty] private string _settingsSourceSummary = "—";
    [ObservableProperty] private string _samplingSummary = "—";
    [ObservableProperty] private string _seedSummary = "—";
    [ObservableProperty] private string _excelBehaviorSummary = "—";
    [ObservableProperty] private string _convergenceSummary = "—";
    [ObservableProperty] private string _warningPauseSummary = "—";

    // Live stats (updated periodically)
    [ObservableProperty] private string _liveMean = "—";
    [ObservableProperty] private string _liveMedian = "—";
    [ObservableProperty] private string _liveStdDev = "—";
    [ObservableProperty] private string _liveP5 = "—";
    [ObservableProperty] private string _liveP95 = "—";

    // Live histogram
    [ObservableProperty] private HistogramData? _liveHistogramData;

    // Convergence
    [ObservableProperty]
    private ObservableCollection<ConvergenceIndicatorViewModel> _convergenceIndicators = new();

    /// <summary>Whether the simulation is currently running.</summary>
    [ObservableProperty] private bool _isRunning;

    /// <summary>Whether a simulation error occurred.</summary>
    [ObservableProperty] private bool _hasError;

    /// <summary>Error type for display (e.g., "Cell Reference Error").</summary>
    [ObservableProperty] private string _errorType = string.Empty;

    /// <summary>Human-readable error message.</summary>
    [ObservableProperty] private string _errorMessage = string.Empty;

    /// <summary>Event raised when the user clicks the Stop button.</summary>
    public event EventHandler? StopRequested;

    /// <summary>Event raised when the user clicks Retry after an error.</summary>
    public event EventHandler? RetryRequested;

    /// <summary>
    /// Update progress from a SimulationProgressEventArgs.
    /// </summary>
    public void UpdateProgress(int completed, int total, TimeSpan elapsed, TimeSpan? estimatedRemaining)
    {
        CompletedIterations = completed;
        TotalIterations = total;
        ProgressPercent = total > 0 ? (double)completed / total * 100 : 0;
        ElapsedTime = FormatTimeSpan(elapsed);
        EstimatedRemaining = estimatedRemaining.HasValue ? $"~{FormatTimeSpan(estimatedRemaining.Value)}" : "—";
        IterationRate = FormatIterationRate(completed, elapsed);
    }

    /// <summary>
    /// Update live stats from partial results.
    /// </summary>
    public void UpdateLiveStats(SummaryStatistics stats)
    {
        LiveMean = NumberFormatter.Format(stats.Mean);
        LiveMedian = NumberFormatter.Format(stats.Median);
        LiveStdDev = NumberFormatter.Format(stats.StdDev);
        LiveP5 = NumberFormatter.Format(stats.P5);
        LiveP95 = NumberFormatter.Format(stats.P95);
    }

    /// <summary>
    /// Update convergence indicators.
    /// </summary>
    public void UpdateConvergence(IReadOnlyList<ConvergenceIndicator> indicators)
    {
        ConvergenceIndicators.Clear();
        foreach (var ind in indicators)
        {
            ConvergenceIndicators.Add(new ConvergenceIndicatorViewModel
            {
                StatName = ind.StatName,
                Status = ind.Status,
                StatusText = ind.Status switch
                {
                    ConvergenceStatus.Stable => "stable",
                    ConvergenceStatus.Drifting => "drifting",
                    ConvergenceStatus.Unstable => "unstable",
                    _ => "..."
                },
                StatusColor = ind.Status switch
                {
                    ConvergenceStatus.Stable => "#10B981",    // green
                    ConvergenceStatus.Drifting => "#F59E0B",  // amber
                    ConvergenceStatus.Unstable => "#EF4444",  // red
                    _ => "#94A3B8"                            // gray
                }
            });
        }
    }

    /// <summary>
    /// Update the live histogram chart from interim data.
    /// </summary>
    public void UpdateLiveChart(HistogramData? histogram)
    {
        if (histogram != null)
            LiveHistogramData = histogram;
    }

    [RelayCommand]
    private void StopSimulation()
    {
        StopRequested?.Invoke(this, EventArgs.Empty);
    }

    [RelayCommand]
    private void RetrySimulation()
    {
        HasError = false;
        RetryRequested?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Show an error state in the run view.
    /// </summary>
    public void ShowError(Exception ex)
    {
        IsRunning = false;
        HasError = true;
        ErrorType = ClassifyError(ex);
        ErrorMessage = ex.Message;
    }

    private static string ClassifyError(Exception ex)
    {
        return ex switch
        {
            InvalidOperationException => "Configuration Error",
            System.Runtime.InteropServices.COMException => "Excel Error",
            OutOfMemoryException => "Memory Error",
            _ => "Simulation Error"
        };
    }

    /// <summary>
    /// Reset the view model for a new run.
    /// </summary>
    public void Reset(int totalIterations, RunSettingsSummary? settingsSummary = null)
    {
        TotalIterations = totalIterations;
        CompletedIterations = 0;
        ProgressPercent = 0;
        ElapsedTime = "0.0s";
        EstimatedRemaining = "—";
        IterationRate = "—";
        RunModeSummary = GetRunModeSummary(totalIterations);
        ApplySettingsSummary(settingsSummary);
        LiveMean = "—";
        LiveMedian = "—";
        LiveStdDev = "—";
        LiveP5 = "—";
        LiveP95 = "—";
        ConvergenceIndicators.Clear();
        LiveHistogramData = null;
        IsRunning = true;
        HasError = false;
    }

    public void ApplySettingsSummary(RunSettingsSummary? settingsSummary)
    {
        SettingsSourceSummary = settingsSummary?.SettingsSource ?? "—";
        SamplingSummary = settingsSummary?.Sampling ?? "—";
        SeedSummary = settingsSummary?.Seed ?? "—";
        ExcelBehaviorSummary = settingsSummary?.ExcelBehavior ?? "—";
        ConvergenceSummary = settingsSummary?.Convergence ?? "—";
        WarningPauseSummary = settingsSummary?.WarningPause ?? "—";
    }

    private static string FormatTimeSpan(TimeSpan ts)
    {
        if (ts.TotalMinutes >= 1)
            return $"{ts.Minutes}m {ts.Seconds}s";
        return $"{ts.TotalSeconds:F1}s";
    }

    private static string FormatIterationRate(int completed, TimeSpan elapsed)
    {
        if (completed <= 0 || elapsed.TotalSeconds <= 0.05)
            return "—";

        var rate = completed / elapsed.TotalSeconds;
        return rate >= 1000
            ? $"{rate:N0} iter/sec"
            : $"{rate:N1} iter/sec";
    }

    private static string GetRunModeSummary(int totalIterations)
    {
        return totalIterations switch
        {
            <= 1_000 => "Preview",
            <= 5_000 => "Standard",
            <= 25_000 => "Full",
            _ => "Deep"
        };
    }
}

public sealed class RunSettingsSummary
{
    public string SettingsSource { get; init; } = "—";

    public string Sampling { get; init; } = "—";

    public string Seed { get; init; } = "—";

    public string ExcelBehavior { get; init; } = "—";

    public string Convergence { get; init; } = "—";

    public string WarningPause { get; init; } = "—";
}

/// <summary>
/// View model for a single convergence indicator row.
/// </summary>
public partial class ConvergenceIndicatorViewModel : ObservableObject
{
    [ObservableProperty] private string _statName = string.Empty;
    [ObservableProperty] private ConvergenceStatus _status;
    [ObservableProperty] private string _statusText = "...";
    [ObservableProperty] private string _statusColor = "#94A3B8";
}
