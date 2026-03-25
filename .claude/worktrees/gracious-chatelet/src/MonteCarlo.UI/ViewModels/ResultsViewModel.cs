using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.Converters;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// View model for the Results dashboard showing histogram, CDF, stats, and target analysis.
/// </summary>
public partial class ResultsViewModel : ObservableObject
{
    [ObservableProperty] private SimulationResult? _simulationResult;
    [ObservableProperty] private string? _selectedOutputId;
    [ObservableProperty] private SummaryStatistics? _currentStats;
    [ObservableProperty] private HistogramData? _histogramData;
    [ObservableProperty] private bool _showCDF;
    [ObservableProperty] private string _targetValueText = string.Empty;
    [ObservableProperty] private double? _probabilityAboveTarget;
    [ObservableProperty] private double? _probabilityBelowTarget;

    // Sensitivity / Tornado data
    [ObservableProperty] private IReadOnlyList<SensitivityResult>? _sensitivityResults;
    [ObservableProperty] private double _baseOutputValue;

    /// <summary>Available output IDs for the dropdown.</summary>
    [ObservableProperty] private ObservableCollection<OutputItem> _availableOutputs = new();

    // Headline stat
    [ObservableProperty] private string _headlineValue = "—";
    [ObservableProperty] private string _headlineLabel = string.Empty;
    [ObservableProperty] private string _headlineRange = string.Empty;

    // Stats panel (pre-formatted)
    [ObservableProperty] private string _meanFormatted = "—";
    [ObservableProperty] private string _stdDevFormatted = "—";
    [ObservableProperty] private string _p5Formatted = "—";
    [ObservableProperty] private string _p95Formatted = "—";
    [ObservableProperty] private string _minFormatted = "—";
    [ObservableProperty] private string _maxFormatted = "—";
    [ObservableProperty] private string _skewnessFormatted = "—";
    [ObservableProperty] private string _kurtosisFormatted = "—";

    // Percentile values for chart markers
    [ObservableProperty] private double? _p10Value;
    [ObservableProperty] private double? _p50Value;
    [ObservableProperty] private double? _p90Value;
    [ObservableProperty] private double? _targetValueNumeric;
    [ObservableProperty] private string? _targetAnnotation;

    /// <summary>Number of iterations in the simulation.</summary>
    [ObservableProperty] private string _iterationCountFormatted = "—";

    /// <summary>Whether results are available for display.</summary>
    public bool HasResults => SimulationResult != null;

    /// <summary>
    /// Event raised when the user clicks "Export to Sheet".
    /// The Addin layer subscribes to this to call ResultsExporter.
    /// </summary>
    public event EventHandler? ExportRequested;

    /// <summary>Sensitivity data by output ID, populated on completion.</summary>
    private Dictionary<string, IReadOnlyList<SensitivityResult>> _sensitivityByOutput = new();

    /// <summary>Stats data by output ID, populated on completion.</summary>
    private Dictionary<string, SummaryStatistics> _statsByOutput = new();

    partial void OnSimulationResultChanged(SimulationResult? value)
    {
        OnPropertyChanged(nameof(HasResults));
    }

    /// <summary>
    /// Load results from a completed simulation, including pre-computed stats and sensitivity.
    /// </summary>
    public void LoadResults(SimulationResult result,
        Dictionary<string, SummaryStatistics>? statsByOutput = null,
        Dictionary<string, IReadOnlyList<SensitivityResult>>? sensitivityByOutput = null)
    {
        _statsByOutput = statsByOutput ?? new();
        _sensitivityByOutput = sensitivityByOutput ?? new();

        SimulationResult = result;
        AvailableOutputs.Clear();

        foreach (var output in result.Config.Outputs)
            AvailableOutputs.Add(new OutputItem(output.Id, output.Label));

        IterationCountFormatted = result.IterationCount.ToString("N0");

        if (AvailableOutputs.Count > 0)
            SelectedOutputId = AvailableOutputs[0].Id;
    }

    partial void OnSelectedOutputIdChanged(string? value)
    {
        if (value == null || SimulationResult == null) return;
        RecomputeStats();
    }

    partial void OnTargetValueTextChanged(string value)
    {
        if (double.TryParse(value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double target) && CurrentStats != null)
        {
            TargetValueNumeric = target;
            ProbabilityAboveTarget = CurrentStats.ProbabilityAbove(target);
            ProbabilityBelowTarget = CurrentStats.ProbabilityBelow(target);
            TargetAnnotation = $"P(X > {NumberFormatter.Format(target)}) = {ProbabilityAboveTarget:P1}";
        }
        else
        {
            TargetValueNumeric = null;
            ProbabilityAboveTarget = null;
            ProbabilityBelowTarget = null;
            TargetAnnotation = null;
        }
    }

    [RelayCommand]
    private void ToggleChartMode()
    {
        ShowCDF = !ShowCDF;
    }

    [RelayCommand]
    private void ExportToSheet()
    {
        ExportRequested?.Invoke(this, EventArgs.Empty);
    }

    [RelayCommand]
    private void CopyStatsToClipboard()
    {
        if (CurrentStats == null || SimulationResult == null) return;

        var outputLabel = AvailableOutputs.FirstOrDefault(o => o.Id == SelectedOutputId)?.Label ?? SelectedOutputId;
        var stats = CurrentStats;

        var text = $"""
            MonteCarlo.XL Simulation Results — {outputLabel}
            Iterations: {SimulationResult.IterationCount:N0}
            ──────────────────────────
            Mean:        {NumberFormatter.Format(stats.Mean)}
            Median:      {NumberFormatter.Format(stats.Median)}
            Std Dev:     {NumberFormatter.Format(stats.StdDev)}
            P5:          {NumberFormatter.Format(stats.P5)}
            P10:         {NumberFormatter.Format(stats.P10)}
            P90:         {NumberFormatter.Format(stats.P90)}
            P95:         {NumberFormatter.Format(stats.P95)}
            Min:         {NumberFormatter.Format(stats.Min)}
            Max:         {NumberFormatter.Format(stats.Max)}
            Skewness:    {stats.Skewness:F3}
            Kurtosis:    {stats.Kurtosis:F3}
            """;

        System.Windows.Clipboard.SetText(text);
    }

    private void RecomputeStats()
    {
        if (SimulationResult == null || SelectedOutputId == null) return;

        int outputIndex = -1;
        for (int i = 0; i < SimulationResult.Config.Outputs.Count; i++)
        {
            if (SimulationResult.Config.Outputs[i].Id == SelectedOutputId)
            {
                outputIndex = i;
                break;
            }
        }

        if (outputIndex < 0) return;

        // Use pre-computed stats if available, otherwise compute fresh
        SummaryStatistics stats;
        if (_statsByOutput.TryGetValue(SelectedOutputId, out var cached))
        {
            stats = cached;
        }
        else
        {
            var outputValues = SimulationResult.GetOutputValues(SelectedOutputId);
            stats = new SummaryStatistics(outputValues);
        }

        CurrentStats = stats;
        HistogramData = stats.ToHistogram();

        var output = SimulationResult.Config.Outputs[outputIndex];

        // Headline stat
        HeadlineValue = NumberFormatter.Format(stats.Median);
        HeadlineLabel = $"Median {output.Label}";
        HeadlineRange = $"95% CI: {NumberFormatter.FormatRange(stats.P5, stats.P95)}";

        // Stats panel
        MeanFormatted = NumberFormatter.Format(stats.Mean);
        StdDevFormatted = NumberFormatter.Format(stats.StdDev);
        P5Formatted = NumberFormatter.Format(stats.P5);
        P95Formatted = NumberFormatter.Format(stats.P95);
        MinFormatted = NumberFormatter.Format(stats.Min);
        MaxFormatted = NumberFormatter.Format(stats.Max);
        SkewnessFormatted = stats.Skewness.ToString("F3");
        KurtosisFormatted = stats.Kurtosis.ToString("F3");

        // Chart percentile markers
        P10Value = stats.P10;
        P50Value = stats.Median;
        P90Value = stats.P90;

        // Sensitivity / Tornado data
        if (_sensitivityByOutput.TryGetValue(SelectedOutputId, out var sensitivity))
        {
            SensitivityResults = sensitivity;
            BaseOutputValue = stats.Mean;
        }
        else
        {
            // Compute sensitivity on demand if not pre-computed
            try
            {
                var computed = SensitivityAnalysis.Analyze(SimulationResult, outputIndex);
                SensitivityResults = computed;
                BaseOutputValue = stats.Mean;
            }
            catch
            {
                SensitivityResults = null;
                BaseOutputValue = 0;
            }
        }

        // Recompute target if set
        OnTargetValueTextChanged(TargetValueText);
    }
}

/// <summary>
/// Simple data class for the output dropdown.
/// </summary>
public record OutputItem(string Id, string Label)
{
    public override string ToString() => Label;
}
