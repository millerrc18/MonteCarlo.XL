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
    private const int MaxScenarioInputs = 6;

    [ObservableProperty] private SimulationResult? _simulationResult;
    [ObservableProperty] private string? _selectedOutputId;
    [ObservableProperty] private SummaryStatistics? _currentStats;
    [ObservableProperty] private HistogramData? _histogramData;
    [ObservableProperty] private bool _showCDF;
    [ObservableProperty] private string _targetValueText = string.Empty;
    [ObservableProperty] private double? _probabilityAboveTarget;
    [ObservableProperty] private double? _probabilityBelowTarget;

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

    // Scatter plot properties
    [ObservableProperty] private string? _selectedScatterInputId;
    [ObservableProperty] private double[]? _scatterInputValues;
    [ObservableProperty] private double[]? _scatterOutputValues;
    [ObservableProperty] private string? _scatterInputLabel;
    [ObservableProperty] private bool _hasScatterData;
    [ObservableProperty] private ObservableCollection<string> _availableInputs = new();

    // Scenario analysis properties
    [ObservableProperty] private ObservableCollection<ScenarioFilterOption> _scenarioOptions = new(new ScenarioFilterOption[]
    {
        new(ScenarioFilterMode.WorstPercent.ToString(), "Worst tail"),
        new(ScenarioFilterMode.BestPercent.ToString(), "Best tail"),
        new(ScenarioFilterMode.AtOrBelowTarget.ToString(), "At or below target"),
        new(ScenarioFilterMode.AboveTarget.ToString(), "Above target")
    });
    [ObservableProperty] private string _selectedScenarioMode = ScenarioFilterMode.WorstPercent.ToString();
    [ObservableProperty] private string _scenarioTailPercentText = "10";
    [ObservableProperty] private ObservableCollection<ScenarioInputSummaryItem> _scenarioInputSummaries = new();
    [ObservableProperty] private string _scenarioStatusText = "Run a simulation to compare tail cases.";
    [ObservableProperty] private string _scenarioThresholdText = string.Empty;
    [ObservableProperty] private bool _hasScenarioSummaries;

    /// <summary>Sensitivity analysis results keyed by output ID.</summary>
    public Dictionary<string, IReadOnlyList<SensitivityResult>>? SensitivityResults { get; private set; }

    /// <summary>Whether results are available for display.</summary>
    public bool HasResults => SimulationResult != null;

    /// <summary>Whether the selected scenario mode uses the target value.</summary>
    public bool ScenarioUsesTarget =>
        SelectedScenarioMode == ScenarioFilterMode.AtOrBelowTarget.ToString()
        || SelectedScenarioMode == ScenarioFilterMode.AboveTarget.ToString();

    partial void OnSimulationResultChanged(SimulationResult? value)
    {
        OnPropertyChanged(nameof(HasResults));
    }

    /// <summary>
    /// Load results from a completed simulation.
    /// </summary>
    public void LoadResults(SimulationResult result, Dictionary<string, IReadOnlyList<SensitivityResult>>? sensitivityResults = null)
    {
        SimulationResult = result;
        SensitivityResults = sensitivityResults;
        AvailableOutputs.Clear();

        foreach (var output in result.Config.Outputs)
            AvailableOutputs.Add(new OutputItem(output.Id, output.Label));

        IterationCountFormatted = result.IterationCount.ToString("N0");

        // Populate available inputs for scatter plot
        AvailableInputs.Clear();
        foreach (var input in result.Config.Inputs)
            AvailableInputs.Add(input.Id);

        if (AvailableOutputs.Count > 0)
            SelectedOutputId = AvailableOutputs[0].Id;

        if (AvailableInputs.Count > 0)
            SelectedScatterInputId = AvailableInputs[0];
    }

    partial void OnSelectedOutputIdChanged(string? value)
    {
        if (value == null || SimulationResult == null) return;
        RecomputeStats();
        UpdateScatterData();
        UpdateScenarioAnalysis();
    }

    partial void OnSelectedScatterInputIdChanged(string? value)
    {
        UpdateScatterData();
    }

    partial void OnSelectedScenarioModeChanged(string value)
    {
        OnPropertyChanged(nameof(ScenarioUsesTarget));
        UpdateScenarioAnalysis();
    }

    partial void OnScenarioTailPercentTextChanged(string value)
    {
        UpdateScenarioAnalysis();
    }

    private void UpdateScatterData()
    {
        if (SimulationResult == null || SelectedScatterInputId == null || SelectedOutputId == null)
        {
            ScatterInputValues = null;
            ScatterOutputValues = null;
            ScatterInputLabel = null;
            HasScatterData = false;
            return;
        }

        try
        {
            ScatterInputValues = SimulationResult.GetInputSamples(SelectedScatterInputId);
            ScatterOutputValues = SimulationResult.GetOutputValues(SelectedOutputId);

            // Find the label for the selected input
            var input = SimulationResult.Config.Inputs.FirstOrDefault(i => i.Id == SelectedScatterInputId);
            ScatterInputLabel = input?.Label ?? SelectedScatterInputId;
            HasScatterData = true;
        }
        catch
        {
            ScatterInputValues = null;
            ScatterOutputValues = null;
            ScatterInputLabel = null;
            HasScatterData = false;
        }
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

        UpdateScenarioAnalysis();
    }

    [RelayCommand]
    private void ToggleChartMode()
    {
        ShowCDF = !ShowCDF;
    }

    [RelayCommand]
    private void CopyStatsToClipboard()
    {
        if (CurrentStats == null) return;

        var s = CurrentStats;
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Output\t{SelectedOutputId}");
        sb.AppendLine($"Mean\t{s.Mean:G6}");
        sb.AppendLine($"Median\t{s.Median:G6}");
        sb.AppendLine($"Std Dev\t{s.StdDev:G6}");
        sb.AppendLine($"Variance\t{s.Variance:G6}");
        sb.AppendLine($"Skewness\t{s.Skewness:G6}");
        sb.AppendLine($"Kurtosis\t{s.Kurtosis:G6}");
        sb.AppendLine($"Min\t{s.Min:G6}");
        sb.AppendLine($"Max\t{s.Max:G6}");
        sb.AppendLine($"P5\t{s.P5:G6}");
        sb.AppendLine($"P10\t{s.P10:G6}");
        sb.AppendLine($"P25\t{s.P25:G6}");
        sb.AppendLine($"P50\t{s.P50:G6}");
        sb.AppendLine($"P75\t{s.P75:G6}");
        sb.AppendLine($"P90\t{s.P90:G6}");
        sb.AppendLine($"P95\t{s.P95:G6}");
        sb.AppendLine($"Iterations\t{s.Count}");

        System.Windows.Clipboard.SetText(sb.ToString());
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

        var outputValues = SimulationResult.GetOutputValues(SelectedOutputId);
        var stats = new SummaryStatistics(outputValues);
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

        // Recompute target if set
        OnTargetValueTextChanged(TargetValueText);
    }

    private void UpdateScenarioAnalysis()
    {
        ScenarioInputSummaries.Clear();
        HasScenarioSummaries = false;
        ScenarioThresholdText = string.Empty;

        if (SimulationResult == null || SelectedOutputId == null)
        {
            ScenarioStatusText = "Run a simulation to compare tail cases.";
            return;
        }

        var outputIndex = FindSelectedOutputIndex();
        if (outputIndex < 0)
        {
            ScenarioStatusText = "Select an output to analyze scenarios.";
            return;
        }

        if (!Enum.TryParse<ScenarioFilterMode>(SelectedScenarioMode, out var mode))
        {
            ScenarioStatusText = "Choose a scenario filter.";
            return;
        }

        double thresholdOrFraction;
        if (mode is ScenarioFilterMode.WorstPercent or ScenarioFilterMode.BestPercent)
        {
            if (!TryParseScenarioTailFraction(ScenarioTailPercentText, out thresholdOrFraction))
            {
                ScenarioStatusText = "Enter a tail size between 0 and 100.";
                return;
            }
        }
        else
        {
            if (TargetValueNumeric is not double target)
            {
                ScenarioStatusText = "Enter a target value above to analyze target-hit cases.";
                return;
            }

            thresholdOrFraction = target;
        }

        try
        {
            var analysis = ScenarioAnalysis.Analyze(SimulationResult, outputIndex, mode, thresholdOrFraction);
            ScenarioStatusText =
                $"{analysis.Description}: {analysis.MatchedIterations:N0} of {analysis.TotalIterations:N0} runs ({analysis.MatchedFraction:P1}).";
            ScenarioThresholdText = $"Cutoff: {NumberFormatter.Format(analysis.Threshold)}";

            foreach (var summary in analysis.InputSummaries.Take(MaxScenarioInputs))
                ScenarioInputSummaries.Add(ScenarioInputSummaryItem.From(summary));

            HasScenarioSummaries = ScenarioInputSummaries.Count > 0;
            if (!HasScenarioSummaries)
                ScenarioStatusText += " No runs matched this filter.";
        }
        catch (Exception ex) when (ex is ArgumentException or ArgumentOutOfRangeException)
        {
            ScenarioStatusText = ex.Message;
        }
    }

    private int FindSelectedOutputIndex()
    {
        if (SimulationResult == null || SelectedOutputId == null)
            return -1;

        for (var i = 0; i < SimulationResult.Config.Outputs.Count; i++)
        {
            if (SimulationResult.Config.Outputs[i].Id == SelectedOutputId)
                return i;
        }

        return -1;
    }

    private static bool TryParseScenarioTailFraction(string text, out double fraction)
    {
        fraction = 0;
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var trimmed = text.Trim();
        if (trimmed.EndsWith("%", StringComparison.Ordinal))
            trimmed = trimmed[..^1];

        if (!double.TryParse(
                trimmed,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var value))
            return false;

        fraction = value > 1 ? value / 100.0 : value;
        return fraction > 0 && fraction < 1;
    }
}

/// <summary>
/// Simple data class for the output dropdown.
/// </summary>
public record OutputItem(string Id, string Label)
{
    public override string ToString() => Label;
}

/// <summary>
/// Scenario filter option for the results view.
/// </summary>
public record ScenarioFilterOption(string Mode, string Label);

/// <summary>
/// Formatted conditional input summary for scenario analysis.
/// </summary>
public record ScenarioInputSummaryItem(
    string InputLabel,
    string ScenarioMean,
    string OverallMean,
    string Delta,
    string DeltaPercent,
    string ScenarioRange,
    string OverallRange)
{
    public static ScenarioInputSummaryItem From(ScenarioInputSummary summary)
    {
        return new ScenarioInputSummaryItem(
            summary.InputLabel,
            NumberFormatter.Format(summary.ScenarioMean),
            NumberFormatter.Format(summary.OverallMean),
            FormatSigned(summary.Delta),
            summary.DeltaPercent is double pct ? pct.ToString("+0.0%;-0.0%;0.0%") : "n/a",
            NumberFormatter.FormatRange(summary.ScenarioP10, summary.ScenarioP90),
            NumberFormatter.FormatRange(summary.OverallP10, summary.OverallP90));
    }

    private static string FormatSigned(double value)
    {
        var formatted = NumberFormatter.Format(value);
        return value > 0 ? $"+{formatted}" : formatted;
    }
}
