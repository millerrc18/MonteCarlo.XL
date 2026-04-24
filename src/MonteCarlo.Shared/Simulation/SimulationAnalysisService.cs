using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.Shared.Interop;

namespace MonteCarlo.Shared.Simulation;

public sealed class SimulationAnalysisService
{
    private readonly SimulationSamplingService _samplingService = new();

    public AnalyzeSimulationResponse Analyze(AnalyzeSimulationRequest request)
    {
        EnsureVersion(request.Version);

        var config = _samplingService.BuildConfig(
            request.Profile,
            request.Settings with { IterationCount = request.InputMatrix.Length });

        var inputMatrix = SimulationSamplingService.ToRectangular(request.InputMatrix);
        var outputMatrix = SimulationSamplingService.ToRectangular(request.OutputMatrix);
        var result = new SimulationResult(
            config,
            inputMatrix,
            outputMatrix,
            TimeSpan.FromMilliseconds(request.ElapsedMilliseconds));

        var outputs = new List<OutputAnalysisDto>(config.Outputs.Count);
        for (var outputIndex = 0; outputIndex < config.Outputs.Count; outputIndex++)
        {
            var output = config.Outputs[outputIndex];
            var values = result.GetOutputValues(output.Id);
            var summary = new SummaryStatistics(values);
            var histogram = summary.ToHistogram(Math.Min(50, Math.Max(10, (int)Math.Sqrt(values.Length))));
            var (kdeX, kdeY) = histogram.ComputeKDE();
            var sensitivity = SensitivityAnalysis.Analyze(result, outputIndex)
                .Select(item => new SensitivityItemDto(
                    item.InputId,
                    item.InputLabel,
                    item.RankCorrelation,
                    item.ContributionToVariance))
                .ToArray();
            var worstCase = ToScenarioSummary(ScenarioAnalysis.Analyze(result, outputIndex, ScenarioFilterMode.WorstPercent, 0.10));
            var bestCase = ToScenarioSummary(ScenarioAnalysis.Analyze(result, outputIndex, ScenarioFilterMode.BestPercent, 0.10));

            TargetAnalysisDto? target = null;
            if (request.Settings.OutputTarget.HasValue)
            {
                target = new TargetAnalysisDto(
                    request.Settings.OutputTarget.Value,
                    summary.ProbabilityAbove(request.Settings.OutputTarget.Value),
                    summary.ProbabilityBelow(request.Settings.OutputTarget.Value));
            }

            outputs.Add(new OutputAnalysisDto(
                output.Id,
                output.Label,
                new SummaryStatisticsDto(
                    summary.Count,
                    summary.Mean,
                    summary.Median,
                    summary.Mode,
                    summary.StdDev,
                    summary.Variance,
                    summary.Min,
                    summary.Max,
                    summary.P1,
                    summary.P5,
                    summary.P10,
                    summary.P25,
                    summary.P50,
                    summary.P75,
                    summary.P90,
                    summary.P95,
                    summary.P99,
                    target?.ProbabilityAbove ?? 0,
                    target?.ProbabilityAtOrBelow ?? 0),
                new HistogramDto(
                    histogram.BinEdges,
                    histogram.BinCenters,
                    histogram.Frequencies,
                    histogram.RelativeFrequencies,
                    kdeX,
                    kdeY),
                new CdfDto(
                    summary.SortedValues,
                    Enumerable.Range(0, summary.Count)
                        .Select(index => (index + 1.0) / summary.Count)
                        .ToArray()),
                sensitivity,
                worstCase,
                bestCase,
                target));
        }

        return new AnalyzeSimulationResponse(
            BridgeProtocol.Version,
            result.IterationCount,
            request.ElapsedMilliseconds,
            outputs);
    }

    public EvaluateGoalSeekMetricResponse EvaluateGoalSeekMetric(EvaluateGoalSeekMetricRequest request)
    {
        EnsureVersion(request.Version);
        var metric = GoalSeekUnderUncertainty.EvaluateMetric(
            request.OutputValues,
            new GoalSeekOptions
            {
                Metric = request.Metric,
                OutputTarget = request.OutputTarget,
                Percentile = request.Percentile
            });

        return new EvaluateGoalSeekMetricResponse(BridgeProtocol.Version, metric);
    }

    private static ScenarioSummaryDto ToScenarioSummary(ScenarioAnalysisResult result) =>
        new(
            result.Description,
            result.Threshold,
            result.MatchedIterations,
            result.MatchedFraction,
            result.InputSummaries.Select(summary => new ScenarioInputSummaryDto(
                summary.InputId,
                summary.InputLabel,
                summary.OverallMean,
                summary.ScenarioMean,
                summary.Delta,
                summary.DeltaPercent,
                summary.OverallP10,
                summary.OverallP50,
                summary.OverallP90,
                summary.ScenarioP10,
                summary.ScenarioP50,
                summary.ScenarioP90)).ToArray());

    private static void EnsureVersion(int version)
    {
        if (version != BridgeProtocol.Version)
            throw new InvalidOperationException($"Unsupported bridge version {version}. Expected {BridgeProtocol.Version}.");
    }
}
