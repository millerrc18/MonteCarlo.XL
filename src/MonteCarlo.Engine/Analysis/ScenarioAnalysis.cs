using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Filters simulation iterations into scenarios and summarizes how inputs differ in those cases.
/// </summary>
public static class ScenarioAnalysis
{
    /// <summary>
    /// Analyze the input assumptions behind a selected scenario for one output.
    /// </summary>
    public static ScenarioAnalysisResult Analyze(
        SimulationResult result,
        int outputIndex,
        ScenarioFilterMode mode,
        double thresholdOrFraction)
    {
        ArgumentNullException.ThrowIfNull(result);
        if (outputIndex < 0 || outputIndex >= result.Config.Outputs.Count)
            throw new ArgumentOutOfRangeException(nameof(outputIndex));

        var output = result.Config.Outputs[outputIndex];
        var outputValues = result.GetOutputValues(output.Id);
        var outputStats = new SummaryStatistics(outputValues);
        var (threshold, description, matcher) = BuildMatcher(outputStats, mode, thresholdOrFraction);

        var matched = new bool[result.IterationCount];
        var matchedCount = 0;
        for (var i = 0; i < result.IterationCount; i++)
        {
            matched[i] = matcher(outputValues[i]);
            if (matched[i])
                matchedCount++;
        }

        var summaries = new List<ScenarioInputSummary>();
        if (matchedCount > 0)
        {
            foreach (var input in result.Config.Inputs)
            {
                var values = result.GetInputSamples(input.Id);
                var scenarioValues = new double[matchedCount];
                var scenarioIndex = 0;
                for (var i = 0; i < values.Length; i++)
                {
                    if (matched[i])
                        scenarioValues[scenarioIndex++] = values[i];
                }

                var overallStats = new SummaryStatistics(values);
                var scenarioStats = new SummaryStatistics(scenarioValues);
                var delta = scenarioStats.Mean - overallStats.Mean;
                double? deltaPercent = Math.Abs(overallStats.Mean) > 1e-12
                    ? delta / Math.Abs(overallStats.Mean)
                    : null;

                summaries.Add(new ScenarioInputSummary(
                    input.Id,
                    input.Label,
                    overallStats.Mean,
                    scenarioStats.Mean,
                    delta,
                    deltaPercent,
                    overallStats.P10,
                    overallStats.Median,
                    overallStats.P90,
                    scenarioStats.P10,
                    scenarioStats.Median,
                    scenarioStats.P90));
            }
        }

        return new ScenarioAnalysisResult(
            output.Id,
            output.Label,
            mode,
            description,
            threshold,
            result.IterationCount,
            matchedCount,
            summaries
                .OrderByDescending(s => Math.Abs(s.DeltaPercent ?? s.Delta))
                .ToList());
    }

    private static (double Threshold, string Description, Func<double, bool> Matcher) BuildMatcher(
        SummaryStatistics outputStats,
        ScenarioFilterMode mode,
        double thresholdOrFraction)
    {
        return mode switch
        {
            ScenarioFilterMode.WorstPercent => BuildTailMatcher(
                outputStats,
                thresholdOrFraction,
                lowerTail: true),

            ScenarioFilterMode.BestPercent => BuildTailMatcher(
                outputStats,
                thresholdOrFraction,
                lowerTail: false),

            ScenarioFilterMode.AtOrBelowTarget => (
                thresholdOrFraction,
                $"Runs at or below {thresholdOrFraction:G6}",
                value => value <= thresholdOrFraction),

            ScenarioFilterMode.AboveTarget => (
                thresholdOrFraction,
                $"Runs above {thresholdOrFraction:G6}",
                value => value > thresholdOrFraction),

            _ => throw new ArgumentOutOfRangeException(nameof(mode), mode, "Unknown scenario filter mode.")
        };
    }

    private static (double Threshold, string Description, Func<double, bool> Matcher) BuildTailMatcher(
        SummaryStatistics outputStats,
        double fraction,
        bool lowerTail)
    {
        if (fraction <= 0 || fraction >= 1)
            throw new ArgumentOutOfRangeException(nameof(fraction), fraction, "Tail fraction must be greater than 0 and less than 1.");

        if (lowerTail)
        {
            var threshold = outputStats.Percentile(fraction);
            return (threshold, $"Worst {fraction * 100:G3}% of runs", value => value <= threshold);
        }
        else
        {
            var threshold = outputStats.Percentile(1.0 - fraction);
            return (threshold, $"Best {fraction * 100:G3}% of runs", value => value >= threshold);
        }
    }
}

/// <summary>
/// Supported filters for scenario analysis.
/// </summary>
public enum ScenarioFilterMode
{
    /// <summary>Lowest output values by percentile.</summary>
    WorstPercent,

    /// <summary>Highest output values by percentile.</summary>
    BestPercent,

    /// <summary>Output values at or below a user target.</summary>
    AtOrBelowTarget,

    /// <summary>Output values above a user target.</summary>
    AboveTarget
}

/// <summary>
/// Summary of one scenario filter applied to one output.
/// </summary>
public sealed record ScenarioAnalysisResult(
    string OutputId,
    string OutputLabel,
    ScenarioFilterMode Mode,
    string Description,
    double Threshold,
    int TotalIterations,
    int MatchedIterations,
    IReadOnlyList<ScenarioInputSummary> InputSummaries)
{
    /// <summary>Share of iterations that matched the scenario filter.</summary>
    public double MatchedFraction => TotalIterations == 0 ? 0 : (double)MatchedIterations / TotalIterations;
}

/// <summary>
/// Conditional input summary for iterations that matched a scenario.
/// </summary>
public sealed record ScenarioInputSummary(
    string InputId,
    string InputLabel,
    double OverallMean,
    double ScenarioMean,
    double Delta,
    double? DeltaPercent,
    double OverallP10,
    double OverallP50,
    double OverallP90,
    double ScenarioP10,
    double ScenarioP50,
    double ScenarioP90);
