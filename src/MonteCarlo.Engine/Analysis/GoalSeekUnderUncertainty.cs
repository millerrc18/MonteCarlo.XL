namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Finds the decision value needed to reach a target simulation statistic.
/// </summary>
public static class GoalSeekUnderUncertainty
{
    /// <summary>
    /// Solve a monotonic goal-seek problem by repeatedly simulating midpoint decision values.
    /// </summary>
    public static async Task<GoalSeekResult> SolveAsync(
        GoalSeekOptions options,
        Func<double, CancellationToken, Task<IReadOnlyList<double>>> simulateOutputValues,
        CancellationToken cancellationToken = default)
    {
        ValidateOptions(options);
        ArgumentNullException.ThrowIfNull(simulateOutputValues);

        var history = new List<GoalSeekIteration>();
        var lower = options.LowerBound;
        var upper = options.UpperBound;

        var lowerMetric = await EvaluateDecisionAsync(options, simulateOutputValues, lower, history, 0, cancellationToken)
            .ConfigureAwait(false);
        var upperMetric = await EvaluateDecisionAsync(options, simulateOutputValues, upper, history, 0, cancellationToken)
            .ConfigureAwait(false);

        var bestDecision = lower;
        var bestMetric = lowerMetric;
        if (Math.Abs(upperMetric - options.DesiredMetricValue) < Math.Abs(bestMetric - options.DesiredMetricValue))
        {
            bestDecision = upper;
            bestMetric = upperMetric;
        }

        var bracketLow = Math.Min(lowerMetric, upperMetric);
        var bracketHigh = Math.Max(lowerMetric, upperMetric);
        if (options.DesiredMetricValue < bracketLow || options.DesiredMetricValue > bracketHigh)
        {
            return new GoalSeekResult(
                GoalSeekStatus.TargetNotBracketed,
                bestDecision,
                bestMetric,
                options.DesiredMetricValue,
                bestMetric - options.DesiredMetricValue,
                0,
                history);
        }

        for (var iteration = 1; iteration <= options.MaxIterations; iteration++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var midpoint = (lower + upper) / 2.0;
            var metric = await EvaluateDecisionAsync(options, simulateOutputValues, midpoint, history, iteration, cancellationToken)
                .ConfigureAwait(false);
            var error = metric - options.DesiredMetricValue;

            if (Math.Abs(error) < Math.Abs(bestMetric - options.DesiredMetricValue))
            {
                bestDecision = midpoint;
                bestMetric = metric;
            }

            if (Math.Abs(error) <= options.MetricTolerance)
            {
                return new GoalSeekResult(
                    GoalSeekStatus.Converged,
                    midpoint,
                    metric,
                    options.DesiredMetricValue,
                    error,
                    iteration,
                    history);
            }

            var metricIsBelowTarget = metric < options.DesiredMetricValue;
            if (options.HigherDecisionIncreasesMetric)
            {
                if (metricIsBelowTarget)
                    lower = midpoint;
                else
                    upper = midpoint;
            }
            else
            {
                if (metricIsBelowTarget)
                    upper = midpoint;
                else
                    lower = midpoint;
            }
        }

        return new GoalSeekResult(
            GoalSeekStatus.MaxIterations,
            bestDecision,
            bestMetric,
            options.DesiredMetricValue,
            bestMetric - options.DesiredMetricValue,
            options.MaxIterations,
            history);
    }

    /// <summary>
    /// Synchronous convenience wrapper for in-memory tests and non-UI callers.
    /// </summary>
    public static GoalSeekResult Solve(
        GoalSeekOptions options,
        Func<double, IReadOnlyList<double>> simulateOutputValues)
    {
        ArgumentNullException.ThrowIfNull(simulateOutputValues);

        return SolveAsync(
            options,
            (decisionValue, _) => Task.FromResult(simulateOutputValues(decisionValue)),
            CancellationToken.None).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Compute the configured metric from one simulation's output samples.
    /// </summary>
    public static double EvaluateMetric(IReadOnlyList<double> outputValues, GoalSeekOptions options)
    {
        ValidateOptions(options);
        if (outputValues == null || outputValues.Count == 0)
            throw new ArgumentException("At least one output value is required.", nameof(outputValues));

        var stats = new SummaryStatistics(outputValues.ToArray());
        return options.Metric switch
        {
            GoalSeekMetric.Mean => stats.Mean,
            GoalSeekMetric.ProbabilityAboveTarget => stats.ProbabilityAbove(options.OutputTarget),
            GoalSeekMetric.ProbabilityAtOrBelowTarget => stats.ProbabilityBelow(options.OutputTarget),
            GoalSeekMetric.Percentile => stats.Percentile(options.Percentile),
            _ => throw new ArgumentOutOfRangeException(nameof(options), options.Metric, "Unknown goal seek metric.")
        };
    }

    private static async Task<double> EvaluateDecisionAsync(
        GoalSeekOptions options,
        Func<double, CancellationToken, Task<IReadOnlyList<double>>> simulateOutputValues,
        double decisionValue,
        ICollection<GoalSeekIteration> history,
        int iteration,
        CancellationToken cancellationToken)
    {
        var outputValues = await simulateOutputValues(decisionValue, cancellationToken).ConfigureAwait(false);
        var metric = EvaluateMetric(outputValues, options);
        history.Add(new GoalSeekIteration(
            iteration,
            decisionValue,
            metric,
            metric - options.DesiredMetricValue));
        return metric;
    }

    private static void ValidateOptions(GoalSeekOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        if (!double.IsFinite(options.LowerBound) || !double.IsFinite(options.UpperBound))
            throw new ArgumentException("Bounds must be finite numbers.", nameof(options));
        if (options.LowerBound >= options.UpperBound)
            throw new ArgumentException("LowerBound must be less than UpperBound.", nameof(options));
        if (!double.IsFinite(options.DesiredMetricValue))
            throw new ArgumentException("DesiredMetricValue must be finite.", nameof(options));
        if (options.MaxIterations <= 0)
            throw new ArgumentException("MaxIterations must be positive.", nameof(options));
        if (options.MetricTolerance <= 0 || !double.IsFinite(options.MetricTolerance))
            throw new ArgumentException("MetricTolerance must be a positive finite number.", nameof(options));

        if (options.Metric is GoalSeekMetric.ProbabilityAboveTarget or GoalSeekMetric.ProbabilityAtOrBelowTarget)
        {
            if (options.DesiredMetricValue < 0 || options.DesiredMetricValue > 1)
                throw new ArgumentException("Desired probability must be between 0 and 1.", nameof(options));
            if (!double.IsFinite(options.OutputTarget))
                throw new ArgumentException("OutputTarget must be finite for probability metrics.", nameof(options));
        }

        if (options.Metric == GoalSeekMetric.Percentile
            && (options.Percentile < 0 || options.Percentile > 1 || !double.IsFinite(options.Percentile)))
        {
            throw new ArgumentException("Percentile must be between 0 and 1.", nameof(options));
        }
    }
}

/// <summary>
/// Simulation statistic to match during uncertainty goal seek.
/// </summary>
public enum GoalSeekMetric
{
    /// <summary>Match the mean output value.</summary>
    Mean,

    /// <summary>Match P(output &gt; OutputTarget).</summary>
    ProbabilityAboveTarget,

    /// <summary>Match P(output &lt;= OutputTarget).</summary>
    ProbabilityAtOrBelowTarget,

    /// <summary>Match a percentile output value.</summary>
    Percentile
}

/// <summary>
/// Final status of a goal-seek solve.
/// </summary>
public enum GoalSeekStatus
{
    /// <summary>The requested metric tolerance was reached.</summary>
    Converged,

    /// <summary>The desired metric was outside the lower/upper bound metrics.</summary>
    TargetNotBracketed,

    /// <summary>The solver used all allowed iterations without reaching tolerance.</summary>
    MaxIterations
}

/// <summary>
/// Options for a monotonic uncertainty goal-seek solve.
/// </summary>
public sealed class GoalSeekOptions
{
    public double LowerBound { get; init; }
    public double UpperBound { get; init; }
    public GoalSeekMetric Metric { get; init; } = GoalSeekMetric.Mean;
    public double OutputTarget { get; init; }
    public double DesiredMetricValue { get; init; }
    public double Percentile { get; init; } = 0.5;
    public int MaxIterations { get; init; } = 25;
    public double MetricTolerance { get; init; } = 0.005;
    public bool HigherDecisionIncreasesMetric { get; init; } = true;
}

/// <summary>
/// Result of a goal-seek solve.
/// </summary>
public sealed record GoalSeekResult(
    GoalSeekStatus Status,
    double BestDecisionValue,
    double BestMetricValue,
    double DesiredMetricValue,
    double Error,
    int Iterations,
    IReadOnlyList<GoalSeekIteration> History);

/// <summary>
/// One simulated decision point evaluated during goal seek.
/// </summary>
public sealed record GoalSeekIteration(
    int Iteration,
    double DecisionValue,
    double MetricValue,
    double Error);
