namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Monitors convergence of simulation statistics during a run.
/// A statistic is "converged" when its value changes by less than a tolerance
/// over a rolling window of checkpoints.
/// </summary>
public class ConvergenceChecker
{
    private readonly int _checkpointInterval;
    private readonly int _windowSize;
    private readonly double _tolerance;
    private readonly List<CheckpointData> _checkpoints = new();

    /// <summary>
    /// Creates a new convergence checker.
    /// </summary>
    /// <param name="checkpointInterval">How often to take a checkpoint (in iterations).</param>
    /// <param name="windowSize">Number of recent checkpoints to compare (default 3).</param>
    /// <param name="tolerance">Maximum relative change rate to consider converged (default 0.5%).</param>
    public ConvergenceChecker(int checkpointInterval = 500, int windowSize = 3, double tolerance = 0.005)
    {
        _checkpointInterval = checkpointInterval;
        _windowSize = windowSize;
        _tolerance = tolerance;
    }

    /// <summary>
    /// Record a checkpoint at the current iteration.
    /// </summary>
    public void RecordCheckpoint(int iteration, double[] outputValues)
    {
        if (outputValues.Length == 0) return;

        var stats = new SummaryStatistics(outputValues);
        _checkpoints.Add(new CheckpointData
        {
            Iteration = iteration,
            Mean = stats.Mean,
            Median = stats.Median,
            P90 = stats.P90,
            StdDev = stats.StdDev
        });
    }

    /// <summary>
    /// Check convergence status for all tracked statistics.
    /// </summary>
    public IReadOnlyList<ConvergenceIndicator> CheckAll()
    {
        var indicators = new List<ConvergenceIndicator>
        {
            CheckStat("Mean", _checkpoints.Select(c => c.Mean).ToArray()),
            CheckStat("P50", _checkpoints.Select(c => c.Median).ToArray()),
            CheckStat("P90", _checkpoints.Select(c => c.P90).ToArray()),
            CheckStat("Std Dev", _checkpoints.Select(c => c.StdDev).ToArray())
        };
        return indicators;
    }

    /// <summary>
    /// Check convergence of a single series of checkpoint values.
    /// </summary>
    public ConvergenceIndicator CheckStat(string name, double[] values)
    {
        if (values.Length < 2)
        {
            return new ConvergenceIndicator
            {
                StatName = name,
                Status = ConvergenceStatus.InsufficientData,
                CurrentValue = values.Length > 0 ? values[^1] : 0,
                ChangeRate = double.NaN
            };
        }

        int windowStart = Math.Max(0, values.Length - _windowSize);
        double startValue = values[windowStart];
        double currentValue = values[^1];

        double changeRate = startValue != 0
            ? Math.Abs(currentValue - startValue) / Math.Abs(startValue)
            : Math.Abs(currentValue - startValue);

        // Determine if the trend is improving (change rate decreasing over time)
        bool improving = false;
        if (values.Length >= 3)
        {
            int prevStart = Math.Max(0, values.Length - _windowSize - 1);
            double prevChangeRate = startValue != 0
                ? Math.Abs(values[windowStart] - values[prevStart]) / Math.Abs(values[prevStart])
                : Math.Abs(values[windowStart] - values[prevStart]);
            improving = changeRate < prevChangeRate;
        }

        ConvergenceStatus status;
        if (changeRate < _tolerance)
            status = ConvergenceStatus.Stable;
        else if (improving)
            status = ConvergenceStatus.Drifting;
        else
            status = ConvergenceStatus.Unstable;

        return new ConvergenceIndicator
        {
            StatName = name,
            Status = status,
            CurrentValue = currentValue,
            ChangeRate = changeRate
        };
    }

    /// <summary>
    /// Check if a series of values has converged.
    /// </summary>
    public bool IsConverged(double[] checkpointValues, double? tolerance = null)
    {
        double tol = tolerance ?? _tolerance;
        if (checkpointValues.Length < _windowSize) return false;

        int windowStart = checkpointValues.Length - _windowSize;
        double startValue = checkpointValues[windowStart];
        double currentValue = checkpointValues[^1];

        double changeRate = startValue != 0
            ? Math.Abs(currentValue - startValue) / Math.Abs(startValue)
            : Math.Abs(currentValue - startValue);

        return changeRate < tol;
    }

    /// <summary>Number of checkpoints recorded.</summary>
    public int CheckpointCount => _checkpoints.Count;

    /// <summary>Reset all checkpoint data.</summary>
    public void Reset() => _checkpoints.Clear();

    private class CheckpointData
    {
        public int Iteration { get; init; }
        public double Mean { get; init; }
        public double Median { get; init; }
        public double P90 { get; init; }
        public double StdDev { get; init; }
    }
}

/// <summary>
/// Convergence status for a single statistic.
/// </summary>
public class ConvergenceIndicator
{
    /// <summary>Name of the statistic (e.g., "Mean", "P50").</summary>
    public string StatName { get; init; } = string.Empty;

    /// <summary>Convergence status.</summary>
    public ConvergenceStatus Status { get; init; }

    /// <summary>Current value of the statistic.</summary>
    public double CurrentValue { get; init; }

    /// <summary>Relative change rate over the window.</summary>
    public double ChangeRate { get; init; }
}

/// <summary>
/// Convergence status enumeration.
/// </summary>
public enum ConvergenceStatus
{
    /// <summary>Not enough checkpoints to determine convergence.</summary>
    InsufficientData,

    /// <summary>Statistic has stabilized (change &lt; tolerance).</summary>
    Stable,

    /// <summary>Still changing but improving (change rate decreasing).</summary>
    Drifting,

    /// <summary>Change rate is not improving or increasing.</summary>
    Unstable
}
