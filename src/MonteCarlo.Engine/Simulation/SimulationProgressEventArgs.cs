using MonteCarlo.Engine.Analysis;

namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Reports progress of a running simulation.
/// </summary>
public class SimulationProgressEventArgs : EventArgs
{
    /// <summary>
    /// Number of iterations completed so far.
    /// </summary>
    public int CompletedIterations { get; init; }

    /// <summary>
    /// Total number of iterations in this run.
    /// </summary>
    public int TotalIterations { get; init; }

    /// <summary>
    /// Completion percentage (0.0 to 1.0).
    /// </summary>
    public double PercentComplete => TotalIterations > 0
        ? (double)CompletedIterations / TotalIterations
        : 0.0;

    /// <summary>
    /// Time elapsed since the simulation started.
    /// </summary>
    public TimeSpan Elapsed { get; init; }

    /// <summary>
    /// Estimated time remaining based on current pace.
    /// </summary>
    public TimeSpan EstimatedRemaining { get; init; }

    /// <summary>
    /// Interim histogram data computed periodically during the simulation.
    /// </summary>
    public HistogramData? InterimHistogram { get; init; }

    /// <summary>
    /// Interim sorted output values used to build the histogram.
    /// </summary>
    public double[]? InterimSortedValues { get; init; }
}
