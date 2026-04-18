using MonteCarlo.Engine.Analysis;

namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Event args raised when convergence is checked during a simulation run.
/// </summary>
public class ConvergenceEventArgs : EventArgs
{
    /// <summary>
    /// Convergence indicators for each tracked statistic.
    /// </summary>
    public required IReadOnlyList<ConvergenceIndicator> Indicators { get; init; }

    /// <summary>
    /// True if all indicators have reached <see cref="ConvergenceStatus.Stable"/>.
    /// </summary>
    public bool AllConverged { get; init; }

    /// <summary>
    /// The iteration at which convergence was checked.
    /// </summary>
    public int Iteration { get; init; }
}
