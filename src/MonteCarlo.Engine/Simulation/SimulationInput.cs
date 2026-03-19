using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Defines a single uncertain input for a Monte Carlo simulation.
/// </summary>
public class SimulationInput
{
    /// <summary>
    /// Unique identifier (e.g., cell reference "B4").
    /// </summary>
    public required string Id { get; set; }

    /// <summary>
    /// Human-readable label (e.g., "Material Cost").
    /// </summary>
    public required string Label { get; set; }

    /// <summary>
    /// The probability distribution to sample from.
    /// </summary>
    public required IDistribution Distribution { get; set; }

    /// <summary>
    /// The deterministic value currently in the cell (used as baseline for sensitivity).
    /// </summary>
    public double BaseValue { get; set; }
}
