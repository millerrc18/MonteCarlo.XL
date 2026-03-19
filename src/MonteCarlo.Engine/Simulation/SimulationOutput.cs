namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Defines a single output to track during a Monte Carlo simulation.
/// </summary>
public class SimulationOutput
{
    /// <summary>
    /// Unique identifier (e.g., cell reference "D10").
    /// </summary>
    public required string Id { get; set; }

    /// <summary>
    /// Human-readable label (e.g., "Net Profit").
    /// </summary>
    public required string Label { get; set; }
}
