using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Manages which cells are tagged as simulation outputs.
/// </summary>
public interface IOutputTagManager
{
    /// <summary>
    /// Tag a cell as a simulation output.
    /// </summary>
    void TagOutput(CellReference cell, string label);

    /// <summary>
    /// Remove a cell's output tag.
    /// </summary>
    void UntagOutput(CellReference cell);

    /// <summary>
    /// Get all tagged outputs.
    /// </summary>
    IReadOnlyList<TaggedOutput> GetAllOutputs();

    /// <summary>
    /// Check if a cell is tagged as an output.
    /// </summary>
    bool IsTagged(CellReference cell);

    /// <summary>
    /// Convert tagged outputs to engine SimulationOutput objects.
    /// </summary>
    List<SimulationOutput> ToSimulationOutputs();

    /// <summary>
    /// Remove all output tags.
    /// </summary>
    void Clear();
}
