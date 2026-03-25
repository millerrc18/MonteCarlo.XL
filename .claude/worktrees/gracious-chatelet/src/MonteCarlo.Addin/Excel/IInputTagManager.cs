using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Manages which cells are tagged as simulation inputs.
/// </summary>
public interface IInputTagManager
{
    /// <summary>
    /// Tag a cell as a simulation input with a specific distribution.
    /// </summary>
    void TagInput(CellReference cell, string label, string distributionName, Dictionary<string, double> parameters);

    /// <summary>
    /// Remove a cell's input tag.
    /// </summary>
    void UntagInput(CellReference cell);

    /// <summary>
    /// Get all tagged inputs.
    /// </summary>
    IReadOnlyList<TaggedInput> GetAllInputs();

    /// <summary>
    /// Check if a cell is tagged as an input.
    /// </summary>
    bool IsTagged(CellReference cell);

    /// <summary>
    /// Convert tagged inputs to engine SimulationInput objects.
    /// Reads current cell values as base values.
    /// </summary>
    List<SimulationInput> ToSimulationInputs(IWorkbookManager workbook);

    /// <summary>
    /// Remove all input tags.
    /// </summary>
    void Clear();
}
