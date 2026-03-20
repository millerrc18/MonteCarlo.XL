using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Tracks which cells the user has designated as simulation outputs.
/// </summary>
public class OutputTagManager : IOutputTagManager
{
    private readonly Dictionary<CellReference, TaggedOutput> _outputs = new();

    /// <inheritdoc />
    public void TagOutput(CellReference cell, string label)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentException.ThrowIfNullOrWhiteSpace(label);

        _outputs[cell] = new TaggedOutput
        {
            Cell = cell,
            Label = label
        };
    }

    /// <inheritdoc />
    public void UntagOutput(CellReference cell)
    {
        _outputs.Remove(cell);
    }

    /// <inheritdoc />
    public IReadOnlyList<TaggedOutput> GetAllOutputs() => _outputs.Values.ToList();

    /// <inheritdoc />
    public bool IsTagged(CellReference cell) => _outputs.ContainsKey(cell);

    /// <inheritdoc />
    public List<SimulationOutput> ToSimulationOutputs()
    {
        return _outputs.Values.Select(tagged => new SimulationOutput
        {
            Id = tagged.Cell.FullReference,
            Label = tagged.Label
        }).ToList();
    }

    /// <inheritdoc />
    public void Clear() => _outputs.Clear();
}
