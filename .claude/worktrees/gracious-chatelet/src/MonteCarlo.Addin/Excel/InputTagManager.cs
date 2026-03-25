using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Tracks which cells the user has designated as simulation inputs.
/// </summary>
public class InputTagManager : IInputTagManager
{
    private readonly Dictionary<CellReference, TaggedInput> _inputs = new();

    /// <inheritdoc />
    public void TagInput(CellReference cell, string label, string distributionName, Dictionary<string, double> parameters)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentException.ThrowIfNullOrWhiteSpace(label);
        ArgumentException.ThrowIfNullOrWhiteSpace(distributionName);

        _inputs[cell] = new TaggedInput
        {
            Cell = cell,
            Label = label,
            DistributionName = distributionName,
            Parameters = parameters ?? new()
        };
    }

    /// <inheritdoc />
    public void UntagInput(CellReference cell)
    {
        _inputs.Remove(cell);
    }

    /// <inheritdoc />
    public IReadOnlyList<TaggedInput> GetAllInputs() => _inputs.Values.ToList();

    /// <inheritdoc />
    public bool IsTagged(CellReference cell) => _inputs.ContainsKey(cell);

    /// <inheritdoc />
    public List<SimulationInput> ToSimulationInputs(IWorkbookManager workbook)
    {
        var result = new List<SimulationInput>();

        foreach (var tagged in _inputs.Values)
        {
            double baseValue = workbook.ReadCellValue(tagged.Cell.SheetName, tagged.Cell.CellAddress);
            var distribution = DistributionFactory.Create(tagged.DistributionName, tagged.Parameters);

            result.Add(new SimulationInput
            {
                Id = tagged.Cell.FullReference,
                Label = tagged.Label,
                Distribution = distribution,
                BaseValue = baseValue
            });
        }

        return result;
    }

    /// <inheritdoc />
    public void Clear() => _inputs.Clear();
}
