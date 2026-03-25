namespace MonteCarlo.Addin.Excel;

/// <summary>
/// A cell tagged as a simulation input with its distribution configuration.
/// </summary>
public class TaggedInput
{
    /// <summary>
    /// The cell reference for this input.
    /// </summary>
    public required CellReference Cell { get; set; }

    /// <summary>
    /// Human-readable label (e.g., "Material Cost").
    /// </summary>
    public required string Label { get; set; }

    /// <summary>
    /// Distribution type name (e.g., "Normal", "Triangular").
    /// </summary>
    public required string DistributionName { get; set; }

    /// <summary>
    /// Distribution parameters (e.g., { "mean": 100, "stdDev": 10 }).
    /// </summary>
    public Dictionary<string, double> Parameters { get; set; } = new();
}
