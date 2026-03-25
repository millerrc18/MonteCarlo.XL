namespace MonteCarlo.Addin.Excel;

/// <summary>
/// A cell tagged as a simulation output.
/// </summary>
public class TaggedOutput
{
    /// <summary>
    /// The cell reference for this output.
    /// </summary>
    public required CellReference Cell { get; set; }

    /// <summary>
    /// Human-readable label (e.g., "Net Profit").
    /// </summary>
    public required string Label { get; set; }
}
