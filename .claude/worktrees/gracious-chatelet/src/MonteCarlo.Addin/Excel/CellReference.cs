namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Represents a reference to a specific cell in a workbook.
/// </summary>
public class CellReference
{
    /// <summary>
    /// The worksheet name (e.g., "Sheet1").
    /// </summary>
    public required string SheetName { get; set; }

    /// <summary>
    /// The cell address in A1 notation (e.g., "B4").
    /// </summary>
    public required string CellAddress { get; set; }

    /// <summary>
    /// Full reference string (e.g., "'Sheet1'!B4").
    /// </summary>
    public string FullReference => $"'{SheetName}'!{CellAddress}";

    public override bool Equals(object? obj) =>
        obj is CellReference other &&
        string.Equals(SheetName, other.SheetName, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(CellAddress, other.CellAddress, StringComparison.OrdinalIgnoreCase);

    public override int GetHashCode() =>
        HashCode.Combine(
            SheetName.ToUpperInvariant(),
            CellAddress.ToUpperInvariant());

    public override string ToString() => FullReference;
}
