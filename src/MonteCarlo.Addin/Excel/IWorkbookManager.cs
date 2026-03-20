namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Abstracts all Excel workbook read/write operations for testability.
/// </summary>
public interface IWorkbookManager
{
    /// <summary>
    /// Read the current computed value of a cell. Throws if the cell does not contain a numeric value.
    /// </summary>
    double ReadCellValue(string sheetName, string cellAddress);

    /// <summary>
    /// Read multiple cell values in a batch (minimizes COM round-trips).
    /// </summary>
    Dictionary<string, double> ReadCellValues(IEnumerable<CellReference> cells);

    /// <summary>
    /// Write a value to a cell.
    /// </summary>
    void WriteCellValue(string sheetName, string cellAddress, double value);

    /// <summary>
    /// Write a block of values to a range starting at <paramref name="topLeftAddress"/>.
    /// </summary>
    void WriteRange(string sheetName, string topLeftAddress, double[,] values);

    /// <summary>
    /// Write column headers and data to a new or existing sheet.
    /// </summary>
    void WriteResultsSheet(string sheetName, string[] headers, double[,] data);

    /// <summary>
    /// Get the currently selected cell's address and sheet name. Null if no cell is selected.
    /// </summary>
    CellReference? GetActiveCell();

    /// <summary>
    /// Check if a sheet exists in the active workbook.
    /// </summary>
    bool SheetExists(string sheetName);

    /// <summary>
    /// Create a new sheet, or clear an existing one.
    /// </summary>
    void EnsureSheet(string sheetName, bool clearIfExists = true);
}
