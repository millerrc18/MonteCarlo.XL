using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Implements workbook read/write operations via ExcelDna COM interop.
/// </summary>
public class WorkbookManager : IWorkbookManager
{
    private Application App => (Application)ExcelDnaUtil.Application;

    /// <inheritdoc />
    public double ReadCellValue(string sheetName, string cellAddress)
    {
        var sheet = GetSheet(sheetName);
        var range = (Range)sheet.Range[cellAddress];
        object? value = range.Value2;

        if (value is double d)
            return d;

        if (value is int i)
            return i;

        throw new InvalidOperationException(
            $"Cell '{sheetName}'!{cellAddress} does not contain a numeric value (current value: {value ?? "empty"}).");
    }

    /// <inheritdoc />
    public Dictionary<string, double> ReadCellValues(IEnumerable<CellReference> cells)
    {
        var results = new Dictionary<string, double>();

        // Group by sheet to minimize COM round-trips
        foreach (var group in cells.GroupBy(c => c.SheetName, StringComparer.OrdinalIgnoreCase))
        {
            var sheet = GetSheet(group.Key);
            foreach (var cell in group)
            {
                var range = (Range)sheet.Range[cell.CellAddress];
                object? value = range.Value2;

                if (value is double d)
                    results[cell.FullReference] = d;
                else if (value is int i)
                    results[cell.FullReference] = i;
                else
                    throw new InvalidOperationException(
                        $"Cell {cell.FullReference} does not contain a numeric value (current value: {value ?? "empty"}).");
            }
        }

        return results;
    }

    /// <inheritdoc />
    public void WriteCellValue(string sheetName, string cellAddress, double value)
    {
        var sheet = GetSheet(sheetName);
        var range = (Range)sheet.Range[cellAddress];
        range.Value2 = value;
    }

    /// <inheritdoc />
    public void WriteRange(string sheetName, string topLeftAddress, double[,] values)
    {
        var sheet = GetSheet(sheetName);
        int rows = values.GetLength(0);
        int cols = values.GetLength(1);
        var topLeft = (Range)sheet.Range[topLeftAddress];
        var range = topLeft.Resize[rows, cols];

        var app = App;
        using var excelState = ExcelStateScope.Capture(app, "Write range");
        excelState.Apply(screenUpdating: false, calculation: XlCalculation.xlCalculationManual);

        range.Value2 = values;
    }

    /// <inheritdoc />
    public void WriteResultsSheet(string sheetName, string[] headers, double[,] data)
    {
        EnsureSheet(sheetName);
        var sheet = GetSheet(sheetName);

        var app = App;
        using var excelState = ExcelStateScope.Capture(app, "Write results sheet");
        excelState.Apply(screenUpdating: false, calculation: XlCalculation.xlCalculationManual);

        // Write headers in row 1
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = (Range)sheet.Cells[1, c + 1];
            cell.Value2 = headers[c];
            cell.Font.Bold = true;
        }

        // Write data starting at row 2
        int rows = data.GetLength(0);
        int cols = data.GetLength(1);
        if (rows > 0 && cols > 0)
        {
            var topLeft = (Range)sheet.Cells[2, 1];
            var range = topLeft.Resize[rows, cols];
            range.Value2 = data;
        }
    }

    /// <inheritdoc />
    public CellReference? GetActiveCell()
    {
        try
        {
            var cell = App.ActiveCell;
            if (cell == null) return null;

            var sheet = (Worksheet)cell.Worksheet;
            return new CellReference
            {
                SheetName = sheet.Name,
                CellAddress = cell.Address[false, false] // No absolute references ($)
            };
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc />
    public bool SheetExists(string sheetName)
    {
        try
        {
            var workbook = App.ActiveWorkbook;
            if (workbook == null) return false;

            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc />
    public void EnsureSheet(string sheetName, bool clearIfExists = true)
    {
        var workbook = App.ActiveWorkbook
            ?? throw new InvalidOperationException("No active workbook.");

        if (SheetExists(sheetName))
        {
            if (clearIfExists)
            {
                var sheet = (Worksheet)workbook.Sheets[sheetName];
                sheet.Cells.Clear();
            }
        }
        else
        {
            var newSheet = (Worksheet)workbook.Sheets.Add();
            newSheet.Name = sheetName;
        }
    }

    private Worksheet GetSheet(string sheetName)
    {
        var workbook = App.ActiveWorkbook
            ?? throw new InvalidOperationException("No active workbook.");

        try
        {
            return (Worksheet)workbook.Sheets[sheetName];
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            throw new ArgumentException($"Sheet '{sheetName}' does not exist in the active workbook.");
        }
    }
}
