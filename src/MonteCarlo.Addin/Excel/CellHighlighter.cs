using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Applies visual formatting to tagged cells via Excel COM interop.
/// </summary>
public class CellHighlighter : ICellHighlighter
{
    // Light blue: #DBEAFE → RGB(219, 234, 254)
    private static readonly int InputColor = ColorToOle(219, 234, 254);

    // Light green: #DCFCE7 → RGB(220, 252, 231)
    private static readonly int OutputColor = ColorToOle(220, 252, 231);

    private Application App => (Application)ExcelDnaUtil.Application;

    /// <inheritdoc />
    public void HighlightInput(CellReference cell)
    {
        SetCellColor(cell, InputColor);
    }

    /// <inheritdoc />
    public void HighlightOutput(CellReference cell)
    {
        SetCellColor(cell, OutputColor);
    }

    /// <inheritdoc />
    public void ClearHighlight(CellReference cell)
    {
        try
        {
            var sheet = (Worksheet)App.ActiveWorkbook.Sheets[cell.SheetName];
            var range = (Range)sheet.Range[cell.CellAddress];
            range.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
        }
        catch
        {
            // Silently ignore if the cell/sheet no longer exists
        }
    }

    /// <inheritdoc />
    public void RefreshAll(IInputTagManager inputs, IOutputTagManager outputs)
    {
        var app = App;
        app.ScreenUpdating = false;
        try
        {
            foreach (var input in inputs.GetAllInputs())
                HighlightInput(input.Cell);

            foreach (var output in outputs.GetAllOutputs())
                HighlightOutput(output.Cell);
        }
        finally
        {
            app.ScreenUpdating = true;
        }
    }

    private void SetCellColor(CellReference cell, int oleColor)
    {
        try
        {
            var sheet = (Worksheet)App.ActiveWorkbook.Sheets[cell.SheetName];
            var range = (Range)sheet.Range[cell.CellAddress];
            range.Interior.Color = oleColor;
        }
        catch
        {
            // Silently ignore if the cell/sheet no longer exists
        }
    }

    /// <summary>
    /// Converts RGB values to OLE color format used by Excel.
    /// </summary>
    private static int ColorToOle(int r, int g, int b) => r | (g << 8) | (b << 16);
}
