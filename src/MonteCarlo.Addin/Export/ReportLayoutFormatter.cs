using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Export;

internal static class ReportLayoutFormatter
{
    public static void ApplyPdfFriendlyLayout(
        Worksheet sheet,
        string reportTitle,
        int lastRow,
        int lastColumn,
        IReadOnlyList<int>? manualPageBreakRows = null)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        lastRow = Math.Max(1, lastRow);
        lastColumn = Math.Max(1, lastColumn);

        var app = (Application)sheet.Application;
        var pageSetup = sheet.PageSetup;

        pageSetup.PrintArea = sheet.Range[sheet.Cells[1, 1], sheet.Cells[lastRow, lastColumn]]
            .Address[true, true, XlReferenceStyle.xlA1];
        pageSetup.Orientation = XlPageOrientation.xlLandscape;
        pageSetup.PaperSize = XlPaperSize.xlPaperLetter;
        pageSetup.Zoom = false;
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = false;
        pageSetup.CenterHorizontally = true;
        pageSetup.CenterVertically = false;
        pageSetup.Order = XlOrder.xlOverThenDown;
        pageSetup.PrintGridlines = false;
        pageSetup.PrintHeadings = false;
        pageSetup.PrintTitleRows = "$1:$3";
        pageSetup.LeftMargin = app.InchesToPoints(0.35);
        pageSetup.RightMargin = app.InchesToPoints(0.35);
        pageSetup.TopMargin = app.InchesToPoints(0.45);
        pageSetup.BottomMargin = app.InchesToPoints(0.45);
        pageSetup.HeaderMargin = app.InchesToPoints(0.2);
        pageSetup.FooterMargin = app.InchesToPoints(0.2);
        pageSetup.CenterHeader = reportTitle;
        pageSetup.LeftFooter = "MonteCarlo.XL";
        pageSetup.RightFooter = "Page &P of &N";

        try
        {
            sheet.ResetAllPageBreaks();
        }
        catch
        {
            // Ignore page-break reset issues on protected or transient sheets.
        }

        if (manualPageBreakRows != null)
        {
            foreach (var breakRow in manualPageBreakRows
                         .Where(row => row > 1 && row < lastRow)
                         .Distinct()
                         .OrderBy(row => row))
            {
                try
                {
                    sheet.HPageBreaks.Add(sheet.Cells[breakRow, 1]);
                }
                catch
                {
                    // Ignore duplicate or invalid manual break locations.
                }
            }
        }

        try
        {
            sheet.DisplayPageBreaks = false;
        }
        catch
        {
            // Some Excel states reject toggling page-break display; layout still applies.
        }
    }
}
