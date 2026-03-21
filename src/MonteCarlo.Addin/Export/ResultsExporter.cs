using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Export;

/// <summary>
/// Exports simulation results to formatted Excel sheets with statistics tables,
/// sensitivity analysis, input assumptions, and embedded chart images.
/// </summary>
public class ResultsExporter
{
    private const string SheetPrefix = "MC Results — ";
    private const string RawDataSheetPrefix = "MC Raw Data — ";

    /// <summary>
    /// Export a complete summary sheet for one output.
    /// </summary>
    public void ExportSummary(
        SimulationResult result,
        SummaryStatistics stats,
        IReadOnlyList<SensitivityResult>? sensitivity,
        SimulationProfile profile,
        int outputIndex,
        byte[]? histogramImage = null,
        byte[]? tornadoImage = null)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        var output = result.Config.Outputs[outputIndex];
        string sheetName = TruncateSheetName($"{SheetPrefix}{output.Label}");

        // Create or clear the sheet
        var sheet = EnsureSheet(workbook, sheetName);

        try
        {
            app.ScreenUpdating = false;

            int row = 1;

            // Title section
            row = WriteTitleSection(sheet, row, output, result, profile);
            row++;

            // Summary statistics
            row = WriteSummaryStats(sheet, row, stats);
            row++;

            // Percentiles
            row = WritePercentiles(sheet, row, stats);
            row++;

            // Sensitivity
            if (sensitivity != null && sensitivity.Count > 0)
            {
                row = WriteSensitivity(sheet, row, sensitivity);
                row++;
            }

            // Input assumptions
            row = WriteInputAssumptions(sheet, row, profile);
            row++;

            // Auto-fit columns
            sheet.Columns["A:D"].AutoFit();

            // Embed chart images (right side)
            int imageColumn = 5; // Column E
            int imageRow = 1;

            if (histogramImage != null)
            {
                EmbedImage(sheet, histogramImage, imageRow, imageColumn, 500, 280);
                imageRow += 18;
            }

            if (tornadoImage != null)
            {
                EmbedImage(sheet, tornadoImage, imageRow, imageColumn, 500, 300);
            }

            // Sheet tab color (blue)
            sheet.Tab.Color = 0xF6823B; // #3B82F6 in BGR
        }
        finally
        {
            app.ScreenUpdating = true;
        }
    }

    /// <summary>
    /// Export raw simulation data to a data sheet.
    /// </summary>
    public void ExportRawData(SimulationResult result, int outputIndex)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        var output = result.Config.Outputs[outputIndex];
        string sheetName = TruncateSheetName($"{RawDataSheetPrefix}{output.Label}");
        var sheet = EnsureSheet(workbook, sheetName);

        try
        {
            app.ScreenUpdating = false;

            int colCount = result.Config.Inputs.Count + 1; // inputs + 1 output
            int rowCount = result.IterationCount;

            // Headers
            sheet.Cells[1, 1].Value2 = "Iteration";
            for (int j = 0; j < result.Config.Inputs.Count; j++)
                sheet.Cells[1, j + 2].Value2 = result.Config.Inputs[j].Label;
            sheet.Cells[1, colCount + 1].Value2 = output.Label;

            // Format header row
            var headerRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, colCount + 1]];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = 0xFAF8F8; // #F8FAFC surface color in BGR

            // Write data as 2D array for performance
            var data = new object[rowCount, colCount + 1];
            var inputSamples = new double[result.Config.Inputs.Count][];
            for (int j = 0; j < result.Config.Inputs.Count; j++)
                inputSamples[j] = result.GetInputSamples(result.Config.Inputs[j].Id);
            var outputValues = result.GetOutputValues(result.Config.Outputs[outputIndex].Id);

            for (int i = 0; i < rowCount; i++)
            {
                data[i, 0] = i + 1;
                for (int j = 0; j < result.Config.Inputs.Count; j++)
                    data[i, j + 1] = inputSamples[j][i];
                data[i, colCount] = outputValues[i];
            }

            var dataRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[rowCount + 1, colCount + 1]];
            dataRange.Value2 = data;

            sheet.Columns[$"A:{GetColumnLetter(colCount + 1)}"].AutoFit();
            sheet.Tab.Color = 0xB981F5; // Violet in BGR
        }
        finally
        {
            app.ScreenUpdating = true;
        }
    }

    #region Private Helpers

    private static int WriteTitleSection(Worksheet sheet, int row, SimulationOutput output,
        SimulationResult result, SimulationProfile profile)
    {
        var titleCell = sheet.Cells[row, 1];
        titleCell.Value2 = "MonteCarlo.XL Simulation Results";
        titleCell.Font.Size = 16;
        titleCell.Font.Bold = true;
        row++;

        sheet.Cells[row, 1].Value2 = $"Output: {output.Label}";
        sheet.Cells[row, 1].Font.Size = 12;
        row++;

        sheet.Cells[row, 1].Value2 = $"Iterations: {result.IterationCount:N0}";
        sheet.Cells[row, 2].Value2 = $"Date: {DateTime.Now:yyyy-MM-dd}";
        row++;

        return row;
    }

    private static int WriteSummaryStats(Worksheet sheet, int row, SummaryStatistics stats)
    {
        WriteSectionHeader(sheet, row, "SUMMARY STATISTICS");
        row++;

        WriteTableHeader(sheet, row, "Statistic", "Value");
        row++;

        var items = new (string Label, double Value)[]
        {
            ("Mean", stats.Mean),
            ("Median", stats.Median),
            ("Std Dev", stats.StdDev),
            ("Variance", stats.Variance),
            ("Minimum", stats.Min),
            ("Maximum", stats.Max),
            ("Skewness", stats.Skewness),
            ("Kurtosis", stats.Kurtosis)
        };

        foreach (var (label, value) in items)
        {
            sheet.Cells[row, 1].Value2 = label;
            sheet.Cells[row, 2].Value2 = value;
            sheet.Cells[row, 2].NumberFormat = GetNumberFormat(value);

            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 2);
            row++;
        }

        return row;
    }

    private static int WritePercentiles(Worksheet sheet, int row, SummaryStatistics stats)
    {
        WriteSectionHeader(sheet, row, "PERCENTILES");
        row++;

        WriteTableHeader(sheet, row, "Percentile", "Value");
        row++;

        var percentiles = new (string Label, double Value)[]
        {
            ("P1", stats.P1), ("P5", stats.P5), ("P10", stats.P10),
            ("P25", stats.P25), ("P50", stats.Median),
            ("P75", stats.P75), ("P90", stats.P90), ("P95", stats.P95), ("P99", stats.P99)
        };

        foreach (var (label, value) in percentiles)
        {
            sheet.Cells[row, 1].Value2 = label;
            sheet.Cells[row, 2].Value2 = value;
            sheet.Cells[row, 2].NumberFormat = GetNumberFormat(value);

            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 2);
            row++;
        }

        return row;
    }

    private static int WriteSensitivity(Worksheet sheet, int row, IReadOnlyList<SensitivityResult> results)
    {
        WriteSectionHeader(sheet, row, $"SENSITIVITY (TOP {Math.Min(results.Count, 10)} INPUTS)");
        row++;

        sheet.Cells[row, 1].Value2 = "Input";
        sheet.Cells[row, 2].Value2 = "Rank Corr";
        sheet.Cells[row, 3].Value2 = "Swing";
        sheet.Cells[row, 4].Value2 = "% Variance";
        FormatHeaderRow(sheet, row, 4);
        row++;

        int count = Math.Min(results.Count, 10);
        for (int i = 0; i < count; i++)
        {
            var r = results[i];
            sheet.Cells[row, 1].Value2 = r.InputLabel;
            sheet.Cells[row, 2].Value2 = r.RankCorrelation;
            sheet.Cells[row, 2].NumberFormat = "0.000";
            sheet.Cells[row, 3].Value2 = r.Swing;
            sheet.Cells[row, 3].NumberFormat = GetNumberFormat(r.Swing);
            sheet.Cells[row, 4].Value2 = r.ContributionToVariance / 100.0;
            sheet.Cells[row, 4].NumberFormat = "0.0%";

            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 4);
            row++;
        }

        return row;
    }

    private static int WriteInputAssumptions(Worksheet sheet, int row, SimulationProfile profile)
    {
        WriteSectionHeader(sheet, row, "INPUT ASSUMPTIONS");
        row++;

        sheet.Cells[row, 1].Value2 = "Input";
        sheet.Cells[row, 2].Value2 = "Cell";
        sheet.Cells[row, 3].Value2 = "Distribution";
        FormatHeaderRow(sheet, row, 3);
        row++;

        foreach (var input in profile.Inputs)
        {
            sheet.Cells[row, 1].Value2 = input.Label;
            sheet.Cells[row, 2].Value2 = $"{input.SheetName}!{input.CellAddress}";
            sheet.Cells[row, 3].Value2 = FormatDistribution(input);

            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 3);
            row++;
        }

        return row;
    }

    private static void WriteSectionHeader(Worksheet sheet, int row, string title)
    {
        var cell = sheet.Cells[row, 1];
        cell.Value2 = title;
        cell.Font.Size = 14;
        cell.Font.Bold = true;

        var range = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 4]];
        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
        range.Borders[XlBordersIndex.xlEdgeBottom].Color = 0xE0E8E2; // #E2E8F0 border in BGR
    }

    private static void WriteTableHeader(Worksheet sheet, int row, params string[] headers)
    {
        for (int i = 0; i < headers.Length; i++)
            sheet.Cells[row, i + 1].Value2 = headers[i];
        FormatHeaderRow(sheet, row, headers.Length);
    }

    private static void FormatHeaderRow(Worksheet sheet, int row, int colCount)
    {
        var range = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, colCount]];
        range.Font.Bold = true;
        range.Interior.Color = 0xFAF8F8; // #F8FAFC
    }

    private static void ApplyAltRowShading(Worksheet sheet, int row, int colCount)
    {
        var range = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, colCount]];
        range.Interior.Color = 0xFAF8F8; // #F8FAFC
    }

    private static void EmbedImage(Worksheet sheet, byte[] imageBytes, int row, int col,
        int widthPx, int heightPx)
    {
        string tempPath = ChartImageRenderer.SaveToTempFile(imageBytes);
        try
        {
            var cell = (Range)sheet.Cells[row, col];
            double left = cell.Left;
            double top = cell.Top;

            // Convert pixels to points (72 pts per inch, 96 px per inch)
            double widthPts = widthPx * 72.0 / 96.0;
            double heightPts = heightPx * 72.0 / 96.0;

            sheet.Shapes.AddPicture(
                tempPath,
                MsoTriState.msoFalse,   // LinkToFile
                MsoTriState.msoCTrue,    // SaveWithDocument
                (float)left, (float)top,
                (float)widthPts, (float)heightPts);
        }
        finally
        {
            try { File.Delete(tempPath); } catch { }
        }
    }

    private static Worksheet EnsureSheet(Workbook workbook, string name)
    {
        // Check if sheet exists
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name == name)
            {
                ws.Cells.Clear();
                return ws;
            }
        }

        // Create new sheet at the end
        var lastSheet = workbook.Worksheets[workbook.Worksheets.Count];
        var newSheet = (Worksheet)workbook.Worksheets.Add(After: lastSheet);
        newSheet.Name = name;
        return newSheet;
    }

    private static string TruncateSheetName(string name)
    {
        // Excel sheet names max 31 chars
        return name.Length > 31 ? name[..31] : name;
    }

    private static string GetNumberFormat(double value)
    {
        double abs = Math.Abs(value);
        if (abs >= 1000) return "#,##0";
        if (abs >= 1) return "0.00";
        if (abs >= 0.01) return "0.000";
        return "0.0000";
    }

    private static string FormatDistribution(SavedInput input)
    {
        var p = input.Parameters;
        return input.DistributionName.ToLowerInvariant() switch
        {
            "normal" => $"Normal(μ={p.GetValueOrDefault("mean")}, σ={p.GetValueOrDefault("stdDev")})",
            "triangular" => $"Triangular({p.GetValueOrDefault("min")}, {p.GetValueOrDefault("mode")}, {p.GetValueOrDefault("max")})",
            "pert" => $"PERT({p.GetValueOrDefault("min")}, {p.GetValueOrDefault("mode")}, {p.GetValueOrDefault("max")})",
            "lognormal" => $"Lognormal(μ={p.GetValueOrDefault("mu")}, σ={p.GetValueOrDefault("sigma")})",
            "uniform" => $"Uniform({p.GetValueOrDefault("min")}, {p.GetValueOrDefault("max")})",
            "discrete" => "Discrete(...)",
            _ => input.DistributionName
        };
    }

    private static string GetColumnLetter(int col)
    {
        string result = string.Empty;
        while (col > 0)
        {
            col--;
            result = (char)('A' + col % 26) + result;
            col /= 26;
        }
        return result;
    }

    #endregion
}
