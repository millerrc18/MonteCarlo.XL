using System.Diagnostics;
using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Engine.Simulation;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Services;

/// <summary>
/// Runs lightweight local performance diagnostics and writes them to Excel.
/// </summary>
internal sealed class PerformanceBenchmarkService
{
    private const string BenchmarkSheetPrefix = "MC Benchmark";

    public async Task RunAsync()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook is available.");

        using var excelState = ExcelStateScope.Capture(app, "Performance benchmark", restoreSelection: true);
        excelState.Apply(
            screenUpdating: false,
            enableEvents: false,
            statusBar: "MonteCarlo.XL: running benchmark diagnostics...");

        var workbookMetrics = MeasureWorkbookRecalc(app, workbook);
        var engineMetrics = await SimulationBenchmark.RunAsync(
            inputCount: Math.Max(5, Math.Min(50, workbookMetrics.WorksheetCount * 10)),
            iterationCount: 10_000).ConfigureAwait(false);

        WriteBenchmarkSheet(app, workbookMetrics, engineMetrics);
        StartupDiagnostics.Log(
            $"Performance benchmark completed. WorkbookRecalcMs={workbookMetrics.RecalcTime.TotalMilliseconds:0.0}, " +
            $"EngineIterationsPerSecond={engineMetrics.IterationsPerSecond:0}.");
    }

    private static WorkbookBenchmarkResult MeasureWorkbookRecalc(Application app, Workbook workbook)
    {
        var sw = Stopwatch.StartNew();
        app.CalculateFullRebuild();
        sw.Stop();

        return new WorkbookBenchmarkResult(
            WorkbookName: workbook.Name,
            WorksheetCount: workbook.Worksheets.Count,
            UsedCellCount: CountUsedCells(workbook),
            RecalcTime: sw.Elapsed);
    }

    private static long CountUsedCells(Workbook workbook)
    {
        long total = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                Range usedRange = sheet.UsedRange;
                total += Convert.ToInt64(usedRange.CountLarge);
            }
            catch { }
        }

        return total;
    }

    private static void WriteBenchmarkSheet(
        Application app,
        WorkbookBenchmarkResult workbookMetrics,
        SimulationBenchmark.BenchmarkResult engineMetrics)
    {
        var workbook = app.ActiveWorkbook;
        if (workbook == null)
            return;

        var sheetName = GetUniqueSheetName(workbook, BenchmarkSheetPrefix);
        var lastSheet = workbook.Worksheets[workbook.Worksheets.Count];
        var sheet = (Worksheet)workbook.Worksheets.Add(After: lastSheet);
        sheet.Name = sheetName;

        var row = 1;
        sheet.Cells[row, 1].Value2 = "MonteCarlo.XL Performance Benchmark";
        sheet.Cells[row, 1].Font.Bold = true;
        sheet.Cells[row, 1].Font.Size = 16;
        row += 2;

        sheet.Cells[row, 1].Value2 = "Workbook Recalc";
        sheet.Cells[row, 1].Font.Bold = true;
        row++;
        row = WriteMetric(sheet, row, "Workbook", workbookMetrics.WorkbookName);
        row = WriteMetric(sheet, row, "Worksheets", workbookMetrics.WorksheetCount.ToString("N0"));
        row = WriteMetric(sheet, row, "Used cells", workbookMetrics.UsedCellCount.ToString("N0"));
        row = WriteMetric(sheet, row, "Full rebuild recalc", $"{workbookMetrics.RecalcTime.TotalSeconds:0.000} seconds");
        row++;

        sheet.Cells[row, 1].Value2 = "Synthetic Engine Benchmark";
        sheet.Cells[row, 1].Font.Bold = true;
        row++;
        row = WriteMetric(sheet, row, "Inputs", engineMetrics.InputCount.ToString("N0"));
        row = WriteMetric(sheet, row, "Iterations", engineMetrics.IterationCount.ToString("N0"));
        row = WriteMetric(sheet, row, "Total time", $"{engineMetrics.TotalTime.TotalSeconds:0.000} seconds");
        row = WriteMetric(sheet, row, "Iterations/sec", engineMetrics.IterationsPerSecond.ToString("N0"));
        row = WriteMetric(sheet, row, "Microseconds/iteration", engineMetrics.MicrosecondsPerIteration.ToString("0.00"));
        row++;

        sheet.Cells[row, 1].Value2 = "Use this sheet to compare workbook recalc cost against raw engine throughput. If workbook recalc dominates, lower iteration counts while modeling and use Full or Deep presets only for final reports.";

        sheet.Columns["A:B"].AutoFit();
        sheet.Tab.Color = 0x81E6D9;
    }

    private static int WriteMetric(Worksheet sheet, int row, string label, string value)
    {
        sheet.Cells[row, 1].Value2 = label;
        sheet.Cells[row, 1].Font.Bold = true;
        sheet.Cells[row, 2].Value2 = value;
        return row + 1;
    }

    private static string GetUniqueSheetName(Workbook workbook, string baseName)
    {
        if (!SheetExists(workbook, baseName))
            return baseName;

        for (var suffix = 2; suffix < 1000; suffix++)
        {
            var suffixText = $" ({suffix})";
            var candidateBase = baseName.Length + suffixText.Length > 31
                ? baseName[..(31 - suffixText.Length)]
                : baseName;
            var candidate = $"{candidateBase}{suffixText}";
            if (!SheetExists(workbook, candidate))
                return candidate;
        }

        return $"{baseName[..Math.Min(baseName.Length, 24)]} {DateTime.Now:HHmmss}";
    }

    private static bool SheetExists(Workbook workbook, string name)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private sealed record WorkbookBenchmarkResult(
        string WorkbookName,
        int WorksheetCount,
        long UsedCellCount,
        TimeSpan RecalcTime);
}
