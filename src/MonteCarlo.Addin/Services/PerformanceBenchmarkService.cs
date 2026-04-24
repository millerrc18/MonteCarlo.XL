using System.Diagnostics;
using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.Export;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.Services;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Services;

/// <summary>
/// Runs lightweight local performance diagnostics and writes them to Excel.
/// </summary>
internal sealed class PerformanceBenchmarkService
{
    private const string BenchmarkSheetPrefix = "MC Benchmark";
    private const int ExportBenchmarkInputCount = 8;
    private const int ExportBenchmarkIterationCount = 2_000;

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
        var exportMetrics = MeasureExportCosts(app, workbook);

        WriteBenchmarkSheet(app, workbookMetrics, engineMetrics, exportMetrics);
        StartupDiagnostics.Log(
            $"Performance benchmark completed. WorkbookRecalcMs={workbookMetrics.RecalcTime.TotalMilliseconds:0.0}, " +
            $"EngineIterationsPerSecond={engineMetrics.IterationsPerSecond:0}, " +
            $"SummaryExportMs={exportMetrics.SummaryExportTime.TotalMilliseconds:0.0}, " +
            $"RawExportMs={exportMetrics.RawExportTime.TotalMilliseconds:0.0}.");
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
        SimulationBenchmark.BenchmarkResult engineMetrics,
        ExportBenchmarkResult exportMetrics)
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

        sheet.Cells[row, 1].Value2 = "Synthetic Export Cost";
        sheet.Cells[row, 1].Font.Bold = true;
        row++;
        row = WriteMetric(sheet, row, "Synthetic inputs", exportMetrics.InputCount.ToString("N0"));
        row = WriteMetric(sheet, row, "Synthetic iterations", exportMetrics.IterationCount.ToString("N0"));
        row = WriteMetric(sheet, row, "Summary export", $"{exportMetrics.SummaryExportTime.TotalSeconds:0.000} seconds");
        row = WriteMetric(sheet, row, "Raw data export", $"{exportMetrics.RawExportTime.TotalSeconds:0.000} seconds");
        row++;

        sheet.Cells[row, 1].Value2 =
            "Use this sheet to compare workbook recalc cost against raw engine throughput and export overhead. " +
            "If workbook recalc dominates, lower iteration counts while modeling and use Full or Deep presets only for final reports. " +
            "If export cost dominates, keep raw-data exports smaller and prefer summary reports during iteration.";

        sheet.Columns["A:B"].AutoFit();
        sheet.Tab.Color = 0x81E6D9;
    }

    private static ExportBenchmarkResult MeasureExportCosts(Application app, Workbook workbook)
    {
        var benchmarkRun = BuildSyntheticExportRun();
        var exporter = new ResultsExporter();
        var defaultSettings = UserSettings.Default;

        var summaryExportTime = MeasureTemporaryExport(
            app,
            workbook,
            () => exporter.ExportSummary(
                benchmarkRun.Result,
                benchmarkRun.Stats,
                benchmarkRun.Sensitivity,
                benchmarkRun.Profile,
                outputIndex: 0,
                createNewSheet: true,
                percentiles: defaultSettings.GetDefaultPercentileFractions(),
                targetValue: benchmarkRun.Stats.Mean,
                effectiveSettings: defaultSettings,
                usesWorkbookOverrides: false));

        var rawExportTime = MeasureTemporaryExport(
            app,
            workbook,
            () => exporter.ExportRawData(benchmarkRun.Result, outputIndex: 0, createNewSheet: true));

        return new ExportBenchmarkResult(
            benchmarkRun.Result.Config.Inputs.Count,
            benchmarkRun.Result.IterationCount,
            summaryExportTime,
            rawExportTime);
    }

    private static TimeSpan MeasureTemporaryExport(Application app, Workbook workbook, System.Action exportAction)
    {
        var sheetCountBefore = workbook.Worksheets.Count;
        var sw = Stopwatch.StartNew();
        exportAction();
        sw.Stop();

        if (workbook.Worksheets.Count > sheetCountBefore)
        {
            var createdSheet = (Worksheet)workbook.Worksheets[workbook.Worksheets.Count];
            DeleteWorksheet(app, createdSheet);
        }

        return sw.Elapsed;
    }

    private static void DeleteWorksheet(Application app, Worksheet sheet)
    {
        using var excelState = ExcelStateScope.Capture(app, "Delete benchmark export sheet");
        excelState.Apply(displayAlerts: false, screenUpdating: false);
        sheet.Delete();
    }

    private static SyntheticBenchmarkRun BuildSyntheticExportRun()
    {
        var config = new SimulationConfig
        {
            IterationCount = ExportBenchmarkIterationCount,
            RandomSeed = 42,
            Sampling = SamplingMethod.LatinHypercube,
            ParallelExecution = false
        };

        var profile = new SimulationProfile
        {
            Name = "Benchmark Synthetic",
            IterationCount = ExportBenchmarkIterationCount,
            RandomSeed = 42
        };

        var inputMatrix = new double[ExportBenchmarkIterationCount, ExportBenchmarkInputCount];
        var outputMatrix = new double[ExportBenchmarkIterationCount, 1];
        for (var inputIndex = 0; inputIndex < ExportBenchmarkInputCount; inputIndex++)
        {
            var parameters = new Dictionary<string, double>
            {
                ["mean"] = 100 + inputIndex * 7,
                ["stdDev"] = 12 + inputIndex
            };

            config.Inputs.Add(new SimulationInput
            {
                Id = $"BenchmarkInput_{inputIndex + 1}",
                Label = $"Benchmark Input {inputIndex + 1}",
                Distribution = DistributionFactory.Create("Normal", parameters, seed: 42 + inputIndex),
                BaseValue = parameters["mean"]
            });

            profile.Inputs.Add(new SavedInput
            {
                SheetName = "Benchmark Synthetic",
                CellAddress = $"B{inputIndex + 2}",
                Label = $"Benchmark Input {inputIndex + 1}",
                DistributionName = "Normal",
                Parameters = new Dictionary<string, double>(parameters)
            });
        }

        config.Outputs.Add(new SimulationOutput
        {
            Id = "BenchmarkOutput_1",
            Label = "Benchmark Output"
        });

        profile.Outputs.Add(new SavedOutput
        {
            SheetName = "Benchmark Synthetic",
            CellAddress = "N2",
            Label = "Benchmark Output"
        });

        for (var iteration = 0; iteration < ExportBenchmarkIterationCount; iteration++)
        {
            double weightedSum = 0;
            for (var inputIndex = 0; inputIndex < ExportBenchmarkInputCount; inputIndex++)
            {
                var sample = config.Inputs[inputIndex].Distribution.Sample();
                inputMatrix[iteration, inputIndex] = sample;
                weightedSum += sample * (1.0 + inputIndex * 0.04);
            }

            var cyclicalNoise = ((iteration % 17) - 8) * 1.5;
            outputMatrix[iteration, 0] = weightedSum / ExportBenchmarkInputCount + cyclicalNoise;
        }

        var result = new SimulationResult(
            config,
            inputMatrix,
            outputMatrix,
            elapsedTime: TimeSpan.FromMilliseconds(250));
        var stats = new SummaryStatistics(result.GetOutputValues(config.Outputs[0].Id));
        var sensitivity = SensitivityAnalysis.Analyze(result, outputIndex: 0);

        return new SyntheticBenchmarkRun(result, profile, stats, sensitivity);
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

    private sealed record ExportBenchmarkResult(
        int InputCount,
        int IterationCount,
        TimeSpan SummaryExportTime,
        TimeSpan RawExportTime);

    private sealed record SyntheticBenchmarkRun(
        SimulationResult Result,
        SimulationProfile Profile,
        SummaryStatistics Stats,
        IReadOnlyList<SensitivityResult> Sensitivity);
}
