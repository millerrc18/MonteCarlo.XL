using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.Services;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using Microsoft.Office.Interop.Excel;
using ExcelAxisType = Microsoft.Office.Interop.Excel.XlAxisType;
using ExcelChartType = Microsoft.Office.Interop.Excel.XlChartType;
using ExcelSeriesCollection = Microsoft.Office.Interop.Excel.SeriesCollection;

namespace MonteCarlo.Addin.Export;

internal sealed class StressAnalysisExporter
{
    private const string SheetPrefix = "MC Stress Analysis";
    private const int ChartColumn = 7; // Column G
    private const int PrintColumn = 12; // Column L
    private const int HistogramDataColumn = 20; // Column T
    private const int CdfDataColumn = 24; // Column X

    public void ExportComparison(
        SimulationResult baseline,
        SimulationResult stressed,
        StressRunOptions options,
        int comparisonSeed)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(stressed);
        ArgumentNullException.ThrowIfNull(options);

        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null)
            return;

        using var excelState = ExcelStateScope.Capture(app, "Stress analysis export", restoreSelection: true);
        var sheetName = GetUniqueSheetName(workbook, SheetPrefix);
        var lastSheet = workbook.Worksheets[workbook.Worksheets.Count];
        var sheet = (Worksheet)workbook.Worksheets.Add(After: lastSheet);
        sheet.Name = sheetName;

        excelState.Apply(screenUpdating: false, statusBar: "MonteCarlo.XL: writing stress analysis report...");

        var outputComparisons = BuildOutputComparisons(baseline, stressed);
        var primaryComparison = outputComparisons.FirstOrDefault(
                                    comparison => string.Equals(
                                        comparison.OutputId,
                                        options.PrimaryOutputId,
                                        StringComparison.OrdinalIgnoreCase))
                                ?? outputComparisons[0];

        var row = 1;
        row = WriteTitle(sheet, row, options.PrimaryOutputLabel);
        row++;

        row = WriteMetadata(sheet, row, baseline, stressed, options, comparisonSeed);
        row++;

        row = WriteStressRules(sheet, row, options.Plan);
        row++;

        row = WriteOutputImpactRanking(sheet, row, outputComparisons);
        row++;

        var primarySectionRow = row;
        row = WritePrimaryOutputDetail(sheet, row, primaryComparison);
        row++;

        AddHistogramComparisonChart(
            sheet,
            primaryComparison,
            primarySectionRow,
            ChartColumn,
            500,
            280);
        AddCdfComparisonChart(
            sheet,
            primaryComparison,
            primarySectionRow + 18,
            ChartColumn,
            500,
            280);

        row = Math.Max(row, primarySectionRow + 36) + 2;

        sheet.Columns["A:F"].AutoFit();
        sheet.Tab.Color = 0x51A0EB;

        ReportLayoutFormatter.ApplyPdfFriendlyLayout(
            sheet,
            "MonteCarlo.XL Stress Analysis",
            row - 1,
            PrintColumn,
            [primarySectionRow]);
    }

    private static int WriteTitle(Worksheet sheet, int row, string primaryOutputLabel)
    {
        var titleCell = sheet.Cells[row, 1];
        titleCell.Value2 = "MonteCarlo.XL Stress Analysis";
        titleCell.Font.Size = 16;
        titleCell.Font.Bold = true;
        row++;

        sheet.Cells[row, 1].Value2 = $"Primary output: {primaryOutputLabel}";
        sheet.Cells[row, 1].Font.Size = 12;
        return row + 1;
    }

    private static int WriteMetadata(
        Worksheet sheet,
        int row,
        SimulationResult baseline,
        SimulationResult stressed,
        StressRunOptions options,
        int comparisonSeed)
    {
        WriteSectionHeader(sheet, row, "RUN METADATA");
        row++;
        WriteTableHeader(sheet, row, "Field", "Value");
        row++;

        row = WriteLabelValue(sheet, row, "Generated at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        row = WriteLabelValue(sheet, row, "Iterations per run", options.IterationsPerRun.ToString("N0"));
        row = WriteLabelValue(sheet, row, "Comparison seed", comparisonSeed.ToString());
        row = WriteLabelValue(sheet, row, "Sampling method", baseline.Config.Sampling.ToString());
        row = WriteLabelValue(sheet, row, "Baseline elapsed", FormatElapsed(baseline.ElapsedTime));
        row = WriteLabelValue(sheet, row, "Stress elapsed", FormatElapsed(stressed.ElapsedTime));
        row = WriteLabelValue(sheet, row, "Inputs stressed", options.Plan.Count.ToString("N0"));
        row = WriteLabelValue(sheet, row, "Outputs compared", baseline.Config.Outputs.Count.ToString("N0"));
        return row;
    }

    private static int WriteStressRules(Worksheet sheet, int row, StressInputPlan plan)
    {
        WriteSectionHeader(sheet, row, "STRESS RULES");
        row++;
        WriteTableHeader(sheet, row, "Input", "Mode", "Value", "Description");
        row++;

        if (plan.Count == 0)
        {
            sheet.Cells[row, 1].Value2 = "No stress rules were defined.";
            return row + 1;
        }

        foreach (var rule in plan.Rules)
        {
            sheet.Cells[row, 1].Value2 = rule.InputLabel;
            sheet.Cells[row, 2].Value2 = DescribeMode(rule.Mode);
            sheet.Cells[row, 3].Value2 = rule.Value;
            sheet.Cells[row, 3].NumberFormat = GetNumberFormat(rule.Value);
            sheet.Cells[row, 4].Value2 = rule.Describe();
            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 4);
            row++;
        }

        return row;
    }

    private static int WriteOutputImpactRanking(
        Worksheet sheet,
        int row,
        IReadOnlyList<OutputComparison> comparisons)
    {
        WriteSectionHeader(sheet, row, "OUTPUT IMPACT RANKING");
        row++;
        WriteTableHeader(
            sheet,
            row,
            "Rank",
            "Output",
            "Baseline Mean",
            "Stress Mean",
            "Mean Delta",
            "Mean Delta %",
            "Baseline P50",
            "Stress P50",
            "P5-P95 Shift");
        row++;

        var ranked = comparisons
            .OrderByDescending(comparison => Math.Abs(comparison.MeanDelta))
            .ToList();

        for (var index = 0; index < ranked.Count; index++)
        {
            var comparison = ranked[index];
            sheet.Cells[row, 1].Value2 = index + 1;
            sheet.Cells[row, 2].Value2 = comparison.OutputLabel;
            WriteValueCell(sheet, row, 3, comparison.Baseline.Mean);
            WriteValueCell(sheet, row, 4, comparison.Stressed.Mean);
            WriteValueCell(sheet, row, 5, comparison.MeanDelta);
            sheet.Cells[row, 6].Value2 = comparison.MeanDeltaPercent;
            sheet.Cells[row, 6].NumberFormat = "0.0%;-0.0%;0.0%";
            WriteValueCell(sheet, row, 7, comparison.Baseline.Median);
            WriteValueCell(sheet, row, 8, comparison.Stressed.Median);
            WriteValueCell(sheet, row, 9, comparison.P5P95SpreadDelta);

            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 9);
            row++;
        }

        return row;
    }

    private static int WritePrimaryOutputDetail(Worksheet sheet, int row, OutputComparison comparison)
    {
        WriteSectionHeader(sheet, row, $"PRIMARY OUTPUT DETAIL - {comparison.OutputLabel}");
        row++;
        WriteTableHeader(sheet, row, "Statistic", "Baseline", "Stressed", "Delta");
        row++;

        var items = new (string Label, double Baseline, double Stressed)[]
        {
            ("Mean", comparison.Baseline.Mean, comparison.Stressed.Mean),
            ("Median", comparison.Baseline.Median, comparison.Stressed.Median),
            ("Std Dev", comparison.Baseline.StdDev, comparison.Stressed.StdDev),
            ("P5", comparison.Baseline.P5, comparison.Stressed.P5),
            ("P95", comparison.Baseline.P95, comparison.Stressed.P95),
            ("P5-P95 Spread", comparison.Baseline.P95 - comparison.Baseline.P5, comparison.Stressed.P95 - comparison.Stressed.P5),
            ("Minimum", comparison.Baseline.Min, comparison.Stressed.Min),
            ("Maximum", comparison.Baseline.Max, comparison.Stressed.Max)
        };

        foreach (var item in items)
        {
            sheet.Cells[row, 1].Value2 = item.Label;
            WriteValueCell(sheet, row, 2, item.Baseline);
            WriteValueCell(sheet, row, 3, item.Stressed);
            WriteValueCell(sheet, row, 4, item.Stressed - item.Baseline);
            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 4);
            row++;
        }

        return row;
    }

    private static IReadOnlyList<OutputComparison> BuildOutputComparisons(
        SimulationResult baseline,
        SimulationResult stressed)
    {
        var comparisons = new List<OutputComparison>(baseline.Config.Outputs.Count);
        foreach (var output in baseline.Config.Outputs)
        {
            var baselineStats = new SummaryStatistics(baseline.GetOutputValues(output.Id));
            var stressedStats = new SummaryStatistics(stressed.GetOutputValues(output.Id));
            comparisons.Add(new OutputComparison(output.Id, output.Label, baselineStats, stressedStats));
        }

        return comparisons;
    }

    private static void AddHistogramComparisonChart(
        Worksheet sheet,
        OutputComparison comparison,
        int row,
        int col,
        int widthPx,
        int heightPx)
    {
        var (centers, baselineFreq, stressedFreq) = BuildHistogramSeries(
            comparison.Baseline.SortedValues,
            comparison.Stressed.SortedValues);

        const int firstDataRow = 2;
        sheet.Cells[1, HistogramDataColumn].Value2 = "Bin";
        sheet.Cells[1, HistogramDataColumn + 1].Value2 = "Baseline";
        sheet.Cells[1, HistogramDataColumn + 2].Value2 = "Stressed";

        var data = new object[centers.Length, 3];
        for (var i = 0; i < centers.Length; i++)
        {
            data[i, 0] = centers[i];
            data[i, 1] = baselineFreq[i];
            data[i, 2] = stressedFreq[i];
        }

        var lastDataRow = firstDataRow + centers.Length - 1;
        var dataRange = sheet.Range[
            sheet.Cells[firstDataRow, HistogramDataColumn],
            sheet.Cells[lastDataRow, HistogramDataColumn + 2]];
        dataRange.Value2 = data;

        var chartObject = AddChartObject(sheet, row, col, widthPx, heightPx);
        var chart = chartObject.Chart;
        chart.ChartType = ExcelChartType.xlColumnClustered;
        chart.HasTitle = true;
        chart.ChartTitle.Text = $"{comparison.OutputLabel} Histogram Comparison";
        chart.HasLegend = true;

        var seriesCollection = (ExcelSeriesCollection)chart.SeriesCollection();
        var baselineSeries = (Series)seriesCollection.NewSeries();
        baselineSeries.Name = "Baseline";
        baselineSeries.XValues = sheet.Range[sheet.Cells[firstDataRow, HistogramDataColumn], sheet.Cells[lastDataRow, HistogramDataColumn]];
        baselineSeries.Values = sheet.Range[sheet.Cells[firstDataRow, HistogramDataColumn + 1], sheet.Cells[lastDataRow, HistogramDataColumn + 1]];

        var stressedSeries = (Series)seriesCollection.NewSeries();
        stressedSeries.Name = "Stressed";
        stressedSeries.XValues = sheet.Range[sheet.Cells[firstDataRow, HistogramDataColumn], sheet.Cells[lastDataRow, HistogramDataColumn]];
        stressedSeries.Values = sheet.Range[sheet.Cells[firstDataRow, HistogramDataColumn + 2], sheet.Cells[lastDataRow, HistogramDataColumn + 2]];

        var categoryAxis = (Axis)chart.Axes(ExcelAxisType.xlCategory);
        categoryAxis.TickLabels.NumberFormat = GetNumberFormat(comparison.Baseline.Mean);
        categoryAxis.HasTitle = true;
        categoryAxis.AxisTitle.Text = "Output value";

        var valueAxis = (Axis)chart.Axes(ExcelAxisType.xlValue);
        valueAxis.TickLabels.NumberFormat = "0.0%";
        valueAxis.HasTitle = true;
        valueAxis.AxisTitle.Text = "Relative frequency";
    }

    private static void AddCdfComparisonChart(
        Worksheet sheet,
        OutputComparison comparison,
        int row,
        int col,
        int widthPx,
        int heightPx)
    {
        var (xValues, baselineProbabilities, stressedProbabilities) = BuildCdfSeries(
            comparison.Baseline,
            comparison.Stressed);

        const int firstDataRow = 2;
        sheet.Cells[1, CdfDataColumn].Value2 = "Output";
        sheet.Cells[1, CdfDataColumn + 1].Value2 = "Baseline";
        sheet.Cells[1, CdfDataColumn + 2].Value2 = "Stressed";

        var data = new object[xValues.Length, 3];
        for (var i = 0; i < xValues.Length; i++)
        {
            data[i, 0] = xValues[i];
            data[i, 1] = baselineProbabilities[i];
            data[i, 2] = stressedProbabilities[i];
        }

        var lastDataRow = firstDataRow + xValues.Length - 1;
        var dataRange = sheet.Range[
            sheet.Cells[firstDataRow, CdfDataColumn],
            sheet.Cells[lastDataRow, CdfDataColumn + 2]];
        dataRange.Value2 = data;

        var chartObject = AddChartObject(sheet, row, col, widthPx, heightPx);
        var chart = chartObject.Chart;
        chart.ChartType = ExcelChartType.xlXYScatterLinesNoMarkers;
        chart.HasTitle = true;
        chart.ChartTitle.Text = $"{comparison.OutputLabel} CDF Comparison";
        chart.HasLegend = true;

        var seriesCollection = (ExcelSeriesCollection)chart.SeriesCollection();
        var baselineSeries = (Series)seriesCollection.NewSeries();
        baselineSeries.Name = "Baseline";
        baselineSeries.XValues = sheet.Range[sheet.Cells[firstDataRow, CdfDataColumn], sheet.Cells[lastDataRow, CdfDataColumn]];
        baselineSeries.Values = sheet.Range[sheet.Cells[firstDataRow, CdfDataColumn + 1], sheet.Cells[lastDataRow, CdfDataColumn + 1]];

        var stressedSeries = (Series)seriesCollection.NewSeries();
        stressedSeries.Name = "Stressed";
        stressedSeries.XValues = sheet.Range[sheet.Cells[firstDataRow, CdfDataColumn], sheet.Cells[lastDataRow, CdfDataColumn]];
        stressedSeries.Values = sheet.Range[sheet.Cells[firstDataRow, CdfDataColumn + 2], sheet.Cells[lastDataRow, CdfDataColumn + 2]];

        var categoryAxis = (Axis)chart.Axes(ExcelAxisType.xlCategory);
        categoryAxis.TickLabels.NumberFormat = GetNumberFormat(comparison.Baseline.Mean);
        categoryAxis.HasTitle = true;
        categoryAxis.AxisTitle.Text = "Output value";

        var valueAxis = (Axis)chart.Axes(ExcelAxisType.xlValue);
        valueAxis.MinimumScale = 0;
        valueAxis.MaximumScale = 1;
        valueAxis.TickLabels.NumberFormat = "0%";
        valueAxis.HasTitle = true;
        valueAxis.AxisTitle.Text = "Cumulative probability";
    }

    private static (double[] Centers, double[] BaselineFrequencies, double[] StressedFrequencies) BuildHistogramSeries(
        double[] baselineValues,
        double[] stressedValues,
        int binCount = 24)
    {
        var min = Math.Min(baselineValues[0], stressedValues[0]);
        var max = Math.Max(baselineValues[^1], stressedValues[^1]);
        if (Math.Abs(max - min) < double.Epsilon)
        {
            min -= 0.5;
            max += 0.5;
            binCount = 1;
        }

        var width = (max - min) / binCount;
        var centers = new double[binCount];
        var baselineFreq = new double[binCount];
        var stressedFreq = new double[binCount];

        for (var i = 0; i < binCount; i++)
            centers[i] = min + ((i + 0.5) * width);

        AddHistogramCounts(baselineValues, min, width, baselineFreq);
        AddHistogramCounts(stressedValues, min, width, stressedFreq);

        return (centers, baselineFreq, stressedFreq);
    }

    private static void AddHistogramCounts(double[] values, double min, double width, double[] frequencies)
    {
        foreach (var value in values)
        {
            var bin = width <= 0
                ? 0
                : Math.Min(frequencies.Length - 1, Math.Max(0, (int)((value - min) / width)));
            frequencies[bin] += 1.0 / values.Length;
        }
    }

    private static (double[] X, double[] Baseline, double[] Stressed) BuildCdfSeries(
        SummaryStatistics baseline,
        SummaryStatistics stressed,
        int pointCount = 120)
    {
        var min = Math.Min(baseline.Min, stressed.Min);
        var max = Math.Max(baseline.Max, stressed.Max);
        if (Math.Abs(max - min) < double.Epsilon)
        {
            min -= 0.5;
            max += 0.5;
        }

        var x = new double[pointCount];
        var baselineY = new double[pointCount];
        var stressedY = new double[pointCount];
        var step = (max - min) / (pointCount - 1);

        for (var i = 0; i < pointCount; i++)
        {
            x[i] = min + (step * i);
            baselineY[i] = baseline.ProbabilityBelow(x[i]);
            stressedY[i] = stressed.ProbabilityBelow(x[i]);
        }

        return (x, baselineY, stressedY);
    }

    private static ChartObject AddChartObject(Worksheet sheet, int row, int col, int widthPx, int heightPx)
    {
        var cell = (Range)sheet.Cells[row, col];
        var widthPts = widthPx * 72.0 / 96.0;
        var heightPts = heightPx * 72.0 / 96.0;
        var chartObjects = (ChartObjects)sheet.ChartObjects(Type.Missing);
        return chartObjects.Add(cell.Left, cell.Top, widthPts, heightPts);
    }

    private static void WriteSectionHeader(Worksheet sheet, int row, string title)
    {
        var range = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 4]];
        range.Merge();
        range.Value2 = title;
        range.Font.Bold = true;
        range.Font.Size = 11;
        range.Interior.Color = 0xFAF8F8;
    }

    private static void WriteTableHeader(Worksheet sheet, int row, params string[] headers)
    {
        for (var i = 0; i < headers.Length; i++)
        {
            sheet.Cells[row, i + 1].Value2 = headers[i];
        }

        var range = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, headers.Length]];
        range.Font.Bold = true;
        range.Interior.Color = 0xFAF8F8;
    }

    private static int WriteLabelValue(Worksheet sheet, int row, string label, string value)
    {
        sheet.Cells[row, 1].Value2 = label;
        sheet.Cells[row, 1].Font.Bold = true;
        sheet.Cells[row, 2].Value2 = value;
        return row + 1;
    }

    private static void WriteValueCell(Worksheet sheet, int row, int col, double value)
    {
        sheet.Cells[row, col].Value2 = value;
        sheet.Cells[row, col].NumberFormat = GetNumberFormat(value);
    }

    private static void ApplyAltRowShading(Worksheet sheet, int row, int colCount)
    {
        var range = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, colCount]];
        range.Interior.Color = 0xF4F1EF;
    }

    private static string DescribeMode(StressRuleMode mode) =>
        mode switch
        {
            StressRuleMode.FixedValue => "Fixed value",
            StressRuleMode.AddShift => "Add shift",
            StressRuleMode.RangeScale => "Range scale",
            _ => mode.ToString()
        };

    private static string GetNumberFormat(double value)
    {
        var magnitude = Math.Abs(value);
        if (magnitude >= 1000)
            return "#,##0.0";
        if (magnitude >= 1 || magnitude == 0)
            return "0.00";
        if (magnitude >= 0.01)
            return "0.000";
        return "0.0000";
    }

    private static string FormatElapsed(TimeSpan elapsed) =>
        elapsed.TotalMinutes >= 1
            ? elapsed.ToString(@"mm\:ss")
            : elapsed.ToString(@"ss\.ff\s");

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

    private sealed record OutputComparison(
        string OutputId,
        string OutputLabel,
        SummaryStatistics Baseline,
        SummaryStatistics Stressed)
    {
        public double MeanDelta => Stressed.Mean - Baseline.Mean;

        public double MeanDeltaPercent =>
            Math.Abs(Baseline.Mean) > double.Epsilon
                ? MeanDelta / Baseline.Mean
                : 0;

        public double P5P95SpreadDelta =>
            (Stressed.P95 - Stressed.P5) - (Baseline.P95 - Baseline.P5);
    }
}
