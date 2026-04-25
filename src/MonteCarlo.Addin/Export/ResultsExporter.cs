using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.Services;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using ExcelAxisType = Microsoft.Office.Interop.Excel.XlAxisType;
using ExcelChartType = Microsoft.Office.Interop.Excel.XlChartType;
using ExcelSeriesCollection = Microsoft.Office.Interop.Excel.SeriesCollection;

namespace MonteCarlo.Addin.Export;

/// <summary>
/// Exports simulation results to formatted Excel sheets with statistics tables,
/// sensitivity analysis, input assumptions, and embedded chart images.
/// </summary>
public class ResultsExporter
{
    private const string SheetPrefix = "MC Results — ";
    private const string RawDataSheetPrefix = "MC Raw Data — ";
    private const int ChartColumn = 5; // Column E
    private const int ReportPrintColumn = 11; // Column K
    private const int ChartBlockHeightRows = 54;
    private const int ChartBlockHeightRowsWithoutSensitivity = 36;

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
        byte[]? tornadoImage = null,
        bool createNewSheet = true,
        IReadOnlyList<double>? percentiles = null,
        double? targetValue = null,
        UserSettings? effectiveSettings = null,
        bool usesWorkbookOverrides = false)
    {
        var section = BuildSection(result, profile, outputIndex, stats, sensitivity);
        ExportSummaryReport(
            result,
            profile,
            [section],
            createNewSheet,
            percentiles,
            targetValue,
            effectiveSettings,
            usesWorkbookOverrides,
            reportScope: "Selected output",
            preferredSheetName: $"{SheetPrefix}{section.Output.Label}",
            histogramImage,
            tornadoImage);
    }

    /// <summary>
    /// Export a complete summary sheet containing all outputs in one report.
    /// </summary>
    public void ExportSummaryAllOutputs(
        SimulationResult result,
        SimulationProfile profile,
        int selectedOutputIndex,
        bool createNewSheet = true,
        IReadOnlyList<double>? percentiles = null,
        double? targetValue = null,
        UserSettings? effectiveSettings = null,
        bool usesWorkbookOverrides = false,
        IReadOnlyDictionary<string, IReadOnlyList<SensitivityResult>>? sensitivityByOutputId = null)
    {
        var orderedIndices = BuildOutputOrder(result, selectedOutputIndex);
        var sections = new List<SummaryExportSection>(orderedIndices.Count);

        foreach (var outputIndex in orderedIndices)
        {
            var output = result.Config.Outputs[outputIndex];
            IReadOnlyList<SensitivityResult>? sensitivity = null;
            if (sensitivityByOutputId != null)
                sensitivityByOutputId.TryGetValue(output.Id, out sensitivity);
            sections.Add(BuildSection(
                result,
                profile,
                outputIndex,
                new SummaryStatistics(result.GetOutputValues(output.Id)),
                sensitivity));
        }

        ExportSummaryReport(
            result,
            profile,
            sections,
            createNewSheet,
            percentiles,
            targetValue,
            effectiveSettings,
            usesWorkbookOverrides,
            reportScope: $"All outputs ({sections.Count})",
            preferredSheetName: $"{SheetPrefix}All Outputs");
    }

    /// <summary>
    /// Export raw simulation data to a data sheet.
    /// </summary>
    public void ExportRawData(SimulationResult result, int outputIndex, bool createNewSheet = true)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        using var excelState = ExcelStateScope.Capture(app, "Export raw data", restoreSelection: true);
        var output = result.Config.Outputs[outputIndex];
        string sheetName = GetExportSheetName(workbook, $"{RawDataSheetPrefix}{output.Label}", createNewSheet);
        var sheet = EnsureSheet(workbook, sheetName, clearIfExists: !createNewSheet);

        excelState.Apply(screenUpdating: false, statusBar: "MonteCarlo.XL: exporting raw data...");

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

    #region Private Helpers

    private void ExportSummaryReport(
        SimulationResult result,
        SimulationProfile profile,
        IReadOnlyList<SummaryExportSection> sections,
        bool createNewSheet,
        IReadOnlyList<double>? percentiles,
        double? targetValue,
        UserSettings? effectiveSettings,
        bool usesWorkbookOverrides,
        string reportScope,
        string preferredSheetName,
        byte[]? histogramImage = null,
        byte[]? tornadoImage = null)
    {
        if (sections.Count == 0)
            return;

        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null)
            return;

        using var excelState = ExcelStateScope.Capture(app, "Export summary", restoreSelection: true);
        var sheetName = GetExportSheetName(workbook, preferredSheetName, createNewSheet);
        var sheet = EnsureSheet(workbook, sheetName, clearIfExists: !createNewSheet);

        excelState.Apply(screenUpdating: false, statusBar: "MonteCarlo.XL: exporting summary...");

        var primarySection = sections[0];
        var outputSummary = SummarizeOutputs(sections);
        var primaryOutputCell = primarySection.SavedOutput == null
            ? primarySection.Output.Id
            : $"{primarySection.SavedOutput.SheetName}!{primarySection.SavedOutput.CellAddress}";
        var manualPageBreakRows = new List<int>();

        var row = 1;
        row = WriteTitleSection(
            sheet,
            row,
            sections.Count == 1
                ? $"Output: {primarySection.Output.Label}"
                : $"Outputs: {sections.Count} in one worksheet");
        row++;

        row = WriteReportMetadata(
            sheet,
            row,
            workbook,
            result,
            profile,
            reportScope,
            outputSummary,
            primaryOutputCell,
            effectiveSettings,
            usesWorkbookOverrides);
        row++;

        for (var sectionIndex = 0; sectionIndex < sections.Count; sectionIndex++)
        {
            var section = sections[sectionIndex];
            var sectionStartRow = row;
            if (sectionIndex > 0)
                manualPageBreakRows.Add(sectionStartRow);
            row = WriteOutputSectionHeader(sheet, row, section, sectionIndex + 1, sections.Count);
            row++;

            row = WriteSummaryStats(sheet, row, section.Stats);
            row++;

            row = WritePercentiles(sheet, row, section.Stats, percentiles);
            row++;

            row = WriteTargetAnalysis(sheet, row, section.Stats, targetValue);
            row++;

            if (section.Sensitivity != null && section.Sensitivity.Count > 0)
            {
                row = WriteSensitivity(sheet, row, section.Sensitivity);
                row++;
            }

            row = WriteScenarioAnalysis(sheet, row, result, section.OutputIndex, targetValue);
            row++;

            if (sections.Count == 1 && histogramImage != null)
                EmbedImage(sheet, histogramImage, sectionStartRow, ChartColumn, 500, 280);
            else
                AddHistogramChart(sheet, section.Stats, section.Output.Label, sectionStartRow, ChartColumn, 500, 280);

            AddCdfChart(sheet, section.Stats, section.Output.Label, sectionStartRow + 18, ChartColumn, 500, 280);

            if (sections.Count == 1 && tornadoImage != null)
            {
                EmbedImage(sheet, tornadoImage, sectionStartRow + 36, ChartColumn, 500, 300);
            }
            else if (section.Sensitivity != null && section.Sensitivity.Count > 0)
            {
                AddSensitivityChart(sheet, section.Sensitivity, section.Stats.Median, sectionStartRow + 36, ChartColumn, 500, 300);
            }

            var chartBottomRow = sectionStartRow + (
                section.Sensitivity != null && section.Sensitivity.Count > 0 || tornadoImage != null
                    ? ChartBlockHeightRows
                    : ChartBlockHeightRowsWithoutSensitivity);
            row = Math.Max(row, chartBottomRow) + 2;
        }

        var assumptionsStartRow = row;
        if (sections.Count > 1)
            manualPageBreakRows.Add(assumptionsStartRow);

        row = WriteInputAssumptions(sheet, row, profile);
        row++;

        row = WriteCorrelationAssumptions(sheet, row, profile);
        row++;

        sheet.Columns["A:D"].AutoFit();
        sheet.Tab.Color = 0xF6823B; // #3B82F6 in BGR
        ReportLayoutFormatter.ApplyPdfFriendlyLayout(
            sheet,
            "MonteCarlo.XL Simulation Report",
            row - 1,
            ReportPrintColumn,
            manualPageBreakRows);
    }

    private static SummaryExportSection BuildSection(
        SimulationResult result,
        SimulationProfile profile,
        int outputIndex,
        SummaryStatistics stats,
        IReadOnlyList<SensitivityResult>? sensitivity)
    {
        var output = result.Config.Outputs[outputIndex];
        return new SummaryExportSection(
            outputIndex,
            output,
            GetSavedOutput(profile, output, outputIndex),
            stats,
            sensitivity);
    }

    private static List<int> BuildOutputOrder(SimulationResult result, int selectedOutputIndex)
    {
        var indices = Enumerable.Range(0, result.Config.Outputs.Count).ToList();
        if (selectedOutputIndex <= 0 || selectedOutputIndex >= indices.Count)
            return indices;

        indices.Remove(selectedOutputIndex);
        indices.Insert(0, selectedOutputIndex);
        return indices;
    }

    private static int WriteTitleSection(Worksheet sheet, int row, string subtitle)
    {
        var titleCell = sheet.Cells[row, 1];
        titleCell.Value2 = "MonteCarlo.XL Simulation Report";
        titleCell.Font.Size = 16;
        titleCell.Font.Bold = true;
        row++;

        sheet.Cells[row, 1].Value2 = subtitle;
        sheet.Cells[row, 1].Font.Size = 12;
        row++;

        return row;
    }

    private static int WriteReportMetadata(
        Worksheet sheet,
        int row,
        Workbook workbook,
        SimulationResult result,
        SimulationProfile profile,
        string reportScope,
        string outputSummary,
        string primaryOutputCell,
        UserSettings? effectiveSettings,
        bool usesWorkbookOverrides)
    {
        WriteSectionHeader(sheet, row, "REPORT METADATA");
        row++;

        WriteTableHeader(sheet, row, "Field", "Value");
        row++;

        row = WriteMetadataRow(sheet, row, "Generated at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        row = WriteMetadataRow(sheet, row, "Workbook", SafeWorkbookName(workbook));
        row = WriteMetadataRow(sheet, row, "Workbook path", SafeWorkbookFullName(workbook));
        row = WriteMetadataRow(sheet, row, "Profile", profile.Name);
        row = WriteMetadataRow(sheet, row, "Report scope", reportScope);
        row = WriteMetadataRow(sheet, row, "Outputs included", outputSummary);
        row = WriteMetadataRow(sheet, row, "Primary output cell", primaryOutputCell);
        row = WriteMetadataRow(sheet, row, "Iterations", result.IterationCount.ToString("N0"));
        row = WriteMetadataRow(sheet, row, "Elapsed time", FormatElapsed(result.ElapsedTime));
        row = WriteMetadataRow(sheet, row, "Inputs", result.Config.Inputs.Count.ToString("N0"));
        row = WriteMetadataRow(sheet, row, "Outputs", result.Config.Outputs.Count.ToString("N0"));
        row = WriteMetadataRow(sheet, row, "Sampling method", result.Config.Sampling.ToString());
        row = WriteMetadataRow(sheet, row, "Random seed", result.Config.RandomSeed?.ToString() ?? "Random each run");
        row = WriteMetadataRow(sheet, row, "Convergence auto-stop", result.Config.AutoStopOnConvergence
            ? $"On, minimum {result.Config.ConvergenceMinIterations:N0} iterations"
            : "Off");
        if (effectiveSettings != null)
        {
            row = WriteMetadataRow(sheet, row, "Settings scope", usesWorkbookOverrides ? "Workbook override" : "Windows defaults");
            row = WriteMetadataRow(sheet, row, "Seed preference", DescribeSeedPreference(effectiveSettings));
            row = WriteMetadataRow(sheet, row, "Pause on Model Check warnings", effectiveSettings.PauseOnPreflightWarnings ? "On" : "Off");
            row = WriteMetadataRow(sheet, row, "Export default", effectiveSettings.CreateNewWorksheetForExports
                ? "New worksheet per export"
                : "Reuse export worksheet");
            row = WriteMetadataRow(sheet, row, "Excel calculation", DescribeExcelCalculation(effectiveSettings.ExcelCalculationBehavior));
            row = WriteMetadataRow(sheet, row, "Suspend screen updating", effectiveSettings.SuspendScreenUpdating ? "Yes" : "No");
            row = WriteMetadataRow(sheet, row, "Suspend Excel events", effectiveSettings.SuspendEvents ? "Yes" : "No");
            row = WriteMetadataRow(sheet, row, "Default percentiles", effectiveSettings.DefaultPercentiles);
        }
        row = WriteMetadataRow(sheet, row, "Correlation", DescribeCorrelation(profile));

        return row;
    }

    private static int WriteOutputSectionHeader(
        Worksheet sheet,
        int row,
        SummaryExportSection section,
        int sectionNumber,
        int totalSections)
    {
        var sectionTitle = totalSections == 1
            ? $"OUTPUT — {section.Output.Label}"
            : $"OUTPUT {sectionNumber} OF {totalSections} — {section.Output.Label}";
        WriteSectionHeader(sheet, row, sectionTitle);
        row++;

        sheet.Cells[row, 1].Value2 = "Cell";
        sheet.Cells[row, 2].Value2 = section.SavedOutput == null
            ? section.Output.Id
            : $"{section.SavedOutput.SheetName}!{section.SavedOutput.CellAddress}";

        return row + 1;
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

    private static string DescribeSeedPreference(UserSettings settings) =>
        settings.SeedMode switch
        {
            SeedMode.Fixed => $"Fixed {settings.FixedRandomSeed}",
            SeedMode.Prompt => $"Prompt at run time (default {settings.FixedRandomSeed})",
            _ => "Random each run"
        };

    private static string DescribeExcelCalculation(ExcelCalculationBehavior behavior) =>
        behavior switch
        {
            ExcelCalculationBehavior.Automatic => "Automatic",
            ExcelCalculationBehavior.Manual => "Manual",
            _ => "Keep current"
        };

    private static int WritePercentiles(
        Worksheet sheet,
        int row,
        SummaryStatistics stats,
        IReadOnlyList<double>? configuredPercentiles)
    {
        WriteSectionHeader(sheet, row, "PERCENTILES");
        row++;

        WriteTableHeader(sheet, row, "Percentile", "Value");
        row++;

        var percentiles = configuredPercentiles is { Count: > 0 }
            ? configuredPercentiles
            : new[] { 0.01, 0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99 };

        foreach (var percentile in percentiles)
        {
            var label = $"P{percentile * 100:G4}";
            var value = stats.Percentile(percentile);
            sheet.Cells[row, 1].Value2 = label;
            sheet.Cells[row, 2].Value2 = value;
            sheet.Cells[row, 2].NumberFormat = GetNumberFormat(value);

            if (row % 2 == 0)
                ApplyAltRowShading(sheet, row, 2);
            row++;
        }

        return row;
    }

    private static int WriteTargetAnalysis(Worksheet sheet, int row, SummaryStatistics stats, double? targetValue)
    {
        WriteSectionHeader(sheet, row, "TARGET ANALYSIS");
        row++;

        if (targetValue == null)
        {
            sheet.Cells[row, 1].Value2 = "No target value was entered in the Results view before export.";
            return row + 1;
        }

        WriteTableHeader(sheet, row, "Metric", "Value");
        row++;

        var target = targetValue.Value;
        row = WriteMetricRow(sheet, row, "Target value", target, GetNumberFormat(target));
        row = WriteMetricRow(sheet, row, "P(output <= target)", stats.ProbabilityBelow(target), "0.0%");
        row = WriteMetricRow(sheet, row, "P(output > target)", stats.ProbabilityAbove(target), "0.0%");
        row = WriteMetricRow(sheet, row, "Target percentile rank", stats.ProbabilityBelow(target), "0.0%");

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

    private static int WriteScenarioAnalysis(
        Worksheet sheet,
        int row,
        SimulationResult result,
        int outputIndex,
        double? targetValue)
    {
        WriteSectionHeader(sheet, row, "SCENARIO ANALYSIS");
        row++;

        sheet.Cells[row, 1].Value2 = targetValue == null
            ? "Tail cases compare input means inside the scenario against all simulation runs. Enter a target before export to include target-hit cases."
            : "Tail and target-hit cases compare input means inside the scenario against all simulation runs.";
        row++;

        WriteTableHeader(sheet, row, "Scenario", "Input", "Scenario Mean", "Delta vs All");
        row++;

        var scenarios = new List<ScenarioAnalysisResult>
        {
            ScenarioAnalysis.Analyze(result, outputIndex, ScenarioFilterMode.WorstPercent, 0.10),
            ScenarioAnalysis.Analyze(result, outputIndex, ScenarioFilterMode.BestPercent, 0.10)
        };

        if (targetValue is double target)
        {
            scenarios.Add(ScenarioAnalysis.Analyze(result, outputIndex, ScenarioFilterMode.AtOrBelowTarget, target));
            scenarios.Add(ScenarioAnalysis.Analyze(result, outputIndex, ScenarioFilterMode.AboveTarget, target));
        }

        foreach (var scenario in scenarios)
        {
            var topInputs = scenario.InputSummaries.Take(5).ToList();
            if (topInputs.Count == 0)
            {
                sheet.Cells[row, 1].Value2 = $"{scenario.Description} ({scenario.MatchedFraction:0.0%})";
                sheet.Cells[row, 2].Value2 = "No matching runs";
                sheet.Cells[row, 3].Value2 = string.Empty;
                sheet.Cells[row, 4].Value2 = string.Empty;
                row++;
                continue;
            }

            foreach (var input in topInputs)
            {
                sheet.Cells[row, 1].Value2 =
                    $"{scenario.Description} ({scenario.MatchedFraction:0.0%})";
                sheet.Cells[row, 2].Value2 = input.InputLabel;
                sheet.Cells[row, 3].Value2 = input.ScenarioMean;
                sheet.Cells[row, 3].NumberFormat = GetNumberFormat(input.ScenarioMean);
                sheet.Cells[row, 4].Value2 = input.Delta;
                sheet.Cells[row, 4].NumberFormat = GetNumberFormat(input.Delta);

                if (row % 2 == 0)
                    ApplyAltRowShading(sheet, row, 4);

                row++;
            }
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

    private static int WriteCorrelationAssumptions(Worksheet sheet, int row, SimulationProfile profile)
    {
        WriteSectionHeader(sheet, row, "CORRELATION ASSUMPTIONS");
        row++;

        var matrix = profile.CorrelationMatrix;
        if (matrix == null)
        {
            sheet.Cells[row, 1].Value2 = "Inputs are sampled independently.";
            return row + 1;
        }

        var inputCount = profile.Inputs.Count;
        if (matrix.GetLength(0) != inputCount || matrix.GetLength(1) != inputCount)
        {
            sheet.Cells[row, 1].Value2 = "Correlation matrix size does not match the saved input list.";
            return row + 1;
        }

        WriteTableHeader(sheet, row, "Input A", "Input B", "Spearman Rank Corr");
        row++;

        var wroteAny = false;
        for (var i = 0; i < inputCount; i++)
        {
            for (var j = i + 1; j < inputCount; j++)
            {
                var correlation = matrix[i, j];
                if (Math.Abs(correlation) < 1e-12)
                    continue;

                sheet.Cells[row, 1].Value2 = profile.Inputs[i].Label;
                sheet.Cells[row, 2].Value2 = profile.Inputs[j].Label;
                sheet.Cells[row, 3].Value2 = correlation;
                sheet.Cells[row, 3].NumberFormat = "0.000";

                if (row % 2 == 0)
                    ApplyAltRowShading(sheet, row, 3);

                wroteAny = true;
                row++;
            }
        }

        if (!wroteAny)
        {
            sheet.Cells[row, 1].Value2 = "No non-zero correlations are configured.";
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

    private static void AddHistogramChart(Worksheet sheet, SummaryStatistics stats, string outputLabel,
        int row, int col, int widthPx, int heightPx)
    {
        var histogram = stats.ToHistogram();
        const int dataRow = 1;
        const int dataCol = 27; // AA

        sheet.Cells[dataRow, dataCol].Value2 = "Bin";
        sheet.Cells[dataRow, dataCol + 1].Value2 = "Relative Frequency";

        var data = new object[histogram.BinCenters.Length, 2];
        for (var i = 0; i < histogram.BinCenters.Length; i++)
        {
            data[i, 0] = histogram.BinCenters[i];
            data[i, 1] = histogram.RelativeFrequencies[i];
        }

        var firstDataRow = dataRow + 1;
        var lastDataRow = firstDataRow + histogram.BinCenters.Length - 1;
        var dataRange = sheet.Range[sheet.Cells[firstDataRow, dataCol], sheet.Cells[lastDataRow, dataCol + 1]];
        dataRange.Value2 = data;

        var chartObject = AddChartObject(sheet, row, col, widthPx, heightPx);
        var chart = chartObject.Chart;
        chart.ChartType = ExcelChartType.xlColumnClustered;
        chart.HasTitle = true;
        chart.ChartTitle.Text = $"{outputLabel} Distribution";
        chart.HasLegend = false;

        var series = (Series)((ExcelSeriesCollection)chart.SeriesCollection()).NewSeries();
        series.Name = "Relative Frequency";
        series.XValues = sheet.Range[sheet.Cells[firstDataRow, dataCol], sheet.Cells[lastDataRow, dataCol]];
        series.Values = sheet.Range[sheet.Cells[firstDataRow, dataCol + 1], sheet.Cells[lastDataRow, dataCol + 1]];

        var categoryAxis = (Axis)chart.Axes(ExcelAxisType.xlCategory);
        categoryAxis.TickLabels.NumberFormat = GetNumberFormat(stats.Mean);
        categoryAxis.HasTitle = true;
        categoryAxis.AxisTitle.Text = "Output value";

        var valueAxis = (Axis)chart.Axes(ExcelAxisType.xlValue);
        valueAxis.TickLabels.NumberFormat = "0.0%";
        valueAxis.HasTitle = true;
        valueAxis.AxisTitle.Text = "Relative frequency";
    }

    private static void AddCdfChart(Worksheet sheet, SummaryStatistics stats, string outputLabel,
        int row, int col, int widthPx, int heightPx)
    {
        const int maxPoints = 101;
        var sortedValues = stats.SortedValues;
        var pointCount = Math.Min(maxPoints, sortedValues.Length);
        if (pointCount == 0)
            return;

        const int dataRow = 1;
        const int dataCol = 34; // AH

        sheet.Cells[dataRow, dataCol].Value2 = "Output Value";
        sheet.Cells[dataRow, dataCol + 1].Value2 = "Cumulative Probability";

        var data = new object[pointCount, 2];
        for (var i = 0; i < pointCount; i++)
        {
            var index = pointCount == 1
                ? 0
                : (int)Math.Round(i * (sortedValues.Length - 1) / (double)(pointCount - 1));

            data[i, 0] = sortedValues[index];
            data[i, 1] = (index + 1) / (double)sortedValues.Length;
        }

        var firstDataRow = dataRow + 1;
        var lastDataRow = firstDataRow + pointCount - 1;
        var dataRange = sheet.Range[sheet.Cells[firstDataRow, dataCol], sheet.Cells[lastDataRow, dataCol + 1]];
        dataRange.Value2 = data;

        var chartObject = AddChartObject(sheet, row, col, widthPx, heightPx);
        var chart = chartObject.Chart;
        chart.ChartType = ExcelChartType.xlXYScatterLinesNoMarkers;
        chart.HasTitle = true;
        chart.ChartTitle.Text = $"{outputLabel} Cumulative Probability";
        chart.HasLegend = false;

        var series = (Series)((ExcelSeriesCollection)chart.SeriesCollection()).NewSeries();
        series.Name = "Cumulative Probability";
        series.XValues = sheet.Range[sheet.Cells[firstDataRow, dataCol], sheet.Cells[lastDataRow, dataCol]];
        series.Values = sheet.Range[sheet.Cells[firstDataRow, dataCol + 1], sheet.Cells[lastDataRow, dataCol + 1]];

        var categoryAxis = (Axis)chart.Axes(ExcelAxisType.xlCategory);
        categoryAxis.TickLabels.NumberFormat = GetNumberFormat(stats.Mean);
        categoryAxis.HasTitle = true;
        categoryAxis.AxisTitle.Text = "Output value";

        var valueAxis = (Axis)chart.Axes(ExcelAxisType.xlValue);
        valueAxis.MinimumScale = 0;
        valueAxis.MaximumScale = 1;
        valueAxis.TickLabels.NumberFormat = "0%";
        valueAxis.HasTitle = true;
        valueAxis.AxisTitle.Text = "Cumulative probability";
    }

    private static void AddSensitivityChart(Worksheet sheet, IReadOnlyList<SensitivityResult> sensitivity,
        double baseValue, int row, int col, int widthPx, int heightPx)
    {
        var items = sensitivity
            .OrderByDescending(r => r.Swing)
            .Take(10)
            .Reverse()
            .ToList();

        if (items.Count == 0)
            return;

        const int dataRow = 1;
        const int dataCol = 30; // AD

        sheet.Cells[dataRow, dataCol].Value2 = "Input";
        sheet.Cells[dataRow, dataCol + 1].Value2 = "P10 Impact";
        sheet.Cells[dataRow, dataCol + 2].Value2 = "P90 Impact";

        var data = new object[items.Count, 3];
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            data[i, 0] = item.InputLabel;
            data[i, 1] = item.OutputAtInputP10 - baseValue;
            data[i, 2] = item.OutputAtInputP90 - baseValue;
        }

        var firstDataRow = dataRow + 1;
        var lastDataRow = firstDataRow + items.Count - 1;
        var dataRange = sheet.Range[sheet.Cells[firstDataRow, dataCol], sheet.Cells[lastDataRow, dataCol + 2]];
        dataRange.Value2 = data;

        var labels = sheet.Range[sheet.Cells[firstDataRow, dataCol], sheet.Cells[lastDataRow, dataCol]];
        var chartObject = AddChartObject(sheet, row, col, widthPx, heightPx);
        var chart = chartObject.Chart;
        chart.ChartType = ExcelChartType.xlBarClustered;
        chart.HasTitle = true;
        chart.ChartTitle.Text = "Sensitivity";
        chart.HasLegend = true;

        var seriesCollection = (ExcelSeriesCollection)chart.SeriesCollection();

        var p10Series = (Series)seriesCollection.NewSeries();
        p10Series.Name = "Input P10";
        p10Series.XValues = labels;
        p10Series.Values = sheet.Range[sheet.Cells[firstDataRow, dataCol + 1], sheet.Cells[lastDataRow, dataCol + 1]];

        var p90Series = (Series)seriesCollection.NewSeries();
        p90Series.Name = "Input P90";
        p90Series.XValues = labels;
        p90Series.Values = sheet.Range[sheet.Cells[firstDataRow, dataCol + 2], sheet.Cells[lastDataRow, dataCol + 2]];

        var categoryAxis = (Axis)chart.Axes(ExcelAxisType.xlCategory);
        categoryAxis.ReversePlotOrder = true;

        var valueAxis = (Axis)chart.Axes(ExcelAxisType.xlValue);
        valueAxis.HasTitle = true;
        valueAxis.AxisTitle.Text = "Change from median output";
    }

    private static ChartObject AddChartObject(Worksheet sheet, int row, int col, int widthPx, int heightPx)
    {
        var cell = (Range)sheet.Cells[row, col];

        // Convert pixels to points (72 pts per inch, 96 px per inch)
        double widthPts = widthPx * 72.0 / 96.0;
        double heightPts = heightPx * 72.0 / 96.0;

        var chartObjects = (ChartObjects)sheet.ChartObjects(Type.Missing);
        return chartObjects.Add(cell.Left, cell.Top, widthPts, heightPts);
    }

    private static Worksheet EnsureSheet(Workbook workbook, string name, bool clearIfExists = true)
    {
        // Check if sheet exists
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name == name)
            {
                if (clearIfExists)
                {
                    ClearShapes(ws);
                    ws.Cells.Clear();
                }

                return ws;
            }
        }

        // Create new sheet at the end
        var lastSheet = workbook.Worksheets[workbook.Worksheets.Count];
        var newSheet = (Worksheet)workbook.Worksheets.Add(After: lastSheet);
        newSheet.Name = name;
        return newSheet;
    }

    private static void ClearShapes(Worksheet sheet)
    {
        for (var i = sheet.Shapes.Count; i >= 1; i--)
            sheet.Shapes.Item(i).Delete();
    }

    private static string GetExportSheetName(Workbook workbook, string baseName, bool createNewSheet)
    {
        var cleanBaseName = TruncateSheetName(SanitizeSheetName(baseName));
        return createNewSheet ? GetUniqueSheetName(workbook, cleanBaseName) : cleanBaseName;
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
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private static string SanitizeSheetName(string name)
    {
        var invalidChars = new[] { '\\', '/', '?', '*', '[', ']', ':' };
        foreach (var invalidChar in invalidChars)
            name = name.Replace(invalidChar, '-');

        return string.IsNullOrWhiteSpace(name) ? "MonteCarlo Export" : name.Trim();
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
            "beta" => $"Beta(α={p.GetValueOrDefault("alpha")}, β={p.GetValueOrDefault("beta")})",
            "exponential" => $"Exponential(λ={p.GetValueOrDefault("rate")})",
            "weibull" => $"Weibull(k={p.GetValueOrDefault("shape")}, λ={p.GetValueOrDefault("scale")})",
            "poisson" => $"Poisson(λ={p.GetValueOrDefault("lambda")})",
            "gamma" => $"Gamma(shape={p.GetValueOrDefault("shape")}, rate={p.GetValueOrDefault("rate")})",
            "logistic" => $"Logistic(μ={p.GetValueOrDefault("mu")}, s={p.GetValueOrDefault("s")})",
            "gev" => $"GEV(μ={p.GetValueOrDefault("mu")}, σ={p.GetValueOrDefault("sigma")}, ξ={p.GetValueOrDefault("xi")})",
            "binomial" => $"Binomial(n={p.GetValueOrDefault("n")}, p={p.GetValueOrDefault("p")})",
            "geometric" => $"Geometric(p={p.GetValueOrDefault("p")})",
            "discrete" => "Discrete(...)",
            _ => input.DistributionName
        };
    }

    private static int WriteMetricRow(Worksheet sheet, int row, string label, double value, string numberFormat)
    {
        sheet.Cells[row, 1].Value2 = label;
        sheet.Cells[row, 2].Value2 = value;
        sheet.Cells[row, 2].NumberFormat = numberFormat;

        if (row % 2 == 0)
            ApplyAltRowShading(sheet, row, 2);

        return row + 1;
    }

    private static int WriteMetadataRow(Worksheet sheet, int row, string label, string value)
    {
        sheet.Cells[row, 1].Value2 = label;
        sheet.Cells[row, 2].Value2 = value;

        if (row % 2 == 0)
            ApplyAltRowShading(sheet, row, 2);

        return row + 1;
    }

    private static SavedOutput? GetSavedOutput(
        SimulationProfile profile,
        SimulationOutput output,
        int outputIndex)
    {
        if (outputIndex >= 0 && outputIndex < profile.Outputs.Count)
            return profile.Outputs[outputIndex];

        return profile.Outputs.FirstOrDefault(saved =>
            string.Equals(saved.Label, output.Label, StringComparison.OrdinalIgnoreCase));
    }

    private static string SafeWorkbookName(Workbook workbook)
    {
        try
        {
            return workbook.Name;
        }
        catch
        {
            return "Unknown workbook";
        }
    }

    private static string SafeWorkbookFullName(Workbook workbook)
    {
        try
        {
            return workbook.FullName;
        }
        catch
        {
            return "Unsaved workbook";
        }
    }

    private static string FormatElapsed(TimeSpan elapsed)
    {
        return elapsed.TotalMinutes < 1
            ? $"{elapsed.TotalSeconds:0.00} seconds"
            : elapsed.ToString(@"hh\:mm\:ss");
    }

    private static string DescribeCorrelation(SimulationProfile profile)
    {
        var matrix = profile.CorrelationMatrix;
        if (matrix == null)
            return "Independent inputs";

        var nonZeroPairs = 0;
        var maxAbsCorrelation = 0.0;
        var inputCount = Math.Min(matrix.GetLength(0), matrix.GetLength(1));
        for (var i = 0; i < inputCount; i++)
        {
            for (var j = i + 1; j < inputCount; j++)
            {
                var absCorrelation = Math.Abs(matrix[i, j]);
                if (absCorrelation < 1e-12)
                    continue;

                nonZeroPairs++;
                maxAbsCorrelation = Math.Max(maxAbsCorrelation, absCorrelation);
            }
        }

        return nonZeroPairs == 0
            ? "Matrix configured, no non-zero pairs"
            : $"{nonZeroPairs:N0} correlated pair(s), max |ρ| {maxAbsCorrelation:0.000}";
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

    private static string SummarizeOutputs(IReadOnlyList<SummaryExportSection> sections)
    {
        if (sections.Count == 1)
            return sections[0].Output.Label;

        const int maxNames = 3;
        var labels = sections.Select(section => section.Output.Label).ToList();
        if (labels.Count <= maxNames)
            return string.Join(", ", labels);

        return $"{string.Join(", ", labels.Take(maxNames))}, +{labels.Count - maxNames} more";
    }

    #endregion

    private sealed record SummaryExportSection(
        int OutputIndex,
        SimulationOutput Output,
        SavedOutput? SavedOutput,
        SummaryStatistics Stats,
        IReadOnlyList<SensitivityResult>? Sensitivity);
}
