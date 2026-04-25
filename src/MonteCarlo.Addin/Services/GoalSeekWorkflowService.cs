using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.Services;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using ExcelCalculation = Microsoft.Office.Interop.Excel.XlCalculation;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using WinForms = System.Windows.Forms;

namespace MonteCarlo.Addin.Services;

/// <summary>
/// Runs a lightweight Excel-hosted goal seek workflow around the engine solver.
/// </summary>
internal sealed class GoalSeekWorkflowService
{
    private const string GoalSeekSheetPrefix = "MC Goal Seek";

    public async Task RunAsync(
        SimulationOrchestrator orchestrator,
        IReadOnlyList<TaggedInput> inputs,
        IReadOnlyList<TaggedOutput> outputs,
        double[,]? correlationMatrix,
        SamplingMethod samplingMethod,
        int? randomSeed,
        ExcelExecutionOptions? executionOptions = null)
    {
        if (outputs.Count == 0)
        {
            ShowMessage("Add at least one output before running Goal Seek.", WinForms.MessageBoxIcon.Warning);
            return;
        }

        var app = (ExcelApplication)ExcelDnaUtil.Application;
        if (app.ActiveCell is not ExcelRange decisionCell)
        {
            ShowMessage("Select the decision cell first, then run Goal Seek.", WinForms.MessageBoxIcon.Warning);
            return;
        }

        var decisionSheet = (ExcelWorksheet)decisionCell.Worksheet;
        var decisionReference = new CellReference
        {
            SheetName = decisionSheet.Name,
            CellAddress = decisionCell.Address[false, false]
        };

        if (inputs.Any(input => input.Cell.Equals(decisionReference)))
        {
            ShowMessage(
                "The selected decision cell is already configured as a Monte Carlo input.\r\n\r\n" +
                "Choose a deterministic model driver cell that affects the output, not an uncertain input cell.",
                WinForms.MessageBoxIcon.Warning);
            return;
        }

        var originalValue = decisionCell.Value2;
        var originalFormula = decisionCell.Formula;
        var dialog = new GoalSeekOptionsDialog(outputs, decisionReference);
        if (dialog.ShowDialog() != WinForms.DialogResult.OK || dialog.Options == null)
            return;

        var runOptions = dialog.Options;
        var solverOptions = new GoalSeekOptions
        {
            LowerBound = runOptions.LowerBound,
            UpperBound = runOptions.UpperBound,
            Metric = runOptions.Metric,
            OutputTarget = runOptions.OutputTarget ?? 0,
            DesiredMetricValue = runOptions.DesiredMetricValue,
            Percentile = runOptions.Percentile ?? 0.5,
            MaxIterations = runOptions.MaxSolverIterations,
            MetricTolerance = runOptions.Tolerance,
            HigherDecisionIncreasesMetric = runOptions.HigherDecisionIncreasesMetric
        };

        using var excelState = ExcelStateScope.Capture(app, "Goal seek", restoreSelection: true);
        var execution = executionOptions ?? ExcelExecutionOptions.Default;
        try
        {
            excelState.Apply(
                screenUpdating: execution.SuspendScreenUpdating ? false : null,
                enableEvents: execution.SuspendEvents ? false : null,
                calculation: execution.CalculationBehavior switch
                {
                    ExcelCalculationBehavior.Automatic => ExcelCalculation.xlCalculationAutomatic,
                    ExcelCalculationBehavior.Manual => ExcelCalculation.xlCalculationManual,
                    _ => null
                },
                statusBar: "MonteCarlo.XL: running Goal Seek...");

            var result = await GoalSeekUnderUncertainty.SolveAsync(
                solverOptions,
                async (decisionValue, cancellationToken) =>
                {
                    await RunOnExcelThreadAsync(
                        () =>
                        {
                            app.StatusBar = $"MonteCarlo.XL Goal Seek: testing {decisionReference.FullReference} = {decisionValue:G6}...";
                            decisionCell.Value2 = decisionValue;
                            app.Calculate();
                        },
                        cancellationToken).ConfigureAwait(false);

                    await orchestrator.RunSimulationAsync(
                        runOptions.IterationsPerTrial,
                        randomSeed,
                        samplingMethod,
                        autoStopOnConvergence: false,
                        correlationMatrix,
                        execution).ConfigureAwait(false);

                    var simulationResult = orchestrator.LastResult
                        ?? throw new InvalidOperationException("Goal Seek simulation did not produce results.");

                    return simulationResult.GetOutputValues(runOptions.OutputFullReference);
                }).ConfigureAwait(false);

            await RunOnExcelThreadAsync(
                () =>
                {
                    WriteGoalSeekReport(app, decisionReference, runOptions, result);
                    ShowMessage(FormatCompletionMessage(result, runOptions), WinForms.MessageBoxIcon.Information);
                },
                CancellationToken.None).ConfigureAwait(false);
        }
        finally
        {
            await RunOnExcelThreadAsync(
                () =>
                {
                    if (originalFormula is string formula && formula.StartsWith("=", StringComparison.Ordinal))
                        decisionCell.Formula = formula;
                    else
                        decisionCell.Value2 = originalValue;

                    app.Calculate();
                },
                CancellationToken.None).ConfigureAwait(false);
        }
    }

    private static void WriteGoalSeekReport(
        ExcelApplication app,
        CellReference decisionReference,
        GoalSeekRunOptions options,
        GoalSeekResult result)
    {
        var workbook = app.ActiveWorkbook;
        if (workbook == null)
            return;

        var sheetName = GetUniqueSheetName(workbook, GoalSeekSheetPrefix);
        var lastSheet = workbook.Worksheets[workbook.Worksheets.Count];
        var sheet = (ExcelWorksheet)workbook.Worksheets.Add(After: lastSheet);
        sheet.Name = sheetName;

        var row = 1;
        sheet.Cells[row, 1].Value2 = "MonteCarlo.XL Goal Seek";
        sheet.Cells[row, 1].Font.Bold = true;
        sheet.Cells[row, 1].Font.Size = 16;
        row += 2;

        row = WriteLabelValue(sheet, row, "Decision cell", decisionReference.FullReference);
        row = WriteLabelValue(sheet, row, "Output", options.OutputLabel);
        row = WriteLabelValue(sheet, row, "Target metric", DescribeMetric(options));
        row = WriteLabelValue(sheet, row, "Desired value", FormatMetricValue(options, options.DesiredMetricValue));
        row = WriteLabelValue(sheet, row, "Best decision value", result.BestDecisionValue.ToString("G10"));
        row = WriteLabelValue(sheet, row, "Best metric value", FormatMetricValue(options, result.BestMetricValue));
        row = WriteLabelValue(sheet, row, "Error", result.Error.ToString("0.0000"));
        row = WriteLabelValue(sheet, row, "Status", result.Status.ToString());
        row = WriteLabelValue(sheet, row, "Iterations per trial", options.IterationsPerTrial.ToString("N0"));
        row++;

        row = WriteConfidenceGuidance(sheet, row, options, result);
        row++;

        sheet.Cells[row, 1].Value2 = "Solver History";
        sheet.Cells[row, 1].Font.Bold = true;
        row++;

        sheet.Cells[row, 1].Value2 = "Step";
        sheet.Cells[row, 2].Value2 = "Decision Value";
        sheet.Cells[row, 3].Value2 = "Metric Value";
        sheet.Cells[row, 4].Value2 = "Error";
        var header = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 4]];
        header.Font.Bold = true;
        header.Interior.Color = 0xFAF8F8;
        row++;

        foreach (var item in result.History)
        {
            sheet.Cells[row, 1].Value2 = item.Iteration;
            sheet.Cells[row, 2].Value2 = item.DecisionValue;
            sheet.Cells[row, 3].Value2 = item.MetricValue;
            sheet.Cells[row, 3].NumberFormat = GetMetricNumberFormat(options);
            sheet.Cells[row, 4].Value2 = item.Error;
            sheet.Cells[row, 4].NumberFormat = "0.0000";
            row++;
        }

        sheet.Columns["A:D"].AutoFit();
        sheet.Tab.Color = 0x81A1F5;
    }

    private static int WriteLabelValue(ExcelWorksheet sheet, int row, string label, string value)
    {
        sheet.Cells[row, 1].Value2 = label;
        sheet.Cells[row, 1].Font.Bold = true;
        sheet.Cells[row, 2].Value2 = value;
        return row + 1;
    }

    private static int WriteConfidenceGuidance(
        ExcelWorksheet sheet,
        int row,
        GoalSeekRunOptions options,
        GoalSeekResult result)
    {
        var assessment = AssessConfidence(options, result);
        sheet.Cells[row, 1].Value2 = "Confidence Guidance";
        sheet.Cells[row, 1].Font.Bold = true;
        row++;

        row = WriteLabelValue(sheet, row, "Assessment", assessment.Headline);
        row = WriteLabelValue(sheet, row, "Why", assessment.Why);
        row = WriteLabelValue(sheet, row, "Bounds context", assessment.BoundsContext);
        row = WriteLabelValue(sheet, row, "Recommendation", assessment.Recommendation);
        return row;
    }

    private static string FormatCompletionMessage(GoalSeekResult result, GoalSeekRunOptions options) =>
        FormatCompletionMessage(result, options, AssessConfidence(options, result));

    private static string FormatCompletionMessage(
        GoalSeekResult result,
        GoalSeekRunOptions options,
        GoalSeekConfidenceAssessment assessment) =>
        $"Goal Seek finished with status: {result.Status}\r\n\r\n" +
        $"Best decision value: {result.BestDecisionValue:G10}\r\n" +
        $"Best {DescribeMetric(options)}: {FormatMetricValue(options, result.BestMetricValue)}\r\n" +
        $"Desired value: {FormatMetricValue(options, options.DesiredMetricValue)}\r\n" +
        $"Confidence: {assessment.Headline}\r\n\r\n" +
        "A Goal Seek report sheet was added to the workbook.";

    private static string DescribeMetric(GoalSeekRunOptions options) =>
        options.Metric switch
        {
            GoalSeekMetric.Mean => "Mean(output)",
            GoalSeekMetric.ProbabilityAboveTarget => $"P(output > {options.OutputTarget:G6})",
            GoalSeekMetric.ProbabilityAtOrBelowTarget => $"P(output <= {options.OutputTarget:G6})",
            GoalSeekMetric.Percentile => $"P{((options.Percentile ?? 0.5) * 100.0):0.#}(output)",
            _ => options.Metric.ToString()
        };

    private static string FormatMetricValue(GoalSeekRunOptions options, double value) =>
        options.Metric is GoalSeekMetric.ProbabilityAboveTarget or GoalSeekMetric.ProbabilityAtOrBelowTarget
            ? value.ToString("P1")
            : value.ToString("G10");

    private static string GetMetricNumberFormat(GoalSeekRunOptions options) =>
        options.Metric is GoalSeekMetric.ProbabilityAboveTarget or GoalSeekMetric.ProbabilityAtOrBelowTarget
            ? "0.0%"
            : "0.0000";

    private static GoalSeekConfidenceAssessment AssessConfidence(GoalSeekRunOptions options, GoalSeekResult result)
    {
        var absoluteError = Math.Abs(result.Error);
        var decisionSpan = Math.Abs(options.UpperBound - options.LowerBound);
        var nearestBoundDistance = Math.Min(
            Math.Abs(result.BestDecisionValue - options.LowerBound),
            Math.Abs(options.UpperBound - result.BestDecisionValue));
        var nearestBoundFraction = decisionSpan <= 0 ? 0 : nearestBoundDistance / decisionSpan;
        var testedBothSides = result.History.Any(item => item.Error <= 0) && result.History.Any(item => item.Error >= 0);

        var nearbyPoints = result.History
            .OrderBy(item => Math.Abs(item.DecisionValue - result.BestDecisionValue))
            .Take(3)
            .ToList();
        var localMetricSpread = nearbyPoints.Count >= 2
            ? nearbyPoints.Max(item => item.MetricValue) - nearbyPoints.Min(item => item.MetricValue)
            : 0;

        var errorText = FormatMetricValue(options, absoluteError);
        var toleranceText = FormatMetricValue(options, options.Tolerance);
        var boundsText = nearestBoundFraction switch
        {
            <= 0.05 => $"Best value sits very close to a search bound ({nearestBoundFraction:P0} of the range from the nearest edge).",
            <= 0.15 => $"Best value sits somewhat close to a search bound ({nearestBoundFraction:P0} of the range from the nearest edge).",
            _ => $"Best value stays comfortably inside the tested range ({nearestBoundFraction:P0} of the range from the nearest edge)."
        };

        if (result.Status == GoalSeekStatus.TargetNotBracketed)
        {
            return new GoalSeekConfidenceAssessment(
                "Low - target not bracketed",
                "The requested metric was outside the values reached at the lower and upper decision bounds, so the solver could only return the closer edge.",
                boundsText,
                "Widen or reposition the decision bounds, then rerun Goal Seek.");
        }

        if (result.Status == GoalSeekStatus.Converged
            && absoluteError <= options.Tolerance * 0.5
            && nearestBoundFraction >= 0.10
            && testedBothSides)
        {
            return new GoalSeekConfidenceAssessment(
                "High - clean bracket and low residual",
                $"Residual error {errorText} is comfortably inside the tolerance {toleranceText}, and the solver tested both sides of the target near the chosen value.",
                boundsText,
                $"Use the report value as a strong starting point. Nearby tested metric spread was {FormatMetricValue(options, localMetricSpread)}.");
        }

        if (result.Status == GoalSeekStatus.Converged || absoluteError <= options.Tolerance * 1.5)
        {
            return new GoalSeekConfidenceAssessment(
                "Medium - usable, but worth a spot check",
                $"Residual error {errorText} is near the requested tolerance {toleranceText}. The solver found a workable answer, but it is closer to the edge or the local response is still moving.",
                boundsText,
                "Sanity-check the workbook around the reported decision value with one or two nearby manual trials before using it in a final recommendation.");
        }

        return new GoalSeekConfidenceAssessment(
            "Low - approximate only",
            $"Residual error {errorText} stayed above the requested tolerance {toleranceText}, so the solver stopped with the best value it had rather than a clean hit.",
            boundsText,
            "Increase solver steps or iterations per trial, and consider widening the decision bounds before rerunning.");
    }

    private static Task RunOnExcelThreadAsync(Action action, CancellationToken cancellationToken)
    {
        var tcs = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                action();
                tcs.TrySetResult(null);
            }
            catch (OperationCanceledException ex)
            {
                tcs.TrySetCanceled(ex.CancellationToken);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        });

        return tcs.Task.WaitAsync(cancellationToken);
    }

    private static string GetUniqueSheetName(ExcelWorkbook workbook, string baseName)
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

    private static bool SheetExists(ExcelWorkbook workbook, string name)
    {
        foreach (ExcelWorksheet sheet in workbook.Worksheets)
        {
            if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }

    private static void ShowMessage(string message, WinForms.MessageBoxIcon icon) =>
        WinForms.MessageBox.Show(
            message,
            "MonteCarlo.XL Goal Seek",
            WinForms.MessageBoxButtons.OK,
            icon);
}

internal sealed record GoalSeekRunOptions(
    string OutputFullReference,
    string OutputLabel,
    double LowerBound,
    double UpperBound,
    GoalSeekMetric Metric,
    double DesiredMetricValue,
    double? OutputTarget,
    double? Percentile,
    int IterationsPerTrial,
    int MaxSolverIterations,
    double Tolerance,
    bool HigherDecisionIncreasesMetric);

internal sealed record GoalSeekConfidenceAssessment(
    string Headline,
    string Why,
    string BoundsContext,
    string Recommendation);

internal sealed class GoalSeekOptionsDialog : WinForms.Form
{
    private readonly WinForms.ComboBox _outputComboBox = new();
    private readonly WinForms.ComboBox _metricComboBox = new();
    private readonly WinForms.TextBox _lowerBoundTextBox = new() { Text = "0" };
    private readonly WinForms.TextBox _upperBoundTextBox = new() { Text = "100" };
    private readonly WinForms.TextBox _outputTargetTextBox = new();
    private readonly WinForms.TextBox _percentileTextBox = new() { Text = "90" };
    private readonly WinForms.TextBox _desiredMetricTextBox = new() { Text = "0.8" };
    private readonly WinForms.TextBox _iterationsTextBox = new() { Text = "1000" };
    private readonly WinForms.TextBox _maxIterationsTextBox = new() { Text = "12" };
    private readonly WinForms.TextBox _toleranceTextBox = new() { Text = "0.005" };
    private readonly WinForms.Label _outputTargetLabel = CreateRowLabel("Output target");
    private readonly WinForms.Label _percentileLabel = CreateRowLabel("Percentile (0-100)");
    private readonly WinForms.Label _desiredMetricLabel = CreateRowLabel("Desired probability");
    private readonly WinForms.Label _toleranceLabel = CreateRowLabel("Probability tolerance");
    private readonly WinForms.CheckBox _increasingCheckBox = new()
    {
        Text = "Higher decision value increases the selected metric",
        Checked = true,
        AutoSize = true
    };

    public GoalSeekRunOptions? Options { get; private set; }

    public GoalSeekOptionsDialog(IReadOnlyList<TaggedOutput> outputs, CellReference decisionReference)
    {
        Text = "Goal Seek Under Uncertainty";
        Width = 460;
        Height = 430;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var panel = new WinForms.TableLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            Padding = new WinForms.Padding(14),
            ColumnCount = 2,
            RowCount = 13
        };
        panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 42));
        panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 58));

        AddReadOnlyRow(panel, "Decision cell", decisionReference.FullReference);
        AddControlRow(panel, "Output", _outputComboBox);
        AddControlRow(panel, "Metric", _metricComboBox);
        AddControlRow(panel, "Lower decision bound", _lowerBoundTextBox);
        AddControlRow(panel, "Upper decision bound", _upperBoundTextBox);
        AddControlRow(panel, _outputTargetLabel, _outputTargetTextBox);
        AddControlRow(panel, _percentileLabel, _percentileTextBox);
        AddControlRow(panel, _desiredMetricLabel, _desiredMetricTextBox);
        AddControlRow(panel, "Iterations per trial", _iterationsTextBox);
        AddControlRow(panel, "Max solver steps", _maxIterationsTextBox);
        AddControlRow(panel, _toleranceLabel, _toleranceTextBox);

        panel.Controls.Add(new WinForms.Label());
        panel.Controls.Add(_increasingCheckBox);

        var buttons = new WinForms.FlowLayoutPanel
        {
            FlowDirection = WinForms.FlowDirection.RightToLeft,
            Dock = WinForms.DockStyle.Fill
        };
        var okButton = new WinForms.Button { Text = "Run Goal Seek", DialogResult = WinForms.DialogResult.OK, AutoSize = true };
        var cancelButton = new WinForms.Button { Text = "Cancel", DialogResult = WinForms.DialogResult.Cancel, AutoSize = true };
        buttons.Controls.Add(okButton);
        buttons.Controls.Add(cancelButton);
        panel.Controls.Add(new WinForms.Label());
        panel.Controls.Add(buttons);

        Controls.Add(panel);
        AcceptButton = okButton;
        CancelButton = cancelButton;

        foreach (var output in outputs)
            _outputComboBox.Items.Add(new OutputOption(output));

        if (_outputComboBox.Items.Count > 0)
            _outputComboBox.SelectedIndex = 0;

        foreach (var metric in MetricOption.All)
            _metricComboBox.Items.Add(metric);

        _metricComboBox.SelectedItem = MetricOption.Default;
        _metricComboBox.SelectedIndexChanged += (_, _) => UpdateMetricControls();
        UpdateMetricControls();

        okButton.Click += (_, _) =>
        {
            if (!TryReadOptions(out var options, out var error))
            {
                WinForms.MessageBox.Show(error, "Goal Seek Under Uncertainty", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                DialogResult = WinForms.DialogResult.None;
                return;
            }

            Options = options;
        };
    }

    private bool TryReadOptions(out GoalSeekRunOptions? options, out string error)
    {
        options = null;
        error = string.Empty;

        if (_outputComboBox.SelectedItem is not OutputOption output)
        {
            error = "Choose an output.";
            return false;
        }

        var metric = _metricComboBox.SelectedItem is MetricOption metricOption
            ? metricOption.Metric
            : GoalSeekMetric.ProbabilityAboveTarget;

        if (!TryParseDouble(_lowerBoundTextBox.Text, "lower decision bound", out var lower, out error)
            || !TryParseDouble(_upperBoundTextBox.Text, "upper decision bound", out var upper, out error)
            || !TryParseDouble(_toleranceTextBox.Text, "metric tolerance", out var tolerance, out error))
        {
            return false;
        }

        if (lower >= upper)
        {
            error = "Lower decision bound must be less than upper decision bound.";
            return false;
        }

        if (tolerance <= 0)
        {
            error = "Metric tolerance must be greater than zero.";
            return false;
        }

        if (!int.TryParse(_iterationsTextBox.Text, out var iterations) || iterations <= 0)
        {
            error = "Iterations per trial must be a positive whole number.";
            return false;
        }

        if (!int.TryParse(_maxIterationsTextBox.Text, out var maxIterations) || maxIterations <= 0)
        {
            error = "Max solver steps must be a positive whole number.";
            return false;
        }

        double? outputTarget = null;
        if (metric is GoalSeekMetric.ProbabilityAboveTarget or GoalSeekMetric.ProbabilityAtOrBelowTarget)
        {
            if (!TryParseDouble(_outputTargetTextBox.Text, "output target", out var parsedOutputTarget, out error))
                return false;

            outputTarget = parsedOutputTarget;
        }

        double? percentile = null;
        if (metric == GoalSeekMetric.Percentile)
        {
            if (!TryParseDouble(_percentileTextBox.Text, "percentile", out var percentilePercent, out error))
                return false;

            if (percentilePercent < 0 || percentilePercent > 100)
            {
                error = "Percentile must be between 0 and 100.";
                return false;
            }

            percentile = percentilePercent / 100.0;
        }

        if (!TryParseDouble(
                _desiredMetricTextBox.Text,
                metric is GoalSeekMetric.ProbabilityAboveTarget or GoalSeekMetric.ProbabilityAtOrBelowTarget
                    ? "desired probability"
                    : "desired metric value",
                out var desiredMetricValue,
                out error))
        {
            return false;
        }

        if (metric is GoalSeekMetric.ProbabilityAboveTarget or GoalSeekMetric.ProbabilityAtOrBelowTarget
            && (desiredMetricValue < 0 || desiredMetricValue > 1))
        {
            error = "Desired probability must be between 0 and 1.";
            return false;
        }

        options = new GoalSeekRunOptions(
            output.Output.Cell.FullReference,
            output.Output.Label,
            lower,
            upper,
            metric,
            desiredMetricValue,
            outputTarget,
            percentile,
            iterations,
            maxIterations,
            tolerance,
            _increasingCheckBox.Checked);
        return true;
    }

    private void UpdateMetricControls()
    {
        var metric = _metricComboBox.SelectedItem is MetricOption metricOption
            ? metricOption.Metric
            : GoalSeekMetric.ProbabilityAboveTarget;

        var usesTarget = metric is GoalSeekMetric.ProbabilityAboveTarget or GoalSeekMetric.ProbabilityAtOrBelowTarget;
        var usesPercentile = metric == GoalSeekMetric.Percentile;

        _outputTargetLabel.Enabled = usesTarget;
        _outputTargetTextBox.Enabled = usesTarget;
        _percentileLabel.Enabled = usesPercentile;
        _percentileTextBox.Enabled = usesPercentile;

        switch (metric)
        {
            case GoalSeekMetric.Mean:
                _desiredMetricLabel.Text = "Desired mean";
                _desiredMetricTextBox.Text = "100";
                _toleranceLabel.Text = "Metric tolerance";
                _toleranceTextBox.Text = "0.5";
                break;

            case GoalSeekMetric.Percentile:
                _desiredMetricLabel.Text = "Desired percentile value";
                _desiredMetricTextBox.Text = "100";
                _toleranceLabel.Text = "Metric tolerance";
                _toleranceTextBox.Text = "0.5";
                if (string.IsNullOrWhiteSpace(_percentileTextBox.Text))
                    _percentileTextBox.Text = "90";
                break;

            default:
                _desiredMetricLabel.Text = "Desired probability";
                _desiredMetricTextBox.Text = "0.8";
                _toleranceLabel.Text = "Probability tolerance";
                _toleranceTextBox.Text = "0.005";
                break;
        }
    }

    private static bool TryParseDouble(string text, string label, out double value, out string error)
    {
        if (double.TryParse(
                text,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out value)
            && double.IsFinite(value))
        {
            error = string.Empty;
            return true;
        }

        error = $"Enter a valid number for {label}.";
        return false;
    }

    private static void AddReadOnlyRow(WinForms.TableLayoutPanel panel, string label, string value)
    {
        panel.Controls.Add(CreateRowLabel(label));
        panel.Controls.Add(new WinForms.Label { Text = value, AutoSize = true, Anchor = WinForms.AnchorStyles.Left });
    }

    private static void AddControlRow(WinForms.TableLayoutPanel panel, string label, WinForms.Control control) =>
        AddControlRow(panel, CreateRowLabel(label), control);

    private static void AddControlRow(WinForms.TableLayoutPanel panel, WinForms.Label label, WinForms.Control control)
    {
        panel.Controls.Add(label);
        if (control is WinForms.ComboBox comboBox)
            comboBox.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
        control.Dock = WinForms.DockStyle.Fill;
        panel.Controls.Add(control);
    }

    private static WinForms.Label CreateRowLabel(string label) => new()
    {
        Text = label,
        AutoSize = true,
        Anchor = WinForms.AnchorStyles.Left
    };

    private sealed class OutputOption
    {
        public OutputOption(TaggedOutput output)
        {
            Output = output;
        }

        public TaggedOutput Output { get; }

        public override string ToString() => $"{Output.Label} ({Output.Cell.FullReference})";
    }

    private sealed class MetricOption
    {
        public static IReadOnlyList<MetricOption> All { get; } =
        [
            new MetricOption(GoalSeekMetric.ProbabilityAboveTarget, "Probability above target"),
            new MetricOption(GoalSeekMetric.ProbabilityAtOrBelowTarget, "Probability at or below target"),
            new MetricOption(GoalSeekMetric.Mean, "Mean"),
            new MetricOption(GoalSeekMetric.Percentile, "Percentile")
        ];

        public static MetricOption Default => All[0];

        private MetricOption(GoalSeekMetric metric, string label)
        {
            Metric = metric;
            Label = label;
        }

        public GoalSeekMetric Metric { get; }

        public string Label { get; }

        public override string ToString() => Label;
    }
}
