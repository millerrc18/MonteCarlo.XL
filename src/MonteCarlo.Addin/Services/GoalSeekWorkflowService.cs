using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
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
        int? randomSeed)
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
            Metric = GoalSeekMetric.ProbabilityAboveTarget,
            OutputTarget = runOptions.OutputTarget,
            DesiredMetricValue = runOptions.DesiredProbability,
            MaxIterations = runOptions.MaxSolverIterations,
            MetricTolerance = runOptions.Tolerance,
            HigherDecisionIncreasesMetric = runOptions.HigherDecisionIncreasesMetric
        };

        using var excelState = ExcelStateScope.Capture(app, "Goal seek", restoreSelection: true);
        try
        {
            excelState.Apply(
                screenUpdating: false,
                enableEvents: false,
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
                        correlationMatrix).ConfigureAwait(false);

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
        row = WriteLabelValue(sheet, row, "Target metric", $"P(output > {options.OutputTarget:G6})");
        row = WriteLabelValue(sheet, row, "Desired probability", options.DesiredProbability.ToString("0.0%"));
        row = WriteLabelValue(sheet, row, "Best decision value", result.BestDecisionValue.ToString("G10"));
        row = WriteLabelValue(sheet, row, "Best probability", result.BestMetricValue.ToString("0.0%"));
        row = WriteLabelValue(sheet, row, "Error", result.Error.ToString("0.0000"));
        row = WriteLabelValue(sheet, row, "Status", result.Status.ToString());
        row = WriteLabelValue(sheet, row, "Iterations per trial", options.IterationsPerTrial.ToString("N0"));
        row++;

        sheet.Cells[row, 1].Value2 = "Solver History";
        sheet.Cells[row, 1].Font.Bold = true;
        row++;

        sheet.Cells[row, 1].Value2 = "Step";
        sheet.Cells[row, 2].Value2 = "Decision Value";
        sheet.Cells[row, 3].Value2 = "Probability";
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
            sheet.Cells[row, 3].NumberFormat = "0.0%";
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

    private static string FormatCompletionMessage(GoalSeekResult result, GoalSeekRunOptions options) =>
        $"Goal Seek finished with status: {result.Status}\r\n\r\n" +
        $"Best decision value: {result.BestDecisionValue:G10}\r\n" +
        $"Best P(output > {options.OutputTarget:G6}): {result.BestMetricValue:P1}\r\n" +
        $"Desired probability: {options.DesiredProbability:P1}\r\n\r\n" +
        "A Goal Seek report sheet was added to the workbook.";

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
    double OutputTarget,
    double DesiredProbability,
    int IterationsPerTrial,
    int MaxSolverIterations,
    double Tolerance,
    bool HigherDecisionIncreasesMetric);

internal sealed class GoalSeekOptionsDialog : WinForms.Form
{
    private readonly WinForms.ComboBox _outputComboBox = new();
    private readonly WinForms.TextBox _lowerBoundTextBox = new() { Text = "0" };
    private readonly WinForms.TextBox _upperBoundTextBox = new() { Text = "100" };
    private readonly WinForms.TextBox _outputTargetTextBox = new();
    private readonly WinForms.TextBox _desiredProbabilityTextBox = new() { Text = "0.8" };
    private readonly WinForms.TextBox _iterationsTextBox = new() { Text = "1000" };
    private readonly WinForms.TextBox _maxIterationsTextBox = new() { Text = "12" };
    private readonly WinForms.TextBox _toleranceTextBox = new() { Text = "0.005" };
    private readonly WinForms.CheckBox _increasingCheckBox = new()
    {
        Text = "Higher decision value increases the probability metric",
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
            RowCount = 11
        };
        panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 42));
        panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 58));

        AddReadOnlyRow(panel, "Decision cell", decisionReference.FullReference);
        AddComboRow(panel, "Output", _outputComboBox);
        AddTextRow(panel, "Lower decision bound", _lowerBoundTextBox);
        AddTextRow(panel, "Upper decision bound", _upperBoundTextBox);
        AddTextRow(panel, "Output target", _outputTargetTextBox);
        AddTextRow(panel, "Desired probability", _desiredProbabilityTextBox);
        AddTextRow(panel, "Iterations per trial", _iterationsTextBox);
        AddTextRow(panel, "Max solver steps", _maxIterationsTextBox);
        AddTextRow(panel, "Probability tolerance", _toleranceTextBox);

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

        if (!TryParseDouble(_lowerBoundTextBox.Text, "lower decision bound", out var lower, out error)
            || !TryParseDouble(_upperBoundTextBox.Text, "upper decision bound", out var upper, out error)
            || !TryParseDouble(_outputTargetTextBox.Text, "output target", out var outputTarget, out error)
            || !TryParseDouble(_desiredProbabilityTextBox.Text, "desired probability", out var desiredProbability, out error)
            || !TryParseDouble(_toleranceTextBox.Text, "probability tolerance", out var tolerance, out error))
        {
            return false;
        }

        if (lower >= upper)
        {
            error = "Lower decision bound must be less than upper decision bound.";
            return false;
        }

        if (desiredProbability < 0 || desiredProbability > 1)
        {
            error = "Desired probability must be between 0 and 1.";
            return false;
        }

        if (tolerance <= 0)
        {
            error = "Probability tolerance must be greater than zero.";
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

        options = new GoalSeekRunOptions(
            output.Output.Cell.FullReference,
            output.Output.Label,
            lower,
            upper,
            outputTarget,
            desiredProbability,
            iterations,
            maxIterations,
            tolerance,
            _increasingCheckBox.Checked);
        return true;
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
        panel.Controls.Add(new WinForms.Label { Text = label, AutoSize = true, Anchor = WinForms.AnchorStyles.Left });
        panel.Controls.Add(new WinForms.Label { Text = value, AutoSize = true, Anchor = WinForms.AnchorStyles.Left });
    }

    private static void AddComboRow(WinForms.TableLayoutPanel panel, string label, WinForms.ComboBox comboBox)
    {
        panel.Controls.Add(new WinForms.Label { Text = label, AutoSize = true, Anchor = WinForms.AnchorStyles.Left });
        comboBox.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
        comboBox.Dock = WinForms.DockStyle.Fill;
        panel.Controls.Add(comboBox);
    }

    private static void AddTextRow(WinForms.TableLayoutPanel panel, string label, WinForms.TextBox textBox)
    {
        panel.Controls.Add(new WinForms.Label { Text = label, AutoSize = true, Anchor = WinForms.AnchorStyles.Left });
        textBox.Dock = WinForms.DockStyle.Fill;
        panel.Controls.Add(textBox);
    }

    private sealed class OutputOption
    {
        public OutputOption(TaggedOutput output)
        {
            Output = output;
        }

        public TaggedOutput Output { get; }

        public override string ToString() => $"{Output.Label} ({Output.Cell.FullReference})";
    }
}
