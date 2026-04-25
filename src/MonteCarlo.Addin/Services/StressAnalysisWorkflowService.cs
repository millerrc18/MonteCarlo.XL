using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.Export;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.Services;
using WinForms = System.Windows.Forms;

namespace MonteCarlo.Addin.Services;

internal sealed class StressAnalysisWorkflowService
{
    public async Task RunAsync(
        SimulationOrchestrator orchestrator,
        IReadOnlyList<TaggedInput> inputs,
        IReadOnlyList<TaggedOutput> outputs,
        string? defaultOutputId,
        int defaultIterations,
        double[,]? correlationMatrix,
        SamplingMethod samplingMethod,
        int comparisonSeed,
        bool autoStopOnConvergence,
        ExcelExecutionOptions? executionOptions = null)
    {
        if (inputs.Count == 0)
        {
            ShowMessage(
                "Add at least one configured input before running Stress Analysis.\r\n\r\n" +
                "The current workflow stresses saved input assumptions from the setup model manager.",
                WinForms.MessageBoxIcon.Warning);
            return;
        }

        if (outputs.Count == 0)
        {
            ShowMessage("Add at least one output before running Stress Analysis.", WinForms.MessageBoxIcon.Warning);
            return;
        }

        using var dialog = new StressAnalysisOptionsDialog(inputs, outputs, defaultOutputId, defaultIterations);
        if (dialog.ShowDialog() != WinForms.DialogResult.OK || dialog.Options == null)
            return;

        var options = dialog.Options;
        var execution = executionOptions ?? ExcelExecutionOptions.Default;

        await orchestrator.RunSimulationAsync(
            options.IterationsPerRun,
            comparisonSeed,
            samplingMethod,
            autoStopOnConvergence,
            correlationMatrix,
            execution).ConfigureAwait(false);

        var baseline = orchestrator.LastResult
            ?? throw new InvalidOperationException("Stress Analysis baseline run did not produce results.");

        await orchestrator.RunSimulationAsync(
            options.IterationsPerRun,
            comparisonSeed,
            samplingMethod,
            autoStopOnConvergence,
            correlationMatrix,
            execution,
            inputTransform: (input, sample) => options.Plan.Apply(input, sample)).ConfigureAwait(false);

        var stressed = orchestrator.LastResult
            ?? throw new InvalidOperationException("Stress Analysis stressed run did not produce results.");

        await RunOnExcelThreadAsync(
            () =>
            {
                new StressAnalysisExporter().ExportComparison(baseline, stressed, options, comparisonSeed);
                ShowMessage(
                    "Stress analysis complete. MonteCarlo.XL added a comparison sheet with the stressed assumptions, output impact ranking, and primary-output histogram/CDF comparison.",
                    WinForms.MessageBoxIcon.Information);
            },
            CancellationToken.None).ConfigureAwait(false);
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

    private static void ShowMessage(string message, WinForms.MessageBoxIcon icon) =>
        WinForms.MessageBox.Show(
            message,
            "MonteCarlo.XL Stress Analysis",
            WinForms.MessageBoxButtons.OK,
            icon);
}

internal sealed class StressAnalysisOptionsDialog : WinForms.Form
{
    private readonly WinForms.ComboBox _outputComboBox = new();
    private readonly WinForms.TextBox _iterationsTextBox = new();
    private readonly WinForms.DataGridView _rulesGrid = new();

    public StressRunOptions? Options { get; private set; }

    public StressAnalysisOptionsDialog(
        IReadOnlyList<TaggedInput> inputs,
        IReadOnlyList<TaggedOutput> outputs,
        string? defaultOutputId,
        int defaultIterations)
    {
        Text = "Stress Analysis";
        Width = 920;
        Height = 560;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        _iterationsTextBox.Text = defaultIterations.ToString();

        var panel = new WinForms.TableLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            Padding = new WinForms.Padding(14),
            ColumnCount = 2,
            RowCount = 6
        };
        panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Absolute, 170));
        panel.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100));

        AddControlRow(panel, "Primary output", _outputComboBox);
        AddControlRow(panel, "Iterations per run", _iterationsTextBox);

        var helperText = new WinForms.Label
        {
            AutoSize = true,
            Text =
                "MonteCarlo.XL will run the current model twice using the same comparison seed: first baseline, then stressed.\r\n" +
                "Fixed value replaces all samples. Add shift adds a constant. Range scale multiplies the deviation from the input's base workbook value.",
            Margin = new WinForms.Padding(0, 0, 0, 8)
        };
        panel.SetColumnSpan(helperText, 2);
        panel.Controls.Add(helperText);

        ConfigureRulesGrid(inputs);
        panel.SetColumnSpan(_rulesGrid, 2);
        panel.Controls.Add(_rulesGrid);

        var footer = new WinForms.Label
        {
            AutoSize = true,
            Text = "Select one or more inputs to stress. Unchecked rows are left unchanged.",
            Margin = new WinForms.Padding(0, 8, 0, 0)
        };
        panel.SetColumnSpan(footer, 2);
        panel.Controls.Add(footer);

        var buttons = new WinForms.FlowLayoutPanel
        {
            FlowDirection = WinForms.FlowDirection.RightToLeft,
            Dock = WinForms.DockStyle.Fill
        };
        var okButton = new WinForms.Button { Text = "Run Stress Analysis", DialogResult = WinForms.DialogResult.OK, AutoSize = true };
        var cancelButton = new WinForms.Button { Text = "Cancel", DialogResult = WinForms.DialogResult.Cancel, AutoSize = true };
        buttons.Controls.Add(okButton);
        buttons.Controls.Add(cancelButton);
        panel.SetColumnSpan(buttons, 2);
        panel.Controls.Add(buttons);

        Controls.Add(panel);
        AcceptButton = okButton;
        CancelButton = cancelButton;

        foreach (var output in outputs)
        {
            var option = new StressOutputOption(output);
            _outputComboBox.Items.Add(option);
            if (defaultOutputId != null && string.Equals(output.Cell.FullReference, defaultOutputId, StringComparison.OrdinalIgnoreCase))
                _outputComboBox.SelectedItem = option;
        }

        if (_outputComboBox.SelectedItem == null && _outputComboBox.Items.Count > 0)
            _outputComboBox.SelectedIndex = 0;

        okButton.Click += (_, _) =>
        {
            if (!TryReadOptions(out var options, out var error))
            {
                WinForms.MessageBox.Show(error, "Stress Analysis", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                DialogResult = WinForms.DialogResult.None;
                return;
            }

            Options = options;
        };
    }

    private void ConfigureRulesGrid(IReadOnlyList<TaggedInput> inputs)
    {
        _rulesGrid.AllowUserToAddRows = false;
        _rulesGrid.AllowUserToDeleteRows = false;
        _rulesGrid.AllowUserToResizeRows = false;
        _rulesGrid.AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill;
        _rulesGrid.RowHeadersVisible = false;
        _rulesGrid.SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect;
        _rulesGrid.MultiSelect = false;
        _rulesGrid.Height = 310;

        _rulesGrid.Columns.Add(new WinForms.DataGridViewCheckBoxColumn
        {
            HeaderText = "Apply",
            FillWeight = 12
        });
        _rulesGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
        {
            HeaderText = "Input",
            ReadOnly = true,
            FillWeight = 34
        });
        _rulesGrid.Columns.Add(new WinForms.DataGridViewComboBoxColumn
        {
            HeaderText = "Mode",
            DataSource = StressModeOption.All,
            DisplayMember = nameof(StressModeOption.Label),
            ValueMember = nameof(StressModeOption.Mode),
            FillWeight = 24
        });
        _rulesGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
        {
            HeaderText = "Value",
            FillWeight = 18
        });
        _rulesGrid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
        {
            HeaderText = "Cell",
            ReadOnly = true,
            FillWeight = 12
        });

        foreach (var input in inputs)
        {
            var rowIndex = _rulesGrid.Rows.Add(
                false,
                input.Label,
                StressRuleMode.AddShift,
                "0",
                input.Cell.FullReference);
            _rulesGrid.Rows[rowIndex].Tag = input;
        }
    }

    private bool TryReadOptions(out StressRunOptions? options, out string error)
    {
        options = null;
        error = string.Empty;

        if (_outputComboBox.SelectedItem is not StressOutputOption output)
        {
            error = "Choose the primary output to highlight in the stress report.";
            return false;
        }

        if (!int.TryParse(_iterationsTextBox.Text, out var iterations) || iterations <= 0)
        {
            error = "Iterations per run must be a positive whole number.";
            return false;
        }

        var rules = new List<StressInputRule>();
        foreach (WinForms.DataGridViewRow row in _rulesGrid.Rows)
        {
            var apply = row.Cells[0].Value is bool isApplied && isApplied;
            if (!apply)
                continue;

            if (row.Tag is not TaggedInput input)
                continue;

            var mode = row.Cells[2].Value is StressRuleMode selectedMode
                ? selectedMode
                : StressRuleMode.AddShift;

            var rawValue = Convert.ToString(row.Cells[3].Value)?.Trim();
            if (!double.TryParse(
                    rawValue,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var stressValue))
            {
                error = $"Enter a numeric stress value for '{input.Label}'.";
                return false;
            }

            if (mode == StressRuleMode.RangeScale && stressValue < 0)
            {
                error = $"Range scale for '{input.Label}' must be zero or greater.";
                return false;
            }

            rules.Add(new StressInputRule(input.Cell.FullReference, input.Label, mode, stressValue));
        }

        if (rules.Count == 0)
        {
            error = "Select at least one input stress rule.";
            return false;
        }

        options = new StressRunOptions(
            output.Output.Cell.FullReference,
            output.Output.Label,
            iterations,
            new StressInputPlan(rules));
        return true;
    }

    private static void AddControlRow(WinForms.TableLayoutPanel panel, string labelText, WinForms.Control control)
    {
        panel.Controls.Add(new WinForms.Label
        {
            Text = labelText,
            AutoSize = true,
            Anchor = WinForms.AnchorStyles.Left,
            Margin = new WinForms.Padding(0, 6, 10, 6)
        });

        control.Dock = WinForms.DockStyle.Fill;
        control.Margin = new WinForms.Padding(0, 3, 0, 3);
        panel.Controls.Add(control);
    }

    private sealed record StressOutputOption(TaggedOutput Output)
    {
        public override string ToString() => Output.Label;
    }

    private sealed record StressModeOption(StressRuleMode Mode, string Label)
    {
        public static IReadOnlyList<StressModeOption> All { get; } =
        [
            new(StressRuleMode.FixedValue, "Fixed value"),
            new(StressRuleMode.AddShift, "Add shift"),
            new(StressRuleMode.RangeScale, "Range scale")
        ];
    }
}
