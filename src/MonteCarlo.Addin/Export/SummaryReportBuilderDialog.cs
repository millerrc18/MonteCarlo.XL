using WinForms = System.Windows.Forms;

namespace MonteCarlo.Addin.Export;

internal sealed record SummaryReportBuilderSelection(bool ExportAllOutputs, SummaryReportOptions Options);

internal sealed class SummaryReportBuilderDialog : WinForms.Form
{
    private readonly WinForms.RadioButton _selectedOutputRadioButton = new()
    {
        AutoSize = true,
        Text = "Current output only"
    };

    private readonly WinForms.RadioButton _allOutputsRadioButton = new()
    {
        AutoSize = true,
        Text = "All outputs in one report sheet"
    };

    private readonly WinForms.CheckBox _metadataCheckBox = CreateCheckBox("Report metadata and run settings");
    private readonly WinForms.CheckBox _summaryStatsCheckBox = CreateCheckBox("Summary statistics");
    private readonly WinForms.CheckBox _percentilesCheckBox = CreateCheckBox("Percentiles");
    private readonly WinForms.CheckBox _targetAnalysisCheckBox = CreateCheckBox("Target analysis");
    private readonly WinForms.CheckBox _sensitivityCheckBox = CreateCheckBox("Sensitivity table and tornado");
    private readonly WinForms.CheckBox _scenarioAnalysisCheckBox = CreateCheckBox("Scenario analysis");
    private readonly WinForms.CheckBox _chartsCheckBox = CreateCheckBox("Histogram and CDF charts");
    private readonly WinForms.CheckBox _inputAssumptionsCheckBox = CreateCheckBox("Input assumptions");
    private readonly WinForms.CheckBox _correlationAssumptionsCheckBox = CreateCheckBox("Correlation assumptions");

    public SummaryReportBuilderSelection? Selection { get; private set; }

    public SummaryReportBuilderDialog(int outputCount, SummaryReportBuilderSelection? defaults = null)
    {
        Text = "Build Summary Report";
        Width = 560;
        Height = outputCount > 1 ? 560 : 500;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        defaults ??= new SummaryReportBuilderSelection(false, SummaryReportOptions.Default);
        ApplyDefaults(defaults.Options, defaults.ExportAllOutputs && outputCount > 1);

        var layout = new WinForms.TableLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            Padding = new WinForms.Padding(14),
            ColumnCount = 1,
            RowCount = 5
        };

        var introLabel = new WinForms.Label
        {
            AutoSize = true,
            Text =
                "Choose what MonteCarlo.XL should include in the summary report sheet. " +
                "These options affect Export Summary only; raw-data export is unchanged.",
            Margin = new WinForms.Padding(0, 0, 0, 10)
        };
        layout.Controls.Add(introLabel);

        if (outputCount > 1)
            layout.Controls.Add(BuildScopeGroup(outputCount));

        layout.Controls.Add(BuildSectionGroup());
        layout.Controls.Add(BuildPresetButtons());

        var footerLabel = new WinForms.Label
        {
            AutoSize = true,
            Text = "Charts adds histogram/CDF. Tornado appears only when sensitivity is included and data is available.",
            Margin = new WinForms.Padding(0, 8, 0, 0)
        };
        layout.Controls.Add(footerLabel);

        var buttons = new WinForms.FlowLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            FlowDirection = WinForms.FlowDirection.RightToLeft,
            AutoSize = true
        };
        var okButton = new WinForms.Button
        {
            Text = "Export Report",
            DialogResult = WinForms.DialogResult.OK,
            AutoSize = true
        };
        var cancelButton = new WinForms.Button
        {
            Text = "Cancel",
            DialogResult = WinForms.DialogResult.Cancel,
            AutoSize = true
        };
        buttons.Controls.Add(okButton);
        buttons.Controls.Add(cancelButton);
        layout.Controls.Add(buttons);

        Controls.Add(layout);
        AcceptButton = okButton;
        CancelButton = cancelButton;

        okButton.Click += (_, _) =>
        {
            var options = ReadOptions();
            if (!options.IncludesAnySections)
            {
                WinForms.MessageBox.Show(
                    "Select at least one report section before exporting.",
                    "Build Summary Report",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Warning);
                DialogResult = WinForms.DialogResult.None;
                return;
            }

            Selection = new SummaryReportBuilderSelection(
                ExportAllOutputs: outputCount > 1 && _allOutputsRadioButton.Checked,
                Options: options);
        };
    }

    private WinForms.Control BuildScopeGroup(int outputCount)
    {
        var group = new WinForms.GroupBox
        {
            Text = "Report scope",
            Dock = WinForms.DockStyle.Top,
            AutoSize = true,
            Padding = new WinForms.Padding(10)
        };

        var panel = new WinForms.FlowLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            FlowDirection = WinForms.FlowDirection.TopDown,
            AutoSize = true,
            WrapContents = false
        };

        panel.Controls.Add(_selectedOutputRadioButton);
        panel.Controls.Add(_allOutputsRadioButton);
        panel.Controls.Add(new WinForms.Label
        {
            AutoSize = true,
            Text = $"This simulation currently has {outputCount:N0} outputs available for export.",
            Margin = new WinForms.Padding(22, 4, 0, 0)
        });

        group.Controls.Add(panel);
        return group;
    }

    private WinForms.Control BuildSectionGroup()
    {
        var group = new WinForms.GroupBox
        {
            Text = "Sections",
            Dock = WinForms.DockStyle.Top,
            AutoSize = true,
            Padding = new WinForms.Padding(10)
        };

        var table = new WinForms.TableLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true
        };
        table.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 50));
        table.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 50));

        table.Controls.Add(_metadataCheckBox, 0, 0);
        table.Controls.Add(_summaryStatsCheckBox, 1, 0);
        table.Controls.Add(_percentilesCheckBox, 0, 1);
        table.Controls.Add(_targetAnalysisCheckBox, 1, 1);
        table.Controls.Add(_sensitivityCheckBox, 0, 2);
        table.Controls.Add(_scenarioAnalysisCheckBox, 1, 2);
        table.Controls.Add(_chartsCheckBox, 0, 3);
        table.Controls.Add(_inputAssumptionsCheckBox, 1, 3);
        table.Controls.Add(_correlationAssumptionsCheckBox, 0, 4);

        group.Controls.Add(table);
        return group;
    }

    private WinForms.Control BuildPresetButtons()
    {
        var panel = new WinForms.FlowLayoutPanel
        {
            Dock = WinForms.DockStyle.Top,
            FlowDirection = WinForms.FlowDirection.LeftToRight,
            AutoSize = true,
            Margin = new WinForms.Padding(0, 8, 0, 0)
        };

        var allButton = new WinForms.Button { Text = "All Sections", AutoSize = true };
        var presentationButton = new WinForms.Button { Text = "Presentation", AutoSize = true };
        var diagnosticsButton = new WinForms.Button { Text = "Diagnostics", AutoSize = true };

        allButton.Click += (_, _) => ApplyDefaults(SummaryReportOptions.Default, _allOutputsRadioButton.Checked);
        presentationButton.Click += (_, _) => ApplyDefaults(
            SummaryReportOptions.Default with
            {
                IncludeMetadata = true,
                IncludeSummaryStatistics = true,
                IncludePercentiles = true,
                IncludeTargetAnalysis = true,
                IncludeSensitivity = false,
                IncludeScenarioAnalysis = false,
                IncludeCharts = true,
                IncludeInputAssumptions = false,
                IncludeCorrelationAssumptions = false
            },
            _allOutputsRadioButton.Checked);
        diagnosticsButton.Click += (_, _) => ApplyDefaults(
            SummaryReportOptions.Default with
            {
                IncludeMetadata = true,
                IncludeSummaryStatistics = true,
                IncludePercentiles = true,
                IncludeTargetAnalysis = true,
                IncludeSensitivity = true,
                IncludeScenarioAnalysis = true,
                IncludeCharts = false,
                IncludeInputAssumptions = true,
                IncludeCorrelationAssumptions = true
            },
            _allOutputsRadioButton.Checked);

        panel.Controls.Add(new WinForms.Label
        {
            AutoSize = true,
            Text = "Presets:",
            Margin = new WinForms.Padding(0, 7, 6, 0)
        });
        panel.Controls.Add(allButton);
        panel.Controls.Add(presentationButton);
        panel.Controls.Add(diagnosticsButton);
        return panel;
    }

    private void ApplyDefaults(SummaryReportOptions options, bool exportAllOutputs)
    {
        _selectedOutputRadioButton.Checked = !exportAllOutputs;
        _allOutputsRadioButton.Checked = exportAllOutputs;
        _metadataCheckBox.Checked = options.IncludeMetadata;
        _summaryStatsCheckBox.Checked = options.IncludeSummaryStatistics;
        _percentilesCheckBox.Checked = options.IncludePercentiles;
        _targetAnalysisCheckBox.Checked = options.IncludeTargetAnalysis;
        _sensitivityCheckBox.Checked = options.IncludeSensitivity;
        _scenarioAnalysisCheckBox.Checked = options.IncludeScenarioAnalysis;
        _chartsCheckBox.Checked = options.IncludeCharts;
        _inputAssumptionsCheckBox.Checked = options.IncludeInputAssumptions;
        _correlationAssumptionsCheckBox.Checked = options.IncludeCorrelationAssumptions;
    }

    private SummaryReportOptions ReadOptions() =>
        new(
            IncludeMetadata: _metadataCheckBox.Checked,
            IncludeSummaryStatistics: _summaryStatsCheckBox.Checked,
            IncludePercentiles: _percentilesCheckBox.Checked,
            IncludeTargetAnalysis: _targetAnalysisCheckBox.Checked,
            IncludeSensitivity: _sensitivityCheckBox.Checked,
            IncludeScenarioAnalysis: _scenarioAnalysisCheckBox.Checked,
            IncludeCharts: _chartsCheckBox.Checked,
            IncludeInputAssumptions: _inputAssumptionsCheckBox.Checked,
            IncludeCorrelationAssumptions: _correlationAssumptionsCheckBox.Checked);

    private static WinForms.CheckBox CreateCheckBox(string text) =>
        new()
        {
            AutoSize = true,
            Checked = true,
            Text = text,
            Margin = new WinForms.Padding(0, 4, 18, 4)
        };
}
