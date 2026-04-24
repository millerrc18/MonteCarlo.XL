using WinForms = System.Windows.Forms;

namespace MonteCarlo.Addin.Services;

internal sealed record RunSeedSelection(int? RandomSeed, string Description);

internal sealed class RunSeedPromptDialog : WinForms.Form
{
    private readonly WinForms.RadioButton _randomRadio = new()
    {
        Text = "Use a random seed for this run",
        Checked = true,
        AutoSize = true
    };

    private readonly WinForms.RadioButton _fixedRadio = new()
    {
        Text = "Use a fixed seed for this run",
        AutoSize = true
    };

    private readonly WinForms.TextBox _fixedSeedTextBox = new();

    public RunSeedSelection? Selection { get; private set; }

    public RunSeedPromptDialog(string workflowName, int suggestedFixedSeed)
    {
        Text = "Choose Run Seed";
        Width = 420;
        Height = 250;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var panel = new WinForms.TableLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            Padding = new WinForms.Padding(14),
            ColumnCount = 1,
            RowCount = 7
        };

        var introLabel = new WinForms.Label
        {
            Text = $"The current settings are configured to ask for a seed before {workflowName.ToLowerInvariant()}.\r\n\r\nChoose how this run should behave:",
            AutoSize = true,
            MaximumSize = new System.Drawing.Size(360, 0)
        };

        _fixedSeedTextBox.Text = suggestedFixedSeed.ToString();
        _fixedSeedTextBox.Width = 120;
        _fixedSeedTextBox.Enabled = false;

        var fixedSeedPanel = new WinForms.FlowLayoutPanel
        {
            FlowDirection = WinForms.FlowDirection.LeftToRight,
            AutoSize = true
        };
        fixedSeedPanel.Controls.Add(new WinForms.Label { Text = "Seed:", AutoSize = true, Margin = new WinForms.Padding(0, 6, 4, 0) });
        fixedSeedPanel.Controls.Add(_fixedSeedTextBox);

        var buttons = new WinForms.FlowLayoutPanel
        {
            FlowDirection = WinForms.FlowDirection.RightToLeft,
            Dock = WinForms.DockStyle.Fill
        };
        var okButton = new WinForms.Button { Text = "Continue", DialogResult = WinForms.DialogResult.OK, AutoSize = true };
        var cancelButton = new WinForms.Button { Text = "Cancel", DialogResult = WinForms.DialogResult.Cancel, AutoSize = true };
        buttons.Controls.Add(okButton);
        buttons.Controls.Add(cancelButton);

        panel.Controls.Add(introLabel);
        panel.Controls.Add(_randomRadio);
        panel.Controls.Add(_fixedRadio);
        panel.Controls.Add(fixedSeedPanel);
        panel.Controls.Add(new WinForms.Label
        {
            Text = "This choice applies only to the current run.",
            AutoSize = true,
            MaximumSize = new System.Drawing.Size(360, 0)
        });
        panel.Controls.Add(new WinForms.Label());
        panel.Controls.Add(buttons);

        Controls.Add(panel);
        AcceptButton = okButton;
        CancelButton = cancelButton;

        _randomRadio.CheckedChanged += (_, _) => UpdateSeedControls();
        _fixedRadio.CheckedChanged += (_, _) => UpdateSeedControls();

        okButton.Click += (_, _) =>
        {
            if (_fixedRadio.Checked)
            {
                if (!int.TryParse(_fixedSeedTextBox.Text, out var fixedSeed) || fixedSeed < 0)
                {
                    WinForms.MessageBox.Show(
                        "Enter a fixed seed of zero or greater.",
                        "Choose Run Seed",
                        WinForms.MessageBoxButtons.OK,
                        WinForms.MessageBoxIcon.Warning);
                    DialogResult = WinForms.DialogResult.None;
                    return;
                }

                Selection = new RunSeedSelection(fixedSeed, $"Fixed {fixedSeed}");
                return;
            }

            Selection = new RunSeedSelection(null, "Random this run");
        };
    }

    private void UpdateSeedControls()
    {
        _fixedSeedTextBox.Enabled = _fixedRadio.Checked;
    }
}
