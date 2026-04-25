using WinForms = System.Windows.Forms;

namespace MonteCarlo.Addin.Services;

internal enum RestoreConflictResolution
{
    OverwriteAll,
    SkipChangedCells
}

internal sealed record FormulaRestoreConflict(
    FormulaSwapEntry Entry,
    string CurrentState);

internal sealed record FormulaRestorePreview(
    IReadOnlyList<FormulaSwapEntry> Entries,
    IReadOnlyList<FormulaRestoreConflict> Conflicts);

internal sealed class FormulaCatalogPreviewDialog : WinForms.Form
{
    public FormulaCatalogPreviewDialog(IReadOnlyList<FormulaSwapEntry> entries)
    {
        Text = "Replace MC Formulas";
        Width = 940;
        Height = 560;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new WinForms.TableLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            Padding = new WinForms.Padding(14),
            ColumnCount = 1,
            RowCount = 4
        };

        layout.Controls.Add(new WinForms.Label
        {
            AutoSize = true,
            Text =
                "Review the MC.* formulas that will be replaced with their current values. " +
                "MonteCarlo.XL will also save a restore map inside the workbook.",
            Margin = new WinForms.Padding(0, 0, 0, 8)
        });

        layout.Controls.Add(new WinForms.Label
        {
            AutoSize = true,
            Text = BuildSheetSummary(entries),
            Margin = new WinForms.Padding(0, 0, 0, 8)
        });

        layout.Controls.Add(BuildGrid(entries, includeCurrentState: false));

        var buttons = BuildButtons("Replace Formulas");
        layout.Controls.Add(buttons.Panel);

        Controls.Add(layout);
        AcceptButton = buttons.OkButton;
        CancelButton = buttons.CancelButton;
    }

    private static string BuildSheetSummary(IReadOnlyList<FormulaSwapEntry> entries)
    {
        var sheetSummary = entries
            .GroupBy(entry => entry.SheetName)
            .Select(group => $"{group.Key} ({group.Count():N0})")
            .ToList();
        return $"Found {entries.Count:N0} MC.* formulas across {sheetSummary.Count:N0} worksheet(s): {string.Join(", ", sheetSummary)}";
    }

    internal static WinForms.DataGridView BuildGrid(
        IReadOnlyList<FormulaSwapEntry> entries,
        bool includeCurrentState,
        IReadOnlyDictionary<string, string>? currentStateByKey = null)
    {
        var grid = new WinForms.DataGridView
        {
            Dock = WinForms.DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false
        };

        grid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
        {
            HeaderText = "Sheet",
            FillWeight = 18
        });
        grid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
        {
            HeaderText = "Cell",
            FillWeight = 12
        });
        grid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
        {
            HeaderText = "Current Value",
            FillWeight = 18
        });
        if (includeCurrentState)
        {
            grid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
            {
                HeaderText = "Current Cell State",
                FillWeight = 26
            });
        }
        grid.Columns.Add(new WinForms.DataGridViewTextBoxColumn
        {
            HeaderText = "Formula",
            FillWeight = includeCurrentState ? 26 : 52
        });

        foreach (var entry in entries)
        {
            var key = BuildKey(entry);
            var currentState = includeCurrentState && currentStateByKey != null && currentStateByKey.TryGetValue(key, out var state)
                ? state
                : string.Empty;

            if (includeCurrentState)
            {
                grid.Rows.Add(entry.SheetName, entry.CellAddress, entry.ValueText, currentState, entry.Formula);
            }
            else
            {
                grid.Rows.Add(entry.SheetName, entry.CellAddress, entry.ValueText, entry.Formula);
            }
        }

        return grid;
    }

    internal static string BuildKey(FormulaSwapEntry entry) =>
        $"{entry.SheetName}!{entry.CellAddress}";

    internal static (WinForms.FlowLayoutPanel Panel, WinForms.Button OkButton, WinForms.Button CancelButton) BuildButtons(string okText)
    {
        var panel = new WinForms.FlowLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            FlowDirection = WinForms.FlowDirection.RightToLeft,
            AutoSize = true
        };

        var okButton = new WinForms.Button
        {
            Text = okText,
            DialogResult = WinForms.DialogResult.OK,
            AutoSize = true
        };
        var cancelButton = new WinForms.Button
        {
            Text = "Cancel",
            DialogResult = WinForms.DialogResult.Cancel,
            AutoSize = true
        };

        panel.Controls.Add(okButton);
        panel.Controls.Add(cancelButton);
        return (panel, okButton, cancelButton);
    }
}

internal sealed class FormulaRestorePreviewDialog : WinForms.Form
{
    private readonly WinForms.CheckBox _skipChangedCellsCheckBox = new()
    {
        AutoSize = true,
        Checked = true,
        Text = "Skip cells whose current value or formula changed after replacement"
    };

    public RestoreConflictResolution Resolution =>
        _skipChangedCellsCheckBox.Visible && _skipChangedCellsCheckBox.Checked
            ? RestoreConflictResolution.SkipChangedCells
            : RestoreConflictResolution.OverwriteAll;

    public FormulaRestorePreviewDialog(FormulaRestorePreview preview)
    {
        Text = "Restore MC Formulas";
        Width = 980;
        Height = 600;
        FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        StartPosition = WinForms.FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;

        var layout = new WinForms.TableLayoutPanel
        {
            Dock = WinForms.DockStyle.Fill,
            Padding = new WinForms.Padding(14),
            ColumnCount = 1,
            RowCount = 5
        };

        var hasConflicts = preview.Conflicts.Count > 0;
        layout.Controls.Add(new WinForms.Label
        {
            AutoSize = true,
            Text = hasConflicts
                ? "Some cells no longer match the values that were originally written during replacement. Review them before restoring formulas."
                : "Review the cells that will have their MC.* formulas restored from the workbook restore map.",
            Margin = new WinForms.Padding(0, 0, 0, 8)
        });

        layout.Controls.Add(new WinForms.Label
        {
            AutoSize = true,
            Text = hasConflicts
                ? $"Restore map contains {preview.Entries.Count:N0} cells, and {preview.Conflicts.Count:N0} currently look changed."
                : $"Restore map contains {preview.Entries.Count:N0} cells and none currently look changed.",
            Margin = new WinForms.Padding(0, 0, 0, 8)
        });

        _skipChangedCellsCheckBox.Visible = hasConflicts;
        if (hasConflicts)
        {
            layout.Controls.Add(_skipChangedCellsCheckBox);
        }

        var currentStateByKey = preview.Conflicts.ToDictionary(
            conflict => FormulaCatalogPreviewDialog.BuildKey(conflict.Entry),
            conflict => conflict.CurrentState,
            StringComparer.OrdinalIgnoreCase);
        layout.Controls.Add(FormulaCatalogPreviewDialog.BuildGrid(preview.Entries, includeCurrentState: true, currentStateByKey));

        var buttons = FormulaCatalogPreviewDialog.BuildButtons("Restore Formulas");
        layout.Controls.Add(buttons.Panel);

        Controls.Add(layout);
        AcceptButton = buttons.OkButton;
        CancelButton = buttons.CancelButton;
    }
}
