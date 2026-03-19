# TASK-016: Correlation Matrix UI

## Context

Read `ROADMAP.md` for full project context. TASK-015 built the Iman-Conover engine. This task builds the UI for users to define the correlation matrix between their simulation inputs.

## Dependencies

- TASK-015 (CorrelationMatrix model + validation)
- TASK-007 (GlobalStyles)
- TASK-008 (SetupView — the correlation editor integrates into or links from the setup flow)

## Objective

Build a `CorrelationView` in the task pane where users can define pairwise correlations between their inputs via an editable matrix grid. The UI must validate that the matrix is positive semi-definite and guide the user if it's not.

## Design

### Layout

```
┌──────────────────────────────────────┐
│ INPUT CORRELATIONS          [? Help] │
├──────────────────────────────────────┤
│                                      │
│ Tip: Enter Spearman rank             │
│ correlations (-1 to +1).             │
│ Leave as 0 for independent inputs.   │
│                                      │
│        MatCost  Labor  IntRate       │
│ MatCost  [1.0]  [0.6]  [0.0]        │
│ Labor    [0.6]  [1.0]  [-0.3]       │
│ IntRate  [0.0]  [-0.3] [1.0]        │
│                                      │
│ ┌──────────────────────────────────┐ │
│ │ ✓ Matrix is valid (PSD)         │ │  ← Green if valid
│ └──────────────────────────────────┘ │
│                                      │
│    — OR —                            │
│                                      │
│ ┌──────────────────────────────────┐ │
│ │ ⚠ Matrix is not positive semi-  │ │  ← Amber warning
│ │   definite. [Auto-fix]          │ │
│ └──────────────────────────────────┘ │
│                                      │
│ [Clear All]         [Apply & Close]  │
└──────────────────────────────────────┘
```

### Matrix Grid

- The grid shows only inputs that have been configured in the Setup view
- Diagonal cells (self-correlation) are fixed at 1.0, grayed out, non-editable
- Upper-triangle cells are editable; lower-triangle cells mirror them automatically (symmetric)
- Each cell accepts a decimal value from -1.0 to 1.0
- Color-coded cells:
  - Positive correlation: blue shading (intensity proportional to magnitude)
  - Negative correlation: red/orange shading (intensity proportional to magnitude)
  - Zero: white/neutral
  - e.g., 0.8 → strong blue; -0.5 → medium orange; 0.0 → white

### Validation Feedback

After any cell edit:
1. Validate the matrix using `CorrelationMatrix.Validate()`
2. If valid: show green checkmark + "Matrix is valid"
3. If not PSD: show amber warning + "Auto-fix" button
4. If values out of range: show red error on the specific cell

**Auto-fix:** When clicked, calls `CorrelationMatrix.EnsurePositiveSemiDefinite()` and updates the grid cells with the corrected values. Show a brief animation or highlight on the cells that changed.

### Access from Setup View

Add a "Correlations" button or link at the bottom of the inputs section in SetupView:

```
│ INPUTS                       [+ Add]│
│ ┌──────────────────────────────────┐ │
│ │ ... input cards ...              │ │
│ └──────────────────────────────────┘ │
│ [📊 Define Correlations (3 inputs)]  │  ← Links to CorrelationView
```

The correlation button shows the count of inputs and is disabled if fewer than 2 inputs are configured.

## CorrelationViewModel

```csharp
public partial class CorrelationViewModel : ObservableObject
{
    [ObservableProperty] private ObservableCollection<string> _inputLabels;
    [ObservableProperty] private double[,] _matrixValues;
    [ObservableProperty] private bool _isValid;
    [ObservableProperty] private string _validationMessage;
    [ObservableProperty] private bool _canAutoFix;

    /// Called when any cell value changes.
    public void OnCellValueChanged(int row, int col, double value)
    {
        // Mirror to the symmetric cell
        _matrixValues[col, row] = value;
        // Re-validate
        ValidateMatrix();
    }

    [RelayCommand]
    private void AutoFix()
    {
        var matrix = new CorrelationMatrix(_matrixValues);
        var fixed = matrix.EnsurePositiveSemiDefinite();
        MatrixValues = fixed.ToArray();
        ValidateMatrix();
    }

    [RelayCommand]
    private void ClearAll()
    {
        // Reset all off-diagonal values to 0
    }

    [RelayCommand]
    private void Apply()
    {
        // Save to SimulationConfig and navigate back to SetupView
    }
}
```

## WPF Grid Control

Build a custom `CorrelationMatrixGrid` control:

```csharp
public class CorrelationMatrixGrid : UserControl
{
    // Dynamically generates a Grid with TextBox cells based on input count
    // Row/column headers show input labels
    // Diagonal cells are read-only
    // Lower triangle mirrors upper triangle
}
```

For the cell color coding, use a value-to-color converter:

```csharp
public class CorrelationColorConverter : IValueConverter
{
    public object Convert(object value, ...)
    {
        double corr = (double)value;
        if (corr > 0)
            return new SolidColorBrush(Color.FromArgb(
                (byte)(Math.Abs(corr) * 180),  // Alpha scales with magnitude
                59, 130, 246));                  // Blue-500 RGB
        else if (corr < 0)
            return new SolidColorBrush(Color.FromArgb(
                (byte)(Math.Abs(corr) * 180),
                249, 115, 22));                  // Orange-500 RGB
        return Brushes.Transparent;
    }
}
```

### Scrolling for Many Inputs

If there are more than ~6 inputs, the matrix won't fit in the 380px task pane. Wrap the grid in a `ScrollViewer` with both horizontal and vertical scrolling. Keep the row/column headers fixed (sticky) — this requires either a frozen header row/column or a more sophisticated layout.

For simplicity, if inputs > 8, show a message suggesting the user focus on the most important correlations and leave others at 0.

## Config Persistence Update

The correlation matrix needs to be saved as part of the `SimulationProfile` (TASK-011). Update `SavedProfile` to include:

```csharp
public class SimulationProfile
{
    // ... existing fields ...
    public double[,]? CorrelationMatrix { get; set; }  // null = no correlation (independent)
}
```

And update `ConfigPersistence` to serialize/deserialize the matrix.

## File Structure

```
MonteCarlo.UI/
├── Views/
│   ├── CorrelationView.xaml/.cs
│   └── CorrelationMatrixGrid.xaml/.cs    # Custom grid control
├── ViewModels/
│   └── CorrelationViewModel.cs
└── Converters/
    └── CorrelationColorConverter.cs
```

## Commit Strategy

```
feat(ui): add CorrelationMatrixGrid control with color-coded editable cells
feat(ui): add CorrelationView with validation feedback and auto-fix
feat(ui): add correlation access from SetupView
feat(ui): persist correlation matrix in SimulationProfile
```

## Done When

- [ ] Matrix grid renders with input labels and editable cells
- [ ] Diagonal locked at 1.0, lower triangle mirrors upper
- [ ] Cells color-coded by correlation strength and sign
- [ ] Real-time PSD validation with clear status message
- [ ] Auto-fix button corrects non-PSD matrices
- [ ] Correlation matrix persists with the workbook config
- [ ] Accessible from Setup view when 2+ inputs configured
- [ ] Scrollable for many inputs
- [ ] `dotnet build` clean
