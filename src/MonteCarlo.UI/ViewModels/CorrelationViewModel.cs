using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Correlation;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// ViewModel for the correlation matrix editor view.
/// </summary>
public partial class CorrelationViewModel : ObservableObject
{
    [ObservableProperty] private ObservableCollection<string> _inputLabels = new();
    [ObservableProperty] private double[,]? _matrixValues;
    [ObservableProperty] private bool _isValid = true;
    [ObservableProperty] private string _validationMessage = "No inputs configured.";
    [ObservableProperty] private string? _warningMessage;
    [ObservableProperty] private string? _diagnosticMessage;
    [ObservableProperty] private string? _rangeStatusMessage;
    [ObservableProperty] private bool _isIndependent = true;
    [ObservableProperty] private bool _canAutoFix;

    /// <summary>
    /// Raised when the user clicks Apply & Close.
    /// </summary>
    public event Action<double[,]?>? Applied;

    /// <summary>
    /// Raised when the user wants to import a matrix from the selected Excel range.
    /// </summary>
    public event Action? ImportRequested;

    /// <summary>
    /// Raised when the user wants to export the current matrix to Excel.
    /// </summary>
    public event Action? ExportRequested;

    /// <summary>
    /// Raised when the user wants to navigate back.
    /// </summary>
    public event Action? CloseRequested;

    /// <summary>
    /// Initializes the correlation editor with current inputs and optional existing matrix.
    /// </summary>
    /// <param name="inputLabels">Labels for the simulation inputs.</param>
    /// <param name="existingMatrix">Existing correlation values, or null for identity.</param>
    public void Initialize(IList<string> inputLabels, double[,]? existingMatrix = null)
    {
        InputLabels = new ObservableCollection<string>(inputLabels);
        int n = inputLabels.Count;

        if (n < 2)
        {
            MatrixValues = null;
            IsValid = false;
            ValidationMessage = "At least 2 inputs are required for correlation.";
            CanAutoFix = false;
            return;
        }

        if (existingMatrix != null && existingMatrix.GetLength(0) == n)
        {
            MatrixValues = (double[,])existingMatrix.Clone();
        }
        else
        {
            // Default: identity (independent inputs)
            var identity = new double[n, n];
            for (int i = 0; i < n; i++)
                identity[i, i] = 1.0;
            MatrixValues = identity;
        }

        ValidateMatrix();
    }

    /// <summary>
    /// Called when a cell value changes in the grid.
    /// </summary>
    public void OnCellValueChanged(int row, int col, double value)
    {
        if (MatrixValues == null) return;

        value = Math.Clamp(value, -1.0, 1.0);
        MatrixValues[row, col] = value;
        MatrixValues[col, row] = value; // Mirror

        // Trigger UI update by re-setting the property
        OnPropertyChanged(nameof(MatrixValues));
        ValidateMatrix();
    }

    [RelayCommand]
    private void AutoFix()
    {
        if (MatrixValues == null) return;

        var matrix = new CorrelationMatrix(MatrixValues);
        var diagnostics = matrix.Analyze(InputLabels);
        var corrected = matrix.EnsurePositiveSemiDefinite();
        MatrixValues = corrected.ToArray();
        RangeStatusMessage = BuildAutoFixStatusMessage(diagnostics);
        ValidateMatrix();
    }

    [RelayCommand]
    private void RequestImport()
    {
        RangeStatusMessage = "Reading selected Excel range...";
        ImportRequested?.Invoke();
    }

    [RelayCommand]
    private void RequestExport()
    {
        RangeStatusMessage = "Writing matrix to selected Excel range...";
        ExportRequested?.Invoke();
    }

    public void LoadImportedMatrix(double[,] matrix, string sourceAddress)
    {
        MatrixValues = (double[,])matrix.Clone();
        RangeStatusMessage = $"Imported matrix from {sourceAddress}.";
        ValidateMatrix();
    }

    public void ShowRangeError(string message)
    {
        RangeStatusMessage = message;
    }

    public void MarkExported(string targetAddress)
    {
        RangeStatusMessage = $"Exported matrix to {targetAddress}.";
    }

    [RelayCommand]
    private void ClearAll()
    {
        if (MatrixValues == null) return;

        int n = MatrixValues.GetLength(0);
        var cleared = new double[n, n];
        for (int i = 0; i < n; i++)
            cleared[i, i] = 1.0;
        MatrixValues = cleared;
        ValidateMatrix();
    }

    [RelayCommand]
    private void Apply()
    {
        // Return null if it's an identity matrix (no correlation)
        if (MatrixValues != null && IsIdentity(MatrixValues))
        {
            Applied?.Invoke(null);
        }
        else
        {
            Applied?.Invoke(MatrixValues);
        }
        CloseRequested?.Invoke();
    }

    [RelayCommand]
    private void Cancel()
    {
        CloseRequested?.Invoke();
    }

    private void ValidateMatrix()
    {
        WarningMessage = null;
        DiagnosticMessage = null;
        IsIndependent = false;

        if (MatrixValues == null)
        {
            IsValid = false;
            ValidationMessage = "No matrix to validate.";
            CanAutoFix = false;
            return;
        }

        if (IsIdentity(MatrixValues))
        {
            IsValid = true;
            IsIndependent = true;
            ValidationMessage = "No correlations configured. Inputs will be sampled independently.";
            CanAutoFix = false;
            return;
        }

        try
        {
            var matrix = new CorrelationMatrix(MatrixValues);
            matrix.Validate();
            var diagnostics = matrix.Analyze(InputLabels);
            IsValid = true;
            ValidationMessage = "✓ Matrix is valid (positive semi-definite)";
            WarningMessage = BuildWarningMessage(diagnostics);
            DiagnosticMessage = BuildDiagnosticMessage(diagnostics);
            CanAutoFix = false;
        }
        catch (ArgumentException ex)
        {
            IsValid = false;
            if (ex.Message.Contains("positive semi-definite", StringComparison.OrdinalIgnoreCase))
            {
                var diagnostics = new CorrelationMatrix(MatrixValues).Analyze(InputLabels);
                ValidationMessage = $"⚠ Matrix is not positive semi-definite (min eigenvalue {diagnostics.MinimumEigenvalue:0.###}).";
                WarningMessage = BuildWarningMessage(diagnostics);
                DiagnosticMessage = BuildDiagnosticMessage(diagnostics);
                CanAutoFix = true;
            }
            else
            {
                ValidationMessage = $"✗ {ex.Message}";
                CanAutoFix = false;
            }
        }
    }

    private static string BuildAutoFixStatusMessage(CorrelationDiagnostics diagnostics)
    {
        if (diagnostics.RecommendedAdjustments.Count == 0)
            return "Auto-fix replaced the matrix with the nearest positive semi-definite version.";

        var topAdjustment = diagnostics.RecommendedAdjustments[0];
        return $"Auto-fix adjusted {diagnostics.RecommendedAdjustments.Count} pair(s). First change: {FormatAdjustment(topAdjustment)}.";
    }

    private static string? BuildWarningMessage(CorrelationDiagnostics diagnostics)
    {
        var warnings = new List<string>();
        var maxAbs = diagnostics.HighestMagnitudePairs.FirstOrDefault()?.AbsoluteValue ?? 0.0;

        if (maxAbs >= 0.95)
            warnings.Add("One or more correlations are extremely high (|r| >= 0.95), which can make the matrix fragile.");
        else if (maxAbs >= 0.90)
            warnings.Add("One or more correlations are very high (|r| >= 0.90).");

        if (diagnostics.MinimumEigenvalue < 0.05)
            warnings.Add("The matrix is close to singular. Small edits may make it invalid.");

        return warnings.Count == 0 ? null : string.Join(" ", warnings);
    }

    private static string? BuildDiagnosticMessage(CorrelationDiagnostics diagnostics)
    {
        if (!diagnostics.IsPositiveSemiDefinite)
        {
            if (diagnostics.RecommendedAdjustments.Count == 0)
                return "Auto-fix can move the matrix to the nearest positive semi-definite version.";

            var suggestions = string.Join("; ", diagnostics.RecommendedAdjustments.Select(FormatAdjustment));
            return $"Suggested pair adjustments for the nearest valid matrix: {suggestions}.";
        }

        var fragilePairs = diagnostics.HighestMagnitudePairs
            .Where(pair => pair.AbsoluteValue >= 0.90)
            .Select(FormatPair)
            .ToArray();

        if (fragilePairs.Length == 0 && diagnostics.MinimumEigenvalue >= 0.05)
            return null;

        if (fragilePairs.Length == 0)
            return $"Minimum eigenvalue is {diagnostics.MinimumEigenvalue:0.###}. The matrix is valid but fragile.";

        return $"Highest-magnitude pairs: {string.Join("; ", fragilePairs)}.";
    }

    private static string FormatPair(CorrelationPairSummary pair)
    {
        return $"{pair.RowLabel} vs {pair.ColumnLabel} {pair.Value:0.###}";
    }

    private static string FormatAdjustment(CorrelationPairAdjustment adjustment)
    {
        var verb = adjustment.ReducesMagnitude ? "reduce" : "adjust";
        return $"{adjustment.RowLabel} vs {adjustment.ColumnLabel}: {verb} {adjustment.CurrentValue:0.###} to {adjustment.RecommendedValue:0.###}";
    }

    private static bool IsIdentity(double[,] matrix)
    {
        int n = matrix.GetLength(0);
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n; j++)
            {
                double expected = i == j ? 1.0 : 0.0;
                if (Math.Abs(matrix[i, j] - expected) > 1e-10)
                    return false;
            }
        }
        return true;
    }
}
