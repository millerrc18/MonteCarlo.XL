namespace MonteCarlo.UI.Services;

/// <summary>
/// Service that mediates cell selection between the WPF UI and Excel.
/// The Addin layer implements the actual Excel event hookup;
/// this service provides the contract the UI uses.
/// </summary>
public class CellSelectionService
{
    private Action<string, string, string?>? _callback;
    private bool _isActive;

    /// <summary>Whether cell selection mode is currently active.</summary>
    public bool IsActive => _isActive;

    /// <summary>
    /// Begin cell selection. The callback receives (sheetName, cellAddress, suggestedLabel).
    /// </summary>
    public void BeginSelection(Action<string, string, string?> callback)
    {
        _callback = callback;
        _isActive = true;
        SelectionStarted?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Called by the Addin layer when the user clicks a cell in Excel.
    /// </summary>
    public void NotifyCellSelected(string sheetName, string cellAddress, string? suggestedLabel)
    {
        if (!_isActive) return;
        _isActive = false;
        _callback?.Invoke(sheetName, cellAddress, suggestedLabel);
        _callback = null;
        SelectionCompleted?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Cancel the current selection.</summary>
    public void CancelSelection()
    {
        _isActive = false;
        _callback = null;
        SelectionCompleted?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Raised when selection mode starts.</summary>
    public event EventHandler? SelectionStarted;

    /// <summary>Raised when selection mode ends (completed or cancelled).</summary>
    public event EventHandler? SelectionCompleted;
}
