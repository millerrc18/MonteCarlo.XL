namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Applies visual formatting to tagged cells so the user can identify inputs/outputs at a glance.
/// </summary>
public interface ICellHighlighter
{
    /// <summary>
    /// Apply input highlighting (subtle blue background) to a cell.
    /// </summary>
    void HighlightInput(CellReference cell);

    /// <summary>
    /// Apply output highlighting (subtle green background) to a cell.
    /// </summary>
    void HighlightOutput(CellReference cell);

    /// <summary>
    /// Remove highlighting from a cell (sets to no fill).
    /// </summary>
    void ClearHighlight(CellReference cell);

    /// <summary>
    /// Refresh all highlights based on current input/output tags.
    /// </summary>
    void RefreshAll(IInputTagManager inputs, IOutputTagManager outputs);
}
