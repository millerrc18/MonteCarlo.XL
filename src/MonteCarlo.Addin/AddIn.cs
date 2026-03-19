using ExcelDna.Integration;

namespace MonteCarlo.Addin;

/// <summary>
/// ExcelDna add-in entry point. Handles add-in lifecycle events.
/// </summary>
public class AddIn : IExcelAddIn
{
    /// <summary>
    /// Called when the add-in is loaded into Excel.
    /// </summary>
    public void AutoOpen()
    {
        // TODO: Initialize ribbon, register task pane, set up event handlers
    }

    /// <summary>
    /// Called when the add-in is unloaded from Excel.
    /// </summary>
    public void AutoClose()
    {
        // TODO: Clean up resources, unregister event handlers
    }
}
