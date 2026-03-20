using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.Services;
using MonteCarlo.Addin.TaskPane;

namespace MonteCarlo.Addin;

/// <summary>
/// ExcelDna add-in entry point. Manages lifecycle and shared services.
/// </summary>
public class AddIn : IExcelAddIn
{
    /// <summary>Shared task pane controller — accessed by the ribbon callbacks.</summary>
    internal static TaskPaneController? TaskPane { get; private set; }

    /// <summary>Shared workbook manager for Excel I/O.</summary>
    internal static WorkbookManager? Workbook { get; private set; }

    /// <summary>Input tag manager.</summary>
    internal static InputTagManager? InputTags { get; private set; }

    /// <summary>Output tag manager.</summary>
    internal static OutputTagManager? OutputTags { get; private set; }

    /// <summary>Cell highlighter.</summary>
    internal static CellHighlighter? Highlighter { get; private set; }

    /// <summary>Config persistence service.</summary>
    internal static ConfigPersistence? ConfigPersistence { get; private set; }

    /// <summary>Simulation orchestrator — central coordinator.</summary>
    internal static SimulationOrchestrator? Orchestrator { get; private set; }

    /// <summary>
    /// Called when the add-in is loaded into Excel.
    /// </summary>
    public void AutoOpen()
    {
        TaskPane = new TaskPaneController();
        Workbook = new WorkbookManager();
        InputTags = new InputTagManager();
        OutputTags = new OutputTagManager();
        Highlighter = new CellHighlighter();
        ConfigPersistence = new ConfigPersistence();
        Orchestrator = new SimulationOrchestrator(
            Workbook, InputTags, OutputTags, ConfigPersistence);

        // Auto-load saved config if present
        try
        {
            var profile = ConfigPersistence.Load();
            if (profile != null)
            {
                // Profile loaded — SetupViewModel will be populated when task pane opens
            }
        }
        catch
        {
            // Config load failure is not fatal
        }
    }

    /// <summary>
    /// Called when the add-in is unloaded from Excel.
    /// </summary>
    public void AutoClose()
    {
        TaskPane?.Dispose();
        TaskPane = null;
        Workbook = null;
        InputTags = null;
        OutputTags = null;
        Highlighter = null;
        ConfigPersistence = null;
        Orchestrator = null;
    }
}
