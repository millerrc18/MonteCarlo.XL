using System.Runtime.InteropServices;
using System.Windows.Forms;
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

    /// <summary>Task pane to Excel service bridge.</summary>
    internal static TaskPaneIntegration? Integration { get; private set; }

    /// <summary>
    /// Called when the add-in is loaded into Excel.
    /// </summary>
    public void AutoOpen()
    {
        try
        {
            InstallAssemblyResolver();
            InstallExceptionHandlers();
            InitializeSkiaSharp();
            StartupDiagnostics.LogStartup();

            TaskPane = new TaskPaneController();
            Workbook = new WorkbookManager();
            InputTags = new InputTagManager();
            OutputTags = new OutputTagManager();
            Highlighter = new CellHighlighter();
            ConfigPersistence = new ConfigPersistence();
            Orchestrator = new SimulationOrchestrator(
                Workbook, InputTags, OutputTags, ConfigPersistence);
            Integration = new TaskPaneIntegration(
                TaskPane,
                InputTags,
                OutputTags,
                Highlighter,
                Orchestrator);

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

            // Register keyboard shortcuts
            RegisterKeyboardShortcuts();
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("MonteCarlo.XL failed to initialize.", ex);
            MessageBox.Show(
                $"MonteCarlo.XL failed to initialize:\n\n{ex.GetType().Name}: {ex.Message}\n\n{ex.StackTrace}\n\nDiagnostics: {StartupDiagnostics.LogPath}",
                "MonteCarlo.XL Startup Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Registers global keyboard shortcuts in Excel via Application.OnKey().
    /// </summary>
    private static void RegisterKeyboardShortcuts()
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;

            // Ctrl+Shift+R — Run simulation
            app.OnKey("^+R", "MonteCarloRunSimulation");

            // Ctrl+Shift+S — Stop simulation
            app.OnKey("^+S", "MonteCarloStopSimulation");

            // Ctrl+Shift+T — Toggle task pane
            app.OnKey("^+T", "MonteCarloToggleTaskPane");
        }
        catch
        {
            // Keyboard shortcut registration is non-fatal
        }
    }

    /// <summary>
    /// Unregisters keyboard shortcuts on cleanup.
    /// </summary>
    private static void UnregisterKeyboardShortcuts()
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            app.OnKey("^+R");
            app.OnKey("^+S");
            app.OnKey("^+T");
        }
        catch { }
    }

    /// <summary>
    /// Bridges ExcelDna's custom AssemblyLoadContext with the default one so
    /// that WPF's BAML parser can find packed assemblies (like SkiaSharp.Views.WPF)
    /// that ExcelDna loaded into its own ALC.
    /// </summary>
    private static void InstallAssemblyResolver()
    {
        // Force-load assemblies that WPF BAML will need later.
        // These C# references go through ExcelDna's ALC which can unpack them.
        // Once loaded, they become visible to AppDomain.GetAssemblies().
        try
        {
            _ = typeof(SkiaSharp.Views.WPF.SKElement).Assembly;
            _ = typeof(LiveChartsCore.SkiaSharpView.WPF.CartesianChart).Assembly;
            _ = typeof(MonteCarlo.Charts.Controls.HistogramChart).Assembly;
            _ = typeof(MonteCarlo.UI.Views.MainTaskPaneControl).Assembly;
            StartupDiagnostics.Log("Pre-loaded WPF/SkiaSharp assemblies for BAML resolution");
        }
        catch (Exception ex)
        {
            StartupDiagnostics.Log($"Assembly pre-load warning: {ex.Message}");
        }

        // Bridge: when the default ALC or BAML parser can't find an assembly,
        // check all loaded assemblies (including ExcelDna's ALC).
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var requestedName = new System.Reflection.AssemblyName(args.Name).Name;
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.GetName().Name == requestedName)
                    return asm;
            }
            StartupDiagnostics.Log($"AssemblyResolve MISS: {args.Name}");
            return null;
        };

        System.Runtime.Loader.AssemblyLoadContext.Default.Resolving += (_, name) =>
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.GetName().Name == name.Name)
                    return asm;
            }
            return null;
        };
    }

    /// <summary>
    /// Installs global exception handlers to catch and log unhandled exceptions
    /// that would otherwise crash Excel (e.g., in WPF dispatcher callbacks).
    /// </summary>
    private static void InstallExceptionHandlers()
    {
        // Catch unhandled exceptions on the WPF dispatcher thread
        if (System.Windows.Application.Current != null)
        {
            System.Windows.Application.Current.DispatcherUnhandledException += (_, e) =>
            {
                StartupDiagnostics.LogException("WPF DispatcherUnhandledException", e.Exception);
                e.Handled = true;
            };
        }

        System.Windows.Threading.Dispatcher.CurrentDispatcher.UnhandledException += (_, e) =>
        {
            StartupDiagnostics.LogException("Dispatcher UnhandledException", e.Exception);
            e.Handled = true;
        };

        // Last resort — log before the process dies
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
                StartupDiagnostics.LogException("AppDomain UnhandledException (fatal)", ex);
        };

        // Unobserved Task exceptions
        System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            StartupDiagnostics.LogException("UnobservedTaskException", e.Exception);
            e.SetObserved();
        };
    }

    /// <summary>
    /// Preloads native SkiaSharp/HarfBuzzSharp DLLs and warms up the SkiaSharp
    /// type system so that any initialization failure surfaces here (inside
    /// AutoOpen's try-catch) rather than in an unprotected WPF binding callback.
    /// </summary>
    private static void InitializeSkiaSharp()
    {
        var xllDir = System.IO.Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? "";
        StartupDiagnostics.Log($"SkiaSharp init: xllDir={xllDir}");

        // Step 1: Preload native DLLs from the .xll directory
        foreach (var lib in new[] { "libSkiaSharp.dll", "libHarfBuzzSharp.dll" })
        {
            var path = System.IO.Path.Combine(xllDir, lib);
            var exists = System.IO.File.Exists(path);
            StartupDiagnostics.Log($"  {lib}: exists={exists}, path={path}");

            if (exists)
            {
                if (NativeLibrary.TryLoad(path, out var handle))
                    StartupDiagnostics.Log($"  {lib}: loaded at 0x{handle:X}");
                else
                    StartupDiagnostics.Log($"  {lib}: TryLoad FAILED");
            }
        }

        // Step 2: Warm up SkiaSharp — force the static constructor to run NOW
        // so any failure is caught by AutoOpen's try-catch instead of crashing
        // the process via an unhandled WPF dispatcher exception.
        try
        {
            var typeface = SkiaSharp.SKTypeface.Default;
            StartupDiagnostics.Log($"  SKTypeface.Default: {typeface?.FamilyName ?? "null"}");
        }
        catch (Exception ex)
        {
            StartupDiagnostics.Log($"  SKTypeface warmup FAILED: {ex.GetType().Name}: {ex.Message}");
            if (ex.InnerException != null)
                StartupDiagnostics.Log($"  Inner: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
        }
    }

    /// <summary>
    /// Called when the add-in is unloaded from Excel.
    /// </summary>
    public void AutoClose()
    {
        UnregisterKeyboardShortcuts();
        Integration?.Dispose();
        Integration = null;
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
