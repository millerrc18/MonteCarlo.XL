using ExcelDna.Integration;

namespace MonteCarlo.Addin;

/// <summary>
/// ExcelDna-registered macros that serve as targets for keyboard shortcuts
/// registered via Application.OnKey(). Each method name matches the macro
/// name passed to OnKey() in AddIn.RegisterKeyboardShortcuts().
/// </summary>
public static class KeyboardShortcuts
{
    /// <summary>
    /// Ctrl+Shift+R — Run simulation.
    /// </summary>
    [ExcelCommand(Name = "MonteCarloRunSimulation")]
    public static void MonteCarloRunSimulation()
    {
        AddIn.TaskPane?.Show();

        // Trigger the simulation via the wired MainViewModel
        var host = AddIn.TaskPane?.Host;
        if (host != null)
        {
            var setupVM = host.MainViewModel.SetupViewModel;
            if (setupVM.CanRun)
                setupVM.RunSimulationCommand.Execute(null);
        }
    }

    /// <summary>
    /// Ctrl+Shift+S — Stop simulation.
    /// </summary>
    [ExcelCommand(Name = "MonteCarloStopSimulation")]
    public static void MonteCarloStopSimulation()
    {
        AddIn.Orchestrator?.CancelSimulation();
    }

    /// <summary>
    /// Ctrl+Shift+T — Toggle task pane visibility.
    /// </summary>
    [ExcelCommand(Name = "MonteCarloToggleTaskPane")]
    public static void MonteCarloToggleTaskPane()
    {
        AddIn.TaskPane?.Toggle();
    }
}
