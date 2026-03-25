using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace MonteCarlo.Addin;

/// <summary>
/// Defines the custom ribbon tab for MonteCarlo.XL in Excel.
/// </summary>
[ComVisible(true)]
public class MonteCarloRibbon : ExcelRibbon
{
    /// <summary>
    /// Returns the Ribbon XML that defines the custom tab, groups, and buttons.
    /// </summary>
    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
          <ribbon>
            <tabs>
              <tab id='MonteCarloTab' label='MonteCarlo.XL'>
                <group id='SimulationGroup' label='Simulation'>
                  <button id='RunButton' label='Run' size='large'
                          imageMso='PlayMacro' onAction='OnRunSimulation'
                          screentip='Run Simulation'
                          supertip='Start a Monte Carlo simulation with the current input/output configuration.' />
                  <button id='StopButton' label='Stop' size='large'
                          imageMso='CancelRequest' onAction='OnStopSimulation'
                          screentip='Stop Simulation'
                          supertip='Cancel the currently running simulation.' />
                </group>
                <group id='ViewGroup' label='View'>
                  <toggleButton id='TaskPaneButton' label='Task Pane' size='large'
                          imageMso='SheetShow' onAction='OnToggleTaskPane'
                          getPressed='GetTaskPanePressed'
                          screentip='Toggle Task Pane'
                          supertip='Show or hide the MonteCarlo.XL task pane.' />
                </group>
                <group id='SettingsGroup' label='Settings'>
                  <button id='SettingsButton' label='Settings' size='normal'
                          imageMso='PropertySheet' onAction='OnOpenSettings'
                          screentip='Settings'
                          supertip='Open simulation settings (iteration count, seed, theme).' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    /// <summary>
    /// Callback: Run the simulation.
    /// Shows the task pane (which triggers wiring if first time),
    /// then fires the RunSimulation command on the SetupViewModel.
    /// </summary>
    public void OnRunSimulation(IRibbonControl control)
    {
        try
        {
            // Ensure the task pane is visible (creates it and wires events if needed)
            AddIn.TaskPane?.Show();

            // Trigger the simulation via the wired-up MainViewModel
            var host = AddIn.TaskPane?.Host;
            if (host != null)
            {
                var mainVM = host.MainViewModel;
                var setupVM = mainVM.SetupViewModel;

                if (setupVM.CanRun)
                {
                    // This fires SetupViewModel.RunSimulationRequested → MainViewModel → Wiring → Orchestrator
                    setupVM.RunSimulationCommand.Execute(null);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error running simulation:\n{ex}", "MonteCarlo.XL Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Callback: Stop the running simulation.
    /// </summary>
    public void OnStopSimulation(IRibbonControl control)
    {
        AddIn.Orchestrator?.CancelSimulation();
    }

    /// <summary>
    /// Callback: Toggle the task pane visibility.
    /// </summary>
    public void OnToggleTaskPane(IRibbonControl control, bool pressed)
    {
        try
        {
            AddIn.TaskPane?.Toggle();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error toggling task pane:\n{ex}", "MonteCarlo.XL Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Returns whether the task pane toggle button is pressed.
    /// </summary>
    public bool GetTaskPanePressed(IRibbonControl control)
    {
        return AddIn.TaskPane?.IsVisible ?? false;
    }

    /// <summary>
    /// Callback: Open settings view in the task pane.
    /// </summary>
    public void OnOpenSettings(IRibbonControl control)
    {
        try
        {
            AddIn.TaskPane?.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening settings:\n{ex}", "MonteCarlo.XL Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
