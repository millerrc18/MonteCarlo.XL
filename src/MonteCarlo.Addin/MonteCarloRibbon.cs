using ExcelDna.Integration.CustomUI;

namespace MonteCarlo.Addin;

/// <summary>
/// Defines the custom ribbon tab for MonteCarlo.XL in Excel.
/// </summary>
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
                          imageMso='PlayMacro' onAction='OnRunSimulation' />
                  <button id='StopButton' label='Stop' size='large'
                          imageMso='CancelRequest' onAction='OnStopSimulation' />
                </group>
                <group id='ViewGroup' label='View'>
                  <button id='TaskPaneButton' label='Task Pane' size='large'
                          imageMso='SheetInsert' onAction='OnToggleTaskPane' />
                  <button id='SettingsButton' label='Settings' size='normal'
                          imageMso='PropertySheet' onAction='OnOpenSettings' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    /// <summary>
    /// Callback: Run the simulation.
    /// </summary>
    public void OnRunSimulation(IRibbonControl control)
    {
        // TODO: Trigger simulation engine
    }

    /// <summary>
    /// Callback: Stop the running simulation.
    /// </summary>
    public void OnStopSimulation(IRibbonControl control)
    {
        // TODO: Cancel simulation
    }

    /// <summary>
    /// Callback: Toggle the task pane visibility.
    /// </summary>
    public void OnToggleTaskPane(IRibbonControl control)
    {
        // TODO: Show/hide WPF task pane
    }

    /// <summary>
    /// Callback: Open settings view.
    /// </summary>
    public void OnOpenSettings(IRibbonControl control)
    {
        // TODO: Navigate to settings
    }
}
