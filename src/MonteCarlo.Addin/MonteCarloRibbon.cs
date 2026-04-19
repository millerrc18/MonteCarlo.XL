using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using MonteCarlo.Addin.Services;

namespace MonteCarlo.Addin;

/// <summary>
/// Defines the custom ribbon tab for MonteCarlo.XL in Excel.
/// </summary>
[ComVisible(true)]
public class MonteCarloRibbon : ExcelRibbon
{
    private IRibbonUI? _ribbon;

    /// <summary>
    /// Returns the Ribbon XML that defines the custom tab, groups, and buttons.
    /// </summary>
    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
          <ribbon>
            <tabs>
              <tab id='MonteCarloTab' label='MonteCarlo.XL'>
                <group id='WorkspaceGroup' label='Workspace'>
                  <toggleButton id='TaskPaneButton' label='Task Pane' size='large'
                          imageMso='FileOpen' onAction='OnToggleTaskPane'
                          getPressed='GetTaskPanePressed'
                          screentip='Show Task Pane'
                          supertip='Show or hide the MonteCarlo.XL task pane.' />
                  <button id='OpenSetupButton' label='Setup'
                          imageMso='PropertySheet'
                          onAction='OnOpenSetup'
                          screentip='Open Setup'
                          supertip='Open the setup view for inputs, outputs, iteration count, and seed.' />
                  <button id='OpenResultsButton' label='Results'
                          imageMso='ChartInsert'
                          onAction='OnOpenResults'
                          screentip='Open Results'
                          supertip='Open the results dashboard for the latest simulation.' />
                  <button id='SettingsButton' label='Settings'
                          imageMso='PropertySheet' onAction='OnOpenSettings'
                          screentip='Open Settings'
                          supertip='Open add-in settings.' />
                </group>
                <group id='ModelGroup' label='Model Setup'>
                  <button id='AddInputButton' label='Add Input' size='large'
                          imageMso='NameManager' onAction='OnBeginAddInput'
                          screentip='Add Input Assumption'
                          supertip='Open the input editor so you can select a worksheet cell and assign a probability distribution.' />
                  <button id='AddOutputButton' label='Add Output' size='large'
                          imageMso='TraceDependents' onAction='OnBeginAddOutput'
                          screentip='Add Output Forecast'
                          supertip='Open the output editor so you can select the result cell to capture during simulation.' />
                  <button id='CorrelationsButton' label='Correlations'
                          imageMso='TraceDependents' onAction='OnOpenCorrelations'
                          screentip='Define Correlations'
                          supertip='Open the input correlation matrix editor for rank-correlated simulation inputs.' />
                </group>
                <group id='SimulationGroup' label='Simulation'>
                  <button id='ModelCheckButton' label='Model Check'
                          imageMso='ReviewAcceptChange' onAction='OnModelCheck'
                          screentip='Check Model'
                          supertip='Validate the current simulation setup before running.' />
                  <button id='RunButton' label='Run' size='large'
                          imageMso='PlayMacro' onAction='OnRunSimulation'
                          screentip='Run Simulation'
                          supertip='Start a Monte Carlo simulation with the current input/output configuration.' />
                  <button id='StopButton' label='Stop' size='large'
                          imageMso='CancelRequest' onAction='OnStopSimulation'
                          screentip='Stop Simulation'
                          supertip='Cancel the currently running simulation.' />
                  <menu id='IterationPresetMenu' label='Iterations'
                        imageMso='CalculateNow'
                        screentip='Iteration Presets'
                        supertip='Set a common simulation iteration count.'>
                    <button id='Iterations1000Button' label='Preview - 1,000' imageMso='CalculateNow' onAction='OnSetIterations1000' />
                    <button id='Iterations5000Button' label='Standard - 5,000' imageMso='CalculateNow' onAction='OnSetIterations5000' />
                    <button id='Iterations10000Button' label='Standard+ - 10,000' imageMso='CalculateNow' onAction='OnSetIterations10000' />
                    <button id='Iterations25000Button' label='Full - 25,000' imageMso='CalculateNow' onAction='OnSetIterations25000' />
                    <button id='Iterations50000Button' label='Deep - 50,000' imageMso='CalculateNow' onAction='OnSetIterations50000' />
                  </menu>
                </group>
                <group id='ResultsGroup' label='Results'>
                  <button id='ExportSummaryButton' label='Export Summary'
                          imageMso='FileSaveAsExcelXlsx'
                          onAction='OnExportSummary'
                          screentip='Export Summary'
                          supertip='Export the selected output statistics and input assumptions to a worksheet.' />
                  <button id='ExportRawDataButton' label='Export Raw Data'
                          imageMso='FileSaveAsExcelXlsx'
                          onAction='OnExportRawData'
                          screentip='Export Raw Data'
                          supertip='Export iteration-level input samples and output values to a worksheet.' />
                  <button id='CopyStatsButton' label='Copy Stats'
                          imageMso='Copy' onAction='OnCopyStats'
                          screentip='Copy Statistics'
                          supertip='Copy the currently selected output statistics from the results view to the clipboard.' />
                  <button id='OpenResultsButton2' label='View Results'
                          imageMso='ChartInsert'
                          onAction='OnOpenResults'
                          screentip='View Results'
                          supertip='Open the results dashboard.' />
                </group>
                <group id='SharingGroup' label='Sharing'>
                  <button id='ReplaceMcFormulasButton' label='Replace MC Formulas'
                          imageMso='PasteValues'
                          onAction='OnReplaceMcFormulas'
                          screentip='Replace MC Formulas'
                          supertip='Replace MC.* formulas in the active workbook with their current values and save a restore map in the workbook.' />
                  <button id='RestoreMcFormulasButton' label='Restore MC Formulas'
                          imageMso='Undo'
                          onAction='OnRestoreMcFormulas'
                          screentip='Restore MC Formulas'
                          supertip='Restore MC.* formulas from the workbook restore map.' />
                </group>
                <group id='SupportGroup' label='Support'>
                  <button id='OpenLogButton' label='Startup Log'
                          imageMso='FileFind' onAction='OnOpenDiagnosticsLog'
                          screentip='Open Startup Log'
                          supertip='Open the MonteCarlo.XL diagnostics log used for Excel loading and runtime errors.' />
                  <button id='OpenLogFolderButton' label='Log Folder'
                          imageMso='FileOpen'
                          onAction='OnOpenDiagnosticsFolder'
                          screentip='Open Log Folder'
                          supertip='Open the local MonteCarlo.XL diagnostics folder.' />
                  <button id='AboutButton' label='About'
                          imageMso='Info' onAction='OnAbout'
                          screentip='About MonteCarlo.XL'
                          supertip='Show version, keyboard shortcuts, and diagnostic path.' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    /// <summary>
    /// Callback: capture the ribbon object for future invalidation.
    /// </summary>
    public void OnLoad(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
    }

    /// <summary>
    /// Callback: Run the simulation.
    /// </summary>
    public void OnRunSimulation(IRibbonControl control)
    {
        RunRibbonAction("Run simulation", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.RequestRunFromRibbon();
        });
    }

    /// <summary>
    /// Callback: Stop the running simulation.
    /// </summary>
    public void OnStopSimulation(IRibbonControl control)
    {
        RunRibbonAction("Stop simulation", () => AddIn.Orchestrator?.CancelSimulation());
    }

    /// <summary>
    /// Callback: Validate the current simulation setup.
    /// </summary>
    public void OnModelCheck(IRibbonControl control)
    {
        RunRibbonAction("Model check", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.ShowPreflight();
        });
    }

    /// <summary>
    /// Callback: Toggle the task pane visibility.
    /// </summary>
    public void OnToggleTaskPane(IRibbonControl control, bool pressed)
    {
        RunRibbonAction("Toggle task pane", () =>
        {
            if (pressed)
                ShowTaskPaneAndWire();
            else
                AddIn.TaskPane?.Hide();

            _ribbon?.InvalidateControl("TaskPaneButton");
        });
    }

    /// <summary>
    /// Returns whether the task pane toggle button is pressed.
    /// </summary>
    public bool GetTaskPanePressed(IRibbonControl control)
    {
        return AddIn.TaskPane?.IsVisible ?? false;
    }

    /// <summary>
    /// Callback: Open the setup view in the task pane.
    /// </summary>
    public void OnOpenSetup(IRibbonControl control)
    {
        RunRibbonAction("Open setup", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.ShowSetup();
        });
    }

    /// <summary>
    /// Callback: Open the results view in the task pane.
    /// </summary>
    public void OnOpenResults(IRibbonControl control)
    {
        RunRibbonAction("Open results", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.ShowResults();
        });
    }

    /// <summary>
    /// Callback: Open settings view in the task pane.
    /// </summary>
    public void OnOpenSettings(IRibbonControl control)
    {
        RunRibbonAction("Open settings", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.ShowSettings();
        });
    }

    /// <summary>
    /// Callback: Start the add-input task pane flow.
    /// </summary>
    public void OnBeginAddInput(IRibbonControl control)
    {
        RunRibbonAction("Add input", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.BeginAddInput();
        });
    }

    /// <summary>
    /// Callback: Start the add-output task pane flow.
    /// </summary>
    public void OnBeginAddOutput(IRibbonControl control)
    {
        RunRibbonAction("Add output", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.BeginAddOutput();
        });
    }

    /// <summary>
    /// Callback: Open the input correlation matrix editor.
    /// </summary>
    public void OnOpenCorrelations(IRibbonControl control)
    {
        RunRibbonAction("Open correlations", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.ShowCorrelations();
        });
    }

    public void OnSetIterations1000(IRibbonControl control) => SetIterations(1000);

    public void OnSetIterations5000(IRibbonControl control) => SetIterations(5000);

    public void OnSetIterations10000(IRibbonControl control) => SetIterations(10000);

    public void OnSetIterations25000(IRibbonControl control) => SetIterations(25000);

    public void OnSetIterations50000(IRibbonControl control) => SetIterations(50000);

    /// <summary>
    /// Callback: Copy current result statistics to the clipboard.
    /// </summary>
    public void OnCopyStats(IRibbonControl control)
    {
        RunRibbonAction("Copy statistics", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.CopyStatsToClipboard();
        });
    }

    /// <summary>
    /// Callback: Export current result summary to a worksheet.
    /// </summary>
    public void OnExportSummary(IRibbonControl control)
    {
        RunRibbonAction("Export summary", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.ExportCurrentSummary();
        });
    }

    /// <summary>
    /// Callback: Export current raw simulation data to a worksheet.
    /// </summary>
    public void OnExportRawData(IRibbonControl control)
    {
        RunRibbonAction("Export raw data", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.ExportCurrentRawData();
        });
    }

    /// <summary>
    /// Callback: Replace MC.* formulas with current values for workbook sharing.
    /// </summary>
    public void OnReplaceMcFormulas(IRibbonControl control)
    {
        RunRibbonAction("Replace MC formulas", () =>
        {
            var confirm = MessageBox.Show(
                "Replace all MC.* formulas in the active workbook with their current values?\r\n\r\n" +
                "A restore map will be saved inside the workbook so you can restore the formulas later.",
                "Replace MC Formulas",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
                return;

            var result = new FunctionSwapService().ReplaceMcFormulasWithCurrentValues();
            MessageBox.Show(
                result.Message,
                "Replace MC Formulas",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        });
    }

    /// <summary>
    /// Callback: Restore MC.* formulas from the workbook restore map.
    /// </summary>
    public void OnRestoreMcFormulas(IRibbonControl control)
    {
        RunRibbonAction("Restore MC formulas", () =>
        {
            var confirm = MessageBox.Show(
                "Restore MC.* formulas from the workbook restore map?\r\n\r\n" +
                "This will overwrite the current values in those cells.",
                "Restore MC Formulas",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes)
                return;

            var result = new FunctionSwapService().RestoreMcFormulas();
            MessageBox.Show(
                result.Message,
                "Restore MC Formulas",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        });
    }

    /// <summary>
    /// Callback: Open the diagnostics log.
    /// </summary>
    public void OnOpenDiagnosticsLog(IRibbonControl control)
    {
        RunRibbonAction("Open diagnostics log", () =>
        {
            Directory.CreateDirectory(StartupDiagnostics.LogDirectory);
            if (!File.Exists(StartupDiagnostics.LogPath))
                File.WriteAllText(StartupDiagnostics.LogPath, "MonteCarlo.XL diagnostics log has not recorded entries yet.\r\n");

            Process.Start(new ProcessStartInfo
            {
                FileName = StartupDiagnostics.LogPath,
                UseShellExecute = true
            });
        });
    }

    /// <summary>
    /// Callback: Open the diagnostics folder.
    /// </summary>
    public void OnOpenDiagnosticsFolder(IRibbonControl control)
    {
        RunRibbonAction("Open diagnostics folder", () =>
        {
            Directory.CreateDirectory(StartupDiagnostics.LogDirectory);
            Process.Start(new ProcessStartInfo
            {
                FileName = StartupDiagnostics.LogDirectory,
                UseShellExecute = true
            });
        });
    }

    /// <summary>
    /// Callback: Show add-in metadata and shortcuts.
    /// </summary>
    public void OnAbout(IRibbonControl control)
    {
        RunRibbonAction("Show about", () =>
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown";
            MessageBox.Show(
                $"MonteCarlo.XL\r\nVersion: {version}\r\nExcel-DNA add-in\r\n\r\n" +
                "Keyboard shortcuts:\r\n" +
                "Ctrl+Shift+R  Run simulation\r\n" +
                "Ctrl+Shift+S  Stop simulation\r\n" +
                "Ctrl+Shift+T  Toggle task pane\r\n\r\n" +
                $"Diagnostics:\r\n{StartupDiagnostics.LogPath}",
                "About MonteCarlo.XL",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        });
    }

    private static void SetIterations(int iterationCount)
    {
        RunRibbonAction($"Set iterations to {iterationCount:N0}", () =>
        {
            ShowTaskPaneAndWire();
            AddIn.Integration?.SetIterationCount(iterationCount);
        });
    }

    private static void ShowTaskPaneAndWire()
    {
        if (AddIn.TaskPane == null)
            throw new InvalidOperationException("MonteCarlo.XL has not finished initializing the task pane service.");

        AddIn.TaskPane.Show();
        AddIn.Integration?.EnsureWired();
    }

    private static void RunRibbonAction(string actionName, Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException($"{actionName} failed.", ex);
            MessageBox.Show(
                $"{actionName} failed:\r\n\r\n{ex.GetType().Name}: {ex.Message}\r\n\r\nDiagnostics: {StartupDiagnostics.LogPath}",
                "MonteCarlo.XL Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
