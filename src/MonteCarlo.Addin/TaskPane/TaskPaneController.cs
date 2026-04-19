using System.Windows.Forms;
using System.Windows.Threading;
using ExcelDna.Integration.CustomUI;
using MonteCarlo.Addin.Services;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.Addin.TaskPane;

/// <summary>
/// Manages the Excel Custom Task Pane lifecycle — create, show, hide, dispose.
/// </summary>
public class TaskPaneController : IDisposable
{
    private CustomTaskPane? _taskPane;
    private TaskPaneHost? _host;
    private bool _disposed;

    /// <summary>
    /// Raised when the WPF task pane host has been created.
    /// </summary>
    public event EventHandler? Created;

    /// <summary>
    /// Whether the task pane is currently visible.
    /// </summary>
    public bool IsVisible => _taskPane?.Visible ?? false;

    /// <summary>
    /// Gets the root task pane view model, if the pane has been created.
    /// </summary>
    public MainViewModel? ViewModel => _host?.ViewModel;

    /// <summary>
    /// Gets the dispatcher for the hosted WPF control.
    /// </summary>
    public Dispatcher? Dispatcher => _host?.Dispatcher;

    /// <summary>
    /// Toggle task pane visibility. Creates the pane on first call.
    /// </summary>
    public void Toggle()
    {
        if (!EnsureCreated())
            return;

        _taskPane!.Visible = !_taskPane.Visible;
    }

    /// <summary>
    /// Show the task pane. Creates it if it doesn't exist.
    /// </summary>
    public void Show()
    {
        if (!EnsureCreated())
            return;

        _taskPane!.Visible = true;
    }

    /// <summary>
    /// Hide the task pane.
    /// </summary>
    public void Hide()
    {
        if (_taskPane != null)
            _taskPane.Visible = false;
    }

    private bool EnsureCreated()
    {
        if (_taskPane != null)
            return true;

        try
        {
            _host = new TaskPaneHost();
            _taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(_host, "MonteCarlo.XL");
            _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Width = 380;
            Created?.Invoke(this, EventArgs.Empty);
            return true;
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException("Failed to create task pane.", ex);
            MessageBox.Show(
                $"Failed to create task pane:\n\n{ex.GetType().Name}: {ex.Message}\n\nDiagnostics: {StartupDiagnostics.LogPath}",
                "MonteCarlo.XL Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            // Clean up partial state
            _host?.Dispose();
            _host = null;
            _taskPane = null;
            return false;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _host?.Dispose();
        _host = null;

        if (_taskPane != null)
        {
            try
            {
                _taskPane.Visible = false;
                _taskPane.Delete();
            }
            catch { }
            _taskPane = null;
        }
    }
}
