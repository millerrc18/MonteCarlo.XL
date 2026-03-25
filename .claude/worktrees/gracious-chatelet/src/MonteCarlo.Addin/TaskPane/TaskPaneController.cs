using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

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
    /// Raised once when the task pane host is first created.
    /// Used by AddIn to wire the simulation pipeline to the MainViewModel.
    /// </summary>
    public event Action<TaskPaneHost>? TaskPaneCreated;

    /// <summary>
    /// Whether the task pane is currently visible.
    /// </summary>
    public bool IsVisible => _taskPane?.Visible ?? false;

    /// <summary>
    /// Gets the TaskPaneHost, or null if the task pane hasn't been created yet.
    /// </summary>
    internal TaskPaneHost? Host => _host;

    /// <summary>
    /// Toggle task pane visibility. Creates the pane on first call.
    /// </summary>
    public void Toggle()
    {
        EnsureCreated();
        _taskPane!.Visible = !_taskPane.Visible;
    }

    /// <summary>
    /// Show the task pane. Creates it if it doesn't exist.
    /// </summary>
    public void Show()
    {
        EnsureCreated();
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

    private void EnsureCreated()
    {
        if (_taskPane != null) return;

        try
        {
            _host = new TaskPaneHost();
            _taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(_host, "MonteCarlo.XL");
            _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Width = 380;

            // Notify listeners that the pane was created (for wiring)
            TaskPaneCreated?.Invoke(_host);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to create task pane:\n\n{ex.GetType().Name}: {ex.Message}\n\n{ex.StackTrace}",
                "MonteCarlo.XL Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            // Clean up partial state
            _host?.Dispose();
            _host = null;
            _taskPane = null;
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
