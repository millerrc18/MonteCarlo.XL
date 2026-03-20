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
    /// Whether the task pane is currently visible.
    /// </summary>
    public bool IsVisible => _taskPane?.Visible ?? false;

    /// <summary>
    /// Toggle task pane visibility. Creates the pane on first call.
    /// </summary>
    public void Toggle()
    {
        if (_taskPane == null)
            CreateTaskPane();

        _taskPane!.Visible = !_taskPane.Visible;
    }

    /// <summary>
    /// Show the task pane. Creates it if it doesn't exist.
    /// </summary>
    public void Show()
    {
        if (_taskPane == null)
            CreateTaskPane();

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

    private void CreateTaskPane()
    {
        _host = new TaskPaneHost();
        _taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(_host, "MonteCarlo.XL");
        _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
        _taskPane.Width = 380;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _host?.Dispose();
        _host = null;

        if (_taskPane != null)
        {
            _taskPane.Visible = false;
            _taskPane.Delete();
            _taskPane = null;
        }
    }
}
