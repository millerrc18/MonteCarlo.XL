using System.Windows.Forms;
using System.Windows.Forms.Integration;
using MonteCarlo.UI.Views;

namespace MonteCarlo.Addin.TaskPane;

/// <summary>
/// WinForms UserControl that hosts the WPF MainTaskPaneControl via ElementHost.
/// Required because ExcelDna Custom Task Panes use WinForms containers.
/// </summary>
public class TaskPaneHost : UserControl
{
    private readonly ElementHost _host;
    private readonly MainTaskPaneControl _wpfControl;

    public TaskPaneHost()
    {
        _wpfControl = new MainTaskPaneControl();
        _host = new ElementHost
        {
            Child = _wpfControl,
            Dock = DockStyle.Fill
        };
        Controls.Add(_host);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _host.Dispose();
        }
        base.Dispose(disposing);
    }
}
