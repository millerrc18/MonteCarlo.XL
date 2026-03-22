using System.Windows;
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
        // Ensure WPF Application exists and resource dictionaries are loaded.
        // When hosted inside Excel, there is no App.xaml — we must create
        // the Application object and merge resources manually.
        EnsureWpfApplication();

        _wpfControl = new MainTaskPaneControl();
        _host = new ElementHost
        {
            Child = _wpfControl,
            Dock = DockStyle.Fill
        };
        Controls.Add(_host);
    }

    /// <summary>
    /// Creates a WPF Application if one doesn't exist and loads the
    /// theme + style resource dictionaries so StaticResource lookups work.
    /// </summary>
    private static void EnsureWpfApplication()
    {
        if (System.Windows.Application.Current != null) return;

        var app = new System.Windows.Application
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown
        };

        // Load LightTheme first (provides BackgroundBrush, SurfaceBrush, etc.)
        app.Resources.MergedDictionaries.Add(new ResourceDictionary
        {
            Source = new Uri("pack://application:,,,/MonteCarlo.UI;component/Styles/LightTheme.xaml")
        });

        // Then GlobalStyles (provides HeadlineSmall, GhostButton, TabBarButton, etc.)
        app.Resources.MergedDictionaries.Add(new ResourceDictionary
        {
            Source = new Uri("pack://application:,,,/MonteCarlo.UI;component/Styles/GlobalStyles.xaml")
        });
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
