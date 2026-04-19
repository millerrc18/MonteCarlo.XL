using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Threading;
using MonteCarlo.Addin.Services;
using MonteCarlo.UI.ViewModels;
using MonteCarlo.UI.Views;

namespace MonteCarlo.Addin.TaskPane;

/// <summary>
/// Explicit COM default interface required for Excel-DNA custom task panes on .NET 6+.
/// </summary>
[ComVisible(true)]
[Guid("741A7A8B-5ED1-4D67-A31F-566F7C6F2C6B")]
public interface ITaskPaneHost
{
}

/// <summary>
/// WinForms UserControl that hosts the WPF MainTaskPaneControl via ElementHost.
/// Required because ExcelDna Custom Task Panes use WinForms containers.
/// </summary>
[ComVisible(true)]
[Guid("BE37791C-4D58-4D08-9121-22E3F85CAF66")]
[ComDefaultInterface(typeof(ITaskPaneHost))]
public class TaskPaneHost : UserControl, ITaskPaneHost
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
    /// Gets the root WPF task pane view model once the pane is created.
    /// </summary>
    public MainViewModel? ViewModel => _wpfControl.DataContext as MainViewModel;

    /// <summary>
    /// Gets the dispatcher that owns the hosted WPF control.
    /// </summary>
    public Dispatcher Dispatcher => _wpfControl.Dispatcher;

    /// <summary>
    /// Creates a WPF Application if one doesn't exist and loads the
    /// theme + style resource dictionaries so StaticResource lookups work.
    /// </summary>
    private static void EnsureWpfApplication()
    {
        var app = System.Windows.Application.Current;
        if (app == null)
        {
            app = new System.Windows.Application
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown
            };
        }

        EnsureMergedDictionary(app.Resources, "pack://application:,,,/MonteCarlo.UI;component/Styles/LightTheme.xaml");
        EnsureMergedDictionary(app.Resources, "pack://application:,,,/MonteCarlo.UI;component/Styles/GlobalStyles.xaml");
    }

    private static void EnsureMergedDictionary(ResourceDictionary resources, string source)
    {
        var uri = new Uri(source, UriKind.Absolute);
        if (resources.MergedDictionaries.Any(dictionary => dictionary.Source == uri))
            return;

        try
        {
            resources.MergedDictionaries.Add(new ResourceDictionary { Source = uri });
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException($"Failed to load WPF resource dictionary: {source}", ex);
            throw;
        }
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
