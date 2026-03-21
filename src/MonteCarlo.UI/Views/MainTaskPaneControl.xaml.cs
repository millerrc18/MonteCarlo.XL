using System.Windows;
using System.Windows.Controls;
using MonteCarlo.UI.Services;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Root task pane control that hosts view navigation and the content area.
/// Initializes theme on load and serves as the theme dictionary host.
/// </summary>
public partial class MainTaskPaneControl : UserControl
{
    private readonly ThemeManager _themeManager = new();

    public MainTaskPaneControl()
    {
        // Merge GlobalStyles first
        Resources.MergedDictionaries.Add(new ResourceDictionary
        {
            Source = new Uri("pack://application:,,,/MonteCarlo.UI;component/Styles/GlobalStyles.xaml")
        });

        // Apply saved theme (inserts theme dict at position 0)
        var savedTheme = _themeManager.LoadPreference();
        _themeManager.ApplyTheme(savedTheme, Resources.MergedDictionaries);

        InitializeComponent();
    }
}
