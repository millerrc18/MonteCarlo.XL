using System.Windows;
using System.Windows.Controls;
using MonteCarlo.UI.Services;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Settings view — theme configuration and about information.
/// </summary>
public partial class SettingsView : UserControl
{
    private readonly ThemeManager _themeManager = new();
    private readonly UserSettingsService _settingsService = new();
    private bool _isInitializing = true;

    public SettingsView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        var currentTheme = _themeManager.LoadPreference();
        var settings = _settingsService.Load();
        _isInitializing = true;

        switch (currentTheme)
        {
            case ThemeManager.Theme.Light:
                LightRadio.IsChecked = true;
                break;
            case ThemeManager.Theme.Dark:
                DarkRadio.IsChecked = true;
                break;
            case ThemeManager.Theme.System:
            default:
                SystemRadio.IsChecked = true;
                break;
        }

        CreateNewExportSheetCheckBox.IsChecked = settings.CreateNewWorksheetForExports;

        _isInitializing = false;
    }

    private void OnThemeChanged(object sender, RoutedEventArgs e)
    {
        if (_isInitializing) return;

        ThemeManager.Theme selectedTheme;
        if (LightRadio.IsChecked == true)
            selectedTheme = ThemeManager.Theme.Light;
        else if (DarkRadio.IsChecked == true)
            selectedTheme = ThemeManager.Theme.Dark;
        else
            selectedTheme = ThemeManager.Theme.System;

        // Find the top-level control's merged dictionaries
        var topControl = FindTopLevelControl();
        if (topControl != null)
        {
            _themeManager.ApplyTheme(selectedTheme, topControl);
        }

        _themeManager.SavePreference(selectedTheme);
    }

    private void OnExportSettingsChanged(object sender, RoutedEventArgs e)
    {
        if (_isInitializing) return;

        _settingsService.Save(new UserSettings
        {
            CreateNewWorksheetForExports = CreateNewExportSheetCheckBox.IsChecked == true
        });
    }

    private FrameworkElement? FindTopLevelControl()
    {
        DependencyObject current = this;
        FrameworkElement? last = null;

        while (current != null)
        {
            if (current is FrameworkElement fe)
                last = fe;
            current = LogicalTreeHelper.GetParent(current)
                      ?? System.Windows.Media.VisualTreeHelper.GetParent(current);
        }

        return last;
    }
}
