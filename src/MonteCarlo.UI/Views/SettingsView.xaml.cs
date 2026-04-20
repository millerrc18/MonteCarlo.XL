using System.Windows;
using System.Windows.Controls;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.Services;
using MonteCarlo.UI.ViewModels;

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
        DefaultRunPresetComboBox.ItemsSource = RunPresetOption.Defaults;
        SamplingMethodComboBox.ItemsSource = Enum.GetValues<SamplingMethod>();
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
        DefaultRunPresetComboBox.SelectedItem = RunPresetOption.FindByIterations(settings.DefaultIterationCount);
        DefaultIterationsTextBox.Text = settings.DefaultIterationCount.ToString();
        RandomSeedRadio.IsChecked = settings.SeedMode == SeedMode.Random;
        FixedSeedRadio.IsChecked = settings.SeedMode == SeedMode.Fixed;
        FixedSeedTextBox.Text = settings.FixedRandomSeed.ToString();
        SamplingMethodComboBox.SelectedItem = settings.SamplingMethod;
        AutoStopOnConvergenceCheckBox.IsChecked = settings.AutoStopOnConvergence;
        PauseOnPreflightWarningsCheckBox.IsChecked = settings.PauseOnPreflightWarnings;
        DefaultPercentilesTextBox.Text = settings.DefaultPercentiles;
        SettingsSavedText.Text = string.Empty;

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

        if (TryReadSettingsFromControls(out var settings, out _))
            _settingsService.Save(settings);
    }

    private void OnDefaultRunPresetChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializing) return;

        if (DefaultRunPresetComboBox.SelectedItem is RunPresetOption preset)
            DefaultIterationsTextBox.Text = preset.Iterations.ToString();
    }

    private void OnDefaultIterationsChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializing) return;

        if (int.TryParse(DefaultIterationsTextBox.Text, out var iterations))
            DefaultRunPresetComboBox.SelectedItem = RunPresetOption.FindByIterations(iterations);
        else
            DefaultRunPresetComboBox.SelectedItem = null;
    }

    private void OnSaveDefaultsClicked(object sender, RoutedEventArgs e)
    {
        if (!TryReadSettingsFromControls(out var settings, out var error))
        {
            SettingsSavedText.Foreground = FindResource("Red500Brush") as System.Windows.Media.Brush;
            SettingsSavedText.Text = error;
            return;
        }

        _settingsService.Save(settings);
        SettingsSavedText.Foreground = FindResource("Emerald500Brush") as System.Windows.Media.Brush;
        SettingsSavedText.Text = "Defaults saved.";
    }

    private bool TryReadSettingsFromControls(out UserSettings settings, out string error)
    {
        settings = UserSettings.Default;
        error = string.Empty;

        if (!int.TryParse(DefaultIterationsTextBox.Text, out var defaultIterations) || defaultIterations <= 0)
        {
            error = "Iterations must be a positive whole number.";
            return false;
        }

        if (!int.TryParse(FixedSeedTextBox.Text, out var fixedSeed) || fixedSeed < 0)
        {
            error = "Fixed seed must be zero or greater.";
            return false;
        }

        var percentiles = DefaultPercentilesTextBox.Text.Trim();
        if (!IsValidPercentileList(percentiles))
        {
            error = "Percentiles must be comma-separated values from 0 to 100.";
            return false;
        }

        settings = new UserSettings
        {
            CreateNewWorksheetForExports = CreateNewExportSheetCheckBox.IsChecked == true,
            DefaultIterationCount = defaultIterations,
            SeedMode = FixedSeedRadio.IsChecked == true ? SeedMode.Fixed : SeedMode.Random,
            FixedRandomSeed = fixedSeed,
            SamplingMethod = SamplingMethodComboBox.SelectedItem is SamplingMethod method
                ? method
                : UserSettings.Default.SamplingMethod,
            AutoStopOnConvergence = AutoStopOnConvergenceCheckBox.IsChecked == true,
            PauseOnPreflightWarnings = PauseOnPreflightWarningsCheckBox.IsChecked == true,
            DefaultPercentiles = percentiles
        };

        return true;
    }

    private static bool IsValidPercentileList(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        foreach (var part in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!double.TryParse(
                    part,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var percentile)
                || percentile < 0
                || percentile > 100)
            {
                return false;
            }
        }

        return true;
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
