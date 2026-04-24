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
    private enum SettingsScope
    {
        Global,
        Workbook
    }

    private readonly ThemeManager _themeManager = new();
    private readonly UserSettingsService _settingsService = new();
    private UserSettings _globalSettings = UserSettings.Default;
    private WorkbookUserSettingsOverrides? _workbookOverrides;
    private SettingsScope _currentScope = SettingsScope.Global;
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
        _globalSettings = _settingsService.Load();
        _workbookOverrides = WorkbookSettingsBridge.IsAvailable
            ? WorkbookSettingsBridge.Load()
            : null;
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

        WorkbookScopeRadio.IsEnabled = WorkbookSettingsBridge.IsAvailable;
        LoadScopeIntoControls(
            WorkbookSettingsBridge.IsAvailable && _workbookOverrides?.HasAnyValues == true
                ? SettingsScope.Workbook
                : SettingsScope.Global);
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

    private void OnSettingsScopeChanged(object sender, RoutedEventArgs e)
    {
        if (_isInitializing) return;

        LoadScopeIntoControls(WorkbookScopeRadio.IsChecked == true
            ? SettingsScope.Workbook
            : SettingsScope.Global);
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

        if (_currentScope == SettingsScope.Global)
        {
            _settingsService.Save(settings);
            _globalSettings = settings;
            SettingsSavedText.Text = "Defaults saved.";
        }
        else
        {
            if (!WorkbookSettingsBridge.IsAvailable)
            {
                SettingsSavedText.Foreground = FindResource("Red500Brush") as System.Windows.Media.Brush;
                SettingsSavedText.Text = "Workbook overrides are only available inside Excel.";
                return;
            }

            _workbookOverrides = UserSettings.BuildOverrides(_globalSettings, settings);
            WorkbookSettingsBridge.Save(_workbookOverrides);
            SettingsSavedText.Text = _workbookOverrides == null
                ? "Workbook overrides cleared. Global defaults apply."
                : "Workbook overrides saved.";
        }

        SettingsSavedText.Foreground = FindResource("Emerald500Brush") as System.Windows.Media.Brush;
        UpdateScopePresentation();
    }

    private void OnClearWorkbookOverridesClicked(object sender, RoutedEventArgs e)
    {
        if (!WorkbookSettingsBridge.IsAvailable)
            return;

        _workbookOverrides = null;
        WorkbookSettingsBridge.Save(null);
        SettingsSavedText.Foreground = FindResource("Emerald500Brush") as System.Windows.Media.Brush;
        SettingsSavedText.Text = "Workbook overrides cleared.";
        LoadScopeIntoControls(SettingsScope.Workbook);
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

    private void LoadScopeIntoControls(SettingsScope scope)
    {
        _currentScope = scope;

        var settings = scope == SettingsScope.Global
            ? _globalSettings
            : UserSettings.ApplyOverrides(_globalSettings, _workbookOverrides);

        _isInitializing = true;
        GlobalScopeRadio.IsChecked = scope == SettingsScope.Global;
        WorkbookScopeRadio.IsChecked = scope == SettingsScope.Workbook;
        LoadSettingsIntoControls(settings);
        UpdateScopePresentation();
        _isInitializing = false;
    }

    private void LoadSettingsIntoControls(UserSettings settings)
    {
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
    }

    private void UpdateScopePresentation()
    {
        var hasWorkbookOverrides = _workbookOverrides?.HasAnyValues == true;

        if (!WorkbookSettingsBridge.IsAvailable)
        {
            SettingsScopeSummaryText.Text = "Workbook overrides are available only when the task pane is hosted inside Excel.";
            SaveSettingsButton.Content = "Save Defaults";
            ClearWorkbookOverridesButton.IsEnabled = false;
            return;
        }

        if (_currentScope == SettingsScope.Global)
        {
            SettingsScopeSummaryText.Text = "These values are stored for the current Windows user and become the fallback defaults for all workbooks.";
            SaveSettingsButton.Content = "Save Defaults";
            ClearWorkbookOverridesButton.IsEnabled = false;
            return;
        }

        SettingsScopeSummaryText.Text = hasWorkbookOverrides
            ? "These values are stored inside the active workbook and override the Windows defaults only for this workbook."
            : "No workbook override is saved yet. Saving here stores only the values that differ from the Windows defaults.";
        SaveSettingsButton.Content = hasWorkbookOverrides ? "Save Workbook Overrides" : "Save Workbook Settings";
        ClearWorkbookOverridesButton.IsEnabled = hasWorkbookOverrides;
    }
}
