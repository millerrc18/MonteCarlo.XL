using System.Windows;
using MonteCarlo.Charts.Themes;
using Microsoft.Win32;

namespace MonteCarlo.UI.Services;

/// <summary>
/// Manages theme switching (Light / Dark / System) by swapping
/// theme ResourceDictionaries at runtime.
/// </summary>
public class ThemeManager
{
    /// <summary>
    /// Available themes.
    /// </summary>
    public enum Theme { Light, Dark, System }

    private const string ThemeRegistryKey = @"Software\MonteCarlo.XL";
    private const string ThemeRegistryValue = "Theme";

    private ResourceDictionary? _currentThemeDict;

    /// <summary>
    /// Gets the currently applied theme.
    /// </summary>
    public Theme CurrentTheme { get; private set; } = Theme.Light;

    /// <summary>
    /// Applies the specified theme to the application.
    /// </summary>
    /// <param name="theme">The theme to apply.</param>
    /// <param name="mergedDictionaries">
    /// The MergedDictionaries collection to update (typically from a top-level control).
    /// </param>
    public void ApplyTheme(Theme theme, System.Collections.ObjectModel.Collection<ResourceDictionary> mergedDictionaries)
    {
        var resolvedTheme = theme == Theme.System ? DetectSystemTheme() : theme;

        var uri = resolvedTheme switch
        {
            Theme.Dark => new Uri("pack://application:,,,/MonteCarlo.UI;component/Styles/DarkTheme.xaml"),
            _ => new Uri("pack://application:,,,/MonteCarlo.UI;component/Styles/LightTheme.xaml")
        };

        var newDict = new ResourceDictionary { Source = uri };

        // Remove old theme dictionary if present
        if (_currentThemeDict != null)
        {
            mergedDictionaries.Remove(_currentThemeDict);
        }

        // Add new theme dictionary (should be first so GlobalStyles can override)
        mergedDictionaries.Insert(0, newDict);
        _currentThemeDict = newDict;
        CurrentTheme = theme;

        // Sync chart theme colors
        ChartTheme.SetDarkMode(resolvedTheme == Theme.Dark);
    }

    /// <summary>
    /// Saves the theme preference to the Windows registry.
    /// </summary>
    public void SavePreference(Theme theme)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(ThemeRegistryKey);
            key.SetValue(ThemeRegistryValue, theme.ToString());
        }
        catch
        {
            // Registry write may fail in restricted environments — non-fatal
        }
    }

    /// <summary>
    /// Loads the saved theme preference from the Windows registry.
    /// </summary>
    public Theme LoadPreference()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(ThemeRegistryKey);
            var value = key?.GetValue(ThemeRegistryValue) as string;
            if (Enum.TryParse<Theme>(value, out var theme))
                return theme;
        }
        catch { }

        return Theme.System; // Default
    }

    /// <summary>
    /// Detects whether Windows is using dark mode.
    /// </summary>
    private static Theme DetectSystemTheme()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            var value = key?.GetValue("AppsUseLightTheme");
            if (value is int intValue)
                return intValue == 0 ? Theme.Dark : Theme.Light;
        }
        catch { }

        return Theme.Light; // Fallback
    }
}
