using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Media3D;
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
    private const string LightThemePath = "/MonteCarlo.UI;component/Styles/LightTheme.xaml";
    private const string DarkThemePath = "/MonteCarlo.UI;component/Styles/DarkTheme.xaml";

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

        _currentThemeDict = ReplaceThemeDictionary(mergedDictionaries, resolvedTheme);
        CurrentTheme = theme;

        // Sync chart theme colors
        ChartTheme.SetDarkMode(resolvedTheme == Theme.Dark);
    }

    /// <summary>
    /// Applies the specified theme to a root control and every loaded child control.
    /// </summary>
    public void ApplyTheme(Theme theme, FrameworkElement root)
    {
        var resolvedTheme = theme == Theme.System ? DetectSystemTheme() : theme;
        ApplyTheme(theme, root.Resources.MergedDictionaries);
        ApplyThemeToDescendants(root, resolvedTheme);
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

    private static void ApplyThemeToDescendants(DependencyObject root, Theme resolvedTheme)
    {
        var visited = new HashSet<DependencyObject>();
        ApplyThemeToDescendants(root, resolvedTheme, visited);
    }

    private static void ApplyThemeToDescendants(
        DependencyObject root,
        Theme resolvedTheme,
        HashSet<DependencyObject> visited)
    {
        if (!visited.Add(root))
            return;

        if (root is FrameworkElement element)
            ReplaceThemeDictionary(element.Resources.MergedDictionaries, resolvedTheme);

        foreach (var child in LogicalTreeHelper.GetChildren(root).OfType<DependencyObject>())
            ApplyThemeToDescendants(child, resolvedTheme, visited);

        var visualChildCount = root is Visual or Visual3D ? VisualTreeHelper.GetChildrenCount(root) : 0;
        for (var i = 0; i < visualChildCount; i++)
            ApplyThemeToDescendants(VisualTreeHelper.GetChild(root, i), resolvedTheme, visited);
    }

    private static ResourceDictionary ReplaceThemeDictionary(
        System.Collections.ObjectModel.Collection<ResourceDictionary> mergedDictionaries,
        Theme resolvedTheme)
    {
        for (var i = mergedDictionaries.Count - 1; i >= 0; i--)
        {
            if (IsThemeDictionary(mergedDictionaries[i]))
                mergedDictionaries.RemoveAt(i);
        }

        var newDict = new ResourceDictionary { Source = GetThemeUri(resolvedTheme) };
        mergedDictionaries.Insert(0, newDict);
        return newDict;
    }

    private static bool IsThemeDictionary(ResourceDictionary dictionary)
    {
        var source = dictionary.Source?.OriginalString;
        return source != null
            && (source.Contains(LightThemePath, StringComparison.OrdinalIgnoreCase)
                || source.Contains(DarkThemePath, StringComparison.OrdinalIgnoreCase));
    }

    private static Uri GetThemeUri(Theme resolvedTheme)
    {
        return resolvedTheme == Theme.Dark
            ? new Uri($"pack://application:,,,{DarkThemePath}")
            : new Uri($"pack://application:,,,{LightThemePath}");
    }
}
