using Microsoft.Win32;

namespace MonteCarlo.UI.Services;

/// <summary>
/// Persists lightweight task pane preferences for the current Windows user.
/// </summary>
public class UserSettingsService
{
    private const string RegistryKeyPath = @"Software\MonteCarlo.XL";
    private const string CreateNewWorksheetForExportsValue = "CreateNewWorksheetForExports";

    public UserSettings Load()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
            return new UserSettings
            {
                CreateNewWorksheetForExports = ReadBool(
                    key,
                    CreateNewWorksheetForExportsValue,
                    defaultValue: true)
            };
        }
        catch
        {
            return UserSettings.Default;
        }
    }

    public void Save(UserSettings settings)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath);
            key.SetValue(
                CreateNewWorksheetForExportsValue,
                settings.CreateNewWorksheetForExports ? 1 : 0,
                RegistryValueKind.DWord);
        }
        catch
        {
            // Registry write may fail in restricted environments — non-fatal.
        }
    }

    private static bool ReadBool(RegistryKey? key, string valueName, bool defaultValue)
    {
        var value = key?.GetValue(valueName);
        return value switch
        {
            int intValue => intValue != 0,
            string stringValue when bool.TryParse(stringValue, out var boolValue) => boolValue,
            _ => defaultValue
        };
    }
}

public class UserSettings
{
    public static UserSettings Default => new()
    {
        CreateNewWorksheetForExports = true
    };

    /// <summary>
    /// When true, each export creates a uniquely named worksheet instead of replacing the prior export sheet.
    /// </summary>
    public bool CreateNewWorksheetForExports { get; init; }
}
