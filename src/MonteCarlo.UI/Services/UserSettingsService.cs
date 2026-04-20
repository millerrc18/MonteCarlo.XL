using Microsoft.Win32;
using MonteCarlo.Engine.Simulation;
using System.Globalization;

namespace MonteCarlo.UI.Services;

/// <summary>
/// Persists lightweight task pane preferences for the current Windows user.
/// </summary>
public class UserSettingsService
{
    private const string RegistryKeyPath = @"Software\MonteCarlo.XL";
    private const string CreateNewWorksheetForExportsValue = "CreateNewWorksheetForExports";
    private const string DefaultIterationCountValue = "DefaultIterationCount";
    private const string SeedModeValue = "SeedMode";
    private const string FixedRandomSeedValue = "FixedRandomSeed";
    private const string SamplingMethodValue = "SamplingMethod";
    private const string AutoStopOnConvergenceValue = "AutoStopOnConvergence";
    private const string PauseOnPreflightWarningsValue = "PauseOnPreflightWarnings";
    private const string DefaultPercentilesValue = "DefaultPercentiles";

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
                    defaultValue: true),
                DefaultIterationCount = ReadInt(
                    key,
                    DefaultIterationCountValue,
                    UserSettings.Default.DefaultIterationCount,
                    minValue: 1),
                SeedMode = ReadEnum(
                    key,
                    SeedModeValue,
                    UserSettings.Default.SeedMode),
                FixedRandomSeed = ReadInt(
                    key,
                    FixedRandomSeedValue,
                    UserSettings.Default.FixedRandomSeed,
                    minValue: 0),
                SamplingMethod = ReadEnum(
                    key,
                    SamplingMethodValue,
                    UserSettings.Default.SamplingMethod),
                AutoStopOnConvergence = ReadBool(
                    key,
                    AutoStopOnConvergenceValue,
                    UserSettings.Default.AutoStopOnConvergence),
                PauseOnPreflightWarnings = ReadBool(
                    key,
                    PauseOnPreflightWarningsValue,
                    UserSettings.Default.PauseOnPreflightWarnings),
                DefaultPercentiles = ReadString(
                    key,
                    DefaultPercentilesValue,
                    UserSettings.Default.DefaultPercentiles)
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
            key.SetValue(
                DefaultIterationCountValue,
                settings.DefaultIterationCount,
                RegistryValueKind.DWord);
            key.SetValue(
                SeedModeValue,
                settings.SeedMode.ToString(),
                RegistryValueKind.String);
            key.SetValue(
                FixedRandomSeedValue,
                settings.FixedRandomSeed,
                RegistryValueKind.DWord);
            key.SetValue(
                SamplingMethodValue,
                settings.SamplingMethod.ToString(),
                RegistryValueKind.String);
            key.SetValue(
                AutoStopOnConvergenceValue,
                settings.AutoStopOnConvergence ? 1 : 0,
                RegistryValueKind.DWord);
            key.SetValue(
                PauseOnPreflightWarningsValue,
                settings.PauseOnPreflightWarnings ? 1 : 0,
                RegistryValueKind.DWord);
            key.SetValue(
                DefaultPercentilesValue,
                settings.DefaultPercentiles,
                RegistryValueKind.String);
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

    private static int ReadInt(RegistryKey? key, string valueName, int defaultValue, int minValue)
    {
        var value = key?.GetValue(valueName);
        var parsed = value switch
        {
            int intValue => intValue,
            string stringValue when int.TryParse(stringValue, out var intValue) => intValue,
            _ => defaultValue
        };

        return Math.Max(parsed, minValue);
    }

    private static T ReadEnum<T>(RegistryKey? key, string valueName, T defaultValue)
        where T : struct, Enum
    {
        var value = key?.GetValue(valueName)?.ToString();
        return Enum.TryParse<T>(value, out var parsed) ? parsed : defaultValue;
    }

    private static string ReadString(RegistryKey? key, string valueName, string defaultValue)
    {
        return key?.GetValue(valueName)?.ToString() ?? defaultValue;
    }
}

public enum SeedMode
{
    Random,
    Fixed
}

public class UserSettings
{
    public static UserSettings Default => new()
    {
        CreateNewWorksheetForExports = true,
        DefaultIterationCount = 5000,
        SeedMode = SeedMode.Random,
        FixedRandomSeed = 42,
        SamplingMethod = SamplingMethod.LatinHypercube,
        AutoStopOnConvergence = false,
        PauseOnPreflightWarnings = true,
        DefaultPercentiles = "1,5,10,25,50,75,90,95,99"
    };

    /// <summary>
    /// When true, each export creates a uniquely named worksheet instead of replacing the prior export sheet.
    /// </summary>
    public bool CreateNewWorksheetForExports { get; init; }

    /// <summary>Default iteration count for new workbook setups.</summary>
    public int DefaultIterationCount { get; init; }

    /// <summary>Whether new setups use a random seed or a fixed seed by default.</summary>
    public SeedMode SeedMode { get; init; }

    /// <summary>Default fixed seed used when <see cref="SeedMode"/> is fixed.</summary>
    public int FixedRandomSeed { get; init; }

    /// <summary>Sampling method used by simulation runs.</summary>
    public SamplingMethod SamplingMethod { get; init; }

    /// <summary>Whether simulations should request convergence auto-stop from the engine.</summary>
    public bool AutoStopOnConvergence { get; init; }

    /// <summary>Whether warnings from Model Check should pause before starting a run.</summary>
    public bool PauseOnPreflightWarnings { get; init; }

    /// <summary>Comma-separated percentile list for future report-builder defaults.</summary>
    public string DefaultPercentiles { get; init; } = string.Empty;

    /// <summary>Parsed percentile fractions in the 0.0 to 1.0 range.</summary>
    public IReadOnlyList<double> GetDefaultPercentileFractions() =>
        ParsePercentileFractions(DefaultPercentiles);

    public static IReadOnlyList<double> ParsePercentileFractions(string text)
    {
        var percentiles = new List<double>();
        foreach (var part in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (double.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                && value >= 0
                && value <= 100)
            {
                percentiles.Add(value / 100.0);
            }
        }

        return percentiles.Count == 0
            ? Default.DefaultPercentiles.Split(',').Select(p => double.Parse(p, CultureInfo.InvariantCulture) / 100.0).ToArray()
            : percentiles;
    }
}
