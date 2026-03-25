using System.Globalization;

namespace MonteCarlo.UI.Converters;

/// <summary>
/// Smart number formatting utility that handles values across orders of magnitude.
/// Produces clean, readable output for financial, percentage, and raw numeric values.
/// </summary>
public static class NumberFormatter
{
    /// <summary>
    /// Format a numeric value with intelligent magnitude detection.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <param name="hint">Optional hint: "currency", "percent", or null for auto.</param>
    public static string Format(double value, string? hint = null)
    {
        if (double.IsNaN(value)) return "NaN";
        if (double.IsPositiveInfinity(value)) return "+Inf";
        if (double.IsNegativeInfinity(value)) return "-Inf";

        if (hint == "percent")
            return value.ToString("P1", CultureInfo.InvariantCulture);

        var abs = Math.Abs(value);

        if (hint == "currency")
        {
            if (abs >= 1_000_000_000)
                return $"${value / 1_000_000_000:F1}B";
            if (abs >= 1_000_000)
                return $"${value / 1_000_000:F1}M";
            if (abs >= 1_000)
                return $"${value / 1_000:F1}K";
            return $"${value:F0}";
        }

        if (abs >= 1_000_000_000)
            return $"{value / 1_000_000_000:F1}B";
        if (abs >= 1_000_000)
            return $"{value / 1_000_000:F1}M";
        if (abs >= 1_000)
            return $"{value / 1_000:F1}K";

        if (abs >= 1)
            return value.ToString("F2", CultureInfo.InvariantCulture);
        if (abs >= 0.01)
            return value.ToString("F3", CultureInfo.InvariantCulture);
        if (abs >= 0.0001)
            return value.ToString("F4", CultureInfo.InvariantCulture);

        return value.ToString("G4", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Format a value for chart axis labels (shorter).
    /// </summary>
    public static string FormatAxis(double value)
    {
        var abs = Math.Abs(value);
        if (abs >= 1_000_000_000)
            return $"{value / 1_000_000_000:F1}B";
        if (abs >= 1_000_000)
            return $"{value / 1_000_000:F1}M";
        if (abs >= 1_000)
            return $"{value / 1_000:F1}K";
        if (abs >= 1)
            return value.ToString("F1", CultureInfo.InvariantCulture);
        return value.ToString("G3", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Format a range (e.g., P5–P95 interval).
    /// </summary>
    public static string FormatRange(double low, double high, string? hint = null)
    {
        return $"{Format(low, hint)} — {Format(high, hint)}";
    }
}
