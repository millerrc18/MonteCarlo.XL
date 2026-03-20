using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MonteCarlo.UI.Converters;

/// <summary>
/// Shows an element when the bound distribution name matches the ConverterParameter.
/// Supports comma-separated parameter values (e.g., "Triangular,PERT").
/// </summary>
public class DistributionVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not string current || parameter is not string targets)
            return Visibility.Collapsed;

        var names = targets.Split(',', StringSplitOptions.TrimEntries);
        return names.Any(n => n.Equals(current, StringComparison.OrdinalIgnoreCase))
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
