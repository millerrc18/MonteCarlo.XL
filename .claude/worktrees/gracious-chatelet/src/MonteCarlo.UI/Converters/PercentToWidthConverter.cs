using System.Globalization;
using System.Windows.Data;

namespace MonteCarlo.UI.Converters;

/// <summary>
/// MultiValueConverter: takes (percent, containerWidth) and returns the proportional width.
/// Used for the progress bar fill.
/// </summary>
public class PercentToWidthConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values.Length < 2) return 0.0;

        double percent = values[0] is double p ? p : 0;
        double containerWidth = values[1] is double w ? w : 0;

        return Math.Max(0, Math.Min(containerWidth, containerWidth * percent / 100.0));
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
