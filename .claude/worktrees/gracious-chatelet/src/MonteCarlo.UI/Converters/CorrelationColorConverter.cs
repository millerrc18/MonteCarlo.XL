using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace MonteCarlo.UI.Converters;

/// <summary>
/// Converts a correlation value (-1 to +1) to a background color brush.
/// Positive = blue, Negative = orange, Zero = transparent.
/// Alpha scales with magnitude.
/// </summary>
public class CorrelationColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not double corr)
            return Brushes.Transparent;

        if (Math.Abs(corr) < 0.01)
            return Brushes.Transparent;

        byte alpha = (byte)(Math.Min(Math.Abs(corr), 1.0) * 180);

        if (corr > 0)
            return new SolidColorBrush(Color.FromArgb(alpha, 59, 130, 246)); // Blue-500

        return new SolidColorBrush(Color.FromArgb(alpha, 249, 115, 22)); // Orange-500
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
