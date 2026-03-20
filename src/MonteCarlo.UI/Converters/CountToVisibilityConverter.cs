using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MonteCarlo.UI.Converters;

/// <summary>
/// Returns Visible when the count is 0, Collapsed otherwise.
/// Used for empty-state messages.
/// </summary>
public class CountToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        int count = value is int i ? i : 0;
        return count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
