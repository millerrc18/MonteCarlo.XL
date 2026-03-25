using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MonteCarlo.UI.Converters;

/// <summary>
/// Converts bool to Visibility. True = Visible, False = Collapsed.
/// Set ConverterParameter to "invert" for inverse behavior.
/// </summary>
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool flag = value is bool b && b;
        bool invert = parameter is string s && s.Equals("invert", StringComparison.OrdinalIgnoreCase);
        if (invert) flag = !flag;
        return flag ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
