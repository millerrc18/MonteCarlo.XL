using System.Globalization;
using System.Windows.Data;

namespace MonteCarlo.UI.Converters;

/// <summary>
/// Returns true if the bound string value matches the ConverterParameter.
/// Used for RadioButton IsChecked binding with string-based view names.
/// </summary>
public class StringMatchConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is string s && parameter is string p &&
               string.Equals(s, p, StringComparison.OrdinalIgnoreCase);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is true && parameter is string p ? p : Binding.DoNothing;
    }
}
