using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MonteCarlo.UI.Controls;

/// <summary>
/// Enhanced numeric input TextBox that supports:
/// - Paste handling
/// - Scientific notation (e.g., "1e6" → 1000000)
/// - Stripping commas and currency symbols on input
/// - Formatted display on blur (e.g., "1000000" → "1,000,000")
/// </summary>
public class NumberInputTextBox : TextBox
{
    public static readonly DependencyProperty NumericValueProperty =
        DependencyProperty.Register(nameof(NumericValue), typeof(double?), typeof(NumberInputTextBox),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnNumericValueChanged));

    public static readonly DependencyProperty FormatStringProperty =
        DependencyProperty.Register(nameof(FormatString), typeof(string), typeof(NumberInputTextBox),
            new PropertyMetadata("N0"));

    /// <summary>The parsed numeric value.</summary>
    public double? NumericValue
    {
        get => (double?)GetValue(NumericValueProperty);
        set => SetValue(NumericValueProperty, value);
    }

    /// <summary>Format string for display on blur (default: "N0" — number with thousand separators).</summary>
    public string FormatString
    {
        get => (string)GetValue(FormatStringProperty);
        set => SetValue(FormatStringProperty, value);
    }

    private bool _isUpdating;

    public NumberInputTextBox()
    {
        LostFocus += OnLostFocus;
        GotFocus += OnGotFocus;
        DataObject.AddPastingHandler(this, OnPaste);
    }

    protected override void OnTextChanged(TextChangedEventArgs e)
    {
        base.OnTextChanged(e);

        if (_isUpdating) return;

        var parsed = ParseInput(Text);
        if (parsed.HasValue)
        {
            _isUpdating = true;
            NumericValue = parsed.Value;
            _isUpdating = false;
        }
    }

    private static void OnNumericValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var control = (NumberInputTextBox)d;
        if (control._isUpdating) return;

        control._isUpdating = true;
        if (e.NewValue is double val)
        {
            // When focused, show raw value; when not focused, show formatted
            if (control.IsFocused)
                control.Text = val.ToString(CultureInfo.InvariantCulture);
            else
                control.Text = val.ToString(control.FormatString, CultureInfo.CurrentCulture);
        }
        else
        {
            control.Text = string.Empty;
        }
        control._isUpdating = false;
    }

    private void OnGotFocus(object sender, RoutedEventArgs e)
    {
        // Show raw (editable) value on focus
        if (NumericValue.HasValue)
        {
            _isUpdating = true;
            Text = NumericValue.Value.ToString(CultureInfo.InvariantCulture);
            _isUpdating = false;
            SelectAll();
        }
    }

    private void OnLostFocus(object sender, RoutedEventArgs e)
    {
        // Show formatted value on blur
        var parsed = ParseInput(Text);
        _isUpdating = true;

        if (parsed.HasValue)
        {
            NumericValue = parsed.Value;
            Text = parsed.Value.ToString(FormatString, CultureInfo.CurrentCulture);
        }

        _isUpdating = false;
    }

    private void OnPaste(object sender, DataObjectPastingEventArgs e)
    {
        if (e.DataObject.GetDataPresent(typeof(string)))
        {
            var text = (string)e.DataObject.GetData(typeof(string));
            var cleaned = CleanInput(text);
            var parsed = ParseInput(cleaned);

            if (parsed.HasValue)
            {
                e.CancelCommand();
                _isUpdating = true;
                NumericValue = parsed.Value;
                Text = parsed.Value.ToString(CultureInfo.InvariantCulture);
                _isUpdating = false;
                CaretIndex = Text.Length;
            }
            else
            {
                e.CancelCommand(); // Reject non-numeric paste
            }
        }
    }

    /// <summary>
    /// Cleans input by removing commas, currency symbols, and whitespace.
    /// </summary>
    private static string CleanInput(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return string.Empty;

        return input
            .Replace(",", "")
            .Replace("$", "")
            .Replace("€", "")
            .Replace("£", "")
            .Replace("¥", "")
            .Replace(" ", "")
            .Trim();
    }

    /// <summary>
    /// Parses the input string, supporting scientific notation (e.g., "1e6", "2.5E-3").
    /// </summary>
    private static double? ParseInput(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return null;

        var cleaned = CleanInput(input);
        if (string.IsNullOrEmpty(cleaned))
            return null;

        if (double.TryParse(cleaned,
                NumberStyles.Float | NumberStyles.AllowExponent,
                CultureInfo.InvariantCulture,
                out double result))
        {
            return result;
        }

        return null;
    }
}
