using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace MonteCarlo.UI.Controls;

/// <summary>
/// A text input with built-in validation states: default, valid (green border),
/// and invalid (red border + error message). Validation is debounced at 300ms.
/// </summary>
public partial class ValidatedTextBox : UserControl
{
    private readonly DispatcherTimer _debounceTimer;

    private static readonly SolidColorBrush RedBrush = new(Color.FromRgb(239, 68, 68));
    private static readonly SolidColorBrush GreenBrush = new(Color.FromRgb(16, 185, 129));

    public static readonly DependencyProperty TextProperty =
        DependencyProperty.Register(nameof(Text), typeof(string), typeof(ValidatedTextBox),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnTextPropertyChanged));

    public static readonly DependencyProperty ValidationFuncProperty =
        DependencyProperty.Register(nameof(ValidationFunc), typeof(Func<string, string?>), typeof(ValidatedTextBox),
            new PropertyMetadata(null));

    public static readonly DependencyProperty IsValidProperty =
        DependencyProperty.Register(nameof(IsValid), typeof(bool), typeof(ValidatedTextBox),
            new PropertyMetadata(true));

    /// <summary>The text content of the input.</summary>
    public string Text
    {
        get => (string)GetValue(TextProperty);
        set => SetValue(TextProperty, value);
    }

    /// <summary>
    /// Validation function: takes the text, returns null if valid, or an error message string if invalid.
    /// </summary>
    public Func<string, string?>? ValidationFunc
    {
        get => (Func<string, string?>?)GetValue(ValidationFuncProperty);
        set => SetValue(ValidationFuncProperty, value);
    }

    /// <summary>Whether the current value passes validation.</summary>
    public bool IsValid
    {
        get => (bool)GetValue(IsValidProperty);
        set => SetValue(IsValidProperty, value);
    }

    public ValidatedTextBox()
    {
        InitializeComponent();

        _debounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(300)
        };
        _debounceTimer.Tick += (_, _) =>
        {
            _debounceTimer.Stop();
            RunValidation();
        };
    }

    private static void OnTextPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var control = (ValidatedTextBox)d;
        if (control.InnerTextBox.Text != (string)e.NewValue)
            control.InnerTextBox.Text = (string)e.NewValue;
    }

    private void OnTextChanged(object sender, TextChangedEventArgs e)
    {
        Text = InnerTextBox.Text;
        _debounceTimer.Stop();
        _debounceTimer.Start();
    }

    private void RunValidation()
    {
        var validate = ValidationFunc;
        if (validate == null)
        {
            SetDefaultState();
            return;
        }

        var error = validate(Text);
        if (error == null)
        {
            SetValidState();
        }
        else
        {
            SetInvalidState(error);
        }
    }

    private void SetDefaultState()
    {
        InputBorder.BorderBrush = (Brush)FindResource("BorderBrush");
        ErrorMessage.Visibility = Visibility.Collapsed;
        IsValid = true;
    }

    private void SetValidState()
    {
        InputBorder.BorderBrush = GreenBrush;
        ErrorMessage.Visibility = Visibility.Collapsed;
        IsValid = true;
    }

    private void SetInvalidState(string error)
    {
        InputBorder.BorderBrush = RedBrush;
        ErrorMessage.Text = error;
        ErrorMessage.Visibility = Visibility.Visible;
        IsValid = false;
    }
}
