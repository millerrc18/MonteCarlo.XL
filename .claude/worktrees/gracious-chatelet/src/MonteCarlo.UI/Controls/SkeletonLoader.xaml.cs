using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace MonteCarlo.UI.Controls;

/// <summary>
/// Animated skeleton loading placeholder. Shows a pulsing gray rectangle
/// that can be used as a stand-in while content is being computed.
/// Adapts to dark/light theme via <see cref="IsDarkMode"/>.
/// </summary>
public partial class SkeletonLoader : UserControl
{
    public static readonly DependencyProperty IsDarkModeProperty =
        DependencyProperty.Register(nameof(IsDarkMode), typeof(bool), typeof(SkeletonLoader),
            new PropertyMetadata(false, OnThemeChanged));

    /// <summary>Set to true to use dark-mode skeleton colors.</summary>
    public bool IsDarkMode
    {
        get => (bool)GetValue(IsDarkModeProperty);
        set => SetValue(IsDarkModeProperty, value);
    }

    public SkeletonLoader()
    {
        InitializeComponent();
    }

    private static void OnThemeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var control = (SkeletonLoader)d;
        bool dark = (bool)e.NewValue;

        // Update the animation colors based on theme
        var from = dark ? Color.FromRgb(51, 65, 85) : Color.FromRgb(226, 232, 240);   // slate-700 / slate-200
        var to = dark ? Color.FromRgb(71, 85, 105) : Color.FromRgb(203, 213, 225);     // slate-600 / slate-300

        control.SkeletonBrush.Color = from;

        // Restart the storyboard with new colors isn't straightforward in XAML triggers,
        // so we just set the base color and let the animation continue with a visual update
    }
}
