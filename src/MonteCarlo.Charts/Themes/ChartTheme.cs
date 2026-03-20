using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using SkiaSharp;

namespace MonteCarlo.Charts.Themes;

/// <summary>
/// Defines the shared chart color palette and LiveCharts2 configuration
/// following the visual design language from ROADMAP.md Section 4.
/// </summary>
public static class ChartTheme
{
    // Primary data colors
    public static readonly SKColor Blue500 = SKColor.Parse("#3B82F6");
    public static readonly SKColor Blue500_85 = Blue500.WithAlpha(217); // 0.85 opacity
    public static readonly SKColor Blue500_08 = Blue500.WithAlpha(20);  // 0.08 opacity
    public static readonly SKColor DarkBlue = SKColor.Parse("#1E40AF");

    // Accent colors
    public static readonly SKColor Violet500 = SKColor.Parse("#8B5CF6");
    public static readonly SKColor Emerald500 = SKColor.Parse("#10B981");
    public static readonly SKColor Red500 = SKColor.Parse("#EF4444");
    public static readonly SKColor Orange500 = SKColor.Parse("#F97316");
    public static readonly SKColor Amber500 = SKColor.Parse("#F59E0B");

    // Neutral colors
    public static readonly SKColor Slate500 = SKColor.Parse("#64748B");
    public static readonly SKColor Border = SKColor.Parse("#E2E8F0");
    public static readonly SKColor GridLine = Border.WithAlpha(38); // 0.15 opacity

    // Typography
    public static readonly string FontFamily = "Segoe UI Variable";
    public static readonly float AxisLabelSize = 11f;
    public static readonly float AnnotationLabelSize = 10f;

    /// <summary>
    /// Creates a standard X-axis for charts.
    /// </summary>
    public static Axis CreateXAxis(string? name = null, Func<double, string>? labeler = null)
    {
        return new Axis
        {
            Name = name,
            NamePaint = name != null ? new SolidColorPaint(Slate500) { SKTypeface = SKTypeface.FromFamilyName(FontFamily) } : null,
            LabelsPaint = new SolidColorPaint(Slate500)
            {
                SKTypeface = SKTypeface.FromFamilyName(FontFamily),
                FontFamily = FontFamily
            },
            TextSize = AxisLabelSize,
            SeparatorsPaint = new SolidColorPaint(GridLine) { StrokeThickness = 1 },
            Labeler = labeler ?? (v => v.ToString("N0"))
        };
    }

    /// <summary>
    /// Creates a standard Y-axis for charts.
    /// </summary>
    public static Axis CreateYAxis(string? name = null, Func<double, string>? labeler = null)
    {
        return new Axis
        {
            Name = name,
            NamePaint = name != null ? new SolidColorPaint(Slate500) { SKTypeface = SKTypeface.FromFamilyName(FontFamily) } : null,
            LabelsPaint = new SolidColorPaint(Slate500)
            {
                SKTypeface = SKTypeface.FromFamilyName(FontFamily),
                FontFamily = FontFamily
            },
            TextSize = AxisLabelSize,
            SeparatorsPaint = new SolidColorPaint(GridLine) { StrokeThickness = 1 },
            Labeler = labeler ?? (v => v.ToString("P1")),
            MinLimit = 0
        };
    }

    /// <summary>
    /// Creates a percentile marker section (vertical dashed line).
    /// </summary>
    public static RectangularSection CreatePercentileMarker(double value, string label)
    {
        return new RectangularSection
        {
            Xi = value,
            Xj = value,
            Stroke = new SolidColorPaint(Amber500)
            {
                StrokeThickness = 1.5f,
                PathEffect = new LiveChartsCore.SkiaSharpView.Painting.Effects.DashEffect(new[] { 4f, 4f })
            },
            Label = label,
            LabelPaint = new SolidColorPaint(Amber500)
            {
                SKTypeface = SKTypeface.FromFamilyName(FontFamily),
                FontFamily = FontFamily
            },
            LabelSize = AnnotationLabelSize
        };
    }

    /// <summary>
    /// Creates a target line section (vertical dashed red line).
    /// </summary>
    public static RectangularSection CreateTargetLine(double value, string label)
    {
        return new RectangularSection
        {
            Xi = value,
            Xj = value,
            Stroke = new SolidColorPaint(Red500)
            {
                StrokeThickness = 2f,
                PathEffect = new LiveChartsCore.SkiaSharpView.Painting.Effects.DashEffect(new[] { 6f, 4f })
            },
            Label = label,
            LabelPaint = new SolidColorPaint(Red500)
            {
                SKTypeface = SKTypeface.FromFamilyName(FontFamily),
                FontFamily = FontFamily
            },
            LabelSize = AnnotationLabelSize
        };
    }
}
