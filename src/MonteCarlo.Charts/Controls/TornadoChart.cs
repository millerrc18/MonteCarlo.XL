using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MonteCarlo.Engine.Analysis;
using SkiaSharp;
using SkiaSharp.Views.Desktop;
using SkiaSharp.Views.WPF;

namespace MonteCarlo.Charts.Controls;

/// <summary>
/// Custom SkiaSharp WPF control that renders a tornado sensitivity chart.
/// Horizontal bars extend left (decrease, orange) and right (increase, blue)
/// from a center baseline, sorted by absolute impact.
/// </summary>
public class TornadoChart : SKElement
{
    #region Dependency Properties

    public static readonly DependencyProperty ResultsProperty =
        DependencyProperty.Register(nameof(Results), typeof(IReadOnlyList<SensitivityResult>), typeof(TornadoChart),
            new PropertyMetadata(null, OnDataChanged));

    public static readonly DependencyProperty BaseOutputValueProperty =
        DependencyProperty.Register(nameof(BaseOutputValue), typeof(double), typeof(TornadoChart),
            new PropertyMetadata(0.0, OnDataChanged));

    public static readonly DependencyProperty MaxInputsToShowProperty =
        DependencyProperty.Register(nameof(MaxInputsToShow), typeof(int), typeof(TornadoChart),
            new PropertyMetadata(10, OnDataChanged));

    public static readonly DependencyProperty ShowAllProperty =
        DependencyProperty.Register(nameof(ShowAll), typeof(bool), typeof(TornadoChart),
            new PropertyMetadata(false, OnDataChanged));

    public static readonly DependencyProperty ValueFormatterProperty =
        DependencyProperty.Register(nameof(ValueFormatter), typeof(Func<double, string>), typeof(TornadoChart),
            new PropertyMetadata(null, OnDataChanged));

    /// <summary>Sensitivity results to render (sorted by swing descending).</summary>
    public IReadOnlyList<SensitivityResult>? Results
    {
        get => (IReadOnlyList<SensitivityResult>?)GetValue(ResultsProperty);
        set => SetValue(ResultsProperty, value);
    }

    /// <summary>Base output value (center baseline).</summary>
    public double BaseOutputValue
    {
        get => (double)GetValue(BaseOutputValueProperty);
        set => SetValue(BaseOutputValueProperty, value);
    }

    /// <summary>Maximum inputs to show (default 10).</summary>
    public int MaxInputsToShow
    {
        get => (int)GetValue(MaxInputsToShowProperty);
        set => SetValue(MaxInputsToShowProperty, value);
    }

    /// <summary>Whether to show all inputs (override MaxInputsToShow).</summary>
    public bool ShowAll
    {
        get => (bool)GetValue(ShowAllProperty);
        set => SetValue(ShowAllProperty, value);
    }

    /// <summary>Value formatter for annotations.</summary>
    public Func<double, string>? ValueFormatter
    {
        get => (Func<double, string>?)GetValue(ValueFormatterProperty);
        set => SetValue(ValueFormatterProperty, value);
    }

    #endregion

    #region Colors

    private static readonly SKColor BarIncreaseColor = SKColor.Parse("#3B82F6");
    private static readonly SKColor BarDecreaseColor = SKColor.Parse("#F97316");
    private static readonly SKColor BarHoverIncreaseColor = SKColor.Parse("#60A5FA");
    private static readonly SKColor BarHoverDecreaseColor = SKColor.Parse("#FB923C");
    private static readonly SKColor BaselineColor = SKColor.Parse("#E2E8F0");
    private static readonly SKColor LabelColor = SKColor.Parse("#0F172A");
    private static readonly SKColor SubLabelColor = SKColor.Parse("#94A3B8");
    private static readonly SKColor AnnotationColor = SKColor.Parse("#64748B");

    #endregion

    #region Layout Constants

    private const float TopMargin = 24f;
    private const float BottomMargin = 32f;
    private const float LeftPadding = 8f;
    private const float RightPadding = 8f;
    private const float BarHeight = 28f;
    private const float BarGap = 6f;
    private const float BarCornerRadius = 2f;
    private const float LabelFontSize = 12f;
    private const float SubLabelFontSize = 10f;
    private const float AnnotationFontSize = 10f;
    private const float LabelRightPadding = 12f;
    private const float AnnotationLeftPadding = 8f;

    #endregion

    private int _hoveredIndex = -1;
    private readonly List<SKRect> _barRects = new();
    private SKTypeface? _typeface;

    private static void OnDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((TornadoChart)d).InvalidateVisual();
    }

    /// <summary>
    /// Gets the items to display (capped by MaxInputsToShow unless ShowAll).
    /// </summary>
    private IReadOnlyList<SensitivityResult> GetVisibleItems()
    {
        var results = Results;
        if (results == null || results.Count == 0)
            return Array.Empty<SensitivityResult>();

        if (ShowAll || results.Count <= MaxInputsToShow)
            return results;

        return results.Take(MaxInputsToShow).ToList();
    }

    /// <summary>
    /// Calculates the desired height based on visible items.
    /// </summary>
    public double DesiredChartHeight
    {
        get
        {
            var items = GetVisibleItems();
            if (items.Count == 0) return 80;
            return TopMargin + BottomMargin + items.Count * (BarHeight + BarGap);
        }
    }

    protected override void OnPaintSurface(SKPaintSurfaceEventArgs e)
    {
        var canvas = e.Surface.Canvas;
        canvas.Clear(SKColors.Transparent);

        var items = GetVisibleItems();
        if (items.Count == 0)
        {
            DrawEmptyState(canvas, e.Info);
            return;
        }

        _typeface = SKTypeface.FromFamilyName("Segoe UI Variable") ??
                    SKTypeface.FromFamilyName("Segoe UI") ??
                    SKTypeface.Default;

        var formatter = ValueFormatter ?? (v => v.ToString("N0"));
        float w = e.Info.Width;
        float h = e.Info.Height;
        double baseValue = BaseOutputValue;

        // Measure label column width
        float labelColumnWidth = MeasureLabelColumnWidth(items) + LabelRightPadding + LeftPadding;
        float annotationWidth = MeasureAnnotationWidth(items, formatter) + AnnotationLeftPadding + RightPadding;

        float chartAreaWidth = w - labelColumnWidth - annotationWidth;
        if (chartAreaWidth < 40) chartAreaWidth = 40;

        float centerX = labelColumnWidth + chartAreaWidth / 2f;

        // Find max swing for scaling
        double maxSwing = items.Max(r => Math.Max(
            Math.Abs(r.OutputAtInputP10 - baseValue),
            Math.Abs(r.OutputAtInputP90 - baseValue)));

        if (maxSwing <= 0) maxSwing = 1;
        float scaleFactor = (chartAreaWidth / 2f) / (float)maxSwing;

        // Store bar rects for hit-testing
        _barRects.Clear();

        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i];
            float y = TopMargin + i * (BarHeight + BarGap);
            bool isHovered = i == _hoveredIndex;

            // Determine which side each bar goes
            double leftValue = Math.Min(item.OutputAtInputP10, item.OutputAtInputP90);
            double rightValue = Math.Max(item.OutputAtInputP10, item.OutputAtInputP90);

            float leftBarWidth = (float)Math.Abs(leftValue - baseValue) * scaleFactor;
            float rightBarWidth = (float)Math.Abs(rightValue - baseValue) * scaleFactor;

            // Determine colors: if P10 < base, left bar is orange (decrease), right is blue (increase)
            bool p10IsLower = item.OutputAtInputP10 < item.OutputAtInputP90;

            // Draw left bar (decrease side)
            var leftRect = new SKRect(centerX - leftBarWidth, y, centerX, y + BarHeight);
            using var leftPaint = new SKPaint
            {
                Color = isHovered
                    ? (p10IsLower ? BarHoverDecreaseColor : BarHoverIncreaseColor)
                    : (p10IsLower ? BarDecreaseColor : BarIncreaseColor),
                IsAntialias = true
            };
            canvas.DrawRoundRect(new SKRoundRect(leftRect, BarCornerRadius, 0, 0, BarCornerRadius), leftPaint);

            // Draw right bar (increase side)
            var rightRect = new SKRect(centerX, y, centerX + rightBarWidth, y + BarHeight);
            using var rightPaint = new SKPaint
            {
                Color = isHovered
                    ? (p10IsLower ? BarHoverIncreaseColor : BarHoverDecreaseColor)
                    : (p10IsLower ? BarIncreaseColor : BarDecreaseColor),
                IsAntialias = true
            };
            canvas.DrawRoundRect(new SKRoundRect(rightRect, 0, BarCornerRadius, BarCornerRadius, 0), rightPaint);

            // Store full rect for hit-testing
            _barRects.Add(new SKRect(centerX - leftBarWidth, y, centerX + rightBarWidth, y + BarHeight));

            // Draw label (left side)
            using var labelPaint = new SKPaint
            {
                Typeface = _typeface,
                TextSize = LabelFontSize,
                Color = LabelColor,
                IsAntialias = true,
                TextAlign = SKTextAlign.Right
            };
            canvas.DrawText(item.InputLabel, labelColumnWidth - LabelRightPadding, y + BarHeight / 2f + LabelFontSize / 3f, labelPaint);

            // Draw annotation (right side) — output range
            using var annotPaint = new SKPaint
            {
                Typeface = _typeface,
                TextSize = AnnotationFontSize,
                Color = AnnotationColor,
                IsAntialias = true,
                TextAlign = SKTextAlign.Left
            };
            string annotation = $"{formatter(item.OutputAtInputP10)} — {formatter(item.OutputAtInputP90)}";
            float annotX = labelColumnWidth + chartAreaWidth + AnnotationLeftPadding;
            canvas.DrawText(annotation, annotX, y + BarHeight / 2f + AnnotationFontSize / 3f, annotPaint);
        }

        // Draw center baseline
        using var baselinePaint = new SKPaint
        {
            Color = BaselineColor,
            StrokeWidth = 1,
            IsAntialias = true
        };
        float topY = TopMargin - 4;
        float bottomY = TopMargin + items.Count * (BarHeight + BarGap);
        canvas.DrawLine(centerX, topY, centerX, bottomY, baselinePaint);

        // Draw base value label below baseline
        using var baseValuePaint = new SKPaint
        {
            Typeface = _typeface,
            TextSize = AnnotationFontSize,
            Color = AnnotationColor,
            IsAntialias = true,
            TextAlign = SKTextAlign.Center
        };
        canvas.DrawText($"Base: {formatter(baseValue)}", centerX, bottomY + 16, baseValuePaint);

        // Draw header labels
        using var headerPaint = new SKPaint
        {
            Typeface = _typeface,
            TextSize = 9f,
            Color = SubLabelColor,
            IsAntialias = true,
            TextAlign = SKTextAlign.Center
        };
        canvas.DrawText("◀ Decreases", centerX - chartAreaWidth / 4f, 14, headerPaint);
        canvas.DrawText("Increases ▶", centerX + chartAreaWidth / 4f, 14, headerPaint);

        // Draw tooltip if hovered
        if (_hoveredIndex >= 0 && _hoveredIndex < items.Count)
            DrawTooltip(canvas, items[_hoveredIndex], baseValue, formatter, w, h);
    }

    private void DrawEmptyState(SKCanvas canvas, SKImageInfo info)
    {
        _typeface ??= SKTypeface.FromFamilyName("Segoe UI Variable") ??
                      SKTypeface.FromFamilyName("Segoe UI") ??
                      SKTypeface.Default;

        using var paint = new SKPaint
        {
            Typeface = _typeface,
            TextSize = 13,
            Color = SubLabelColor,
            IsAntialias = true,
            TextAlign = SKTextAlign.Center
        };
        canvas.DrawText("No sensitivity data available", info.Width / 2f, info.Height / 2f, paint);
    }

    private void DrawTooltip(SKCanvas canvas, SensitivityResult item, double baseValue,
        Func<double, string> formatter, float canvasWidth, float canvasHeight)
    {
        string line1 = item.InputLabel;
        string line2 = $"P10 → Output: {formatter(item.OutputAtInputP10)}";
        string line3 = $"P90 → Output: {formatter(item.OutputAtInputP90)}";
        string line4 = $"Swing: {formatter(item.Swing)} ({item.ContributionToVariance:F1}% of variance)";

        using var bgPaint = new SKPaint
        {
            Color = new SKColor(15, 23, 42, 230), // near-black
            IsAntialias = true
        };
        using var textPaint = new SKPaint
        {
            Typeface = _typeface,
            TextSize = 11,
            Color = SKColors.White,
            IsAntialias = true
        };
        using var titlePaint = new SKPaint
        {
            Typeface = _typeface,
            TextSize = 12,
            Color = SKColors.White,
            IsAntialias = true,
            FakeBoldText = true
        };

        float padding = 10;
        float lineHeight = 16;
        float tooltipWidth = Math.Max(
            Math.Max(titlePaint.MeasureText(line1), textPaint.MeasureText(line2)),
            Math.Max(textPaint.MeasureText(line3), textPaint.MeasureText(line4))
        ) + padding * 2;
        float tooltipHeight = lineHeight * 4 + padding * 2;

        // Position tooltip near the hovered bar
        float tooltipX = Math.Min(canvasWidth - tooltipWidth - 8, canvasWidth / 2f);
        float tooltipY = TopMargin + _hoveredIndex * (BarHeight + BarGap) - tooltipHeight - 4;
        if (tooltipY < 4) tooltipY = TopMargin + (_hoveredIndex + 1) * (BarHeight + BarGap) + 4;

        var rect = new SKRect(tooltipX, tooltipY, tooltipX + tooltipWidth, tooltipY + tooltipHeight);
        canvas.DrawRoundRect(rect, 6, 6, bgPaint);

        float textX = tooltipX + padding;
        float textY = tooltipY + padding + 12;
        canvas.DrawText(line1, textX, textY, titlePaint);
        canvas.DrawText(line2, textX, textY + lineHeight, textPaint);
        canvas.DrawText(line3, textX, textY + lineHeight * 2, textPaint);
        canvas.DrawText(line4, textX, textY + lineHeight * 3, textPaint);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);

        var pos = e.GetPosition(this);
        float x = (float)(pos.X * (RenderSize.Width > 0 ? ActualWidth / RenderSize.Width : 1));
        float y = (float)(pos.Y * (RenderSize.Height > 0 ? ActualHeight / RenderSize.Height : 1));

        // Simple Y-band hit testing
        int newHovered = -1;
        for (int i = 0; i < _barRects.Count; i++)
        {
            float barY = TopMargin + i * (BarHeight + BarGap);
            if (y >= barY && y <= barY + BarHeight)
            {
                newHovered = i;
                break;
            }
        }

        if (newHovered != _hoveredIndex)
        {
            _hoveredIndex = newHovered;
            InvalidateVisual();
        }
    }

    protected override void OnMouseLeave(MouseEventArgs e)
    {
        base.OnMouseLeave(e);
        if (_hoveredIndex != -1)
        {
            _hoveredIndex = -1;
            InvalidateVisual();
        }
    }

    private float MeasureLabelColumnWidth(IReadOnlyList<SensitivityResult> items)
    {
        using var paint = new SKPaint
        {
            Typeface = _typeface,
            TextSize = LabelFontSize,
            IsAntialias = true
        };
        return items.Max(r => paint.MeasureText(r.InputLabel));
    }

    private float MeasureAnnotationWidth(IReadOnlyList<SensitivityResult> items, Func<double, string> formatter)
    {
        using var paint = new SKPaint
        {
            Typeface = _typeface,
            TextSize = AnnotationFontSize,
            IsAntialias = true
        };
        return items.Max(r => paint.MeasureText($"{formatter(r.OutputAtInputP10)} — {formatter(r.OutputAtInputP90)}"));
    }
}
