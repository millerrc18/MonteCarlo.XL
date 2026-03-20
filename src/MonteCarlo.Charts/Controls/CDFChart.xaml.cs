using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using MonteCarlo.Charts.Themes;
using MonteCarlo.Engine.Analysis;

namespace MonteCarlo.Charts.Controls;

/// <summary>
/// LiveCharts2 CDF chart — smooth S-curve with area fill and percentile reference lines.
/// </summary>
public partial class CDFChart : UserControl
{
    public static readonly DependencyProperty StatsProperty =
        DependencyProperty.Register(nameof(Stats), typeof(SummaryStatistics), typeof(CDFChart),
            new PropertyMetadata(null, OnDataChanged));

    public static readonly DependencyProperty ValueFormatterProperty =
        DependencyProperty.Register(nameof(ValueFormatter), typeof(Func<double, string>), typeof(CDFChart),
            new PropertyMetadata(null, OnDataChanged));

    /// <summary>Summary statistics containing the sorted values for the CDF.</summary>
    public SummaryStatistics? Stats
    {
        get => (SummaryStatistics?)GetValue(StatsProperty);
        set => SetValue(StatsProperty, value);
    }

    /// <summary>Value formatter for axis labels.</summary>
    public Func<double, string>? ValueFormatter
    {
        get => (Func<double, string>?)GetValue(ValueFormatterProperty);
        set => SetValue(ValueFormatterProperty, value);
    }

    public ObservableCollection<ISeries> Series { get; } = new();
    public ObservableCollection<Axis> XAxes { get; } = new();
    public ObservableCollection<Axis> YAxes { get; } = new();
    public ObservableCollection<RectangularSection> Sections { get; } = new();

    public CDFChart()
    {
        InitializeComponent();
    }

    private static void OnDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((CDFChart)d).RebuildChart();
    }

    private void RebuildChart()
    {
        Series.Clear();
        XAxes.Clear();
        YAxes.Clear();
        Sections.Clear();

        var stats = Stats;
        if (stats == null) return;

        var formatter = ValueFormatter ?? (v => v.ToString("N0"));
        var sorted = stats.SortedValues;
        int n = sorted.Length;

        // Subsample for performance — take ~200 points for a smooth curve
        int step = Math.Max(1, n / 200);
        var points = new List<LiveChartsCore.Defaults.ObservablePoint>();

        for (int i = 0; i < n; i += step)
        {
            double x = sorted[i];
            double y = (double)(i + 1) / n; // CDF value
            points.Add(new LiveChartsCore.Defaults.ObservablePoint(x, y));
        }

        // Ensure last point is included
        if (points.Count == 0 || points[^1].X != sorted[^1])
            points.Add(new LiveChartsCore.Defaults.ObservablePoint(sorted[^1], 1.0));

        Series.Add(new LineSeries<LiveChartsCore.Defaults.ObservablePoint>
        {
            Values = points,
            Stroke = new SolidColorPaint(ChartTheme.Blue500) { StrokeThickness = 2 },
            Fill = new SolidColorPaint(ChartTheme.Blue500_08),
            GeometrySize = 0,
            LineSmoothness = 0.5
        });

        XAxes.Add(ChartTheme.CreateXAxis(labeler: formatter));

        YAxes.Add(ChartTheme.CreateYAxis(
            labeler: v => v.ToString("P0")
        ));

        // Percentile reference lines (horizontal)
        var p10 = stats.P10;
        var p50 = stats.Median;
        var p90 = stats.P90;

        Sections.Add(CreateHorizontalRef(0.10, p10, $"P10: {formatter(p10)}"));
        Sections.Add(CreateHorizontalRef(0.50, p50, $"P50: {formatter(p50)}"));
        Sections.Add(CreateHorizontalRef(0.90, p90, $"P90: {formatter(p90)}"));
    }

    private static RectangularSection CreateHorizontalRef(double yValue, double xValue, string label)
    {
        return new RectangularSection
        {
            Yi = yValue,
            Yj = yValue,
            Stroke = new SolidColorPaint(ChartTheme.Amber500)
            {
                StrokeThickness = 1f,
                PathEffect = new LiveChartsCore.SkiaSharpView.Painting.Effects.DashEffect(new[] { 4f, 4f })
            },
            Label = label,
            LabelPaint = new SolidColorPaint(ChartTheme.Amber500)
            {
                SKTypeface = SkiaSharp.SKTypeface.FromFamilyName(ChartTheme.FontFamily),
                FontFamily = ChartTheme.FontFamily
            },
            LabelSize = ChartTheme.AnnotationLabelSize
        };
    }
}
