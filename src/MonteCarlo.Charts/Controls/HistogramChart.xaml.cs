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
/// LiveCharts2 histogram chart with percentile markers and optional target line.
/// Follows the visual design language from ROADMAP.md Section 4.
/// </summary>
public partial class HistogramChart : UserControl
{
    public static readonly DependencyProperty HistogramDataProperty =
        DependencyProperty.Register(nameof(HistogramData), typeof(HistogramData), typeof(HistogramChart),
            new PropertyMetadata(null, OnDataChanged));

    public static readonly DependencyProperty P10Property =
        DependencyProperty.Register(nameof(P10), typeof(double?), typeof(HistogramChart),
            new PropertyMetadata(null, OnMarkersChanged));

    public static readonly DependencyProperty P50Property =
        DependencyProperty.Register(nameof(P50), typeof(double?), typeof(HistogramChart),
            new PropertyMetadata(null, OnMarkersChanged));

    public static readonly DependencyProperty P90Property =
        DependencyProperty.Register(nameof(P90), typeof(double?), typeof(HistogramChart),
            new PropertyMetadata(null, OnMarkersChanged));

    public static readonly DependencyProperty TargetValueProperty =
        DependencyProperty.Register(nameof(TargetValue), typeof(double?), typeof(HistogramChart),
            new PropertyMetadata(null, OnMarkersChanged));

    public static readonly DependencyProperty TargetLabelProperty =
        DependencyProperty.Register(nameof(TargetLabel), typeof(string), typeof(HistogramChart),
            new PropertyMetadata(null, OnMarkersChanged));

    public static readonly DependencyProperty ValueFormatterProperty =
        DependencyProperty.Register(nameof(ValueFormatter), typeof(Func<double, string>), typeof(HistogramChart),
            new PropertyMetadata(null, OnDataChanged));

    /// <summary>Histogram bin data to render.</summary>
    public HistogramData? HistogramData
    {
        get => (HistogramData?)GetValue(HistogramDataProperty);
        set => SetValue(HistogramDataProperty, value);
    }

    public double? P10 { get => (double?)GetValue(P10Property); set => SetValue(P10Property, value); }
    public double? P50 { get => (double?)GetValue(P50Property); set => SetValue(P50Property, value); }
    public double? P90 { get => (double?)GetValue(P90Property); set => SetValue(P90Property, value); }
    public double? TargetValue { get => (double?)GetValue(TargetValueProperty); set => SetValue(TargetValueProperty, value); }
    public string? TargetLabel { get => (string?)GetValue(TargetLabelProperty); set => SetValue(TargetLabelProperty, value); }
    public Func<double, string>? ValueFormatter { get => (Func<double, string>?)GetValue(ValueFormatterProperty); set => SetValue(ValueFormatterProperty, value); }

    /// <summary>LiveCharts series bound to the chart.</summary>
    public ObservableCollection<ISeries> Series { get; } = new();

    /// <summary>X axes configuration.</summary>
    public ObservableCollection<Axis> XAxes { get; } = new();

    /// <summary>Y axes configuration.</summary>
    public ObservableCollection<Axis> YAxes { get; } = new();

    /// <summary>Sections (percentile markers, target line).</summary>
    public ObservableCollection<RectangularSection> Sections { get; } = new();

    public HistogramChart()
    {
        InitializeComponent();
    }

    private static void OnDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((HistogramChart)d).RebuildChart();
    }

    private static void OnMarkersChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((HistogramChart)d).RebuildSections();
    }

    private void RebuildChart()
    {
        Series.Clear();
        XAxes.Clear();
        YAxes.Clear();
        Sections.Clear();

        var data = HistogramData;
        if (data == null) return;

        var formatter = ValueFormatter ?? (v => v.ToString("N0"));

        // Build column series from histogram bins
        var values = new List<double>(data.RelativeFrequencies);

        Series.Add(new ColumnSeries<double>
        {
            Values = values,
            Fill = new SolidColorPaint(ChartTheme.Blue500_85),
            Stroke = null,
            Padding = 1,
            MaxBarWidth = double.MaxValue
        });

        // X axis — map bin index to bin center value
        var binCenters = data.BinCenters;
        XAxes.Add(ChartTheme.CreateXAxis(labeler: v =>
        {
            int idx = (int)Math.Round(v);
            if (idx >= 0 && idx < binCenters.Length)
                return formatter(binCenters[idx]);
            return "";
        }));

        // Y axis — relative frequency
        YAxes.Add(ChartTheme.CreateYAxis(labeler: v => v.ToString("P1")));

        RebuildSections();
    }

    private void RebuildSections()
    {
        Sections.Clear();
        var data = HistogramData;
        if (data == null) return;

        var formatter = ValueFormatter ?? (v => v.ToString("N0"));
        var binCenters = data.BinCenters;
        var binEdges = data.BinEdges;

        // Helper: convert a value to the bin index (interpolated)
        double ValueToIndex(double val)
        {
            if (binEdges.Length < 2) return 0;
            double binWidth = binEdges[1] - binEdges[0];
            return (val - binEdges[0]) / binWidth - 0.5;
        }

        if (P10.HasValue)
            Sections.Add(ChartTheme.CreatePercentileMarker(ValueToIndex(P10.Value), $"P10: {formatter(P10.Value)}"));
        if (P50.HasValue)
            Sections.Add(ChartTheme.CreatePercentileMarker(ValueToIndex(P50.Value), $"P50: {formatter(P50.Value)}"));
        if (P90.HasValue)
            Sections.Add(ChartTheme.CreatePercentileMarker(ValueToIndex(P90.Value), $"P90: {formatter(P90.Value)}"));

        if (TargetValue.HasValue)
            Sections.Add(ChartTheme.CreateTargetLine(ValueToIndex(TargetValue.Value), TargetLabel ?? $"Target: {formatter(TargetValue.Value)}"));
    }
}
