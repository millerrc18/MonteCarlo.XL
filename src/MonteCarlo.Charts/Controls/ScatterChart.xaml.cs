using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using LiveChartsCore.SkiaSharpView.WPF;
using MonteCarlo.Charts.Themes;

namespace MonteCarlo.Charts.Controls;

/// <summary>
/// LiveCharts2 scatter chart — plots input vs output values to show correlation.
/// </summary>
public partial class ScatterChart : UserControl
{
    public static readonly DependencyProperty InputValuesProperty =
        DependencyProperty.Register(nameof(InputValues), typeof(double[]), typeof(ScatterChart),
            new PropertyMetadata(null, OnDataChanged));

    public static readonly DependencyProperty OutputValuesProperty =
        DependencyProperty.Register(nameof(OutputValues), typeof(double[]), typeof(ScatterChart),
            new PropertyMetadata(null, OnDataChanged));

    public static readonly DependencyProperty InputLabelProperty =
        DependencyProperty.Register(nameof(InputLabel), typeof(string), typeof(ScatterChart),
            new PropertyMetadata(null, OnDataChanged));

    /// <summary>Input sample values (one per iteration).</summary>
    public double[]? InputValues
    {
        get => (double[]?)GetValue(InputValuesProperty);
        set => SetValue(InputValuesProperty, value);
    }

    /// <summary>Output values (one per iteration, same length as InputValues).</summary>
    public double[]? OutputValues
    {
        get => (double[]?)GetValue(OutputValuesProperty);
        set => SetValue(OutputValuesProperty, value);
    }

    /// <summary>Label for the X-axis (input name).</summary>
    public string? InputLabel
    {
        get => (string?)GetValue(InputLabelProperty);
        set => SetValue(InputLabelProperty, value);
    }

    public ObservableCollection<ISeries> Series { get; } = new();
    public ObservableCollection<Axis> XAxes { get; } = new();
    public ObservableCollection<Axis> YAxes { get; } = new();

    public ScatterChart()
    {
        _ = typeof(CartesianChart).Assembly;
        InitializeComponent();
    }

    private static void OnDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((ScatterChart)d).Rebuild();
    }

    private void Rebuild()
    {
        Series.Clear();
        XAxes.Clear();
        YAxes.Clear();

        var inputs = InputValues;
        var outputs = OutputValues;
        if (inputs == null || outputs == null || inputs.Length == 0 || outputs.Length == 0)
            return;

        int n = Math.Min(inputs.Length, outputs.Length);

        // Subsample to max 500 points for performance
        const int maxPoints = 500;
        int step = Math.Max(1, n / maxPoints);
        var points = new List<LiveChartsCore.Defaults.ObservablePoint>();

        for (int i = 0; i < n; i += step)
        {
            points.Add(new LiveChartsCore.Defaults.ObservablePoint(inputs[i], outputs[i]));
        }

        Series.Add(new ScatterSeries<LiveChartsCore.Defaults.ObservablePoint>
        {
            Values = points,
            Fill = new SolidColorPaint(ChartTheme.Blue500_85),
            Stroke = null,
            GeometrySize = 4
        });

        XAxes.Add(ChartTheme.CreateXAxis(name: InputLabel));
        YAxes.Add(ChartTheme.CreateYAxis(name: "Output"));
    }
}
