using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Renders a lightweight histogram plus fitted distribution curve overlay.
/// </summary>
public partial class DistributionFitPreviewControl : UserControl
{
    public static readonly DependencyProperty HistogramBarsProperty =
        DependencyProperty.Register(
            nameof(HistogramBars),
            typeof(IReadOnlyList<DistributionFitHistogramBar>),
            typeof(DistributionFitPreviewControl),
            new PropertyMetadata(null, OnPreviewDataChanged));

    public static readonly DependencyProperty CurvePointsProperty =
        DependencyProperty.Register(
            nameof(CurvePoints),
            typeof(IReadOnlyList<(double X, double Y)>),
            typeof(DistributionFitPreviewControl),
            new PropertyMetadata(null, OnPreviewDataChanged));

    public IReadOnlyList<DistributionFitHistogramBar>? HistogramBars
    {
        get => (IReadOnlyList<DistributionFitHistogramBar>?)GetValue(HistogramBarsProperty);
        set => SetValue(HistogramBarsProperty, value);
    }

    public IReadOnlyList<(double X, double Y)>? CurvePoints
    {
        get => (IReadOnlyList<(double X, double Y)>?)GetValue(CurvePointsProperty);
        set => SetValue(CurvePointsProperty, value);
    }

    public DistributionFitPreviewControl()
    {
        InitializeComponent();
        SizeChanged += (_, _) => Render();
    }

    private static void OnPreviewDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((DistributionFitPreviewControl)d).Render();
    }

    private void Render()
    {
        PreviewCanvas.Children.Clear();

        var bars = HistogramBars;
        var curve = CurvePoints;
        if ((bars == null || bars.Count == 0) && (curve == null || curve.Count < 2))
            return;

        var width = PreviewCanvas.ActualWidth;
        var height = PreviewCanvas.ActualHeight;
        if (width <= 0 || height <= 0)
            return;

        var xMin = double.PositiveInfinity;
        var xMax = double.NegativeInfinity;
        var yMax = 0.0;

        if (bars != null)
        {
            foreach (var bar in bars)
            {
                xMin = Math.Min(xMin, bar.Center - bar.HalfWidth);
                xMax = Math.Max(xMax, bar.Center + bar.HalfWidth);
                yMax = Math.Max(yMax, bar.Height);
            }
        }

        if (curve != null)
        {
            foreach (var point in curve)
            {
                xMin = Math.Min(xMin, point.X);
                xMax = Math.Max(xMax, point.X);
                yMax = Math.Max(yMax, point.Y);
            }
        }

        if (!double.IsFinite(xMin) || !double.IsFinite(xMax) || xMax <= xMin || yMax <= 0)
            return;

        const double padLeft = 6;
        const double padRight = 6;
        const double padTop = 4;
        const double padBottom = 6;
        var plotWidth = Math.Max(1, width - padLeft - padRight);
        var plotHeight = Math.Max(1, height - padTop - padBottom);

        double ScaleX(double x) => padLeft + ((x - xMin) / (xMax - xMin)) * plotWidth;
        double ScaleY(double y) => height - padBottom - (y / yMax) * plotHeight;

        if (bars != null)
        {
            foreach (var bar in bars)
            {
                var left = ScaleX(bar.Center - bar.HalfWidth);
                var right = ScaleX(bar.Center + bar.HalfWidth);
                var top = ScaleY(bar.Height);

                var rect = new Rectangle
                {
                    Width = Math.Max(1, right - left),
                    Height = Math.Max(1, height - padBottom - top),
                    Fill = new SolidColorBrush(Color.FromArgb(70, 148, 163, 184)),
                    Stroke = new SolidColorBrush(Color.FromArgb(90, 148, 163, 184)),
                    StrokeThickness = 0.5
                };

                Canvas.SetLeft(rect, left);
                Canvas.SetTop(rect, top);
                PreviewCanvas.Children.Add(rect);
            }
        }

        if (curve != null && curve.Count >= 2)
        {
            var fillPoints = new PointCollection();
            var linePoints = new PointCollection();
            fillPoints.Add(new Point(ScaleX(curve[0].X), height - padBottom));

            foreach (var point in curve)
            {
                var scaled = new Point(ScaleX(point.X), ScaleY(point.Y));
                fillPoints.Add(scaled);
                linePoints.Add(scaled);
            }

            fillPoints.Add(new Point(ScaleX(curve[^1].X), height - padBottom));

            var fill = new Polygon
            {
                Points = fillPoints,
                Fill = new SolidColorBrush(Color.FromArgb(26, 59, 130, 246)),
                StrokeThickness = 0
            };
            PreviewCanvas.Children.Add(fill);

            var line = new Polyline
            {
                Points = linePoints,
                Stroke = new SolidColorBrush(Color.FromRgb(59, 130, 246)),
                StrokeThickness = 1.5,
                StrokeLineJoin = PenLineJoin.Round
            };
            PreviewCanvas.Children.Add(line);
        }
    }
}
