using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace MonteCarlo.UI.Views;

/// <summary>
/// A mini distribution preview sparkline using a WPF Polyline.
/// Shows the PDF shape — no axes or labels, just the curve.
/// </summary>
public partial class DistributionPreviewControl : UserControl
{
    public static readonly DependencyProperty PointsProperty =
        DependencyProperty.Register(
            nameof(Points),
            typeof(IReadOnlyList<(double X, double Y)>),
            typeof(DistributionPreviewControl),
            new PropertyMetadata(null, OnPointsChanged));

    /// <summary>The PDF sample points to render.</summary>
    public IReadOnlyList<(double X, double Y)>? Points
    {
        get => (IReadOnlyList<(double X, double Y)>?)GetValue(PointsProperty);
        set => SetValue(PointsProperty, value);
    }

    public DistributionPreviewControl()
    {
        InitializeComponent();
        SizeChanged += (_, _) => Render();
    }

    private static void OnPointsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((DistributionPreviewControl)d).Render();
    }

    private void Render()
    {
        PreviewCanvas.Children.Clear();
        var points = Points;
        if (points == null || points.Count < 2) return;

        double w = PreviewCanvas.ActualWidth;
        double h = PreviewCanvas.ActualHeight;
        if (w <= 0 || h <= 0) return;

        double xMin = points[0].X;
        double xMax = points[^1].X;
        double yMax = 0;
        foreach (var pt in points)
            if (pt.Y > yMax) yMax = pt.Y;

        if (xMax <= xMin || yMax <= 0) return;

        double xScale = w / (xMax - xMin);
        double yScale = (h - 4) / yMax; // 2px padding top/bottom
        const double yPad = 2;

        // Build polyline for the fill
        var fillPoints = new PointCollection();
        fillPoints.Add(new Point(0, h));

        var linePoints = new PointCollection();

        foreach (var pt in points)
        {
            double px = (pt.X - xMin) * xScale;
            double py = h - yPad - pt.Y * yScale;
            fillPoints.Add(new Point(px, py));
            linePoints.Add(new Point(px, py));
        }

        fillPoints.Add(new Point(w, h));

        // Fill polygon
        var fill = new Polygon
        {
            Points = fillPoints,
            Fill = new SolidColorBrush(Color.FromArgb(30, 59, 130, 246)), // Blue500 at 12% opacity
            StrokeThickness = 0
        };
        PreviewCanvas.Children.Add(fill);

        // Line
        var line = new Polyline
        {
            Points = linePoints,
            Stroke = new SolidColorBrush(Color.FromRgb(59, 130, 246)), // Blue500
            StrokeThickness = 1.5,
            StrokeLineJoin = PenLineJoin.Round
        };
        PreviewCanvas.Children.Add(line);
    }
}
