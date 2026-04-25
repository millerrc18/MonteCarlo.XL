using CommunityToolkit.Mvvm.ComponentModel;
using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// View model for a configured simulation input displayed as a card.
/// </summary>
public partial class InputCardViewModel : ObservableObject
{
    /// <summary>Cell address in A1 notation (e.g., "B4").</summary>
    public string CellAddress { get; }

    /// <summary>Worksheet name.</summary>
    public string SheetName { get; }

    /// <summary>Human-readable label.</summary>
    [ObservableProperty]
    private string _label;

    /// <summary>Distribution type name (e.g., "Normal").</summary>
    public string DistributionName { get; }

    /// <summary>Human-readable parameter summary (e.g., "Normal(μ=100, σ=10)").</summary>
    public string ParameterSummary { get; }

    /// <summary>Distribution mean formatted for manager tables.</summary>
    public string MeanSummary => FormatFinite(Distribution.Mean);

    /// <summary>Distribution mean as a numeric value for table sorting.</summary>
    public double MeanValue => Distribution.Mean;

    /// <summary>Distribution 5th percentile formatted for manager tables.</summary>
    public string P5Summary => FormatFinite(Distribution.Percentile(0.05));

    /// <summary>Distribution 5th percentile as a numeric value for table sorting.</summary>
    public double P5Value => Distribution.Percentile(0.05);

    /// <summary>Distribution 95th percentile formatted for manager tables.</summary>
    public string P95Summary => FormatFinite(Distribution.Percentile(0.95));

    /// <summary>Distribution 95th percentile as a numeric value for table sorting.</summary>
    public double P95Value => Distribution.Percentile(0.95);

    /// <summary>Full cell reference (e.g., "Sheet1!B4").</summary>
    public string FullReference => $"{SheetName}!{CellAddress}";

    /// <summary>The distribution instance for preview rendering.</summary>
    public IDistribution Distribution { get; }

    /// <summary>Points for the mini distribution preview curve.</summary>
    public IReadOnlyList<(double X, double Y)> PreviewPoints { get; }

    /// <summary>Raw parameter dictionary for persistence.</summary>
    public Dictionary<string, double> Parameters { get; }

    public InputCardViewModel(
        string sheetName,
        string cellAddress,
        string label,
        string distributionName,
        Dictionary<string, double> parameters,
        IDistribution distribution)
    {
        SheetName = sheetName;
        CellAddress = cellAddress;
        _label = label;
        DistributionName = distributionName;
        Parameters = parameters;
        Distribution = distribution;
        ParameterSummary = distribution.ParameterSummary();
        PreviewPoints = ComputePreviewPointsStatic(distribution);
    }

    /// <summary>
    /// Computes PDF preview points for a distribution (used by editor preview too).
    /// </summary>
    public static IReadOnlyList<(double X, double Y)> ComputePreviewPointsStatic(IDistribution dist)
    {
        if (IsDiscreteDistribution(dist))
            return ComputeDiscretePreviewPoints(dist);

        const int pointCount = 50;
        var points = new List<(double X, double Y)>(pointCount);

        // Use P1 to P99 as the range to avoid infinite tails
        double xMin = dist.Percentile(0.01);
        double xMax = dist.Percentile(0.99);

        if (xMax <= xMin)
        {
            xMin = dist.Mean - 3 * Math.Max(dist.StdDev, 1);
            xMax = dist.Mean + 3 * Math.Max(dist.StdDev, 1);
        }

        double step = (xMax - xMin) / (pointCount - 1);

        for (int i = 0; i < pointCount; i++)
        {
            double x = xMin + i * step;
            double y = dist.PDF(x);
            if (double.IsNaN(y) || double.IsInfinity(y))
                y = 0;
            points.Add((x, y));
        }

        return points;
    }

    private static bool IsDiscreteDistribution(IDistribution dist)
    {
        return dist.Name.Equals("Binomial", StringComparison.OrdinalIgnoreCase)
            || dist.Name.Equals("Geometric", StringComparison.OrdinalIgnoreCase)
            || dist.Name.Equals("Poisson", StringComparison.OrdinalIgnoreCase)
            || dist.Name.Equals("Discrete", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<(double X, double Y)> ComputeDiscretePreviewPoints(IDistribution dist)
    {
        const int maxBars = 80;
        var values = GetDiscretePreviewValues(dist, maxBars);
        if (values.Count == 0)
            return Array.Empty<(double X, double Y)>();

        var points = new List<(double X, double Y)>(values.Count * 4);
        var width = ComputeBarWidth(values);

        foreach (var x in values)
        {
            var y = dist.PDF(x);
            if (double.IsNaN(y) || double.IsInfinity(y) || y < 0)
                y = 0;

            points.Add((x - width, 0));
            points.Add((x - width, y));
            points.Add((x + width, y));
            points.Add((x + width, 0));
        }

        return points;
    }

    private static List<double> GetDiscretePreviewValues(IDistribution dist, int maxBars)
    {
        if (dist.Name.Equals("Binomial", StringComparison.OrdinalIgnoreCase)
            || dist.Name.Equals("Geometric", StringComparison.OrdinalIgnoreCase)
            || dist.Name.Equals("Poisson", StringComparison.OrdinalIgnoreCase))
        {
            var start = SafeFinite(dist.Percentile(0.01), dist.Minimum);
            var end = SafeFinite(dist.Percentile(0.99), dist.Mean + 3 * Math.Max(dist.StdDev, 1));

            if (double.IsFinite(dist.Minimum))
                start = Math.Max(start, dist.Minimum);
            if (double.IsFinite(dist.Maximum))
                end = Math.Min(end, dist.Maximum);

            var kMin = (int)Math.Floor(start);
            var kMax = (int)Math.Ceiling(end);
            if (kMax < kMin)
                kMax = kMin;

            var range = kMax - kMin + 1;
            var step = Math.Max(1, (int)Math.Ceiling(range / (double)maxBars));
            var values = new List<double>();

            for (var k = kMin; k <= kMax; k += step)
                values.Add(k);

            if (values.Count == 0 || Math.Abs(values[^1] - kMax) > 1e-9)
                values.Add(kMax);

            return values;
        }

        return GetFiniteDiscreteValuesFromPercentiles(dist, maxBars);
    }

    private static List<double> GetFiniteDiscreteValuesFromPercentiles(IDistribution dist, int maxBars)
    {
        var values = new List<double>();
        for (var i = 1; i < 100 && values.Count < maxBars; i++)
        {
            var x = dist.Percentile(i / 100.0);
            if (!double.IsFinite(x))
                continue;

            if (!values.Any(existing => Math.Abs(existing - x) < 1e-9))
                values.Add(x);
        }

        values.Sort();
        return values;
    }

    private static double ComputeBarWidth(IReadOnlyList<double> values)
    {
        if (values.Count < 2)
            return 0.45;

        var minGap = double.PositiveInfinity;
        for (var i = 1; i < values.Count; i++)
        {
            var gap = Math.Abs(values[i] - values[i - 1]);
            if (gap > 0 && gap < minGap)
                minGap = gap;
        }

        return double.IsFinite(minGap) ? minGap * 0.4 : 0.45;
    }

    private static double SafeFinite(double value, double fallback)
    {
        if (double.IsFinite(value))
            return value;

        return double.IsFinite(fallback) ? fallback : 0;
    }

    private static string FormatFinite(double value)
    {
        return double.IsFinite(value)
            ? value.ToString("G5", System.Globalization.CultureInfo.InvariantCulture)
            : "n/a";
    }
}
