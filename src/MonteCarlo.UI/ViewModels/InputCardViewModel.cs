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
    public string Label { get; }

    /// <summary>Distribution type name (e.g., "Normal").</summary>
    public string DistributionName { get; }

    /// <summary>Human-readable parameter summary (e.g., "Normal(μ=100, σ=10)").</summary>
    public string ParameterSummary { get; }

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
        Label = label;
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
}
