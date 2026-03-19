using MathNet.Numerics.Distributions;
using MathNet.Numerics.Statistics;

namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Computes comprehensive descriptive statistics for a set of simulation output values.
/// All statistics are computed eagerly at construction time.
/// </summary>
public class SummaryStatistics
{
    private static readonly double[] StandardPercentiles = { 0.01, 0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99 };

    /// <summary>
    /// Creates a new SummaryStatistics from raw output values.
    /// </summary>
    /// <param name="values">The simulation output values to analyze.</param>
    public SummaryStatistics(double[] values)
    {
        if (values == null || values.Length == 0)
            throw new ArgumentException("Values array must not be null or empty.", nameof(values));

        Count = values.Length;
        SortedValues = new double[Count];
        Array.Copy(values, SortedValues, Count);
        Array.Sort(SortedValues);

        // Central tendency
        Mean = Statistics.Mean(values);
        Median = InterpolatedPercentile(0.50);

        // Spread
        Min = SortedValues[0];
        Max = SortedValues[Count - 1];
        Range = Max - Min;

        if (Count > 1)
        {
            Variance = Statistics.Variance(values);
            StdDev = Math.Sqrt(Variance);
        }
        else
        {
            Variance = 0;
            StdDev = 0;
        }

        // Shape
        if (Count > 2 && StdDev > 0)
        {
            // Compute skewness: E[(X - mean)^3] / stddev^3
            double sumCubed = 0;
            double sumFourth = 0;
            foreach (double v in values)
            {
                double z = (v - Mean) / StdDev;
                double z2 = z * z;
                sumCubed += z2 * z;
                sumFourth += z2 * z2;
            }
            // Using sample skewness (adjusted)
            double n = Count;
            Skewness = (n / ((n - 1) * (n - 2))) * sumCubed;
            // Excess kurtosis (Fisher definition, Normal = 0)
            Kurtosis = ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * sumFourth
                       - (3.0 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3));
        }
        else
        {
            Skewness = 0;
            Kurtosis = 0;
        }

        // Standard percentiles
        P1 = InterpolatedPercentile(0.01);
        P5 = InterpolatedPercentile(0.05);
        P10 = InterpolatedPercentile(0.10);
        P25 = InterpolatedPercentile(0.25);
        P50 = Median;
        P75 = InterpolatedPercentile(0.75);
        P90 = InterpolatedPercentile(0.90);
        P95 = InterpolatedPercentile(0.95);
        P99 = InterpolatedPercentile(0.99);

        // Mode estimation via histogram binning (Sturges' rule)
        int binCount = Math.Max(1, (int)(1 + 3.322 * Math.Log10(Count)));
        if (Range > 0)
        {
            double binWidth = Range / binCount;
            int bestBin = 0;
            int bestCount = 0;
            var bins = new int[binCount];
            foreach (double v in SortedValues)
            {
                int bin = Math.Min((int)((v - Min) / binWidth), binCount - 1);
                bins[bin]++;
                if (bins[bin] > bestCount)
                {
                    bestCount = bins[bin];
                    bestBin = bin;
                }
            }
            Mode = Min + (bestBin + 0.5) * binWidth;
        }
        else
        {
            Mode = Min;
        }
    }

    /// <summary>Number of values.</summary>
    public int Count { get; }

    /// <summary>The sorted values array (for CDF charting).</summary>
    public double[] SortedValues { get; }

    // Central tendency
    /// <summary>Arithmetic mean.</summary>
    public double Mean { get; }
    /// <summary>Median (P50).</summary>
    public double Median { get; }
    /// <summary>Estimated mode (center of most frequent histogram bin).</summary>
    public double Mode { get; }

    // Spread
    /// <summary>Sample standard deviation.</summary>
    public double StdDev { get; }
    /// <summary>Sample variance.</summary>
    public double Variance { get; }
    /// <summary>Minimum value.</summary>
    public double Min { get; }
    /// <summary>Maximum value.</summary>
    public double Max { get; }
    /// <summary>Max - Min.</summary>
    public double Range { get; }

    // Shape
    /// <summary>Sample skewness (adjusted).</summary>
    public double Skewness { get; }
    /// <summary>Excess kurtosis (Fisher definition, Normal = 0).</summary>
    public double Kurtosis { get; }

    // Standard percentiles
    /// <summary>1st percentile.</summary>
    public double P1 { get; }
    /// <summary>5th percentile.</summary>
    public double P5 { get; }
    /// <summary>10th percentile.</summary>
    public double P10 { get; }
    /// <summary>25th percentile.</summary>
    public double P25 { get; }
    /// <summary>50th percentile (median).</summary>
    public double P50 { get; }
    /// <summary>75th percentile.</summary>
    public double P75 { get; }
    /// <summary>90th percentile.</summary>
    public double P90 { get; }
    /// <summary>95th percentile.</summary>
    public double P95 { get; }
    /// <summary>99th percentile.</summary>
    public double P99 { get; }

    /// <summary>
    /// Computes an arbitrary percentile using linear interpolation
    /// (same method as Excel's PERCENTILE.INC).
    /// </summary>
    /// <param name="p">Percentile as a fraction (0.0 to 1.0).</param>
    public double Percentile(double p)
    {
        if (p < 0.0 || p > 1.0)
            throw new ArgumentOutOfRangeException(nameof(p), "Percentile must be between 0.0 and 1.0.");
        return InterpolatedPercentile(p);
    }

    /// <summary>
    /// Computes the confidence interval for the mean using the t-distribution.
    /// </summary>
    /// <param name="confidence">Confidence level (e.g., 0.95 for 95%).</param>
    public (double Lower, double Upper) MeanConfidenceInterval(double confidence = 0.95)
    {
        if (confidence <= 0 || confidence >= 1)
            throw new ArgumentOutOfRangeException(nameof(confidence));

        if (Count < 2)
            return (Mean, Mean);

        double alpha = 1.0 - confidence;
        double tCritical = StudentT.InvCDF(0, 1, Count - 1, 1.0 - alpha / 2.0);
        double margin = tCritical * StdDev / Math.Sqrt(Count);

        return (Mean - margin, Mean + margin);
    }

    /// <summary>
    /// Probability that a value exceeds the threshold: P(X > threshold).
    /// </summary>
    public double ProbabilityAbove(double threshold)
    {
        return 1.0 - ProbabilityBelow(threshold);
    }

    /// <summary>
    /// Probability that a value is at or below the threshold: P(X &lt;= threshold).
    /// </summary>
    public double ProbabilityBelow(double threshold)
    {
        // Binary search for the count of values <= threshold
        int index = Array.BinarySearch(SortedValues, threshold);
        int countBelow;
        if (index >= 0)
        {
            // Exact match found — find the last occurrence
            while (index < Count - 1 && SortedValues[index + 1] == threshold)
                index++;
            countBelow = index + 1;
        }
        else
        {
            // ~index is the insertion point (count of values less than threshold)
            countBelow = ~index;
        }
        return (double)countBelow / Count;
    }

    /// <summary>
    /// Probability that a value falls between lower and upper: P(lower &lt; X &lt;= upper).
    /// </summary>
    public double ProbabilityBetween(double lower, double upper)
    {
        if (upper < lower)
            throw new ArgumentException("Upper must be >= lower.");
        return ProbabilityBelow(upper) - ProbabilityBelow(lower);
    }

    /// <summary>
    /// Creates histogram data with the specified number of bins.
    /// </summary>
    public HistogramData ToHistogram(int binCount = 50)
    {
        return new HistogramData(SortedValues, binCount);
    }

    private double InterpolatedPercentile(double p)
    {
        if (Count == 1) return SortedValues[0];

        double rank = p * (Count - 1);
        int lower = (int)Math.Floor(rank);
        int upper = (int)Math.Ceiling(rank);

        if (lower == upper || upper >= Count)
            return SortedValues[Math.Min(lower, Count - 1)];

        double fraction = rank - lower;
        return SortedValues[lower] + fraction * (SortedValues[upper] - SortedValues[lower]);
    }
}
