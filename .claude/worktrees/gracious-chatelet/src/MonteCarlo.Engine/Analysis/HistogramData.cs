namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Pre-binned histogram data ready for charting.
/// </summary>
public class HistogramData
{
    /// <summary>
    /// Bin edge values. Length = binCount + 1.
    /// </summary>
    public double[] BinEdges { get; }

    /// <summary>
    /// Center value of each bin. Length = binCount.
    /// </summary>
    public double[] BinCenters { get; }

    /// <summary>
    /// Number of samples in each bin. Length = binCount.
    /// </summary>
    public int[] Frequencies { get; }

    /// <summary>
    /// Frequencies divided by total count. Length = binCount.
    /// </summary>
    public double[] RelativeFrequencies { get; }

    /// <summary>
    /// Width of each bin (uniform).
    /// </summary>
    public double BinWidth { get; }

    /// <summary>
    /// Creates histogram data from sorted sample values.
    /// </summary>
    /// <param name="sortedValues">The sample values, already sorted ascending.</param>
    /// <param name="binCount">Number of bins to create.</param>
    public HistogramData(double[] sortedValues, int binCount = 50)
    {
        if (sortedValues == null || sortedValues.Length == 0)
            throw new ArgumentException("Values array must not be null or empty.", nameof(sortedValues));

        if (binCount <= 0)
            throw new ArgumentException("Bin count must be positive.", nameof(binCount));

        int n = sortedValues.Length;
        double min = sortedValues[0];
        double max = sortedValues[n - 1];

        // Edge case: all values identical
        if (Math.Abs(max - min) < double.Epsilon)
        {
            binCount = 1;
            BinEdges = new[] { min - 0.5, max + 0.5 };
            BinCenters = new[] { min };
            Frequencies = new[] { n };
            RelativeFrequencies = new[] { 1.0 };
            BinWidth = 1.0;
            return;
        }

        BinWidth = (max - min) / binCount;
        BinEdges = new double[binCount + 1];
        BinCenters = new double[binCount];
        Frequencies = new int[binCount];

        for (int i = 0; i <= binCount; i++)
            BinEdges[i] = min + i * BinWidth;

        // Ensure last edge exactly equals max (avoid floating point drift)
        BinEdges[binCount] = max;

        for (int i = 0; i < binCount; i++)
            BinCenters[i] = (BinEdges[i] + BinEdges[i + 1]) / 2.0;

        // Assign values to bins using sorted order for efficiency
        foreach (double v in sortedValues)
        {
            int bin = (int)((v - min) / BinWidth);
            // Clamp to valid range (last value lands exactly on max)
            if (bin >= binCount) bin = binCount - 1;
            if (bin < 0) bin = 0;
            Frequencies[bin]++;
        }

        RelativeFrequencies = new double[binCount];
        for (int i = 0; i < binCount; i++)
            RelativeFrequencies[i] = (double)Frequencies[i] / n;
    }
}
