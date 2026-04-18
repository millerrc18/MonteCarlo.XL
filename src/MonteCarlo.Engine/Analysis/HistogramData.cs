namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Pre-binned histogram data ready for charting.
/// </summary>
public class HistogramData
{
    // Stored for KDE computation
    private readonly double[] _sortedValues;
    private readonly double _stdDev;
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

        _sortedValues = sortedValues;

        // Compute standard deviation for KDE bandwidth
        double sum = 0, sumSq = 0;
        for (int i = 0; i < n; i++)
        {
            sum += sortedValues[i];
            sumSq += sortedValues[i] * sortedValues[i];
        }
        double mean = sum / n;
        _stdDev = Math.Sqrt(sumSq / n - mean * mean);

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

    /// <summary>
    /// Computes a Gaussian kernel density estimate over evenly spaced points.
    /// Returns arrays scaled to match the histogram's relative frequency scale.
    /// </summary>
    /// <param name="points">Number of evaluation points from min to max.</param>
    /// <returns>X values and corresponding Y (density * BinWidth) values.</returns>
    public (double[] X, double[] Y) ComputeKDE(int points = 200)
    {
        int n = _sortedValues.Length;
        double min = _sortedValues[0];
        double max = _sortedValues[^1];

        var x = new double[points];
        var y = new double[points];

        // Degenerate case
        if (_stdDev < double.Epsilon || n < 2)
            return (x, y);

        // Silverman's rule of thumb for bandwidth
        double h = 1.06 * _stdDev * Math.Pow(n, -0.2);

        // Subsample for performance: use every kth value
        int subsampleStep = Math.Max(1, n / 5000);
        int subsampleCount = 0;
        for (int i = 0; i < n; i += subsampleStep)
            subsampleCount++;

        double step = (max - min) / (points - 1);

        for (int j = 0; j < points; j++)
        {
            x[j] = min + j * step;
            double density = 0;
            for (int i = 0; i < n; i += subsampleStep)
            {
                double u = (x[j] - _sortedValues[i]) / h;
                // Gaussian kernel: (1/sqrt(2*pi)) * exp(-0.5 * u^2)
                density += Math.Exp(-0.5 * u * u);
            }
            // Normalize: (1 / (subsampleCount * h)) * density * (1/sqrt(2*pi))
            // Then scale by BinWidth to match relative frequency
            y[j] = density / (subsampleCount * h * Math.Sqrt(2.0 * Math.PI)) * BinWidth;
        }

        return (x, y);
    }
}
