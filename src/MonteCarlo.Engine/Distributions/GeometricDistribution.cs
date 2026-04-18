namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Geometric distribution for modeling the number of trials until first success.
/// This is a discrete distribution — <see cref="Sample"/> returns integers cast to double.
/// Implemented manually (no MathNet wrapper).
/// </summary>
public sealed class GeometricDistribution : IDistribution
{
    private readonly double _p;
    private readonly Random _rng;

    /// <summary>
    /// Creates a new Geometric distribution.
    /// </summary>
    /// <param name="p">Probability of success per trial. Must be in (0, 1].</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public GeometricDistribution(double p, int? seed = null)
    {
        if (p <= 0 || p > 1)
            throw new ArgumentOutOfRangeException(nameof(p), p, "Probability must be in (0, 1].");

        _p = p;
        _rng = seed.HasValue ? new Random(seed.Value) : new Random();
    }

    /// <inheritdoc />
    public string Name => "Geometric";

    /// <inheritdoc />
    public double Mean => 1.0 / _p;

    /// <inheritdoc />
    public double Variance => (1.0 - _p) / (_p * _p);

    /// <inheritdoc />
    public double StdDev => Math.Sqrt(Variance);

    /// <inheritdoc />
    public double Minimum => 1.0;

    /// <inheritdoc />
    public double Maximum => double.PositiveInfinity;

    /// <inheritdoc />
    public double Sample()
    {
        if (_p >= 1.0) return 1.0;

        double u = _rng.NextDouble();
        while (u <= 0.0 || u >= 1.0)
            u = _rng.NextDouble();
        return Math.Ceiling(Math.Log(1.0 - u) / Math.Log(1.0 - _p));
    }

    /// <inheritdoc />
    public double[] Sample(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(count);
        var samples = new double[count];
        for (int i = 0; i < count; i++)
            samples[i] = Sample();
        return samples;
    }

    /// <inheritdoc />
    public double PDF(double x)
    {
        int k = (int)Math.Round(x);
        if (k < 1 || Math.Abs(x - k) > 1e-10)
            return 0.0;
        // P(X = k) = (1-p)^(k-1) * p
        return Math.Pow(1.0 - _p, k - 1) * _p;
    }

    /// <inheritdoc />
    public double CDF(double x)
    {
        if (x < 1.0) return 0.0;
        int k = (int)Math.Floor(x);
        // CDF(k) = 1 - (1-p)^k
        return 1.0 - Math.Pow(1.0 - _p, k);
    }

    /// <inheritdoc />
    public double Percentile(double probability)
    {
        if (probability < 0 || probability > 1)
            throw new ArgumentOutOfRangeException(nameof(probability), probability, "Percentile must be between 0 and 1.");

        if (probability <= 0.0) return 1.0;
        if (probability >= 1.0) return double.PositiveInfinity;

        if (_p >= 1.0) return 1.0;

        return Math.Ceiling(Math.Log(1.0 - probability) / Math.Log(1.0 - _p));
    }

    /// <inheritdoc />
    public string ParameterSummary() => $"Geometric(p={_p})";
}
