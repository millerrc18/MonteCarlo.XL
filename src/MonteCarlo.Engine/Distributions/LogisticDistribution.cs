namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Logistic distribution for modeling growth curves and probability models.
/// Implemented manually (no MathNet wrapper).
/// </summary>
public sealed class LogisticDistribution : IDistribution
{
    private readonly double _mu;
    private readonly double _s;
    private readonly Random _rng;

    /// <summary>
    /// Creates a new Logistic distribution.
    /// </summary>
    /// <param name="mu">Location parameter (μ).</param>
    /// <param name="s">Scale parameter (s). Must be greater than 0.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public LogisticDistribution(double mu, double s, int? seed = null)
    {
        if (s <= 0)
            throw new ArgumentOutOfRangeException(nameof(s), s, "Scale parameter must be greater than 0.");

        _mu = mu;
        _s = s;
        _rng = seed.HasValue ? new Random(seed.Value) : new Random();
    }

    /// <inheritdoc />
    public string Name => "Logistic";

    /// <inheritdoc />
    public double Mean => _mu;

    /// <inheritdoc />
    public double Variance => _s * _s * Math.PI * Math.PI / 3.0;

    /// <inheritdoc />
    public double StdDev => Math.Sqrt(Variance);

    /// <inheritdoc />
    public double Minimum => double.NegativeInfinity;

    /// <inheritdoc />
    public double Maximum => double.PositiveInfinity;

    /// <inheritdoc />
    public double Sample()
    {
        double u = _rng.NextDouble();
        // Avoid log(0) or log(inf)
        while (u <= 0.0 || u >= 1.0)
            u = _rng.NextDouble();
        return _mu + _s * Math.Log(u / (1.0 - u));
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
        double z = (x - _mu) / _s;
        double expNegZ = Math.Exp(-z);
        double denom = _s * (1.0 + expNegZ) * (1.0 + expNegZ);
        return expNegZ / denom;
    }

    /// <inheritdoc />
    public double CDF(double x)
    {
        double z = (x - _mu) / _s;
        return 1.0 / (1.0 + Math.Exp(-z));
    }

    /// <inheritdoc />
    public double Percentile(double p)
    {
        if (p <= 0 || p >= 1)
            throw new ArgumentOutOfRangeException(nameof(p), p, "Percentile must be between 0 and 1 (exclusive).");

        return _mu + _s * Math.Log(p / (1.0 - p));
    }

    /// <inheritdoc />
    public string ParameterSummary() => $"Logistic(μ={_mu}, s={_s})";
}
