using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Binomial distribution for modeling the number of successes in n independent trials.
/// This is a discrete distribution — <see cref="Sample"/> returns integers cast to double.
/// Wraps <see cref="Binomial"/> from MathNet.Numerics.
/// </summary>
public sealed class BinomialDistribution : IDistribution
{
    private readonly Binomial _inner;
    private readonly int _n;
    private readonly double _p;

    /// <summary>
    /// Creates a new Binomial distribution.
    /// </summary>
    /// <param name="n">Number of trials. Must be greater than 0.</param>
    /// <param name="p">Probability of success per trial. Must be between 0 and 1 inclusive.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public BinomialDistribution(int n, double p, int? seed = null)
    {
        if (n <= 0)
            throw new ArgumentOutOfRangeException(nameof(n), n, "Number of trials must be greater than 0.");
        if (p < 0 || p > 1)
            throw new ArgumentOutOfRangeException(nameof(p), p, "Probability must be between 0 and 1.");

        _n = n;
        _p = p;
        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        // MathNet Binomial constructor is (p, n, rng)
        _inner = new Binomial(p, n, rng);
    }

    /// <inheritdoc />
    public string Name => "Binomial";

    /// <inheritdoc />
    public double Mean => _inner.Mean;

    /// <inheritdoc />
    public double Variance => _inner.Variance;

    /// <inheritdoc />
    public double StdDev => _inner.StdDev;

    /// <inheritdoc />
    public double Minimum => 0.0;

    /// <inheritdoc />
    public double Maximum => _n;

    /// <inheritdoc />
    public double Sample() => _inner.Sample();

    /// <inheritdoc />
    public double[] Sample(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(count);
        var samples = new double[count];
        for (int i = 0; i < count; i++)
            samples[i] = _inner.Sample();
        return samples;
    }

    /// <inheritdoc />
    public double PDF(double x)
    {
        int k = (int)Math.Round(x);
        if (k < 0 || k > _n || Math.Abs(x - k) > 1e-10)
            return 0.0;
        return _inner.Probability(k);
    }

    /// <inheritdoc />
    public double CDF(double x)
    {
        if (x < 0) return 0.0;
        if (x >= _n) return 1.0;
        return _inner.CumulativeDistribution(x);
    }

    /// <inheritdoc />
    public double Percentile(double p)
    {
        if (p < 0 || p > 1)
            throw new ArgumentOutOfRangeException(nameof(p), p, "Percentile must be between 0 and 1.");

        if (p <= 0.0) return 0.0;
        if (p >= 1.0) return _n;

        // Search for smallest k where CDF(k) >= p
        for (int k = 0; k <= _n; k++)
        {
            if (_inner.CumulativeDistribution(k) >= p)
                return k;
        }
        return _n;
    }

    /// <inheritdoc />
    public string ParameterSummary() => $"Binomial(n={_n}, p={_p})";
}
