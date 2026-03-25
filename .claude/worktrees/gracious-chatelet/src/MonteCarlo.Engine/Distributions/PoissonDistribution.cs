using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Poisson distribution for modeling the count of events in a fixed interval.
/// Common uses include defect counts, arrival counts, and insurance claims.
/// This is a discrete distribution — <see cref="Sample"/> returns integers cast to double.
/// Wraps <see cref="Poisson"/> from MathNet.Numerics.
/// </summary>
public sealed class PoissonDistribution : IDistribution
{
    private readonly Poisson _inner;

    /// <summary>
    /// Creates a new Poisson distribution.
    /// </summary>
    /// <param name="lambda">Expected number of events λ. Must be greater than 0.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public PoissonDistribution(double lambda, int? seed = null)
    {
        if (lambda <= 0)
            throw new ArgumentOutOfRangeException(nameof(lambda), lambda, "Lambda must be greater than 0.");

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new Poisson(lambda, rng);
    }

    /// <inheritdoc />
    public string Name => "Poisson";

    /// <inheritdoc />
    public double Mean => _inner.Mean;

    /// <inheritdoc />
    public double Variance => _inner.Variance;

    /// <inheritdoc />
    public double StdDev => _inner.StdDev;

    /// <inheritdoc />
    public double Minimum => 0.0;

    /// <inheritdoc />
    public double Maximum => double.PositiveInfinity;

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
        // Poisson PMF is only defined for non-negative integers
        int k = (int)Math.Round(x);
        if (k < 0 || Math.Abs(x - k) > 1e-10)
            return 0.0;
        return _inner.Probability(k);
    }

    /// <inheritdoc />
    public double CDF(double x)
    {
        if (x < 0) return 0.0;
        return _inner.CumulativeDistribution(x);
    }

    /// <inheritdoc />
    public double Percentile(double p)
    {
        if (p < 0 || p > 1)
            throw new ArgumentOutOfRangeException(nameof(p), p, "Percentile must be between 0 and 1.");

        // For discrete distributions, return the smallest integer k where CDF(k) >= p
        // Math.NET's Poisson doesn't have InverseCumulativeDistribution,
        // so we search manually.
        if (p <= 0.0) return 0.0;
        if (p >= 1.0) return double.PositiveInfinity;

        int k = 0;
        while (_inner.CumulativeDistribution(k) < p)
        {
            k++;
            if (k > _inner.Lambda * 20 + 100) break; // safety limit
        }
        return k;
    }

    /// <inheritdoc />
    public string ParameterSummary() => $"Poisson(λ={_inner.Lambda})";
}
