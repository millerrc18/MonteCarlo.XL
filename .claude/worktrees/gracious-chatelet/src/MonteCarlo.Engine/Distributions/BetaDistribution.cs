using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Beta distribution for modeling values bounded in [0, 1], such as
/// percentages, conversion rates, yields, and probabilities.
/// Wraps <see cref="Beta"/> from MathNet.Numerics.
/// </summary>
public sealed class BetaDistribution : IDistribution
{
    private readonly Beta _inner;

    /// <summary>
    /// Creates a new Beta distribution.
    /// </summary>
    /// <param name="alpha">Shape parameter α. Must be greater than 0.</param>
    /// <param name="beta">Shape parameter β. Must be greater than 0.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public BetaDistribution(double alpha, double beta, int? seed = null)
    {
        if (alpha <= 0)
            throw new ArgumentOutOfRangeException(nameof(alpha), alpha, "Alpha must be greater than 0.");
        if (beta <= 0)
            throw new ArgumentOutOfRangeException(nameof(beta), beta, "Beta must be greater than 0.");

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new Beta(alpha, beta, rng);
    }

    /// <inheritdoc />
    public string Name => "Beta";

    /// <inheritdoc />
    public double Mean => _inner.Mean;

    /// <inheritdoc />
    public double Variance => _inner.Variance;

    /// <inheritdoc />
    public double StdDev => _inner.StdDev;

    /// <inheritdoc />
    public double Minimum => 0.0;

    /// <inheritdoc />
    public double Maximum => 1.0;

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
    public double PDF(double x) => _inner.Density(x);

    /// <inheritdoc />
    public double CDF(double x) => _inner.CumulativeDistribution(x);

    /// <inheritdoc />
    public double Percentile(double p) => _inner.InverseCumulativeDistribution(p);

    /// <inheritdoc />
    public string ParameterSummary() => $"Beta(α={_inner.A}, β={_inner.B})";
}
