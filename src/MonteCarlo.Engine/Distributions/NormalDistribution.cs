using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Normal (Gaussian) distribution for symmetric uncertainty modeling.
/// Wraps <see cref="Normal"/> from MathNet.Numerics.
/// </summary>
public sealed class NormalDistribution : IDistribution
{
    private readonly Normal _inner;

    /// <summary>
    /// Creates a new Normal distribution.
    /// </summary>
    /// <param name="mean">The mean (μ) of the distribution.</param>
    /// <param name="stdDev">The standard deviation (σ). Must be greater than 0.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public NormalDistribution(double mean, double stdDev, int? seed = null)
    {
        if (stdDev <= 0)
            throw new ArgumentOutOfRangeException(nameof(stdDev), stdDev, "Standard deviation must be greater than 0.");

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new Normal(mean, stdDev, rng);
    }

    /// <inheritdoc />
    public string Name => "Normal";

    /// <inheritdoc />
    public double Mean => _inner.Mean;

    /// <inheritdoc />
    public double Variance => _inner.Variance;

    /// <inheritdoc />
    public double StdDev => _inner.StdDev;

    /// <inheritdoc />
    public double Minimum => double.NegativeInfinity;

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
    public double PDF(double x) => _inner.Density(x);

    /// <inheritdoc />
    public double CDF(double x) => _inner.CumulativeDistribution(x);

    /// <inheritdoc />
    public double Percentile(double p) => _inner.InverseCumulativeDistribution(p);

    /// <inheritdoc />
    public string ParameterSummary() => $"Normal(μ={_inner.Mean}, σ={_inner.StdDev})";
}
