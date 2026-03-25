using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Exponential distribution for modeling time between events in a Poisson process.
/// Commonly used for arrival times, failure intervals, and memoryless processes.
/// Wraps <see cref="Exponential"/> from MathNet.Numerics.
/// </summary>
public sealed class ExponentialDistribution : IDistribution
{
    private readonly Exponential _inner;

    /// <summary>
    /// Creates a new Exponential distribution.
    /// </summary>
    /// <param name="rate">Rate parameter λ. Must be greater than 0. Mean = 1/λ.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public ExponentialDistribution(double rate, int? seed = null)
    {
        if (rate <= 0)
            throw new ArgumentOutOfRangeException(nameof(rate), rate, "Rate must be greater than 0.");

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new Exponential(rate, rng);
    }

    /// <inheritdoc />
    public string Name => "Exponential";

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
    public double PDF(double x) => _inner.Density(x);

    /// <inheritdoc />
    public double CDF(double x) => _inner.CumulativeDistribution(x);

    /// <inheritdoc />
    public double Percentile(double p) => _inner.InverseCumulativeDistribution(p);

    /// <inheritdoc />
    public string ParameterSummary() => $"Exponential(λ={_inner.Rate})";
}
