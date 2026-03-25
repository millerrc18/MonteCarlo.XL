using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Continuous uniform distribution where all values in a range are equally likely.
/// Wraps <see cref="ContinuousUniform"/> from MathNet.Numerics.
/// </summary>
public sealed class UniformDistribution : IDistribution
{
    private readonly ContinuousUniform _inner;
    private readonly double _min;
    private readonly double _max;

    /// <summary>
    /// Creates a new Uniform distribution.
    /// </summary>
    /// <param name="min">The minimum value.</param>
    /// <param name="max">The maximum value. Must be greater than min.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public UniformDistribution(double min, double max, int? seed = null)
    {
        if (min >= max)
            throw new ArgumentException($"Minimum ({min}) must be less than maximum ({max}).");

        _min = min;
        _max = max;

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new ContinuousUniform(min, max, rng);
    }

    /// <inheritdoc />
    public string Name => "Uniform";

    /// <inheritdoc />
    public double Mean => (_min + _max) / 2.0;

    /// <inheritdoc />
    public double Variance => (_max - _min) * (_max - _min) / 12.0;

    /// <inheritdoc />
    public double StdDev => Math.Sqrt(Variance);

    /// <inheritdoc />
    public double Minimum => _min;

    /// <inheritdoc />
    public double Maximum => _max;

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
    public string ParameterSummary() => $"Uniform(min={_min}, max={_max})";
}
