using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Triangular distribution for expert estimates with min/mode/max.
/// Wraps <see cref="Triangular"/> from MathNet.Numerics.
/// </summary>
public sealed class TriangularDistribution : IDistribution
{
    private readonly Triangular _inner;
    private readonly double _min;
    private readonly double _mode;
    private readonly double _max;

    /// <summary>
    /// Creates a new Triangular distribution.
    /// </summary>
    /// <param name="min">The minimum value. Must be less than or equal to mode.</param>
    /// <param name="mode">The most likely value. Must be between min and max.</param>
    /// <param name="max">The maximum value. Must be greater than or equal to mode.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public TriangularDistribution(double min, double mode, double max, int? seed = null)
    {
        if (min >= max)
            throw new ArgumentException($"Minimum ({min}) must be less than maximum ({max}).");
        if (mode < min || mode > max)
            throw new ArgumentOutOfRangeException(nameof(mode), mode, $"Mode must be between min ({min}) and max ({max}).");

        _min = min;
        _mode = mode;
        _max = max;

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new Triangular(min, max, mode, rng);
    }

    /// <inheritdoc />
    public string Name => "Triangular";

    /// <inheritdoc />
    public double Mean => (_min + _mode + _max) / 3.0;

    /// <inheritdoc />
    public double Variance => (_min * _min + _mode * _mode + _max * _max
                               - _min * _mode - _min * _max - _mode * _max) / 18.0;

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
    public string ParameterSummary() => $"Triangular(min={_min}, mode={_mode}, max={_max})";
}
