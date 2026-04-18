using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Gamma distribution for modeling waiting times, insurance claims, and rainfall.
/// Wraps <see cref="Gamma"/> from MathNet.Numerics.
/// </summary>
public sealed class GammaDistribution : IDistribution
{
    private readonly Gamma _inner;
    private readonly double _shape;
    private readonly double _rate;

    /// <summary>
    /// Creates a new Gamma distribution.
    /// </summary>
    /// <param name="shape">Shape parameter (α). Must be greater than 0.</param>
    /// <param name="rate">Rate parameter (β). Must be greater than 0.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public GammaDistribution(double shape, double rate, int? seed = null)
    {
        if (shape <= 0)
            throw new ArgumentOutOfRangeException(nameof(shape), shape, "Shape must be greater than 0.");
        if (rate <= 0)
            throw new ArgumentOutOfRangeException(nameof(rate), rate, "Rate must be greater than 0.");

        _shape = shape;
        _rate = rate;
        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new Gamma(shape, rate, rng);
    }

    /// <inheritdoc />
    public string Name => "Gamma";

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
    public string ParameterSummary() => $"Gamma(shape={_shape}, rate={_rate})";
}
