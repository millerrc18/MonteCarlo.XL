using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Weibull distribution for failure rates, time-to-event, reliability analysis,
/// and wind speed modeling. Wraps <see cref="Weibull"/> from MathNet.Numerics.
/// </summary>
public sealed class WeibullDistribution : IDistribution
{
    private readonly Weibull _inner;

    /// <summary>
    /// Creates a new Weibull distribution.
    /// </summary>
    /// <param name="shape">Shape parameter k. Must be greater than 0.</param>
    /// <param name="scale">Scale parameter λ. Must be greater than 0.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public WeibullDistribution(double shape, double scale, int? seed = null)
    {
        if (shape <= 0)
            throw new ArgumentOutOfRangeException(nameof(shape), shape, "Shape must be greater than 0.");
        if (scale <= 0)
            throw new ArgumentOutOfRangeException(nameof(scale), scale, "Scale must be greater than 0.");

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new Weibull(shape, scale, rng);
    }

    /// <inheritdoc />
    public string Name => "Weibull";

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
    public double Percentile(double p)
    {
        if (p < 0.0 || p > 1.0)
            throw new ArgumentOutOfRangeException(nameof(p), "Percentile must be between 0.0 and 1.0.");
        if (p == 0.0) return 0.0;
        if (p == 1.0) return double.PositiveInfinity;
        // Analytical inverse CDF: x = scale * (-ln(1-p))^(1/shape)
        return _inner.Scale * Math.Pow(-Math.Log(1.0 - p), 1.0 / _inner.Shape);
    }

    /// <inheritdoc />
    public string ParameterSummary() => $"Weibull(k={_inner.Shape}, λ={_inner.Scale})";
}
