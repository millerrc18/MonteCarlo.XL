using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Lognormal distribution for right-skewed positive quantities (revenue, prices).
/// Wraps <see cref="LogNormal"/> from MathNet.Numerics.
/// Parameters are μ and σ of the underlying normal distribution.
/// </summary>
public sealed class LognormalDistribution : IDistribution
{
    private readonly LogNormal _inner;
    private readonly double _mu;
    private readonly double _sigma;

    /// <summary>
    /// Creates a new Lognormal distribution.
    /// </summary>
    /// <param name="mu">The mean (μ) of the underlying normal distribution.</param>
    /// <param name="sigma">The standard deviation (σ) of the underlying normal. Must be greater than 0.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public LognormalDistribution(double mu, double sigma, int? seed = null)
    {
        if (sigma <= 0)
            throw new ArgumentOutOfRangeException(nameof(sigma), sigma, "Sigma must be greater than 0.");

        _mu = mu;
        _sigma = sigma;

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _inner = new LogNormal(mu, sigma, rng);
    }

    /// <inheritdoc />
    public string Name => "Lognormal";

    /// <inheritdoc />
    public double Mean => Math.Exp(_mu + _sigma * _sigma / 2.0);

    /// <inheritdoc />
    public double Variance
    {
        get
        {
            double s2 = _sigma * _sigma;
            return (Math.Exp(s2) - 1.0) * Math.Exp(2.0 * _mu + s2);
        }
    }

    /// <inheritdoc />
    public double StdDev => Math.Sqrt(Variance);

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
    public string ParameterSummary() => $"Lognormal(μ={_mu}, σ={_sigma})";
}
