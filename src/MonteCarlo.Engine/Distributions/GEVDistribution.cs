namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Generalized Extreme Value (GEV) distribution for modeling extremes
/// (maxima or minima of samples). Encompasses Gumbel (ξ=0), Frechet (ξ>0),
/// and Weibull-type (ξ&lt;0) distributions.
/// Implemented manually (no MathNet wrapper).
/// </summary>
public sealed class GEVDistribution : IDistribution
{
    private readonly double _mu;
    private readonly double _sigma;
    private readonly double _xi;
    private readonly Random _rng;

    /// <summary>
    /// Creates a new GEV distribution.
    /// </summary>
    /// <param name="mu">Location parameter (μ).</param>
    /// <param name="sigma">Scale parameter (σ). Must be greater than 0.</param>
    /// <param name="xi">Shape parameter (ξ). ξ=0 gives Gumbel, ξ>0 Frechet, ξ&lt;0 Weibull-type.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public GEVDistribution(double mu, double sigma, double xi, int? seed = null)
    {
        if (sigma <= 0)
            throw new ArgumentOutOfRangeException(nameof(sigma), sigma, "Scale parameter must be greater than 0.");

        _mu = mu;
        _sigma = sigma;
        _xi = xi;
        _rng = seed.HasValue ? new Random(seed.Value) : new Random();
    }

    /// <inheritdoc />
    public string Name => "GEV";

    /// <inheritdoc />
    public double Mean
    {
        get
        {
            if (Math.Abs(_xi) < 1e-10)
            {
                // Gumbel case: mu + sigma * Euler-Mascheroni constant
                return _mu + _sigma * 0.5772156649;
            }
            if (_xi < 1.0)
            {
                return _mu + _sigma * (MathNet.Numerics.SpecialFunctions.Gamma(1.0 - _xi) - 1.0) / _xi;
            }
            return double.PositiveInfinity;
        }
    }

    /// <inheritdoc />
    public double Variance
    {
        get
        {
            if (Math.Abs(_xi) < 1e-10)
            {
                // Gumbel variance = (pi * sigma)^2 / 6
                return _sigma * _sigma * Math.PI * Math.PI / 6.0;
            }
            if (_xi < 0.5)
            {
                double g1 = MathNet.Numerics.SpecialFunctions.Gamma(1.0 - _xi);
                double g2 = MathNet.Numerics.SpecialFunctions.Gamma(1.0 - 2.0 * _xi);
                return _sigma * _sigma * (g2 - g1 * g1) / (_xi * _xi);
            }
            return double.PositiveInfinity;
        }
    }

    /// <inheritdoc />
    public double StdDev => Math.Sqrt(Variance);

    /// <inheritdoc />
    public double Minimum
    {
        get
        {
            if (_xi > 0) return double.NegativeInfinity;
            if (Math.Abs(_xi) < 1e-10) return double.NegativeInfinity;
            // xi < 0: support is (-inf, mu - sigma/xi)
            return double.NegativeInfinity;
        }
    }

    /// <inheritdoc />
    public double Maximum
    {
        get
        {
            if (_xi < 0) return _mu - _sigma / _xi;
            return double.PositiveInfinity;
        }
    }

    /// <inheritdoc />
    public double Sample()
    {
        double u = _rng.NextDouble();
        while (u <= 0.0 || u >= 1.0)
            u = _rng.NextDouble();
        return Percentile(u);
    }

    /// <inheritdoc />
    public double[] Sample(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(count);
        var samples = new double[count];
        for (int i = 0; i < count; i++)
            samples[i] = Sample();
        return samples;
    }

    /// <inheritdoc />
    public double PDF(double x)
    {
        double t = Tx(x);
        if (double.IsNaN(t) || t <= 0) return 0.0;

        if (Math.Abs(_xi) < 1e-10)
        {
            // Gumbel PDF
            double z = (x - _mu) / _sigma;
            double expZ = Math.Exp(-z);
            return (1.0 / _sigma) * expZ * Math.Exp(-expZ);
        }

        return (1.0 / _sigma) * Math.Pow(t, _xi + 1.0) * Math.Exp(-t);
    }

    /// <inheritdoc />
    public double CDF(double x)
    {
        double t = Tx(x);
        if (double.IsNaN(t)) return _xi > 0 ? 0.0 : 1.0;
        if (t <= 0) return _xi > 0 ? 0.0 : 1.0;
        return Math.Exp(-t);
    }

    /// <inheritdoc />
    public double Percentile(double p)
    {
        if (p <= 0 || p >= 1)
            throw new ArgumentOutOfRangeException(nameof(p), p, "Percentile must be between 0 and 1 (exclusive).");

        double logP = -Math.Log(p);

        if (Math.Abs(_xi) < 1e-10)
        {
            // Gumbel case
            return _mu - _sigma * Math.Log(logP);
        }

        return _mu + _sigma * (Math.Pow(logP, -_xi) - 1.0) / _xi;
    }

    /// <inheritdoc />
    public string ParameterSummary() => $"GEV(μ={_mu}, σ={_sigma}, ξ={_xi})";

    /// <summary>
    /// Helper function t(x) used in PDF and CDF calculations.
    /// </summary>
    private double Tx(double x)
    {
        if (Math.Abs(_xi) < 1e-10)
        {
            double z = (x - _mu) / _sigma;
            return Math.Exp(-z);
        }

        double val = 1.0 + _xi * (x - _mu) / _sigma;
        if (val <= 0) return double.NaN;
        return Math.Pow(val, -1.0 / _xi);
    }
}
