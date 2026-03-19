using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// PERT distribution for smoother expert estimates. Implemented as a scaled Beta distribution.
/// Produces less extreme tails than the Triangular distribution for the same min/mode/max.
/// </summary>
public sealed class PERTDistribution : IDistribution
{
    private readonly Beta _beta;
    private readonly double _min;
    private readonly double _mode;
    private readonly double _max;
    private readonly double _lambda;
    private readonly double _alpha;
    private readonly double _betaParam;

    /// <summary>
    /// Creates a new PERT distribution.
    /// </summary>
    /// <param name="min">The minimum value.</param>
    /// <param name="mode">The most likely value.</param>
    /// <param name="max">The maximum value.</param>
    /// <param name="lambda">Shape parameter controlling peakedness. Default is 4.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public PERTDistribution(double min, double mode, double max, double lambda = 4.0, int? seed = null)
    {
        if (min >= max)
            throw new ArgumentException($"Minimum ({min}) must be less than maximum ({max}).");
        if (mode < min || mode > max)
            throw new ArgumentOutOfRangeException(nameof(mode), mode, $"Mode must be between min ({min}) and max ({max}).");
        if (lambda <= 0)
            throw new ArgumentOutOfRangeException(nameof(lambda), lambda, "Lambda must be greater than 0.");

        _min = min;
        _mode = mode;
        _max = max;
        _lambda = lambda;

        double range = max - min;
        _alpha = 1.0 + lambda * (mode - min) / range;
        _betaParam = 1.0 + lambda * (max - mode) / range;

        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _beta = new Beta(_alpha, _betaParam, rng);
    }

    /// <inheritdoc />
    public string Name => "PERT";

    /// <inheritdoc />
    public double Mean => (_min + _lambda * _mode + _max) / (_lambda + 2.0);

    /// <inheritdoc />
    public double Variance
    {
        get
        {
            double mean = Mean;
            double range = _max - _min;
            // Variance of scaled Beta: range^2 * Var(Beta)
            return range * range * _beta.Variance;
        }
    }

    /// <inheritdoc />
    public double StdDev => Math.Sqrt(Variance);

    /// <inheritdoc />
    public double Minimum => _min;

    /// <inheritdoc />
    public double Maximum => _max;

    /// <inheritdoc />
    public double Sample() => _min + _beta.Sample() * (_max - _min);

    /// <inheritdoc />
    public double[] Sample(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(count);
        var samples = new double[count];
        double range = _max - _min;
        for (int i = 0; i < count; i++)
            samples[i] = _min + _beta.Sample() * range;
        return samples;
    }

    /// <inheritdoc />
    public double PDF(double x)
    {
        if (x < _min || x > _max)
            return 0.0;

        double range = _max - _min;
        // Transform x to [0,1] for Beta PDF, then divide by range (Jacobian)
        double t = (x - _min) / range;
        return _beta.Density(t) / range;
    }

    /// <inheritdoc />
    public double CDF(double x)
    {
        if (x <= _min) return 0.0;
        if (x >= _max) return 1.0;

        double t = (x - _min) / (_max - _min);
        return _beta.CumulativeDistribution(t);
    }

    /// <inheritdoc />
    public double Percentile(double p)
    {
        double t = _beta.InverseCumulativeDistribution(p);
        return _min + t * (_max - _min);
    }

    /// <inheritdoc />
    public string ParameterSummary() => $"PERT(min={_min}, mode={_mode}, max={_max}, λ={_lambda})";
}
