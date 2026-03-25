using MathNet.Numerics.Distributions;

namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Discrete distribution modeling a finite set of outcomes with associated probabilities.
/// Uses <see cref="Categorical"/> from MathNet.Numerics internally.
/// </summary>
public sealed class DiscreteDistribution : IDistribution
{
    private readonly Categorical _categorical;
    private readonly double[] _values;
    private readonly double[] _probabilities;

    /// <summary>
    /// Creates a new Discrete distribution.
    /// </summary>
    /// <param name="values">The possible outcome values.</param>
    /// <param name="probabilities">The probability of each outcome. Must sum to 1.0 (within tolerance).</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    public DiscreteDistribution(double[] values, double[] probabilities, int? seed = null)
    {
        ArgumentNullException.ThrowIfNull(values);
        ArgumentNullException.ThrowIfNull(probabilities);

        if (values.Length == 0)
            throw new ArgumentException("Values array must not be empty.", nameof(values));
        if (values.Length != probabilities.Length)
            throw new ArgumentException(
                $"Values array length ({values.Length}) must equal probabilities array length ({probabilities.Length}).");

        for (int i = 0; i < probabilities.Length; i++)
        {
            if (probabilities[i] < 0)
                throw new ArgumentOutOfRangeException(nameof(probabilities),
                    $"Probability at index {i} is negative ({probabilities[i]}).");
        }

        double sum = probabilities.Sum();
        if (Math.Abs(sum - 1.0) > 1e-9)
            throw new ArgumentException(
                $"Probabilities must sum to 1.0, but sum is {sum}.", nameof(probabilities));

        // Store sorted copies for consistent CDF/Percentile behavior
        var indices = Enumerable.Range(0, values.Length).OrderBy(i => values[i]).ToArray();
        _values = indices.Select(i => values[i]).ToArray();
        _probabilities = indices.Select(i => probabilities[i]).ToArray();

        // Categorical needs the original (unsorted) probability order for correct index mapping,
        // but we sorted, so we use sorted probabilities and map back via sorted values.
        var rng = seed.HasValue ? new Random(seed.Value) : new Random();
        _categorical = new Categorical(_probabilities, rng);
    }

    /// <inheritdoc />
    public string Name => "Discrete";

    /// <inheritdoc />
    public double Mean
    {
        get
        {
            double sum = 0;
            for (int i = 0; i < _values.Length; i++)
                sum += _values[i] * _probabilities[i];
            return sum;
        }
    }

    /// <inheritdoc />
    public double Variance
    {
        get
        {
            double mean = Mean;
            double sum = 0;
            for (int i = 0; i < _values.Length; i++)
            {
                double diff = _values[i] - mean;
                sum += diff * diff * _probabilities[i];
            }
            return sum;
        }
    }

    /// <inheritdoc />
    public double StdDev => Math.Sqrt(Variance);

    /// <inheritdoc />
    public double Minimum => _values[0];

    /// <inheritdoc />
    public double Maximum => _values[^1];

    /// <inheritdoc />
    public double Sample() => _values[_categorical.Sample()];

    /// <inheritdoc />
    public double[] Sample(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(count);
        var samples = new double[count];
        for (int i = 0; i < count; i++)
            samples[i] = _values[_categorical.Sample()];
        return samples;
    }

    /// <inheritdoc />
    public double PDF(double x)
    {
        for (int i = 0; i < _values.Length; i++)
        {
            if (Math.Abs(_values[i] - x) < 1e-12)
                return _probabilities[i];
        }
        return 0.0;
    }

    /// <inheritdoc />
    public double CDF(double x)
    {
        double sum = 0;
        for (int i = 0; i < _values.Length; i++)
        {
            if (_values[i] <= x)
                sum += _probabilities[i];
            else
                break; // Values are sorted, so we can stop early
        }
        return sum;
    }

    /// <inheritdoc />
    public double Percentile(double p)
    {
        if (p < 0 || p > 1)
            throw new ArgumentOutOfRangeException(nameof(p), p, "Percentile must be between 0 and 1.");

        double cumulative = 0;
        for (int i = 0; i < _values.Length; i++)
        {
            cumulative += _probabilities[i];
            if (cumulative >= p - 1e-12)
                return _values[i];
        }
        return _values[^1];
    }

    /// <inheritdoc />
    public string ParameterSummary()
    {
        var pairs = _values.Zip(_probabilities, (v, p) => $"{v}:{p:F2}");
        return $"Discrete({string.Join(", ", pairs)})";
    }
}
