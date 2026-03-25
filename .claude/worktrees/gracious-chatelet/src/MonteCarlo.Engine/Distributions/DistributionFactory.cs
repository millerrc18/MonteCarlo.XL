namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Factory for creating distributions by name and parameter dictionary.
/// Used by the UI layer to create distributions from user input.
/// </summary>
public static class DistributionFactory
{
    private static readonly IReadOnlyList<string> _available = new[]
    {
        "Normal", "Triangular", "PERT", "Lognormal", "Uniform", "Discrete",
        "Beta", "Weibull", "Exponential", "Poisson"
    };

    /// <summary>
    /// Gets the list of available distribution names.
    /// </summary>
    public static IReadOnlyList<string> AvailableDistributions => _available;

    /// <summary>
    /// Creates a distribution by name and parameter dictionary.
    /// </summary>
    /// <param name="name">Distribution name (case-insensitive).</param>
    /// <param name="parameters">Dictionary of parameter names to values.</param>
    /// <param name="seed">Optional RNG seed for reproducibility.</param>
    /// <returns>A new <see cref="IDistribution"/> instance.</returns>
    /// <exception cref="ArgumentException">Thrown when the distribution name is unknown or required parameters are missing.</exception>
    public static IDistribution Create(string name, Dictionary<string, double> parameters, int? seed = null)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentNullException.ThrowIfNull(parameters);

        return name.ToLowerInvariant() switch
        {
            "normal" => CreateNormal(parameters, seed),
            "triangular" => CreateTriangular(parameters, seed),
            "pert" => CreatePERT(parameters, seed),
            "lognormal" => CreateLognormal(parameters, seed),
            "uniform" => CreateUniform(parameters, seed),
            "discrete" => CreateDiscrete(parameters, seed),
            "beta" => CreateBeta(parameters, seed),
            "weibull" => CreateWeibull(parameters, seed),
            "exponential" => CreateExponential(parameters, seed),
            "poisson" => CreatePoisson(parameters, seed),
            _ => throw new ArgumentException($"Unknown distribution: '{name}'. Available: {string.Join(", ", _available)}.", nameof(name))
        };
    }

    private static double GetRequired(Dictionary<string, double> parameters, string key)
    {
        if (!parameters.TryGetValue(key, out double value))
            throw new ArgumentException($"Missing required parameter: '{key}'.");
        return value;
    }

    private static NormalDistribution CreateNormal(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "mean"), GetRequired(p, "stdDev"), seed);

    private static TriangularDistribution CreateTriangular(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "min"), GetRequired(p, "mode"), GetRequired(p, "max"), seed);

    private static PERTDistribution CreatePERT(Dictionary<string, double> p, int? seed)
    {
        double min = GetRequired(p, "min");
        double mode = GetRequired(p, "mode");
        double max = GetRequired(p, "max");
        double lambda = p.TryGetValue("lambda", out double l) ? l : 4.0;
        return new PERTDistribution(min, mode, max, lambda, seed);
    }

    private static LognormalDistribution CreateLognormal(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "mu"), GetRequired(p, "sigma"), seed);

    private static UniformDistribution CreateUniform(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "min"), GetRequired(p, "max"), seed);

    private static DiscreteDistribution CreateDiscrete(Dictionary<string, double> p, int? seed)
    {
        // Discrete expects value_0, value_1, ... and prob_0, prob_1, ... keys
        var values = new List<double>();
        var probs = new List<double>();

        for (int i = 0; ; i++)
        {
            string vKey = $"value_{i}";
            string pKey = $"prob_{i}";
            if (!p.ContainsKey(vKey) || !p.ContainsKey(pKey))
                break;
            values.Add(p[vKey]);
            probs.Add(p[pKey]);
        }

        if (values.Count == 0)
            throw new ArgumentException("Discrete distribution requires at least one value_0/prob_0 pair.");

        return new DiscreteDistribution(values.ToArray(), probs.ToArray(), seed);
    }

    private static BetaDistribution CreateBeta(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "alpha"), GetRequired(p, "beta"), seed);

    private static WeibullDistribution CreateWeibull(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "shape"), GetRequired(p, "scale"), seed);

    private static ExponentialDistribution CreateExponential(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "rate"), seed);

    private static PoissonDistribution CreatePoisson(Dictionary<string, double> p, int? seed) =>
        new(GetRequired(p, "lambda"), seed);
}
