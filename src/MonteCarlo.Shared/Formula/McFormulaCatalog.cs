namespace MonteCarlo.Shared.Formula;

public static class McFormulaCatalog
{
    private static readonly IReadOnlyList<McFormulaDefinition> _definitions =
    [
        new(
            "Normal",
            "Normal",
            "Normal distribution. Returns the mean in normal mode; samples during simulation.",
            [
                new("mean", "Mean (mu)", "Mean (mu)"),
                new("stdDev", "Standard deviation (sigma)", "Standard deviation (sigma), must be > 0")
            ]),
        new(
            "Triangular",
            "Triangular",
            "Triangular distribution. Returns the mode in normal mode.",
            [
                new("min", "Minimum", "Minimum value"),
                new("mode", "Mode", "Most likely value"),
                new("max", "Maximum", "Maximum value")
            ]),
        new(
            "PERT",
            "PERT",
            "PERT distribution. Returns the mode in normal mode.",
            [
                new("min", "Minimum", "Minimum value"),
                new("mode", "Mode", "Most likely value"),
                new("max", "Maximum", "Maximum value")
            ]),
        new(
            "Lognormal",
            "Lognormal",
            "Lognormal distribution. Returns exp(mu + sigma^2 / 2) in normal mode.",
            [
                new("mu", "Mu", "Mean of the underlying normal distribution (mu)"),
                new("sigma", "Sigma", "Std dev of the underlying normal distribution (sigma), must be > 0")
            ]),
        new(
            "Uniform",
            "Uniform",
            "Uniform distribution. Returns the midpoint in normal mode.",
            [
                new("min", "Minimum", "Minimum value"),
                new("max", "Maximum", "Maximum value")
            ]),
        new(
            "Beta",
            "Beta",
            "Beta distribution [0,1]. Returns alpha/(alpha+beta) in normal mode.",
            [
                new("alpha", "Alpha", "Shape parameter alpha, must be > 0"),
                new("beta", "Beta", "Shape parameter beta, must be > 0")
            ]),
        new(
            "Weibull",
            "Weibull",
            "Weibull distribution. Returns scale * Gamma(1 + 1/shape) in normal mode.",
            [
                new("shape", "Shape (k)", "Shape parameter (k), must be > 0"),
                new("scale", "Scale (lambda)", "Scale parameter (lambda), must be > 0")
            ]),
        new(
            "Exponential",
            "Exponential",
            "Exponential distribution. Returns 1/rate in normal mode.",
            [
                new("rate", "Rate (lambda)", "Rate parameter (lambda), must be > 0")
            ]),
        new(
            "Poisson",
            "Poisson",
            "Poisson distribution. Returns lambda in normal mode.",
            [
                new("lambda", "Lambda", "Expected event count (lambda), must be > 0")
            ]),
        new(
            "Gamma",
            "Gamma",
            "Gamma distribution. Returns shape/rate in normal mode.",
            [
                new("shape", "Shape (alpha)", "Shape parameter (alpha), must be > 0"),
                new("rate", "Rate (beta)", "Rate parameter (beta), must be > 0")
            ]),
        new(
            "Logistic",
            "Logistic",
            "Logistic distribution. Returns mu in normal mode.",
            [
                new("mu", "Location (mu)", "Location parameter (mu)"),
                new("s", "Scale (s)", "Scale parameter (s), must be > 0")
            ]),
        new(
            "GEV",
            "GEV",
            "Generalized Extreme Value distribution. Returns expected value in normal mode.",
            [
                new("mu", "Location (mu)", "Location parameter (mu)"),
                new("sigma", "Scale (sigma)", "Scale parameter (sigma), must be > 0"),
                new("xi", "Shape (xi)", "Shape parameter (xi)")
            ]),
        new(
            "Binomial",
            "Binomial",
            "Binomial distribution. Returns n*p in normal mode.",
            [
                new("n", "Trials (n)", "Number of trials, must be > 0"),
                new("p", "Probability (p)", "Probability of success, must be in [0,1]")
            ]),
        new(
            "Geometric",
            "Geometric",
            "Geometric distribution. Returns 1/p in normal mode.",
            [
                new("p", "Probability (p)", "Probability of success, must be in (0,1]")
            ])
    ];

    private static readonly IReadOnlyDictionary<string, McFormulaDefinition> _byFunctionName =
        _definitions.ToDictionary(
            definition => definition.FunctionName,
            definition => definition,
            StringComparer.OrdinalIgnoreCase);

    public static IReadOnlyList<McFormulaDefinition> Definitions => _definitions;

    public static bool TryGetByFunctionName(string functionName, out McFormulaDefinition definition) =>
        _byFunctionName.TryGetValue(functionName, out definition!);
}
