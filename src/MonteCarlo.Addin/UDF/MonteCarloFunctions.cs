using ExcelDna.Integration;
using MonteCarlo.Shared.Formula;

namespace MonteCarlo.Addin.UDF;

/// <summary>
/// Custom Excel functions (UDFs) for Monte Carlo simulation.
/// In normal mode, each function returns the distribution's expected value.
/// During simulation, the orchestrator overrides these cells with sampled values.
/// </summary>
public static class MonteCarloFunctions
{
    [ExcelFunction(
        Name = "MC.Normal",
        Description = "Normal distribution. Returns the mean in normal mode; samples during simulation.",
        Category = "MonteCarlo.XL")]
    public static object MCNormal(
        [ExcelArgument(Name = "mean", Description = "Mean (μ)")] double mean,
        [ExcelArgument(Name = "stdDev", Description = "Standard deviation (σ), must be > 0")] double stdDev)
        => Evaluate("Normal", mean, stdDev);

    [ExcelFunction(
        Name = "MC.Triangular",
        Description = "Triangular distribution. Returns the mode in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCTriangular(
        [ExcelArgument(Name = "min", Description = "Minimum value")] double min,
        [ExcelArgument(Name = "mode", Description = "Most likely value")] double mode,
        [ExcelArgument(Name = "max", Description = "Maximum value")] double max)
        => Evaluate("Triangular", min, mode, max);

    [ExcelFunction(
        Name = "MC.PERT",
        Description = "PERT distribution. Returns the mode in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCPERT(
        [ExcelArgument(Name = "min", Description = "Minimum value")] double min,
        [ExcelArgument(Name = "mode", Description = "Most likely value")] double mode,
        [ExcelArgument(Name = "max", Description = "Maximum value")] double max)
        => Evaluate("PERT", min, mode, max);

    [ExcelFunction(
        Name = "MC.Lognormal",
        Description = "Lognormal distribution. Returns exp(μ + σ²/2) in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCLognormal(
        [ExcelArgument(Name = "mu", Description = "Mean of the underlying normal distribution (μ)")] double mu,
        [ExcelArgument(Name = "sigma", Description = "Std dev of the underlying normal distribution (σ), must be > 0")] double sigma)
        => Evaluate("Lognormal", mu, sigma);

    [ExcelFunction(
        Name = "MC.Uniform",
        Description = "Uniform distribution. Returns the midpoint in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCUniform(
        [ExcelArgument(Name = "min", Description = "Minimum value")] double min,
        [ExcelArgument(Name = "max", Description = "Maximum value")] double max)
        => Evaluate("Uniform", min, max);

    [ExcelFunction(
        Name = "MC.Beta",
        Description = "Beta distribution [0,1]. Returns α/(α+β) in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCBeta(
        [ExcelArgument(Name = "alpha", Description = "Shape parameter α, must be > 0")] double alpha,
        [ExcelArgument(Name = "beta", Description = "Shape parameter β, must be > 0")] double beta)
        => Evaluate("Beta", alpha, beta);

    [ExcelFunction(
        Name = "MC.Weibull",
        Description = "Weibull distribution. Returns scale × Γ(1 + 1/shape) in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCWeibull(
        [ExcelArgument(Name = "shape", Description = "Shape parameter (k), must be > 0")] double shape,
        [ExcelArgument(Name = "scale", Description = "Scale parameter (λ), must be > 0")] double scale)
        => Evaluate("Weibull", shape, scale);

    [ExcelFunction(
        Name = "MC.Exponential",
        Description = "Exponential distribution. Returns 1/λ in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCExponential(
        [ExcelArgument(Name = "rate", Description = "Rate parameter (λ), must be > 0")] double rate)
        => Evaluate("Exponential", rate);

    [ExcelFunction(
        Name = "MC.Poisson",
        Description = "Poisson distribution. Returns λ in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCPoisson(
        [ExcelArgument(Name = "lambda", Description = "Expected event count (λ), must be > 0")] double lambda)
        => Evaluate("Poisson", lambda);

    [ExcelFunction(
        Name = "MC.Gamma",
        Description = "Gamma distribution. Returns shape/rate in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCGamma(
        [ExcelArgument(Name = "shape", Description = "Shape parameter (α), must be > 0")] double shape,
        [ExcelArgument(Name = "rate", Description = "Rate parameter (β), must be > 0")] double rate)
        => Evaluate("Gamma", shape, rate);

    [ExcelFunction(
        Name = "MC.Logistic",
        Description = "Logistic distribution. Returns μ in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCLogistic(
        [ExcelArgument(Name = "mu", Description = "Location parameter (μ)")] double mu,
        [ExcelArgument(Name = "s", Description = "Scale parameter (s), must be > 0")] double s)
        => Evaluate("Logistic", mu, s);

    [ExcelFunction(
        Name = "MC.GEV",
        Description = "Generalized Extreme Value distribution. Returns expected value in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCGEV(
        [ExcelArgument(Name = "mu", Description = "Location parameter (μ)")] double mu,
        [ExcelArgument(Name = "sigma", Description = "Scale parameter (σ), must be > 0")] double sigma,
        [ExcelArgument(Name = "xi", Description = "Shape parameter (ξ)")] double xi)
        => Evaluate("GEV", mu, sigma, xi);

    [ExcelFunction(
        Name = "MC.Binomial",
        Description = "Binomial distribution. Returns n×p in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCBinomial(
        [ExcelArgument(Name = "n", Description = "Number of trials, must be > 0")] int n,
        [ExcelArgument(Name = "p", Description = "Probability of success, must be in [0,1]")] double p)
        => Evaluate("Binomial", n, p);

    [ExcelFunction(
        Name = "MC.Geometric",
        Description = "Geometric distribution. Returns 1/p in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCGeometric(
        [ExcelArgument(Name = "p", Description = "Probability of success, must be in (0,1]")] double p)
        => Evaluate("Geometric", p);

    private static object Evaluate(string functionName, params double[] arguments)
    {
        return McNormalModeEvaluator.TryEvaluate(functionName, arguments, out var result)
            ? result
            : ExcelError.ExcelErrorValue;
    }
}
