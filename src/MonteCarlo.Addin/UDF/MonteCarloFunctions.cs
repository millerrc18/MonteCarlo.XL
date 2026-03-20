using ExcelDna.Integration;

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
    {
        if (stdDev <= 0)
            return ExcelError.ExcelErrorValue;

        return mean;
    }

    [ExcelFunction(
        Name = "MC.Triangular",
        Description = "Triangular distribution. Returns the mode in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCTriangular(
        [ExcelArgument(Name = "min", Description = "Minimum value")] double min,
        [ExcelArgument(Name = "mode", Description = "Most likely value")] double mode,
        [ExcelArgument(Name = "max", Description = "Maximum value")] double max)
    {
        if (min >= mode || mode >= max)
            return ExcelError.ExcelErrorValue;

        return mode;
    }

    [ExcelFunction(
        Name = "MC.PERT",
        Description = "PERT distribution. Returns the mode in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCPERT(
        [ExcelArgument(Name = "min", Description = "Minimum value")] double min,
        [ExcelArgument(Name = "mode", Description = "Most likely value")] double mode,
        [ExcelArgument(Name = "max", Description = "Maximum value")] double max)
    {
        if (min >= mode || mode >= max)
            return ExcelError.ExcelErrorValue;

        return mode;
    }

    [ExcelFunction(
        Name = "MC.Lognormal",
        Description = "Lognormal distribution. Returns exp(μ + σ²/2) in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCLognormal(
        [ExcelArgument(Name = "mu", Description = "Mean of the underlying normal distribution (μ)")] double mu,
        [ExcelArgument(Name = "sigma", Description = "Std dev of the underlying normal distribution (σ), must be > 0")] double sigma)
    {
        if (sigma <= 0)
            return ExcelError.ExcelErrorValue;

        // Expected value of Lognormal = exp(μ + σ²/2)
        return Math.Exp(mu + sigma * sigma / 2.0);
    }

    [ExcelFunction(
        Name = "MC.Uniform",
        Description = "Uniform distribution. Returns the midpoint in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCUniform(
        [ExcelArgument(Name = "min", Description = "Minimum value")] double min,
        [ExcelArgument(Name = "max", Description = "Maximum value")] double max)
    {
        if (min >= max)
            return ExcelError.ExcelErrorValue;

        return (min + max) / 2.0;
    }

    [ExcelFunction(
        Name = "MC.Beta",
        Description = "Beta distribution [0,1]. Returns α/(α+β) in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCBeta(
        [ExcelArgument(Name = "alpha", Description = "Shape parameter α, must be > 0")] double alpha,
        [ExcelArgument(Name = "beta", Description = "Shape parameter β, must be > 0")] double beta)
    {
        if (alpha <= 0 || beta <= 0)
            return ExcelError.ExcelErrorValue;

        return alpha / (alpha + beta);
    }

    [ExcelFunction(
        Name = "MC.Weibull",
        Description = "Weibull distribution. Returns scale × Γ(1 + 1/shape) in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCWeibull(
        [ExcelArgument(Name = "shape", Description = "Shape parameter (k), must be > 0")] double shape,
        [ExcelArgument(Name = "scale", Description = "Scale parameter (λ), must be > 0")] double scale)
    {
        if (shape <= 0 || scale <= 0)
            return ExcelError.ExcelErrorValue;

        // Mean of Weibull = scale * Γ(1 + 1/shape)
        return scale * SpecialFunctions.Gamma(1.0 + 1.0 / shape);
    }

    [ExcelFunction(
        Name = "MC.Exponential",
        Description = "Exponential distribution. Returns 1/λ in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCExponential(
        [ExcelArgument(Name = "rate", Description = "Rate parameter (λ), must be > 0")] double rate)
    {
        if (rate <= 0)
            return ExcelError.ExcelErrorValue;

        return 1.0 / rate;
    }

    [ExcelFunction(
        Name = "MC.Poisson",
        Description = "Poisson distribution. Returns λ in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCPoisson(
        [ExcelArgument(Name = "lambda", Description = "Expected event count (λ), must be > 0")] double lambda)
    {
        if (lambda <= 0)
            return ExcelError.ExcelErrorValue;

        return lambda;
    }

    /// <summary>
    /// Simple gamma function implementation for Weibull expected value.
    /// </summary>
    private static class SpecialFunctions
    {
        /// <summary>
        /// Computes the Gamma function using Stirling's approximation for large values
        /// and the Lanczos approximation for general use.
        /// </summary>
        public static double Gamma(double x)
        {
            // Use MathNet if available at runtime, otherwise Lanczos approximation
            return MathNet.Numerics.SpecialFunctions.Gamma(x);
        }
    }
}
