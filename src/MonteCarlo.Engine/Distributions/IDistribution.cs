namespace MonteCarlo.Engine.Distributions;

/// <summary>
/// Defines the contract for a probability distribution used in Monte Carlo simulation.
/// </summary>
public interface IDistribution
{
    /// <summary>
    /// Gets the human-readable name of the distribution (e.g., "Normal", "Triangular").
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Gets the mean (expected value) of the distribution.
    /// </summary>
    double Mean { get; }

    /// <summary>
    /// Draws a random sample from the distribution.
    /// </summary>
    double Sample();

    /// <summary>
    /// Evaluates the probability density function at the given point.
    /// </summary>
    double Pdf(double x);

    /// <summary>
    /// Evaluates the cumulative distribution function at the given point.
    /// </summary>
    double Cdf(double x);

    /// <summary>
    /// Returns the value at the given percentile (0.0 to 1.0).
    /// </summary>
    double Percentile(double p);
}
