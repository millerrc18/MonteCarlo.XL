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
    /// Gets the variance of the distribution.
    /// </summary>
    double Variance { get; }

    /// <summary>
    /// Gets the standard deviation of the distribution.
    /// </summary>
    double StdDev { get; }

    /// <summary>
    /// Gets the minimum value of the distribution's support.
    /// </summary>
    double Minimum { get; }

    /// <summary>
    /// Gets the maximum value of the distribution's support.
    /// </summary>
    double Maximum { get; }

    /// <summary>
    /// Draws a single random sample from the distribution.
    /// </summary>
    double Sample();

    /// <summary>
    /// Draws multiple random samples from the distribution.
    /// </summary>
    double[] Sample(int count);

    /// <summary>
    /// Evaluates the probability density function at the given point.
    /// </summary>
    double PDF(double x);

    /// <summary>
    /// Evaluates the cumulative distribution function at the given point.
    /// </summary>
    double CDF(double x);

    /// <summary>
    /// Returns the value at the given percentile (p in 0.0 to 1.0).
    /// </summary>
    double Percentile(double p);

    /// <summary>
    /// Returns a human-readable summary of the distribution's parameters.
    /// </summary>
    string ParameterSummary();
}
