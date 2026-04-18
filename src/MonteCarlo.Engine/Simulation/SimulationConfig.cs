using MonteCarlo.Engine.Correlation;

namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Specifies the sampling method used to generate input samples.
/// </summary>
public enum SamplingMethod
{
    /// <summary>Simple random (Monte Carlo) sampling.</summary>
    MonteCarlo,

    /// <summary>Latin Hypercube Sampling — stratified for better coverage.</summary>
    LatinHypercube
}

/// <summary>
/// Configuration for a Monte Carlo simulation run.
/// </summary>
public class SimulationConfig
{
    /// <summary>
    /// The uncertain inputs with their probability distributions.
    /// </summary>
    public List<SimulationInput> Inputs { get; set; } = new();

    /// <summary>
    /// The output cells to track.
    /// </summary>
    public List<SimulationOutput> Outputs { get; set; } = new();

    /// <summary>
    /// Number of Monte Carlo iterations to run.
    /// </summary>
    public int IterationCount { get; set; } = 5000;

    /// <summary>
    /// Random seed for reproducibility. Null means non-deterministic.
    /// </summary>
    public int? RandomSeed { get; set; }

    /// <summary>
    /// Whether to run evaluator calls in parallel. Default true.
    /// Should be false for Excel recalc mode (COM is single-threaded).
    /// </summary>
    public bool ParallelExecution { get; set; } = true;

    /// <summary>
    /// The sampling method to use. Default is Latin Hypercube for better coverage.
    /// </summary>
    public SamplingMethod Sampling { get; set; } = SamplingMethod.LatinHypercube;

    /// <summary>
    /// Optional Spearman rank correlation matrix between inputs.
    /// Null means inputs are independent (no correlation applied).
    /// </summary>
    public CorrelationMatrix? Correlation { get; set; }

    /// <summary>
    /// If true, the simulation stops early when all tracked statistics converge.
    /// Only applies to sequential execution mode.
    /// </summary>
    public bool AutoStopOnConvergence { get; set; } = false;

    /// <summary>
    /// Minimum number of iterations before convergence checking begins.
    /// </summary>
    public int ConvergenceMinIterations { get; set; } = 500;

    /// <summary>
    /// Validates the configuration and throws <see cref="ArgumentException"/> if invalid.
    /// </summary>
    public void Validate()
    {
        if (Inputs == null || Inputs.Count == 0)
            throw new ArgumentException("At least one input is required.", nameof(Inputs));

        if (Outputs == null || Outputs.Count == 0)
            throw new ArgumentException("At least one output is required.", nameof(Outputs));

        if (IterationCount <= 0)
            throw new ArgumentException("IterationCount must be positive.", nameof(IterationCount));

        var inputIds = new HashSet<string>();
        foreach (var input in Inputs)
        {
            if (string.IsNullOrWhiteSpace(input.Id))
                throw new ArgumentException("All inputs must have a non-empty Id.");
            if (input.Distribution == null)
                throw new ArgumentException($"Input '{input.Id}' has a null distribution.");
            if (!inputIds.Add(input.Id))
                throw new ArgumentException($"Duplicate input Id: '{input.Id}'.");
        }

        var outputIds = new HashSet<string>();
        foreach (var output in Outputs)
        {
            if (string.IsNullOrWhiteSpace(output.Id))
                throw new ArgumentException("All outputs must have a non-empty Id.");
            if (!outputIds.Add(output.Id))
                throw new ArgumentException($"Duplicate output Id: '{output.Id}'.");
        }
    }
}
