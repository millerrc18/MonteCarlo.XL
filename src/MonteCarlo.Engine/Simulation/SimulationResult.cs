namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Contains the full results of a Monte Carlo simulation run.
/// </summary>
public class SimulationResult
{
    private readonly Dictionary<string, int> _inputIndexMap;
    private readonly Dictionary<string, int> _outputIndexMap;

    /// <summary>
    /// Creates a new simulation result.
    /// </summary>
    public SimulationResult(
        SimulationConfig config,
        double[,] inputMatrix,
        double[,] outputMatrix,
        TimeSpan elapsedTime)
    {
        Config = config ?? throw new ArgumentNullException(nameof(config));
        InputMatrix = inputMatrix ?? throw new ArgumentNullException(nameof(inputMatrix));
        OutputMatrix = outputMatrix ?? throw new ArgumentNullException(nameof(outputMatrix));
        ElapsedTime = elapsedTime;
        IterationCount = inputMatrix.GetLength(0);

        _inputIndexMap = new Dictionary<string, int>();
        for (int i = 0; i < config.Inputs.Count; i++)
            _inputIndexMap[config.Inputs[i].Id] = i;

        _outputIndexMap = new Dictionary<string, int>();
        for (int i = 0; i < config.Outputs.Count; i++)
            _outputIndexMap[config.Outputs[i].Id] = i;
    }

    /// <summary>
    /// The configuration used for this simulation.
    /// </summary>
    public SimulationConfig Config { get; }

    /// <summary>
    /// Number of iterations that were executed.
    /// </summary>
    public int IterationCount { get; }

    /// <summary>
    /// Wall-clock time for the simulation run.
    /// </summary>
    public TimeSpan ElapsedTime { get; }

    /// <summary>
    /// Raw input samples: [iteration, inputIndex].
    /// </summary>
    public double[,] InputMatrix { get; }

    /// <summary>
    /// Raw output values: [iteration, outputIndex].
    /// </summary>
    public double[,] OutputMatrix { get; }

    /// <summary>
    /// Gets all samples for a specific input across all iterations.
    /// </summary>
    public double[] GetInputSamples(string inputId)
    {
        if (!_inputIndexMap.TryGetValue(inputId, out int idx))
            throw new ArgumentException($"Unknown input Id: '{inputId}'.", nameof(inputId));

        var samples = new double[IterationCount];
        for (int i = 0; i < IterationCount; i++)
            samples[i] = InputMatrix[i, idx];
        return samples;
    }

    /// <summary>
    /// Gets all values for a specific output across all iterations.
    /// </summary>
    public double[] GetOutputValues(string outputId)
    {
        if (!_outputIndexMap.TryGetValue(outputId, out int idx))
            throw new ArgumentException($"Unknown output Id: '{outputId}'.", nameof(outputId));

        var values = new double[IterationCount];
        for (int i = 0; i < IterationCount; i++)
            values[i] = OutputMatrix[i, idx];
        return values;
    }

    /// <summary>
    /// Gets a single input sample for a specific iteration.
    /// </summary>
    public double GetInputSample(string inputId, int iteration)
    {
        if (!_inputIndexMap.TryGetValue(inputId, out int idx))
            throw new ArgumentException($"Unknown input Id: '{inputId}'.", nameof(inputId));
        if (iteration < 0 || iteration >= IterationCount)
            throw new ArgumentOutOfRangeException(nameof(iteration));

        return InputMatrix[iteration, idx];
    }

    /// <summary>
    /// Gets a single output value for a specific iteration.
    /// </summary>
    public double GetOutputValue(string outputId, int iteration)
    {
        if (!_outputIndexMap.TryGetValue(outputId, out int idx))
            throw new ArgumentException($"Unknown output Id: '{outputId}'.", nameof(outputId));
        if (iteration < 0 || iteration >= IterationCount)
            throw new ArgumentOutOfRangeException(nameof(iteration));

        return OutputMatrix[iteration, idx];
    }
}
