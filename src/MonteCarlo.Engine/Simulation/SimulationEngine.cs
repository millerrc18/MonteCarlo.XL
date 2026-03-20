using System.Diagnostics;
using MonteCarlo.Engine.Correlation;

namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Core Monte Carlo simulation engine. Takes input distributions and an evaluator
/// function, runs N iterations, and produces a results matrix.
/// </summary>
public class SimulationEngine
{
    private const int ProgressReportInterval = 100;

    /// <summary>
    /// Fires periodically during simulation to report progress.
    /// </summary>
    public event EventHandler<SimulationProgressEventArgs>? ProgressChanged;

    /// <summary>
    /// Runs a Monte Carlo simulation asynchronously.
    /// </summary>
    /// <param name="config">Simulation configuration defining inputs, outputs, and iteration count.</param>
    /// <param name="evaluator">
    /// A function that takes sampled input values (keyed by input ID) and returns
    /// output values (keyed by output ID). For fast mode, this is a simple in-memory
    /// calculation. For recalc mode, this drives Excel recalculation.
    /// </param>
    /// <param name="cancellationToken">Token to cancel the simulation.</param>
    /// <returns>The simulation results including raw data and timing.</returns>
    public Task<SimulationResult> RunAsync(
        SimulationConfig config,
        Func<Dictionary<string, double>, Dictionary<string, double>> evaluator,
        CancellationToken cancellationToken = default)
    {
        config.Validate();

        if (evaluator == null)
            throw new ArgumentNullException(nameof(evaluator));

        return Task.Run(() => Execute(config, evaluator, cancellationToken), cancellationToken);
    }

    private SimulationResult Execute(
        SimulationConfig config,
        Func<Dictionary<string, double>, Dictionary<string, double>> evaluator,
        CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();
        int iterations = config.IterationCount;
        int inputCount = config.Inputs.Count;
        int outputCount = config.Outputs.Count;

        // Step 1: Pre-allocate matrices
        var inputMatrix = new double[iterations, inputCount];
        var outputMatrix = new double[iterations, outputCount];

        // Step 2: Pre-generate all input samples (batch sampling is faster
        // and required for future Iman-Conover correlation)
        var allSamples = new double[inputCount][];
        for (int j = 0; j < inputCount; j++)
        {
            allSamples[j] = config.Inputs[j].Distribution.Sample(iterations);
        }

        // Copy samples into the input matrix
        for (int j = 0; j < inputCount; j++)
        {
            for (int i = 0; i < iterations; i++)
            {
                inputMatrix[i, j] = allSamples[j][i];
            }
        }

        // Step 3: Apply Iman-Conover rank correlation if specified
        if (config.Correlation != null)
        {
            ImanConover.Apply(inputMatrix, config.Correlation);
        }

        // Step 4: Build output index map for fast lookup
        var outputIds = config.Outputs.Select(o => o.Id).ToArray();
        var inputIds = config.Inputs.Select(inp => inp.Id).ToArray();

        // Step 5: Evaluate outputs for each iteration
        if (config.ParallelExecution)
        {
            int completedCount = 0;
            var lastProgressReport = sw.Elapsed;

            Parallel.For(0, iterations, new ParallelOptions { CancellationToken = cancellationToken }, i =>
            {
                var inputDict = new Dictionary<string, double>(inputCount);
                for (int j = 0; j < inputCount; j++)
                    inputDict[inputIds[j]] = inputMatrix[i, j];

                var outputs = evaluator(inputDict);

                for (int k = 0; k < outputCount; k++)
                    outputMatrix[i, k] = outputs[outputIds[k]];

                int current = Interlocked.Increment(ref completedCount);

                // Throttled progress reporting
                if (current % ProgressReportInterval == 0 || current == iterations)
                {
                    var elapsed = sw.Elapsed;
                    // Avoid firing too frequently from multiple threads
                    if ((elapsed - lastProgressReport).TotalMilliseconds >= 50 || current == iterations)
                    {
                        lastProgressReport = elapsed;
                        var rate = current / elapsed.TotalSeconds;
                        var remaining = rate > 0
                            ? TimeSpan.FromSeconds((iterations - current) / rate)
                            : TimeSpan.Zero;

                        ProgressChanged?.Invoke(this, new SimulationProgressEventArgs
                        {
                            CompletedIterations = current,
                            TotalIterations = iterations,
                            Elapsed = elapsed,
                            EstimatedRemaining = remaining
                        });
                    }
                }
            });
        }
        else
        {
            // Sequential execution (for recalc mode / single-threaded evaluators)
            for (int i = 0; i < iterations; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var inputDict = new Dictionary<string, double>(inputCount);
                for (int j = 0; j < inputCount; j++)
                    inputDict[inputIds[j]] = inputMatrix[i, j];

                var outputs = evaluator(inputDict);

                for (int k = 0; k < outputCount; k++)
                    outputMatrix[i, k] = outputs[outputIds[k]];

                // Throttled progress reporting
                if ((i + 1) % ProgressReportInterval == 0 || i == iterations - 1)
                {
                    var elapsed = sw.Elapsed;
                    var rate = (i + 1) / elapsed.TotalSeconds;
                    var remaining = rate > 0
                        ? TimeSpan.FromSeconds((iterations - i - 1) / rate)
                        : TimeSpan.Zero;

                    ProgressChanged?.Invoke(this, new SimulationProgressEventArgs
                    {
                        CompletedIterations = i + 1,
                        TotalIterations = iterations,
                        Elapsed = elapsed,
                        EstimatedRemaining = remaining
                    });
                }
            }
        }

        sw.Stop();
        return new SimulationResult(config, inputMatrix, outputMatrix, sw.Elapsed);
    }
}
