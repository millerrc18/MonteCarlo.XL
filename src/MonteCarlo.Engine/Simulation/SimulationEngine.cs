using System.Diagnostics;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Correlation;
using MonteCarlo.Engine.Sampling;

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
    /// Fires when convergence is checked during the simulation (sequential mode only).
    /// </summary>
    public event EventHandler<ConvergenceEventArgs>? ConvergenceChecked;

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
        CancellationToken cancellationToken = default,
        Func<SimulationInput, double, double>? inputTransform = null)
    {
        config.Validate();

        if (evaluator == null)
            throw new ArgumentNullException(nameof(evaluator));

        return Task.Run(() => Execute(config, evaluator, cancellationToken, inputTransform), cancellationToken);
    }

    private SimulationResult Execute(
        SimulationConfig config,
        Func<Dictionary<string, double>, Dictionary<string, double>> evaluator,
        CancellationToken cancellationToken,
        Func<SimulationInput, double, double>? inputTransform)
    {
        var sw = Stopwatch.StartNew();
        int iterations = config.IterationCount;
        int inputCount = config.Inputs.Count;
        int outputCount = config.Outputs.Count;

        // Step 1: Pre-allocate matrices
        var inputMatrix = new double[iterations, inputCount];
        var outputMatrix = new double[iterations, outputCount];

        // Step 2: Pre-generate all input samples
        if (config.Sampling == SamplingMethod.LatinHypercube)
        {
            // Latin Hypercube Sampling: generate stratified [0,1] samples,
            // then transform through each distribution's inverse CDF (Percentile)
            var lhs = new LatinHypercubeSampler(config.RandomSeed);
            var unitSamples = lhs.Generate(iterations, inputCount);

            for (int j = 0; j < inputCount; j++)
            {
                var dist = config.Inputs[j].Distribution;
                for (int i = 0; i < iterations; i++)
                {
                    inputMatrix[i, j] = dist.Percentile(unitSamples[i, j]);
                }
            }
        }
        else
        {
            // Simple Monte Carlo sampling
            var allSamples = new double[inputCount][];
            for (int j = 0; j < inputCount; j++)
            {
                allSamples[j] = config.Inputs[j].Distribution.Sample(iterations);
            }

            for (int j = 0; j < inputCount; j++)
            {
                for (int i = 0; i < iterations; i++)
                {
                    inputMatrix[i, j] = allSamples[j][i];
                }
            }
        }

        // Step 3: Apply Iman-Conover rank correlation if specified
        if (config.Correlation != null)
        {
            ImanConover.Apply(inputMatrix, config.Correlation);
        }

        if (inputTransform != null)
        {
            for (int j = 0; j < inputCount; j++)
            {
                var input = config.Inputs[j];
                for (int i = 0; i < iterations; i++)
                    inputMatrix[i, j] = inputTransform(input, inputMatrix[i, j]);
            }
        }

        // Step 4: Build output index map for fast lookup
        var outputIds = config.Outputs.Select(o => o.Id).ToArray();
        var inputIds = config.Inputs.Select(inp => inp.Id).ToArray();

        // Step 5: Evaluate outputs for each iteration
        if (config.ParallelExecution)
        {
            int completedCount = 0;
            long lastProgressReportTicks = sw.ElapsedTicks;

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
                    long prevTicks = Interlocked.Read(ref lastProgressReportTicks);
                    // Avoid firing too frequently from multiple threads
                    if ((elapsed.Ticks - prevTicks) >= TimeSpan.TicksPerMillisecond * 50 || current == iterations)
                    {
                        Interlocked.Exchange(ref lastProgressReportTicks, elapsed.Ticks);
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
            var convergenceChecker = config.AutoStopOnConvergence ? new ConvergenceChecker() : null;
            int actualIterations = iterations;
            bool stoppedEarly = false;

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

                    // Compute interim histogram every 500 iterations when we have enough data
                    HistogramData? interimHistogram = null;
                    double[]? interimSortedValues = null;
                    if ((i + 1) % 500 == 0 && (i + 1) >= 200 && outputCount >= 1)
                    {
                        var slice = new double[i + 1];
                        for (int s = 0; s <= i; s++)
                            slice[s] = outputMatrix[s, 0];
                        Array.Sort(slice);
                        interimSortedValues = slice;
                        int binCount = Math.Min(50, (int)Math.Sqrt(i + 1));
                        interimHistogram = new HistogramData(slice, binCount);
                    }

                    ProgressChanged?.Invoke(this, new SimulationProgressEventArgs
                    {
                        CompletedIterations = i + 1,
                        TotalIterations = iterations,
                        Elapsed = elapsed,
                        EstimatedRemaining = remaining,
                        InterimHistogram = interimHistogram,
                        InterimSortedValues = interimSortedValues
                    });
                }

                // Convergence checking every 500 iterations after the minimum threshold
                if (convergenceChecker != null
                    && (i + 1) >= config.ConvergenceMinIterations
                    && (i + 1) % 500 == 0
                    && outputCount >= 1)
                {
                    // Extract output values for the first output up to current iteration
                    var outputSlice = new double[i + 1];
                    for (int s = 0; s <= i; s++)
                        outputSlice[s] = outputMatrix[s, 0];

                    convergenceChecker.RecordCheckpoint(i + 1, outputSlice);
                    var indicators = convergenceChecker.CheckAll();
                    bool allConverged = indicators.All(ind => ind.Status == ConvergenceStatus.Stable);

                    ConvergenceChecked?.Invoke(this, new ConvergenceEventArgs
                    {
                        Indicators = indicators,
                        AllConverged = allConverged,
                        Iteration = i + 1
                    });

                    if (allConverged)
                    {
                        actualIterations = i + 1;
                        stoppedEarly = true;
                        break;
                    }
                }
            }

            // Trim matrices if stopped early
            if (stoppedEarly)
            {
                var trimmedInput = new double[actualIterations, inputCount];
                var trimmedOutput = new double[actualIterations, outputCount];
                for (int i = 0; i < actualIterations; i++)
                {
                    for (int j = 0; j < inputCount; j++)
                        trimmedInput[i, j] = inputMatrix[i, j];
                    for (int k = 0; k < outputCount; k++)
                        trimmedOutput[i, k] = outputMatrix[i, k];
                }
                inputMatrix = trimmedInput;
                outputMatrix = trimmedOutput;
            }
        }

        sw.Stop();
        return new SimulationResult(config, inputMatrix, outputMatrix, sw.Elapsed);
    }
}
