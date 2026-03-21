using System.Diagnostics;
using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Benchmarks the simulation engine with synthetic inputs and a trivial evaluator.
/// Measures raw iterations/second to help users understand the overhead
/// versus their model complexity.
/// </summary>
public static class SimulationBenchmark
{
    /// <summary>
    /// Results from a benchmark run.
    /// </summary>
    public record BenchmarkResult(
        int InputCount,
        int IterationCount,
        TimeSpan TotalTime,
        double IterationsPerSecond,
        double MicrosecondsPerIteration);

    /// <summary>
    /// Runs a synthetic benchmark with the specified number of inputs and iterations.
    /// All inputs use Normal distributions and the evaluator returns the sum of inputs.
    /// </summary>
    /// <param name="inputCount">Number of input distributions (default: 20).</param>
    /// <param name="iterationCount">Number of iterations (default: 10,000).</param>
    /// <param name="cancellationToken">Optional cancellation token.</param>
    /// <returns>Benchmark results including timing and throughput.</returns>
    public static async Task<BenchmarkResult> RunAsync(
        int inputCount = 20,
        int iterationCount = 10_000,
        CancellationToken cancellationToken = default)
    {
        var config = new SimulationConfig
        {
            IterationCount = iterationCount,
            RandomSeed = 42,
            UseParallelEvaluation = true
        };

        // Create synthetic Normal inputs
        for (int i = 0; i < inputCount; i++)
        {
            config.Inputs.Add(new SimulationInput
            {
                Id = $"Input_{i}",
                Label = $"Input {i}",
                Distribution = DistributionFactory.Create("Normal", new Dictionary<string, double>
                {
                    ["mean"] = 100 + i,
                    ["stdDev"] = 10
                }, seed: 42 + i),
                BaseValue = 100 + i
            });
        }

        // Single output: sum of all inputs
        config.Outputs.Add(new SimulationOutput
        {
            Id = "Output_Sum",
            Label = "Sum"
        });

        // Trivial evaluator: output = sum of inputs
        var evaluator = (Dictionary<string, double> inputs) =>
        {
            double sum = 0;
            foreach (var value in inputs.Values)
                sum += value;

            return new Dictionary<string, double> { ["Output_Sum"] = sum };
        };

        var engine = new SimulationEngine();
        var sw = Stopwatch.StartNew();

        await engine.RunAsync(config, evaluator, cancellationToken);

        sw.Stop();

        return new BenchmarkResult(
            InputCount: inputCount,
            IterationCount: iterationCount,
            TotalTime: sw.Elapsed,
            IterationsPerSecond: iterationCount / sw.Elapsed.TotalSeconds,
            MicrosecondsPerIteration: sw.Elapsed.TotalMicroseconds / iterationCount);
    }
}
