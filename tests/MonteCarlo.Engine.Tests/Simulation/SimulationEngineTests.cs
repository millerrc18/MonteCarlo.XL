using Xunit;
using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Engine.Tests.Simulation;

public class SimulationEngineTests
{
    [Fact]
    public async Task BasicExecution_ThreeInputsOneOutput_ProducesCorrectDimensions()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "A", Distribution = new NormalDistribution(100, 10, seed: 1) },
                new() { Id = "B", Label = "B", Distribution = new TriangularDistribution(0, 50, 100, seed: 2) },
                new() { Id = "C", Label = "C", Distribution = new UniformDistribution(10, 20, seed: 3) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Sum", Label = "Sum" }
            },
            IterationCount = 1000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Sum"] = inputs["A"] + inputs["B"] + inputs["C"] });

        result.IterationCount.Should().Be(1000);
        result.InputMatrix.GetLength(0).Should().Be(1000);
        result.InputMatrix.GetLength(1).Should().Be(3);
        result.OutputMatrix.GetLength(0).Should().Be(1000);
        result.OutputMatrix.GetLength(1).Should().Be(1);
    }

    [Fact]
    public async Task BasicExecution_OutputMeanApproximatesSumOfInputMeans()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "A", Distribution = new NormalDistribution(100, 10, seed: 42) },
                new() { Id = "B", Label = "B", Distribution = new NormalDistribution(200, 20, seed: 43) },
                new() { Id = "C", Label = "C", Distribution = new NormalDistribution(300, 30, seed: 44) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Sum", Label = "Sum" }
            },
            IterationCount = 10000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Sum"] = inputs["A"] + inputs["B"] + inputs["C"] });

        var outputValues = result.GetOutputValues("Sum");
        double mean = outputValues.Average();
        mean.Should().BeApproximately(600, 5); // 100 + 200 + 300 within statistical tolerance
    }

    [Fact]
    public async Task SeededReproducibility_SameSeedProducesIdenticalResults()
    {
        var config1 = CreateSeededConfig(42);
        var config2 = CreateSeededConfig(42);

        var engine = new SimulationEngine();
        Func<Dictionary<string, double>, Dictionary<string, double>> eval =
            inputs => new Dictionary<string, double> { ["Out"] = inputs["X"] * 2 };

        var result1 = await engine.RunAsync(config1, eval);
        var result2 = await engine.RunAsync(config2, eval);

        for (int i = 0; i < 100; i++)
        {
            result1.OutputMatrix[i, 0].Should().Be(result2.OutputMatrix[i, 0],
                $"iteration {i} should be identical with same seed");
        }
    }

    [Fact]
    public async Task DifferentSeeds_ProduceDifferentResults()
    {
        var config1 = CreateSeededConfig(42);
        var config2 = CreateSeededConfig(99);

        var engine = new SimulationEngine();
        Func<Dictionary<string, double>, Dictionary<string, double>> eval =
            inputs => new Dictionary<string, double> { ["Out"] = inputs["X"] };

        var result1 = await engine.RunAsync(config1, eval);
        var result2 = await engine.RunAsync(config2, eval);

        // At least some values should differ
        bool anyDifferent = false;
        for (int i = 0; i < 100; i++)
        {
            if (Math.Abs(result1.OutputMatrix[i, 0] - result2.OutputMatrix[i, 0]) > 1e-10)
            {
                anyDifferent = true;
                break;
            }
        }
        anyDifferent.Should().BeTrue("different seeds should produce different samples");
    }

    [Fact]
    public async Task Cancellation_StopsExecution()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(0, 1) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 1_000_000,
            ParallelExecution = false // sequential so cancellation is deterministic
        };

        var cts = new CancellationTokenSource();
        var engine = new SimulationEngine();

        // Cancel after a short delay
        cts.CancelAfter(TimeSpan.FromMilliseconds(200));

        Func<Task> act = async () => await engine.RunAsync(config,
            inputs => new Dictionary<string, double> { ["Out"] = inputs["X"] },
            cts.Token);

        await act.Should().ThrowAsync<OperationCanceledException>();
    }

    [Fact]
    public async Task ProgressReporting_FiresAndIncreasesMonotonically()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(0, 1, seed: 1) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 1000,
            ParallelExecution = false
        };

        var progressEvents = new List<SimulationProgressEventArgs>();
        var engine = new SimulationEngine();
        engine.ProgressChanged += (_, e) => progressEvents.Add(e);

        await engine.RunAsync(config,
            inputs => new Dictionary<string, double> { ["Out"] = inputs["X"] });

        progressEvents.Should().NotBeEmpty("progress should be reported");

        // Verify monotonically increasing
        for (int i = 1; i < progressEvents.Count; i++)
        {
            progressEvents[i].CompletedIterations.Should()
                .BeGreaterThanOrEqualTo(progressEvents[i - 1].CompletedIterations);
        }

        // Final event should report all iterations complete
        progressEvents.Last().CompletedIterations.Should().Be(1000);
        progressEvents.Last().TotalIterations.Should().Be(1000);
    }

    [Fact]
    public async Task EmptyInputs_ThrowsArgumentException()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>(),
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 100
        };

        var engine = new SimulationEngine();
        Func<Task> act = async () => await engine.RunAsync(config,
            _ => new Dictionary<string, double> { ["Out"] = 0 });

        await act.Should().ThrowAsync<ArgumentException>();
    }

    [Fact]
    public async Task EmptyOutputs_ThrowsArgumentException()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(0, 1) }
            },
            Outputs = new List<SimulationOutput>(),
            IterationCount = 100
        };

        var engine = new SimulationEngine();
        Func<Task> act = async () => await engine.RunAsync(config,
            _ => new Dictionary<string, double>());

        await act.Should().ThrowAsync<ArgumentException>();
    }

    [Fact]
    public async Task ZeroIterations_ThrowsArgumentException()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(0, 1) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 0
        };

        var engine = new SimulationEngine();
        Func<Task> act = async () => await engine.RunAsync(config,
            _ => new Dictionary<string, double> { ["Out"] = 0 });

        await act.Should().ThrowAsync<ArgumentException>();
    }

    [Fact]
    public async Task SingleIteration_WorksCorrectly()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(100, 10, seed: 1) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 1,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config,
            inputs => new Dictionary<string, double> { ["Out"] = inputs["X"] * 2 });

        result.IterationCount.Should().Be(1);
        result.OutputMatrix[0, 0].Should().Be(result.InputMatrix[0, 0] * 2);
    }

    [Fact]
    public async Task OutputValuesMatchEvaluator_Deterministic()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "A", Distribution = new NormalDistribution(50, 5, seed: 10) },
                new() { Id = "B", Label = "B", Distribution = new UniformDistribution(0, 10, seed: 20) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 500,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config,
            inputs => new Dictionary<string, double> { ["Out"] = inputs["A"] * 2 + inputs["B"] });

        // Verify every row
        for (int i = 0; i < 500; i++)
        {
            double expected = result.InputMatrix[i, 0] * 2 + result.InputMatrix[i, 1];
            result.OutputMatrix[i, 0].Should().BeApproximately(expected, 1e-10,
                $"iteration {i} output should match evaluator");
        }
    }

    [Fact]
    public async Task LargeRun_CompletesInReasonableTime()
    {
        var config = new SimulationConfig
        {
            Inputs = Enumerable.Range(0, 10).Select(i => new SimulationInput
            {
                Id = $"Input{i}",
                Label = $"Input {i}",
                Distribution = new NormalDistribution(100, 10, seed: i)
            }).ToList(),
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Sum", Label = "Sum" }
            },
            IterationCount = 10000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config,
            inputs => new Dictionary<string, double>
            {
                ["Sum"] = inputs.Values.Sum()
            });

        result.IterationCount.Should().Be(10000);
        result.ElapsedTime.Should().BeLessThan(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public async Task GetInputSamples_ReturnsCorrectColumn()
    {
        var config = CreateSeededConfig(42);
        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config,
            inputs => new Dictionary<string, double> { ["Out"] = inputs["X"] });

        var samples = result.GetInputSamples("X");
        samples.Length.Should().Be(100);
        for (int i = 0; i < 100; i++)
            samples[i].Should().Be(result.InputMatrix[i, 0]);
    }

    [Fact]
    public async Task GetOutputValues_ReturnsCorrectColumn()
    {
        var config = CreateSeededConfig(42);
        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config,
            inputs => new Dictionary<string, double> { ["Out"] = inputs["X"] * 3 });

        var values = result.GetOutputValues("Out");
        values.Length.Should().Be(100);
        for (int i = 0; i < 100; i++)
            values[i].Should().Be(result.OutputMatrix[i, 0]);
    }

    [Fact]
    public void GetInputSamples_UnknownId_Throws()
    {
        var result = CreateDummyResult();
        var act = () => result.GetInputSamples("UNKNOWN");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void GetOutputValues_UnknownId_Throws()
    {
        var result = CreateDummyResult();
        var act = () => result.GetOutputValues("UNKNOWN");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public async Task NullEvaluator_ThrowsArgumentNullException()
    {
        var config = CreateSeededConfig(1);
        var engine = new SimulationEngine();

        Func<Task> act = async () => await engine.RunAsync(config, null!);
        await act.Should().ThrowAsync<ArgumentNullException>();
    }

    [Fact]
    public async Task MultipleOutputs_WorksCorrectly()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(10, 1, seed: 1) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Double", Label = "Double" },
                new() { Id = "Square", Label = "Square" }
            },
            IterationCount = 100,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs => new Dictionary<string, double>
        {
            ["Double"] = inputs["X"] * 2,
            ["Square"] = inputs["X"] * inputs["X"]
        });

        result.OutputMatrix.GetLength(1).Should().Be(2);
        for (int i = 0; i < 100; i++)
        {
            double x = result.InputMatrix[i, 0];
            result.OutputMatrix[i, 0].Should().BeApproximately(x * 2, 1e-10);
            result.OutputMatrix[i, 1].Should().BeApproximately(x * x, 1e-10);
        }
    }

    private static SimulationConfig CreateSeededConfig(int seed)
    {
        return new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(0, 1, seed: seed) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 100,
            RandomSeed = seed,
            ParallelExecution = false
        };
    }

    private static SimulationResult CreateDummyResult()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(0, 1) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10
        };
        return new SimulationResult(config, new double[10, 1], new double[10, 1], TimeSpan.Zero);
    }
}
