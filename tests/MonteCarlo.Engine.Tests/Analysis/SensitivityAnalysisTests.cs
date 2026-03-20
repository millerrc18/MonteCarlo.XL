using FluentAssertions;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Engine.Tests.Analysis;

public class SensitivityAnalysisTests
{
    [Fact]
    public async Task KnownSensitivity_HighWeightInputRanksFirst()
    {
        // Output = 3*A + 1*B + 0*C → A should be most sensitive, C least
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "A", Distribution = new NormalDistribution(0, 10, seed: 1) },
                new() { Id = "B", Label = "B", Distribution = new NormalDistribution(0, 10, seed: 2) },
                new() { Id = "C", Label = "C", Distribution = new NormalDistribution(0, 10, seed: 3) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Out"] = 3 * inputs["A"] + 1 * inputs["B"] + 0 * inputs["C"] });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);

        sensitivity.Should().HaveCount(3);
        // Sorted by absolute impact, so A first, B second, C last
        sensitivity[0].InputId.Should().Be("A");
        sensitivity[1].InputId.Should().Be("B");
        sensitivity[2].InputId.Should().Be("C");

        // A should have high positive correlation
        sensitivity[0].RankCorrelation.Should().BeGreaterThan(0.8);
        // B should have moderate positive correlation
        sensitivity[1].RankCorrelation.Should().BeGreaterThan(0.2);
        // C should be near zero
        Math.Abs(sensitivity[2].RankCorrelation).Should().BeLessThan(0.05);
    }

    [Fact]
    public async Task VarianceContributions_SumToApproximately100()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "A", Distribution = new NormalDistribution(0, 10, seed: 10) },
                new() { Id = "B", Label = "B", Distribution = new NormalDistribution(0, 5, seed: 20) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Out"] = inputs["A"] + inputs["B"] });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);
        double totalContribution = sensitivity.Sum(s => s.ContributionToVariance);

        totalContribution.Should().BeApproximately(100.0, 1.0);
    }

    [Fact]
    public async Task OutputAtExtremes_PositiveCorrelation_P90HigherThanP10()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(100, 20, seed: 42) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Out"] = inputs["X"] * 2 });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);

        sensitivity[0].OutputAtInputP90.Should().BeGreaterThan(sensitivity[0].OutputAtInputP10,
            "for positive correlation, high input should produce high output");
    }

    [Fact]
    public async Task NegativeCorrelation_DetectedCorrectly()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "Cost", Label = "Cost", Distribution = new NormalDistribution(100, 20, seed: 42) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Profit", Label = "Profit" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Profit"] = 500 - inputs["Cost"] });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);
        sensitivity[0].RankCorrelation.Should().BeLessThan(-0.9,
            "output = 500 - cost means strong negative correlation");
    }

    [Fact]
    public void NullResult_Throws()
    {
        var act = () => SensitivityAnalysis.Analyze(null!, 0);
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public async Task InvalidOutputIndex_Throws()
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
            IterationCount = 100,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Out"] = inputs["X"] });

        var act = () => SensitivityAnalysis.Analyze(result, 5);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public async Task ResultsSortedByAbsoluteImpact()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "Small", Label = "Small", Distribution = new NormalDistribution(0, 1, seed: 1) },
                new() { Id = "Large", Label = "Large", Distribution = new NormalDistribution(0, 10, seed: 2) },
                new() { Id = "Neg", Label = "Neg", Distribution = new NormalDistribution(0, 5, seed: 3) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double>
            {
                ["Out"] = 10 * inputs["Large"] - 5 * inputs["Neg"] + 0.1 * inputs["Small"]
            });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);

        // Should be sorted by |correlation| descending
        for (int i = 1; i < sensitivity.Count; i++)
        {
            Math.Abs(sensitivity[i].RankCorrelation).Should()
                .BeLessThanOrEqualTo(Math.Abs(sensitivity[i - 1].RankCorrelation));
        }
    }

    [Fact]
    public async Task SingleInput_CorrelationNearOne()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(50, 10, seed: 7) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Out"] = inputs["X"] * 3 + 5 });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);
        sensitivity.Should().HaveCount(1);
        sensitivity[0].RankCorrelation.Should().BeApproximately(1.0, 0.01,
            "monotonic function of single input should have correlation ~1");
        sensitivity[0].ContributionToVariance.Should().BeApproximately(100.0, 0.1);
    }

    [Fact]
    public async Task EqualContributions_BothInputsSimilarCorrelation()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "A", Distribution = new NormalDistribution(0, 10, seed: 100) },
                new() { Id = "B", Label = "B", Distribution = new NormalDistribution(0, 10, seed: 200) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Out"] = inputs["A"] + inputs["B"] });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);

        // Both should have roughly equal correlation ~0.707 (1/sqrt(2))
        sensitivity[0].RankCorrelation.Should().BeApproximately(0.707, 0.05);
        sensitivity[1].RankCorrelation.Should().BeApproximately(0.707, 0.05);

        // Both should contribute ~50% to variance
        sensitivity[0].ContributionToVariance.Should().BeApproximately(50.0, 5.0);
        sensitivity[1].ContributionToVariance.Should().BeApproximately(50.0, 5.0);
    }

    [Fact]
    public async Task Swing_Property_EqualsAbsDifference()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "X", Label = "X", Distribution = new NormalDistribution(100, 20, seed: 42) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, inputs =>
            new Dictionary<string, double> { ["Out"] = inputs["X"] * 2 });

        var sensitivity = SensitivityAnalysis.Analyze(result, 0);
        var s = sensitivity[0];
        s.Swing.Should().Be(Math.Abs(s.OutputAtInputP90 - s.OutputAtInputP10));
        s.Swing.Should().BeGreaterThan(0);
    }

    [Fact]
    public async Task ComputeTornadoSwing_WithEvaluator_ProducesCorrectSwings()
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "A", Distribution = new NormalDistribution(100, 20, seed: 1) },
                new() { Id = "B", Label = "B", Distribution = new NormalDistribution(50, 5, seed: 2) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Out" }
            },
            IterationCount = 10_000,
            ParallelExecution = false
        };

        Func<Dictionary<string, double>, Dictionary<string, double>> evaluator =
            inputs => new Dictionary<string, double> { ["Out"] = 3 * inputs["A"] + inputs["B"] };

        var engine = new SimulationEngine();
        var result = await engine.RunAsync(config, evaluator);

        var swingResults = SensitivityAnalysis.ComputeTornadoSwing(result, 0, evaluator);

        // A has 3x weight and 4x stddev → should have much larger swing
        swingResults.Should().HaveCount(2);
        swingResults[0].Swing.Should().BeGreaterThan(swingResults[1].Swing,
            "A (weight=3, stddev=20) should have larger swing than B (weight=1, stddev=5)");
        swingResults[0].InputId.Should().Be("A");
    }
}
