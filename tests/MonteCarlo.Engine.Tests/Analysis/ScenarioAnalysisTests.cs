using FluentAssertions;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;
using Xunit;

namespace MonteCarlo.Engine.Tests.Analysis;

public class ScenarioAnalysisTests
{
    [Fact]
    public void Analyze_WorstPercent_FiltersLowestOutputAndRanksInputDeltas()
    {
        var result = CreateResult(
            inputA: new[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 },
            inputB: new[] { 10, 10, 10, 10, 10, 10, 10, 10, 10, 10 },
            output: new[] { 10, 20, 30, 40, 50, 60, 70, 80, 90, 100 });

        var analysis = ScenarioAnalysis.Analyze(result, 0, ScenarioFilterMode.WorstPercent, 0.20);

        analysis.Description.Should().Be("Worst 20% of runs");
        analysis.MatchedIterations.Should().Be(2);
        analysis.MatchedFraction.Should().BeApproximately(0.20, 1e-12);
        analysis.Threshold.Should().BeApproximately(28, 1e-12);
        analysis.InputSummaries[0].InputId.Should().Be("A");
        analysis.InputSummaries[0].ScenarioMean.Should().Be(1.5);
        analysis.InputSummaries[0].OverallMean.Should().Be(5.5);
        analysis.InputSummaries[0].Delta.Should().Be(-4);
    }

    [Fact]
    public void Analyze_BestPercent_FiltersHighestOutput()
    {
        var result = CreateResult(
            inputA: new[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 },
            inputB: new[] { 10, 9, 8, 7, 6, 5, 4, 3, 2, 1 },
            output: new[] { 10, 20, 30, 40, 50, 60, 70, 80, 90, 100 });

        var analysis = ScenarioAnalysis.Analyze(result, 0, ScenarioFilterMode.BestPercent, 0.20);

        analysis.MatchedIterations.Should().Be(2);
        analysis.Threshold.Should().BeApproximately(82, 1e-12);
        analysis.InputSummaries.Should().Contain(s => s.InputId == "A" && s.ScenarioMean == 9.5);
    }

    [Fact]
    public void Analyze_AtOrBelowTarget_UsesInclusiveTarget()
    {
        var result = CreateResult(
            inputA: new[] { 1, 2, 3, 4 },
            inputB: new[] { 4, 3, 2, 1 },
            output: new[] { 100, 200, 300, 400 });

        var analysis = ScenarioAnalysis.Analyze(result, 0, ScenarioFilterMode.AtOrBelowTarget, 200);

        analysis.MatchedIterations.Should().Be(2);
        analysis.Threshold.Should().Be(200);
        analysis.InputSummaries.Single(s => s.InputId == "A").ScenarioMean.Should().Be(1.5);
    }

    [Fact]
    public void Analyze_AboveTarget_ReturnsEmptySummariesWhenNoRunsMatch()
    {
        var result = CreateResult(
            inputA: new[] { 1, 2, 3, 4 },
            inputB: new[] { 4, 3, 2, 1 },
            output: new[] { 100, 200, 300, 400 });

        var analysis = ScenarioAnalysis.Analyze(result, 0, ScenarioFilterMode.AboveTarget, 500);

        analysis.MatchedIterations.Should().Be(0);
        analysis.InputSummaries.Should().BeEmpty();
    }

    [Fact]
    public void Analyze_TailPercentOutsideRange_Throws()
    {
        var result = CreateResult(
            inputA: new[] { 1, 2, 3, 4 },
            inputB: new[] { 4, 3, 2, 1 },
            output: new[] { 100, 200, 300, 400 });

        var act = () => ScenarioAnalysis.Analyze(result, 0, ScenarioFilterMode.WorstPercent, 1.0);

        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    private static SimulationResult CreateResult(int[] inputA, int[] inputB, int[] output)
    {
        var config = new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "Input A", Distribution = new UniformDistribution(0, 10) },
                new() { Id = "B", Label = "Input B", Distribution = new UniformDistribution(0, 10) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Y", Label = "Output Y" }
            }
        };

        var inputMatrix = new double[inputA.Length, 2];
        var outputMatrix = new double[output.Length, 1];

        for (var i = 0; i < inputA.Length; i++)
        {
            inputMatrix[i, 0] = inputA[i];
            inputMatrix[i, 1] = inputB[i];
            outputMatrix[i, 0] = output[i];
        }

        return new SimulationResult(config, inputMatrix, outputMatrix, TimeSpan.FromMilliseconds(10));
    }
}
