using FluentAssertions;
using MonteCarlo.Engine.Analysis;
using Xunit;

namespace MonteCarlo.Engine.Tests.Analysis;

/// <summary>
/// Tests for the ConvergenceChecker.
/// </summary>
public class ConvergenceCheckerTests
{
    [Fact]
    public void StableSeries_ShouldConverge()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.005);
        double[] values = { 100.5, 100.2, 100.1, 100.08, 100.05 };

        checker.IsConverged(values).Should().BeTrue();
    }

    [Fact]
    public void UnstableSeries_ShouldNotConverge()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.005);
        double[] values = { 100, 102, 98, 105, 95 };

        checker.IsConverged(values).Should().BeFalse();
    }

    [Fact]
    public void EventuallyConvergingSeries_ShouldConverge()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.005);
        // Starts volatile, then stabilizes
        double[] values = { 80, 120, 95, 105, 100.3, 100.1, 100.05 };

        checker.IsConverged(values).Should().BeTrue();
    }

    [Fact]
    public void TightTolerance_RequiresMoreStability()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.001);
        double[] values = { 100.5, 100.2, 100.1, 100.08, 100.05 };

        // Change from 100.1 to 100.05 is ~0.05% — within 0.1% tolerance
        checker.IsConverged(values).Should().BeTrue();
    }

    [Fact]
    public void TightTolerance_RejectsLargerDrift()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.0001);
        double[] values = { 100.5, 100.2, 100.1, 100.08, 100.05 };

        // 100.1 to 100.05 is 0.05% — exceeds 0.01% tolerance
        checker.IsConverged(values, 0.0001).Should().BeFalse();
    }

    [Fact]
    public void SingleCheckpoint_ShouldNotConverge()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.005);
        double[] values = { 100 };

        checker.IsConverged(values).Should().BeFalse();
    }

    [Fact]
    public void TwoCheckpoints_WithDefaultWindowOf3_ShouldNotConverge()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.005);
        double[] values = { 100, 100.01 };

        checker.IsConverged(values).Should().BeFalse();
    }

    [Fact]
    public void CheckStat_WithInsufficientData_ReportsInsufficientData()
    {
        var checker = new ConvergenceChecker();

        var result = checker.CheckStat("Mean", new[] { 100.0 });

        result.Status.Should().Be(ConvergenceStatus.InsufficientData);
        result.StatName.Should().Be("Mean");
    }

    [Fact]
    public void CheckStat_StableValues_ReportsStable()
    {
        var checker = new ConvergenceChecker(windowSize: 3, tolerance: 0.005);

        var result = checker.CheckStat("Mean", new[] { 100.5, 100.2, 100.1, 100.08, 100.05 });

        result.Status.Should().Be(ConvergenceStatus.Stable);
    }

    [Fact]
    public void RecordCheckpoint_AndCheckAll_ReturnsAllStats()
    {
        var checker = new ConvergenceChecker(checkpointInterval: 100, windowSize: 2, tolerance: 0.01);

        // Generate stable data
        var rng = new Random(42);
        for (int i = 0; i < 5; i++)
        {
            // Each checkpoint has 1000 similar values
            double[] values = Enumerable.Range(0, 1000)
                .Select(_ => 100.0 + rng.NextDouble() * 0.1) // very tight around 100
                .ToArray();
            checker.RecordCheckpoint((i + 1) * 100, values);
        }

        var indicators = checker.CheckAll();

        indicators.Should().HaveCount(4);
        indicators.Select(i => i.StatName).Should().Contain("Mean");
        indicators.Select(i => i.StatName).Should().Contain("P50");
        indicators.Select(i => i.StatName).Should().Contain("P90");
        indicators.Select(i => i.StatName).Should().Contain("Std Dev");
    }

    [Fact]
    public void Reset_ClearsCheckpoints()
    {
        var checker = new ConvergenceChecker();
        checker.RecordCheckpoint(500, new[] { 1.0, 2.0, 3.0 });
        checker.CheckpointCount.Should().Be(1);

        checker.Reset();
        checker.CheckpointCount.Should().Be(0);
    }
}
