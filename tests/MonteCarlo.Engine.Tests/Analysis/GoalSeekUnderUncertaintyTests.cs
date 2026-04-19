using FluentAssertions;
using MonteCarlo.Engine.Analysis;
using Xunit;

namespace MonteCarlo.Engine.Tests.Analysis;

public class GoalSeekUnderUncertaintyTests
{
    [Fact]
    public void Solve_MeanMetric_FindsDecisionValue()
    {
        var options = new GoalSeekOptions
        {
            LowerBound = 0,
            UpperBound = 100,
            Metric = GoalSeekMetric.Mean,
            DesiredMetricValue = 42,
            MetricTolerance = 0.001,
            MaxIterations = 30
        };

        var result = GoalSeekUnderUncertainty.Solve(
            options,
            decisionValue => new[] { decisionValue - 1, decisionValue, decisionValue + 1 });

        result.Status.Should().Be(GoalSeekStatus.Converged);
        result.BestDecisionValue.Should().BeApproximately(42, 0.01);
        result.BestMetricValue.Should().BeApproximately(42, 0.001);
        result.History.Should().NotBeEmpty();
    }

    [Fact]
    public void Solve_ProbabilityAboveTarget_FindsDesiredProbability()
    {
        var options = new GoalSeekOptions
        {
            LowerBound = 0,
            UpperBound = 20,
            Metric = GoalSeekMetric.ProbabilityAboveTarget,
            OutputTarget = 10,
            DesiredMetricValue = 0.75,
            MetricTolerance = 0.01,
            MaxIterations = 30
        };

        var result = GoalSeekUnderUncertainty.Solve(options, decisionValue =>
            Enumerable.Range(0, 1_001)
                .Select(i => decisionValue - 5.0 + i * 0.01)
                .ToArray());

        result.Status.Should().Be(GoalSeekStatus.Converged);
        result.BestMetricValue.Should().BeApproximately(0.75, 0.01);
        result.BestDecisionValue.Should().BeApproximately(12.5, 0.2);
    }

    [Fact]
    public void Solve_DecreasingMetric_UpdatesBoundsInReverse()
    {
        var options = new GoalSeekOptions
        {
            LowerBound = 0,
            UpperBound = 100,
            Metric = GoalSeekMetric.Mean,
            DesiredMetricValue = 25,
            MetricTolerance = 0.001,
            MaxIterations = 30,
            HigherDecisionIncreasesMetric = false
        };

        var result = GoalSeekUnderUncertainty.Solve(
            options,
            decisionValue => new[] { 100 - decisionValue });

        result.Status.Should().Be(GoalSeekStatus.Converged);
        result.BestDecisionValue.Should().BeApproximately(75, 0.01);
        result.BestMetricValue.Should().BeApproximately(25, 0.001);
    }

    [Fact]
    public void Solve_TargetNotBracketed_ReturnsBestBound()
    {
        var options = new GoalSeekOptions
        {
            LowerBound = 0,
            UpperBound = 10,
            Metric = GoalSeekMetric.Mean,
            DesiredMetricValue = 25
        };

        var result = GoalSeekUnderUncertainty.Solve(
            options,
            decisionValue => new[] { decisionValue });

        result.Status.Should().Be(GoalSeekStatus.TargetNotBracketed);
        result.BestDecisionValue.Should().Be(10);
        result.BestMetricValue.Should().Be(10);
        result.Iterations.Should().Be(0);
    }

    [Fact]
    public void EvaluateMetric_Percentile_ReturnsRequestedPercentile()
    {
        var options = new GoalSeekOptions
        {
            LowerBound = 0,
            UpperBound = 10,
            Metric = GoalSeekMetric.Percentile,
            Percentile = 0.90,
            DesiredMetricValue = 0
        };

        var metric = GoalSeekUnderUncertainty.EvaluateMetric(new[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, options);

        metric.Should().BeApproximately(4.6, 1e-12);
    }

    [Fact]
    public void Solve_InvalidProbabilityTarget_Throws()
    {
        var options = new GoalSeekOptions
        {
            LowerBound = 0,
            UpperBound = 10,
            Metric = GoalSeekMetric.ProbabilityAboveTarget,
            OutputTarget = 5,
            DesiredMetricValue = 1.5
        };

        var act = () => GoalSeekUnderUncertainty.Solve(options, _ => new[] { 1.0 });

        act.Should().Throw<ArgumentException>()
            .WithMessage("*Desired probability*");
    }
}
