using Xunit;
using FluentAssertions;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.Engine.Tests.Analysis;

public class SummaryStatisticsTests
{
    [Fact]
    public void KnownDistribution_NormalMeanAndStdDev()
    {
        var dist = new NormalDistribution(100, 10, seed: 42);
        var samples = dist.Sample(100_000);
        var stats = new SummaryStatistics(samples);

        stats.Mean.Should().BeApproximately(100, 0.1);
        stats.StdDev.Should().BeApproximately(10, 0.2);
        stats.Median.Should().BeApproximately(100, 0.2);
        stats.Skewness.Should().BeApproximately(0, 0.05);
        stats.Kurtosis.Should().BeApproximately(0, 0.1);
    }

    [Fact]
    public void Percentiles_NormalDistribution()
    {
        var dist = new NormalDistribution(0, 1, seed: 42);
        var samples = dist.Sample(100_000);
        var stats = new SummaryStatistics(samples);

        stats.P50.Should().BeApproximately(0, 0.02);
        stats.P5.Should().BeApproximately(-1.645, 0.05);
        stats.P95.Should().BeApproximately(1.645, 0.05);
    }

    [Fact]
    public void SkewedDistribution_LognormalHasPositiveSkewness()
    {
        var dist = new LognormalDistribution(0, 0.5, seed: 42);
        var samples = dist.Sample(100_000);
        var stats = new SummaryStatistics(samples);

        stats.Skewness.Should().BeGreaterThan(0);
    }

    [Fact]
    public void ProbabilityAbove_NormalSymmetry()
    {
        var dist = new NormalDistribution(100, 10, seed: 42);
        var samples = dist.Sample(100_000);
        var stats = new SummaryStatistics(samples);

        stats.ProbabilityAbove(100).Should().BeApproximately(0.5, 0.01);
    }

    [Fact]
    public void ProbabilityBelow_NormalTail()
    {
        var dist = new NormalDistribution(100, 10, seed: 42);
        var samples = dist.Sample(100_000);
        var stats = new SummaryStatistics(samples);

        // P(X <= 80) for N(100,10) ~ P(Z <= -2) ~ 0.0228
        stats.ProbabilityBelow(80).Should().BeApproximately(0.0228, 0.005);
    }

    [Fact]
    public void ProbabilityBetween_EqualsBelow_Difference()
    {
        var dist = new NormalDistribution(100, 10, seed: 42);
        var samples = dist.Sample(100_000);
        var stats = new SummaryStatistics(samples);

        double between = stats.ProbabilityBetween(90, 110);
        double expected = stats.ProbabilityBelow(110) - stats.ProbabilityBelow(90);
        between.Should().BeApproximately(expected, 1e-10);
    }

    [Fact]
    public void MeanConfidenceInterval_Contains_TrueMean()
    {
        var dist = new NormalDistribution(0, 1, seed: 42);
        var samples = dist.Sample(100_000);
        var stats = new SummaryStatistics(samples);

        var (lower, upper) = stats.MeanConfidenceInterval(0.95);
        lower.Should().BeLessThan(0);
        upper.Should().BeGreaterThan(0);
    }

    [Fact]
    public void Histogram_FrequenciesSumToCount()
    {
        var dist = new NormalDistribution(0, 1, seed: 42);
        var samples = dist.Sample(10_000);
        var stats = new SummaryStatistics(samples);
        var hist = stats.ToHistogram(50);

        hist.Frequencies.Sum().Should().Be(10_000);
        hist.BinEdges.Length.Should().Be(51); // 50 + 1
        hist.BinCenters.Length.Should().Be(50);
        hist.Frequencies.Should().OnlyContain(f => f >= 0);
    }

    [Fact]
    public void Histogram_BinEdges_SpanMinToMax()
    {
        var dist = new NormalDistribution(0, 1, seed: 42);
        var samples = dist.Sample(10_000);
        var stats = new SummaryStatistics(samples);
        var hist = stats.ToHistogram(50);

        hist.BinEdges[0].Should().Be(stats.Min);
        hist.BinEdges[^1].Should().Be(stats.Max);
    }

    [Fact]
    public void DeterministicValues_ExactResults()
    {
        var values = new double[] { 5, 10, 15, 20, 25 };
        var stats = new SummaryStatistics(values);

        stats.Mean.Should().Be(15);
        stats.Median.Should().Be(15);
        stats.Min.Should().Be(5);
        stats.Max.Should().Be(25);
        stats.Range.Should().Be(20);
        stats.Count.Should().Be(5);
    }

    [Fact]
    public void SingleValue_HandledCorrectly()
    {
        var stats = new SummaryStatistics(new[] { 42.0 });

        stats.Mean.Should().Be(42);
        stats.Median.Should().Be(42);
        stats.Min.Should().Be(42);
        stats.Max.Should().Be(42);
        stats.StdDev.Should().Be(0);
        stats.Variance.Should().Be(0);
        stats.Count.Should().Be(1);
    }

    [Fact]
    public void TwoValues_PercentilesInterpolate()
    {
        var stats = new SummaryStatistics(new[] { 0.0, 100.0 });

        stats.Percentile(0.0).Should().Be(0);
        stats.Percentile(0.5).Should().Be(50);
        stats.Percentile(1.0).Should().Be(100);
        stats.Mean.Should().Be(50);
    }

    [Fact]
    public void AllIdenticalValues_StdDevZero_SingleBinHistogram()
    {
        var values = Enumerable.Repeat(7.0, 1000).ToArray();
        var stats = new SummaryStatistics(values);

        stats.StdDev.Should().Be(0);
        stats.Mean.Should().Be(7);
        stats.Median.Should().Be(7);

        var hist = stats.ToHistogram(50);
        hist.Frequencies.Sum().Should().Be(1000);
        hist.BinCenters.Length.Should().Be(1); // collapsed to single bin
    }

    [Fact]
    public void EmptyArray_Throws()
    {
        var act = () => new SummaryStatistics(Array.Empty<double>());
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void NullArray_Throws()
    {
        var act = () => new SummaryStatistics(null!);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Percentile_OutOfRange_Throws()
    {
        var stats = new SummaryStatistics(new[] { 1.0, 2.0, 3.0 });

        var act1 = () => stats.Percentile(-0.1);
        act1.Should().Throw<ArgumentOutOfRangeException>();

        var act2 = () => stats.Percentile(1.1);
        act2.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void SortedValues_AreSorted()
    {
        var rng = new Random(42);
        var values = Enumerable.Range(0, 1000).Select(_ => rng.NextDouble() * 100).ToArray();
        var stats = new SummaryStatistics(values);

        for (int i = 1; i < stats.SortedValues.Length; i++)
            stats.SortedValues[i].Should().BeGreaterThanOrEqualTo(stats.SortedValues[i - 1]);
    }

    [Fact]
    public void ProbabilityBetween_InvalidRange_Throws()
    {
        var stats = new SummaryStatistics(new[] { 1.0, 2.0, 3.0 });
        var act = () => stats.ProbabilityBetween(10, 5);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void MeanConfidenceInterval_SingleValue_ReturnsPointEstimate()
    {
        var stats = new SummaryStatistics(new[] { 42.0 });
        var (lower, upper) = stats.MeanConfidenceInterval(0.95);
        lower.Should().Be(42);
        upper.Should().Be(42);
    }
}
