using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class ExponentialDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new ExponentialDistribution(0.5);
        dist.Name.Should().Be("Exponential");
        dist.Mean.Should().BeApproximately(2.0, 1e-10); // Mean = 1/λ
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(-0.001)]
    public void Constructor_InvalidRate_Throws(double rate)
    {
        var act = () => new ExponentialDistribution(rate);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("rate");
    }

    // --- Statistical Properties ---

    [Fact]
    public void Properties_AreCorrect()
    {
        var dist = new ExponentialDistribution(0.5);
        dist.Mean.Should().BeApproximately(2.0, 1e-10);
        dist.Variance.Should().BeApproximately(4.0, 1e-10); // Variance = 1/λ²
        dist.StdDev.Should().BeApproximately(2.0, 1e-10);
        dist.Minimum.Should().Be(0.0);
        dist.Maximum.Should().Be(double.PositiveInfinity);
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new ExponentialDistribution(0.5, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(2.0, 2.0 * 0.02);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalStdDev()
    {
        var dist = new ExponentialDistribution(0.5, Seed);
        var samples = dist.Sample(SampleCount);
        double mean = samples.Average();
        double empiricalStdDev = Math.Sqrt(samples.Sum(s => (s - mean) * (s - mean)) / (samples.Length - 1));
        empiricalStdDev.Should().BeApproximately(2.0, 2.0 * 0.03);
    }

    // --- Quantile Round-Trip ---

    [Theory]
    [InlineData(0.01)]
    [InlineData(0.1)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    [InlineData(0.99)]
    public void CDF_Percentile_RoundTrip(double p)
    {
        var dist = new ExponentialDistribution(0.5, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtZero()
    {
        var dist = new ExponentialDistribution(0.5, Seed);
        dist.CDF(0).Should().BeApproximately(0, 1e-10);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new ExponentialDistribution(0.5, 123);
        var dist2 = new ExponentialDistribution(0.5, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- Exponential-Specific Tests ---

    [Fact]
    public void AllSamples_AreNonNegative()
    {
        var dist = new ExponentialDistribution(0.5, Seed);
        var samples = dist.Sample(SampleCount);
        samples.All(s => s >= 0).Should().BeTrue();
    }

    [Fact]
    public void MeanIs_OneOverRate()
    {
        var dist = new ExponentialDistribution(2.0, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(0.5, 0.5 * 0.02); // 1/2 = 0.5
    }

    // --- Batch Sample ---

    [Fact]
    public void Sample_BatchReturnsCorrectCount()
    {
        var dist = new ExponentialDistribution(0.5, Seed);
        dist.Sample(500).Should().HaveCount(500);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new ExponentialDistribution(0.5);
        dist.ParameterSummary().Should().Be("Exponential(λ=0.5)");
    }
}
