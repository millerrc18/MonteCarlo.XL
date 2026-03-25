using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class BetaDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new BetaDistribution(2, 5);
        dist.Name.Should().Be("Beta");
        dist.Mean.Should().BeApproximately(2.0 / 7.0, 1e-10);
    }

    [Theory]
    [InlineData(0, 5)]
    [InlineData(-1, 5)]
    public void Constructor_InvalidAlpha_Throws(double alpha, double beta)
    {
        var act = () => new BetaDistribution(alpha, beta);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("alpha");
    }

    [Theory]
    [InlineData(2, 0)]
    [InlineData(2, -1)]
    public void Constructor_InvalidBeta_Throws(double alpha, double beta)
    {
        var act = () => new BetaDistribution(alpha, beta);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("beta");
    }

    // --- Statistical Properties ---

    [Fact]
    public void Properties_AreCorrect()
    {
        var dist = new BetaDistribution(2, 5);
        dist.Mean.Should().BeApproximately(2.0 / 7.0, 1e-10);
        dist.Minimum.Should().Be(0.0);
        dist.Maximum.Should().Be(1.0);
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new BetaDistribution(2, 5, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(2.0 / 7.0, 0.01);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalStdDev()
    {
        var dist = new BetaDistribution(2, 5, Seed);
        var samples = dist.Sample(SampleCount);
        double mean = samples.Average();
        double empiricalStdDev = Math.Sqrt(samples.Sum(s => (s - mean) * (s - mean)) / (samples.Length - 1));
        empiricalStdDev.Should().BeApproximately(dist.StdDev, dist.StdDev * 0.03);
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
        var dist = new BetaDistribution(2, 5, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtBoundaries()
    {
        var dist = new BetaDistribution(2, 5, Seed);
        dist.CDF(0).Should().BeApproximately(0, 1e-10);
        dist.CDF(1).Should().BeApproximately(1, 1e-10);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new BetaDistribution(2, 5, 123);
        var dist2 = new BetaDistribution(2, 5, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- Beta-Specific Tests ---

    [Fact]
    public void Beta_1_1_IsUniformOnZeroOne()
    {
        var dist = new BetaDistribution(1, 1, Seed);
        var samples = dist.Sample(SampleCount);
        samples.All(s => s >= 0 && s <= 1).Should().BeTrue();
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(0.5, 0.01);
    }

    [Fact]
    public void AllSamples_AreInZeroOneRange()
    {
        var dist = new BetaDistribution(0.5, 0.5, Seed);
        var samples = dist.Sample(SampleCount);
        samples.All(s => s >= 0 && s <= 1).Should().BeTrue();
    }

    [Fact]
    public void Beta_2_5_MeanIsCorrect()
    {
        var dist = new BetaDistribution(2, 5, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        // Mean = α / (α + β) = 2/7 ≈ 0.286
        empiricalMean.Should().BeApproximately(2.0 / 7.0, 0.005);
    }

    // --- Batch Sample ---

    [Fact]
    public void Sample_BatchReturnsCorrectCount()
    {
        var dist = new BetaDistribution(2, 5, Seed);
        dist.Sample(500).Should().HaveCount(500);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new BetaDistribution(2, 5);
        dist.ParameterSummary().Should().Be("Beta(α=2, β=5)");
    }
}
