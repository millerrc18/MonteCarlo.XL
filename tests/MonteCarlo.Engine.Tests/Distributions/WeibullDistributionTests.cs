using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class WeibullDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new WeibullDistribution(2, 100);
        dist.Name.Should().Be("Weibull");
    }

    [Theory]
    [InlineData(0, 100)]
    [InlineData(-1, 100)]
    public void Constructor_InvalidShape_Throws(double shape, double scale)
    {
        var act = () => new WeibullDistribution(shape, scale);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("shape");
    }

    [Theory]
    [InlineData(2, 0)]
    [InlineData(2, -1)]
    public void Constructor_InvalidScale_Throws(double shape, double scale)
    {
        var act = () => new WeibullDistribution(shape, scale);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("scale");
    }

    // --- Statistical Properties ---

    [Fact]
    public void Properties_AreCorrect()
    {
        var dist = new WeibullDistribution(2, 100);
        dist.Minimum.Should().Be(0.0);
        dist.Maximum.Should().Be(double.PositiveInfinity);
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new WeibullDistribution(2, 100, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(dist.Mean, dist.Mean * 0.02);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalStdDev()
    {
        var dist = new WeibullDistribution(2, 100, Seed);
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
        var dist = new WeibullDistribution(2, 100, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtZero()
    {
        var dist = new WeibullDistribution(2, 100, Seed);
        dist.CDF(0).Should().BeApproximately(0, 1e-10);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new WeibullDistribution(2, 100, 123);
        var dist2 = new WeibullDistribution(2, 100, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- Weibull-Specific Tests ---

    [Fact]
    public void AllSamples_AreNonNegative()
    {
        var dist = new WeibullDistribution(2, 100, Seed);
        var samples = dist.Sample(SampleCount);
        samples.All(s => s >= 0).Should().BeTrue();
    }

    [Fact]
    public void Weibull_Shape1_BehavesLikeExponential()
    {
        // Weibull(1, λ) has the same distribution as Exponential(1/λ)
        double scale = 50.0;
        var weibull = new WeibullDistribution(1, scale, Seed);
        double expectedMean = scale; // Weibull(1, λ) mean = λ
        var samples = weibull.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(expectedMean, expectedMean * 0.02);
    }

    // --- Batch Sample ---

    [Fact]
    public void Sample_BatchReturnsCorrectCount()
    {
        var dist = new WeibullDistribution(2, 100, Seed);
        dist.Sample(500).Should().HaveCount(500);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new WeibullDistribution(2.5, 100);
        dist.ParameterSummary().Should().Be("Weibull(k=2.5, λ=100)");
    }
}
