using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class PoissonDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new PoissonDistribution(4.5);
        dist.Name.Should().Be("Poisson");
        dist.Mean.Should().BeApproximately(4.5, 1e-10);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(-0.001)]
    public void Constructor_InvalidLambda_Throws(double lambda)
    {
        var act = () => new PoissonDistribution(lambda);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("lambda");
    }

    // --- Statistical Properties ---

    [Fact]
    public void Properties_AreCorrect()
    {
        var dist = new PoissonDistribution(4.5);
        dist.Mean.Should().BeApproximately(4.5, 1e-10);
        dist.Variance.Should().BeApproximately(4.5, 1e-10); // Variance = λ
        dist.StdDev.Should().BeApproximately(Math.Sqrt(4.5), 1e-10);
        dist.Minimum.Should().Be(0.0);
        dist.Maximum.Should().Be(double.PositiveInfinity);
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new PoissonDistribution(4.5, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(4.5, 4.5 * 0.02);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalVariance()
    {
        var dist = new PoissonDistribution(4.5, Seed);
        var samples = dist.Sample(SampleCount);
        double mean = samples.Average();
        double empiricalVariance = samples.Sum(s => (s - mean) * (s - mean)) / (samples.Length - 1);
        // For Poisson, Variance ≈ λ
        empiricalVariance.Should().BeApproximately(4.5, 4.5 * 0.03);
    }

    // --- Quantile Round-Trip ---

    [Theory]
    [InlineData(0.1)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    public void CDF_Percentile_RoundTrip(double p)
    {
        var dist = new PoissonDistribution(4.5, Seed);
        double value = dist.Percentile(p);
        // For discrete: CDF(Percentile(p)) >= p
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeGreaterThanOrEqualTo(p);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtNegative_IsZero()
    {
        var dist = new PoissonDistribution(4.5, Seed);
        dist.CDF(-1).Should().Be(0.0);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new PoissonDistribution(4.5, 123);
        var dist2 = new PoissonDistribution(4.5, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- Poisson-Specific Tests ---

    [Fact]
    public void AllSamples_AreNonNegativeIntegers()
    {
        var dist = new PoissonDistribution(4.5, Seed);
        var samples = dist.Sample(SampleCount);
        samples.All(s => s >= 0 && s == Math.Floor(s)).Should().BeTrue();
    }

    [Fact]
    public void LargeLambda_ApproximatesNormal()
    {
        // For large λ, Poisson approximates Normal(λ, √λ)
        double lambda = 100;
        var dist = new PoissonDistribution(lambda, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        double mean = samples.Average();
        double empiricalStdDev = Math.Sqrt(samples.Sum(s => (s - mean) * (s - mean)) / (samples.Length - 1));
        empiricalMean.Should().BeApproximately(lambda, lambda * 0.01);
        empiricalStdDev.Should().BeApproximately(Math.Sqrt(lambda), Math.Sqrt(lambda) * 0.05);
    }

    // --- Batch Sample ---

    [Fact]
    public void Sample_BatchReturnsCorrectCount()
    {
        var dist = new PoissonDistribution(4.5, Seed);
        dist.Sample(500).Should().HaveCount(500);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new PoissonDistribution(4.5);
        dist.ParameterSummary().Should().Be("Poisson(λ=4.5)");
    }
}
