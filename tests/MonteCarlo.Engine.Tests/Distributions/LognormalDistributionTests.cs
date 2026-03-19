using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class LognormalDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new LognormalDistribution(4.6, 0.3);
        dist.Name.Should().Be("Lognormal");
        dist.Minimum.Should().Be(0);
        dist.Maximum.Should().Be(double.PositiveInfinity);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Constructor_InvalidSigma_Throws(double sigma)
    {
        var act = () => new LognormalDistribution(0, sigma);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("sigma");
    }

    // --- Statistical Properties ---

    [Fact]
    public void Mean_IsCorrect()
    {
        // Mean of Lognormal = exp(mu + sigma^2/2)
        var dist = new LognormalDistribution(4.6, 0.3);
        double expected = Math.Exp(4.6 + 0.3 * 0.3 / 2.0);
        dist.Mean.Should().BeApproximately(expected, 1e-10);
    }

    [Fact]
    public void Variance_IsCorrect()
    {
        var dist = new LognormalDistribution(4.6, 0.3);
        double s2 = 0.3 * 0.3;
        double expected = (Math.Exp(s2) - 1.0) * Math.Exp(2.0 * 4.6 + s2);
        dist.Variance.Should().BeApproximately(expected, 1e-6);
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new LognormalDistribution(4.6, 0.3, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(dist.Mean, dist.Mean * 0.01);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalStdDev()
    {
        var dist = new LognormalDistribution(4.6, 0.3, Seed);
        var samples = dist.Sample(SampleCount);
        double mean = samples.Average();
        double empiricalStdDev = Math.Sqrt(samples.Sum(s => (s - mean) * (s - mean)) / (samples.Length - 1));
        empiricalStdDev.Should().BeApproximately(dist.StdDev, dist.StdDev * 0.02);
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
        var dist = new LognormalDistribution(4.6, 0.3, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    // --- PDF Integrates to 1 ---

    [Fact]
    public void PDF_IntegratesToOne()
    {
        var dist = new LognormalDistribution(4.6, 0.3, Seed);
        double lower = 0.001; // Avoid 0 where PDF may be 0 or undefined
        double upper = dist.Mean + 6 * dist.StdDev;
        double integral = NumericalIntegrate(x => dist.PDF(x), lower, upper, 10_000);
        integral.Should().BeApproximately(1.0, 1e-3);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtBoundaries()
    {
        var dist = new LognormalDistribution(4.6, 0.3, Seed);
        dist.CDF(0).Should().BeApproximately(0, 1e-10);
        dist.CDF(dist.Mean + 10 * dist.StdDev).Should().BeApproximately(1.0, 1e-3);
    }

    // --- All samples positive ---

    [Fact]
    public void AllSamples_ArePositive()
    {
        var dist = new LognormalDistribution(4.6, 0.3, Seed);
        var samples = dist.Sample(SampleCount);
        samples.Should().OnlyContain(s => s > 0);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new LognormalDistribution(4.6, 0.3, 123);
        var dist2 = new LognormalDistribution(4.6, 0.3, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new LognormalDistribution(4.6, 0.3);
        dist.ParameterSummary().Should().Be("Lognormal(μ=4.6, σ=0.3)");
    }

    private static double NumericalIntegrate(Func<double, double> f, double a, double b, int n)
    {
        double h = (b - a) / n;
        double sum = 0.5 * (f(a) + f(b));
        for (int i = 1; i < n; i++)
            sum += f(a + i * h);
        return sum * h;
    }
}
