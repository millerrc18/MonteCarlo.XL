using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class NormalDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new NormalDistribution(100, 10);
        dist.Mean.Should().Be(100);
        dist.StdDev.Should().Be(10);
        dist.Name.Should().Be("Normal");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(-0.001)]
    public void Constructor_InvalidStdDev_Throws(double stdDev)
    {
        var act = () => new NormalDistribution(0, stdDev);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("stdDev");
    }

    // --- Statistical Properties ---

    [Fact]
    public void Properties_AreCorrect()
    {
        var dist = new NormalDistribution(50, 15);
        dist.Variance.Should().BeApproximately(225, 1e-10);
        dist.StdDev.Should().BeApproximately(15, 1e-10);
        dist.Minimum.Should().Be(double.NegativeInfinity);
        dist.Maximum.Should().Be(double.PositiveInfinity);
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new NormalDistribution(100, 10, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(100, 100 * 0.01);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalStdDev()
    {
        var dist = new NormalDistribution(100, 10, Seed);
        var samples = dist.Sample(SampleCount);
        double mean = samples.Average();
        double empiricalStdDev = Math.Sqrt(samples.Sum(s => (s - mean) * (s - mean)) / (samples.Length - 1));
        empiricalStdDev.Should().BeApproximately(10, 10 * 0.02);
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
        var dist = new NormalDistribution(100, 10, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    // --- PDF Integrates to 1 ---

    [Fact]
    public void PDF_IntegratesToOne()
    {
        var dist = new NormalDistribution(100, 10, Seed);
        double lower = dist.Mean - 6 * dist.StdDev;
        double upper = dist.Mean + 6 * dist.StdDev;
        double integral = NumericalIntegrate(x => dist.PDF(x), lower, upper, 10_000);
        integral.Should().BeApproximately(1.0, 1e-3);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtExtremes()
    {
        var dist = new NormalDistribution(100, 10, Seed);
        dist.CDF(100 - 6 * 10).Should().BeApproximately(0, 1e-6);
        dist.CDF(100 + 6 * 10).Should().BeApproximately(1, 1e-6);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new NormalDistribution(100, 10, 123);
        var dist2 = new NormalDistribution(100, 10, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- Batch Sample ---

    [Fact]
    public void Sample_BatchReturnsCorrectCount()
    {
        var dist = new NormalDistribution(0, 1, Seed);
        dist.Sample(500).Should().HaveCount(500);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new NormalDistribution(100, 10);
        dist.ParameterSummary().Should().Be("Normal(μ=100, σ=10)");
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
