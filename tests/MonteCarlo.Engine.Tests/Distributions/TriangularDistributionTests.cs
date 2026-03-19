using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class TriangularDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new TriangularDistribution(0, 5, 10);
        dist.Name.Should().Be("Triangular");
        dist.Minimum.Should().Be(0);
        dist.Maximum.Should().Be(10);
    }

    [Fact]
    public void Constructor_MinEqualsMax_Throws()
    {
        var act = () => new TriangularDistribution(5, 5, 5);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Constructor_MinGreaterThanMax_Throws()
    {
        var act = () => new TriangularDistribution(10, 5, 0);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Constructor_ModeOutsideRange_Throws()
    {
        var act = () => new TriangularDistribution(0, 15, 10);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("mode");
    }

    [Fact]
    public void Constructor_ModeEqualsMin_Succeeds()
    {
        var dist = new TriangularDistribution(0, 0, 10);
        dist.Mean.Should().BeApproximately(10.0 / 3.0, 1e-10);
    }

    [Fact]
    public void Constructor_ModeEqualsMax_Succeeds()
    {
        var dist = new TriangularDistribution(0, 10, 10);
        dist.Mean.Should().BeApproximately(20.0 / 3.0, 1e-10);
    }

    // --- Statistical Properties ---

    [Fact]
    public void Mean_IsCorrect()
    {
        var dist = new TriangularDistribution(0, 5, 10);
        dist.Mean.Should().BeApproximately(5.0, 1e-10);
    }

    [Fact]
    public void Variance_IsCorrect()
    {
        // Var = (a^2 + b^2 + c^2 - ab - ac - bc) / 18
        // For (0, 5, 10): (0 + 25 + 100 - 0 - 0 - 50) / 18 = 75/18
        var dist = new TriangularDistribution(0, 5, 10);
        dist.Variance.Should().BeApproximately(75.0 / 18.0, 1e-10);
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new TriangularDistribution(0, 5, 10, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(dist.Mean, dist.Mean * 0.01);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalStdDev()
    {
        var dist = new TriangularDistribution(0, 5, 10, Seed);
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
        var dist = new TriangularDistribution(0, 5, 10, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    // --- PDF Integrates to 1 ---

    [Fact]
    public void PDF_IntegratesToOne()
    {
        var dist = new TriangularDistribution(0, 5, 10, Seed);
        double integral = NumericalIntegrate(x => dist.PDF(x), 0, 10, 10_000);
        integral.Should().BeApproximately(1.0, 1e-3);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtBoundaries()
    {
        var dist = new TriangularDistribution(0, 5, 10, Seed);
        dist.CDF(0).Should().BeApproximately(0, 1e-10);
        dist.CDF(10).Should().BeApproximately(1, 1e-10);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new TriangularDistribution(0, 5, 10, 123);
        var dist2 = new TriangularDistribution(0, 5, 10, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new TriangularDistribution(0, 5, 10);
        dist.ParameterSummary().Should().Be("Triangular(min=0, mode=5, max=10)");
    }

    // --- All samples within bounds ---

    [Fact]
    public void AllSamples_WithinBounds()
    {
        var dist = new TriangularDistribution(0, 5, 10, Seed);
        var samples = dist.Sample(SampleCount);
        samples.Should().OnlyContain(s => s >= 0 && s <= 10);
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
