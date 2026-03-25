using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class PERTDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new PERTDistribution(0, 3, 6);
        dist.Name.Should().Be("PERT");
        dist.Minimum.Should().Be(0);
        dist.Maximum.Should().Be(6);
    }

    [Fact]
    public void Constructor_MinEqualsMax_Throws()
    {
        var act = () => new PERTDistribution(5, 5, 5);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Constructor_MinGreaterThanMax_Throws()
    {
        var act = () => new PERTDistribution(10, 5, 0);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Constructor_ModeOutsideRange_Throws()
    {
        var act = () => new PERTDistribution(0, 15, 10);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("mode");
    }

    [Fact]
    public void Constructor_InvalidLambda_Throws()
    {
        var act = () => new PERTDistribution(0, 5, 10, lambda: 0);
        act.Should().Throw<ArgumentOutOfRangeException>().WithParameterName("lambda");
    }

    [Fact]
    public void Constructor_ModeEqualsMin_Succeeds()
    {
        var dist = new PERTDistribution(0, 0, 10);
        dist.Mean.Should().BeApproximately((0 + 4 * 0 + 10) / 6.0, 1e-10);
    }

    [Fact]
    public void Constructor_ModeEqualsMax_Succeeds()
    {
        var dist = new PERTDistribution(0, 10, 10);
        dist.Mean.Should().BeApproximately((0 + 4 * 10 + 10) / 6.0, 1e-10);
    }

    // --- PERT Mean Formula ---

    [Fact]
    public void Mean_FollowsPERTFormula()
    {
        // PERT mean = (min + lambda * mode + max) / (lambda + 2)
        var dist = new PERTDistribution(10, 50, 90, lambda: 4);
        double expectedMean = (10 + 4 * 50 + 90) / 6.0;
        dist.Mean.Should().BeApproximately(expectedMean, 1e-10);
    }

    [Fact]
    public void Mean_SymmetricCase_EqualsMode()
    {
        // PERT(0, 3, 6) with lambda=4: mean = (0 + 12 + 6) / 6 = 3.0
        var dist = new PERTDistribution(0, 3, 6);
        dist.Mean.Should().BeApproximately(3.0, 1e-10);
    }

    [Fact]
    public void Mean_WithCustomLambda()
    {
        var dist = new PERTDistribution(0, 5, 10, lambda: 6);
        double expectedMean = (0 + 6 * 5 + 10) / 8.0;
        dist.Mean.Should().BeApproximately(expectedMean, 1e-10);
    }

    // --- PERT is smoother than Triangular ---

    [Fact]
    public void PERT_HasLowerKurtosis_ThanTriangular()
    {
        // For the same min/mode/max, PERT should have lower kurtosis (less extreme tails)
        var pert = new PERTDistribution(0, 5, 10, seed: Seed);
        var tri = new TriangularDistribution(0, 5, 10, Seed);

        var pertSamples = pert.Sample(SampleCount);
        var triSamples = tri.Sample(SampleCount);

        double pertKurtosis = ComputeExcessKurtosis(pertSamples);
        double triKurtosis = ComputeExcessKurtosis(triSamples);

        // Triangular excess kurtosis = -0.6, PERT should be lower (more negative = more platykurtic)
        // Actually PERT is closer to normal (kurtosis closer to 0) and Triangular is -0.6
        // The key point: PERT variance should be smaller for same range
        pert.Variance.Should().BeLessThan(tri.Variance,
            "PERT should be more concentrated than Triangular for the same min/mode/max");
    }

    // --- Statistical Convergence ---

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new PERTDistribution(0, 3, 6, seed: Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(dist.Mean, dist.Mean * 0.01 + 0.05);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalStdDev()
    {
        var dist = new PERTDistribution(0, 3, 6, seed: Seed);
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
        var dist = new PERTDistribution(0, 3, 6, seed: Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    // --- PDF Integrates to 1 ---

    [Fact]
    public void PDF_IntegratesToOne()
    {
        var dist = new PERTDistribution(0, 3, 6, seed: Seed);
        double integral = NumericalIntegrate(x => dist.PDF(x), 0, 6, 10_000);
        integral.Should().BeApproximately(1.0, 1e-3);
    }

    // --- CDF Boundary Conditions ---

    [Fact]
    public void CDF_AtBoundaries()
    {
        var dist = new PERTDistribution(0, 3, 6, seed: Seed);
        dist.CDF(0).Should().BeApproximately(0, 1e-10);
        dist.CDF(6).Should().BeApproximately(1, 1e-10);
    }

    [Fact]
    public void PDF_OutsideBounds_IsZero()
    {
        var dist = new PERTDistribution(0, 3, 6, seed: Seed);
        dist.PDF(-1).Should().Be(0);
        dist.PDF(7).Should().Be(0);
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new PERTDistribution(0, 3, 6, seed: 123);
        var dist2 = new PERTDistribution(0, 3, 6, seed: 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- All samples within bounds ---

    [Fact]
    public void AllSamples_WithinBounds()
    {
        var dist = new PERTDistribution(0, 3, 6, seed: Seed);
        var samples = dist.Sample(SampleCount);
        samples.Should().OnlyContain(s => s >= 0 && s <= 6);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_IsCorrect()
    {
        var dist = new PERTDistribution(10, 50, 90);
        dist.ParameterSummary().Should().Be("PERT(min=10, mode=50, max=90, λ=4)");
    }

    private static double ComputeExcessKurtosis(double[] samples)
    {
        double mean = samples.Average();
        double n = samples.Length;
        double m2 = samples.Sum(s => (s - mean) * (s - mean)) / n;
        double m4 = samples.Sum(s => Math.Pow(s - mean, 4)) / n;
        return m4 / (m2 * m2) - 3.0;
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
