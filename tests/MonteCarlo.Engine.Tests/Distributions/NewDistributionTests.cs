using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class NewDistributionTests
{
    private const int SampleCount = 50_000;
    private const int Seed = 42;

    // ==================== Gamma Distribution ====================

    [Fact]
    public void Gamma_Constructor_ValidParameters_Succeeds()
    {
        var dist = new GammaDistribution(2.0, 0.5, Seed);
        dist.Name.Should().Be("Gamma");
        dist.Mean.Should().BeApproximately(4.0, 1e-10); // shape/rate = 2/0.5
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(-1, 1)]
    [InlineData(1, 0)]
    [InlineData(1, -1)]
    public void Gamma_Constructor_InvalidParameters_Throws(double shape, double rate)
    {
        var act = () => new GammaDistribution(shape, rate);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Gamma_Sample_ConvergesToTheoreticalMean()
    {
        double shape = 3.0, rate = 1.5;
        var dist = new GammaDistribution(shape, rate, Seed);
        var samples = dist.Sample(SampleCount);
        double expected = shape / rate;
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(expected, expected * 0.02);
    }

    [Fact]
    public void Gamma_Sample_ConvergesToTheoreticalStdDev()
    {
        double shape = 3.0, rate = 1.5;
        var dist = new GammaDistribution(shape, rate, Seed);
        var samples = dist.Sample(SampleCount);
        double mean = samples.Average();
        double empiricalStdDev = Math.Sqrt(samples.Sum(s => (s - mean) * (s - mean)) / (samples.Length - 1));
        double expected = Math.Sqrt(shape / (rate * rate));
        empiricalStdDev.Should().BeApproximately(expected, expected * 0.05);
    }

    [Theory]
    [InlineData(0.01)]
    [InlineData(0.1)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    [InlineData(0.99)]
    public void Gamma_CDF_Percentile_RoundTrip(double p)
    {
        var dist = new GammaDistribution(3.0, 1.5, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    [Fact]
    public void Gamma_DistributionFactory_Creates()
    {
        var parameters = new Dictionary<string, double> { ["shape"] = 2.0, ["rate"] = 1.0 };
        var dist = DistributionFactory.Create("Gamma", parameters, Seed);
        dist.Should().BeOfType<GammaDistribution>();
        dist.Name.Should().Be("Gamma");
    }

    // ==================== Logistic Distribution ====================

    [Fact]
    public void Logistic_Constructor_ValidParameters_Succeeds()
    {
        var dist = new LogisticDistribution(5.0, 2.0, Seed);
        dist.Name.Should().Be("Logistic");
        dist.Mean.Should().Be(5.0);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Logistic_Constructor_InvalidScale_Throws(double s)
    {
        var act = () => new LogisticDistribution(0, s);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Logistic_Sample_ConvergesToTheoreticalMean()
    {
        double mu = 10.0, s = 3.0;
        var dist = new LogisticDistribution(mu, s, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(mu, Math.Abs(mu) * 0.02 + 0.1);
    }

    [Fact]
    public void Logistic_Sample_ConvergesToTheoreticalStdDev()
    {
        double mu = 10.0, s = 3.0;
        var dist = new LogisticDistribution(mu, s, Seed);
        var samples = dist.Sample(SampleCount);
        double mean = samples.Average();
        double empiricalStdDev = Math.Sqrt(samples.Sum(x => (x - mean) * (x - mean)) / (samples.Length - 1));
        double expectedStdDev = s * Math.PI / Math.Sqrt(3.0);
        empiricalStdDev.Should().BeApproximately(expectedStdDev, expectedStdDev * 0.05);
    }

    [Theory]
    [InlineData(0.01)]
    [InlineData(0.1)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    [InlineData(0.99)]
    public void Logistic_CDF_Percentile_RoundTrip(double p)
    {
        var dist = new LogisticDistribution(5.0, 2.0, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    [Fact]
    public void Logistic_DistributionFactory_Creates()
    {
        var parameters = new Dictionary<string, double> { ["mu"] = 0.0, ["s"] = 1.0 };
        var dist = DistributionFactory.Create("Logistic", parameters, Seed);
        dist.Should().BeOfType<LogisticDistribution>();
        dist.Name.Should().Be("Logistic");
    }

    // ==================== GEV Distribution ====================

    [Fact]
    public void GEV_Constructor_ValidParameters_Succeeds()
    {
        var dist = new GEVDistribution(0.0, 1.0, 0.0, Seed);
        dist.Name.Should().Be("GEV");
    }

    [Fact]
    public void GEV_Constructor_InvalidSigma_Throws()
    {
        var act = () => new GEVDistribution(0, 0, 0);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void GEV_Gumbel_Sample_ConvergesToTheoreticalMean()
    {
        // Gumbel (xi=0): mean = mu + sigma * euler_mascheroni
        double mu = 5.0, sigma = 2.0, xi = 0.0;
        var dist = new GEVDistribution(mu, sigma, xi, Seed);
        var samples = dist.Sample(SampleCount);
        double expected = mu + sigma * 0.5772156649;
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(expected, Math.Abs(expected) * 0.02 + 0.1);
    }

    [Fact]
    public void GEV_Frechet_Sample_ConvergesToTheoreticalMean()
    {
        // Frechet (xi>0, xi<1): mean = mu + sigma * (Gamma(1-xi) - 1) / xi
        double mu = 0.0, sigma = 1.0, xi = 0.3;
        var dist = new GEVDistribution(mu, sigma, xi, Seed);
        var samples = dist.Sample(SampleCount);
        double expected = dist.Mean;
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(expected, Math.Abs(expected) * 0.05 + 0.2);
    }

    [Theory]
    [InlineData(0.01)]
    [InlineData(0.1)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    [InlineData(0.99)]
    public void GEV_Gumbel_CDF_Percentile_RoundTrip(double p)
    {
        var dist = new GEVDistribution(5.0, 2.0, 0.0, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    [Theory]
    [InlineData(0.01)]
    [InlineData(0.1)]
    [InlineData(0.5)]
    [InlineData(0.9)]
    [InlineData(0.99)]
    public void GEV_Frechet_CDF_Percentile_RoundTrip(double p)
    {
        var dist = new GEVDistribution(0.0, 1.0, 0.3, Seed);
        double value = dist.Percentile(p);
        double roundTrip = dist.CDF(value);
        roundTrip.Should().BeApproximately(p, 1e-6);
    }

    [Fact]
    public void GEV_DistributionFactory_Creates()
    {
        var parameters = new Dictionary<string, double> { ["mu"] = 0.0, ["sigma"] = 1.0, ["xi"] = 0.0 };
        var dist = DistributionFactory.Create("GEV", parameters, Seed);
        dist.Should().BeOfType<GEVDistribution>();
        dist.Name.Should().Be("GEV");
    }

    // ==================== Binomial Distribution ====================

    [Fact]
    public void Binomial_Constructor_ValidParameters_Succeeds()
    {
        var dist = new BinomialDistribution(10, 0.5, Seed);
        dist.Name.Should().Be("Binomial");
        dist.Mean.Should().BeApproximately(5.0, 1e-10);
    }

    [Theory]
    [InlineData(0, 0.5)]
    [InlineData(-1, 0.5)]
    [InlineData(10, -0.1)]
    [InlineData(10, 1.1)]
    public void Binomial_Constructor_InvalidParameters_Throws(int n, double p)
    {
        var act = () => new BinomialDistribution(n, p);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Binomial_Sample_ConvergesToTheoreticalMean()
    {
        int n = 20;
        double p = 0.3;
        var dist = new BinomialDistribution(n, p, Seed);
        var samples = dist.Sample(SampleCount);
        double expected = n * p;
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(expected, expected * 0.02 + 0.1);
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    public void Binomial_CDF_Percentile_RoundTrip(double p)
    {
        var dist = new BinomialDistribution(20, 0.5, Seed);
        double value = dist.Percentile(p);
        // For discrete distributions, verify CDF(Percentile(p)) >= p
        // and CDF(Percentile(p) - 1) < p (i.e. it's the smallest such integer)
        double cdfAtValue = dist.CDF(value);
        cdfAtValue.Should().BeGreaterThanOrEqualTo(p);
        if (value > 0)
        {
            double cdfBelow = dist.CDF(value - 1);
            cdfBelow.Should().BeLessThan(p + 1e-10);
        }
    }

    [Fact]
    public void Binomial_DistributionFactory_Creates()
    {
        var parameters = new Dictionary<string, double> { ["n"] = 10, ["p"] = 0.5 };
        var dist = DistributionFactory.Create("Binomial", parameters, Seed);
        dist.Should().BeOfType<BinomialDistribution>();
        dist.Name.Should().Be("Binomial");
    }

    // ==================== Geometric Distribution ====================

    [Fact]
    public void Geometric_Constructor_ValidParameters_Succeeds()
    {
        var dist = new GeometricDistribution(0.25, Seed);
        dist.Name.Should().Be("Geometric");
        dist.Mean.Should().BeApproximately(4.0, 1e-10);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-0.1)]
    [InlineData(1.1)]
    public void Geometric_Constructor_InvalidParameters_Throws(double p)
    {
        var act = () => new GeometricDistribution(p);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Geometric_Sample_ConvergesToTheoreticalMean()
    {
        double p = 0.2;
        var dist = new GeometricDistribution(p, Seed);
        var samples = dist.Sample(SampleCount);
        double expected = 1.0 / p;
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(expected, expected * 0.02 + 0.1);
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(0.75)]
    [InlineData(0.9)]
    public void Geometric_CDF_Percentile_RoundTrip(double p)
    {
        var dist = new GeometricDistribution(0.3, Seed);
        double value = dist.Percentile(p);
        // For discrete: CDF(Percentile(p)) >= p
        double cdfAtValue = dist.CDF(value);
        cdfAtValue.Should().BeGreaterThanOrEqualTo(p);
        if (value > 1)
        {
            double cdfBelow = dist.CDF(value - 1);
            cdfBelow.Should().BeLessThan(p + 1e-10);
        }
    }

    [Fact]
    public void Geometric_DistributionFactory_Creates()
    {
        var parameters = new Dictionary<string, double> { ["p"] = 0.5 };
        var dist = DistributionFactory.Create("Geometric", parameters, Seed);
        dist.Should().BeOfType<GeometricDistribution>();
        dist.Name.Should().Be("Geometric");
    }

    [Fact]
    public void Geometric_PEqualsOne_AlwaysReturnsOne()
    {
        var dist = new GeometricDistribution(1.0, Seed);
        var samples = dist.Sample(100);
        samples.Should().AllSatisfy(s => s.Should().Be(1.0));
    }
}
