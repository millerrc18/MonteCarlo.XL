using FluentAssertions;
using MonteCarlo.Engine.Analysis;
using Xunit;

namespace MonteCarlo.Engine.Tests.Analysis;

public class DistributionFitServiceTests
{
    [Fact]
    public void Fit_NormalLikeSamples_RanksNormalNearTop()
    {
        var samples = Enumerable.Range(-50, 101)
            .Select(i => 100 + i * 0.5)
            .ToArray();

        var results = DistributionFitService.Fit(samples, maxResults: 3);

        results.Should().NotBeEmpty();
        results.Select(r => r.DistributionName).Should().Contain("Normal");
        results[0].Score.Should().BeLessThan(0.2);
    }

    [Fact]
    public void Fit_PercentageSamples_IncludesBeta()
    {
        var samples = new[] { 0.08, 0.10, 0.11, 0.13, 0.16, 0.18, 0.20, 0.22, 0.24, 0.28 };

        var results = DistributionFitService.Fit(samples, maxResults: 5);

        results.Select(r => r.DistributionName).Should().Contain("Beta");
    }

    [Fact]
    public void Fit_CountSamples_IncludesDiscreteCountDistributions()
    {
        var samples = new double[] { 1, 2, 2, 3, 3, 4, 4, 4, 5, 6, 6, 7 };

        var results = DistributionFitService.Fit(samples, maxResults: 6);

        results.Select(r => r.DistributionName).Should().Contain(new[] { "Poisson", "Binomial" });
    }

    [Fact]
    public void Fit_TooFewSamples_Throws()
    {
        var act = () => DistributionFitService.Fit(new[] { 1.0, 2.0 });

        act.Should().Throw<ArgumentException>();
    }
}
