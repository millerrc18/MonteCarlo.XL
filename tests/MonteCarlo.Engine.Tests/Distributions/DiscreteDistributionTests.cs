using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class DiscreteDistributionTests
{
    private const int SampleCount = 100_000;
    private const int Seed = 42;

    // --- Construction & Validation ---

    [Fact]
    public void Constructor_ValidParameters_Succeeds()
    {
        var dist = new DiscreteDistribution(
            new[] { 1.0, 2.0, 3.0 },
            new[] { 0.2, 0.5, 0.3 });
        dist.Name.Should().Be("Discrete");
    }

    [Fact]
    public void Constructor_EmptyArrays_Throws()
    {
        var act = () => new DiscreteDistribution(Array.Empty<double>(), Array.Empty<double>());
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Constructor_MismatchedLengths_Throws()
    {
        var act = () => new DiscreteDistribution(
            new[] { 1.0, 2.0 },
            new[] { 0.5, 0.3, 0.2 });
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Constructor_NegativeProbability_Throws()
    {
        var act = () => new DiscreteDistribution(
            new[] { 1.0, 2.0 },
            new[] { 1.5, -0.5 });
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void Constructor_ProbabilitiesDontSumToOne_Throws()
    {
        var act = () => new DiscreteDistribution(
            new[] { 1.0, 2.0 },
            new[] { 0.3, 0.3 });
        act.Should().Throw<ArgumentException>().WithMessage("*sum*");
    }

    [Fact]
    public void Constructor_NullValues_Throws()
    {
        var act = () => new DiscreteDistribution(null!, new[] { 1.0 });
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void Constructor_NullProbabilities_Throws()
    {
        var act = () => new DiscreteDistribution(new[] { 1.0 }, null!);
        act.Should().Throw<ArgumentNullException>();
    }

    // --- Statistical Properties ---

    [Fact]
    public void Mean_IsWeightedSum()
    {
        var dist = new DiscreteDistribution(
            new[] { 10.0, 20.0, 30.0 },
            new[] { 0.2, 0.5, 0.3 });
        double expected = 10 * 0.2 + 20 * 0.5 + 30 * 0.3;
        dist.Mean.Should().BeApproximately(expected, 1e-10);
    }

    [Fact]
    public void Variance_IsCorrect()
    {
        var dist = new DiscreteDistribution(
            new[] { 10.0, 20.0, 30.0 },
            new[] { 0.2, 0.5, 0.3 });
        double mean = 10 * 0.2 + 20 * 0.5 + 30 * 0.3;
        double expected = 0.2 * (10 - mean) * (10 - mean)
                        + 0.5 * (20 - mean) * (20 - mean)
                        + 0.3 * (30 - mean) * (30 - mean);
        dist.Variance.Should().BeApproximately(expected, 1e-10);
    }

    [Fact]
    public void MinMax_AreCorrect()
    {
        var dist = new DiscreteDistribution(
            new[] { 30.0, 10.0, 20.0 },
            new[] { 0.3, 0.2, 0.5 });
        dist.Minimum.Should().Be(10);
        dist.Maximum.Should().Be(30);
    }

    // --- Sampling Convergence ---

    [Fact]
    public void Sample_FrequenciesConvergeToProbabilities()
    {
        var values = new[] { 1.0, 2.0, 3.0 };
        var probs = new[] { 0.2, 0.5, 0.3 };
        var dist = new DiscreteDistribution(values, probs, Seed);

        var samples = dist.Sample(SampleCount);
        var counts = new Dictionary<double, int> { { 1.0, 0 }, { 2.0, 0 }, { 3.0, 0 } };
        foreach (var s in samples)
            counts[s]++;

        (counts[1.0] / (double)SampleCount).Should().BeApproximately(0.2, 0.01);
        (counts[2.0] / (double)SampleCount).Should().BeApproximately(0.5, 0.01);
        (counts[3.0] / (double)SampleCount).Should().BeApproximately(0.3, 0.01);
    }

    [Fact]
    public void Sample_ConvergesToTheoreticalMean()
    {
        var dist = new DiscreteDistribution(
            new[] { 10.0, 20.0, 30.0 },
            new[] { 0.2, 0.5, 0.3 }, Seed);
        var samples = dist.Sample(SampleCount);
        double empiricalMean = samples.Average();
        empiricalMean.Should().BeApproximately(dist.Mean, dist.Mean * 0.01 + 0.5);
    }

    // --- PDF ---

    [Fact]
    public void PDF_ReturnsCorrectProbabilities()
    {
        var dist = new DiscreteDistribution(
            new[] { 10.0, 20.0, 30.0 },
            new[] { 0.2, 0.5, 0.3 });
        dist.PDF(10).Should().BeApproximately(0.2, 1e-10);
        dist.PDF(20).Should().BeApproximately(0.5, 1e-10);
        dist.PDF(30).Should().BeApproximately(0.3, 1e-10);
    }

    [Fact]
    public void PDF_ReturnsZero_ForNonExistentValue()
    {
        var dist = new DiscreteDistribution(
            new[] { 10.0, 20.0, 30.0 },
            new[] { 0.2, 0.5, 0.3 });
        dist.PDF(15).Should().Be(0);
        dist.PDF(0).Should().Be(0);
        dist.PDF(100).Should().Be(0);
    }

    // --- CDF is a step function ---

    [Fact]
    public void CDF_IsStepFunction()
    {
        var dist = new DiscreteDistribution(
            new[] { 10.0, 20.0, 30.0 },
            new[] { 0.2, 0.5, 0.3 });

        dist.CDF(5).Should().Be(0);              // Below all values
        dist.CDF(10).Should().BeApproximately(0.2, 1e-10);   // At first value
        dist.CDF(15).Should().BeApproximately(0.2, 1e-10);   // Between values (same as CDF(10))
        dist.CDF(20).Should().BeApproximately(0.7, 1e-10);   // At second value
        dist.CDF(25).Should().BeApproximately(0.7, 1e-10);   // Between second and third
        dist.CDF(30).Should().BeApproximately(1.0, 1e-10);   // At last value
        dist.CDF(100).Should().BeApproximately(1.0, 1e-10);  // Above all values
    }

    // --- Percentile ---

    [Fact]
    public void Percentile_ReturnsCorrectValues()
    {
        var dist = new DiscreteDistribution(
            new[] { 10.0, 20.0, 30.0 },
            new[] { 0.2, 0.5, 0.3 });

        dist.Percentile(0.1).Should().Be(10);    // p=0.1 < 0.2 → first value
        dist.Percentile(0.2).Should().Be(10);    // p=0.2 = cumulative at first value
        dist.Percentile(0.3).Should().Be(20);    // p=0.3 → second value
        dist.Percentile(0.7).Should().Be(20);    // p=0.7 = cumulative at second value
        dist.Percentile(0.8).Should().Be(30);    // p=0.8 → third value
        dist.Percentile(1.0).Should().Be(30);    // p=1.0 → last value
    }

    // --- Reproducibility ---

    [Fact]
    public void SameSeed_ProducesIdenticalSequences()
    {
        var dist1 = new DiscreteDistribution(new[] { 1.0, 2.0, 3.0 }, new[] { 0.2, 0.5, 0.3 }, 123);
        var dist2 = new DiscreteDistribution(new[] { 1.0, 2.0, 3.0 }, new[] { 0.2, 0.5, 0.3 }, 123);
        var samples1 = dist1.Sample(100);
        var samples2 = dist2.Sample(100);
        samples1.Should().Equal(samples2);
    }

    // --- Unsorted input values are handled ---

    [Fact]
    public void UnsortedValues_AreHandledCorrectly()
    {
        var dist = new DiscreteDistribution(
            new[] { 30.0, 10.0, 20.0 },
            new[] { 0.3, 0.2, 0.5 });

        dist.PDF(10).Should().BeApproximately(0.2, 1e-10);
        dist.PDF(20).Should().BeApproximately(0.5, 1e-10);
        dist.PDF(30).Should().BeApproximately(0.3, 1e-10);
        dist.CDF(15).Should().BeApproximately(0.2, 1e-10);
    }

    // --- ParameterSummary ---

    [Fact]
    public void ParameterSummary_ContainsValuesAndProbabilities()
    {
        var dist = new DiscreteDistribution(
            new[] { 1.0, 2.0 },
            new[] { 0.4, 0.6 });
        var summary = dist.ParameterSummary();
        summary.Should().Contain("Discrete");
        summary.Should().Contain("1");
        summary.Should().Contain("2");
    }
}
