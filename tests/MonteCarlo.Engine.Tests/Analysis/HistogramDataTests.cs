using Xunit;
using FluentAssertions;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.Engine.Tests.Analysis;

public class HistogramDataTests
{
    [Fact]
    public void UniformDistribution_RoughlyEqualBins()
    {
        var dist = new UniformDistribution(0, 1, seed: 42);
        var samples = dist.Sample(100_000);
        Array.Sort(samples);
        var hist = new HistogramData(samples, 10);

        // Each bin should have roughly 10% of samples
        foreach (var freq in hist.Frequencies)
        {
            freq.Should().BeInRange(8000, 12000,
                "uniform distribution should produce roughly equal bin frequencies");
        }
    }

    [Fact]
    public void CustomBinCount_IsRespected()
    {
        var samples = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
        var hist = new HistogramData(samples, 5);

        hist.BinCenters.Length.Should().Be(5);
        hist.Frequencies.Length.Should().Be(5);
        hist.BinEdges.Length.Should().Be(6);
    }

    [Fact]
    public void SingleValue_HandledGracefully()
    {
        var samples = new double[] { 5, 5, 5, 5, 5 };
        var hist = new HistogramData(samples, 10);

        hist.Frequencies.Sum().Should().Be(5);
        hist.BinCenters.Length.Should().Be(1); // collapsed
    }

    [Fact]
    public void RelativeFrequencies_SumToOne()
    {
        var dist = new NormalDistribution(0, 1, seed: 42);
        var samples = dist.Sample(10_000);
        Array.Sort(samples);
        var hist = new HistogramData(samples, 50);

        hist.RelativeFrequencies.Sum().Should().BeApproximately(1.0, 1e-10);
    }

    [Fact]
    public void EmptyArray_Throws()
    {
        var act = () => new HistogramData(Array.Empty<double>());
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void ZeroBinCount_Throws()
    {
        var act = () => new HistogramData(new[] { 1.0, 2.0 }, 0);
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void BinEdges_AreMonotonicallyIncreasing()
    {
        var dist = new NormalDistribution(0, 1, seed: 42);
        var samples = dist.Sample(1000);
        Array.Sort(samples);
        var hist = new HistogramData(samples, 20);

        for (int i = 1; i < hist.BinEdges.Length; i++)
            hist.BinEdges[i].Should().BeGreaterThanOrEqualTo(hist.BinEdges[i - 1]);
    }
}
