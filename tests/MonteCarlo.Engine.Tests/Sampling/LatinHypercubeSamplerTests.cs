using FluentAssertions;
using MonteCarlo.Engine.Sampling;
using Xunit;

namespace MonteCarlo.Engine.Tests.Sampling;

public class LatinHypercubeSamplerTests
{
    private const int Seed = 42;

    [Theory]
    [InlineData(10, 3)]
    [InlineData(100, 5)]
    [InlineData(1, 1)]
    [InlineData(500, 1)]
    public void GenerateSamples_ReturnsCorrectDimensions(int iterations, int dimensions)
    {
        var sampler = new LatinHypercubeSampler(Seed);
        var samples = sampler.Generate(iterations, dimensions);

        samples.GetLength(0).Should().Be(iterations);
        samples.GetLength(1).Should().Be(dimensions);
    }

    [Theory]
    [InlineData(10)]
    [InlineData(100)]
    [InlineData(500)]
    public void GenerateSamples_EachDimensionCoversAllStrata(int iterations)
    {
        var sampler = new LatinHypercubeSampler(Seed);
        int dimensions = 3;
        var samples = sampler.Generate(iterations, dimensions);

        for (int d = 0; d < dimensions; d++)
        {
            // Track which stratum each sample falls into
            var strataHit = new bool[iterations];

            for (int i = 0; i < iterations; i++)
            {
                double value = samples[i, d];

                // Value must be in [0, 1)
                value.Should().BeGreaterThanOrEqualTo(0.0);
                value.Should().BeLessThan(1.0);

                // Determine which stratum this value belongs to
                int stratum = (int)(value * iterations);
                stratum.Should().BeInRange(0, iterations - 1);

                // This stratum should not have been hit yet (exactly one per stratum)
                strataHit[stratum].Should().BeFalse(
                    $"Stratum {stratum} was hit more than once in dimension {d}");
                strataHit[stratum] = true;
            }

            // All strata should be covered
            strataHit.Should().AllBeEquivalentTo(true,
                $"All strata should be hit exactly once in dimension {d}");
        }
    }

    [Fact]
    public void GenerateSamples_ValuesAreUniformOnUnitInterval()
    {
        int iterations = 10_000;
        int dimensions = 3;
        var sampler = new LatinHypercubeSampler(Seed);
        var samples = sampler.Generate(iterations, dimensions);

        for (int d = 0; d < dimensions; d++)
        {
            double sum = 0;
            for (int i = 0; i < iterations; i++)
                sum += samples[i, d];

            double mean = sum / iterations;
            // With 10k stratified samples, mean should be very close to 0.5
            mean.Should().BeApproximately(0.5, 0.01,
                $"Mean of dimension {d} should be approximately 0.5");
        }
    }

    [Fact]
    public void GenerateSamples_Reproducible_WithSameSeed()
    {
        int iterations = 100;
        int dimensions = 4;

        var sampler1 = new LatinHypercubeSampler(Seed);
        var samples1 = sampler1.Generate(iterations, dimensions);

        var sampler2 = new LatinHypercubeSampler(Seed);
        var samples2 = sampler2.Generate(iterations, dimensions);

        for (int i = 0; i < iterations; i++)
        {
            for (int d = 0; d < dimensions; d++)
            {
                samples1[i, d].Should().Be(samples2[i, d],
                    $"Sample [{i},{d}] should be identical with the same seed");
            }
        }
    }

    [Fact]
    public void GenerateSamples_DifferentSeeds_ProduceDifferentResults()
    {
        int iterations = 50;
        int dimensions = 2;

        var sampler1 = new LatinHypercubeSampler(42);
        var samples1 = sampler1.Generate(iterations, dimensions);

        var sampler2 = new LatinHypercubeSampler(99);
        var samples2 = sampler2.Generate(iterations, dimensions);

        // At least some values should differ
        bool anyDifferent = false;
        for (int i = 0; i < iterations && !anyDifferent; i++)
            for (int d = 0; d < dimensions && !anyDifferent; d++)
                if (Math.Abs(samples1[i, d] - samples2[i, d]) > 1e-15)
                    anyDifferent = true;

        anyDifferent.Should().BeTrue("Different seeds should produce different samples");
    }

    [Fact]
    public void GenerateSamples_InvalidIterations_Throws()
    {
        var sampler = new LatinHypercubeSampler(Seed);
        var act = () => sampler.Generate(0, 3);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void GenerateSamples_InvalidDimensions_Throws()
    {
        var sampler = new LatinHypercubeSampler(Seed);
        var act = () => sampler.Generate(10, 0);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }
}
