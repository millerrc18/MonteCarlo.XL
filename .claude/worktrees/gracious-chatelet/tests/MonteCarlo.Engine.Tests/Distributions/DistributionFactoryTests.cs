using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

public class DistributionFactoryTests
{
    // --- Available Distributions ---

    [Fact]
    public void AvailableDistributions_ContainsAllTen()
    {
        var available = DistributionFactory.AvailableDistributions;
        available.Should().HaveCount(10);
        available.Should().Contain("Normal");
        available.Should().Contain("Triangular");
        available.Should().Contain("PERT");
        available.Should().Contain("Lognormal");
        available.Should().Contain("Uniform");
        available.Should().Contain("Discrete");
        available.Should().Contain("Beta");
        available.Should().Contain("Weibull");
        available.Should().Contain("Exponential");
        available.Should().Contain("Poisson");
    }

    // --- Creates Each Type ---

    [Fact]
    public void Create_Normal()
    {
        var dist = DistributionFactory.Create("Normal", new Dictionary<string, double>
        {
            { "mean", 100 },
            { "stdDev", 10 }
        });
        dist.Should().BeOfType<NormalDistribution>();
        dist.Mean.Should().BeApproximately(100, 1e-10);
    }

    [Fact]
    public void Create_Triangular()
    {
        var dist = DistributionFactory.Create("Triangular", new Dictionary<string, double>
        {
            { "min", 0 },
            { "mode", 5 },
            { "max", 10 }
        });
        dist.Should().BeOfType<TriangularDistribution>();
        dist.Mean.Should().BeApproximately(5.0, 1e-10);
    }

    [Fact]
    public void Create_PERT()
    {
        var dist = DistributionFactory.Create("PERT", new Dictionary<string, double>
        {
            { "min", 0 },
            { "mode", 3 },
            { "max", 6 }
        });
        dist.Should().BeOfType<PERTDistribution>();
        dist.Mean.Should().BeApproximately(3.0, 1e-10);
    }

    [Fact]
    public void Create_PERT_WithCustomLambda()
    {
        var dist = DistributionFactory.Create("PERT", new Dictionary<string, double>
        {
            { "min", 0 },
            { "mode", 5 },
            { "max", 10 },
            { "lambda", 6 }
        });
        dist.Should().BeOfType<PERTDistribution>();
        double expectedMean = (0 + 6 * 5 + 10) / 8.0;
        dist.Mean.Should().BeApproximately(expectedMean, 1e-10);
    }

    [Fact]
    public void Create_Lognormal()
    {
        var dist = DistributionFactory.Create("Lognormal", new Dictionary<string, double>
        {
            { "mu", 4.6 },
            { "sigma", 0.3 }
        });
        dist.Should().BeOfType<LognormalDistribution>();
    }

    [Fact]
    public void Create_Uniform()
    {
        var dist = DistributionFactory.Create("Uniform", new Dictionary<string, double>
        {
            { "min", 0 },
            { "max", 10 }
        });
        dist.Should().BeOfType<UniformDistribution>();
        dist.Mean.Should().BeApproximately(5.0, 1e-10);
    }

    [Fact]
    public void Create_Discrete()
    {
        var dist = DistributionFactory.Create("Discrete", new Dictionary<string, double>
        {
            { "value_0", 10 },
            { "prob_0", 0.3 },
            { "value_1", 20 },
            { "prob_1", 0.7 }
        });
        dist.Should().BeOfType<DiscreteDistribution>();
        double expectedMean = 10 * 0.3 + 20 * 0.7;
        dist.Mean.Should().BeApproximately(expectedMean, 1e-10);
    }

    [Fact]
    public void Create_Beta()
    {
        var dist = DistributionFactory.Create("Beta", new Dictionary<string, double>
        {
            { "alpha", 2 },
            { "beta", 5 }
        });
        dist.Should().BeOfType<BetaDistribution>();
        dist.Mean.Should().BeApproximately(2.0 / 7.0, 1e-10);
    }

    [Fact]
    public void Create_Weibull()
    {
        var dist = DistributionFactory.Create("Weibull", new Dictionary<string, double>
        {
            { "shape", 2 },
            { "scale", 100 }
        });
        dist.Should().BeOfType<WeibullDistribution>();
    }

    [Fact]
    public void Create_Exponential()
    {
        var dist = DistributionFactory.Create("Exponential", new Dictionary<string, double>
        {
            { "rate", 0.5 }
        });
        dist.Should().BeOfType<ExponentialDistribution>();
        dist.Mean.Should().BeApproximately(2.0, 1e-10);
    }

    [Fact]
    public void Create_Poisson()
    {
        var dist = DistributionFactory.Create("Poisson", new Dictionary<string, double>
        {
            { "lambda", 4.5 }
        });
        dist.Should().BeOfType<PoissonDistribution>();
        dist.Mean.Should().BeApproximately(4.5, 1e-10);
    }

    // --- Case Insensitive ---

    [Theory]
    [InlineData("normal")]
    [InlineData("NORMAL")]
    [InlineData("Normal")]
    [InlineData("nOrMaL")]
    public void Create_IsCaseInsensitive(string name)
    {
        var dist = DistributionFactory.Create(name, new Dictionary<string, double>
        {
            { "mean", 0 },
            { "stdDev", 1 }
        });
        dist.Should().BeOfType<NormalDistribution>();
    }

    // --- Unknown Name Throws ---

    [Fact]
    public void Create_UnknownName_Throws()
    {
        var act = () => DistributionFactory.Create("Gamma", new Dictionary<string, double>());
        act.Should().Throw<ArgumentException>().WithMessage("*Unknown distribution*Gamma*");
    }

    // --- Missing Parameters Throws ---

    [Fact]
    public void Create_MissingParameters_Throws()
    {
        var act = () => DistributionFactory.Create("Normal", new Dictionary<string, double>
        {
            { "mean", 0 }
            // missing stdDev
        });
        act.Should().Throw<ArgumentException>().WithMessage("*stdDev*");
    }

    // --- Seed Is Passed Through ---

    [Fact]
    public void Create_WithSeed_ProducesReproducibleResults()
    {
        var dist1 = DistributionFactory.Create("Normal",
            new Dictionary<string, double> { { "mean", 0 }, { "stdDev", 1 } }, seed: 42);
        var dist2 = DistributionFactory.Create("Normal",
            new Dictionary<string, double> { { "mean", 0 }, { "stdDev", 1 } }, seed: 42);

        var samples1 = dist1.Sample(50);
        var samples2 = dist2.Sample(50);
        samples1.Should().Equal(samples2);
    }

    // --- Discrete With No Pairs Throws ---

    [Fact]
    public void Create_Discrete_NoPairs_Throws()
    {
        var act = () => DistributionFactory.Create("Discrete", new Dictionary<string, double>());
        act.Should().Throw<ArgumentException>();
    }
}
