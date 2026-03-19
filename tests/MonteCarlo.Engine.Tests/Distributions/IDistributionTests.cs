using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

/// <summary>
/// Tests that the IDistribution interface defines the expected contract.
/// </summary>
public class IDistributionTests
{
    [Fact]
    public void IDistribution_Interface_ShouldExist()
    {
        var interfaceType = typeof(IDistribution);
        interfaceType.Should().NotBeNull();
        interfaceType.IsInterface.Should().BeTrue();
    }

    [Fact]
    public void IDistribution_ShouldDefine_RequiredMembers()
    {
        var interfaceType = typeof(IDistribution);

        interfaceType.GetProperty("Name").Should().NotBeNull();
        interfaceType.GetProperty("Mean").Should().NotBeNull();
        interfaceType.GetProperty("Variance").Should().NotBeNull();
        interfaceType.GetProperty("StdDev").Should().NotBeNull();
        interfaceType.GetProperty("Minimum").Should().NotBeNull();
        interfaceType.GetProperty("Maximum").Should().NotBeNull();
        interfaceType.GetMethod("Sample", Type.EmptyTypes).Should().NotBeNull();
        interfaceType.GetMethod("Sample", new[] { typeof(int) }).Should().NotBeNull();
        interfaceType.GetMethod("PDF").Should().NotBeNull();
        interfaceType.GetMethod("CDF").Should().NotBeNull();
        interfaceType.GetMethod("Percentile").Should().NotBeNull();
        interfaceType.GetMethod("ParameterSummary").Should().NotBeNull();
    }
}
