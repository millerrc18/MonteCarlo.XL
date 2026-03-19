using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using Xunit;

namespace MonteCarlo.Engine.Tests.Distributions;

/// <summary>
/// Placeholder test to verify project structure and test runner configuration.
/// Will be expanded with distribution-specific tests in TASK-002.
/// </summary>
public class IDistributionTests
{
    [Fact]
    public void IDistribution_Interface_ShouldExist()
    {
        // Verify the IDistribution interface is accessible from the test project
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
        interfaceType.GetMethod("Sample").Should().NotBeNull();
        interfaceType.GetMethod("Pdf").Should().NotBeNull();
        interfaceType.GetMethod("Cdf").Should().NotBeNull();
        interfaceType.GetMethod("Percentile").Should().NotBeNull();
    }
}
