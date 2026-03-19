using FluentAssertions;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Engine.Tests.Simulation;

public class SimulationConfigTests
{
    [Fact]
    public void Validate_ValidConfig_DoesNotThrow()
    {
        var config = CreateValidConfig();
        var act = () => config.Validate();
        act.Should().NotThrow();
    }

    [Fact]
    public void Validate_EmptyInputs_Throws()
    {
        var config = CreateValidConfig();
        config.Inputs.Clear();
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*input*");
    }

    [Fact]
    public void Validate_NullInputs_Throws()
    {
        var config = CreateValidConfig();
        config.Inputs = null!;
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Validate_EmptyOutputs_Throws()
    {
        var config = CreateValidConfig();
        config.Outputs.Clear();
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*output*");
    }

    [Fact]
    public void Validate_ZeroIterations_Throws()
    {
        var config = CreateValidConfig();
        config.IterationCount = 0;
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*IterationCount*");
    }

    [Fact]
    public void Validate_NegativeIterations_Throws()
    {
        var config = CreateValidConfig();
        config.IterationCount = -1;
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*IterationCount*");
    }

    [Fact]
    public void Validate_DuplicateInputIds_Throws()
    {
        var config = CreateValidConfig();
        config.Inputs.Add(new SimulationInput
        {
            Id = config.Inputs[0].Id,
            Label = "Duplicate",
            Distribution = new NormalDistribution(0, 1)
        });
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*Duplicate*");
    }

    [Fact]
    public void Validate_DuplicateOutputIds_Throws()
    {
        var config = CreateValidConfig();
        config.Outputs.Add(new SimulationOutput
        {
            Id = config.Outputs[0].Id,
            Label = "Duplicate"
        });
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*Duplicate*");
    }

    [Fact]
    public void Validate_InputWithNullDistribution_Throws()
    {
        var config = CreateValidConfig();
        config.Inputs[0].Distribution = null!;
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*null distribution*");
    }

    [Fact]
    public void Validate_InputWithEmptyId_Throws()
    {
        var config = CreateValidConfig();
        config.Inputs[0].Id = "";
        var act = () => config.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*non-empty*");
    }

    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var config = new SimulationConfig();
        config.IterationCount.Should().Be(5000);
        config.RandomSeed.Should().BeNull();
        config.ParallelExecution.Should().BeTrue();
        config.Inputs.Should().BeEmpty();
        config.Outputs.Should().BeEmpty();
    }

    private static SimulationConfig CreateValidConfig()
    {
        return new SimulationConfig
        {
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A1", Label = "Cost", Distribution = new NormalDistribution(100, 10) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "B1", Label = "Profit" }
            },
            IterationCount = 1000
        };
    }
}
