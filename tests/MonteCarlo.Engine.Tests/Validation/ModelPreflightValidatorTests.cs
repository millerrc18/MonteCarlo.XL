using FluentAssertions;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.Engine.Validation;
using Xunit;

namespace MonteCarlo.Engine.Tests.Validation;

public class ModelPreflightValidatorTests
{
    [Fact]
    public void Validate_ValidProfile_HasNoErrors()
    {
        var report = ModelPreflightValidator.Validate(CreateValidProfile());

        report.HasErrors.Should().BeFalse();
        report.ErrorCount.Should().Be(0);
    }

    [Fact]
    public void Validate_NoInputsOrOutputs_ReturnsBlockingErrors()
    {
        var report = ModelPreflightValidator.Validate(new SimulationProfile());

        report.HasErrors.Should().BeTrue();
        report.Issues.Should().Contain(i => i.Code == "NO_INPUTS");
        report.Issues.Should().Contain(i => i.Code == "NO_OUTPUTS");
    }

    [Fact]
    public void Validate_InvalidDistribution_ReturnsError()
    {
        var profile = CreateValidProfile();
        profile.Inputs[0].Parameters["stdDev"] = 0;

        var report = ModelPreflightValidator.Validate(profile);

        report.HasErrors.Should().BeTrue();
        report.Issues.Should().Contain(i => i.Code == "DISTRIBUTION_INVALID");
    }

    [Fact]
    public void Validate_DuplicateInputCells_ReturnsError()
    {
        var profile = CreateValidProfile();
        profile.Inputs.Add(new SavedInput
        {
            SheetName = "Sheet1",
            CellAddress = "B2",
            Label = "Duplicate demand",
            DistributionName = "Normal",
            Parameters = new Dictionary<string, double>
            {
                ["mean"] = 10,
                ["stdDev"] = 1
            }
        });

        var report = ModelPreflightValidator.Validate(profile);

        report.HasErrors.Should().BeTrue();
        report.Issues.Should().Contain(i => i.Code == "INPUT_DUPLICATE_CELL");
    }

    [Fact]
    public void Validate_CellConfiguredAsInputAndOutput_ReturnsError()
    {
        var profile = CreateValidProfile();
        profile.Outputs[0].CellAddress = "B2";

        var report = ModelPreflightValidator.Validate(profile);

        report.HasErrors.Should().BeTrue();
        report.Issues.Should().Contain(i => i.Code == "CELL_IS_INPUT_AND_OUTPUT");
    }

    [Fact]
    public void Validate_CorrelationSizeMismatch_ReturnsError()
    {
        var profile = CreateValidProfile();
        profile.CorrelationMatrix = new double[,]
        {
            { 1, 0 },
            { 0, 1 }
        };

        var report = ModelPreflightValidator.Validate(profile);

        report.HasErrors.Should().BeTrue();
        report.Issues.Should().Contain(i => i.Code == "CORRELATION_SIZE_MISMATCH");
    }

    [Fact]
    public void Validate_LargeRun_ReturnsWarning()
    {
        var profile = CreateValidProfile();
        profile.IterationCount = 5_000_000;
        for (var i = 0; i < 5; i++)
        {
            profile.Outputs.Add(new SavedOutput
            {
                SheetName = "Sheet1",
                CellAddress = $"B{10 + i}",
                Label = $"Output {i + 2}"
            });
        }

        var report = ModelPreflightValidator.Validate(profile);

        report.WarningCount.Should().BeGreaterThan(0);
        report.Issues.Should().Contain(i => i.Code == "ITERATIONS_HIGH");
        report.Issues.Should().Contain(i => i.Code == "RESULT_SIZE_LARGE");
    }

    private static SimulationProfile CreateValidProfile()
    {
        return new SimulationProfile
        {
            IterationCount = 5_000,
            Inputs =
            {
                new SavedInput
                {
                    SheetName = "Sheet1",
                    CellAddress = "B2",
                    Label = "Demand",
                    DistributionName = "Normal",
                    Parameters = new Dictionary<string, double>
                    {
                        ["mean"] = 100,
                        ["stdDev"] = 15
                    }
                }
            },
            Outputs =
            {
                new SavedOutput
                {
                    SheetName = "Sheet1",
                    CellAddress = "B9",
                    Label = "Profit"
                }
            }
        };
    }
}
