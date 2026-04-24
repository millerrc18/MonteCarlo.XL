using MonteCarlo.Engine.Simulation;
using MonteCarlo.Shared.Interop;
using MonteCarlo.Shared.Simulation;

namespace MonteCarlo.Shared.Tests;

public class SimulationSamplingServiceTests
{
    [Fact]
    public void GenerateInputMatrix_ReturnsStableMatrixForFixedSeed()
    {
        var service = new SimulationSamplingService();
        var profile = new SimulationProfile
        {
            Name = "Smoke",
            Inputs =
            [
                new SavedInput
                {
                    SheetName = "Sheet1",
                    CellAddress = "A1",
                    Label = "A1",
                    DistributionName = "Normal",
                    Parameters = new Dictionary<string, double>
                    {
                        ["mean"] = 100,
                        ["stdDev"] = 10
                    }
                }
            ],
            Outputs =
            [
                new SavedOutput
                {
                    SheetName = "Sheet1",
                    CellAddress = "B1",
                    Label = "B1"
                }
            ]
        };

        var request = new GenerateInputMatrixRequest(
            BridgeProtocol.Version,
            profile,
            new OfficeRunSettings(5, 42, SamplingMethod.LatinHypercube, false));

        var first = service.GenerateInputMatrix(request);
        var second = service.GenerateInputMatrix(request);

        Assert.Equal(first.InputIds, second.InputIds);
        Assert.Equal(first.OutputIds, second.OutputIds);
        Assert.Equal(first.InputMatrix.Length, second.InputMatrix.Length);
        Assert.Equal(first.InputMatrix[0].Length, second.InputMatrix[0].Length);

        for (var row = 0; row < first.InputMatrix.Length; row++)
        {
            for (var column = 0; column < first.InputMatrix[row].Length; column++)
            {
                Assert.Equal(first.InputMatrix[row][column], second.InputMatrix[row][column], 10);
            }
        }
    }
}
