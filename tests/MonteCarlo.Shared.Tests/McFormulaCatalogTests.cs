using MonteCarlo.Shared.Formula;

namespace MonteCarlo.Shared.Tests;

public class McFormulaCatalogTests
{
    [Fact]
    public void Definitions_ExposeExpectedWorksheetFunctions()
    {
        var functionNames = McFormulaCatalog.Definitions.Select(definition => definition.FunctionName).ToArray();

        Assert.Equal(
            [
                "Normal",
                "Triangular",
                "PERT",
                "Lognormal",
                "Uniform",
                "Beta",
                "Weibull",
                "Exponential",
                "Poisson",
                "Gamma",
                "Logistic",
                "GEV",
                "Binomial",
                "Geometric"
            ],
            functionNames);
    }
}
