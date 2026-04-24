using MonteCarlo.Shared.Formula;

namespace MonteCarlo.Shared.Tests;

public class McFormulaParserTests
{
    [Fact]
    public void TryParse_RecognizesKnownFunction()
    {
        var success = McFormulaParser.TryParse("=MC.Gamma(4, 2)", out var parsed);

        Assert.True(success);
        Assert.NotNull(parsed);
        Assert.Equal("Gamma", parsed!.FunctionName);
        Assert.Equal(["4", "2"], parsed.RawArguments);
        Assert.Equal(["shape", "rate"], parsed.Arguments.Select(argument => argument.Name).ToArray());
    }

    [Fact]
    public void TryParse_SplitsNestedArgumentsWithoutBreakingOnInnerCommas()
    {
        var success = McFormulaParser.TryParse("=MC.Normal(AVERAGE(A1,B1), 10)", out var parsed);

        Assert.True(success);
        Assert.NotNull(parsed);
        Assert.Equal(["AVERAGE(A1,B1)", "10"], parsed!.RawArguments);
    }
}
