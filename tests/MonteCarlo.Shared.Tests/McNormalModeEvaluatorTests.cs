using MathNet.Numerics;
using MonteCarlo.Shared.Formula;

namespace MonteCarlo.Shared.Tests;

public class McNormalModeEvaluatorTests
{
    public static IEnumerable<object[]> ValidNormalModeCases()
    {
        yield return ["Normal", new double[] { 42, 3 }, 42d];
        yield return ["Triangular", new double[] { 1, 2, 5 }, 2d];
        yield return ["PERT", new double[] { 10, 12, 20 }, 12d];
        yield return ["Lognormal", new double[] { 1.2, 0.4 }, Math.Exp(1.2 + 0.4 * 0.4 / 2.0)];
        yield return ["Uniform", new double[] { 4, 10 }, 7d];
        yield return ["Beta", new double[] { 2, 3 }, 0.4d];
        yield return ["Weibull", new double[] { 1.5, 8 }, 8 * SpecialFunctions.Gamma(1.0 + 1.0 / 1.5)];
        yield return ["Exponential", new double[] { 0.25 }, 4d];
        yield return ["Poisson", new double[] { 12 }, 12d];
        yield return ["Gamma", new double[] { 6, 2 }, 3d];
        yield return ["Logistic", new double[] { 9, 1.5 }, 9d];
        yield return ["GEV", new double[] { 10, 2, 0 }, 10 + 2 * 0.5772156649];
        yield return ["Binomial", new double[] { 20, 0.35 }, 7d];
        yield return ["Geometric", new double[] { 0.2 }, 5d];
    }

    [Theory]
    [MemberData(nameof(ValidNormalModeCases))]
    public void TryEvaluate_ReturnsExpectedRepresentativeValue(string functionName, double[] arguments, double expected)
    {
        var success = McNormalModeEvaluator.TryEvaluate(functionName, arguments, out var result);

        Assert.True(success);
        Assert.Equal(expected, result, 10);
    }

    [Fact]
    public void TryEvaluate_ReturnsExpectedValueForLognormal()
    {
        var success = McNormalModeEvaluator.TryEvaluate("Lognormal", [1.2, 0.4], out var result);

        Assert.True(success);
        Assert.Equal(Math.Exp(1.2 + 0.4 * 0.4 / 2.0), result, 10);
    }

    [Fact]
    public void TryEvaluate_RejectsInvalidBinomialProbability()
    {
        var success = McNormalModeEvaluator.TryEvaluate("Binomial", [10, 1.4], out _);

        Assert.False(success);
    }
}
