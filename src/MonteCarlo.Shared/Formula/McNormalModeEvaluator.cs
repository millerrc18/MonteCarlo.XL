using MathNet.Numerics;

namespace MonteCarlo.Shared.Formula;

public static class McNormalModeEvaluator
{
    public static bool TryEvaluate(string functionName, IReadOnlyList<double> arguments, out double result)
    {
        result = 0;
        if (!McFormulaCatalog.TryGetByFunctionName(functionName, out var definition))
            return false;

        if (arguments.Count != definition.Arguments.Count)
            return false;

        switch (definition.FunctionName.ToLowerInvariant())
        {
            case "normal":
                result = arguments[0];
                return arguments[1] > 0;

            case "triangular":
            case "pert":
                result = arguments[1];
                return arguments[0] < arguments[2]
                    && arguments[1] >= arguments[0]
                    && arguments[1] <= arguments[2];

            case "lognormal":
                if (arguments[1] <= 0)
                    return false;
                result = Math.Exp(arguments[0] + arguments[1] * arguments[1] / 2.0);
                return true;

            case "uniform":
                if (arguments[0] >= arguments[1])
                    return false;
                result = (arguments[0] + arguments[1]) / 2.0;
                return true;

            case "beta":
                if (arguments[0] <= 0 || arguments[1] <= 0)
                    return false;
                result = arguments[0] / (arguments[0] + arguments[1]);
                return true;

            case "weibull":
                if (arguments[0] <= 0 || arguments[1] <= 0)
                    return false;
                result = arguments[1] * SpecialFunctions.Gamma(1.0 + 1.0 / arguments[0]);
                return true;

            case "exponential":
                if (arguments[0] <= 0)
                    return false;
                result = 1.0 / arguments[0];
                return true;

            case "poisson":
                result = arguments[0];
                return arguments[0] > 0;

            case "gamma":
                if (arguments[0] <= 0 || arguments[1] <= 0)
                    return false;
                result = arguments[0] / arguments[1];
                return true;

            case "logistic":
                result = arguments[0];
                return arguments[1] > 0;

            case "gev":
                if (arguments[1] <= 0)
                    return false;

                if (Math.Abs(arguments[2]) < 1e-10)
                {
                    result = arguments[0] + arguments[1] * 0.5772156649;
                    return true;
                }

                if (arguments[2] < 1.0)
                {
                    result = arguments[0] + arguments[1] * (SpecialFunctions.Gamma(1.0 - arguments[2]) - 1.0) / arguments[2];
                    return true;
                }

                result = double.PositiveInfinity;
                return true;

            case "binomial":
                if (arguments[0] <= 0 || arguments[1] < 0 || arguments[1] > 1)
                    return false;
                result = arguments[0] * arguments[1];
                return true;

            case "geometric":
                if (arguments[0] <= 0 || arguments[0] > 1)
                    return false;
                result = 1.0 / arguments[0];
                return true;

            default:
                return false;
        }
    }
}
