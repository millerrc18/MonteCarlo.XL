using System.Text.RegularExpressions;
using MonteCarlo.Addin.Excel;

namespace MonteCarlo.Addin.UDF;

/// <summary>
/// Scans Excel worksheets for cells containing MC.* formulas
/// and parses them into distribution definitions.
/// </summary>
public class MCFunctionScanner
{
    // Matches =MC.FunctionName(args...) in formulas
    private static readonly Regex MCFormulaPattern = new(
        @"^=MC\.(\w+)\((.+)\)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// Maps MC function names to their parameter names (positional).
    /// </summary>
    private static readonly Dictionary<string, string[]> ParameterMapping = new(StringComparer.OrdinalIgnoreCase)
    {
        { "Normal", new[] { "mean", "stdDev" } },
        { "Triangular", new[] { "min", "mode", "max" } },
        { "PERT", new[] { "min", "mode", "max" } },
        { "Lognormal", new[] { "mu", "sigma" } },
        { "Uniform", new[] { "min", "max" } },
        { "Beta", new[] { "alpha", "beta" } },
        { "Weibull", new[] { "shape", "scale" } },
        { "Exponential", new[] { "rate" } },
        { "Poisson", new[] { "lambda" } }
    };

    /// <summary>
    /// Scans a worksheet's used range for MC.* formulas.
    /// </summary>
    /// <param name="sheet">The Excel worksheet to scan.</param>
    /// <returns>List of detected MC function cells with parsed distribution info.</returns>
    public List<DetectedMCFunction> ScanWorksheet(dynamic sheet)
    {
        var results = new List<DetectedMCFunction>();

        try
        {
            dynamic usedRange = sheet.UsedRange;
            if (usedRange == null) return results;

            foreach (dynamic cell in usedRange)
            {
                try
                {
                    if (cell.HasFormula)
                    {
                        string formula = cell.Formula?.ToString() ?? string.Empty;
                        if (formula.StartsWith("=MC.", StringComparison.OrdinalIgnoreCase))
                        {
                            var parsed = ParseFormula(formula);
                            if (parsed != null)
                            {
                                string address = cell.Address.ToString().Replace("$", "");
                                results.Add(new DetectedMCFunction
                                {
                                    Cell = new CellReference { SheetName = sheet.Name, CellAddress = address },
                                    DistributionName = parsed.DistributionName,
                                    ParameterNames = parsed.ParameterNames,
                                    Formula = formula
                                });
                            }
                        }
                    }
                }
                catch
                {
                    // Skip cells that can't be read (merged cells, errors, etc.)
                }
            }
        }
        catch
        {
            // Sheet may have no used range
        }

        return results;
    }

    /// <summary>
    /// Resolves parameter values for a detected MC function.
    /// Reads the cell's current computed value (which Excel has already evaluated)
    /// to handle both literal and cell-reference arguments.
    /// </summary>
    /// <param name="function">The detected MC function.</param>
    /// <param name="sheet">The worksheet containing the cell.</param>
    /// <returns>Dictionary of parameter name → value, or null if resolution fails.</returns>
    public Dictionary<string, double>? ResolveParameters(DetectedMCFunction function, dynamic sheet)
    {
        try
        {
            // Parse the formula arguments
            var match = MCFormulaPattern.Match(function.Formula);
            if (!match.Success) return null;

            string argsStr = match.Groups[2].Value;
            var argParts = SplitArguments(argsStr);

            if (argParts.Count != function.ParameterNames.Length)
                return null;

            var parameters = new Dictionary<string, double>();

            for (int i = 0; i < argParts.Count; i++)
            {
                string arg = argParts[i].Trim();

                // Try to parse as a literal number
                if (double.TryParse(arg, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double literal))
                {
                    parameters[function.ParameterNames[i]] = literal;
                }
                else
                {
                    // Assume it's a cell reference — read the value
                    try
                    {
                        dynamic refCell = sheet.Range[arg];
                        object? val = refCell.Value2;
                        if (val is double d)
                            parameters[function.ParameterNames[i]] = d;
                        else if (val != null && double.TryParse(val.ToString(), out double parsed))
                            parameters[function.ParameterNames[i]] = parsed;
                        else
                            return null; // Can't resolve
                    }
                    catch
                    {
                        return null;
                    }
                }
            }

            return parameters;
        }
        catch
        {
            return null;
        }
    }

    private static ParsedFormula? ParseFormula(string formula)
    {
        var match = MCFormulaPattern.Match(formula);
        if (!match.Success) return null;

        string funcName = match.Groups[1].Value;

        if (!ParameterMapping.TryGetValue(funcName, out var paramNames))
            return null;

        return new ParsedFormula
        {
            DistributionName = funcName,
            ParameterNames = paramNames
        };
    }

    /// <summary>
    /// Splits function arguments, respecting nested parentheses.
    /// </summary>
    private static List<string> SplitArguments(string argsStr)
    {
        var args = new List<string>();
        int depth = 0;
        int start = 0;

        for (int i = 0; i < argsStr.Length; i++)
        {
            char c = argsStr[i];
            if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == ',' && depth == 0)
            {
                args.Add(argsStr[start..i]);
                start = i + 1;
            }
        }

        if (start < argsStr.Length)
            args.Add(argsStr[start..]);

        return args;
    }

    private class ParsedFormula
    {
        public required string DistributionName { get; init; }
        public required string[] ParameterNames { get; init; }
    }
}

/// <summary>
/// Represents an MC.* function detected in a worksheet cell.
/// </summary>
public class DetectedMCFunction
{
    /// <summary>The cell containing the MC function.</summary>
    public required CellReference Cell { get; init; }

    /// <summary>The distribution name (e.g., "Normal", "PERT").</summary>
    public required string DistributionName { get; init; }

    /// <summary>Parameter names for this distribution type.</summary>
    public required string[] ParameterNames { get; init; }

    /// <summary>The original formula text.</summary>
    public required string Formula { get; init; }
}
