using MonteCarlo.Addin.Excel;
using MonteCarlo.Shared.Formula;

namespace MonteCarlo.Addin.UDF;

/// <summary>
/// Scans Excel worksheets for cells containing MC.* formulas
/// and parses them into distribution definitions.
/// </summary>
public class MCFunctionScanner
{
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
                            if (McFormulaParser.TryParse(formula, out var parsed) && parsed != null)
                            {
                                string address = cell.Address.ToString().Replace("$", "");
                                results.Add(new DetectedMCFunction
                                {
                                    Cell = new CellReference { SheetName = sheet.Name, CellAddress = address },
                                    DistributionName = parsed.DistributionName,
                                    ParameterNames = parsed.Arguments.Select(argument => argument.Name).ToArray(),
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
            if (!McFormulaParser.TryParse(function.Formula, out var parsed) || parsed == null)
                return null;

            var argParts = parsed.RawArguments;

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
                        else if (val != null && double.TryParse(val.ToString(), out double parsedValue))
                            parameters[function.ParameterNames[i]] = parsedValue;
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
