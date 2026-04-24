using System.Text.RegularExpressions;

namespace MonteCarlo.Shared.Formula;

public static class McFormulaParser
{
    private static readonly Regex McFormulaPattern = new(
        @"^=MC\.(\w+)\((.*)\)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static bool TryParse(string? formula, out ParsedMcFormula? parsed)
    {
        parsed = null;
        if (string.IsNullOrWhiteSpace(formula))
            return false;

        var match = McFormulaPattern.Match(formula.Trim());
        if (!match.Success)
            return false;

        var functionName = match.Groups[1].Value;
        if (!McFormulaCatalog.TryGetByFunctionName(functionName, out var definition))
            return false;

        var arguments = SplitArguments(match.Groups[2].Value);
        parsed = new ParsedMcFormula(
            formula.Trim(),
            definition.FunctionName,
            definition.DistributionName,
            arguments,
            definition.Arguments);
        return true;
    }

    public static IReadOnlyList<string> SplitArguments(string argumentList)
    {
        var args = new List<string>();
        var depth = 0;
        var start = 0;

        for (var index = 0; index < argumentList.Length; index++)
        {
            var current = argumentList[index];
            if (current == '(')
            {
                depth++;
            }
            else if (current == ')')
            {
                depth--;
            }
            else if (current == ',' && depth == 0)
            {
                args.Add(argumentList[start..index].Trim());
                start = index + 1;
            }
        }

        if (start < argumentList.Length)
            args.Add(argumentList[start..].Trim());

        return args;
    }
}
