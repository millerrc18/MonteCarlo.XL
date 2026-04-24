namespace MonteCarlo.Shared.Formula;

public sealed record McFormulaArgument(
    string Name,
    string DisplayName,
    string Description);

public sealed record McFormulaDefinition(
    string FunctionName,
    string DistributionName,
    string Description,
    IReadOnlyList<McFormulaArgument> Arguments,
    bool ReturnsErrorForInvalidArguments = true)
{
    public string ExcelName => $"MC.{FunctionName}";
}

public sealed record ParsedMcFormula(
    string Formula,
    string FunctionName,
    string DistributionName,
    IReadOnlyList<string> RawArguments,
    IReadOnlyList<McFormulaArgument> Arguments);
