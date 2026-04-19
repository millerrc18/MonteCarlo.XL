using MonteCarlo.Engine.Correlation;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Engine.Validation;

/// <summary>
/// Severity level for a preflight validation issue.
/// </summary>
public enum PreflightSeverity
{
    Info,
    Warning,
    Error
}

/// <summary>
/// A single actionable preflight validation result.
/// </summary>
public sealed record PreflightIssue(
    PreflightSeverity Severity,
    string Code,
    string Title,
    string Message,
    string SuggestedAction);

/// <summary>
/// Aggregate preflight validation report for a simulation profile.
/// </summary>
public sealed class PreflightReport
{
    public PreflightReport(IReadOnlyList<PreflightIssue> issues)
    {
        Issues = issues;
    }

    public IReadOnlyList<PreflightIssue> Issues { get; }

    public bool HasErrors => ErrorCount > 0;

    public int ErrorCount => Issues.Count(i => i.Severity == PreflightSeverity.Error);

    public int WarningCount => Issues.Count(i => i.Severity == PreflightSeverity.Warning);

    public int InfoCount => Issues.Count(i => i.Severity == PreflightSeverity.Info);
}

/// <summary>
/// Validates a simulation profile before the add-in starts mutating Excel cells.
/// </summary>
public static class ModelPreflightValidator
{
    private const long MemoryWarningBytes = 250L * 1024L * 1024L;

    public static PreflightReport Validate(SimulationProfile? profile)
    {
        var issues = new List<PreflightIssue>();

        if (profile == null)
        {
            AddError(
                issues,
                "PROFILE_MISSING",
                "No simulation profile",
                "MonteCarlo.XL could not build a simulation profile from the current setup.",
                "Open Setup and add at least one input and one output.");
            return new PreflightReport(issues);
        }

        ValidateIterations(profile, issues);
        ValidateInputs(profile, issues);
        ValidateOutputs(profile, issues);
        ValidateCellConflicts(profile, issues);
        ValidateCorrelation(profile, issues);
        ValidateMemory(profile, issues);

        if (issues.Count == 0)
        {
            issues.Add(new PreflightIssue(
                PreflightSeverity.Info,
                "READY",
                "Ready to simulate",
                "No blocking setup issues were found.",
                "Run the simulation when you are ready."));
        }

        return new PreflightReport(issues);
    }

    private static void ValidateIterations(SimulationProfile profile, List<PreflightIssue> issues)
    {
        if (profile.IterationCount <= 0)
        {
            AddError(
                issues,
                "ITERATIONS_INVALID",
                "Invalid iteration count",
                $"Iteration count is {profile.IterationCount:N0}. It must be greater than zero.",
                "Set iterations to a positive value such as 1,000 or 5,000.");
            return;
        }

        if (profile.IterationCount < 100)
        {
            AddWarning(
                issues,
                "ITERATIONS_LOW",
                "Very low iteration count",
                $"Iteration count is {profile.IterationCount:N0}, which may produce noisy results.",
                "Use at least 1,000 iterations for a quick smoke run and more for final analysis.");
        }

        if (profile.IterationCount > 250_000)
        {
            AddWarning(
                issues,
                "ITERATIONS_HIGH",
                "Large simulation run",
                $"Iteration count is {profile.IterationCount:N0}. Excel recalc mode may take a while.",
                "Use a smaller preview run first, then increase iterations for the final run.");
        }
    }

    private static void ValidateInputs(SimulationProfile profile, List<PreflightIssue> issues)
    {
        if (profile.Inputs.Count == 0)
        {
            AddError(
                issues,
                "NO_INPUTS",
                "No inputs configured",
                "The model does not have any uncertainty inputs.",
                "Add at least one input assumption or use MC.* formulas that can be auto-detected.");
            return;
        }

        var seenCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < profile.Inputs.Count; i++)
        {
            var input = profile.Inputs[i];
            var label = string.IsNullOrWhiteSpace(input.Label) ? $"Input {i + 1}" : input.Label;
            var cellKey = CellKey(input.SheetName, input.CellAddress);

            if (string.IsNullOrWhiteSpace(input.SheetName) || string.IsNullOrWhiteSpace(input.CellAddress))
            {
                AddError(
                    issues,
                    "INPUT_CELL_MISSING",
                    "Input cell is missing",
                    $"{label} does not have a worksheet and cell address.",
                    "Edit or remove the input, then select a valid Excel cell.");
            }
            else if (!seenCells.Add(cellKey))
            {
                AddError(
                    issues,
                    "INPUT_DUPLICATE_CELL",
                    "Duplicate input cell",
                    $"More than one input is assigned to {input.SheetName}!{input.CellAddress}.",
                    "Keep one input per worksheet cell.");
            }

            if (string.IsNullOrWhiteSpace(input.Label))
            {
                AddWarning(
                    issues,
                    "INPUT_LABEL_MISSING",
                    "Input label is blank",
                    $"The input at {input.SheetName}!{input.CellAddress} has no label.",
                    "Add a meaningful label so reports and tornado charts are readable.");
            }

            ValidateDistribution(input, label, issues);
        }
    }

    private static void ValidateDistribution(SavedInput input, string label, List<PreflightIssue> issues)
    {
        if (string.IsNullOrWhiteSpace(input.DistributionName))
        {
            AddError(
                issues,
                "DISTRIBUTION_MISSING",
                "Distribution is missing",
                $"{label} does not specify a probability distribution.",
                "Choose a distribution and enter its required parameters.");
            return;
        }

        if (input.Parameters.Any(pair => !double.IsFinite(pair.Value)))
        {
            AddError(
                issues,
                "DISTRIBUTION_PARAMETER_NONFINITE",
                "Distribution parameter is not finite",
                $"{label} has a parameter that is NaN or infinite.",
                "Replace non-finite parameters with valid numeric values.");
            return;
        }

        try
        {
            DistributionFactory.Create(input.DistributionName, input.Parameters);
        }
        catch (Exception ex)
        {
            AddError(
                issues,
                "DISTRIBUTION_INVALID",
                "Distribution parameters are invalid",
                $"{label}: {ex.Message}",
                "Edit the input parameters and confirm the distribution preview renders correctly.");
        }
    }

    private static void ValidateOutputs(SimulationProfile profile, List<PreflightIssue> issues)
    {
        if (profile.Outputs.Count == 0)
        {
            AddError(
                issues,
                "NO_OUTPUTS",
                "No outputs configured",
                "The model does not have any output cells to record.",
                "Add at least one output forecast cell.");
            return;
        }

        var seenCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < profile.Outputs.Count; i++)
        {
            var output = profile.Outputs[i];
            var label = string.IsNullOrWhiteSpace(output.Label) ? $"Output {i + 1}" : output.Label;

            if (string.IsNullOrWhiteSpace(output.SheetName) || string.IsNullOrWhiteSpace(output.CellAddress))
            {
                AddError(
                    issues,
                    "OUTPUT_CELL_MISSING",
                    "Output cell is missing",
                    $"{label} does not have a worksheet and cell address.",
                    "Edit or remove the output, then select a valid Excel cell.");
                continue;
            }

            if (!seenCells.Add(CellKey(output.SheetName, output.CellAddress)))
            {
                AddError(
                    issues,
                    "OUTPUT_DUPLICATE_CELL",
                    "Duplicate output cell",
                    $"More than one output is assigned to {output.SheetName}!{output.CellAddress}.",
                    "Keep one output per worksheet cell.");
            }

            if (string.IsNullOrWhiteSpace(output.Label))
            {
                AddWarning(
                    issues,
                    "OUTPUT_LABEL_MISSING",
                    "Output label is blank",
                    $"The output at {output.SheetName}!{output.CellAddress} has no label.",
                    "Add a meaningful label so exported reports are readable.");
            }
        }
    }

    private static void ValidateCellConflicts(SimulationProfile profile, List<PreflightIssue> issues)
    {
        var inputCells = profile.Inputs
            .Where(i => !string.IsNullOrWhiteSpace(i.SheetName) && !string.IsNullOrWhiteSpace(i.CellAddress))
            .Select(i => CellKey(i.SheetName, i.CellAddress))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var output in profile.Outputs)
        {
            if (inputCells.Contains(CellKey(output.SheetName, output.CellAddress)))
            {
                AddError(
                    issues,
                    "CELL_IS_INPUT_AND_OUTPUT",
                    "Cell is both input and output",
                    $"{output.SheetName}!{output.CellAddress} is configured as both an input and an output.",
                    "Choose a separate output formula cell that depends on the input.");
            }
        }
    }

    private static void ValidateCorrelation(SimulationProfile profile, List<PreflightIssue> issues)
    {
        var matrix = profile.CorrelationMatrix;
        if (matrix == null)
            return;

        if (matrix.GetLength(0) != profile.Inputs.Count || matrix.GetLength(1) != profile.Inputs.Count)
        {
            AddError(
                issues,
                "CORRELATION_SIZE_MISMATCH",
                "Correlation matrix size does not match inputs",
                $"The matrix is {matrix.GetLength(0)}x{matrix.GetLength(1)} but the model has {profile.Inputs.Count} inputs.",
                "Reopen the correlation editor so it can rebuild the matrix for the current inputs.");
            return;
        }

        try
        {
            new CorrelationMatrix(matrix).Validate();
        }
        catch (ArgumentException ex)
        {
            AddError(
                issues,
                "CORRELATION_INVALID",
                "Correlation matrix is invalid",
                ex.Message,
                "Use Auto-fix in the correlation editor or reduce conflicting pairwise correlations.");
        }
    }

    private static void ValidateMemory(SimulationProfile profile, List<PreflightIssue> issues)
    {
        if (profile.IterationCount <= 0)
            return;

        var seriesCount = profile.Inputs.Count + profile.Outputs.Count;
        var estimatedBytes = (long)profile.IterationCount * Math.Max(seriesCount, 1) * sizeof(double);
        if (estimatedBytes < MemoryWarningBytes)
            return;

        AddWarning(
            issues,
            "RESULT_SIZE_LARGE",
            "Large result set",
            $"This run may store about {FormatBytes(estimatedBytes)} of sample data before any raw-data export.",
            "Use a preview run first or reduce raw-data export size for very large models.");
    }

    private static string CellKey(string sheetName, string cellAddress) =>
        $"{sheetName.Trim()}!{cellAddress.Trim()}";

    private static string FormatBytes(long bytes)
    {
        var mb = bytes / 1024.0 / 1024.0;
        return mb < 1024 ? $"{mb:N0} MB" : $"{mb / 1024.0:N1} GB";
    }

    private static void AddError(
        List<PreflightIssue> issues,
        string code,
        string title,
        string message,
        string suggestedAction) =>
        issues.Add(new PreflightIssue(PreflightSeverity.Error, code, title, message, suggestedAction));

    private static void AddWarning(
        List<PreflightIssue> issues,
        string code,
        string title,
        string message,
        string suggestedAction) =>
        issues.Add(new PreflightIssue(PreflightSeverity.Warning, code, title, message, suggestedAction));
}
