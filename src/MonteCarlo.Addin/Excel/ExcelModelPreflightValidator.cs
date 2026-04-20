using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.Engine.Validation;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Adds Excel workbook and cell checks to the engine-level model preflight report.
/// </summary>
internal static class ExcelModelPreflightValidator
{
    public static PreflightReport Validate(SimulationProfile? profile, Application? app)
    {
        var issues = ModelPreflightValidator.Validate(profile)
            .Issues
            .Where(issue => issue.Code != "READY")
            .ToList();

        if (app == null)
        {
            AddWarning(
                issues,
                "EXCEL_CONTEXT_UNAVAILABLE",
                "Excel context unavailable",
                "MonteCarlo.XL could not inspect the active Excel workbook.",
                "Open a workbook and run Model Check from the MonteCarlo.XL ribbon or task pane.");
            return CreateReport(issues);
        }

        ValidateApplicationState(app, issues);

        var workbook = TryGetActiveWorkbook(app, issues);
        if (workbook == null)
            return CreateReport(issues);

        AddWorkbookContext(app, workbook, issues);
        ValidateWorkbookState(workbook, issues);

        if (profile != null)
        {
            ValidateInputCells(profile, workbook, issues);
            ValidateOutputCells(profile, workbook, issues);
        }

        return CreateReport(issues);
    }

    private static void ValidateApplicationState(Application app, List<PreflightIssue> issues)
    {
        TryInspect(issues, "Excel calculation mode", () =>
        {
            if (app.Calculation != XlCalculation.xlCalculationAutomatic)
            {
                AddWarning(
                    issues,
                    "EXCEL_CALCULATION_NOT_AUTOMATIC",
                    "Excel calculation is not automatic",
                    "The active Excel application is not in automatic calculation mode. Output cells may be stale before a run.",
                    "Use the Recover Excel ribbon command or switch Excel calculation back to Automatic before final runs.");
            }
        });

        TryInspect(issues, "Excel events", () =>
        {
            if (!app.EnableEvents)
            {
                AddWarning(
                    issues,
                    "EXCEL_EVENTS_DISABLED",
                    "Excel events are disabled",
                    "Excel event handling is currently disabled, which can interfere with cell selection and workbook automation.",
                    "Use the Recover Excel ribbon command before continuing.");
            }
        });

        TryInspect(issues, "Excel screen updating", () =>
        {
            if (!app.ScreenUpdating)
            {
                AddWarning(
                    issues,
                    "EXCEL_SCREEN_UPDATING_DISABLED",
                    "Excel screen updating is disabled",
                    "Excel screen updating is currently disabled, which may make the workbook appear frozen.",
                    "Use the Recover Excel ribbon command before continuing.");
            }
        });
    }

    private static Workbook? TryGetActiveWorkbook(Application app, List<PreflightIssue> issues)
    {
        return TryInspect(issues, "active workbook", () =>
        {
            var workbook = app.ActiveWorkbook;
            if (workbook == null)
            {
                AddError(
                    issues,
                    "NO_ACTIVE_WORKBOOK",
                    "No active workbook",
                    "Excel does not have an active workbook for MonteCarlo.XL to inspect.",
                    "Open the workbook you want to simulate, then run Model Check again.");
            }

            return workbook;
        });
    }

    private static void AddWorkbookContext(Application app, Workbook workbook, List<PreflightIssue> issues)
    {
        TryInspect(issues, "workbook context", () =>
        {
            var activeSheet = app.ActiveSheet is Worksheet worksheet ? worksheet.Name : "unknown";
            AddInfo(
                issues,
                "WORKBOOK_CONTEXT",
                "Workbook context",
                $"Checking workbook '{workbook.Name}' with active sheet '{activeSheet}'.",
                "Confirm this is the workbook you intend to simulate.");
        });
    }

    private static void ValidateWorkbookState(Workbook workbook, List<PreflightIssue> issues)
    {
        TryInspect(issues, "workbook protection", () =>
        {
            if (workbook.ProtectStructure)
            {
                AddWarning(
                    issues,
                    "WORKBOOK_STRUCTURE_PROTECTED",
                    "Workbook structure is protected",
                    "The workbook structure is protected. Summary exports that create new sheets may fail.",
                    "Unprotect workbook structure or disable new-sheet exports before exporting results.");
            }
        });
    }

    private static void ValidateInputCells(SimulationProfile profile, Workbook workbook, List<PreflightIssue> issues)
    {
        foreach (var input in profile.Inputs)
        {
            if (string.IsNullOrWhiteSpace(input.SheetName) || string.IsNullOrWhiteSpace(input.CellAddress))
                continue;

            var label = string.IsNullOrWhiteSpace(input.Label) ? $"{input.SheetName}!{input.CellAddress}" : input.Label;
            var worksheet = TryGetWorksheet(workbook, input.SheetName, "INPUT_SHEET_MISSING", "Input sheet is missing", issues);
            if (worksheet == null)
                continue;

            var range = TryGetRange(worksheet, input.CellAddress, "INPUT_CELL_INVALID", "Input cell is invalid", issues);
            if (range == null)
                continue;

            var cellLabel = $"{worksheet.Name}!{range.Address[false, false]}";
            if (CellDisplaysError(range))
            {
                AddError(
                    issues,
                    "INPUT_CELL_ERROR",
                    "Input cell has an error value",
                    $"{label} at {cellLabel} currently evaluates to {GetCellText(range)}.",
                    "Fix the cell formula or choose a valid input cell before running.");
            }

            if (IsCellLockedOnProtectedSheet(worksheet, range))
            {
                AddError(
                    issues,
                    "INPUT_CELL_PROTECTED",
                    "Input cell is protected",
                    $"{label} at {cellLabel} is locked on a protected worksheet. Simulation runs need to write sampled values into input cells.",
                    "Unprotect the worksheet or unlock the input cell before running.");
            }

            var formula = GetFormula(range);
            if (!string.IsNullOrWhiteSpace(formula) && !IsMonteCarloFormula(formula))
            {
                AddWarning(
                    issues,
                    "INPUT_CELL_NON_MC_FORMULA",
                    "Input cell contains a non-MC formula",
                    $"{label} at {cellLabel} contains a formula that is not an MC.* distribution formula.",
                    "Confirm the input should be overwritten during simulation, or replace it with an MC.* formula.");
            }
        }
    }

    private static void ValidateOutputCells(SimulationProfile profile, Workbook workbook, List<PreflightIssue> issues)
    {
        foreach (var output in profile.Outputs)
        {
            if (string.IsNullOrWhiteSpace(output.SheetName) || string.IsNullOrWhiteSpace(output.CellAddress))
                continue;

            var label = string.IsNullOrWhiteSpace(output.Label) ? $"{output.SheetName}!{output.CellAddress}" : output.Label;
            var worksheet = TryGetWorksheet(workbook, output.SheetName, "OUTPUT_SHEET_MISSING", "Output sheet is missing", issues);
            if (worksheet == null)
                continue;

            var range = TryGetRange(worksheet, output.CellAddress, "OUTPUT_CELL_INVALID", "Output cell is invalid", issues);
            if (range == null)
                continue;

            var cellLabel = $"{worksheet.Name}!{range.Address[false, false]}";
            if (CellDisplaysError(range))
            {
                AddError(
                    issues,
                    "OUTPUT_CELL_ERROR",
                    "Output cell has an error value",
                    $"{label} at {cellLabel} currently evaluates to {GetCellText(range)}.",
                    "Fix the output formula before running the simulation.");
                continue;
            }

            if (!IsNumeric(range.Value2))
            {
                AddError(
                    issues,
                    "OUTPUT_CELL_NOT_NUMERIC",
                    "Output cell is not numeric",
                    $"{label} at {cellLabel} currently evaluates to '{GetCellText(range)}', which cannot be summarized as a simulation output.",
                    "Choose a numeric formula/result cell as the output forecast.");
            }

            if (string.IsNullOrWhiteSpace(GetFormula(range)))
            {
                AddWarning(
                    issues,
                    "OUTPUT_CELL_STATIC_VALUE",
                    "Output cell is a static value",
                    $"{label} at {cellLabel} does not contain a formula.",
                    "Confirm this is intentional. Forecast outputs usually should be formulas that depend on one or more inputs.");
            }
        }
    }

    private static Worksheet? TryGetWorksheet(
        Workbook workbook,
        string sheetName,
        string code,
        string title,
        List<PreflightIssue> issues)
    {
        return TryInspect(issues, $"worksheet '{sheetName}'", () =>
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(worksheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return worksheet;
            }

            AddError(
                issues,
                code,
                title,
                $"Worksheet '{sheetName}' could not be found in the active workbook.",
                "Edit or remove the affected input/output, or open the workbook that contains that sheet.");
            return null;
        });
    }

    private static Range? TryGetRange(
        Worksheet worksheet,
        string address,
        string code,
        string title,
        List<PreflightIssue> issues)
    {
        return TryInspect(issues, $"{worksheet.Name}!{address}", () =>
        {
            try
            {
                return worksheet.Range[address];
            }
            catch (Exception ex)
            {
                AddError(
                    issues,
                    code,
                    title,
                    $"{worksheet.Name}!{address} could not be resolved: {ex.Message}",
                    "Edit or remove the affected input/output, then select a valid worksheet cell.");
                return null;
            }
        });
    }

    private static bool IsCellLockedOnProtectedSheet(Worksheet worksheet, Range range)
    {
        if (!worksheet.ProtectContents)
            return false;

        return range.Locked is bool locked && locked;
    }

    private static string? GetFormula(Range range)
    {
        return range.HasFormula is bool hasFormula && hasFormula
            ? Convert.ToString(range.Formula, CultureInfo.InvariantCulture)
            : null;
    }

    private static bool IsMonteCarloFormula(string formula) =>
        formula.Contains("MC.", StringComparison.OrdinalIgnoreCase)
        || formula.Contains("MonteCarlo", StringComparison.OrdinalIgnoreCase);

    private static bool CellDisplaysError(Range range)
    {
        if (range.Value2 is ErrorWrapper)
            return true;

        var text = GetCellText(range);
        return text.StartsWith("#", StringComparison.Ordinal);
    }

    private static string GetCellText(Range range) =>
        Convert.ToString(range.Text, CultureInfo.InvariantCulture) ?? string.Empty;

    private static bool IsNumeric(object? value)
    {
        return value switch
        {
            null => false,
            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal => true,
            _ => double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out _)
        };
    }

    private static PreflightReport CreateReport(List<PreflightIssue> issues)
    {
        if (issues.Count == 0)
        {
            AddInfo(
                issues,
                "READY",
                "Ready to simulate",
                "No blocking setup or Excel workbook issues were found.",
                "Run the simulation when you are ready.");
        }

        return new PreflightReport(issues);
    }

    private static T? TryInspect<T>(List<PreflightIssue> issues, string context, Func<T?> inspect)
    {
        try
        {
            return inspect();
        }
        catch (Exception ex)
        {
            AddWarning(
                issues,
                "EXCEL_PREFLIGHT_INSPECTION_FAILED",
                "Could not inspect Excel state",
                $"MonteCarlo.XL could not inspect {context}: {ex.Message}",
                "Review the workbook manually or check the startup log for diagnostics.");
            return default;
        }
    }

    private static void TryInspect(List<PreflightIssue> issues, string context, System.Action inspect) =>
        TryInspect<object>(issues, context, () =>
        {
            inspect();
            return null;
        });

    private static void AddInfo(
        List<PreflightIssue> issues,
        string code,
        string title,
        string message,
        string suggestedAction) =>
        issues.Add(new PreflightIssue(PreflightSeverity.Info, code, title, message, suggestedAction));

    private static void AddWarning(
        List<PreflightIssue> issues,
        string code,
        string title,
        string message,
        string suggestedAction) =>
        issues.Add(new PreflightIssue(PreflightSeverity.Warning, code, title, message, suggestedAction));

    private static void AddError(
        List<PreflightIssue> issues,
        string code,
        string title,
        string message,
        string suggestedAction) =>
        issues.Add(new PreflightIssue(PreflightSeverity.Error, code, title, message, suggestedAction));
}
