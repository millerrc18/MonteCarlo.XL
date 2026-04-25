using System.Text.Json;
using System.Xml;
using System.IO;
using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.UDF;
using Microsoft.Office.Interop.Excel;
using WinForms = System.Windows.Forms;

namespace MonteCarlo.Addin.Services;

/// <summary>
/// Replaces MC.* formulas with their current values for sharing, and restores them later.
/// </summary>
public sealed class FunctionSwapService
{
    private const string CustomXmlNamespace = "urn:montecarlo-xl:function-swap:v1";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <summary>
    /// Catalog all MC.* formulas in the active workbook without changing cells.
    /// </summary>
    public IReadOnlyList<FormulaSwapEntry> CatalogActiveWorkbook()
    {
        var workbook = GetActiveWorkbook();
        return CaptureWorkbook(workbook).Select(c => c.Entry).ToList();
    }

    /// <summary>
    /// Preview the active workbook restore map and flag cells whose values changed after replacement.
    /// </summary>
    internal FormulaRestorePreview? BuildRestorePreview()
    {
        var workbook = GetActiveWorkbook();
        var snapshot = LoadSnapshot(workbook);
        if (snapshot == null || snapshot.Entries.Count == 0)
            return null;

        var conflicts = DetectRestoreConflicts(workbook, snapshot.Entries);
        return new FormulaRestorePreview(snapshot.Entries, conflicts);
    }

    /// <summary>
    /// Replace MC.* formulas with current cell values and save a restore map in workbook custom XML.
    /// </summary>
    public FunctionSwapResult ReplaceMcFormulasWithCurrentValues()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = GetActiveWorkbook();
        var captured = CaptureWorkbook(workbook);
        if (captured.Count == 0)
            return new FunctionSwapResult(0, "No MC.* formulas were found in the active workbook.");

        var snapshot = new FormulaSwapSnapshot
        {
            WorkbookName = workbook.Name,
            CreatedAtUtc = DateTimeOffset.UtcNow,
            Entries = captured.Select(c => c.Entry).ToList()
        };

        SaveSnapshot(workbook, snapshot);

        using var excelState = ExcelStateScope.Capture(app, "Replace MC formulas", restoreSelection: true);
        excelState.Apply(
            screenUpdating: false,
            calculation: XlCalculation.xlCalculationManual,
            statusBar: "MonteCarlo.XL: replacing MC formulas with current values...");

        foreach (var formula in captured)
        {
            var range = GetRange(workbook, formula.Entry);
            range.Value2 = formula.ReplacementValue;
        }

        return new FunctionSwapResult(captured.Count, $"Replaced {captured.Count:N0} MC.* formulas with current values.");
    }

    /// <summary>
    /// Restore MC.* formulas from the active workbook restore map.
    /// </summary>
    public FunctionSwapResult RestoreMcFormulas() =>
        RestoreMcFormulas(RestoreConflictResolution.OverwriteAll);

    internal FunctionSwapResult RestoreMcFormulas(RestoreConflictResolution resolution)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = GetActiveWorkbook();
        var snapshot = LoadSnapshot(workbook);
        if (snapshot == null || snapshot.Entries.Count == 0)
            return new FunctionSwapResult(0, "No MonteCarlo.XL formula restore map was found in this workbook.");

        var conflictsByKey = DetectRestoreConflicts(workbook, snapshot.Entries)
            .ToDictionary(
                conflict => FormulaCatalogPreviewDialog.BuildKey(conflict.Entry),
                conflict => conflict,
                StringComparer.OrdinalIgnoreCase);

        using var excelState = ExcelStateScope.Capture(app, "Restore MC formulas", restoreSelection: true);
        excelState.Apply(
            screenUpdating: false,
            calculation: XlCalculation.xlCalculationManual,
            statusBar: "MonteCarlo.XL: restoring MC formulas...");

        var restored = 0;
        var skipped = 0;
        var alreadyRestored = 0;
        var remainingEntries = new List<FormulaSwapEntry>();
        foreach (var entry in snapshot.Entries)
        {
            try
            {
                var range = GetRange(workbook, entry);
                if (CellAlreadyHasFormula(range, entry.Formula))
                {
                    alreadyRestored++;
                    continue;
                }

                var key = FormulaCatalogPreviewDialog.BuildKey(entry);
                if (resolution == RestoreConflictResolution.SkipChangedCells &&
                    conflictsByKey.ContainsKey(key))
                {
                    skipped++;
                    remainingEntries.Add(entry);
                    continue;
                }

                range.Formula = entry.Formula;
                restored++;
            }
            catch (Exception ex)
            {
                remainingEntries.Add(entry);
                StartupDiagnostics.LogException(
                    $"Restore MC formula failed for {entry.SheetName}!{entry.CellAddress}.",
                    ex);
            }
        }

        if (remainingEntries.Count == 0)
        {
            RemoveExistingSnapshot(workbook);
        }
        else
        {
            SaveSnapshot(workbook, new FormulaSwapSnapshot
            {
                WorkbookName = snapshot.WorkbookName,
                CreatedAtUtc = snapshot.CreatedAtUtc,
                Entries = remainingEntries
            });
        }

        var message = $"Restored {restored:N0} MC.* formulas.";
        if (skipped > 0)
            message += $"\r\nSkipped {skipped:N0} changed cell(s); the restore map was kept for those entries.";
        if (alreadyRestored > 0)
            message += $"\r\nRemoved {alreadyRestored:N0} entry/entries that were already back to formulas.";
        if (remainingEntries.Count > 0 && skipped == 0)
            message += $"\r\n{remainingEntries.Count:N0} entry/entries remain in the restore map because Excel could not restore them.";

        return new FunctionSwapResult(restored, message);
    }

    /// <summary>
    /// Save a separate workbook copy with MC.* formulas replaced by current values.
    /// The active workbook is left unchanged and formula-backed.
    /// </summary>
    public FunctionSwapResult SaveShareableCopy()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var sourceWorkbook = GetActiveWorkbook();

        using var dialog = new WinForms.SaveFileDialog
        {
            Title = "Save Shareable MonteCarlo.XL Copy",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx|Macro-Enabled Workbook (*.xlsm)|*.xlsm",
            FileName = GetDefaultShareFileName(sourceWorkbook),
            InitialDirectory = GetDefaultShareDirectory(sourceWorkbook),
            AddExtension = true,
            OverwritePrompt = true
        };

        if (dialog.ShowDialog() != WinForms.DialogResult.OK)
            return new FunctionSwapResult(0, "Shareable copy cancelled.");

        var destinationPath = dialog.FileName;
        var sourceExtension = Path.GetExtension(sourceWorkbook.Name);
        if (string.IsNullOrWhiteSpace(sourceExtension))
            sourceExtension = ".xlsx";

        var tempPath = Path.Combine(
            Path.GetTempPath(),
            $"mcxl-share-{Guid.NewGuid():N}{sourceExtension}");

        Workbook? copyWorkbook = null;
        using var excelState = ExcelStateScope.Capture(app, "Save shareable copy", restoreSelection: true);

        try
        {
            excelState.Apply(
                screenUpdating: false,
                displayAlerts: false,
                calculation: XlCalculation.xlCalculationManual,
                statusBar: "MonteCarlo.XL: saving shareable copy...");

            sourceWorkbook.SaveCopyAs(tempPath);
            copyWorkbook = app.Workbooks.Open(tempPath, UpdateLinks: 0, ReadOnly: false);

            var captured = CaptureWorkbook(copyWorkbook);
            foreach (var formula in captured)
            {
                var range = GetRange(copyWorkbook, formula.Entry);
                range.Value2 = formula.ReplacementValue;
            }

            RemoveExistingSnapshot(copyWorkbook);
            copyWorkbook.SaveAs(
                Filename: destinationPath,
                FileFormat: GetFileFormat(destinationPath),
                ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);

            return new FunctionSwapResult(
                captured.Count,
                $"Saved shareable copy with {captured.Count:N0} MC.* formulas replaced:\r\n{destinationPath}");
        }
        finally
        {
            try { copyWorkbook?.Close(SaveChanges: false); } catch { }
            try { sourceWorkbook.Activate(); } catch { }
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
        }
    }

    private static Workbook GetActiveWorkbook()
    {
        var app = (Application)ExcelDnaUtil.Application;
        return app.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook is available.");
    }

    private static List<CapturedFormula> CaptureWorkbook(Workbook workbook)
    {
        var scanner = new MCFunctionScanner();
        var captured = new List<CapturedFormula>();

        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            foreach (var detected in scanner.ScanWorksheet(worksheet))
            {
                try
                {
                    var range = (Range)worksheet.Range[detected.Cell.CellAddress];
                    var value = range.Value2;
                    captured.Add(new CapturedFormula(
                        new FormulaSwapEntry
                        {
                            SheetName = worksheet.Name,
                            CellAddress = detected.Cell.CellAddress,
                            Formula = detected.Formula,
                            ValueText = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty
                        },
                        value ?? string.Empty));
                }
                catch (Exception ex)
                {
                    StartupDiagnostics.LogException(
                        $"Catalog MC formula failed for {worksheet.Name}!{detected.Cell.CellAddress}.",
                        ex);
                }
            }
        }

        return captured;
    }

    private static Range GetRange(Workbook workbook, FormulaSwapEntry entry)
    {
        var worksheet = (Worksheet)workbook.Worksheets[entry.SheetName];
        return (Range)worksheet.Range[entry.CellAddress];
    }

    private static List<FormulaRestoreConflict> DetectRestoreConflicts(
        Workbook workbook,
        IEnumerable<FormulaSwapEntry> entries)
    {
        var conflicts = new List<FormulaRestoreConflict>();
        foreach (var entry in entries)
        {
            try
            {
                var range = GetRange(workbook, entry);
                if (CellAlreadyHasFormula(range, entry.Formula))
                    continue;

                var currentValue = NormalizeCellValue(range.Value2);
                if (string.Equals(currentValue, entry.ValueText, StringComparison.Ordinal))
                    continue;

                conflicts.Add(new FormulaRestoreConflict(
                    entry,
                    DescribeCurrentCellState(range, currentValue)));
            }
            catch (Exception ex)
            {
                StartupDiagnostics.LogException(
                    $"Preview restore conflict failed for {entry.SheetName}!{entry.CellAddress}.",
                    ex);
                conflicts.Add(new FormulaRestoreConflict(entry, "Could not read current cell state"));
            }
        }

        return conflicts;
    }

    private static bool CellAlreadyHasFormula(Range range, string formula)
    {
        try
        {
            var hasFormula = range.HasFormula is bool boolValue
                ? boolValue
                : Convert.ToBoolean(range.HasFormula);
            if (!hasFormula)
                return false;

            var currentFormula = Convert.ToString(range.Formula, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            return string.Equals(currentFormula, formula, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string DescribeCurrentCellState(Range range, string normalizedValue)
    {
        try
        {
            var hasFormula = range.HasFormula is bool boolValue
                ? boolValue
                : Convert.ToBoolean(range.HasFormula);
            if (hasFormula)
            {
                var currentFormula = Convert.ToString(range.Formula, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                return $"Formula: {currentFormula}";
            }
        }
        catch { }

        return string.IsNullOrEmpty(normalizedValue) ? "(blank)" : $"Value: {normalizedValue}";
    }

    private static string NormalizeCellValue(object? value) =>
        value switch
        {
            null => string.Empty,
            double number => number.ToString("G17", System.Globalization.CultureInfo.InvariantCulture),
            float number => number.ToString("G9", System.Globalization.CultureInfo.InvariantCulture),
            bool boolValue => boolValue ? "TRUE" : "FALSE",
            DateTime dateTime => dateTime.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)?.Trim() ?? string.Empty
        };

    private static void SaveSnapshot(Workbook workbook, FormulaSwapSnapshot snapshot)
    {
        var json = JsonSerializer.Serialize(snapshot, JsonOptions);
        var xml = WrapInXml(json, snapshot.WorkbookName);

        RemoveExistingSnapshot(workbook);
        dynamic parts = workbook.CustomXMLParts;
        parts.Add(xml);
    }

    private static FormulaSwapSnapshot? LoadSnapshot(Workbook workbook)
    {
        try
        {
            dynamic parts = workbook.CustomXMLParts;
            int count = parts.Count;

            for (var idx = 1; idx <= count; idx++)
            {
                try
                {
                    dynamic part = parts[idx];
                    string xml = part.XML;
                    if (!xml.Contains(CustomXmlNamespace, StringComparison.Ordinal))
                        continue;

                    return DeserializeSnapshotXml(xml);
                }
                catch { }
            }
        }
        catch { }

        return null;
    }

    private static void RemoveExistingSnapshot(Workbook workbook)
    {
        try
        {
            dynamic parts = workbook.CustomXMLParts;
            int count = parts.Count;
            var indicesToRemove = new List<int>();

            for (var idx = 1; idx <= count; idx++)
            {
                try
                {
                    dynamic part = parts[idx];
                    if (!(bool)part.BuiltIn && ((string)part.XML).Contains(CustomXmlNamespace, StringComparison.Ordinal))
                        indicesToRemove.Add(idx);
                }
                catch { }
            }

            for (var i = indicesToRemove.Count - 1; i >= 0; i--)
            {
                try { parts[indicesToRemove[i]].Delete(); } catch { }
            }
        }
        catch { }
    }

    private static string WrapInXml(string json, string workbookName)
    {
        var safeJson = json.Replace("]]>", "]]]]><![CDATA[>", StringComparison.Ordinal);

        return $"""
            <?xml version="1.0" encoding="utf-8"?>
            <MonteCarloFormulaSwap xmlns="{CustomXmlNamespace}">
              <Snapshot workbook="{System.Security.SecurityElement.Escape(workbookName)}"><![CDATA[{safeJson}]]></Snapshot>
            </MonteCarloFormulaSwap>
            """;
    }

    private static FormulaSwapSnapshot? DeserializeSnapshotXml(string xml)
    {
        var doc = new XmlDocument();
        doc.LoadXml(xml);

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("mc", CustomXmlNamespace);
        var node = doc.SelectSingleNode("//mc:Snapshot", ns);
        var json = node?.InnerText;
        return string.IsNullOrWhiteSpace(json)
            ? null
            : JsonSerializer.Deserialize<FormulaSwapSnapshot>(json, JsonOptions);
    }

    private sealed record CapturedFormula(FormulaSwapEntry Entry, object ReplacementValue);

    private static string GetDefaultShareFileName(Workbook workbook)
    {
        var baseName = Path.GetFileNameWithoutExtension(workbook.Name);
        return $"{baseName} - shareable.xlsx";
    }

    private static string GetDefaultShareDirectory(Workbook workbook)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(workbook.Path))
                return workbook.Path;
        }
        catch { }

        return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }

    private static XlFileFormat GetFileFormat(string path) =>
        string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase)
            ? XlFileFormat.xlOpenXMLWorkbookMacroEnabled
            : XlFileFormat.xlOpenXMLWorkbook;
}

/// <summary>
/// One MC.* formula captured before value replacement.
/// </summary>
public sealed class FormulaSwapEntry
{
    public required string SheetName { get; init; }
    public required string CellAddress { get; init; }
    public required string Formula { get; init; }
    public string ValueText { get; init; } = string.Empty;
}

/// <summary>
/// Restore map persisted in workbook custom XML.
/// </summary>
public sealed class FormulaSwapSnapshot
{
    public string WorkbookName { get; init; } = string.Empty;
    public DateTimeOffset CreatedAtUtc { get; init; }
    public List<FormulaSwapEntry> Entries { get; init; } = new();
}

/// <summary>
/// Result message for a formula swap operation.
/// </summary>
public sealed record FunctionSwapResult(int Count, string Message);
