using System.Text.Json;
using System.Xml;
using ExcelDna.Integration;
using MonteCarlo.Addin.Excel;
using MonteCarlo.Addin.UDF;
using Microsoft.Office.Interop.Excel;

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
    public FunctionSwapResult RestoreMcFormulas()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = GetActiveWorkbook();
        var snapshot = LoadSnapshot(workbook);
        if (snapshot == null || snapshot.Entries.Count == 0)
            return new FunctionSwapResult(0, "No MonteCarlo.XL formula restore map was found in this workbook.");

        using var excelState = ExcelStateScope.Capture(app, "Restore MC formulas", restoreSelection: true);
        excelState.Apply(
            screenUpdating: false,
            calculation: XlCalculation.xlCalculationManual,
            statusBar: "MonteCarlo.XL: restoring MC formulas...");

        var restored = 0;
        foreach (var entry in snapshot.Entries)
        {
            try
            {
                var range = GetRange(workbook, entry);
                range.Formula = entry.Formula;
                restored++;
            }
            catch (Exception ex)
            {
                StartupDiagnostics.LogException(
                    $"Restore MC formula failed for {entry.SheetName}!{entry.CellAddress}.",
                    ex);
            }
        }

        if (restored == snapshot.Entries.Count)
            RemoveExistingSnapshot(workbook);

        return new FunctionSwapResult(restored, $"Restored {restored:N0} MC.* formulas.");
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
