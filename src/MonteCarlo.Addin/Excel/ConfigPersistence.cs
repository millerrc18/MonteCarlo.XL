using System.Text.Json;
using System.Xml;
using ExcelDna.Integration;
using MonteCarlo.Engine.Simulation;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Saves and loads simulation configuration inside the Excel workbook
/// using CustomXMLParts. Falls back to a hidden sheet if needed.
/// </summary>
public class ConfigPersistence
{
    private const string CustomXmlNamespace = "urn:montecarlo-xl:config:v1";
    private const string HiddenSheetName = "__MC_Config";
    private const string ConfigCellAddress = "A1";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <summary>
    /// Save the simulation profile into the active workbook's CustomXMLPart.
    /// </summary>
    public void Save(SimulationProfile profile)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        string json = JsonSerializer.Serialize(profile, JsonOptions);
        string xml = WrapInXml(json, profile.Name);

        // Remove existing config part if present
        RemoveExistingPart(workbook);

        // Add new part
        workbook.CustomXMLParts.Add(xml);
    }

    /// <summary>
    /// Load the simulation profile from the active workbook.
    /// Returns null if no config exists.
    /// </summary>
    public SimulationProfile? Load()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return null;

        // Try CustomXMLPart first
        var profile = LoadFromCustomXml(workbook);
        if (profile != null) return profile;

        // Fallback: try hidden sheet
        return LoadFromHiddenSheet(workbook);
    }

    /// <summary>
    /// Delete any saved config from the workbook.
    /// </summary>
    public void Clear()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        RemoveExistingPart(workbook);
        RemoveHiddenSheet(workbook);
    }

    /// <summary>
    /// List all saved profile names.
    /// </summary>
    public List<string> GetProfileNames()
    {
        var profile = Load();
        if (profile == null) return new List<string>();
        return new List<string> { profile.Name };
    }

    private static string WrapInXml(string json, string profileName)
    {
        // Escape CDATA end markers in JSON (edge case)
        string safeJson = json.Replace("]]>", "]]]]><![CDATA[>");

        return $"""
            <?xml version="1.0" encoding="utf-8"?>
            <MonteCarloConfig xmlns="{CustomXmlNamespace}">
              <Profile name="{System.Security.SecurityElement.Escape(profileName)}"><![CDATA[{safeJson}]]></Profile>
            </MonteCarloConfig>
            """;
    }

    private static SimulationProfile? LoadFromCustomXml(Workbook workbook)
    {
        foreach (CustomXMLPart part in workbook.CustomXMLParts)
        {
            try
            {
                if (part.BuiltIn) continue;
                string xml = part.XML;
                if (!xml.Contains(CustomXmlNamespace)) continue;

                // Parse XML and extract JSON from CDATA
                var doc = new XmlDocument();
                doc.LoadXml(xml);

                var nsManager = new XmlNamespaceManager(doc.NameTable);
                nsManager.AddNamespace("mc", CustomXmlNamespace);

                var profileNode = doc.SelectSingleNode("//mc:Profile", nsManager);
                if (profileNode == null) continue;

                string json = profileNode.InnerText;
                return JsonSerializer.Deserialize<SimulationProfile>(json, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });
            }
            catch
            {
                // Skip malformed parts
            }
        }

        return null;
    }

    private static SimulationProfile? LoadFromHiddenSheet(Workbook workbook)
    {
        try
        {
            Worksheet? sheet = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == HiddenSheetName)
                {
                    sheet = ws;
                    break;
                }
            }

            if (sheet == null) return null;

            var cell = sheet.Range[ConfigCellAddress];
            string? json = cell.Value2?.ToString();
            if (string.IsNullOrEmpty(json)) return null;

            return JsonSerializer.Deserialize<SimulationProfile>(json, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });
        }
        catch
        {
            return null;
        }
    }

    private static void RemoveExistingPart(Workbook workbook)
    {
        var toRemove = new List<CustomXMLPart>();
        foreach (CustomXMLPart part in workbook.CustomXMLParts)
        {
            try
            {
                if (!part.BuiltIn && part.XML.Contains(CustomXmlNamespace))
                    toRemove.Add(part);
            }
            catch { }
        }

        foreach (var part in toRemove)
        {
            try { part.Delete(); } catch { }
        }
    }

    private static void RemoveHiddenSheet(Workbook workbook)
    {
        try
        {
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == HiddenSheetName)
                {
                    var app = workbook.Application;
                    app.DisplayAlerts = false;
                    ws.Delete();
                    app.DisplayAlerts = true;
                    break;
                }
            }
        }
        catch { }
    }
}
