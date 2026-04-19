using System.Text.Json;
using System.Xml;
using ExcelDna.Integration;
using MonteCarlo.Engine.Simulation;
using Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Saves and loads simulation configuration inside the Excel workbook
/// using CustomXMLParts. Falls back to a hidden sheet if needed.
/// Uses dynamic COM late-binding for CustomXMLPart to avoid
/// a hard dependency on Microsoft.Office.Core interop assembly.
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

    public void Save(SimulationProfile profile)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        string json = JsonSerializer.Serialize(profile, JsonOptions);
        string xml = WrapInXml(json, profile.Name);

        RemoveExistingPart(workbook);

        dynamic parts = workbook.CustomXMLParts;
        parts.Add(xml);
    }

    public SimulationProfile? Load()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return null;

        var profile = LoadFromCustomXml(workbook);
        if (profile != null) return profile;

        return LoadFromHiddenSheet(workbook);
    }

    public void Clear()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        RemoveExistingPart(workbook);
        RemoveHiddenSheet(workbook);
    }

    public List<string> GetProfileNames()
    {
        var profile = Load();
        if (profile == null) return new List<string>();
        return new List<string> { profile.Name };
    }

    private static string WrapInXml(string json, string profileName)
    {
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
        dynamic parts = workbook.CustomXMLParts;
        int count = parts.Count;

        for (int idx = 1; idx <= count; idx++)
        {
            try
            {
                dynamic part = parts[idx];
                if ((bool)part.BuiltIn) continue;
                string xml = (string)part.XML;
                if (!xml.Contains(CustomXmlNamespace)) continue;

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
            catch { }
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
        try
        {
            dynamic parts = workbook.CustomXMLParts;
            int count = parts.Count;
            var indicesToRemove = new List<int>();

            for (int idx = 1; idx <= count; idx++)
            {
                try
                {
                    dynamic part = parts[idx];
                    if (!(bool)part.BuiltIn && ((string)part.XML).Contains(CustomXmlNamespace))
                        indicesToRemove.Add(idx);
                }
                catch { }
            }

            for (int i = indicesToRemove.Count - 1; i >= 0; i--)
            {
                try { parts[indicesToRemove[i]].Delete(); } catch { }
            }
        }
        catch { }
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
                    using var excelState = ExcelStateScope.Capture(app, "Remove hidden config sheet");
                    excelState.Apply(displayAlerts: false);
                    ws.Delete();
                    break;
                }
            }
        }
        catch { }
    }
}
