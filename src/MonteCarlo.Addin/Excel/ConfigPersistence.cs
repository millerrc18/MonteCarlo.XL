using System.Text.Json;
using System.Text;
using System.Xml;
using ExcelDna.Integration;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.UI.Services;
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

    private sealed record StoredConfigData(
        SimulationProfile? Profile,
        WorkbookUserSettingsOverrides? WorkbookSettingsOverrides);

    public void Save(SimulationProfile profile)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        var existing = LoadStoredConfig(workbook);
        SaveStoredConfig(workbook, profile, existing.WorkbookSettingsOverrides);
    }

    public SimulationProfile? Load()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return null;

        var profile = LoadStoredConfig(workbook).Profile;
        if (profile != null) return profile;

        return LoadFromHiddenSheet(workbook);
    }

    public void SaveWorkbookSettingsOverrides(WorkbookUserSettingsOverrides? workbookSettingsOverrides)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return;

        var existing = LoadStoredConfig(workbook);
        SaveStoredConfig(workbook, existing.Profile, workbookSettingsOverrides);
    }

    public WorkbookUserSettingsOverrides? LoadWorkbookSettingsOverrides()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var workbook = app.ActiveWorkbook;
        if (workbook == null) return null;

        return LoadStoredConfig(workbook).WorkbookSettingsOverrides;
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

    private static string WrapInXml(
        SimulationProfile? profile,
        WorkbookUserSettingsOverrides? workbookSettingsOverrides)
    {
        var builder = new StringBuilder();
        builder.AppendLine("""<?xml version="1.0" encoding="utf-8"?>""");
        builder.AppendLine($"""<MonteCarloConfig xmlns="{CustomXmlNamespace}">""");

        if (profile != null)
        {
            var profileJson = JsonSerializer.Serialize(profile, JsonOptions);
            builder.AppendLine(
                $"""  <Profile name="{System.Security.SecurityElement.Escape(profile.Name)}"><![CDATA[{EscapeCData(profileJson)}]]></Profile>""");
        }

        if (workbookSettingsOverrides != null && workbookSettingsOverrides.HasAnyValues)
        {
            var settingsJson = JsonSerializer.Serialize(workbookSettingsOverrides, JsonOptions);
            builder.AppendLine(
                $"""  <WorkbookSettings><![CDATA[{EscapeCData(settingsJson)}]]></WorkbookSettings>""");
        }

        builder.AppendLine("""</MonteCarloConfig>""");
        return builder.ToString();
    }

    private static string EscapeCData(string text) =>
        text.Replace("]]>", "]]]]><![CDATA[>");

    private static StoredConfigData LoadStoredConfig(Workbook workbook)
    {
        var config = LoadFromCustomXml(workbook);
        if (config.Profile != null || config.WorkbookSettingsOverrides != null)
            return config;

        return new StoredConfigData(LoadFromHiddenSheet(workbook), null);
    }

    private static StoredConfigData LoadFromCustomXml(Workbook workbook)
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

                SimulationProfile? profile = null;
                var profileNode = doc.SelectSingleNode("//mc:Profile", nsManager);
                if (profileNode != null && !string.IsNullOrWhiteSpace(profileNode.InnerText))
                {
                    profile = JsonSerializer.Deserialize<SimulationProfile>(profileNode.InnerText, new JsonSerializerOptions
                    {
                        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                    });
                }

                WorkbookUserSettingsOverrides? workbookSettings = null;
                var settingsNode = doc.SelectSingleNode("//mc:WorkbookSettings", nsManager);
                if (settingsNode != null && !string.IsNullOrWhiteSpace(settingsNode.InnerText))
                {
                    workbookSettings = JsonSerializer.Deserialize<WorkbookUserSettingsOverrides>(settingsNode.InnerText, new JsonSerializerOptions
                    {
                        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                    });
                }

                return new StoredConfigData(profile, workbookSettings);
            }
            catch { }
        }

        return new StoredConfigData(null, null);
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

    private static void SaveStoredConfig(
        Workbook workbook,
        SimulationProfile? profile,
        WorkbookUserSettingsOverrides? workbookSettingsOverrides)
    {
        RemoveExistingPart(workbook);

        if (profile == null && (workbookSettingsOverrides == null || !workbookSettingsOverrides.HasAnyValues))
        {
            RemoveHiddenSheet(workbook);
            return;
        }

        string xml = WrapInXml(profile, workbookSettingsOverrides);

        dynamic parts = workbook.CustomXMLParts;
        parts.Add(xml);
        RemoveHiddenSheet(workbook);
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
