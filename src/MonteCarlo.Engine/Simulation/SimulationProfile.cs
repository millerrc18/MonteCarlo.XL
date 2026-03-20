using System.Text.Json.Serialization;

namespace MonteCarlo.Engine.Simulation;

/// <summary>
/// Represents a complete simulation configuration that can be serialized and persisted.
/// </summary>
public class SimulationProfile
{
    /// <summary>Profile name (for future multi-profile support).</summary>
    [JsonPropertyName("name")]
    public string Name { get; set; } = "Default";

    /// <summary>Number of Monte Carlo iterations.</summary>
    [JsonPropertyName("iterationCount")]
    public int IterationCount { get; set; } = 5000;

    /// <summary>Optional random seed for reproducibility.</summary>
    [JsonPropertyName("randomSeed")]
    public int? RandomSeed { get; set; }

    /// <summary>Configured simulation inputs.</summary>
    [JsonPropertyName("inputs")]
    public List<SavedInput> Inputs { get; set; } = new();

    /// <summary>Configured simulation outputs.</summary>
    [JsonPropertyName("outputs")]
    public List<SavedOutput> Outputs { get; set; } = new();
}

/// <summary>
/// A saved input cell configuration with distribution parameters.
/// </summary>
public class SavedInput
{
    [JsonPropertyName("sheetName")]
    public string SheetName { get; set; } = string.Empty;

    [JsonPropertyName("cellAddress")]
    public string CellAddress { get; set; } = string.Empty;

    [JsonPropertyName("label")]
    public string Label { get; set; } = string.Empty;

    [JsonPropertyName("distributionName")]
    public string DistributionName { get; set; } = string.Empty;

    [JsonPropertyName("parameters")]
    public Dictionary<string, double> Parameters { get; set; } = new();
}

/// <summary>
/// A saved output cell configuration.
/// </summary>
public class SavedOutput
{
    [JsonPropertyName("sheetName")]
    public string SheetName { get; set; } = string.Empty;

    [JsonPropertyName("cellAddress")]
    public string CellAddress { get; set; } = string.Empty;

    [JsonPropertyName("label")]
    public string Label { get; set; } = string.Empty;
}
