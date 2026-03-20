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

    /// <summary>
    /// Optional correlation matrix between inputs. Null means independent (no correlation).
    /// Stored as a flat array in row-major order for JSON serialization.
    /// </summary>
    [JsonPropertyName("correlationMatrix")]
    public double[]? CorrelationMatrixFlat { get; set; }

    /// <summary>
    /// Size of the correlation matrix (K for a K×K matrix).
    /// </summary>
    [JsonPropertyName("correlationMatrixSize")]
    public int CorrelationMatrixSize { get; set; }

    /// <summary>
    /// Sets the correlation matrix from a 2D array.
    /// </summary>
    [JsonIgnore]
    public double[,]? CorrelationMatrix
    {
        get
        {
            if (CorrelationMatrixFlat == null || CorrelationMatrixSize <= 0)
                return null;

            int n = CorrelationMatrixSize;
            var matrix = new double[n, n];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                    matrix[i, j] = CorrelationMatrixFlat[i * n + j];
            return matrix;
        }
        set
        {
            if (value == null)
            {
                CorrelationMatrixFlat = null;
                CorrelationMatrixSize = 0;
                return;
            }

            int n = value.GetLength(0);
            CorrelationMatrixSize = n;
            CorrelationMatrixFlat = new double[n * n];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                    CorrelationMatrixFlat[i * n + j] = value[i, j];
        }
    }
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
