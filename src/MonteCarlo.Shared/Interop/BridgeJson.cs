using System.Text.Json;
using System.Text.Json.Serialization;

namespace MonteCarlo.Shared.Interop;

public static class BridgeJson
{
    public static readonly JsonSerializerOptions DefaultOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    public static string Serialize<T>(T value) =>
        JsonSerializer.Serialize(value, DefaultOptions);

    public static T Deserialize<T>(string json) =>
        JsonSerializer.Deserialize<T>(json, DefaultOptions)
        ?? throw new InvalidOperationException($"Unable to deserialize {typeof(T).Name}.");
}
