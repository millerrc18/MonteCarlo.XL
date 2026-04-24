using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.Shared.Formula;
using MonteCarlo.Shared.Interop;
using MonteCarlo.Shared.Simulation;

namespace MonteCarlo.Engine.Wasm;

[SupportedOSPlatform("browser")]
public static partial class BridgeExports
{
    private static readonly SimulationSamplingService SamplingService = new();
    private static readonly SimulationAnalysisService AnalysisService = new();

    [JSExport]
    public static string GetFormulaCatalogJson()
    {
        var response = new FormulaCatalogResponse(
            BridgeProtocol.Version,
            McFormulaCatalog.Definitions.Select(definition => new FormulaDefinitionDto(
                definition.FunctionName,
                definition.ExcelName,
                definition.DistributionName,
                definition.Description,
                definition.Arguments.Select(argument => new FormulaArgumentDto(
                    argument.Name,
                    argument.DisplayName,
                    argument.Description)).ToArray()))
                .ToArray());

        return BridgeJson.Serialize(response);
    }

    [JSExport]
    public static string ParseFormulaJson(string formula)
    {
        if (!McFormulaParser.TryParse(formula, out var parsed) || parsed == null)
            return BridgeJson.Serialize(new ParseFormulaResponse(BridgeProtocol.Version, false, null, null, null, null));

        return BridgeJson.Serialize(new ParseFormulaResponse(
            BridgeProtocol.Version,
            true,
            parsed.FunctionName,
            parsed.DistributionName,
            parsed.RawArguments.ToArray(),
            parsed.Arguments.Select(argument => argument.Name).ToArray()));
    }

    [JSExport]
    public static string EvaluateFormulaJson(string requestJson)
    {
        var request = BridgeJson.Deserialize<EvaluateFormulaRequest>(requestJson);
        EnsureVersion(request.Version);

        var isValid = McNormalModeEvaluator.TryEvaluate(request.FunctionName, request.Arguments, out var value);
        return BridgeJson.Serialize(new EvaluateFormulaResponse(
            BridgeProtocol.Version,
            isValid,
            isValid ? value : null));
    }

    [JSExport]
    public static string ValidateProfileJson(string requestJson)
    {
        var request = BridgeJson.Deserialize<ValidateProfileRequest>(requestJson);
        return BridgeJson.Serialize(SamplingService.ValidateProfile(request));
    }

    [JSExport]
    public static string GenerateInputMatrixJson(string requestJson)
    {
        var request = BridgeJson.Deserialize<GenerateInputMatrixRequest>(requestJson);
        return BridgeJson.Serialize(SamplingService.GenerateInputMatrix(request));
    }

    [JSExport]
    public static string AnalyzeSimulationJson(string requestJson)
    {
        var request = BridgeJson.Deserialize<AnalyzeSimulationRequest>(requestJson);
        return BridgeJson.Serialize(AnalysisService.Analyze(request));
    }

    [JSExport]
    public static string EvaluateGoalSeekMetricJson(string requestJson)
    {
        var request = BridgeJson.Deserialize<EvaluateGoalSeekMetricRequest>(requestJson);
        return BridgeJson.Serialize(AnalysisService.EvaluateGoalSeekMetric(request));
    }

    [JSExport]
    public static string RunBenchmarkJson(string requestJson)
    {
        var request = BridgeJson.Deserialize<BenchmarkRequest>(requestJson);
        EnsureVersion(request.Version);

        var result = SimulationBenchmark.RunAsync(request.InputCount, request.IterationCount)
            .GetAwaiter()
            .GetResult();

        return BridgeJson.Serialize(new BenchmarkResponse(
            BridgeProtocol.Version,
            result.InputCount,
            result.IterationCount,
            result.IterationsPerSecond,
            result.MicrosecondsPerIteration,
            result.TotalTime.TotalMilliseconds));
    }

    private static void EnsureVersion(int version)
    {
        if (version != BridgeProtocol.Version)
            throw new InvalidOperationException($"Unsupported bridge version {version}. Expected {BridgeProtocol.Version}.");
    }
}
