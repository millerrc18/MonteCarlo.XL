using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Shared.Interop;

public static class BridgeProtocol
{
    public const int Version = 1;
}

public sealed record WorkbookCellDto(string SheetName, string CellAddress)
{
    public string FullReference => $"'{SheetName}'!{CellAddress}";
}

public sealed record OfficeRunSettings(
    int IterationCount,
    int? RandomSeed,
    SamplingMethod SamplingMethod,
    bool AutoStopOnConvergence,
    double? OutputTarget = null);

public sealed record GenerateInputMatrixRequest(
    int Version,
    SimulationProfile Profile,
    OfficeRunSettings Settings);

public sealed record FormulaCatalogResponse(
    int Version,
    IReadOnlyList<FormulaDefinitionDto> Functions);

public sealed record FormulaDefinitionDto(
    string FunctionName,
    string ExcelName,
    string DistributionName,
    string Description,
    IReadOnlyList<FormulaArgumentDto> Arguments);

public sealed record FormulaArgumentDto(
    string Name,
    string DisplayName,
    string Description);

public sealed record ParseFormulaResponse(
    int Version,
    bool IsMatch,
    string? FunctionName,
    string? DistributionName,
    IReadOnlyList<string>? RawArguments,
    IReadOnlyList<string>? ParameterNames);

public sealed record EvaluateFormulaRequest(
    int Version,
    string FunctionName,
    double[] Arguments);

public sealed record EvaluateFormulaResponse(
    int Version,
    bool IsValid,
    double? Value);

public sealed record GenerateInputMatrixResponse(
    int Version,
    string[] InputIds,
    string[] OutputIds,
    double[][] InputMatrix);

public sealed record AnalyzeSimulationRequest(
    int Version,
    SimulationProfile Profile,
    OfficeRunSettings Settings,
    double[][] InputMatrix,
    double[][] OutputMatrix,
    double ElapsedMilliseconds);

public sealed record AnalyzeSimulationResponse(
    int Version,
    int IterationCount,
    double ElapsedMilliseconds,
    IReadOnlyList<OutputAnalysisDto> Outputs);

public sealed record OutputAnalysisDto(
    string OutputId,
    string OutputLabel,
    SummaryStatisticsDto Summary,
    HistogramDto Histogram,
    CdfDto Cdf,
    IReadOnlyList<SensitivityItemDto> Sensitivity,
    ScenarioSummaryDto WorstCase,
    ScenarioSummaryDto BestCase,
    TargetAnalysisDto? Target);

public sealed record SummaryStatisticsDto(
    int Count,
    double Mean,
    double Median,
    double Mode,
    double StdDev,
    double Variance,
    double Min,
    double Max,
    double P1,
    double P5,
    double P10,
    double P25,
    double P50,
    double P75,
    double P90,
    double P95,
    double P99,
    double ProbabilityAboveTarget,
    double ProbabilityAtOrBelowTarget);

public sealed record HistogramDto(
    double[] BinEdges,
    double[] BinCenters,
    int[] Frequencies,
    double[] RelativeFrequencies,
    double[] KdeX,
    double[] KdeY);

public sealed record CdfDto(
    double[] X,
    double[] Y);

public sealed record SensitivityItemDto(
    string InputId,
    string InputLabel,
    double RankCorrelation,
    double StandardizedRegression);

public sealed record ScenarioSummaryDto(
    string Description,
    double Threshold,
    int MatchedIterations,
    double MatchedFraction,
    IReadOnlyList<ScenarioInputSummaryDto> Inputs);

public sealed record ScenarioInputSummaryDto(
    string InputId,
    string InputLabel,
    double OverallMean,
    double ScenarioMean,
    double Delta,
    double? DeltaPercent,
    double OverallP10,
    double OverallP50,
    double OverallP90,
    double ScenarioP10,
    double ScenarioP50,
    double ScenarioP90);

public sealed record TargetAnalysisDto(
    double Target,
    double ProbabilityAbove,
    double ProbabilityAtOrBelow);

public sealed record ValidateProfileRequest(
    int Version,
    SimulationProfile Profile);

public sealed record ValidateProfileResponse(
    int Version,
    IReadOnlyList<PreflightIssueDto> Issues);

public sealed record PreflightIssueDto(
    string Severity,
    string Code,
    string Title,
    string Message,
    string SuggestedAction);

public sealed record EvaluateGoalSeekMetricRequest(
    int Version,
    GoalSeekMetric Metric,
    double OutputTarget,
    double Percentile,
    double[] OutputValues);

public sealed record EvaluateGoalSeekMetricResponse(
    int Version,
    double MetricValue);

public sealed record BenchmarkRequest(
    int Version,
    int InputCount,
    int IterationCount);

public sealed record BenchmarkResponse(
    int Version,
    int InputCount,
    int IterationCount,
    double IterationsPerSecond,
    double MicrosecondsPerIteration,
    double TotalMilliseconds);
