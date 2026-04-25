namespace MonteCarlo.Addin.Export;

public sealed record SummaryReportOptions(
    bool IncludeMetadata,
    bool IncludeSummaryStatistics,
    bool IncludePercentiles,
    bool IncludeTargetAnalysis,
    bool IncludeSensitivity,
    bool IncludeScenarioAnalysis,
    bool IncludeCharts,
    bool IncludeInputAssumptions,
    bool IncludeCorrelationAssumptions)
{
    public static SummaryReportOptions Default { get; } = new(
        IncludeMetadata: true,
        IncludeSummaryStatistics: true,
        IncludePercentiles: true,
        IncludeTargetAnalysis: true,
        IncludeSensitivity: true,
        IncludeScenarioAnalysis: true,
        IncludeCharts: true,
        IncludeInputAssumptions: true,
        IncludeCorrelationAssumptions: true);

    public bool IncludesAnySections =>
        IncludeMetadata ||
        IncludesAnyPerOutputSections ||
        IncludeInputAssumptions ||
        IncludeCorrelationAssumptions;

    public bool IncludesAnyPerOutputSections =>
        IncludeSummaryStatistics ||
        IncludePercentiles ||
        IncludeTargetAnalysis ||
        IncludeSensitivity ||
        IncludeScenarioAnalysis ||
        IncludeCharts;
}
