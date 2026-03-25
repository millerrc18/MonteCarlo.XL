namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Sensitivity of a single input to a specific output.
/// </summary>
public class SensitivityResult
{
    /// <summary>Input identifier.</summary>
    public required string InputId { get; init; }

    /// <summary>Human-readable input label.</summary>
    public required string InputLabel { get; init; }

    /// <summary>Spearman rank correlation between this input and the output.</summary>
    public double RankCorrelation { get; init; }

    /// <summary>Percentage of output variance explained by this input (0–100).</summary>
    public double ContributionToVariance { get; init; }

    /// <summary>Mean output value when this input is in its bottom 10%.</summary>
    public double OutputAtInputP10 { get; init; }

    /// <summary>Mean output value when this input is in its top 10%.</summary>
    public double OutputAtInputP90 { get; init; }

    /// <summary>Absolute swing: |OutputAtInputP90 - OutputAtInputP10|.</summary>
    public double Swing => Math.Abs(OutputAtInputP90 - OutputAtInputP10);
}
