# TASK-004: Sensitivity Analysis Engine

**Priority**: 🟡 High
**Phase**: 2 — Storytelling Layer
**Depends On**: TASK-003 (simulation engine)
**Estimated Effort**: ~2 hours

---

## Objective

Implement the sensitivity analysis module that determines how much each input contributes to output variance. This data drives the tornado chart.

---

## Deliverables

### 1. SensitivityAnalysis Class

```csharp
public class SensitivityResult
{
    public string InputId { get; }
    public string InputLabel { get; }
    public double RankCorrelation { get; }        // Spearman rank correlation with output
    public double RegressionCoefficient { get; }  // Standardized regression coefficient
    public double ContributionToVariance { get; } // Percentage of output variance explained
    public double OutputAtInputP10 { get; }       // Output value when this input is at its P10
    public double OutputAtInputP90 { get; }       // Output value when this input is at its P90
}

public static class SensitivityAnalysis
{
    /// <summary>
    /// Compute sensitivity of a specific output to all inputs.
    /// Uses rank correlation (Spearman) as the primary measure.
    /// </summary>
    public static List<SensitivityResult> Analyze(
        SimulationResult result,
        int outputIndex);
}
```

### 2. Implementation Approach

**Spearman Rank Correlation**: For each input column and the target output column, compute the Spearman rank correlation. This is the primary sensitivity measure and what the tornado chart displays.

**Contribution to Variance**: Square the rank correlations, normalize so they sum to 100%. This gives "percentage of output variance explained by each input."

**Output Range at Input Extremes**: For each input, find iterations where that input's sample falls in the bottom 10% (P10 bucket) and top 10% (P90 bucket). Compute the mean output in each bucket. The difference shows the output swing caused by that input.

### 3. Unit Tests

- **Known sensitivity**: Create a simulation where Output = 3*A + 1*B + 0*C. Verify A has the highest sensitivity, B second, C near zero.
- **Rank ordering**: Verify results are sorted by absolute sensitivity.
- **Variance contribution**: Verify contributions sum to approximately 100% (within tolerance due to correlation between inputs).
- **Output ranges**: Verify P10/P90 output ranges are logically consistent (P90 side should be higher for positively correlated inputs).

---

## Acceptance Criteria

- [ ] `SensitivityAnalysis.Analyze()` returns correct rank correlations
- [ ] Results are sorted by absolute impact
- [ ] Variance contributions sum to ~100%
- [ ] Output-at-extremes values are computed
- [ ] Unit tests pass
- [ ] No UI or Excel dependencies
