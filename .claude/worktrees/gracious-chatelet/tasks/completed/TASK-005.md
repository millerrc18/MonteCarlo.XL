# TASK-005: Sensitivity Analysis Engine

## Context

Read `ROADMAP.md` for full project context. This module powers the tornado chart — it answers "which inputs drive the most variance in the output?" **No Excel dependency.** Pure engine code.

## Objective

Build `MonteCarlo.Engine/Analysis/SensitivityAnalysis.cs` that takes a `SimulationResult` and computes sensitivity metrics for each input relative to a chosen output.

## Design

### SensitivityResult

```csharp
public class SensitivityResult
{
    public string OutputId { get; }
    public List<InputSensitivity> InputSensitivities { get; }   // Sorted by |Impact| descending
}

public class InputSensitivity
{
    public string InputId { get; }
    public string InputLabel { get; }

    // Rank correlation (Spearman) between this input's samples and the output values
    // Range: -1 to +1. Sign indicates direction; magnitude indicates strength.
    public double RankCorrelation { get; }

    // Contribution to variance (from regression): proportion of output variance
    // explained by this input. Range: 0 to 1. All inputs' contributions sum to ~1.
    public double ContributionToVariance { get; }

    // Tornado chart data: what the output looks like when this input is at its extremes
    // while other inputs are at their base values
    public double OutputAtInputP10 { get; }    // Output when this input is at its 10th percentile
    public double OutputAtInputP90 { get; }    // Output when this input is at its 90th percentile
    public double Swing { get; }               // |OutputAtInputP90 - OutputAtInputP10|
}
```

### SensitivityAnalysis

```csharp
public static class SensitivityAnalysis
{
    /// Compute sensitivity using rank correlation (Spearman's rho).
    /// This is the primary method — robust, handles non-linear relationships.
    public static SensitivityResult ComputeRankCorrelation(
        SimulationResult result,
        string outputId
    );

    /// Compute contribution to variance using standardized regression coefficients.
    /// Runs a multiple regression of output on all inputs (on ranks), then
    /// squares the standardized coefficients to get variance contributions.
    public static SensitivityResult ComputeRegressionBased(
        SimulationResult result,
        string outputId
    );

    /// Compute tornado swing values.
    /// For each input: find iterations where that input is near its P10 and P90,
    /// compute the mean output in each group, and report the difference.
    public static SensitivityResult ComputeTornadoSwing(
        SimulationResult result,
        string outputId,
        Func<Dictionary<string, double>, Dictionary<string, double>>? evaluator = null
    );
}
```

## Implementation Notes

### Spearman Rank Correlation (Primary Method)

This is what @RISK uses by default. For each input:

1. Rank-transform the input samples (handle ties with average ranking)
2. Rank-transform the output values
3. Compute Pearson correlation on the ranked data

Use `MathNet.Numerics.Statistics.Correlation.Spearman()` if available, or implement manually:
```
ρ = 1 - (6 * Σ(d²)) / (n * (n² - 1))
```
where d = difference between input rank and output rank for each iteration.

### Standardized Regression Coefficients (Variance Contribution)

1. Rank-transform all inputs and the output
2. Standardize each ranked variable to mean=0, std=1
3. Fit an OLS regression: `output_rank ~ β₁*input1_rank + β₂*input2_rank + ...`
4. The squared standardized coefficients (β²) approximate each input's contribution to variance
5. Normalize so they sum to 1.0

For the regression, use `MathNet.Numerics.LinearRegression.MultipleRegression.QR()` or build the normal equations manually. With rank-transformed data, multicollinearity is minimal (unless inputs are correlated, which is Phase 3).

### Tornado Swing (Conditional Means)

For each input:
1. Find the iterations where that input falls in the bottom 10% (near P10) — take mean output
2. Find the iterations where that input falls in the top 10% (near P90) — take mean output
3. The swing = |mean_output_at_P90 - mean_output_at_P10|

This approach uses the actual simulation data and works well for non-linear models. It's more intuitive for stakeholders than regression coefficients.

Alternative (if an evaluator function is provided): directly evaluate the model with each input at its P10/P90 while holding all other inputs at their base (mean) values. This is cleaner but requires the evaluator.

**Implement both approaches.** The conditional-means approach works without an evaluator; the direct evaluation approach is optional and used when the evaluator is available.

### Sorting

Results should always be sorted by absolute impact descending (largest driver first). This is critical for the tornado chart — the biggest bar goes on top.

## Tests — MonteCarlo.Engine.Tests/Analysis/

### SensitivityAnalysisTests.cs

1. **Linear model, dominant input** — Output = 3*A + 1*B + 0.5*C where A~Normal(100,20), B~Normal(50,5), C~Normal(25,2).
   - A should have the highest rank correlation (it drives the most variance because 3*20 >> 1*5 >> 0.5*2)
   - Rank correlations should all be positive
   - Results should be sorted: A first, B second, C third

2. **Negative relationship** — Output = 100 - 2*A + B. A should have a negative rank correlation.

3. **Independent input** — add an input D~Normal(0,1) that doesn't appear in the output formula. Its rank correlation should be ≈ 0 (within tolerance, say |ρ| < 0.05).

4. **Variance contribution sums to ~1** — for the linear model, assert that contributions sum to between 0.90 and 1.10 (allowing for sampling noise with finite iterations).

5. **Tornado swing direction** — for Output = A + B: swing should show higher output when inputs are at P90 vs P10.

6. **Single input** — should work, rank correlation ≈ 1.0 if output is a monotonic function of input.

7. **All inputs equal contribution** — Output = A + B with A~Normal(0,10) and B~Normal(0,10). Both should have rank correlations ≈ 0.7 (since each explains ~50% of variance, and 1/√2 ≈ 0.707).

## File Structure

```
MonteCarlo.Engine/
└── Analysis/
    ├── SensitivityAnalysis.cs
    ├── SensitivityResult.cs
    └── InputSensitivity.cs

MonteCarlo.Engine.Tests/
└── Analysis/
    └── SensitivityAnalysisTests.cs
```

## Commit Strategy

```
feat(engine): add SensitivityResult and InputSensitivity models
feat(engine): implement rank correlation sensitivity analysis
feat(engine): implement regression-based variance contribution
feat(engine): implement tornado swing computation
test(engine): add sensitivity analysis tests — linear models, direction, independence
```

## Done When

- [ ] Rank correlation computed correctly for each input-output pair
- [ ] Variance contributions computed and normalized
- [ ] Tornado swing values computed (both conditional-means and direct evaluation)
- [ ] Results sorted by absolute impact
- [ ] All tests passing
- [ ] `dotnet build` clean, `dotnet test` green
