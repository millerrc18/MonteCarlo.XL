# TASK-004: Summary Statistics

## Context

Read `ROADMAP.md` for full project context. TASK-003 delivers `SimulationResult` with raw output matrices. This task builds the statistics module that computes everything the results dashboard needs to display.

## Objective

Build `MonteCarlo.Engine/Analysis/SummaryStatistics.cs` — a class that takes a `double[]` of simulation output values and computes a comprehensive statistical summary. This is the bridge between raw results and the UI. **No Excel dependency.**

## Design

### SummaryStatistics

```csharp
public class SummaryStatistics
{
    // Construction — takes the raw output values array
    public SummaryStatistics(double[] values);

    // Central tendency
    public double Mean { get; }
    public double Median { get; }               // Same as P50
    public double Mode { get; }                 // Estimated via KDE or histogram binning

    // Spread
    public double StdDev { get; }
    public double Variance { get; }
    public double Min { get; }
    public double Max { get; }
    public double Range { get; }

    // Shape
    public double Skewness { get; }
    public double Kurtosis { get; }             // Excess kurtosis (Normal = 0)

    // Percentiles — the core set the UI will display
    public double P1 { get; }
    public double P5 { get; }
    public double P10 { get; }
    public double P25 { get; }
    public double P50 { get; }
    public double P75 { get; }
    public double P90 { get; }
    public double P95 { get; }
    public double P99 { get; }

    // Arbitrary percentile
    public double Percentile(double p);         // p in 0..1

    // Confidence interval for the mean
    public (double Lower, double Upper) MeanConfidenceInterval(double confidence = 0.95);

    // Probability queries — these power the target line feature
    public double ProbabilityAbove(double threshold);   // P(X > threshold)
    public double ProbabilityBelow(double threshold);   // P(X <= threshold)
    public double ProbabilityBetween(double lower, double upper);

    // Histogram data — pre-binned for the chart
    public HistogramData ToHistogram(int binCount = 50);

    // The raw sorted values (for CDF chart)
    public double[] SortedValues { get; }
    public int Count { get; }
}
```

### HistogramData

```csharp
public class HistogramData
{
    public double[] BinEdges { get; }          // Length = binCount + 1
    public double[] BinCenters { get; }        // Length = binCount
    public int[] Frequencies { get; }           // Counts per bin
    public double[] RelativeFrequencies { get; } // Frequencies / totalCount
    public double BinWidth { get; }
}
```

## Implementation Notes

### Percentile Calculation
Use linear interpolation between sorted values (same method as Excel's `PERCENTILE.INC`). Given sorted array of n values and probability p:
```
rank = p * (n - 1)
lower = floor(rank)
upper = ceil(rank)
fraction = rank - lower
result = values[lower] + fraction * (values[upper] - values[lower])
```

### Mode Estimation
For continuous simulation output, there's no exact mode. Use histogram binning: compute a histogram with a reasonable bin count (Sturges' rule: `1 + 3.322 * log10(n)`), find the bin with the highest frequency, return its center. This is an estimate — document it as such.

### Histogram Binning
- Default to 50 bins (can be overridden)
- Use equal-width bins from Min to Max
- Edge case: if Min == Max (all values identical), return a single bin

### Mean Confidence Interval
Standard formula using the t-distribution (for unknown population variance):
```
CI = Mean ± t(α/2, n-1) * StdDev / sqrt(n)
```
Use `MathNet.Numerics.Distributions.StudentT` for the critical value.

### Probability Queries
These are simple empirical calculations on the sorted values array:
```
ProbabilityBelow(threshold) = count(values <= threshold) / totalCount
```
Use binary search (`Array.BinarySearch`) on sorted values for O(log n) performance.

### Performance
All statistics should be computed lazily on first access and cached. The constructor should only sort the array and store it. This way, if the UI only needs Mean and P50, we don't compute kurtosis unnecessarily.

Alternatively, compute everything eagerly in the constructor since it's a single pass through sorted data and the array sizes (5,000–50,000) make this trivially fast.

Pick whichever approach you prefer — just be consistent.

## Tests — MonteCarlo.Engine.Tests/Analysis/

### SummaryStatisticsTests.cs

1. **Known distribution validation** — generate 100,000 samples from Normal(100, 10), compute stats, assert:
   - Mean ≈ 100 (within 0.1)
   - StdDev ≈ 10 (within 0.2)
   - Median ≈ 100 (within 0.2)
   - Skewness ≈ 0 (within 0.05)
   - Excess Kurtosis ≈ 0 (within 0.1)

2. **Percentile accuracy** — for Normal(0, 1) with 100k samples:
   - P50 ≈ 0 (within 0.02)
   - P5 ≈ -1.645 (within 0.05)
   - P95 ≈ 1.645 (within 0.05)

3. **Skewed distribution** — Lognormal samples should have positive skewness

4. **Probability queries**:
   - For Normal(100, 10): `ProbabilityAbove(100)` ≈ 0.5 (within 0.01)
   - For Normal(100, 10): `ProbabilityBelow(80)` ≈ 0.0228 (within 0.005)
   - `ProbabilityBetween(lower, upper)` = `ProbabilityBelow(upper) - ProbabilityBelow(lower)`

5. **Confidence interval** — for 100k Normal(0, 1) samples, the 95% CI for the mean should contain 0

6. **Histogram**:
   - Frequencies sum to total count
   - BinEdges has binCount + 1 elements
   - BinEdges[0] == Min, BinEdges[last] == Max
   - All frequencies >= 0

7. **Edge cases**:
   - Single value: Mean == Median == that value, StdDev == 0
   - Two values: verify percentiles interpolate correctly
   - All identical values: StdDev == 0, histogram has 1 bin

8. **Deterministic values** — use `[5, 10, 15, 20, 25]` to verify exact Mean (15), exact Min/Max (5/25), exact Median (15)

### HistogramDataTests.cs

1. Uniform(0, 1) samples → histogram bins should have roughly equal frequencies
2. Custom bin count is respected
3. Single-value edge case handled

## File Structure

```
MonteCarlo.Engine/
└── Analysis/
    ├── SummaryStatistics.cs
    └── HistogramData.cs

MonteCarlo.Engine.Tests/
└── Analysis/
    ├── SummaryStatisticsTests.cs
    └── HistogramDataTests.cs
```

## Commit Strategy

```
feat(engine): add HistogramData model with binning logic
feat(engine): add SummaryStatistics — central tendency, spread, shape, percentiles
feat(engine): add probability queries and confidence intervals to SummaryStatistics
test(engine): add SummaryStatistics tests — known distributions, edge cases, histograms
```

## Done When

- [ ] SummaryStatistics computes all listed properties correctly
- [ ] Probability queries (above, below, between) work
- [ ] Histogram binning produces correct results
- [ ] Mean confidence interval uses t-distribution correctly
- [ ] Edge cases handled (single value, all identical, two values)
- [ ] All tests passing
- [ ] `dotnet build` clean, `dotnet test` green
