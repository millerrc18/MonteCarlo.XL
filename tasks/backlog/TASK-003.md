# TASK-003: Simulation Engine — Core Monte Carlo Loop

**Priority**: 🔴 Critical
**Phase**: 1 — Walking Skeleton
**Depends On**: TASK-002 (distributions)
**Estimated Effort**: ~3 hours

---

## Objective

Build the core simulation engine that takes a set of input distributions, runs N iterations, and produces a results matrix with summary statistics. Pure C# in `MonteCarlo.Engine` — no Excel, no UI.

---

## Deliverables

### 1. SimulationConfig

```csharp
public class SimulationInput
{
    public string Id { get; set; }               // Unique identifier (maps to cell ref later)
    public string Label { get; set; }             // Human-readable label (e.g., "Material Cost")
    public IDistribution Distribution { get; set; }
}

public class SimulationOutput
{
    public string Id { get; set; }
    public string Label { get; set; }
    // In "fast mode" (Phase 1), outputs are evaluated via a user-provided formula delegate.
    // In "recalc mode" (Phase 2), they map to Excel cell references.
}

public class SimulationConfig
{
    public List<SimulationInput> Inputs { get; set; }
    public List<SimulationOutput> Outputs { get; set; }
    public int IterationCount { get; set; } = 5000;
    public int? RandomSeed { get; set; }          // Null = random seed each run
}
```

### 2. SimulationResult

```csharp
public class SimulationResult
{
    public int IterationCount { get; }
    public TimeSpan Elapsed { get; }
    
    // Raw data: [iterationIndex][inputIndex] = sampled value
    public double[][] InputSamples { get; }
    
    // Raw data: [iterationIndex][outputIndex] = computed value
    public double[][] OutputSamples { get; }
    
    // Pre-computed stats for each output
    public OutputStatistics[] OutputStats { get; }
    
    // Config reference (for labeling in charts)
    public SimulationConfig Config { get; }
}

public class OutputStatistics
{
    public string OutputId { get; }
    public string OutputLabel { get; }
    public double Mean { get; }
    public double Median { get; }
    public double StandardDeviation { get; }
    public double Variance { get; }
    public double Skewness { get; }
    public double Kurtosis { get; }
    public double Minimum { get; }
    public double Maximum { get; }
    public double[] Percentiles { get; }          // P1, P5, P10, P25, P50, P75, P90, P95, P99
    
    public double GetPercentile(double p);        // Arbitrary percentile lookup
    public double ProbabilityAbove(double value);  // P(X > value)
    public double ProbabilityBelow(double value);  // P(X <= value)
}
```

### 3. SimulationEngine

```csharp
public class SimulationEngine
{
    /// <summary>
    /// Run a simulation in "fast mode" — inputs are sampled, outputs are computed
    /// via delegate functions. No Excel dependency.
    /// </summary>
    /// <param name="config">Simulation configuration</param>
    /// <param name="outputFunctions">
    /// For each output, a function that takes the sampled input values (keyed by input ID)
    /// and returns the output value. This is how we evaluate the model without Excel.
    /// </param>
    /// <param name="progress">Optional progress callback (0.0 to 1.0)</param>
    /// <param name="cancellationToken">Cancellation support</param>
    public Task<SimulationResult> RunAsync(
        SimulationConfig config,
        IReadOnlyDictionary<string, Func<IReadOnlyDictionary<string, double>, double>> outputFunctions,
        IProgress<double>? progress = null,
        CancellationToken cancellationToken = default);
}
```

**Engine logic:**
1. Create `Random` instance (seeded or not per config)
2. For each iteration:
   a. Sample each input from its distribution → store in `InputSamples[i]`
   b. Build a dictionary of `{ inputId: sampledValue }`
   c. Evaluate each output function with the input dictionary → store in `OutputSamples[i]`
   d. Report progress every 100 iterations (or similar)
   e. Check cancellation token
3. Compute `OutputStatistics` for each output column
4. Return `SimulationResult`

### 4. SummaryStatistics Calculator

```csharp
public static class SummaryStatistics
{
    public static OutputStatistics Compute(string id, string label, double[] samples);
}
```

Uses `MathNet.Numerics.Statistics` for percentiles, moments, etc. The `Percentiles` array should contain: P1, P5, P10, P25, P50, P75, P90, P95, P99.

### 5. Unit Tests

- **Basic simulation**: 3 inputs (Normal, Triangular, Uniform), 1 output (sum of inputs). Verify result dimensions, stats are reasonable.
- **Seeded reproducibility**: Same seed produces identical results across two runs.
- **Cancellation**: Verify the engine respects CancellationToken.
- **Progress reporting**: Verify progress callback is invoked and increases monotonically.
- **Statistics accuracy**: Generate 100,000 samples from Normal(100, 10), verify mean ≈ 100, stddev ≈ 10, P50 ≈ 100, P5 ≈ ~83.5, P95 ≈ ~116.5 (within 1% tolerance).
- **Edge cases**: Single iteration, single input, zero-variance input (Uniform(5, 5) → should throw or handle gracefully).

---

## Acceptance Criteria

- [ ] `SimulationConfig`, `SimulationResult`, `OutputStatistics` data classes exist
- [ ] `SimulationEngine.RunAsync()` works with delegate-based output evaluation
- [ ] Summary statistics computed correctly (validated by tests against known distributions)
- [ ] Cancellation and progress reporting work
- [ ] Seeded simulations are reproducible
- [ ] All unit tests pass
- [ ] No references to Excel, WPF, or UI
