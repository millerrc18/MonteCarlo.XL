# TASK-003: Simulation Engine

## Context

Read `ROADMAP.md` for full project context. TASK-002 delivered the distribution module (`IDistribution` + 6 implementations + `DistributionFactory`). This task builds the core Monte Carlo simulation loop that consumes those distributions.

## Objective

Build the simulation engine in `MonteCarlo.Engine/Simulation/`. This is the heart of the product — it takes a set of input distributions and output definitions, runs N iterations of sampling, and produces a raw results matrix. **This module has ZERO Excel dependency.** It's pure C# math.

## Design

### SimulationConfig

```csharp
public class SimulationConfig
{
    public List<SimulationInput> Inputs { get; set; }
    public List<SimulationOutput> Outputs { get; set; }
    public int IterationCount { get; set; } = 5000;
    public int? RandomSeed { get; set; }           // null = non-deterministic
    public CorrelationMatrix? Correlation { get; set; }  // null = independent (Phase 3)
}

public class SimulationInput
{
    public string Id { get; set; }                  // Unique identifier (e.g., cell reference "B4")
    public string Label { get; set; }               // Human-readable name (e.g., "Material Cost")
    public IDistribution Distribution { get; set; }
    public double BaseValue { get; set; }           // The deterministic value currently in the cell
}

public class SimulationOutput
{
    public string Id { get; set; }                  // Unique identifier (e.g., cell reference "D10")
    public string Label { get; set; }               // Human-readable name (e.g., "Net Profit")
}
```

### SimulationResult

```csharp
public class SimulationResult
{
    public SimulationConfig Config { get; }
    public int IterationCount { get; }
    public TimeSpan ElapsedTime { get; }

    // Raw data access
    public double[] GetInputSamples(string inputId);           // All samples for one input
    public double[] GetOutputValues(string outputId);           // All output values
    public double GetInputSample(string inputId, int iteration);
    public double GetOutputValue(string outputId, int iteration);

    // The full matrices (iterations × inputs/outputs)
    public double[,] InputMatrix { get; }   // [iteration, inputIndex]
    public double[,] OutputMatrix { get; }  // [iteration, outputIndex]
}
```

### SimulationEngine

```csharp
public class SimulationEngine
{
    public event EventHandler<SimulationProgressEventArgs>? ProgressChanged;

    public Task<SimulationResult> RunAsync(
        SimulationConfig config,
        Func<Dictionary<string, double>, Dictionary<string, double>> evaluator,
        CancellationToken cancellationToken = default
    );
}

public class SimulationProgressEventArgs : EventArgs
{
    public int CompletedIterations { get; }
    public int TotalIterations { get; }
    public double PercentComplete { get; }
    public TimeSpan Elapsed { get; }
    public TimeSpan EstimatedRemaining { get; }
}
```

### The `evaluator` Function

This is the key abstraction. The engine doesn't know about Excel — it just calls a function that maps input values to output values:

```csharp
// For "fast mode" (no formula dependency), the evaluator is trivial:
// inputs go in, same values come out as outputs (the Excel layer handles the mapping)

// For "recalc mode" (Phase 3), the Excel layer provides an evaluator that:
// 1. Writes input values to cells
// 2. Triggers Application.Calculate()
// 3. Reads output cell values
// 4. Returns them
```

For now (Phase 1), implement and test with simple evaluator functions — the Excel integration comes later.

### Simulation Loop (Pseudocode)

```
1. Validate config (inputs non-empty, outputs non-empty, iterations > 0)
2. Pre-allocate input matrix [iterations × inputCount]
3. Pre-allocate output matrix [iterations × outputCount]
4. Generate all input samples upfront:
   For each input:
     samples[input] = input.Distribution.Sample(iterationCount)
5. (Phase 3: Apply Iman-Conover correlation if config.Correlation != null)
6. For each iteration (0..iterationCount-1):
     a. Build input dictionary: { inputId → samples[input][iteration] }
     b. Call evaluator(inputs) → outputs dictionary
     c. Store input samples in InputMatrix[iteration, *]
     d. Store output values in OutputMatrix[iteration, *]
     e. Report progress every ~100 iterations (don't fire on every iteration — too expensive)
     f. Check cancellationToken
7. Return SimulationResult with elapsed time
```

**Important: Sample all inputs upfront (step 4) rather than sampling inside the iteration loop.** This is both faster (batch sampling) and necessary for Iman-Conover correlation (Phase 3), which needs the full sample matrix to reorder.

## Performance Considerations

- Use `Parallel.For` for the evaluator calls across iterations when the evaluator is thread-safe (it will be for fast mode). Add a `bool ParallelExecution` option to `SimulationConfig` (default true). Recalc mode must be sequential (Excel COM is single-threaded).
- Progress reporting should be throttled — fire the event at most every 100 iterations or every 100ms, whichever comes first. Don't fire on every iteration.
- Pre-allocate arrays. No `List<double>` growing during the loop.
- The engine should comfortably handle 10,000 iterations × 20 inputs in under 2 seconds for fast mode (in-memory evaluator).

## Tests — MonteCarlo.Engine.Tests/Simulation/

### SimulationEngineTests.cs

1. **Basic execution** — 3 Normal inputs, 1 output (evaluator sums inputs), 1000 iterations. Assert:
   - Result has correct dimensions (1000 iterations, 3 inputs, 1 output)
   - Output mean ≈ sum of input means (within statistical tolerance)

2. **Reproducibility** — same seed produces identical results across two runs

3. **Different seeds produce different results**

4. **Cancellation** — start a 100,000 iteration run, cancel after 500ms, assert the task completes (doesn't hang) and throws `OperationCanceledException`

5. **Progress reporting** — verify ProgressChanged fires, percentages are monotonically increasing, and final event has CompletedIterations == TotalIterations

6. **Empty inputs/outputs** — throws `ArgumentException`

7. **Zero/negative iterations** — throws `ArgumentException`

8. **Single iteration** — edge case, should work correctly

9. **Large run** — 10,000 iterations, 10 inputs, verify it completes in under 5 seconds (performance sanity check, not a hard benchmark)

10. **Output values match evaluator** — use a deterministic evaluator (output = input1 * 2 + input2) and verify each row of the output matrix matches

### SimulationConfigTests.cs

1. Validation tests for the config object itself (if you add a `Validate()` method)

## File Structure

```
MonteCarlo.Engine/
└── Simulation/
    ├── SimulationConfig.cs
    ├── SimulationInput.cs
    ├── SimulationOutput.cs
    ├── SimulationResult.cs
    ├── SimulationEngine.cs
    └── SimulationProgressEventArgs.cs

MonteCarlo.Engine.Tests/
└── Simulation/
    ├── SimulationEngineTests.cs
    └── SimulationConfigTests.cs
```

## Commit Strategy

```
feat(engine): add SimulationConfig, SimulationInput, SimulationOutput models
feat(engine): add SimulationResult with matrix storage and accessors
feat(engine): implement SimulationEngine with async execution and progress reporting
test(engine): add simulation engine tests — execution, reproducibility, cancellation
```

## Done When

- [ ] SimulationEngine.RunAsync executes a full simulation with configurable inputs/outputs
- [ ] Progress reporting works with throttling
- [ ] Cancellation is respected
- [ ] Reproducible results with seeded RNG
- [ ] All tests passing
- [ ] `dotnet build` clean, `dotnet test` green
