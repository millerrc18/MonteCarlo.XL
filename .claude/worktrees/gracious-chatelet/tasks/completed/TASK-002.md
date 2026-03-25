# TASK-002: Distribution Module Implementation

## Context

Read `ROADMAP.md` for full project context. This is a Monte Carlo simulation Excel add-in. The `MonteCarlo.Engine` project already has an `IDistribution` interface in `Distributions/`. Your job is to implement the six Phase 1 distributions and thorough test coverage.

## Objective

Implement all six Phase 1 distributions in `MonteCarlo.Engine/Distributions/`, plus a `DistributionFactory`, with comprehensive unit tests in `MonteCarlo.Engine.Tests/`.

## IDistribution Interface

The interface should support (adjust the existing placeholder if needed):

```csharp
public interface IDistribution
{
    string Name { get; }
    double Mean { get; }
    double Variance { get; }
    double StdDev { get; }
    double Minimum { get; }
    double Maximum { get; }

    double Sample();                          // Single random draw
    double[] Sample(int count);               // Batch random draws
    double PDF(double x);                     // Probability density at x
    double CDF(double x);                     // Cumulative probability P(X <= x)
    double Percentile(double p);              // Inverse CDF / quantile (p in 0..1)
    string ParameterSummary();                // Human-readable, e.g. "Normal(μ=100, σ=10)"
}
```

All distributions should accept an optional `Random` seed for reproducibility.

## Distributions to Implement

### 1. NormalDistribution
- Wraps `MathNet.Numerics.Distributions.Normal`
- Parameters: `mean`, `stdDev`
- Validation: `stdDev > 0`

### 2. TriangularDistribution
- Wraps `MathNet.Numerics.Distributions.Triangular`
- Parameters: `min`, `mode`, `max`
- Validation: `min < mode < max` (also handle `min == mode` or `mode == max` edge cases if Math.NET allows)

### 3. PERTDistribution
- **Math.NET does NOT have a native PERT distribution.** Implement as a transformed Beta distribution.
- Parameters: `min`, `mode`, `max`, optional `lambda` (shape parameter, default = 4)
- Implementation:
  ```
  alpha = 1 + lambda * (mode - min) / (max - min)
  beta  = 1 + lambda * (max - mode) / (max - min)
  ```
  Then sample from `Beta(alpha, beta)` and scale to `[min, max]`:
  ```
  value = min + sample * (max - min)
  ```
- PDF, CDF, Percentile all need the same scaling/transformation applied
- This is the trickiest one — pay extra attention to correctness

### 4. LognormalDistribution
- Wraps `MathNet.Numerics.Distributions.LogNormal`
- Parameters: `mu`, `sigma` (of the underlying normal, NOT of the lognormal itself)
- Validation: `sigma > 0`
- Note: `Mean` of the lognormal is `exp(mu + sigma^2 / 2)`, not `mu`

### 5. UniformDistribution
- Wraps `MathNet.Numerics.Distributions.ContinuousUniform`
- Parameters: `min`, `max`
- Validation: `min < max`

### 6. DiscreteDistribution
- Models a discrete set of outcomes with associated probabilities
- Parameters: `double[] values`, `double[] probabilities`
- Validation: probabilities must sum to 1.0 (within tolerance, e.g. 1e-9); all probabilities >= 0; arrays must be same length and non-empty
- `PDF(x)` returns the probability of the exact value (0 if not in the set)
- `CDF(x)` returns sum of probabilities for all values <= x
- `Percentile(p)` returns the smallest value where CDF >= p
- `Sample()` uses `MathNet.Numerics.Distributions.Categorical` internally, mapping category indices back to values
- `Mean` is the weighted sum of values
- `Minimum`/`Maximum` are min/max of the values array

## DistributionFactory

Create `DistributionFactory.cs`:

```csharp
public static class DistributionFactory
{
    // Create a distribution by name and parameter dictionary
    public static IDistribution Create(string name, Dictionary<string, double> parameters, int? seed = null);

    // List available distribution names
    public static IReadOnlyList<string> AvailableDistributions { get; }
}
```

The factory should be case-insensitive on name. This will be used by the UI layer to create distributions from user input.

## Tests — MonteCarlo.Engine.Tests

### Required Test Coverage

**For every distribution, test:**

1. **Construction & validation**
   - Valid parameters create successfully
   - Invalid parameters throw `ArgumentException` with a clear message (e.g., `stdDev <= 0`, `min >= max`, probabilities don't sum to 1)

2. **Statistical convergence** (100,000 samples)
   - Empirical mean within 1% of theoretical `Mean` (or within 0.5 for small means)
   - Empirical std dev within 2% of theoretical `StdDev`

3. **Quantile round-trip**
   - For p in {0.01, 0.1, 0.25, 0.5, 0.75, 0.9, 0.99}:
     `CDF(Percentile(p))` ≈ p (tolerance 1e-6)
   - This single test pattern catches most implementation bugs

4. **PDF integrates to 1**
   - Numerical integration of PDF over [Minimum, Maximum] ≈ 1.0 (tolerance 1e-3)
   - For Normal/Lognormal with infinite tails, integrate over [Mean - 6*StdDev, Mean + 6*StdDev]

5. **CDF boundary conditions**
   - `CDF(Minimum)` ≈ 0 (or exactly 0 for bounded distributions)
   - `CDF(Maximum)` ≈ 1

6. **Reproducibility**
   - Two distributions created with the same seed produce identical sample sequences

### PERT-specific tests

- Verify the PERT mean formula: `(min + lambda * mode + max) / (lambda + 2)` with default lambda=4
- Compare against a known reference: PERT(0, 3, 6) with lambda=4 should have mean = 3.0
- Verify it's smoother than Triangular: for the same min/mode/max, PERT should have lower kurtosis

### DiscreteDistribution-specific tests

- Sampling frequencies converge to specified probabilities (chi-squared or simple proportion check)
- CDF is a proper step function (test values between discrete points)

### DistributionFactory tests

- Creates each distribution type by name (case-insensitive)
- Throws on unknown distribution name
- Passes parameters through correctly
- `AvailableDistributions` returns all 6 names

## File Structure

```
MonteCarlo.Engine/
└── Distributions/
    ├── IDistribution.cs             (update existing)
    ├── NormalDistribution.cs
    ├── TriangularDistribution.cs
    ├── PERTDistribution.cs
    ├── LognormalDistribution.cs
    ├── UniformDistribution.cs
    ├── DiscreteDistribution.cs
    └── DistributionFactory.cs

MonteCarlo.Engine.Tests/
└── Distributions/
    ├── NormalDistributionTests.cs
    ├── TriangularDistributionTests.cs
    ├── PERTDistributionTests.cs
    ├── LognormalDistributionTests.cs
    ├── UniformDistributionTests.cs
    ├── DiscreteDistributionTests.cs
    └── DistributionFactoryTests.cs
```

## Conventions

- Use `ArgumentException` or `ArgumentOutOfRangeException` for invalid parameters — throw in constructors, fail fast
- All distributions should be immutable after construction (parameters are readonly)
- Use `System.Random` for RNG, accepting an optional seed. If no seed, use a default `Random()` instance
- Use descriptive `ParameterSummary()` output, e.g. `"PERT(min=10, mode=50, max=90, λ=4)"`
- Follow existing code style in the repo

## Commit Strategy

Commit to `main` with meaningful messages:
```
feat(engine): add IDistribution interface with full contract
feat(engine): implement Normal and Triangular distributions
feat(engine): implement PERT distribution as scaled Beta
feat(engine): implement Lognormal, Uniform, Discrete distributions
feat(engine): add DistributionFactory
test(engine): add distribution convergence and round-trip tests
test(engine): add PERT-specific validation and factory tests
```

Split into logical commits as you go — don't lump everything into one.

## Done When

- [ ] All 6 distributions implemented and following IDistribution contract
- [ ] DistributionFactory works for all 6 types
- [ ] All tests listed above are passing
- [ ] `dotnet build` succeeds with no warnings
- [ ] `dotnet test` passes with 0 failures
