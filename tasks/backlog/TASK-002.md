# TASK-002: Distribution Engine — Core 6 Distributions

**Priority**: 🔴 Critical
**Phase**: 1 — Walking Skeleton
**Depends On**: TASK-001 (solution structure)
**Estimated Effort**: ~3 hours

---

## Objective

Implement the distribution abstraction layer and the first 6 distributions in `MonteCarlo.Engine`. This is pure math — no Excel, no UI.

---

## Deliverables

### 1. IDistribution Interface

```csharp
// MonteCarlo.Engine/Distributions/IDistribution.cs
public interface IDistribution
{
    string Name { get; }
    string ShortCode { get; }                    // e.g., "NORM", "TRI", "PERT"
    double Mean { get; }
    double Median { get; }
    double StandardDeviation { get; }
    double Minimum { get; }                      // Support bound (may be double.NegativeInfinity)
    double Maximum { get; }                      // Support bound (may be double.PositiveInfinity)
    
    double Sample(Random random);                // Draw one sample
    double[] SampleN(Random random, int n);      // Draw n samples (can be optimized per-distribution)
    double PDF(double x);                        // Probability density at x
    double CDF(double x);                        // Cumulative probability at x
    double Quantile(double p);                   // Inverse CDF — value at probability p
    
    IReadOnlyDictionary<string, double> Parameters { get; }  // Named params for serialization
}
```

### 2. Implement 6 Distributions

Each wraps `MathNet.Numerics.Distributions` internally. Each must validate parameters in the constructor (throw `ArgumentException` with a clear message for invalid params).

| Class | Wraps | Params | Validation |
|-------|-------|--------|------------|
| `NormalDistribution` | `Normal` | mean, stdDev | stdDev > 0 |
| `TriangularDistribution` | `Triangular` | min, mode, max | min < mode < max (allow min == mode or mode == max but not both) |
| `PERTDistribution` | Custom (Beta transform) | min, mode, max | min < max, min ≤ mode ≤ max |
| `LognormalDistribution` | `LogNormal` | mu, sigma | sigma > 0 |
| `UniformDistribution` | `ContinuousUniform` | min, max | min < max |
| `DiscreteDistribution` | `Categorical` + value map | values[], probabilities[] | lengths match, probs sum to ~1.0, all probs ≥ 0 |

#### PERT Distribution Implementation

PERT is not in Math.NET directly. It's implemented as a Beta distribution with shape parameters derived from the PERT min/mode/max:

```
alpha = 1 + 4 * (mode - min) / (max - min)
beta  = 1 + 4 * (max - mode) / (max - min)
```

Then scale the standard Beta(alpha, beta) from [0,1] to [min, max]:
```
sample = min + Beta.Sample(alpha, beta) * (max - min)
```

### 3. DistributionFactory

```csharp
// MonteCarlo.Engine/Distributions/DistributionFactory.cs
public static class DistributionFactory
{
    public static IDistribution Create(string name, Dictionary<string, double> parameters);
    public static IReadOnlyList<string> AvailableDistributions { get; }
}
```

This is used by the UI layer to create distributions from user input and by the config persistence layer to deserialize saved configurations.

### 4. Unit Tests

Create thorough tests in `MonteCarlo.Engine.Tests/Distributions/`:

For each distribution, test:
- **Construction**: Valid params succeed, invalid params throw `ArgumentException`
- **Sampling**: 10,000 samples — verify sample mean is within 5% of theoretical mean, sample std dev is within 10% of theoretical. Use a fixed seed for reproducibility.
- **PDF**: Spot-check known values (e.g., Normal(0,1).PDF(0) ≈ 0.3989)
- **CDF**: Spot-check known values (e.g., Normal(0,1).CDF(0) == 0.5)
- **Quantile**: Verify Quantile(CDF(x)) ≈ x for several x values (round-trip)
- **Bounds**: Verify all samples fall within [Minimum, Maximum]
- **Parameters dict**: Verify it contains the right keys and values

PERT-specific tests:
- Verify the mean matches the PERT formula: (min + 4*mode + max) / 6
- Verify samples are bounded by [min, max]
- Verify the distribution is more concentrated around the mode than Triangular

---

## Acceptance Criteria

- [ ] `IDistribution` interface exists with all specified members
- [ ] All 6 distribution classes implemented and constructors validate params
- [ ] `DistributionFactory.Create()` works for all 6 distributions
- [ ] Unit tests pass for all distributions (construction, sampling, PDF, CDF, quantile, bounds)
- [ ] No references to Excel, WPF, or UI in the Engine project
- [ ] Code has XML doc comments on all public members
