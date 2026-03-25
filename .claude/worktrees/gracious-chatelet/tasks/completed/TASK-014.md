# TASK-014: Additional Distributions (Beta, Weibull, Exponential, Poisson)

## Context

Read `ROADMAP.md` Section 6 for the full distribution roadmap. TASK-002 delivered the core 6 distributions. This task adds 4 more to cover the most common specialized use cases.

## Dependencies

- TASK-002 (IDistribution interface, DistributionFactory)

## Objective

Add 4 Phase 2 distributions to `MonteCarlo.Engine/Distributions/` and register them in the `DistributionFactory`. Follow the same patterns established in TASK-002.

## Distributions

### 1. BetaDistribution

- **Use case:** Percentages, conversion rates, yields, probabilities — any value bounded in [0, 1]
- **Wraps:** `MathNet.Numerics.Distributions.Beta`
- **Parameters:** `alpha` (α > 0), `beta` (β > 0)
- **Notes:**
  - α = β = 1 → Uniform(0, 1)
  - α = β → symmetric; α < β → left-skewed; α > β → right-skewed
  - Mean = α / (α + β)
  - Minimum = 0, Maximum = 1
- **ParameterSummary:** `"Beta(α=2, β=5)"`

### 2. WeibullDistribution

- **Use case:** Failure rates, time-to-event, reliability analysis, wind speed distributions
- **Wraps:** `MathNet.Numerics.Distributions.Weibull`
- **Parameters:** `shape` (k > 0), `scale` (λ > 0)
- **Notes:**
  - k = 1 → Exponential distribution
  - k = 2 → Rayleigh distribution
  - k > 1 → failure rate increases over time (aging)
  - k < 1 → failure rate decreases over time (burn-in)
  - Minimum = 0, Maximum = +∞
  - Check Math.NET's parameterization: some libraries use (shape, scale) and some use (shape, rate)
- **ParameterSummary:** `"Weibull(k=2.5, λ=100)"`

### 3. ExponentialDistribution

- **Use case:** Time between events (arrivals, failures), memoryless processes
- **Wraps:** `MathNet.Numerics.Distributions.Exponential`
- **Parameters:** `rate` (λ > 0) — the rate parameter, NOT the mean
- **Notes:**
  - Mean = 1/λ
  - Minimum = 0, Maximum = +∞
  - This is a special case of Weibull with shape = 1, but worth having separately because it's so commonly used and the parameterization is simpler
- **ParameterSummary:** `"Exponential(λ=0.5)"`

### 4. PoissonDistribution

- **Use case:** Count of events in a fixed interval (defects, arrivals, claims)
- **Wraps:** `MathNet.Numerics.Distributions.Poisson`
- **Parameters:** `lambda` (λ > 0) — expected number of events
- **Notes:**
  - This is a **discrete** distribution — Sample() returns integers (cast to double)
  - Mean = λ, Variance = λ
  - Minimum = 0, Maximum = +∞ (practically bounded)
  - PDF returns the probability mass function P(X = k)
  - CDF returns P(X ≤ k)
  - Percentile should return the smallest integer k where CDF(k) ≥ p
- **ParameterSummary:** `"Poisson(λ=4.5)"`

## DistributionFactory Update

Register all 4 new distributions in `DistributionFactory`:

```csharp
// Add to the factory's Create method:
"beta" => new BetaDistribution(params["alpha"], params["beta"], seed),
"weibull" => new WeibullDistribution(params["shape"], params["scale"], seed),
"exponential" => new ExponentialDistribution(params["rate"], seed),
"poisson" => new PoissonDistribution(params["lambda"], seed),
```

Update `AvailableDistributions` to include all 10.

## UI Integration

Update the distribution picker in the SetupView (TASK-008) to include the new distributions with appropriate parameter fields:

| Distribution | Fields |
|-------------|--------|
| Beta | Alpha (α), Beta (β) |
| Weibull | Shape (k), Scale (λ) |
| Exponential | Rate (λ) |
| Poisson | Rate (λ) |

These are additive changes to the `DistributionParameterPanel` — should not require restructuring.

## Tests

Follow the same test patterns from TASK-002:

### For each distribution:

1. **Construction & validation** — valid params succeed, invalid throw
2. **Statistical convergence** (100k samples) — mean and std dev within tolerance
3. **Quantile round-trip** — `CDF(Percentile(p)) ≈ p` for p in {0.01, 0.1, 0.25, 0.5, 0.75, 0.9, 0.99}
4. **CDF boundary conditions** — CDF(0) ≈ 0 for the non-negative distributions
5. **Reproducibility** — same seed → same samples

### Distribution-specific tests:

**Beta:**
- Beta(1, 1) should produce samples uniformly distributed in [0, 1]
- All samples must be in [0, 1]
- Beta(2, 5) mean ≈ 0.286

**Weibull:**
- All samples must be ≥ 0
- Weibull(1, λ) should behave like Exponential(1/λ) — compare means

**Exponential:**
- All samples must be ≥ 0
- Mean ≈ 1/λ
- Memoryless property (optional, harder to test): P(X > s+t | X > s) ≈ P(X > t)

**Poisson:**
- All samples must be non-negative integers
- Mean ≈ λ, Variance ≈ λ
- For large λ (e.g., 100), distribution should approximate Normal(λ, √λ)

### DistributionFactory:
- Factory creates all 10 distributions
- `AvailableDistributions` returns 10 items

## File Structure

```
MonteCarlo.Engine/
└── Distributions/
    ├── BetaDistribution.cs
    ├── WeibullDistribution.cs
    ├── ExponentialDistribution.cs
    ├── PoissonDistribution.cs
    └── DistributionFactory.cs          # Updated

MonteCarlo.Engine.Tests/
└── Distributions/
    ├── BetaDistributionTests.cs
    ├── WeibullDistributionTests.cs
    ├── ExponentialDistributionTests.cs
    ├── PoissonDistributionTests.cs
    └── DistributionFactoryTests.cs     # Updated
```

## Commit Strategy

```
feat(engine): add BetaDistribution
feat(engine): add WeibullDistribution
feat(engine): add ExponentialDistribution
feat(engine): add PoissonDistribution
feat(engine): register Phase 2 distributions in DistributionFactory
test(engine): add Phase 2 distribution tests
```

## Done When

- [ ] All 4 distributions implemented following IDistribution contract
- [ ] DistributionFactory creates all 10 distributions
- [ ] All tests passing (convergence, round-trip, edge cases)
- [ ] `dotnet build` clean, `dotnet test` green
