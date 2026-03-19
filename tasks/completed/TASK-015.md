# TASK-015: Iman-Conover Correlation Engine

## Context

Read `ROADMAP.md` for full project context. In real models, inputs are rarely independent — material costs correlate with labor costs, interest rates correlate with property values, etc. Ignoring correlation overstates the spread of outputs and produces misleading results. This task implements the Iman-Conover method to induce rank correlation between simulation inputs.

## Dependencies

- TASK-002 (IDistribution — for sampling)
- TASK-003 (SimulationEngine — the engine calls correlation as step 5 in its loop)

## Objective

Build `MonteCarlo.Engine/Correlation/` with the Iman-Conover algorithm that takes an independently sampled input matrix and rearranges the samples to match a target rank-correlation matrix, while preserving each input's marginal distribution.

## Background: What Iman-Conover Does

Given:
- A sample matrix `X` of dimensions [N iterations × K inputs], where each column is independently sampled from its distribution
- A target rank-correlation matrix `C` of dimensions [K × K]

The algorithm rearranges the rows within each column of `X` so that the Spearman rank correlations between columns match `C`, without changing any column's marginal distribution (the same values are present, just reordered).

This is the same approach @RISK uses.

## Design

### CorrelationMatrix

```csharp
public class CorrelationMatrix
{
    public int Size { get; }                          // K × K
    public double this[int i, int j] { get; set; }    // Access elements

    /// Create a K×K identity matrix (no correlation — default).
    public static CorrelationMatrix Identity(int size);

    /// Create from a flat 2D array.
    public CorrelationMatrix(double[,] matrix);

    /// Validate that the matrix is:
    /// 1. Square
    /// 2. Symmetric: C[i,j] == C[j,i]
    /// 3. Diagonal is all 1.0
    /// 4. All values in [-1, 1]
    /// 5. Positive semi-definite (PSD) — all eigenvalues ≥ 0
    /// Throws ArgumentException with specific message on failure.
    public void Validate();

    /// Check if PSD. If not, attempt to find the nearest PSD matrix.
    /// Returns the adjusted matrix if correction was needed, or this if already PSD.
    public CorrelationMatrix EnsurePositiveSemiDefinite();

    /// Get the underlying array for matrix operations.
    public double[,] ToArray();
}
```

### ImanConover

```csharp
public static class ImanConover
{
    /// Apply Iman-Conover rank correlation to a sample matrix.
    ///
    /// Parameters:
    ///   samples - [N × K] matrix of independently sampled values.
    ///             Each column is one input's samples. Modified in place.
    ///   targetCorrelation - [K × K] target Spearman rank correlation matrix.
    ///
    /// Returns: the rearranged sample matrix (same object, modified in place).
    public static double[,] Apply(double[,] samples, CorrelationMatrix targetCorrelation);
}
```

## Algorithm — Step by Step

Reference: Iman, R.L. and Conover, W.J. (1982), "A distribution-free approach to inducing rank correlation among input variables," Communications in Statistics - Simulation and Computation, 11(3), 311-334.

```
Input: X [N × K] sample matrix, C [K × K] target rank correlation matrix

1. Rank-transform X:
   R[i,j] = rank of X[i,j] within column j (average ranks for ties)

2. Create a score matrix S from the ranks:
   For each column j, for each row i:
     S[i,j] = Φ⁻¹(R[i,j] / (N + 1))
   where Φ⁻¹ is the inverse standard normal CDF (van der Waerden scores).

3. Compute the current correlation of S:
   T = correlation matrix of S (Pearson correlation of the scores)

4. Cholesky decomposition:
   P = Cholesky(C)   — lower triangular matrix from TARGET correlation
   Q = Cholesky(T)   — lower triangular matrix from CURRENT correlation

5. Transformation matrix:
   M = P × Q⁻¹

6. Apply transformation to score matrix:
   S* = S × Mᵀ       (each row of S is transformed)

7. Rank-transform S* to get target ranks:
   R*[i,j] = rank of S*[i,j] within column j

8. Rearrange X according to R*:
   For each column j:
     Sort the original values X[*,j]
     Assign the k-th smallest value to the row that has rank k in R*[*,j]

Result: X now has the target rank correlation structure, and each column
        contains exactly the same values as before (marginal distributions preserved).
```

### Key Implementation Details

**Ranking:** Use average ranking for ties (same as RANK.AVG in Excel). `MathNet.Numerics.Statistics.SortedArrayStatistics` or write a simple ranking function.

**Cholesky decomposition:** Use `MathNet.Numerics.LinearAlgebra.Matrix<double>.Cholesky()`. If the target matrix is not positive definite, Cholesky will fail — this is caught by `CorrelationMatrix.Validate()`.

**Matrix inverse:** Q⁻¹ can be computed as `Q.Inverse()` from Math.NET, or more efficiently by solving `Q × Z = I` via forward substitution (Cholesky factors are triangular).

**Van der Waerden scores:** `MathNet.Numerics.Distributions.Normal.InvCDF(0, 1, rank / (N + 1))`. Using `N + 1` instead of `N` in the denominator avoids ±∞ at the extremes.

### Nearest PSD Matrix (for invalid user inputs)

If the user enters a correlation matrix that isn't positive semi-definite (very common with hand-entered matrices), use the Higham (2002) nearest correlation matrix algorithm, or a simpler approach:

1. Compute eigendecomposition: `C = V × D × Vᵀ`
2. Set any negative eigenvalues to a small positive number (e.g., 1e-10)
3. Reconstruct: `C' = V × D' × Vᵀ`
4. Rescale diagonal back to 1.0

Math.NET's `Matrix<double>.Evd()` provides eigendecomposition.

## Integration with SimulationEngine

In TASK-003, the simulation loop has this comment at step 5:

```
5. (Phase 3: Apply Iman-Conover correlation if config.Correlation != null)
```

Now implement it:

```csharp
// In SimulationEngine.RunAsync, after step 4 (generating all samples):
if (config.Correlation != null)
{
    ImanConover.Apply(inputSampleMatrix, config.Correlation);
}
```

This is a one-line integration. The correlation is applied to the pre-generated sample matrix before the iteration loop begins.

## Tests — MonteCarlo.Engine.Tests/Correlation/

### CorrelationMatrixTests.cs

1. **Identity matrix is valid**
2. **Symmetric matrix with valid values passes validation**
3. **Non-symmetric throws** — C[0,1] ≠ C[1,0]
4. **Diagonal not 1.0 throws**
5. **Values outside [-1,1] throws**
6. **Non-PSD matrix detected and nearest PSD computed**
7. **Nearest PSD matrix IS actually PSD** (verify eigenvalues ≥ 0)

### ImanConoverTests.cs

1. **Marginal distributions preserved** — sort each column before and after; values are identical
2. **Target correlation achieved** — apply correlation matrix [[1, 0.8], [0.8, 1]] to 2 inputs (100k samples). Compute Spearman correlation of result. Assert it's within 0.02 of 0.8.
3. **Negative correlation** — target [[1, -0.7], [-0.7, 1]]. Verify negative Spearman ρ ≈ -0.7.
4. **Identity correlation changes nothing** — samples should be in the same order (or at least have near-zero correlation, same as the independent case)
5. **3×3 correlation** — verify all 3 pairwise correlations match targets
6. **Large N convergence** — with N=50,000, achieved correlation should be very close to target (within 0.005)
7. **Small N (100)** — should still work, but with larger tolerance (within 0.05)

### Integration test with SimulationEngine

8. **End-to-end with engine** — create a 2-input config with correlation 0.9. Run 10,000 iterations. Compute Spearman correlation of the input samples from SimulationResult. Assert ≈ 0.9.

## File Structure

```
MonteCarlo.Engine/
└── Correlation/
    ├── CorrelationMatrix.cs
    └── ImanConover.cs

MonteCarlo.Engine/
└── Simulation/
    └── SimulationEngine.cs           # Updated: apply correlation at step 5

MonteCarlo.Engine.Tests/
└── Correlation/
    ├── CorrelationMatrixTests.cs
    └── ImanConoverTests.cs
```

## Commit Strategy

```
feat(engine): add CorrelationMatrix with validation and nearest-PSD correction
feat(engine): implement Iman-Conover rank correlation algorithm
feat(engine): integrate Iman-Conover into SimulationEngine
test(engine): add correlation matrix validation tests
test(engine): add Iman-Conover tests — preservation, accuracy, edge cases
```

## Done When

- [ ] CorrelationMatrix validates symmetry, PSD, diagonal, and bounds
- [ ] Nearest-PSD correction works for invalid user inputs
- [ ] Iman-Conover correctly rearranges samples to achieve target rank correlation
- [ ] Marginal distributions are exactly preserved (same values, just reordered)
- [ ] Achieved correlation within 0.02 of target for N ≥ 10,000
- [ ] Integrated into SimulationEngine (one-line hook)
- [ ] All tests passing
- [ ] `dotnet build` clean, `dotnet test` green
