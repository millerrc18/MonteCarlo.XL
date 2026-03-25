using MathNet.Numerics.Distributions;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace MonteCarlo.Engine.Correlation;

/// <summary>
/// Implements the Iman-Conover algorithm for inducing rank correlation
/// between simulation input samples while preserving marginal distributions.
/// </summary>
/// <remarks>
/// Reference: Iman, R.L. and Conover, W.J. (1982),
/// "A distribution-free approach to inducing rank correlation among input variables,"
/// Communications in Statistics - Simulation and Computation, 11(3), 311-334.
/// </remarks>
public static class ImanConover
{
    /// <summary>
    /// Applies the Iman-Conover rank correlation to a sample matrix in place.
    /// </summary>
    /// <param name="samples">
    /// [N × K] matrix of independently sampled values.
    /// Each column is one input's samples. Modified in place.
    /// </param>
    /// <param name="targetCorrelation">[K × K] target Spearman rank correlation matrix.</param>
    /// <returns>The rearranged sample matrix (same object, modified in place).</returns>
    public static double[,] Apply(double[,] samples, CorrelationMatrix targetCorrelation)
    {
        ArgumentNullException.ThrowIfNull(samples);
        ArgumentNullException.ThrowIfNull(targetCorrelation);

        int n = samples.GetLength(0);  // iterations
        int k = samples.GetLength(1);  // inputs

        if (k != targetCorrelation.Size)
            throw new ArgumentException(
                $"Sample matrix has {k} columns but correlation matrix is {targetCorrelation.Size}×{targetCorrelation.Size}.");

        if (n < 3)
            throw new ArgumentException("Need at least 3 samples for Iman-Conover.", nameof(samples));

        // Step 1: Rank-transform each column (average ranks for ties)
        var ranks = new double[n, k];
        for (int j = 0; j < k; j++)
        {
            var column = GetColumn(samples, j, n);
            var columnRanks = ComputeRanks(column);
            for (int i = 0; i < n; i++)
                ranks[i, j] = columnRanks[i];
        }

        // Step 2: Van der Waerden scores: S[i,j] = Φ⁻¹(R[i,j] / (N+1))
        var scores = new double[n, k];
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < k; j++)
            {
                scores[i, j] = Normal.InvCDF(0, 1, ranks[i, j] / (n + 1));
            }
        }

        // Step 3: Compute current Pearson correlation of scores
        var S = DenseMatrix.OfArray(scores);
        var T = ComputeCorrelationMatrix(S);

        // Step 4: Cholesky decomposition of target and current correlations
        var C = DenseMatrix.OfArray(targetCorrelation.ToArray());
        var P = C.Cholesky().Factor;  // Lower triangular from target
        var Q = T.Cholesky().Factor;  // Lower triangular from current

        // Step 5: Transformation matrix M = P * Q⁻¹
        var M = P * Q.Inverse();

        // Step 6: Apply transformation to scores: S* = S * M^T
        var transformedScores = S * M.Transpose();

        // Step 7: Rank-transform S* to get target ranks
        var targetRanks = new int[n, k];
        for (int j = 0; j < k; j++)
        {
            var column = new double[n];
            for (int i = 0; i < n; i++)
                column[i] = transformedScores[i, j];

            var sortedIndices = Enumerable.Range(0, n)
                .OrderBy(i => column[i])
                .ToArray();

            for (int rank = 0; rank < n; rank++)
                targetRanks[sortedIndices[rank], j] = rank;
        }

        // Step 8: Rearrange original samples according to target ranks
        for (int j = 0; j < k; j++)
        {
            // Get original column values and sort them
            var originalValues = GetColumn(samples, j, n);
            Array.Sort(originalValues);

            // Assign the k-th smallest value to the row that has rank k
            for (int i = 0; i < n; i++)
            {
                samples[i, j] = originalValues[targetRanks[i, j]];
            }
        }

        return samples;
    }

    private static double[] GetColumn(double[,] matrix, int col, int rows)
    {
        var column = new double[rows];
        for (int i = 0; i < rows; i++)
            column[i] = matrix[i, col];
        return column;
    }

    /// <summary>
    /// Computes average ranks for an array of values.
    /// </summary>
    private static double[] ComputeRanks(double[] values)
    {
        int n = values.Length;
        var ranks = new double[n];

        // Sort indices by value
        var indices = Enumerable.Range(0, n)
            .OrderBy(i => values[i])
            .ToArray();

        // Assign ranks with averaging for ties
        int pos = 0;
        while (pos < n)
        {
            int start = pos;
            // Find all tied values
            while (pos < n - 1 && Math.Abs(values[indices[pos + 1]] - values[indices[start]]) < 1e-15)
                pos++;

            // Average rank for the tied group (1-based ranks)
            double avgRank = (start + pos) / 2.0 + 1.0;
            for (int i = start; i <= pos; i++)
                ranks[indices[i]] = avgRank;

            pos++;
        }

        return ranks;
    }

    /// <summary>
    /// Computes the Pearson correlation matrix of the columns of a matrix.
    /// </summary>
    private static Matrix<double> ComputeCorrelationMatrix(Matrix<double> matrix)
    {
        int k = matrix.ColumnCount;
        int n = matrix.RowCount;

        // Compute column means
        var means = new double[k];
        for (int j = 0; j < k; j++)
            means[j] = matrix.Column(j).Average();

        // Center the matrix
        var centered = matrix.Clone();
        for (int j = 0; j < k; j++)
        {
            for (int i = 0; i < n; i++)
                centered[i, j] -= means[j];
        }

        // Compute covariance matrix: (1/(n-1)) * C^T * C
        var cov = (centered.Transpose() * centered) / (n - 1);

        // Convert to correlation: corr[i,j] = cov[i,j] / sqrt(cov[i,i] * cov[j,j])
        var corr = DenseMatrix.Create(k, k, 0.0);
        for (int i = 0; i < k; i++)
        {
            for (int j = 0; j < k; j++)
            {
                double denom = Math.Sqrt(cov[i, i] * cov[j, j]);
                corr[i, j] = denom > 1e-15 ? cov[i, j] / denom : 0.0;
            }
        }

        return corr;
    }
}
