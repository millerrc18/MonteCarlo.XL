using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace MonteCarlo.Engine.Correlation;

/// <summary>
/// Represents a symmetric correlation matrix with validation and
/// nearest positive semi-definite correction.
/// </summary>
public class CorrelationMatrix
{
    private readonly double[,] _data;

    /// <summary>
    /// Gets the size (K) of the K×K matrix.
    /// </summary>
    public int Size { get; }

    /// <summary>
    /// Creates a correlation matrix from a 2D array.
    /// </summary>
    /// <param name="matrix">A square 2D array. Must be symmetric with 1.0 diagonal.</param>
    public CorrelationMatrix(double[,] matrix)
    {
        ArgumentNullException.ThrowIfNull(matrix);

        int rows = matrix.GetLength(0);
        int cols = matrix.GetLength(1);
        if (rows != cols)
            throw new ArgumentException($"Matrix must be square. Got {rows}×{cols}.");
        if (rows < 2)
            throw new ArgumentException("Matrix must be at least 2×2.");

        Size = rows;
        _data = (double[,])matrix.Clone();
    }

    /// <summary>
    /// Access elements of the correlation matrix.
    /// </summary>
    public double this[int i, int j]
    {
        get => _data[i, j];
        set => _data[i, j] = value;
    }

    /// <summary>
    /// Creates a K×K identity matrix (no correlation — all inputs independent).
    /// </summary>
    public static CorrelationMatrix Identity(int size)
    {
        if (size < 2)
            throw new ArgumentException("Size must be at least 2.", nameof(size));

        var data = new double[size, size];
        for (int i = 0; i < size; i++)
            data[i, i] = 1.0;

        return new CorrelationMatrix(data);
    }

    /// <summary>
    /// Validates that the matrix is a proper correlation matrix:
    /// square, symmetric, diagonal = 1.0, values in [-1, 1], and positive semi-definite.
    /// </summary>
    /// <exception cref="ArgumentException">Thrown with specific message on validation failure.</exception>
    public void Validate()
    {
        // Check diagonal is 1.0
        for (int i = 0; i < Size; i++)
        {
            if (Math.Abs(_data[i, i] - 1.0) > 1e-10)
                throw new ArgumentException($"Diagonal element [{i},{i}] must be 1.0, got {_data[i, i]}.");
        }

        // Check symmetry and bounds
        for (int i = 0; i < Size; i++)
        {
            for (int j = i + 1; j < Size; j++)
            {
                if (Math.Abs(_data[i, j] - _data[j, i]) > 1e-10)
                    throw new ArgumentException(
                        $"Matrix is not symmetric: [{i},{j}]={_data[i, j]} ≠ [{j},{i}]={_data[j, i]}.");

                if (_data[i, j] < -1.0 || _data[i, j] > 1.0)
                    throw new ArgumentException(
                        $"Value at [{i},{j}]={_data[i, j]} is outside [-1, 1].");
            }
        }

        // Check positive semi-definite (all eigenvalues >= 0)
        var m = ToMathNetMatrix();
        var evd = m.Evd();
        var eigenvalues = evd.EigenValues;
        for (int i = 0; i < Size; i++)
        {
            if (eigenvalues[i].Real < -1e-10)
                throw new ArgumentException(
                    $"Matrix is not positive semi-definite. Eigenvalue {i} = {eigenvalues[i].Real:G6}.");
        }
    }

    /// <summary>
    /// Returns true if the matrix is positive semi-definite.
    /// </summary>
    public bool IsPositiveSemiDefinite()
    {
        var m = ToMathNetMatrix();
        var evd = m.Evd();
        var eigenvalues = evd.EigenValues;
        for (int i = 0; i < Size; i++)
        {
            if (eigenvalues[i].Real < -1e-10)
                return false;
        }
        return true;
    }

    /// <summary>
    /// Returns the smallest eigenvalue of the matrix.
    /// </summary>
    public double MinimumEigenvalue()
    {
        var m = ToMathNetMatrix();
        var evd = m.Evd();
        return evd.EigenValues.Min(value => value.Real);
    }

    /// <summary>
    /// If the matrix is not positive semi-definite, computes and returns the
    /// nearest PSD correlation matrix. If already PSD, returns this.
    /// </summary>
    public CorrelationMatrix EnsurePositiveSemiDefinite()
    {
        if (IsPositiveSemiDefinite())
            return this;

        // Eigendecomposition approach:
        // 1. C = V * D * V^T
        // 2. Set negative eigenvalues to a small positive value
        // 3. Reconstruct: C' = V * D' * V^T
        // 4. Rescale diagonal back to 1.0
        var m = ToMathNetMatrix();
        var evd = m.Evd();
        var eigenvalues = evd.EigenValues;
        var V = evd.EigenVectors;

        // Build corrected diagonal matrix
        var dPrime = DenseMatrix.Create(Size, Size, 0.0);
        for (int i = 0; i < Size; i++)
        {
            double val = eigenvalues[i].Real;
            dPrime[i, i] = val < 1e-10 ? 1e-10 : val;
        }

        // Reconstruct
        var corrected = V * dPrime * V.Transpose();

        // Rescale diagonal to 1.0
        var result = new double[Size, Size];
        for (int i = 0; i < Size; i++)
        {
            for (int j = 0; j < Size; j++)
            {
                double scale = Math.Sqrt(corrected[i, i] * corrected[j, j]);
                result[i, j] = corrected[i, j] / scale;
            }
        }

        // Enforce exact symmetry and diagonal
        for (int i = 0; i < Size; i++)
        {
            result[i, i] = 1.0;
            for (int j = i + 1; j < Size; j++)
            {
                double avg = (result[i, j] + result[j, i]) / 2.0;
                result[i, j] = avg;
                result[j, i] = avg;
            }
        }

        return new CorrelationMatrix(result);
    }

    /// <summary>
    /// Gets the underlying array.
    /// </summary>
    public double[,] ToArray()
    {
        return (double[,])_data.Clone();
    }

    private Matrix<double> ToMathNetMatrix()
    {
        return DenseMatrix.OfArray(_data);
    }
}
