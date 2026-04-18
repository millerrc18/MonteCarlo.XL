namespace MonteCarlo.Engine.Sampling;

/// <summary>
/// Latin Hypercube Sampling (LHS) generates stratified random samples on [0,1]
/// with exactly one sample per stratum for each dimension. This provides much
/// better coverage of the input space compared to simple Monte Carlo sampling
/// with the same iteration count.
/// </summary>
public sealed class LatinHypercubeSampler
{
    private readonly Random _rng;

    /// <summary>
    /// Creates a new Latin Hypercube sampler.
    /// </summary>
    /// <param name="seed">Optional RNG seed for reproducibility. Null means non-deterministic.</param>
    public LatinHypercubeSampler(int? seed = null)
    {
        _rng = seed.HasValue ? new Random(seed.Value) : new Random();
    }

    /// <summary>
    /// Generates stratified random samples on [0,1] using Latin Hypercube Sampling.
    /// </summary>
    /// <param name="iterations">Number of samples (rows). Must be positive.</param>
    /// <param name="dimensions">Number of dimensions (columns). Must be positive.</param>
    /// <returns>
    /// A [iterations x dimensions] array where each column has exactly one sample
    /// per stratum [i/n, (i+1)/n] for i in {0, ..., n-1}.
    /// </returns>
    public double[,] Generate(int iterations, int dimensions)
    {
        if (iterations <= 0)
            throw new ArgumentOutOfRangeException(nameof(iterations), iterations, "Iterations must be positive.");
        if (dimensions <= 0)
            throw new ArgumentOutOfRangeException(nameof(dimensions), dimensions, "Dimensions must be positive.");

        var result = new double[iterations, dimensions];

        for (int d = 0; d < dimensions; d++)
        {
            // Create a random permutation of {0, 1, ..., iterations-1} using Fisher-Yates
            var permutation = new int[iterations];
            for (int i = 0; i < iterations; i++)
                permutation[i] = i;

            FisherYatesShuffle(permutation);

            // For each row, sample uniformly within stratum [perm[i]/n, (perm[i]+1)/n]
            for (int i = 0; i < iterations; i++)
            {
                double lower = (double)permutation[i] / iterations;
                double upper = (double)(permutation[i] + 1) / iterations;
                result[i, d] = lower + _rng.NextDouble() * (upper - lower);
            }
        }

        return result;
    }

    /// <summary>
    /// Fisher-Yates (Knuth) shuffle — unbiased O(n) in-place permutation.
    /// </summary>
    private void FisherYatesShuffle(int[] array)
    {
        for (int i = array.Length - 1; i > 0; i--)
        {
            int j = _rng.Next(i + 1);
            (array[i], array[j]) = (array[j], array[i]);
        }
    }
}
