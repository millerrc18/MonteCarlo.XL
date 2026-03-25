using FluentAssertions;
using MonteCarlo.Engine.Correlation;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;
using Xunit;

namespace MonteCarlo.Engine.Tests.Correlation;

public class ImanConoverTests
{
    private const int Seed = 42;

    // --- Marginal distributions preserved ---

    [Fact]
    public void MarginalDistributions_ArePreserved()
    {
        int n = 10_000;
        int k = 2;
        var samples = GenerateIndependentSamples(n, k, Seed);

        // Save sorted column values before
        var beforeCol0 = GetColumnSorted(samples, 0, n);
        var beforeCol1 = GetColumnSorted(samples, 1, n);

        var target = new CorrelationMatrix(new double[,]
        {
            { 1.0, 0.8 },
            { 0.8, 1.0 }
        });

        ImanConover.Apply(samples, target);

        // Sort columns after — should have identical values
        var afterCol0 = GetColumnSorted(samples, 0, n);
        var afterCol1 = GetColumnSorted(samples, 1, n);

        for (int i = 0; i < n; i++)
        {
            afterCol0[i].Should().BeApproximately(beforeCol0[i], 1e-12);
            afterCol1[i].Should().BeApproximately(beforeCol1[i], 1e-12);
        }
    }

    // --- Positive correlation ---

    [Fact]
    public void PositiveCorrelation_IsAchieved()
    {
        int n = 100_000;
        var samples = GenerateIndependentSamples(n, 2, Seed);

        var target = new CorrelationMatrix(new double[,]
        {
            { 1.0, 0.8 },
            { 0.8, 1.0 }
        });

        ImanConover.Apply(samples, target);

        double rho = SpearmanCorrelation(samples, 0, 1, n);
        rho.Should().BeApproximately(0.8, 0.02);
    }

    // --- Negative correlation ---

    [Fact]
    public void NegativeCorrelation_IsAchieved()
    {
        int n = 100_000;
        var samples = GenerateIndependentSamples(n, 2, Seed);

        var target = new CorrelationMatrix(new double[,]
        {
            { 1.0, -0.7 },
            { -0.7, 1.0 }
        });

        ImanConover.Apply(samples, target);

        double rho = SpearmanCorrelation(samples, 0, 1, n);
        rho.Should().BeApproximately(-0.7, 0.02);
    }

    // --- Identity correlation changes nothing meaningful ---

    [Fact]
    public void IdentityCorrelation_PreservesNearZeroCorrelation()
    {
        int n = 10_000;
        var samples = GenerateIndependentSamples(n, 2, Seed);
        var target = CorrelationMatrix.Identity(2);

        ImanConover.Apply(samples, target);

        double rho = SpearmanCorrelation(samples, 0, 1, n);
        rho.Should().BeApproximately(0.0, 0.05);
    }

    // --- 3×3 correlation ---

    [Fact]
    public void ThreeByThree_AllPairsMatchTargets()
    {
        int n = 50_000;
        var samples = GenerateIndependentSamples(n, 3, Seed);

        var target = new CorrelationMatrix(new double[,]
        {
            { 1.0,  0.6, -0.3 },
            { 0.6,  1.0,  0.4 },
            { -0.3, 0.4,  1.0 }
        });

        ImanConover.Apply(samples, target);

        SpearmanCorrelation(samples, 0, 1, n).Should().BeApproximately(0.6, 0.03);
        SpearmanCorrelation(samples, 0, 2, n).Should().BeApproximately(-0.3, 0.03);
        SpearmanCorrelation(samples, 1, 2, n).Should().BeApproximately(0.4, 0.03);
    }

    // --- Large N convergence ---

    [Fact]
    public void LargeN_AchievesHighPrecision()
    {
        int n = 50_000;
        var samples = GenerateIndependentSamples(n, 2, Seed);

        var target = new CorrelationMatrix(new double[,]
        {
            { 1.0, 0.5 },
            { 0.5, 1.0 }
        });

        ImanConover.Apply(samples, target);

        double rho = SpearmanCorrelation(samples, 0, 1, n);
        rho.Should().BeApproximately(0.5, 0.02);
    }

    // --- Small N still works ---

    [Fact]
    public void SmallN_StillProducesReasonableCorrelation()
    {
        int n = 100;
        var samples = GenerateIndependentSamples(n, 2, Seed);

        var target = new CorrelationMatrix(new double[,]
        {
            { 1.0, 0.8 },
            { 0.8, 1.0 }
        });

        ImanConover.Apply(samples, target);

        double rho = SpearmanCorrelation(samples, 0, 1, n);
        rho.Should().BeApproximately(0.8, 0.05);
    }

    // --- Integration test with SimulationEngine ---

    [Fact]
    public async Task EndToEnd_WithEngine_CorrelationApplied()
    {
        var engine = new SimulationEngine();
        var config = new SimulationConfig
        {
            IterationCount = 10_000,
            RandomSeed = 42,
            ParallelExecution = false,
            Inputs = new List<SimulationInput>
            {
                new() { Id = "A", Label = "Input A", Distribution = new NormalDistribution(100, 10, 42) },
                new() { Id = "B", Label = "Input B", Distribution = new NormalDistribution(200, 20, 43) }
            },
            Outputs = new List<SimulationOutput>
            {
                new() { Id = "Out", Label = "Output" }
            },
            Correlation = new CorrelationMatrix(new double[,]
            {
                { 1.0, 0.9 },
                { 0.9, 1.0 }
            })
        };

        var result = await engine.RunAsync(config, inputs =>
        {
            return new Dictionary<string, double>
            {
                { "Out", inputs["A"] + inputs["B"] }
            };
        });

        // Compute Spearman correlation of the two input sample columns
        double rho = SpearmanCorrelationFromResult(result, 0, 1);
        rho.Should().BeApproximately(0.9, 0.02);
    }

    // --- Mismatched dimensions throws ---

    [Fact]
    public void MismatchedDimensions_Throws()
    {
        var samples = new double[100, 3];
        var target = CorrelationMatrix.Identity(2);

        var act = () => ImanConover.Apply(samples, target);
        act.Should().Throw<ArgumentException>().WithMessage("*columns*");
    }

    // --- Too few samples throws ---

    [Fact]
    public void TooFewSamples_Throws()
    {
        var samples = new double[2, 2];
        var target = CorrelationMatrix.Identity(2);

        var act = () => ImanConover.Apply(samples, target);
        act.Should().Throw<ArgumentException>().WithMessage("*3 samples*");
    }

    // --- Helpers ---

    private static double[,] GenerateIndependentSamples(int n, int k, int seed)
    {
        var samples = new double[n, k];
        for (int j = 0; j < k; j++)
        {
            var dist = new NormalDistribution(100 * (j + 1), 10 * (j + 1), seed + j);
            var s = dist.Sample(n);
            for (int i = 0; i < n; i++)
                samples[i, j] = s[i];
        }
        return samples;
    }

    private static double[] GetColumnSorted(double[,] matrix, int col, int n)
    {
        var column = new double[n];
        for (int i = 0; i < n; i++)
            column[i] = matrix[i, col];
        Array.Sort(column);
        return column;
    }

    private static double SpearmanCorrelation(double[,] samples, int col1, int col2, int n)
    {
        var x = new double[n];
        var y = new double[n];
        for (int i = 0; i < n; i++)
        {
            x[i] = samples[i, col1];
            y[i] = samples[i, col2];
        }

        var xRanks = ComputeRanks(x);
        var yRanks = ComputeRanks(y);

        return PearsonCorrelation(xRanks, yRanks);
    }

    private static double SpearmanCorrelationFromResult(SimulationResult result, int col1, int col2)
    {
        int n = result.IterationCount;
        var x = new double[n];
        var y = new double[n];
        for (int i = 0; i < n; i++)
        {
            x[i] = result.InputMatrix[i, col1];
            y[i] = result.InputMatrix[i, col2];
        }

        var xRanks = ComputeRanks(x);
        var yRanks = ComputeRanks(y);

        return PearsonCorrelation(xRanks, yRanks);
    }

    private static double[] ComputeRanks(double[] values)
    {
        int n = values.Length;
        var ranks = new double[n];
        var indices = Enumerable.Range(0, n).OrderBy(i => values[i]).ToArray();

        int pos = 0;
        while (pos < n)
        {
            int start = pos;
            while (pos < n - 1 && Math.Abs(values[indices[pos + 1]] - values[indices[start]]) < 1e-15)
                pos++;

            double avgRank = (start + pos) / 2.0 + 1.0;
            for (int i = start; i <= pos; i++)
                ranks[indices[i]] = avgRank;

            pos++;
        }

        return ranks;
    }

    private static double PearsonCorrelation(double[] x, double[] y)
    {
        int n = x.Length;
        double meanX = x.Average();
        double meanY = y.Average();
        double sumXY = 0, sumX2 = 0, sumY2 = 0;
        for (int i = 0; i < n; i++)
        {
            double dx = x[i] - meanX;
            double dy = y[i] - meanY;
            sumXY += dx * dy;
            sumX2 += dx * dx;
            sumY2 += dy * dy;
        }
        return sumXY / Math.Sqrt(sumX2 * sumY2);
    }
}
