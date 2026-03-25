using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Computes sensitivity of simulation outputs to inputs using Spearman rank correlation.
/// </summary>
public static class SensitivityAnalysis
{
    /// <summary>
    /// Analyze how sensitive a specific output is to each input.
    /// Results are sorted by absolute rank correlation (most impactful first).
    /// </summary>
    /// <param name="result">The completed simulation result.</param>
    /// <param name="outputIndex">Index of the output to analyze.</param>
    /// <returns>Sensitivity results sorted by absolute impact (descending).</returns>
    public static List<SensitivityResult> Analyze(SimulationResult result, int outputIndex)
    {
        if (result == null)
            throw new ArgumentNullException(nameof(result));
        if (outputIndex < 0 || outputIndex >= result.Config.Outputs.Count)
            throw new ArgumentOutOfRangeException(nameof(outputIndex));

        int n = result.IterationCount;
        int inputCount = result.Config.Inputs.Count;

        // Extract output column
        var outputValues = new double[n];
        for (int i = 0; i < n; i++)
            outputValues[i] = result.OutputMatrix[i, outputIndex];

        var outputRanks = ComputeRanks(outputValues);

        var results = new List<SensitivityResult>(inputCount);
        var squaredCorrelations = new double[inputCount];

        for (int j = 0; j < inputCount; j++)
        {
            var input = result.Config.Inputs[j];

            // Extract input column
            var inputValues = new double[n];
            for (int i = 0; i < n; i++)
                inputValues[i] = result.InputMatrix[i, j];

            var inputRanks = ComputeRanks(inputValues);

            // Spearman rank correlation = Pearson correlation of ranks
            double rankCorrelation = PearsonCorrelation(inputRanks, outputRanks);

            squaredCorrelations[j] = rankCorrelation * rankCorrelation;

            // Output at input extremes: find iterations where input is in bottom/top 10%
            var (outputAtP10, outputAtP90) = ComputeOutputAtExtremes(inputValues, outputValues, n);

            results.Add(new SensitivityResult
            {
                InputId = input.Id,
                InputLabel = input.Label,
                RankCorrelation = rankCorrelation,
                OutputAtInputP10 = outputAtP10,
                OutputAtInputP90 = outputAtP90
            });
        }

        // Compute contribution to variance (normalize squared correlations to sum to 100%)
        double totalSquared = squaredCorrelations.Sum();
        if (totalSquared > 0)
        {
            for (int j = 0; j < inputCount; j++)
            {
                results[j] = new SensitivityResult
                {
                    InputId = results[j].InputId,
                    InputLabel = results[j].InputLabel,
                    RankCorrelation = results[j].RankCorrelation,
                    ContributionToVariance = (squaredCorrelations[j] / totalSquared) * 100.0,
                    OutputAtInputP10 = results[j].OutputAtInputP10,
                    OutputAtInputP90 = results[j].OutputAtInputP90
                };
            }
        }

        // Sort by absolute rank correlation descending
        results.Sort((a, b) => Math.Abs(b.RankCorrelation).CompareTo(Math.Abs(a.RankCorrelation)));
        return results;
    }

    /// <summary>
    /// Compute tornado swing values by directly evaluating the model with each input
    /// at its P10/P90 while holding all other inputs at their base values.
    /// This is more precise than conditional means when an evaluator is available.
    /// </summary>
    /// <param name="result">The completed simulation result.</param>
    /// <param name="outputIndex">Index of the output to analyze.</param>
    /// <param name="evaluator">The model evaluator function.</param>
    /// <returns>Sensitivity results with swing values from direct evaluation, sorted by swing descending.</returns>
    public static List<SensitivityResult> ComputeTornadoSwing(
        SimulationResult result,
        int outputIndex,
        Func<Dictionary<string, double>, Dictionary<string, double>> evaluator)
    {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (evaluator == null) throw new ArgumentNullException(nameof(evaluator));
        if (outputIndex < 0 || outputIndex >= result.Config.Outputs.Count)
            throw new ArgumentOutOfRangeException(nameof(outputIndex));

        var config = result.Config;
        string outputId = config.Outputs[outputIndex].Id;

        // Build base case input dictionary (all inputs at mean/base values)
        var baseInputs = new Dictionary<string, double>();
        var inputSamples = new Dictionary<string, double[]>();
        foreach (var input in config.Inputs)
        {
            baseInputs[input.Id] = input.Distribution.Mean;
            inputSamples[input.Id] = result.GetInputSamples(input.Id);
        }

        var results = new List<SensitivityResult>(config.Inputs.Count);

        foreach (var input in config.Inputs)
        {
            var samples = inputSamples[input.Id];
            var sorted = samples.OrderBy(x => x).ToArray();
            double p10Value = sorted[(int)(sorted.Length * 0.10)];
            double p90Value = sorted[(int)(sorted.Length * 0.90)];

            // Evaluate at P10
            var p10Inputs = new Dictionary<string, double>(baseInputs) { [input.Id] = p10Value };
            var p10Outputs = evaluator(p10Inputs);
            double outputAtP10 = p10Outputs[outputId];

            // Evaluate at P90
            var p90Inputs = new Dictionary<string, double>(baseInputs) { [input.Id] = p90Value };
            var p90Outputs = evaluator(p90Inputs);
            double outputAtP90 = p90Outputs[outputId];

            // Also compute rank correlation from the simulation data
            var outputValues = new double[result.IterationCount];
            for (int i = 0; i < result.IterationCount; i++)
                outputValues[i] = result.OutputMatrix[i, outputIndex];

            int inputIdx = config.Inputs.IndexOf(input);
            var inputValues = new double[result.IterationCount];
            for (int i = 0; i < result.IterationCount; i++)
                inputValues[i] = result.InputMatrix[i, inputIdx];

            double rankCorrelation = PearsonCorrelation(ComputeRanks(inputValues), ComputeRanks(outputValues));

            results.Add(new SensitivityResult
            {
                InputId = input.Id,
                InputLabel = input.Label,
                RankCorrelation = rankCorrelation,
                OutputAtInputP10 = outputAtP10,
                OutputAtInputP90 = outputAtP90
            });
        }

        // Sort by swing descending
        results.Sort((a, b) => b.Swing.CompareTo(a.Swing));
        return results;
    }

    private static double[] ComputeRanks(double[] values)
    {
        int n = values.Length;
        var indexed = new (double Value, int Index)[n];
        for (int i = 0; i < n; i++)
            indexed[i] = (values[i], i);

        Array.Sort(indexed, (a, b) => a.Value.CompareTo(b.Value));

        var ranks = new double[n];
        int i2 = 0;
        while (i2 < n)
        {
            int j = i2;
            // Find all ties
            while (j < n && indexed[j].Value == indexed[i2].Value)
                j++;

            // Average rank for ties
            double avgRank = (i2 + j - 1) / 2.0 + 1.0; // 1-based
            for (int k = i2; k < j; k++)
                ranks[indexed[k].Index] = avgRank;

            i2 = j;
        }

        return ranks;
    }

    private static double PearsonCorrelation(double[] x, double[] y)
    {
        int n = x.Length;
        double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0, sumY2 = 0;

        for (int i = 0; i < n; i++)
        {
            sumX += x[i];
            sumY += y[i];
            sumXY += x[i] * y[i];
            sumX2 += x[i] * x[i];
            sumY2 += y[i] * y[i];
        }

        double denom = Math.Sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
        if (denom == 0) return 0;

        return (n * sumXY - sumX * sumY) / denom;
    }

    private static (double OutputAtP10, double OutputAtP90) ComputeOutputAtExtremes(
        double[] inputValues, double[] outputValues, int n)
    {
        // Sort iteration indices by input value
        var indices = Enumerable.Range(0, n).ToArray();
        Array.Sort(inputValues.ToArray(), indices);

        int bucketSize = Math.Max(1, n / 10); // Bottom/top 10%

        double sumP10 = 0;
        for (int i = 0; i < bucketSize; i++)
            sumP10 += outputValues[indices[i]];

        double sumP90 = 0;
        for (int i = n - bucketSize; i < n; i++)
            sumP90 += outputValues[indices[i]];

        return (sumP10 / bucketSize, sumP90 / bucketSize);
    }
}
