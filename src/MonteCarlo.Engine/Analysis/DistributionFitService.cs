using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.Engine.Analysis;

/// <summary>
/// Fits supported distribution candidates to historical numeric samples.
/// </summary>
public static class DistributionFitService
{
    public static IReadOnlyList<DistributionFitResult> Fit(double[] samples, int maxResults = 5)
    {
        if (samples == null)
            throw new ArgumentNullException(nameof(samples));

        var values = samples.Where(double.IsFinite).OrderBy(v => v).ToArray();
        if (values.Length < 3)
            throw new ArgumentException("At least three finite numeric samples are required.", nameof(samples));

        var mean = values.Average();
        var variance = SampleVariance(values, mean);
        var stdDev = Math.Sqrt(Math.Max(variance, 0));
        var min = values[0];
        var max = values[^1];
        var median = Percentile(values, 0.5);
        var looksLikeCountData = LooksLikeDiscreteCountData(values, min, max);
        var candidates = new List<DistributionFitResult>();

        AddCandidate(candidates, values, "Normal", new Dictionary<string, double>
        {
            ["mean"] = mean,
            ["stdDev"] = Math.Max(stdDev, 1e-9)
        });

        if (min < max)
        {
            AddCandidate(candidates, values, "Uniform", new Dictionary<string, double>
            {
                ["min"] = min,
                ["max"] = max
            });

            var mode = Math.Clamp(median, min + (max - min) * 1e-6, max - (max - min) * 1e-6);
            AddCandidate(candidates, values, "Triangular", new Dictionary<string, double>
            {
                ["min"] = min,
                ["mode"] = mode,
                ["max"] = max
            });
            AddCandidate(candidates, values, "PERT", new Dictionary<string, double>
            {
                ["min"] = min,
                ["mode"] = mode,
                ["max"] = max
            });
        }

        if (min > 0)
        {
            var logs = values.Select(value => Math.Log(value)).ToArray();
            var logMean = logs.Average();
            var logStdDev = Math.Sqrt(Math.Max(SampleVariance(logs, logMean), 1e-12));
            AddCandidate(candidates, values, "Lognormal", new Dictionary<string, double>
            {
                ["mu"] = logMean,
                ["sigma"] = logStdDev
            });

            AddCandidate(candidates, values, "Exponential", new Dictionary<string, double>
            {
                ["rate"] = 1.0 / Math.Max(mean, 1e-9)
            });

            if (variance > 0)
            {
                AddCandidate(candidates, values, "Gamma", new Dictionary<string, double>
                {
                    ["shape"] = Math.Max(mean * mean / variance, 1e-9),
                    ["rate"] = Math.Max(mean / variance, 1e-9)
                });
            }

            if (stdDev > 0)
            {
                AddCandidate(candidates, values, "Weibull", EstimateWeibull(mean, stdDev));
            }
        }

        if (values.All(v => v >= 0 && v <= 1) && variance > 0 && variance < mean * (1 - mean))
        {
            var common = mean * (1 - mean) / variance - 1;
            AddCandidate(candidates, values, "Beta", new Dictionary<string, double>
            {
                ["alpha"] = Math.Max(mean * common, 1e-9),
                ["beta"] = Math.Max((1 - mean) * common, 1e-9)
            });
        }

        if (values.All(IsWholeNumber) && min >= 0)
        {
            AddCandidate(candidates, values, "Poisson", new Dictionary<string, double>
            {
                ["lambda"] = Math.Max(mean, 1e-9)
            });

            var n = Math.Max(1, (int)Math.Round(max));
            if (mean <= n)
            {
                AddCandidate(candidates, values, "Binomial", new Dictionary<string, double>
                {
                    ["n"] = n,
                    ["p"] = Math.Clamp(mean / n, 1e-9, 1 - 1e-9)
                });
            }
        }

        if (values.All(IsWholeNumber) && min >= 1)
        {
            AddCandidate(candidates, values, "Geometric", new Dictionary<string, double>
            {
                ["p"] = Math.Clamp(1.0 / Math.Max(mean, 1e-9), 1e-9, 1)
            });
        }

        if (stdDev > 0)
        {
            AddCandidate(candidates, values, "Logistic", new Dictionary<string, double>
            {
                ["mu"] = mean,
                ["s"] = Math.Max(stdDev * Math.Sqrt(3) / Math.PI, 1e-9)
            });
        }

        return candidates
            .OrderBy(c => looksLikeCountData ? CountDataRank(c.DistributionName) : 0)
            .ThenBy(c => c.Score)
            .ThenBy(c => c.DistributionName, StringComparer.OrdinalIgnoreCase)
            .Take(Math.Max(maxResults, 1))
            .ToArray();
    }

    private static void AddCandidate(
        List<DistributionFitResult> candidates,
        double[] sortedValues,
        string distributionName,
        Dictionary<string, double> parameters)
    {
        try
        {
            var distribution = DistributionFactory.Create(distributionName, parameters);
            var score = KolmogorovSmirnovDistance(sortedValues, distribution);
            if (double.IsFinite(score))
                candidates.Add(new DistributionFitResult(distributionName, parameters, score, distribution.ParameterSummary()));
        }
        catch
        {
            // Some moment estimates are invalid for edge-case samples; skip those candidates.
        }
    }

    private static double KolmogorovSmirnovDistance(double[] sortedValues, IDistribution distribution)
    {
        var n = sortedValues.Length;
        var maxDistance = 0.0;
        for (var i = 0; i < n; i++)
        {
            var cdf = Math.Clamp(distribution.CDF(sortedValues[i]), 0, 1);
            var lower = i / (double)n;
            var upper = (i + 1) / (double)n;
            maxDistance = Math.Max(maxDistance, Math.Max(Math.Abs(cdf - lower), Math.Abs(upper - cdf)));
        }

        return maxDistance;
    }

    private static Dictionary<string, double> EstimateWeibull(double mean, double stdDev)
    {
        var coefficientOfVariation = stdDev / Math.Max(mean, 1e-9);
        var shape = Math.Clamp(Math.Pow(coefficientOfVariation, -1.086), 0.1, 10);
        var scale = mean / GammaApproximation(1 + 1 / shape);

        return new Dictionary<string, double>
        {
            ["shape"] = shape,
            ["scale"] = Math.Max(scale, 1e-9)
        };
    }

    private static double SampleVariance(double[] values, double mean)
    {
        if (values.Length < 2)
            return 0;

        var sum = 0.0;
        foreach (var value in values)
        {
            var delta = value - mean;
            sum += delta * delta;
        }

        return sum / (values.Length - 1);
    }

    private static double Percentile(double[] sortedValues, double p)
    {
        var position = p * (sortedValues.Length - 1);
        var lower = (int)Math.Floor(position);
        var upper = (int)Math.Ceiling(position);
        if (lower == upper)
            return sortedValues[lower];

        var weight = position - lower;
        return sortedValues[lower] * (1 - weight) + sortedValues[upper] * weight;
    }

    private static bool IsWholeNumber(double value) => Math.Abs(value - Math.Round(value)) < 1e-9;

    private static bool LooksLikeDiscreteCountData(double[] sortedValues, double min, double max)
    {
        if (min < 0 || !sortedValues.All(IsWholeNumber))
            return false;

        var distinctRatio = sortedValues.Distinct().Count() / (double)sortedValues.Length;
        return max <= 100 || distinctRatio <= 0.75;
    }

    private static int CountDataRank(string distributionName)
    {
        return distributionName is "Poisson" or "Binomial" or "Geometric" ? 0 : 1;
    }

    private static double GammaApproximation(double z)
    {
        // Lanczos approximation, good enough for Weibull moment estimates.
        double[] coefficients =
        {
            676.5203681218851,
            -1259.1392167224028,
            771.32342877765313,
            -176.61502916214059,
            12.507343278686905,
            -0.13857109526572012,
            9.9843695780195716e-6,
            1.5056327351493116e-7
        };

        if (z < 0.5)
            return Math.PI / (Math.Sin(Math.PI * z) * GammaApproximation(1 - z));

        z -= 1;
        var x = 0.99999999999980993;
        for (var i = 0; i < coefficients.Length; i++)
            x += coefficients[i] / (z + i + 1);

        var t = z + coefficients.Length - 0.5;
        return Math.Sqrt(2 * Math.PI) * Math.Pow(t, z + 0.5) * Math.Exp(-t) * x;
    }
}

public sealed record DistributionFitResult(
    string DistributionName,
    IReadOnlyDictionary<string, double> Parameters,
    double Score,
    string ParameterSummary);
