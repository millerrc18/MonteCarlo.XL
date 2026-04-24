using MonteCarlo.Engine.Correlation;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Sampling;
using MonteCarlo.Engine.Simulation;
using MonteCarlo.Engine.Validation;
using MonteCarlo.Shared.Formula;
using MonteCarlo.Shared.Interop;

namespace MonteCarlo.Shared.Simulation;

public sealed class SimulationSamplingService
{
    public GenerateInputMatrixResponse GenerateInputMatrix(GenerateInputMatrixRequest request)
    {
        EnsureVersion(request.Version);
        var config = BuildConfig(request.Profile, request.Settings);
        var inputMatrix = GenerateInputMatrix(config);

        return new GenerateInputMatrixResponse(
            BridgeProtocol.Version,
            config.Inputs.Select(input => input.Id).ToArray(),
            config.Outputs.Select(output => output.Id).ToArray(),
            ToJagged(inputMatrix));
    }

    public ValidateProfileResponse ValidateProfile(ValidateProfileRequest request)
    {
        EnsureVersion(request.Version);
        var report = ModelPreflightValidator.Validate(request.Profile);
        return new ValidateProfileResponse(
            BridgeProtocol.Version,
            report.Issues
                .Select(issue => new PreflightIssueDto(
                    issue.Severity.ToString(),
                    issue.Code,
                    issue.Title,
                    issue.Message,
                    issue.SuggestedAction))
                .ToArray());
    }

    public SimulationConfig BuildConfig(SimulationProfile profile, OfficeRunSettings settings)
    {
        ArgumentNullException.ThrowIfNull(profile);
        ArgumentNullException.ThrowIfNull(settings);

        var config = new SimulationConfig
        {
            IterationCount = settings.IterationCount,
            RandomSeed = settings.RandomSeed,
            ParallelExecution = false,
            Sampling = settings.SamplingMethod,
            AutoStopOnConvergence = settings.AutoStopOnConvergence
        };

        for (var index = 0; index < profile.Inputs.Count; index++)
        {
            var input = profile.Inputs[index];
            var distribution = DistributionFactory.Create(
                input.DistributionName,
                new Dictionary<string, double>(input.Parameters),
                seed: settings.RandomSeed.HasValue ? settings.RandomSeed.Value + index + 1 : null);

            config.Inputs.Add(new SimulationInput
            {
                Id = ToFullReference(input.SheetName, input.CellAddress),
                Label = string.IsNullOrWhiteSpace(input.Label) ? input.CellAddress : input.Label,
                Distribution = distribution,
                BaseValue = McNormalModeBaseValue(input.DistributionName, input.Parameters)
            });
        }

        foreach (var output in profile.Outputs)
        {
            config.Outputs.Add(new SimulationOutput
            {
                Id = ToFullReference(output.SheetName, output.CellAddress),
                Label = string.IsNullOrWhiteSpace(output.Label) ? output.CellAddress : output.Label
            });
        }

        if (profile.CorrelationMatrix is { } matrix)
        {
            var correlation = new CorrelationMatrix(matrix);
            correlation.Validate();
            config.Correlation = correlation;
        }

        return config;
    }

    public static double[,] ToRectangular(double[][] source)
    {
        if (source.Length == 0)
            return new double[0, 0];

        var columns = source[0].Length;
        var matrix = new double[source.Length, columns];
        for (var row = 0; row < source.Length; row++)
        {
            if (source[row].Length != columns)
                throw new ArgumentException("All rows must have the same length.", nameof(source));

            for (var column = 0; column < columns; column++)
                matrix[row, column] = source[row][column];
        }

        return matrix;
    }

    public static double[][] ToJagged(double[,] source)
    {
        var rows = source.GetLength(0);
        var columns = source.GetLength(1);
        var jagged = new double[rows][];

        for (var row = 0; row < rows; row++)
        {
            jagged[row] = new double[columns];
            for (var column = 0; column < columns; column++)
                jagged[row][column] = source[row, column];
        }

        return jagged;
    }

    private static void EnsureVersion(int version)
    {
        if (version != BridgeProtocol.Version)
            throw new InvalidOperationException($"Unsupported bridge version {version}. Expected {BridgeProtocol.Version}.");
    }

    private static string ToFullReference(string sheetName, string cellAddress) =>
        $"'{sheetName}'!{cellAddress}";

    private static double McNormalModeBaseValue(string distributionName, IReadOnlyDictionary<string, double> parameters)
    {
        if (!McFormulaCatalog.TryGetByFunctionName(distributionName, out var definition))
            return 0;

        var orderedArguments = new List<double>(definition.Arguments.Count);
        foreach (var argument in definition.Arguments)
        {
            if (!parameters.TryGetValue(argument.Name, out var value))
                return 0;
            orderedArguments.Add(value);
        }

        return McNormalModeEvaluator.TryEvaluate(definition.FunctionName, orderedArguments, out var result)
            ? result
            : 0;
    }

    private static double[,] GenerateInputMatrix(SimulationConfig config)
    {
        config.Validate();

        var iterations = config.IterationCount;
        var inputCount = config.Inputs.Count;
        var inputMatrix = new double[iterations, inputCount];

        if (config.Sampling == SamplingMethod.LatinHypercube)
        {
            var lhs = new LatinHypercubeSampler(config.RandomSeed);
            var unitSamples = lhs.Generate(iterations, inputCount);

            for (var column = 0; column < inputCount; column++)
            {
                var distribution = config.Inputs[column].Distribution;
                for (var row = 0; row < iterations; row++)
                    inputMatrix[row, column] = distribution.Percentile(unitSamples[row, column]);
            }
        }
        else
        {
            for (var column = 0; column < inputCount; column++)
            {
                var samples = config.Inputs[column].Distribution.Sample(iterations);
                for (var row = 0; row < iterations; row++)
                    inputMatrix[row, column] = samples[row];
            }
        }

        if (config.Correlation != null)
            ImanConover.Apply(inputMatrix, config.Correlation);

        return inputMatrix;
    }
}
