using MonteCarlo.Engine.Simulation;

namespace MonteCarlo.Addin.Services;

internal enum StressRuleMode
{
    FixedValue,
    AddShift,
    RangeScale
}

internal sealed record StressInputRule(
    string InputId,
    string InputLabel,
    StressRuleMode Mode,
    double Value)
{
    public double Apply(SimulationInput input, double sample) =>
        Mode switch
        {
            StressRuleMode.FixedValue => Value,
            StressRuleMode.AddShift => sample + Value,
            StressRuleMode.RangeScale => input.BaseValue + ((sample - input.BaseValue) * Value),
            _ => sample
        };

    public string Describe() =>
        Mode switch
        {
            StressRuleMode.FixedValue => $"{InputLabel}: fixed at {Value:G6}",
            StressRuleMode.AddShift => $"{InputLabel}: shift by {Value:G6}",
            StressRuleMode.RangeScale => $"{InputLabel}: scale range around base by {Value:G6}x",
            _ => $"{InputLabel}: unchanged"
        };
}

internal sealed class StressInputPlan
{
    private readonly IReadOnlyDictionary<string, StressInputRule> _rulesByInputId;

    public StressInputPlan(IReadOnlyList<StressInputRule> rules)
    {
        ArgumentNullException.ThrowIfNull(rules);
        Rules = rules;
        _rulesByInputId = rules.ToDictionary(rule => rule.InputId, StringComparer.OrdinalIgnoreCase);
    }

    public IReadOnlyList<StressInputRule> Rules { get; }

    public int Count => Rules.Count;

    public double Apply(SimulationInput input, double sample) =>
        _rulesByInputId.TryGetValue(input.Id, out var rule)
            ? rule.Apply(input, sample)
            : sample;
}

internal sealed record StressRunOptions(
    string PrimaryOutputId,
    string PrimaryOutputLabel,
    int IterationsPerRun,
    StressInputPlan Plan,
    bool ExportRawData);
