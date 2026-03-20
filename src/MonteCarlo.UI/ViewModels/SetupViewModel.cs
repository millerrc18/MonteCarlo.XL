using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Distributions;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// View model for the Setup view where users configure simulation inputs, outputs, and options.
/// </summary>
public partial class SetupViewModel : ObservableObject
{
    /// <summary>Configured simulation inputs.</summary>
    [ObservableProperty]
    private ObservableCollection<InputCardViewModel> _inputs = new();

    /// <summary>Configured simulation outputs.</summary>
    [ObservableProperty]
    private ObservableCollection<OutputCardViewModel> _outputs = new();

    /// <summary>Number of simulation iterations.</summary>
    [ObservableProperty]
    private int _iterationCount = 5000;

    /// <summary>Optional random seed for reproducibility. Null = auto.</summary>
    [ObservableProperty]
    private int? _randomSeed;

    /// <summary>Whether the seed is locked (user-specified).</summary>
    [ObservableProperty]
    private bool _isSeedLocked;

    // --- Input editor state ---

    /// <summary>Whether the input editor panel is visible.</summary>
    [ObservableProperty]
    private bool _isAddingInput;

    /// <summary>Whether the output editor panel is visible.</summary>
    [ObservableProperty]
    private bool _isAddingOutput;

    /// <summary>Cell address for the new input being added.</summary>
    [ObservableProperty]
    private string _newInputCellAddress = string.Empty;

    /// <summary>Sheet name for the new input being added.</summary>
    [ObservableProperty]
    private string _newInputSheetName = string.Empty;

    /// <summary>Label for the new input being added.</summary>
    [ObservableProperty]
    private string _newInputLabel = string.Empty;

    /// <summary>Selected distribution type for the new input.</summary>
    [ObservableProperty]
    private string _selectedDistribution = "Normal";

    /// <summary>Whether the cell selection mode is active.</summary>
    [ObservableProperty]
    private bool _isSelectingCell;

    /// <summary>Validation error message for the input editor.</summary>
    [ObservableProperty]
    private string? _inputValidationError;

    // --- Output editor state ---

    /// <summary>Cell address for the new output being added.</summary>
    [ObservableProperty]
    private string _newOutputCellAddress = string.Empty;

    /// <summary>Sheet name for the new output being added.</summary>
    [ObservableProperty]
    private string _newOutputSheetName = string.Empty;

    /// <summary>Label for the new output being added.</summary>
    [ObservableProperty]
    private string _newOutputLabel = string.Empty;

    /// <summary>Whether the output cell selection mode is active.</summary>
    [ObservableProperty]
    private bool _isSelectingOutputCell;

    // --- Distribution parameters ---

    /// <summary>Normal: Mean.</summary>
    [ObservableProperty] private string _paramMean = "0";
    /// <summary>Normal: Std Dev.</summary>
    [ObservableProperty] private string _paramStdDev = "1";
    /// <summary>Triangular/PERT/Uniform: Min.</summary>
    [ObservableProperty] private string _paramMin = "0";
    /// <summary>Triangular/PERT: Mode.</summary>
    [ObservableProperty] private string _paramMode = "50";
    /// <summary>Triangular/PERT/Uniform: Max.</summary>
    [ObservableProperty] private string _paramMax = "100";
    /// <summary>Lognormal: Mu.</summary>
    [ObservableProperty] private string _paramMu = "0";
    /// <summary>Lognormal: Sigma.</summary>
    [ObservableProperty] private string _paramSigma = "1";

    /// <summary>Discrete value/probability pairs.</summary>
    [ObservableProperty]
    private ObservableCollection<DiscretePairViewModel> _discretePairs = new()
    {
        new DiscretePairViewModel { Value = "0", Probability = "0.5" },
        new DiscretePairViewModel { Value = "1", Probability = "0.5" }
    };

    /// <summary>Preview points for the current distribution parameter editor.</summary>
    [ObservableProperty]
    private IReadOnlyList<(double X, double Y)>? _editorPreviewPoints;

    /// <summary>Available distribution names.</summary>
    public IReadOnlyList<string> AvailableDistributions { get; } = DistributionFactory.AvailableDistributions;

    /// <summary>Common iteration count presets.</summary>
    public IReadOnlyList<int> IterationPresets { get; } = new[] { 1000, 5000, 10000, 25000, 50000 };

    /// <summary>Whether the Run button should be enabled.</summary>
    public bool CanRun => Inputs.Count > 0 && Outputs.Count > 0;

    /// <summary>Event raised when the user clicks Run Simulation.</summary>
    public event EventHandler? RunSimulationRequested;

    /// <summary>
    /// Event raised when cell selection mode is requested.
    /// The Addin layer subscribes to this and hooks into Excel's SheetSelectionChange event.
    /// </summary>
    public event EventHandler<CellSelectionRequestedEventArgs>? CellSelectionRequested;

    /// <summary>
    /// Event raised when an input is added (for cell highlighting).
    /// </summary>
    public event EventHandler<InputAddedEventArgs>? InputAdded;

    /// <summary>
    /// Event raised when an output is added (for cell highlighting).
    /// </summary>
    public event EventHandler<OutputAddedEventArgs>? OutputAdded;

    /// <summary>
    /// Event raised when an input is removed.
    /// </summary>
    public event EventHandler<InputRemovedEventArgs>? InputRemoved;

    /// <summary>
    /// Event raised when an output is removed.
    /// </summary>
    public event EventHandler<OutputRemovedEventArgs>? OutputRemoved;

    public SetupViewModel()
    {
        Inputs.CollectionChanged += (_, _) => OnPropertyChanged(nameof(CanRun));
        Outputs.CollectionChanged += (_, _) => OnPropertyChanged(nameof(CanRun));
    }

    partial void OnSelectedDistributionChanged(string value)
    {
        UpdateEditorPreview();
    }

    /// <summary>
    /// Called by any parameter TextBox on change to update the preview.
    /// </summary>
    [RelayCommand]
    private void UpdateEditorPreview()
    {
        InputValidationError = null;
        try
        {
            var parameters = BuildParameterDictionary();
            var dist = DistributionFactory.Create(SelectedDistribution, parameters);
            EditorPreviewPoints = InputCardViewModel.ComputePreviewPointsStatic(dist);
        }
        catch (Exception ex)
        {
            EditorPreviewPoints = null;
            InputValidationError = ex.Message;
        }
    }

    [RelayCommand]
    private void ShowAddInput()
    {
        IsAddingInput = true;
        IsAddingOutput = false;
        ResetInputEditor();
    }

    [RelayCommand]
    private void ShowAddOutput()
    {
        IsAddingOutput = true;
        IsAddingInput = false;
        ResetOutputEditor();
    }

    [RelayCommand]
    private void CancelAddInput()
    {
        IsAddingInput = false;
        IsSelectingCell = false;
    }

    [RelayCommand]
    private void CancelAddOutput()
    {
        IsAddingOutput = false;
        IsSelectingOutputCell = false;
    }

    [RelayCommand]
    private void StartCellSelection()
    {
        IsSelectingCell = true;
        CellSelectionRequested?.Invoke(this, new CellSelectionRequestedEventArgs("input"));
    }

    [RelayCommand]
    private void StartOutputCellSelection()
    {
        IsSelectingOutputCell = true;
        CellSelectionRequested?.Invoke(this, new CellSelectionRequestedEventArgs("output"));
    }

    /// <summary>
    /// Called by the CellSelectionService when a cell is selected in Excel.
    /// </summary>
    public void OnCellSelected(string sheetName, string cellAddress, string? suggestedLabel)
    {
        if (IsSelectingCell)
        {
            NewInputSheetName = sheetName;
            NewInputCellAddress = cellAddress;
            if (!string.IsNullOrEmpty(suggestedLabel))
                NewInputLabel = suggestedLabel;
            IsSelectingCell = false;
        }
        else if (IsSelectingOutputCell)
        {
            NewOutputSheetName = sheetName;
            NewOutputCellAddress = cellAddress;
            if (!string.IsNullOrEmpty(suggestedLabel))
                NewOutputLabel = suggestedLabel;
            IsSelectingOutputCell = false;
        }
    }

    [RelayCommand]
    private void ConfirmAddInput()
    {
        InputValidationError = null;

        if (string.IsNullOrWhiteSpace(NewInputCellAddress))
        {
            InputValidationError = "Please select a cell.";
            return;
        }

        if (string.IsNullOrWhiteSpace(NewInputLabel))
        {
            InputValidationError = "Please enter a label.";
            return;
        }

        try
        {
            var parameters = BuildParameterDictionary();
            var dist = DistributionFactory.Create(SelectedDistribution, parameters);

            var card = new InputCardViewModel(
                NewInputSheetName,
                NewInputCellAddress,
                NewInputLabel,
                SelectedDistribution,
                parameters,
                dist);

            Inputs.Add(card);
            IsAddingInput = false;

            InputAdded?.Invoke(this, new InputAddedEventArgs(card));
        }
        catch (Exception ex)
        {
            InputValidationError = ex.Message;
        }
    }

    [RelayCommand]
    private void ConfirmAddOutput()
    {
        if (string.IsNullOrWhiteSpace(NewOutputCellAddress))
            return;

        if (string.IsNullOrWhiteSpace(NewOutputLabel))
            NewOutputLabel = NewOutputCellAddress;

        var card = new OutputCardViewModel(NewOutputSheetName, NewOutputCellAddress, NewOutputLabel);
        Outputs.Add(card);
        IsAddingOutput = false;

        OutputAdded?.Invoke(this, new OutputAddedEventArgs(card));
    }

    [RelayCommand]
    private void RemoveInput(InputCardViewModel input)
    {
        Inputs.Remove(input);
        InputRemoved?.Invoke(this, new InputRemovedEventArgs(input));
    }

    [RelayCommand]
    private void RemoveOutput(OutputCardViewModel output)
    {
        Outputs.Remove(output);
        OutputRemoved?.Invoke(this, new OutputRemovedEventArgs(output));
    }

    [RelayCommand]
    private void AddDiscretePair()
    {
        DiscretePairs.Add(new DiscretePairViewModel { Value = "0", Probability = "0" });
    }

    [RelayCommand]
    private void RemoveDiscretePair(DiscretePairViewModel pair)
    {
        if (DiscretePairs.Count > 1)
            DiscretePairs.Remove(pair);
    }

    [RelayCommand]
    private void RunSimulation()
    {
        RunSimulationRequested?.Invoke(this, EventArgs.Empty);
    }

    [RelayCommand]
    private void ToggleSeedLock()
    {
        IsSeedLocked = !IsSeedLocked;
        if (!IsSeedLocked)
            RandomSeed = null;
        else
            RandomSeed ??= 42;
    }

    private Dictionary<string, double> BuildParameterDictionary()
    {
        var p = new Dictionary<string, double>();
        var distName = SelectedDistribution.ToLowerInvariant();

        switch (distName)
        {
            case "normal":
                p["mean"] = ParseDouble(ParamMean, "Mean");
                p["stdDev"] = ParseDouble(ParamStdDev, "Std Dev");
                break;
            case "triangular":
            case "pert":
                p["min"] = ParseDouble(ParamMin, "Minimum");
                p["mode"] = ParseDouble(ParamMode, "Mode");
                p["max"] = ParseDouble(ParamMax, "Maximum");
                break;
            case "uniform":
                p["min"] = ParseDouble(ParamMin, "Minimum");
                p["max"] = ParseDouble(ParamMax, "Maximum");
                break;
            case "lognormal":
                p["mu"] = ParseDouble(ParamMu, "μ");
                p["sigma"] = ParseDouble(ParamSigma, "σ");
                break;
            case "discrete":
                for (int i = 0; i < DiscretePairs.Count; i++)
                {
                    p[$"value_{i}"] = ParseDouble(DiscretePairs[i].Value, $"Value {i + 1}");
                    p[$"prob_{i}"] = ParseDouble(DiscretePairs[i].Probability, $"Probability {i + 1}");
                }
                break;
        }

        return p;
    }

    private static double ParseDouble(string text, string fieldName)
    {
        if (!double.TryParse(text, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double value))
            throw new ArgumentException($"'{fieldName}' must be a valid number.");
        return value;
    }

    private void ResetInputEditor()
    {
        NewInputCellAddress = string.Empty;
        NewInputSheetName = string.Empty;
        NewInputLabel = string.Empty;
        SelectedDistribution = "Normal";
        ParamMean = "0";
        ParamStdDev = "1";
        ParamMin = "0";
        ParamMode = "50";
        ParamMax = "100";
        ParamMu = "0";
        ParamSigma = "1";
        DiscretePairs = new ObservableCollection<DiscretePairViewModel>
        {
            new() { Value = "0", Probability = "0.5" },
            new() { Value = "1", Probability = "0.5" }
        };
        InputValidationError = null;
        EditorPreviewPoints = null;
    }

    private void ResetOutputEditor()
    {
        NewOutputCellAddress = string.Empty;
        NewOutputSheetName = string.Empty;
        NewOutputLabel = string.Empty;
    }
}

/// <summary>
/// View model for a discrete distribution value/probability pair.
/// </summary>
public partial class DiscretePairViewModel : ObservableObject
{
    [ObservableProperty] private string _value = "0";
    [ObservableProperty] private string _probability = "0";
}

/// <summary>Event args for cell selection requests.</summary>
public class CellSelectionRequestedEventArgs : EventArgs
{
    public string Mode { get; }
    public CellSelectionRequestedEventArgs(string mode) => Mode = mode;
}

/// <summary>Event args when an input is added.</summary>
public class InputAddedEventArgs : EventArgs
{
    public InputCardViewModel Input { get; }
    public InputAddedEventArgs(InputCardViewModel input) => Input = input;
}

/// <summary>Event args when an output is added.</summary>
public class OutputAddedEventArgs : EventArgs
{
    public OutputCardViewModel Output { get; }
    public OutputAddedEventArgs(OutputCardViewModel output) => Output = output;
}

/// <summary>Event args when an input is removed.</summary>
public class InputRemovedEventArgs : EventArgs
{
    public InputCardViewModel Input { get; }
    public InputRemovedEventArgs(InputCardViewModel input) => Input = input;
}

/// <summary>Event args when an output is removed.</summary>
public class OutputRemovedEventArgs : EventArgs
{
    public OutputCardViewModel Output { get; }
    public OutputRemovedEventArgs(OutputCardViewModel output) => Output = output;
}
