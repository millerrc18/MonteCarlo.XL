using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Analysis;
using MonteCarlo.Engine.Distributions;
using MonteCarlo.Engine.Simulation;

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

    /// <summary>Selected named run preset, when the iteration count matches one.</summary>
    [ObservableProperty]
    private RunPresetOption? _selectedRunPreset = RunPresetOption.FindByIterations(5000);

    /// <summary>Optional random seed for reproducibility. Null = auto.</summary>
    [ObservableProperty]
    private int? _randomSeed;

    /// <summary>Whether the seed is locked (user-specified).</summary>
    [ObservableProperty]
    private bool _isSeedLocked;

    /// <summary>Whether the model manager table is visible.</summary>
    [ObservableProperty]
    private bool _isManagerVisible;

    /// <summary>Search text for the input manager table.</summary>
    [ObservableProperty]
    private string _inputManagerSearchText = string.Empty;

    /// <summary>Status text for the input manager table.</summary>
    [ObservableProperty]
    private string _inputManagerStatus = "0 inputs";

    /// <summary>Search text for the output manager table.</summary>
    [ObservableProperty]
    private string _outputManagerSearchText = string.Empty;

    /// <summary>Status text for the output manager table.</summary>
    [ObservableProperty]
    private string _outputManagerStatus = "0 outputs";

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
    /// <summary>Beta: Alpha.</summary>
    [ObservableProperty] private string _paramAlpha = "2";
    /// <summary>Beta: Beta.</summary>
    [ObservableProperty] private string _paramBeta = "5";
    /// <summary>Weibull: Shape.</summary>
    [ObservableProperty] private string _paramShape = "2";
    /// <summary>Weibull: Scale.</summary>
    [ObservableProperty] private string _paramScale = "100";
    /// <summary>Exponential/Poisson: Rate/Lambda.</summary>
    [ObservableProperty] private string _paramRate = "1";
    /// <summary>Poisson: Lambda.</summary>
    [ObservableProperty] private string _paramLambda = "4.5";
    /// <summary>Gamma: Rate.</summary>
    [ObservableProperty] private string _paramRateGamma = "2";
    /// <summary>Logistic: Scale.</summary>
    [ObservableProperty] private string _paramS = "1";
    /// <summary>GEV: Shape.</summary>
    [ObservableProperty] private string _paramXi = "0";
    /// <summary>Binomial: Trials.</summary>
    [ObservableProperty] private string _paramN = "10";
    /// <summary>Binomial/Geometric: Probability.</summary>
    [ObservableProperty] private string _paramP = "0.5";

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

    /// <summary>Selected plain-English distribution suggestion.</summary>
    [ObservableProperty]
    private DistributionSuggestionViewModel? _selectedDistributionSuggestion;

    /// <summary>Filtered distribution suggestions shown in the helper.</summary>
    [ObservableProperty]
    private ObservableCollection<DistributionSuggestionViewModel> _filteredDistributionSuggestions = new();

    /// <summary>Selected helper category for narrowing distribution suggestions.</summary>
    [ObservableProperty]
    private string _selectedDistributionSuggestionCategory = "All";

    /// <summary>Selected helper complexity for narrowing distribution suggestions.</summary>
    [ObservableProperty]
    private string _selectedDistributionSuggestionComplexity = "All";

    /// <summary>Search text for narrowing distribution suggestions.</summary>
    [ObservableProperty]
    private string _distributionSuggestionSearchText = string.Empty;

    /// <summary>Status text for the helper filter results.</summary>
    [ObservableProperty]
    private string _distributionSuggestionStatus = string.Empty;

    /// <summary>Prompt answer for the primary kind of uncertainty.</summary>
    [ObservableProperty]
    private string _selectedDistributionWizardDomain = "Any";

    /// <summary>Prompt answer for what information the analyst already has.</summary>
    [ObservableProperty]
    private string _selectedDistributionWizardEvidence = "Any";

    /// <summary>Prompt answer for the expected shape.</summary>
    [ObservableProperty]
    private string _selectedDistributionWizardShape = "Any";

    /// <summary>Status text for the guided distribution helper prompts.</summary>
    [ObservableProperty]
    private string _distributionWizardStatus = "Answer the prompts to rank good starting points.";

    /// <summary>Distribution fits computed from a selected Excel range.</summary>
    [ObservableProperty]
    private ObservableCollection<DistributionFitResultViewModel> _distributionFitResults = new();

    /// <summary>Selected fit result to apply to the input editor.</summary>
    [ObservableProperty]
    private DistributionFitResultViewModel? _selectedDistributionFitResult;

    /// <summary>Status text for distribution fitting.</summary>
    [ObservableProperty]
    private string? _distributionFitStatus;

    /// <summary>Histogram bars for the selected fit preview.</summary>
    [ObservableProperty]
    private IReadOnlyList<DistributionFitHistogramBar> _distributionFitPreviewBars = Array.Empty<DistributionFitHistogramBar>();

    /// <summary>Distribution curve points for the selected fit preview.</summary>
    [ObservableProperty]
    private IReadOnlyList<(double X, double Y)>? _distributionFitPreviewCurvePoints;

    /// <summary>Confidence or caution text for the selected fit.</summary>
    [ObservableProperty]
    private string? _distributionFitWarning;

    /// <summary>Available distribution names.</summary>
    public IReadOnlyList<string> AvailableDistributions { get; } = DistributionFactory.AvailableDistributions;

    /// <summary>Plain-English distribution suggestions for the input wizard.</summary>
    public IReadOnlyList<DistributionSuggestionViewModel> DistributionSuggestions { get; } =
        DistributionSuggestionViewModel.Defaults;

    /// <summary>Distribution helper categories.</summary>
    public IReadOnlyList<string> DistributionSuggestionCategories { get; } =
        DistributionSuggestionViewModel.Categories;

    /// <summary>Distribution helper complexity filters.</summary>
    public IReadOnlyList<string> DistributionSuggestionComplexities { get; } =
        DistributionSuggestionViewModel.Complexities;

    /// <summary>Guided prompt choices for the type of uncertainty being modeled.</summary>
    public IReadOnlyList<string> DistributionWizardDomains { get; } =
        DistributionSuggestionViewModel.WizardDomains;

    /// <summary>Guided prompt choices for what evidence is available.</summary>
    public IReadOnlyList<string> DistributionWizardEvidenceOptions { get; } =
        DistributionSuggestionViewModel.WizardEvidenceOptions;

    /// <summary>Guided prompt choices for the expected shape.</summary>
    public IReadOnlyList<string> DistributionWizardShapes { get; } =
        DistributionSuggestionViewModel.WizardShapes;

    /// <summary>Common iteration count presets.</summary>
    public IReadOnlyList<int> IterationPresets { get; } = new[] { 1000, 5000, 10000, 25000, 50000 };

    /// <summary>Named speed/accuracy presets for common simulation sizes.</summary>
    public IReadOnlyList<RunPresetOption> RunPresets { get; } = RunPresetOption.Defaults;

    /// <summary>Whether the Run button should be enabled.</summary>
    public bool CanRun => Inputs.Count > 0 && Outputs.Count > 0;

    /// <summary>Whether the Define Correlations button is enabled (2+ inputs).</summary>
    public bool CanDefineCorrelations => Inputs.Count >= 2;

    /// <summary>Saved correlation matrix values. Null means independent.</summary>
    public double[,]? CorrelationMatrixValues { get; set; }

    /// <summary>Label for the model manager toggle button.</summary>
    public string ManagerToggleLabel => IsManagerVisible ? "Hide Manager" : "Open Manager";

    /// <summary>Sortable/filterable view of configured inputs for the manager table.</summary>
    public ICollectionView InputManagerView { get; }

    /// <summary>Sortable/filterable view of configured outputs for the manager table.</summary>
    public ICollectionView OutputManagerView { get; }

    private InputCardViewModel? _editingInput;
    private OutputCardViewModel? _editingOutput;
    private readonly HashSet<InputCardViewModel> _observedInputs = new();
    private readonly HashSet<OutputCardViewModel> _observedOutputs = new();
    private double[] _distributionFitSourceSamples = Array.Empty<double>();

    partial void OnIterationCountChanged(int value)
    {
        var matchingPreset = RunPresets.FirstOrDefault(p => p.Iterations == value);
        if (!Equals(SelectedRunPreset, matchingPreset))
            SelectedRunPreset = matchingPreset;
    }

    partial void OnSelectedRunPresetChanged(RunPresetOption? value)
    {
        if (value != null && IterationCount != value.Iterations)
            IterationCount = value.Iterations;
    }

    partial void OnSelectedDistributionSuggestionCategoryChanged(string value) =>
        RefreshDistributionSuggestions();

    partial void OnSelectedDistributionSuggestionComplexityChanged(string value) =>
        RefreshDistributionSuggestions();

    partial void OnDistributionSuggestionSearchTextChanged(string value) =>
        RefreshDistributionSuggestions();

    partial void OnSelectedDistributionWizardDomainChanged(string value) =>
        RefreshDistributionSuggestions();

    partial void OnSelectedDistributionWizardEvidenceChanged(string value) =>
        RefreshDistributionSuggestions();

    partial void OnSelectedDistributionWizardShapeChanged(string value) =>
        RefreshDistributionSuggestions();

    partial void OnSelectedDistributionFitResultChanged(DistributionFitResultViewModel? value) =>
        UpdateDistributionFitPreview();

    partial void OnInputManagerSearchTextChanged(string value) =>
        RefreshInputManagerView();

    partial void OnOutputManagerSearchTextChanged(string value) =>
        RefreshOutputManagerView();

    /// <summary>Event raised when the user wants to open the correlation editor.</summary>
    public event Action? CorrelationEditorRequested;

    /// <summary>Event raised when the user wants to run model preflight checks.</summary>
    public event Action? PreflightRequested;

    /// <summary>Event raised when the user asks to fit a distribution from the current Excel selection.</summary>
    public event Action? DistributionFitRequested;

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

    /// <summary>
    /// Event raised when the user asks Excel to select or highlight a configured setup cell.
    /// </summary>
    public event EventHandler<SetupCellActionEventArgs>? CellActionRequested;

    /// <summary>
    /// Event raised when the user asks Excel to repaint all input/output highlights.
    /// </summary>
    public event EventHandler? RefreshHighlightsRequested;

    public SetupViewModel()
    {
        InputManagerView = new ListCollectionView(Inputs);
        InputManagerView.Filter = FilterInputManagerItem;
        InputManagerView.SortDescriptions.Add(new SortDescription(nameof(InputCardViewModel.Label), ListSortDirection.Ascending));

        OutputManagerView = new ListCollectionView(Outputs);
        OutputManagerView.Filter = FilterOutputManagerItem;
        OutputManagerView.SortDescriptions.Add(new SortDescription(nameof(OutputCardViewModel.Label), ListSortDirection.Ascending));

        RefreshDistributionSuggestions();
        Inputs.CollectionChanged += OnInputsCollectionChanged;
        Outputs.CollectionChanged += OnOutputsCollectionChanged;
        RefreshInputManagerView();
        RefreshOutputManagerView();
    }

    [RelayCommand]
    private void OpenCorrelationEditor()
    {
        CorrelationEditorRequested?.Invoke();
    }

    [RelayCommand]
    private void RunPreflight()
    {
        PreflightRequested?.Invoke();
    }

    [RelayCommand]
    private void ApplyDistributionSuggestion()
    {
        var suggestion = SelectedDistributionSuggestion;
        if (suggestion == null)
            return;

        NewInputLabel = string.IsNullOrWhiteSpace(NewInputLabel)
            ? suggestion.LabelExample
            : NewInputLabel;
        ApplyDetectedDistribution(suggestion.DistributionName, new Dictionary<string, double>(suggestion.Parameters));
    }

    [RelayCommand]
    private void ResetDistributionWizard()
    {
        SelectedDistributionWizardDomain = "Any";
        SelectedDistributionWizardEvidence = "Any";
        SelectedDistributionWizardShape = "Any";
    }

    [RelayCommand]
    private void RequestDistributionFit()
    {
        DistributionFitStatus = "Select a numeric Excel range for fitting.";
        DistributionFitRequested?.Invoke();
    }

    [RelayCommand]
    private void ApplyDistributionFit()
    {
        var fit = SelectedDistributionFitResult;
        if (fit == null)
            return;

        NewInputLabel = string.IsNullOrWhiteSpace(NewInputLabel)
            ? $"{fit.DistributionName} fit"
            : NewInputLabel;
        ApplyDetectedDistribution(fit.DistributionName, new Dictionary<string, double>(fit.Parameters));
        DistributionFitStatus = $"Applied {fit.DistributionName} fit.";
    }

    public void LoadDistributionFitResults(
        IReadOnlyList<DistributionFitResult> results,
        string sourceAddress,
        IReadOnlyList<double>? sourceSamples = null)
    {
        _distributionFitSourceSamples = sourceSamples?
            .Where(double.IsFinite)
            .ToArray() ?? Array.Empty<double>();
        DistributionFitResults = new ObservableCollection<DistributionFitResultViewModel>(
            results.Select(result => new DistributionFitResultViewModel(result)));
        SelectedDistributionFitResult = DistributionFitResults.FirstOrDefault();
        DistributionFitStatus = DistributionFitResults.Count == 0
            ? $"No supported fits were found for {sourceAddress}."
            : $"Fitted {DistributionFitResults.Count} candidates from {sourceAddress}.";
        UpdateDistributionFitPreview();
    }

    public void ShowDistributionFitError(string message)
    {
        DistributionFitResults.Clear();
        SelectedDistributionFitResult = null;
        DistributionFitStatus = message;
        DistributionFitPreviewBars = Array.Empty<DistributionFitHistogramBar>();
        DistributionFitPreviewCurvePoints = null;
        DistributionFitWarning = null;
        _distributionFitSourceSamples = Array.Empty<double>();
    }

    [RelayCommand]
    private void ToggleManager()
    {
        IsManagerVisible = !IsManagerVisible;
    }

    partial void OnIsManagerVisibleChanged(bool value)
    {
        OnPropertyChanged(nameof(ManagerToggleLabel));
        RefreshInputManagerView();
        RefreshOutputManagerView();
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

    private void UpdateDistributionFitPreview()
    {
        var fit = SelectedDistributionFitResult;
        if (fit == null || _distributionFitSourceSamples.Length < 3)
        {
            DistributionFitPreviewBars = Array.Empty<DistributionFitHistogramBar>();
            DistributionFitPreviewCurvePoints = null;
            DistributionFitWarning = null;
            return;
        }

        try
        {
            var distribution = DistributionFactory.Create(fit.DistributionName, fit.Parameters);
            DistributionFitPreviewCurvePoints = InputCardViewModel.ComputePreviewPointsStatic(distribution);
            DistributionFitPreviewBars = BuildDistributionFitHistogram(_distributionFitSourceSamples);
            DistributionFitWarning = BuildDistributionFitWarning(fit, _distributionFitSourceSamples.Length);
        }
        catch (Exception ex)
        {
            DistributionFitPreviewBars = Array.Empty<DistributionFitHistogramBar>();
            DistributionFitPreviewCurvePoints = null;
            DistributionFitWarning = $"Preview unavailable: {ex.Message}";
        }
    }

    private static IReadOnlyList<DistributionFitHistogramBar> BuildDistributionFitHistogram(IReadOnlyList<double> samples)
    {
        if (samples.Count == 0)
            return Array.Empty<DistributionFitHistogramBar>();

        var sorted = samples
            .Where(double.IsFinite)
            .OrderBy(value => value)
            .ToArray();
        if (sorted.Length == 0)
            return Array.Empty<DistributionFitHistogramBar>();

        var min = sorted[0];
        var max = sorted[^1];
        var distinctCount = sorted.Distinct().Count();
        if (sorted.All(value => Math.Abs(value - Math.Round(value)) < 1e-9) && distinctCount <= 20)
        {
            return sorted
                .GroupBy(value => value)
                .OrderBy(group => group.Key)
                .Select(group => new DistributionFitHistogramBar(
                    group.Key,
                    0.4,
                    group.Count() / (double)sorted.Length))
                .ToArray();
        }

        if (max <= min)
        {
            return new[]
            {
                new DistributionFitHistogramBar(min, 0.5, 1.0)
            };
        }

        var binCount = Math.Clamp((int)Math.Round(Math.Sqrt(sorted.Length)), 8, 20);
        var width = (max - min) / binCount;
        if (width <= 0)
            width = 1;

        var counts = new int[binCount];
        foreach (var value in sorted)
        {
            var index = (int)Math.Floor((value - min) / width);
            index = Math.Clamp(index, 0, binCount - 1);
            counts[index]++;
        }

        var bars = new List<DistributionFitHistogramBar>(binCount);
        for (var i = 0; i < binCount; i++)
        {
            var center = min + width * (i + 0.5);
            var height = counts[i] / (sorted.Length * width);
            bars.Add(new DistributionFitHistogramBar(center, width / 2, height));
        }

        return bars;
    }

    private static string? BuildDistributionFitWarning(DistributionFitResultViewModel fit, int sampleCount)
    {
        if (sampleCount < 12)
            return $"Only {sampleCount} samples were fitted. Treat this as a rough starting point and sanity-check the tails.";

        if (fit.Score >= 0.20)
            return $"The best KS distance is {fit.Score:0.000}, which is weak. Review the overlay before using this fit.";

        if (fit.Score >= 0.12)
            return $"The best KS distance is {fit.Score:0.000}. Use caution and compare the curve to your historical shape.";

        if (sampleCount < 30)
            return $"The fit looks reasonable, but it is based on only {sampleCount} samples.";

        return null;
    }

    [RelayCommand]
    private void ShowAddInput()
    {
        IsAddingInput = true;
        IsAddingOutput = false;
        _editingInput = null;
        ResetInputEditor();
    }

    /// <summary>
    /// Public entry point for host integrations such as the Excel ribbon.
    /// </summary>
    public void BeginAddInput() => ShowAddInput();

    [RelayCommand]
    private void ShowAddOutput()
    {
        IsAddingOutput = true;
        IsAddingInput = false;
        _editingOutput = null;
        ResetOutputEditor();
    }

    /// <summary>
    /// Public entry point for host integrations such as the Excel ribbon.
    /// </summary>
    public void BeginAddOutput() => ShowAddOutput();

    [RelayCommand]
    private void CancelAddInput()
    {
        IsAddingInput = false;
        IsSelectingCell = false;
        _editingInput = null;
    }

    [RelayCommand]
    private void CancelAddOutput()
    {
        IsAddingOutput = false;
        IsSelectingOutputCell = false;
        _editingOutput = null;
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
    public void OnCellSelected(
        string sheetName,
        string cellAddress,
        string? suggestedLabel,
        string? detectedDistribution = null,
        Dictionary<string, double>? detectedParameters = null)
    {
        if (IsSelectingCell)
        {
            NewInputSheetName = sheetName;
            NewInputCellAddress = cellAddress;
            if (!string.IsNullOrEmpty(suggestedLabel))
                NewInputLabel = suggestedLabel;

            if (!string.IsNullOrEmpty(detectedDistribution) && detectedParameters != null)
                ApplyDetectedDistribution(detectedDistribution, detectedParameters);

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

    private void ApplyDetectedDistribution(string distributionName, Dictionary<string, double> parameters)
    {
        // Match the canonical case from AvailableDistributions (dropdown uses exact strings)
        var canonical = AvailableDistributions.FirstOrDefault(
            d => string.Equals(d, distributionName, StringComparison.OrdinalIgnoreCase));
        if (canonical == null) return;

        SelectedDistribution = canonical;

        string Fmt(double v) => v.ToString("G", System.Globalization.CultureInfo.InvariantCulture);

        switch (distributionName.ToLowerInvariant())
        {
            case "normal":
                if (parameters.TryGetValue("mean", out var mean)) ParamMean = Fmt(mean);
                if (parameters.TryGetValue("stdDev", out var sd)) ParamStdDev = Fmt(sd);
                break;
            case "triangular":
            case "pert":
                if (parameters.TryGetValue("min", out var tmin)) ParamMin = Fmt(tmin);
                if (parameters.TryGetValue("mode", out var tmode)) ParamMode = Fmt(tmode);
                if (parameters.TryGetValue("max", out var tmax)) ParamMax = Fmt(tmax);
                break;
            case "uniform":
                if (parameters.TryGetValue("min", out var umin)) ParamMin = Fmt(umin);
                if (parameters.TryGetValue("max", out var umax)) ParamMax = Fmt(umax);
                break;
            case "lognormal":
                if (parameters.TryGetValue("mu", out var mu)) ParamMu = Fmt(mu);
                if (parameters.TryGetValue("sigma", out var sigma)) ParamSigma = Fmt(sigma);
                break;
            case "beta":
                if (parameters.TryGetValue("alpha", out var alpha)) ParamAlpha = Fmt(alpha);
                if (parameters.TryGetValue("beta", out var beta)) ParamBeta = Fmt(beta);
                break;
            case "weibull":
                if (parameters.TryGetValue("shape", out var shape)) ParamShape = Fmt(shape);
                if (parameters.TryGetValue("scale", out var scale)) ParamScale = Fmt(scale);
                break;
            case "exponential":
                if (parameters.TryGetValue("rate", out var rate)) ParamRate = Fmt(rate);
                break;
            case "poisson":
                if (parameters.TryGetValue("lambda", out var lambda)) ParamLambda = Fmt(lambda);
                break;
            case "gamma":
                if (parameters.TryGetValue("shape", out var gshape)) ParamShape = Fmt(gshape);
                if (parameters.TryGetValue("rate", out var grate)) ParamRateGamma = Fmt(grate);
                break;
            case "logistic":
                if (parameters.TryGetValue("mu", out var lmu)) ParamMu = Fmt(lmu);
                if (parameters.TryGetValue("s", out var ls)) ParamS = Fmt(ls);
                break;
            case "gev":
                if (parameters.TryGetValue("mu", out var gmu)) ParamMu = Fmt(gmu);
                if (parameters.TryGetValue("sigma", out var gsigma)) ParamSigma = Fmt(gsigma);
                if (parameters.TryGetValue("xi", out var gxi)) ParamXi = Fmt(gxi);
                break;
            case "binomial":
                if (parameters.TryGetValue("n", out var bn)) ParamN = Fmt(bn);
                if (parameters.TryGetValue("p", out var bp)) ParamP = Fmt(bp);
                break;
            case "geometric":
                if (parameters.TryGetValue("p", out var geop)) ParamP = Fmt(geop);
                break;
            case "discrete":
                var pairs = new ObservableCollection<DiscretePairViewModel>();
                for (var i = 0; parameters.ContainsKey($"value_{i}") || parameters.ContainsKey($"prob_{i}"); i++)
                {
                    pairs.Add(new DiscretePairViewModel
                    {
                        Value = parameters.TryGetValue($"value_{i}", out var value) ? Fmt(value) : "0",
                        Probability = parameters.TryGetValue($"prob_{i}", out var probability) ? Fmt(probability) : "0"
                    });
                }

                if (pairs.Count > 0)
                    DiscretePairs = pairs;
                break;
        }

        UpdateEditorPreview();
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

            if (_editingInput != null)
            {
                var oldInput = _editingInput;
                var index = Inputs.IndexOf(oldInput);
                if (index >= 0)
                    Inputs[index] = card;
                else
                    Inputs.Add(card);

                InputRemoved?.Invoke(this, new InputRemovedEventArgs(oldInput));
                InputAdded?.Invoke(this, new InputAddedEventArgs(card));
                _editingInput = null;
            }
            else
            {
                Inputs.Add(card);
                InputAdded?.Invoke(this, new InputAddedEventArgs(card));
            }

            IsAddingInput = false;
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
        if (_editingOutput != null)
        {
            var oldOutput = _editingOutput;
            var index = Outputs.IndexOf(oldOutput);
            if (index >= 0)
                Outputs[index] = card;
            else
                Outputs.Add(card);

            OutputRemoved?.Invoke(this, new OutputRemovedEventArgs(oldOutput));
            OutputAdded?.Invoke(this, new OutputAddedEventArgs(card));
            _editingOutput = null;
        }
        else
        {
            Outputs.Add(card);
            OutputAdded?.Invoke(this, new OutputAddedEventArgs(card));
        }

        IsAddingOutput = false;
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
    private void EditInput(InputCardViewModel input)
    {
        _editingInput = input;
        IsAddingInput = true;
        IsAddingOutput = false;
        NewInputSheetName = input.SheetName;
        NewInputCellAddress = input.CellAddress;
        NewInputLabel = input.Label;
        ApplyDetectedDistribution(input.DistributionName, input.Parameters);
    }

    [RelayCommand]
    private void DuplicateInput(InputCardViewModel input)
    {
        _editingInput = null;
        IsAddingInput = true;
        IsAddingOutput = false;
        NewInputSheetName = string.Empty;
        NewInputCellAddress = string.Empty;
        NewInputLabel = $"{input.Label} copy";
        ApplyDetectedDistribution(input.DistributionName, input.Parameters);
    }

    [RelayCommand]
    private void JumpToInput(InputCardViewModel input)
    {
        CellActionRequested?.Invoke(this, SetupCellActionEventArgs.JumpToInput(input));
    }

    [RelayCommand]
    private void HighlightInput(InputCardViewModel input)
    {
        CellActionRequested?.Invoke(this, SetupCellActionEventArgs.HighlightInput(input));
    }

    [RelayCommand]
    private void EditOutput(OutputCardViewModel output)
    {
        _editingOutput = output;
        IsAddingOutput = true;
        IsAddingInput = false;
        NewOutputSheetName = output.SheetName;
        NewOutputCellAddress = output.CellAddress;
        NewOutputLabel = output.Label;
    }

    [RelayCommand]
    private void DuplicateOutput(OutputCardViewModel output)
    {
        _editingOutput = null;
        IsAddingOutput = true;
        IsAddingInput = false;
        NewOutputSheetName = string.Empty;
        NewOutputCellAddress = string.Empty;
        NewOutputLabel = $"{output.Label} copy";
    }

    [RelayCommand]
    private void JumpToOutput(OutputCardViewModel output)
    {
        CellActionRequested?.Invoke(this, SetupCellActionEventArgs.JumpToOutput(output));
    }

    [RelayCommand]
    private void HighlightOutput(OutputCardViewModel output)
    {
        CellActionRequested?.Invoke(this, SetupCellActionEventArgs.HighlightOutput(output));
    }

    [RelayCommand]
    private void RefreshHighlights()
    {
        RefreshHighlightsRequested?.Invoke(this, EventArgs.Empty);
    }

    [RelayCommand]
    private void ClearInputs()
    {
        foreach (var input in Inputs.ToList())
            RemoveInput(input);
    }

    [RelayCommand]
    private void ClearOutputs()
    {
        foreach (var output in Outputs.ToList())
            RemoveOutput(output);
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

    /// <summary>
    /// Builds a serializable profile from the current setup state for validation and persistence.
    /// </summary>
    public SimulationProfile BuildSimulationProfile()
    {
        var profile = new SimulationProfile
        {
            IterationCount = IterationCount,
            RandomSeed = RandomSeed,
            CorrelationMatrix = CorrelationMatrixValues
        };

        foreach (var input in Inputs)
        {
            profile.Inputs.Add(new SavedInput
            {
                SheetName = input.SheetName,
                CellAddress = input.CellAddress,
                Label = input.Label,
                DistributionName = input.DistributionName,
                Parameters = new Dictionary<string, double>(input.Parameters)
            });
        }

        foreach (var output in Outputs)
        {
            profile.Outputs.Add(new SavedOutput
            {
                SheetName = output.SheetName,
                CellAddress = output.CellAddress,
                Label = output.Label
            });
        }

        return profile;
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

    private void OnInputsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(CanRun));
        OnPropertyChanged(nameof(CanDefineCorrelations));
        RebuildObservedInputs();
        RefreshInputManagerView();
    }

    private void OnOutputsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(CanRun));
        RebuildObservedOutputs();
        RefreshOutputManagerView();
    }

    private void RebuildObservedInputs()
    {
        foreach (var input in _observedInputs.Except(Inputs).ToList())
        {
            input.PropertyChanged -= OnInputCardPropertyChanged;
            _observedInputs.Remove(input);
        }

        foreach (var input in Inputs.Where(input => !_observedInputs.Contains(input)))
        {
            input.PropertyChanged += OnInputCardPropertyChanged;
            _observedInputs.Add(input);
        }
    }

    private void RebuildObservedOutputs()
    {
        foreach (var output in _observedOutputs.Except(Outputs).ToList())
        {
            output.PropertyChanged -= OnOutputCardPropertyChanged;
            _observedOutputs.Remove(output);
        }

        foreach (var output in Outputs.Where(output => !_observedOutputs.Contains(output)))
        {
            output.PropertyChanged += OnOutputCardPropertyChanged;
            _observedOutputs.Add(output);
        }
    }

    private void OnInputCardPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(InputCardViewModel.Label))
            RefreshInputManagerView();
    }

    private void OnOutputCardPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(OutputCardViewModel.Label))
            RefreshOutputManagerView();
    }

    private bool FilterInputManagerItem(object item)
    {
        if (item is not InputCardViewModel input)
            return false;

        var query = InputManagerSearchText.Trim();
        if (string.IsNullOrWhiteSpace(query))
            return true;

        return Contains(input.Label, query)
            || Contains(input.FullReference, query)
            || Contains(input.DistributionName, query)
            || Contains(input.ParameterSummary, query)
            || Contains(input.MeanSummary, query)
            || Contains(input.P5Summary, query)
            || Contains(input.P95Summary, query);
    }

    private bool FilterOutputManagerItem(object item)
    {
        if (item is not OutputCardViewModel output)
            return false;

        var query = OutputManagerSearchText.Trim();
        if (string.IsNullOrWhiteSpace(query))
            return true;

        return Contains(output.Label, query)
            || Contains(output.FullReference, query);
    }

    private void RefreshInputManagerView()
    {
        InputManagerView.Refresh();
        var filtered = InputManagerView.Cast<object>().Count();
        InputManagerStatus = filtered == Inputs.Count
            ? $"{filtered} input{(filtered == 1 ? string.Empty : "s")}"
            : $"Showing {filtered} of {Inputs.Count} inputs";
    }

    private void RefreshOutputManagerView()
    {
        OutputManagerView.Refresh();
        var filtered = OutputManagerView.Cast<object>().Count();
        OutputManagerStatus = filtered == Outputs.Count
            ? $"{filtered} output{(filtered == 1 ? string.Empty : "s")}"
            : $"Showing {filtered} of {Outputs.Count} outputs";
    }

    private static bool Contains(string? value, string query) =>
        !string.IsNullOrWhiteSpace(value)
        && value.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;

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
            case "beta":
                p["alpha"] = ParseDouble(ParamAlpha, "Alpha");
                p["beta"] = ParseDouble(ParamBeta, "Beta");
                break;
            case "weibull":
                p["shape"] = ParseDouble(ParamShape, "Shape");
                p["scale"] = ParseDouble(ParamScale, "Scale");
                break;
            case "exponential":
                p["rate"] = ParseDouble(ParamRate, "Rate");
                break;
            case "poisson":
                p["lambda"] = ParseDouble(ParamLambda, "Lambda");
                break;
            case "gamma":
                p["shape"] = ParseDouble(ParamShape, "Shape");
                p["rate"] = ParseDouble(ParamRateGamma, "Rate");
                break;
            case "logistic":
                p["mu"] = ParseDouble(ParamMu, "μ");
                p["s"] = ParseDouble(ParamS, "s");
                break;
            case "gev":
                p["mu"] = ParseDouble(ParamMu, "μ");
                p["sigma"] = ParseDouble(ParamSigma, "σ");
                p["xi"] = ParseDouble(ParamXi, "ξ");
                break;
            case "binomial":
                p["n"] = ParseDouble(ParamN, "n (trials)");
                p["p"] = ParseDouble(ParamP, "p (probability)");
                break;
            case "geometric":
                p["p"] = ParseDouble(ParamP, "p (probability)");
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
        ParamAlpha = "2";
        ParamBeta = "5";
        ParamShape = "2";
        ParamScale = "100";
        ParamRate = "1";
        ParamLambda = "4.5";
        ParamRateGamma = "2";
        ParamS = "1";
        ParamXi = "0";
        ParamN = "10";
        ParamP = "0.5";
        DiscretePairs = new ObservableCollection<DiscretePairViewModel>
        {
            new() { Value = "0", Probability = "0.5" },
            new() { Value = "1", Probability = "0.5" }
        };
        InputValidationError = null;
        EditorPreviewPoints = null;
    }

    private void RefreshDistributionSuggestions()
    {
        IEnumerable<DistributionSuggestionViewModel> suggestions = DistributionSuggestions;

        if (!string.Equals(SelectedDistributionSuggestionCategory, "All", StringComparison.OrdinalIgnoreCase))
        {
            suggestions = suggestions.Where(s =>
                string.Equals(s.Category, SelectedDistributionSuggestionCategory, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.Equals(SelectedDistributionSuggestionComplexity, "All", StringComparison.OrdinalIgnoreCase))
        {
            suggestions = suggestions.Where(s =>
                string.Equals(s.Complexity, SelectedDistributionSuggestionComplexity, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(DistributionSuggestionSearchText))
        {
            var searchText = DistributionSuggestionSearchText.Trim();
            suggestions = suggestions.Where(s => s.Matches(searchText));
        }

        var filtered = suggestions.ToList();
        var wizardActive =
            !string.Equals(SelectedDistributionWizardDomain, "Any", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(SelectedDistributionWizardEvidence, "Any", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(SelectedDistributionWizardShape, "Any", StringComparison.OrdinalIgnoreCase);

        if (filtered.Count == 0)
        {
            DistributionWizardStatus = "No starting points match the current filters.";
        }
        else if (wizardActive)
        {
            var ranked = filtered
                .Select(suggestion => new
                {
                    Suggestion = suggestion,
                    Score = suggestion.GetWizardScore(
                        SelectedDistributionWizardDomain,
                        SelectedDistributionWizardEvidence,
                        SelectedDistributionWizardShape)
                })
                .ToList();

            var matched = ranked
                .Where(item => item.Score > 0)
                .OrderByDescending(item => item.Score)
                .ThenBy(item => item.Suggestion.DisplayName, StringComparer.OrdinalIgnoreCase)
                .Select(item => item.Suggestion)
                .ToList();

            if (matched.Count > 0)
            {
                filtered = matched;
                DistributionWizardStatus = matched.Count == 1
                    ? "1 guided starting point matches your answers."
                    : $"{matched.Count} guided starting points match your answers.";
            }
            else
            {
                filtered = ranked
                    .OrderBy(item => item.Suggestion.DisplayName, StringComparer.OrdinalIgnoreCase)
                    .Select(item => item.Suggestion)
                    .ToList();
                DistributionWizardStatus = "No exact guided match yet, so the full list is shown.";
            }
        }
        else
        {
            DistributionWizardStatus = "Answer the prompts to rank good starting points.";
        }

        FilteredDistributionSuggestions = new ObservableCollection<DistributionSuggestionViewModel>(filtered);
        DistributionSuggestionStatus = filtered.Count == 1
            ? "1 starting point matches."
            : $"{filtered.Count} starting points match.";

        if (SelectedDistributionSuggestion == null || !filtered.Contains(SelectedDistributionSuggestion))
            SelectedDistributionSuggestion = filtered.FirstOrDefault();
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

public sealed class DistributionSuggestionViewModel
{
    public DistributionSuggestionViewModel(
        string title,
        string distributionName,
        string useCase,
        string parameterHint,
        string labelExample,
        Dictionary<string, double> parameters,
        string category = "General",
        string complexity = "Common",
        string keywords = "",
        string wizardDomain = "Any",
        string wizardEvidence = "Any",
        string wizardShape = "Any")
    {
        Title = title;
        DistributionName = distributionName;
        UseCase = useCase;
        ParameterHint = parameterHint;
        LabelExample = labelExample;
        Parameters = parameters;
        Category = category;
        Complexity = complexity;
        Keywords = keywords;
        WizardDomain = wizardDomain;
        WizardEvidence = wizardEvidence;
        WizardShape = wizardShape;
    }

    public string Title { get; }
    public string DistributionName { get; }
    public string UseCase { get; }
    public string ParameterHint { get; }
    public string LabelExample { get; }
    public Dictionary<string, double> Parameters { get; }
    public string Category { get; }
    public string Complexity { get; }
    public string Keywords { get; }
    public string WizardDomain { get; }
    public string WizardEvidence { get; }
    public string WizardShape { get; }
    public string DisplayName => $"{Title} ({DistributionName})";
    public string HelperSummary => $"{Category} • {Complexity}";

    public bool Matches(string searchText)
    {
        if (string.IsNullOrWhiteSpace(searchText))
            return true;

        return Title.Contains(searchText, StringComparison.OrdinalIgnoreCase)
               || DistributionName.Contains(searchText, StringComparison.OrdinalIgnoreCase)
               || UseCase.Contains(searchText, StringComparison.OrdinalIgnoreCase)
               || ParameterHint.Contains(searchText, StringComparison.OrdinalIgnoreCase)
               || Category.Contains(searchText, StringComparison.OrdinalIgnoreCase)
               || Keywords.Contains(searchText, StringComparison.OrdinalIgnoreCase);
    }

    public int GetWizardScore(string domain, string evidence, string shape)
    {
        var score = 0;
        if (!string.Equals(domain, "Any", StringComparison.OrdinalIgnoreCase)
            && string.Equals(WizardDomain, domain, StringComparison.OrdinalIgnoreCase))
            score++;

        if (!string.Equals(evidence, "Any", StringComparison.OrdinalIgnoreCase)
            && string.Equals(WizardEvidence, evidence, StringComparison.OrdinalIgnoreCase))
            score++;

        if (!string.Equals(shape, "Any", StringComparison.OrdinalIgnoreCase)
            && string.Equals(WizardShape, shape, StringComparison.OrdinalIgnoreCase))
            score++;

        return score;
    }

    public static IReadOnlyList<string> WizardDomains { get; } = new[]
    {
        "Any",
        "Three-point estimate",
        "General bounded range",
        "Symmetric around a center",
        "Positive amount",
        "Percentage or rate",
        "Count or trials",
        "Waiting time / reliability",
        "Scenario outcomes",
        "Extreme values"
    };

    public static IReadOnlyList<string> WizardEvidenceOptions { get; } = new[]
    {
        "Any",
        "Historical data",
        "Low / likely / high estimate",
        "Only bounds / no center",
        "Known trials or event rate",
        "Named scenarios"
    };

    public static IReadOnlyList<string> WizardShapes { get; } = new[]
    {
        "Any",
        "Symmetric",
        "Right-skewed",
        "Heavy tails",
        "Bounded",
        "Discrete"
    };

    public static IReadOnlyList<string> Categories { get; } = new[]
    {
        "All",
        "Three-point",
        "Symmetric",
        "Positive skew",
        "Percentage",
        "Range",
        "Scenarios",
        "Count",
        "Waiting time",
        "Reliability",
        "Extreme value"
    };

    public static IReadOnlyList<string> Complexities { get; } = new[]
    {
        "All",
        "Common",
        "Advanced"
    };

    public static IReadOnlyList<DistributionSuggestionViewModel> Defaults { get; } = new[]
    {
        new DistributionSuggestionViewModel(
            "Three-point estimate",
            "PERT",
            "Use when you know low, most likely, and high estimates and want a smooth expert-judgment curve.",
            "Starts with min 80, mode 100, max 140.",
            "Estimate",
            new Dictionary<string, double> { ["min"] = 80, ["mode"] = 100, ["max"] = 140 },
            "Three-point",
            "Common",
            "low most likely high estimate pert expert judgment project cost schedule",
            "Three-point estimate",
            "Low / likely / high estimate",
            "Bounded"),
        new DistributionSuggestionViewModel(
            "Simple low-likely-high",
            "Triangular",
            "Use when the three-point estimate should be direct and transparent.",
            "Starts with min 80, mode 100, max 140.",
            "Three point input",
            new Dictionary<string, double> { ["min"] = 80, ["mode"] = 100, ["max"] = 140 },
            "Three-point",
            "Common",
            "low likely high simple transparent project cost schedule",
            "Three-point estimate",
            "Low / likely / high estimate",
            "Bounded"),
        new DistributionSuggestionViewModel(
            "Symmetric forecast error",
            "Normal",
            "Use when values vary around a center and upside/downside misses are roughly balanced.",
            "Starts with mean 0 and standard deviation 1.",
            "Forecast error",
            new Dictionary<string, double> { ["mean"] = 0, ["stdDev"] = 1 },
            "Symmetric",
            "Common",
            "balanced around mean forecast error variance uncertainty",
            "Symmetric around a center",
            "Historical data",
            "Symmetric"),
        new DistributionSuggestionViewModel(
            "Heavier-tail forecast error",
            "Logistic",
            "Use when errors are centered but large misses happen more often than a Normal curve implies.",
            "Starts with location 0 and scale 1.",
            "Heavy-tail error",
            new Dictionary<string, double> { ["mu"] = 0, ["s"] = 1 },
            "Symmetric",
            "Advanced",
            "heavy tails centered forecast error large misses",
            "Symmetric around a center",
            "Historical data",
            "Heavy tails"),
        new DistributionSuggestionViewModel(
            "Positive skewed amount",
            "Lognormal",
            "Use for positive values with plausible large upside outcomes, such as deal size or repair cost.",
            "Starts with mu 0 and sigma 0.35.",
            "Positive amount",
            new Dictionary<string, double> { ["mu"] = 0, ["sigma"] = 0.35 },
            "Positive skew",
            "Common",
            "positive cost revenue deal size repair cost upside skew",
            "Positive amount",
            "Historical data",
            "Right-skewed"),
        new DistributionSuggestionViewModel(
            "Positive total or duration",
            "Gamma",
            "Use for positive continuous totals built from multiple waiting-time-like effects.",
            "Starts with shape 3 and rate 0.4.",
            "Processing time",
            new Dictionary<string, double> { ["shape"] = 3, ["rate"] = 0.4 },
            "Positive skew",
            "Advanced",
            "positive duration total processing time waiting time aggregate skew",
            "Positive amount",
            "Historical data",
            "Right-skewed"),
        new DistributionSuggestionViewModel(
            "Percentage or rate",
            "Beta",
            "Use for values constrained between 0 and 1, such as conversion, churn, or defect rates.",
            "Starts with alpha 8 and beta 32.",
            "Conversion rate",
            new Dictionary<string, double> { ["alpha"] = 8, ["beta"] = 32 },
            "Percentage",
            "Common",
            "percentage rate conversion churn defect probability bounded zero one",
            "Percentage or rate",
            "Historical data",
            "Bounded"),
        new DistributionSuggestionViewModel(
            "No better view than a range",
            "Uniform",
            "Use when every value in a known range is equally plausible.",
            "Starts with min 0 and max 1.",
            "Bounded range",
            new Dictionary<string, double> { ["min"] = 0, ["max"] = 1 },
            "Range",
            "Common",
            "bounded range equal probability min max no better view",
            "General bounded range",
            "Only bounds / no center",
            "Bounded"),
        new DistributionSuggestionViewModel(
            "Exact scenarios",
            "Discrete",
            "Use for a small set of known outcomes with assigned probabilities.",
            "Starts with two outcomes: 0 at 50% and 1 at 50%.",
            "Scenario outcome",
            new Dictionary<string, double> { ["value_0"] = 0, ["prob_0"] = 0.5, ["value_1"] = 1, ["prob_1"] = 0.5 },
            "Scenarios",
            "Common",
            "scenario exact outcomes probabilities discrete cases",
            "Scenario outcomes",
            "Named scenarios",
            "Discrete"),
        new DistributionSuggestionViewModel(
            "Count in a fixed period",
            "Poisson",
            "Use for independent event counts over a defined period.",
            "Starts with lambda 4.5.",
            "Event count",
            new Dictionary<string, double> { ["lambda"] = 4.5 },
            "Count",
            "Common",
            "count events arrivals fixed period incidents",
            "Count or trials",
            "Known trials or event rate",
            "Discrete"),
        new DistributionSuggestionViewModel(
            "Successes out of trials",
            "Binomial",
            "Use when there are a fixed number of independent yes/no attempts.",
            "Starts with 20 trials and 35% probability.",
            "Prospect wins",
            new Dictionary<string, double> { ["n"] = 20, ["p"] = 0.35 },
            "Count",
            "Common",
            "successes trials yes no wins conversions attempts probability",
            "Count or trials",
            "Known trials or event rate",
            "Discrete"),
        new DistributionSuggestionViewModel(
            "Attempts until first success",
            "Geometric",
            "Use for the number of tries until the first win or event.",
            "Starts with 25% success probability.",
            "Calls until win",
            new Dictionary<string, double> { ["p"] = 0.25 },
            "Count",
            "Advanced",
            "tries until first success attempts calls win probability",
            "Count or trials",
            "Known trials or event rate",
            "Discrete"),
        new DistributionSuggestionViewModel(
            "Time to failure",
            "Weibull",
            "Use for component life or reliability where failure risk changes over time.",
            "Starts with shape 1.6 and scale 1200.",
            "Failure time",
            new Dictionary<string, double> { ["shape"] = 1.6, ["scale"] = 1200 },
            "Reliability",
            "Advanced",
            "failure time reliability component life hazard aging",
            "Waiting time / reliability",
            "Historical data",
            "Right-skewed"),
        new DistributionSuggestionViewModel(
            "Waiting time to event",
            "Exponential",
            "Use for waiting time until an event with a constant hazard rate.",
            "Starts with rate 0.2.",
            "Wait time",
            new Dictionary<string, double> { ["rate"] = 0.2 },
            "Waiting time",
            "Common",
            "waiting time event constant hazard arrivals service",
            "Waiting time / reliability",
            "Known trials or event rate",
            "Right-skewed"),
        new DistributionSuggestionViewModel(
            "Extreme maximum or minimum",
            "GEV",
            "Use for block maxima or minima such as annual peak demand or worst annual loss.",
            "Starts with location 100, scale 12, shape 0.1.",
            "Annual maximum",
            new Dictionary<string, double> { ["mu"] = 100, ["sigma"] = 12, ["xi"] = 0.1 },
            "Extreme value",
            "Advanced",
            "extreme maximum minimum annual peak demand worst loss tail",
            "Extreme values",
            "Historical data",
            "Heavy tails")
    };
}

public sealed class DistributionFitResultViewModel
{
    public DistributionFitResultViewModel(DistributionFitResult result)
    {
        DistributionName = result.DistributionName;
        Parameters = result.Parameters.ToDictionary(pair => pair.Key, pair => pair.Value);
        Score = result.Score;
        ParameterSummary = result.ParameterSummary;
    }

    public string DistributionName { get; }
    public Dictionary<string, double> Parameters { get; }
    public double Score { get; }
    public string ParameterSummary { get; }
    public string DisplayName => $"{DistributionName} fit (KS {Score:0.000})";
    public string ScoreSummary => $"KS distance {Score:0.000}";
}

public sealed record DistributionFitHistogramBar(double Center, double HalfWidth, double Height);

public sealed class RunPresetOption
{
    public RunPresetOption(string name, int iterations, string description)
    {
        Name = name;
        Iterations = iterations;
        Description = description;
    }

    public string Name { get; }
    public int Iterations { get; }
    public string Description { get; }
    public string DisplayName => $"{Name} ({Iterations:N0})";

    public static IReadOnlyList<RunPresetOption> Defaults { get; } = new[]
    {
        new RunPresetOption("Preview", 1_000, "Fast smoke checks while building a model."),
        new RunPresetOption("Standard", 5_000, "Balanced speed and stability for everyday analysis."),
        new RunPresetOption("Full", 25_000, "More stable percentiles for reports and decisions."),
        new RunPresetOption("Deep", 50_000, "Higher precision for final checks on important models.")
    };

    public static RunPresetOption? FindByIterations(int iterations) =>
        Defaults.FirstOrDefault(p => p.Iterations == iterations);

    public override string ToString() => DisplayName;
}

/// <summary>Event args for cell selection requests.</summary>
public class CellSelectionRequestedEventArgs : EventArgs
{
    public string Mode { get; }
    public CellSelectionRequestedEventArgs(string mode) => Mode = mode;
}

public enum SetupCellAction
{
    Jump,
    Highlight
}

public enum SetupCellRole
{
    Input,
    Output
}

/// <summary>Event args when a configured input or output cell should be selected or highlighted in Excel.</summary>
public class SetupCellActionEventArgs : EventArgs
{
    public string SheetName { get; }
    public string CellAddress { get; }
    public SetupCellRole Role { get; }
    public SetupCellAction Action { get; }

    private SetupCellActionEventArgs(string sheetName, string cellAddress, SetupCellRole role, SetupCellAction action)
    {
        SheetName = sheetName;
        CellAddress = cellAddress;
        Role = role;
        Action = action;
    }

    public static SetupCellActionEventArgs JumpToInput(InputCardViewModel input) =>
        new(input.SheetName, input.CellAddress, SetupCellRole.Input, SetupCellAction.Jump);

    public static SetupCellActionEventArgs HighlightInput(InputCardViewModel input) =>
        new(input.SheetName, input.CellAddress, SetupCellRole.Input, SetupCellAction.Highlight);

    public static SetupCellActionEventArgs JumpToOutput(OutputCardViewModel output) =>
        new(output.SheetName, output.CellAddress, SetupCellRole.Output, SetupCellAction.Jump);

    public static SetupCellActionEventArgs HighlightOutput(OutputCardViewModel output) =>
        new(output.SheetName, output.CellAddress, SetupCellRole.Output, SetupCellAction.Highlight);
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
