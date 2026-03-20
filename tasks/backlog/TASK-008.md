# TASK-008: Setup View

## Context

Read `ROADMAP.md` for full project context. The task pane shell (TASK-007) has a placeholder SetupView. This task replaces it with the real configuration UI where users define simulation inputs, outputs, and distributions.

## Dependencies

- TASK-002 (DistributionFactory — to create distributions from user selections)
- TASK-006 (WorkbookManager, InputTagManager, OutputTagManager, CellHighlighter)
- TASK-007 (Task pane shell, GlobalStyles, navigation)

## Objective

Build the Setup view where users configure their simulation: select input cells, assign distributions with parameters, select output cells, and set simulation options (iteration count, seed). This is the primary pre-simulation workflow.

## Design

### Layout (380px task pane width)

```
┌──────────────────────────────────┐
│ SIMULATION SETUP                 │
├──────────────────────────────────┤
│ Iterations: [  5,000  ▾]        │
│ Seed: [ Auto      ] [lock icon] │
├──────────────────────────────────┤
│ INPUTS                    [+ Add]│
│ ┌──────────────────────────────┐ │
│ │ 📊 Material Cost     Sheet1!B4│ │
│ │    Normal(μ=100, σ=10)       │ │
│ │    [mini distribution curve] │ │
│ │                     [Edit][✕]│ │
│ ├──────────────────────────────┤ │
│ │ 📊 Labor Hours       Sheet1!B5│ │
│ │    PERT(80, 100, 130)        │ │
│ │    [mini distribution curve] │ │
│ │                     [Edit][✕]│ │
│ └──────────────────────────────┘ │
├──────────────────────────────────┤
│ OUTPUTS                   [+ Add]│
│ ┌──────────────────────────────┐ │
│ │ 📈 Net Profit       Sheet1!D10│ │
│ │                          [✕] │ │
│ ├──────────────────────────────┤ │
│ │ 📈 IRR              Sheet1!D12│ │
│ │                          [✕] │ │
│ └──────────────────────────────┘ │
├──────────────────────────────────┤
│ [       ▶ Run Simulation       ] │
└──────────────────────────────────┘
```

### Add Input Flow

When the user clicks "+ Add" on inputs:

1. The task pane shows an "input editor" inline (expands in place, or slides in):
   ```
   ┌──────────────────────────────┐
   │ NEW INPUT                    │
   │ Cell: [Click a cell...] [📌] │  ← User clicks cell in Excel, address populates
   │ Label: [Material Cost     ]  │  ← Auto-filled from cell comment or left of cell
   │ Distribution: [Normal     ▾] │
   │ ┌──────────────────────────┐ │
   │ │ Mean (μ):  [100        ] │ │  ← Parameters change based on distribution
   │ │ Std Dev (σ): [10       ] │ │
   │ └──────────────────────────┘ │
   │ [mini preview of the curve]  │  ← DistributionPreview from MonteCarlo.Charts
   │ [Cancel]          [Add Input]│
   └──────────────────────────────┘
   ```

2. When "Add Input" is clicked:
   - Validate parameters (delegate to DistributionFactory)
   - Add to InputTagManager
   - Highlight the cell in Excel (blue)
   - Collapse the editor, show the input in the list

### Cell Selection

When the user clicks the cell-select button (📌), the add-in enters "cell selection mode":
- The task pane shows a "Click a cell in Excel..." prompt
- When the user clicks a cell in Excel, capture its address via `Application.SheetSelectionChange` event
- Populate the Cell field with the reference (e.g., "Sheet1!B4")
- Auto-suggest a label from the cell to the left or above (common spreadsheet convention)
- Exit selection mode

### Distribution Parameter Forms

Each distribution shows different parameter fields:

| Distribution | Fields |
|-------------|--------|
| Normal | Mean (μ), Std Dev (σ) |
| Triangular | Minimum, Mode (Most Likely), Maximum |
| PERT | Minimum, Mode (Most Likely), Maximum |
| Lognormal | μ (log mean), σ (log std dev) |
| Uniform | Minimum, Maximum |
| Discrete | Value/Probability pairs (dynamic rows, + add row button) |

All numeric fields should:
- Accept decimal input
- Show validation errors inline (red border + message)
- Update the distribution preview in real-time as the user types (debounced 300ms)

### Add Output Flow

Simpler than inputs — just select a cell and give it a label:

1. Click "+ Add" on outputs
2. Select a cell in Excel
3. Enter a label (auto-suggested)
4. Click "Add Output"
5. Cell gets highlighted green

### Distribution Preview

The small distribution curve shown on each input card should use the `DistributionPreview` control from `MonteCarlo.Charts` (or build it here if TASK-009 hasn't created it yet). It's a tiny (300×60px) line chart showing the PDF shape. No axes, no labels — just the curve shape so the user can see "this is skewed right" vs "this is symmetric."

If the Charts project isn't ready, use a simple WPF `Polyline` drawing the PDF at ~50 evenly spaced points. This can be replaced later.

### Iteration Count

Provide a dropdown or input field with common presets:
- 1,000 (quick test)
- 5,000 (default — good balance of speed and accuracy)
- 10,000 (standard)
- 25,000 (high precision)
- 50,000 (maximum precision)
- Custom (user types a number)

### Run Simulation Button

The big button at the bottom:
- Disabled when: no inputs configured OR no outputs configured
- On click: validate config → build SimulationConfig → navigate to RunView → start simulation
- The actual simulation orchestration (calling SimulationEngine, collecting results) can be handled in the MainViewModel or a dedicated `SimulationOrchestrator` service

## ViewModels

### SetupViewModel

```csharp
public partial class SetupViewModel : ObservableObject
{
    [ObservableProperty] private ObservableCollection<InputCardViewModel> _inputs;
    [ObservableProperty] private ObservableCollection<OutputCardViewModel> _outputs;
    [ObservableProperty] private int _iterationCount = 5000;
    [ObservableProperty] private int? _randomSeed;
    [ObservableProperty] private bool _isAddingInput;
    [ObservableProperty] private bool _isAddingOutput;
    [ObservableProperty] private bool _canRun;  // true when inputs.Count > 0 && outputs.Count > 0

    // Input editor state
    [ObservableProperty] private string _newInputCellAddress;
    [ObservableProperty] private string _newInputLabel;
    [ObservableProperty] private string _selectedDistribution = "Normal";
    [ObservableProperty] private Dictionary<string, double> _distributionParameters;

    [RelayCommand] private void AddInput();
    [RelayCommand] private void RemoveInput(InputCardViewModel input);
    [RelayCommand] private void EditInput(InputCardViewModel input);
    [RelayCommand] private void AddOutput();
    [RelayCommand] private void RemoveOutput(OutputCardViewModel output);
    [RelayCommand] private void StartCellSelection();
    [RelayCommand] private void RunSimulation();
}
```

### InputCardViewModel

```csharp
public partial class InputCardViewModel : ObservableObject
{
    public string CellAddress { get; }
    public string SheetName { get; }
    public string Label { get; }
    public string DistributionName { get; }
    public string ParameterSummary { get; }     // e.g., "Normal(μ=100, σ=10)"
    public IDistribution Distribution { get; }  // For the preview chart
}
```

## Implementation Notes

### Cell Selection via Excel Events

```csharp
// In the AddIn or a service class
app.SheetSelectionChange += (Worksheet sheet, Range target) =>
{
    if (_isInCellSelectionMode)
    {
        var address = target.Address.Replace("$", "");
        var sheetName = sheet.Name;
        // Notify the SetupViewModel
        _cellSelectionCallback?.Invoke(new CellReference(sheetName, address));
        _isInCellSelectionMode = false;
    }
};
```

### Auto-Label Suggestion

When a cell is selected, try to find a descriptive label:
1. Check the cell directly to the left — if it contains text, use that
2. Else check the cell directly above
3. Else check for a named range that includes this cell
4. Else use the cell address as the label

### Validation

Validate in real-time as the user types distribution parameters:
- Use `DistributionFactory.Create()` in a try/catch — if it throws, show the error message
- Disable "Add Input" button until parameters are valid
- Show validation state with border colors: default → red on error → green on valid

### Scrolling

The inputs and outputs lists can grow long. Wrap them in `ScrollViewer` elements. The entire setup view should also scroll if the content exceeds the task pane height.

## File Structure

```
MonteCarlo.UI/
├── Views/
│   ├── SetupView.xaml/.cs                  # Main setup view
│   ├── InputEditorControl.xaml/.cs         # Inline input editor (cell + dist + params)
│   ├── InputCardControl.xaml/.cs           # Display card for a configured input
│   ├── OutputCardControl.xaml/.cs          # Display card for a configured output
│   └── DistributionParameterPanel.xaml/.cs # Dynamic parameter fields based on distribution type
├── ViewModels/
│   ├── SetupViewModel.cs
│   ├── InputCardViewModel.cs
│   └── OutputCardViewModel.cs
└── Services/
    └── CellSelectionService.cs             # Manages cell selection mode between Excel and WPF
```

## Commit Strategy

```
feat(ui): add SetupView layout with inputs/outputs lists and iteration config
feat(ui): add InputEditorControl with distribution picker and parameter fields
feat(ui): add cell selection service — click-to-select cells from Excel
feat(ui): add distribution preview sparkline on input cards
feat(ui): add validation and real-time parameter feedback
```

## Done When

- [ ] User can add input cells with a click-to-select workflow
- [ ] Distribution dropdown shows all 6 Phase 1 distributions
- [ ] Parameter fields update dynamically per distribution type
- [ ] Distribution preview sparkline renders on each input card
- [ ] User can add output cells
- [ ] Inputs and outputs can be removed
- [ ] Iteration count is configurable with presets
- [ ] Run button is disabled until at least 1 input + 1 output configured
- [ ] Cell highlighting applies when inputs/outputs are added
- [ ] Uses GlobalStyles consistently (colors, typography, cards, buttons)
- [ ] `dotnet build` clean
