# TASK-006: Excel I/O Layer

## Context

Read `ROADMAP.md` for full project context. The engine (TASK-002 through TASK-005) is pure C# with no Excel dependency. This task builds the bridge between the engine and Excel — reading cell values, writing results, highlighting tagged cells, and managing input/output tracking.

## Objective

Build the `MonteCarlo.Addin/Excel/` layer that reads from and writes to the active workbook. This is the only code in the project that touches Excel COM interop (via ExcelDna).

## Design

### WorkbookManager

The central class for all Excel read/write operations.

```csharp
public class WorkbookManager
{
    /// Read the current value of a cell. Returns the computed value (not the formula).
    public double ReadCellValue(string sheetName, string cellAddress);

    /// Read multiple cell values in a batch (minimizes COM round-trips).
    public Dictionary<string, double> ReadCellValues(IEnumerable<CellReference> cells);

    /// Write a value to a cell.
    public void WriteCellValue(string sheetName, string cellAddress, double value);

    /// Write a block of values to a range (for writing simulation results to a sheet).
    public void WriteRange(string sheetName, string topLeftAddress, double[,] values);

    /// Write column headers + data to a new or existing sheet.
    public void WriteResultsSheet(string sheetName, string[] headers, double[,] data);

    /// Get the currently selected cell's address and sheet name.
    public CellReference? GetActiveCell();

    /// Check if a sheet exists in the active workbook.
    public bool SheetExists(string sheetName);

    /// Create a new sheet (or clear an existing one).
    public void EnsureSheet(string sheetName, bool clearIfExists = true);
}

public class CellReference
{
    public string SheetName { get; set; }
    public string CellAddress { get; set; }     // e.g., "B4"
    public string FullReference => $"'{SheetName}'!{CellAddress}";
}
```

### InputTagManager

Tracks which cells the user has designated as simulation inputs.

```csharp
public class InputTagManager
{
    /// Tag a cell as a simulation input with a specific distribution.
    public void TagInput(CellReference cell, string label, IDistribution distribution);

    /// Remove a cell's input tag.
    public void UntagInput(CellReference cell);

    /// Get all tagged inputs.
    public IReadOnlyList<TaggedInput> GetAllInputs();

    /// Check if a cell is tagged.
    public bool IsTagged(CellReference cell);

    /// Convert tagged inputs to SimulationInput objects for the engine.
    public List<SimulationInput> ToSimulationInputs(WorkbookManager workbook);
}

public class TaggedInput
{
    public CellReference Cell { get; set; }
    public string Label { get; set; }
    public string DistributionName { get; set; }       // e.g., "Normal"
    public Dictionary<string, double> Parameters { get; set; }  // e.g., { "mean": 100, "stdDev": 10 }
}
```

### OutputTagManager

Tracks which cells are simulation outputs.

```csharp
public class OutputTagManager
{
    public void TagOutput(CellReference cell, string label);
    public void UntagOutput(CellReference cell);
    public IReadOnlyList<TaggedOutput> GetAllOutputs();
    public bool IsTagged(CellReference cell);
    public List<SimulationOutput> ToSimulationOutputs();
}

public class TaggedOutput
{
    public CellReference Cell { get; set; }
    public string Label { get; set; }
}
```

### CellHighlighter

Applies visual formatting to tagged cells so the user can see at a glance which cells are inputs/outputs.

```csharp
public class CellHighlighter
{
    /// Apply input highlighting (subtle blue background).
    public void HighlightInput(CellReference cell);

    /// Apply output highlighting (subtle green background).
    public void HighlightOutput(CellReference cell);

    /// Remove highlighting from a cell.
    public void ClearHighlight(CellReference cell);

    /// Refresh all highlights (call after loading a saved config).
    public void RefreshAll(InputTagManager inputs, OutputTagManager outputs);
}
```

## Implementation Notes

### ExcelDna API Access

Use `ExcelDnaUtil.Application` to get the Excel `Application` object:

```csharp
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

var app = (Application)ExcelDnaUtil.Application;
var workbook = app.ActiveWorkbook;
var sheet = (Worksheet)workbook.Sheets["Sheet1"];
var range = sheet.Range["B4"];
double value = Convert.ToDouble(range.Value2);
```

### Batching COM Calls

COM interop calls to Excel are expensive. When reading multiple cells:

```csharp
// BAD — one COM call per cell
foreach (var cell in cells)
    values[cell] = ReadCellValue(cell.SheetName, cell.CellAddress);

// GOOD — read a contiguous range in one call if possible,
// or at minimum batch by sheet
var sheet = (Worksheet)workbook.Sheets[sheetName];
// For non-contiguous cells, use a Union range or read sheet-by-sheet
```

For `WriteRange`, always write the full 2D array in one call:
```csharp
var range = sheet.Range[topLeft].Resize[rows, cols];
range.Value2 = values;  // Single COM call for the entire block
```

### Cell Highlighting Colors

Keep these subtle — the cells still need to be readable:
```
Input cells:  Light blue background (#DBEAFE / RGB 219, 234, 254)
Output cells: Light green background (#DCFCE7 / RGB 220, 252, 231)
```

Use `range.Interior.Color` to set the background. Store the original color before highlighting so `ClearHighlight` can restore it (or just set to no fill).

### Error Handling

- If a cell contains text, a formula error (#REF!, #VALUE!), or is empty, `ReadCellValue` should throw a descriptive exception: `"Cell 'Sheet1'!B4 does not contain a numeric value (current value: #REF!)"`
- If the sheet doesn't exist, throw early with the sheet name in the message
- Wrap all COM calls in try/catch to handle cases where the workbook has been closed or the cell reference is invalid

### Screen Updating

When performing batch operations (writing results, highlighting many cells):
```csharp
app.ScreenUpdating = false;
app.Calculation = XlCalculation.xlCalculationManual;
try
{
    // ... batch operations ...
}
finally
{
    app.Calculation = XlCalculation.xlCalculationAutomatic;
    app.ScreenUpdating = true;
}
```

## Tests

This layer is hard to unit test because it depends on Excel COM. The approach:

1. **Define interfaces** (`IWorkbookManager`, `IInputTagManager`, etc.) so the UI and engine integration code can be tested with mocks
2. **Manual smoke testing** with the sample workbooks once the add-in is runnable
3. **Integration tests** are deferred until we have a CI environment with Excel installed (or we use a shim)

For now, create the interfaces and implementations. Tests for this layer will be manual + implicit (exercised through end-to-end testing in later tasks).

## File Structure

```
MonteCarlo.Addin/
└── Excel/
    ├── IWorkbookManager.cs        # Interface
    ├── WorkbookManager.cs         # Implementation
    ├── CellReference.cs
    ├── IInputTagManager.cs
    ├── InputTagManager.cs
    ├── TaggedInput.cs
    ├── IOutputTagManager.cs
    ├── OutputTagManager.cs
    ├── TaggedOutput.cs
    ├── ICellHighlighter.cs
    └── CellHighlighter.cs
```

## Commit Strategy

```
feat(addin): add CellReference model and IWorkbookManager interface
feat(addin): implement WorkbookManager with batched reads/writes
feat(addin): add InputTagManager and OutputTagManager
feat(addin): add CellHighlighter with input/output color coding
```

## Done When

- [ ] WorkbookManager reads/writes cells via ExcelDna COM interop
- [ ] Batch reading and writing implemented for performance
- [ ] InputTagManager and OutputTagManager track tagged cells
- [ ] CellHighlighter applies/removes visual formatting
- [ ] All classes have corresponding interfaces for testability
- [ ] Screen updating disabled during batch operations
- [ ] `dotnet build` clean (full testing deferred to integration)
