# TASK-017: Custom Excel Functions — =MC.Normal(), =MC.PERT(), etc.

## Context

Read `ROADMAP.md` Section 7 for the input tagging design. TASK-008 built the GUI-driven input tagging workflow. This task adds the power-user alternative: type distribution functions directly in cells. This is how @RISK's `=RiskNormal()` functions work.

## Dependencies

- TASK-002 (DistributionFactory, all distributions)
- TASK-003 (SimulationEngine — detects MC function cells as inputs)
- TASK-013 (SimulationOrchestrator — updated to scan for MC functions)

## Objective

Register custom Excel functions (UDFs) via ExcelDna that let users type distributions directly in cells. In normal mode, the functions return the expected value. During simulation, the engine detects these cells and samples from the specified distribution.

## Design

### Function Signatures

```
=MC.Normal(mean, stdDev)
=MC.Triangular(min, mode, max)
=MC.PERT(min, mode, max)
=MC.Lognormal(mu, sigma)
=MC.Uniform(min, max)
=MC.Beta(alpha, beta)
=MC.Weibull(shape, scale)
=MC.Exponential(rate)
=MC.Poisson(lambda)
```

### Behavior

**In normal mode (no simulation running):**
Each function returns the distribution's expected value so the spreadsheet works as a regular model:

| Function | Returns |
|----------|---------|
| MC.Normal(100, 10) | 100 (mean) |
| MC.Triangular(50, 100, 150) | 100 (mode) |
| MC.PERT(50, 100, 150) | 100 (mode) |
| MC.Lognormal(4.6, 0.3) | ~104.5 (exp(μ + σ²/2)) |
| MC.Uniform(0, 100) | 50 (midpoint) |
| MC.Beta(2, 5) | ~0.286 (α/(α+β)) |
| MC.Weibull(2, 100) | ~88.6 (scale × Γ(1 + 1/shape)) |
| MC.Exponential(0.5) | 2.0 (1/λ) |
| MC.Poisson(4.5) | 4.5 (λ) |

**During simulation (for each iteration):**
The SimulationOrchestrator detects cells containing MC.* formulas, creates the appropriate IDistribution, and treats them as simulation inputs. The evaluator writes the sampled value directly to the cell (overriding the formula temporarily), then restores the formula after the simulation completes.

### ExcelDna UDF Registration

```csharp
using ExcelDna.Integration;

public static class MonteCarloFunctions
{
    [ExcelFunction(
        Name = "MC.Normal",
        Description = "Normal distribution. Returns the mean in normal mode; samples during simulation.",
        Category = "MonteCarlo.XL")]
    public static object MCNormal(
        [ExcelArgument(Name = "mean", Description = "Mean (μ)")] double mean,
        [ExcelArgument(Name = "stdDev", Description = "Standard deviation (σ)")] double stdDev)
    {
        if (stdDev <= 0)
            return ExcelError.ExcelErrorValue;

        // During simulation, the orchestrator overrides this cell's value
        // In normal mode, return the expected value
        return mean;
    }

    [ExcelFunction(
        Name = "MC.Triangular",
        Description = "Triangular distribution. Returns the mode in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCTriangular(
        [ExcelArgument(Name = "min", Description = "Minimum")] double min,
        [ExcelArgument(Name = "mode", Description = "Most likely")] double mode,
        [ExcelArgument(Name = "max", Description = "Maximum")] double max)
    {
        if (min >= mode || mode >= max)
            return ExcelError.ExcelErrorValue;

        return mode;
    }

    [ExcelFunction(
        Name = "MC.PERT",
        Description = "PERT distribution. Returns the mode in normal mode.",
        Category = "MonteCarlo.XL")]
    public static object MCPERT(
        [ExcelArgument(Name = "min", Description = "Minimum")] double min,
        [ExcelArgument(Name = "mode", Description = "Most likely")] double mode,
        [ExcelArgument(Name = "max", Description = "Maximum")] double max)
    {
        if (min >= mode || mode >= max)
            return ExcelError.ExcelErrorValue;

        return mode;
    }

    // ... same pattern for all distributions ...
}
```

### UDF Auto-Detection

When the simulation starts, the orchestrator needs to find all cells containing MC.* formulas. Two approaches:

**Approach A: Scan formulas (simpler, recommended for Phase 3)**

Before each simulation, scan the used range for cells whose formula starts with `=MC.`:

```csharp
public List<DetectedMCFunction> ScanForMCFunctions(Worksheet sheet)
{
    var usedRange = sheet.UsedRange;
    var results = new List<DetectedMCFunction>();

    foreach (Range cell in usedRange)
    {
        if (cell.HasFormula && cell.Formula.ToString().StartsWith("=MC."))
        {
            var parsed = ParseMCFormula(cell.Formula.ToString());
            results.Add(new DetectedMCFunction
            {
                Cell = new CellReference(sheet.Name, cell.Address.Replace("$", "")),
                DistributionName = parsed.Name,
                Parameters = parsed.Parameters
            });
        }
    }

    return results;
}
```

**Approach B: Registration during UDF execution (more sophisticated)**

Each time an MC.* function executes, register the calling cell in a global dictionary:

```csharp
[ExcelFunction(Name = "MC.Normal", ...)]
public static object MCNormal(double mean, double stdDev)
{
    var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
    if (caller != null)
    {
        MCFunctionRegistry.Register(caller, "Normal", new { mean, stdDev });
    }
    return mean;
}
```

Approach A is simpler and more reliable. Use Approach A.

### Formula Parsing

Parse `=MC.Normal(100, 10)` into a distribution name and parameters:

```csharp
public class ParsedMCFormula
{
    public string Name { get; set; }               // "Normal"
    public Dictionary<string, double> Parameters { get; set; }
}

// Parser maps positional args to named params based on function name:
// MC.Normal(arg1, arg2) → { mean: arg1, stdDev: arg2 }
// MC.Triangular(arg1, arg2, arg3) → { min: arg1, mode: arg2, max: arg3 }
```

Handle cell references as arguments: `=MC.Normal(B2, B3)` — the formula contains cell references, not literal numbers. When scanning, read the cell's **current value** (not formula) for the parameters, or resolve the references.

**Simplification:** Since we read the cell's current computed value anyway (via `cell.Value2`), and the formula's arguments are evaluated by Excel before we scan, we can just read the parameter values from the function's result context. Actually — the simplest approach is: during scan, use `XlCall.Excel(XlCall.xlfEvaluate, formula)` or parse the formula and evaluate each argument.

The most pragmatic approach: when scanning, record the cell address and the formula text. When building SimulationInputs, re-parse the formula to extract the distribution name, then use the cell's current parameter values (which Excel has already computed). For literal arguments, this is straightforward. For cell-reference arguments, read those cells.

## Integration with Orchestrator

Update `SimulationOrchestrator.RunSimulationAsync()`:

```csharp
// After loading tagged inputs from InputTagManager:
var mcFunctions = ScanForMCFunctions(activeSheet);
foreach (var func in mcFunctions)
{
    // Skip if already tagged via the GUI (avoid duplicates)
    if (!inputManager.IsTagged(func.Cell))
    {
        var dist = DistributionFactory.Create(func.DistributionName, func.Parameters);
        var baseValue = workbook.ReadCellValue(func.Cell.SheetName, func.Cell.CellAddress);
        config.Inputs.Add(new SimulationInput
        {
            Id = func.Cell.FullReference,
            Label = func.Cell.CellAddress,  // Or auto-detect from adjacent cells
            Distribution = dist,
            BaseValue = baseValue
        });
    }
}
```

### Formula Restoration

During simulation, input cell formulas are temporarily replaced with sampled values. After simulation (or cancellation), **restore the original formulas**:

```csharp
// Before simulation: save formulas
var originalFormulas = new Dictionary<string, string>();
foreach (var input in mcFunctionInputs)
{
    var cell = sheet.Range[input.Cell.CellAddress];
    originalFormulas[input.Cell.FullReference] = cell.Formula.ToString();
}

// After simulation: restore formulas
foreach (var (cellRef, formula) in originalFormulas)
{
    var cell = sheet.Range[...];
    cell.Formula = formula;
}
```

## File Structure

```
MonteCarlo.Addin/
├── UDF/
│   ├── MonteCarloFunctions.cs       # All MC.* Excel functions
│   └── MCFunctionScanner.cs         # Scans worksheets for MC.* formulas
└── Services/
    └── SimulationOrchestrator.cs    # Updated: scan for MC functions before running
```

## Tests

### MonteCarloFunctionsTests

Unit tests for the return values (these don't need Excel):

1. MCNormal(100, 10) returns 100
2. MCTriangular(50, 100, 150) returns 100
3. MCNormal(100, -10) returns ExcelError (invalid σ)
4. MCTriangular(100, 50, 150) returns ExcelError (min > mode)

### MCFunctionScannerTests

Manual testing (requires Excel):

1. Workbook with MC.Normal in B4, MC.PERT in B5 → scanner finds both
2. Workbook with no MC functions → scanner returns empty list
3. MC.Normal(B2, B3) with cell references → correctly resolves parameter values

## Commit Strategy

```
feat(addin): register MC.Normal, MC.Triangular, MC.PERT Excel functions via ExcelDna
feat(addin): register MC.Lognormal, MC.Uniform, MC.Beta, MC.Weibull, MC.Exponential, MC.Poisson
feat(addin): add MCFunctionScanner to detect MC.* formulas in worksheets
feat(addin): integrate MC function detection into SimulationOrchestrator
feat(addin): add formula save/restore for MC function cells during simulation
```

## Done When

- [ ] All 9 MC.* functions registered and visible in Excel's function wizard
- [ ] Functions return expected values in normal mode
- [ ] Functions return #VALUE! for invalid parameters
- [ ] Scanner detects all MC.* cells in a worksheet
- [ ] Detected MC functions are automatically included as simulation inputs
- [ ] Original formulas restored after simulation completes or is cancelled
- [ ] GUI-tagged inputs and MC.* function inputs coexist without duplicates
- [ ] `dotnet build` clean
