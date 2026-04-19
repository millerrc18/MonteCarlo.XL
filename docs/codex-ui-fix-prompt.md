# Codex Prompt: Complete UI Support for 5 New Distributions in MonteCarlo.XL

## Project

MonteCarlo.XL is an Excel add-in for Monte Carlo simulation, built with ExcelDna 1.8 + .NET 8 + WPF + CommunityToolkit.Mvvm. The project root is `C:\Users\mille\OneDrive\04 - Projects\MonteCarlo.XL`.

The core simulation engine (`MonteCarlo.Engine`) already supports 15 probability distributions via `IDistribution` implementations registered in `DistributionFactory`. Excel UDFs for all 15 are in `src/MonteCarlo.Addin/UDF/MonteCarloFunctions.cs`. The formula auto-detection (`MCFunctionScanner`) also handles all 15.

**The problem:** The WPF task pane Setup UI (`SetupViewModel` + `DistributionParameterPanel.xaml`) only supports parameter entry for the original 10 distributions. The 5 newer distributions — **Gamma, Logistic, GEV, Binomial, Geometric** — can be selected from the dropdown but:
1. No parameter input fields appear for them
2. Confirming the add-input flow throws because `BuildParameterDictionary` has no case for them
3. Auto-detected formulas for these distributions don't populate parameter fields
4. `ResetInputEditor` doesn't reset their parameters

## Your task

Make the 5 new distributions fully usable from the task pane UI: user can pick them from the dropdown, enter parameters, preview the distribution shape, and confirm — same as the original 10.

## Files to modify

### 1. `src/MonteCarlo.UI/ViewModels/SetupViewModel.cs`

Add `[ObservableProperty]` fields for the new parameters (follow the existing pattern around line 85-112). Each property uses `private string _paramX = "default"` and the source generator creates `ParamX`.

Required new parameters:
- `_paramRateGamma = "2"` — Gamma rate (shape reuses existing `_paramShape`)
- `_paramS = "1"` — Logistic scale (mu reuses existing `_paramMu`)
- `_paramXi = "0"` — GEV shape (mu reuses existing `_paramMu`, sigma reuses existing `_paramSigma`)
- `_paramN = "10"` — Binomial trials (integer)
- `_paramP = "0.5"` — Binomial/Geometric probability

Add cases to `BuildParameterDictionary` (around line 396-437):
```csharp
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
```

Add cases to `ApplyDetectedDistribution` (same file, added recently). Follow the existing case pattern:
```csharp
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
```

Add resets to `ResetInputEditor` (near line 450-470) after the existing param resets:
```csharp
ParamRateGamma = "2";
ParamS = "1";
ParamXi = "0";
ParamN = "10";
ParamP = "0.5";
```

### 2. `src/MonteCarlo.UI/Views/DistributionParameterPanel.xaml` (and its code-behind)

Read this file to understand the existing pattern — it uses `DataTrigger`s on `SelectedDistribution` to show/hide parameter input rows. Each distribution has a `StackPanel` visible only when selected.

Add five new StackPanels (one per new distribution) following the exact pattern of the existing panels. Label text should match the parameter names in the engine:
- **Gamma**: "Shape (k)" bound to `ParamShape`, "Rate (β)" bound to `ParamRateGamma`
- **Logistic**: "Location (μ)" bound to `ParamMu`, "Scale (s)" bound to `ParamS`
- **GEV**: "Location (μ)" bound to `ParamMu`, "Scale (σ)" bound to `ParamSigma`, "Shape (ξ)" bound to `ParamXi` (with tooltip: "ξ=0 Gumbel, ξ>0 Fréchet, ξ<0 Weibull")
- **Binomial**: "Trials (n)" bound to `ParamN`, "Probability (p)" bound to `ParamP`
- **Geometric**: "Probability (p)" bound to `ParamP` (tooltip: "Number of trials until first success")

Use the existing `DataTrigger` pattern:
```xml
<StackPanel>
    <StackPanel.Style>
        <Style TargetType="StackPanel">
            <Setter Property="Visibility" Value="Collapsed"/>
            <Style.Triggers>
                <DataTrigger Binding="{Binding SelectedDistribution}" Value="Gamma">
                    <Setter Property="Visibility" Value="Visible"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </StackPanel.Style>
    <!-- parameter inputs -->
</StackPanel>
```

### 3. Verify `InputCardViewModel.ComputePreviewPointsStatic` handles all 15 distributions

Read `src/MonteCarlo.UI/ViewModels/InputCardViewModel.cs`. The `ComputePreviewPointsStatic` method likely generates PDF curve points via sampling or explicit PDF. For discrete distributions (Binomial, Geometric, Poisson), it should produce a step function or bar values, not a continuous curve. If this method doesn't handle the new distributions correctly, fix it — but only if broken. Test it by selecting each new distribution in the dropdown and verifying the preview chart renders.

## Testing

### Build
```bash
cd "C:\Users\mille\OneDrive\04 - Projects\MonteCarlo.XL"
dotnet build src/MonteCarlo.Addin/MonteCarlo.Addin.csproj -c Release
```
Expected: 0 errors, 0 warnings.

### Unit tests
```bash
dotnet test tests/MonteCarlo.Engine.Tests -v n
```
Expected: 408 tests passing, 0 failures. (No new engine tests needed — this is a UI-only fix.)

### Manual verification

1. Close any running Excel instance: `Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force`
2. Deploy:
   ```powershell
   Copy-Item "src/MonteCarlo.Addin/bin/Release/net8.0-windows/publish/MonteCarlo.Addin-AddIn64-packed.xll" "$env:APPDATA/Microsoft/Excel/XLSTART/MonteCarlo.XL.xll" -Force
   ```
3. Open Excel, open `samples/Product Launch Model.xlsx`
4. Click ribbon "Add Input" → click cell `B12` (contains `=MC.Binomial(1, 0.3)`)
5. Verify the setup form shows:
   - Distribution dropdown = "Binomial"
   - Trials (n) = 1
   - Probability (p) = 0.3
   - Preview chart renders
6. Confirm — no exception, input card appears in the list
7. Repeat for each of the 5 new distributions using manual dropdown selection with custom parameters

### Log check
If anything fails, check: `C:\Users\mille\AppData\Local\MonteCarlo.XL\Logs\startup.log`

## Constraints

- **Follow existing patterns exactly.** The SetupViewModel and XAML already have consistent conventions. Don't refactor; just extend.
- **Don't touch the engine.** Distribution classes and the factory are already complete and tested.
- **Don't break existing behavior.** All 10 original distributions must continue to work identically.
- **Reuse existing param properties when sensible.** Gamma reuses `ParamShape`, Logistic reuses `ParamMu`, GEV reuses `ParamMu`+`ParamSigma`. This matches the engine's parameter naming.
- **Single commit.** Scope the change to this one logical unit.

## Commit message
```
feat(ui): parameter inputs for Gamma, Logistic, GEV, Binomial, Geometric
```

## Reference files (already correct — do not modify)

- `src/MonteCarlo.Engine/Distributions/GammaDistribution.cs` (and the other 4)
- `src/MonteCarlo.Engine/Distributions/DistributionFactory.cs` — registration is correct
- `src/MonteCarlo.Addin/UDF/MonteCarloFunctions.cs` — UDFs exist
- `src/MonteCarlo.Addin/UDF/MCFunctionScanner.cs` — ParameterMapping has all 15
- `src/MonteCarlo.Addin/TaskPane/TaskPaneIntegration.cs` — auto-detect logic works
