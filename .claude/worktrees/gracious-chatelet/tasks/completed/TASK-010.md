# TASK-010: Tornado Chart

## Context

Read `ROADMAP.md` Section 4 for exact visual specs. The tornado chart answers the most important analytical question: "Which inputs drive the most risk in my output?" It's a horizontal bar chart sorted by impact magnitude. This is a **custom WPF control rendered with SkiaSharp**, not a LiveCharts2 chart — tornado diagrams have a specialized layout that general charting libraries handle poorly.

## Dependencies

- TASK-005 (SensitivityAnalysis → SensitivityResult, InputSensitivity)
- TASK-007 (GlobalStyles)

## Objective

Build a custom `TornadoChart` WPF control in `MonteCarlo.Charts/Controls/` that renders a tornado diagram from a `SensitivityResult`. The chart should be beautiful enough to screenshot and drop into a board presentation.

## Design

### Visual Layout

```
          ◀ Decreases output    Increases output ▶

 Material Cost  ████████████████|██████████████████████  $1.8M — $3.2M
    Labor Hours  ████████████|████████████████           $2.0M — $2.9M
  Interest Rate  ██████████|████████████                 $2.1M — $2.8M
   Sales Volume  ██████|██████████                       $2.2M — $2.7M
     Tax Rate    ████|██████                             $2.3M — $2.6M
                               |
                          Base: $2.45M
```

### Visual Specifications (from ROADMAP.md)

- **Bars extending right (increase):** `#3B82F6` (blue-500)
- **Bars extending left (decrease):** `#F97316` (orange-500)
- **Center baseline:** 1px `#E2E8F0` (border color), representing the base/mean output
- **Bar height:** ~28px per input, 6px gap between bars
- **Bar corners:** 2px radius on the outer ends only (the end away from the baseline)
- **Sorted:** Largest absolute swing at the top, smallest at the bottom
- **Top N:** Show the top 10 inputs by default. If fewer than 10, show all. If more than 10, add a "Show all N inputs" toggle.

### Labels and Annotations

- **Left axis (input labels):** Input label text, right-aligned, in `Segoe UI Variable` 12px, `#0F172A`
- **Distribution info:** Below the label in a smaller, muted font (11px `#94A3B8`): "Normal(μ=100, σ=10)"
- **Bar-tip values:** At the right end of the longest bar (or beyond), show the output range: "$1.8M — $3.2M" — the output value when that input is at P10 vs P90
- **Base value label:** Below the center baseline, "Base: $2.45M"

### Interactions

- **Hover:** When the mouse enters a bar pair, highlight it (slight brightness increase) and show a tooltip:
  ```
  Material Cost
  P10 → Output: $1.8M
  P90 → Output: $3.2M
  Swing: $1.4M (57% of total variance)
  ```
- **Click:** Optional — could navigate to highlight that input cell in Excel (nice-to-have, not required)

## Implementation — SkiaSharp WPF Control

### Why SkiaSharp, not LiveCharts2?

Tornado charts need:
- Bidirectional bars from a center baseline (not standard)
- Labels on the left axis with multi-line text (label + distribution info)
- Precise control over bar-tip annotations
- Clean hover/tooltip behavior on individual bar pairs

This is specialized enough that building it with SkiaSharp is less work than fighting a general charting library into this shape.

### WPF Control Structure

```csharp
public class TornadoChart : SKElement  // SkiaSharp.Views.WPF
{
    // Dependency Properties
    public static readonly DependencyProperty SensitivityResultProperty = ...;
    public static readonly DependencyProperty BaseOutputValueProperty = ...;
    public static readonly DependencyProperty MaxInputsToShowProperty = ...;  // Default 10
    public static readonly DependencyProperty NumberFormatProperty = ...;

    // SkiaSharp rendering
    protected override void OnPaintSurface(SKPaintSurfaceEventArgs e)
    {
        var canvas = e.Surface.Canvas;
        canvas.Clear(SKColors.Transparent);

        DrawBars(canvas, ...);
        DrawBaseline(canvas, ...);
        DrawLabels(canvas, ...);
        DrawAnnotations(canvas, ...);
    }

    // Mouse interaction for hover
    protected override void OnMouseMove(MouseEventArgs e)
    {
        // Hit-test against bar rects to determine which input is hovered
        // Trigger tooltip or highlight
    }
}
```

### Drawing Algorithm

```
1. Calculate layout:
   - Label column width: measure the widest label text (~120px typical)
   - Chart area: remaining width after label column and right margin
   - Bar area: split equally left/right from center baseline
   - Center X = labelColumnWidth + chartArea / 2

2. Calculate scale:
   - Find the largest absolute swing across all inputs
   - Scale factor = (chartArea / 2) / maxAbsoluteSwing

3. For each input (top to bottom, sorted by |swing| descending):
   y = topMargin + index * (barHeight + barGap)

   leftBarWidth  = |OutputAtInputP10 - baseValue| * scaleFactor
   rightBarWidth = |OutputAtInputP90 - baseValue| * scaleFactor

   Draw left bar (orange): Rect from (centerX - leftBarWidth, y) to (centerX, y + barHeight)
   Draw right bar (blue):  Rect from (centerX, y) to (centerX + rightBarWidth, y + barHeight)

4. Draw center baseline: vertical line at centerX, full height
5. Draw labels: right-aligned text at (labelColumnWidth - padding, y + barHeight/2)
6. Draw annotations: bar-tip values beyond the longest bar
```

### Color Constants (SkiaSharp)

```csharp
private static readonly SKColor BarIncreaseColor = SKColor.Parse("#3B82F6");  // Blue
private static readonly SKColor BarDecreaseColor = SKColor.Parse("#F97316");  // Orange
private static readonly SKColor BaselineColor = SKColor.Parse("#E2E8F0");
private static readonly SKColor LabelColor = SKColor.Parse("#0F172A");
private static readonly SKColor SubLabelColor = SKColor.Parse("#94A3B8");
private static readonly SKColor AnnotationColor = SKColor.Parse("#64748B");
```

### Font Loading

```csharp
// Use Segoe UI Variable (Win 11) or Segoe UI (Win 10 fallback)
var typeface = SKTypeface.FromFamilyName("Segoe UI Variable") ??
               SKTypeface.FromFamilyName("Segoe UI") ??
               SKTypeface.Default;

var labelPaint = new SKPaint
{
    Typeface = typeface,
    TextSize = 12,
    Color = LabelColor,
    IsAntialias = true
};
```

## Data Binding

The control accepts a `SensitivityResult` and a base output value:

```xml
<charts:TornadoChart
    SensitivityResult="{Binding TornadoData}"
    BaseOutputValue="{Binding BaseOutputValue}"
    MaxInputsToShow="10"
    NumberFormat="{Binding NumberFormatter}"
    Height="300"
    Margin="0,8"/>
```

The chart auto-sizes vertically based on the number of inputs: `Height = topMargin + bottomMargin + inputCount * (barHeight + barGap)`. Provide a `MinHeight` and `MaxHeight` with scrolling if there are many inputs.

## File Structure

```
MonteCarlo.Charts/
└── Controls/
    └── TornadoChart.cs              # Custom SkiaSharp WPF control (no XAML needed)

MonteCarlo.Charts.Tests/             # Optional
└── TornadoChartLayoutTests.cs       # Test the layout math (positions, scaling) without rendering
```

## Commit Strategy

```
feat(charts): add TornadoChart SkiaSharp control — layout engine and rendering
feat(charts): add tornado hover interaction and tooltips
feat(charts): add bar-tip annotations and base value label
```

## Done When

- [ ] Tornado chart renders bars sorted by absolute impact
- [ ] Left bars are orange (decrease), right bars are blue (increase)
- [ ] Center baseline with base value label
- [ ] Input labels on the left axis with distribution info in muted text below
- [ ] Bar-tip annotations show output range values
- [ ] Top 10 inputs shown by default with "show all" toggle
- [ ] Hover highlights a bar pair and shows a tooltip
- [ ] Chart resizes correctly when the task pane width changes
- [ ] Visual quality matches the design spec in ROADMAP.md Section 4
- [ ] `dotnet build` clean
