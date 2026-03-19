# TASK-005: Chart Controls — Histogram + Tornado

**Priority**: 🟡 High
**Phase**: 1/2 — Walking Skeleton + Storytelling
**Depends On**: TASK-003 (simulation engine), TASK-004 (sensitivity analysis)
**Estimated Effort**: ~5 hours

---

## Objective

Build the two most important chart controls: the histogram (with PDF overlay and percentile markers) and the tornado chart (sensitivity). These are WPF controls in `MonteCarlo.Charts` — reusable, themed, and beautiful.

---

## Design Requirements

Refer to **ROADMAP.md Section 4 (Visual Design Language)** for the full spec. Key points:

- **Color palette**: Blue (#3B82F6) primary, orange (#F97316) for negative tornado bars, amber (#F59E0B) percentile markers, red (#EF4444) target line
- **Typography**: Segoe UI Variable / Segoe UI fallback. Tabular numerics for data labels.
- **Whitespace**: Generous padding. Let data breathe.
- **Bar radius**: 2px border radius on histogram bars, slight transparency (0.85 opacity)
- **Animations**: Smooth transitions when data updates (bars growing, curve fading in)

---

## Deliverables

### 1. ChartTheme System

```csharp
// MonteCarlo.Charts/Themes/ChartTheme.cs
public class ChartTheme
{
    public Color PrimaryData { get; }       // #3B82F6
    public Color SecondaryData { get; }     // #8B5CF6
    public Color PositiveAccent { get; }    // #10B981
    public Color NegativeAccent { get; }    // #EF4444
    public Color TornadoPositive { get; }   // #3B82F6
    public Color TornadoNegative { get; }   // #F97316
    public Color PercentileMarker { get; }  // #F59E0B
    public Color TargetLine { get; }        // #EF4444
    public Color Neutral { get; }           // #64748B
    public Color Background { get; }
    public Color Surface { get; }
    public Color Border { get; }
    public string FontFamily { get; }
    // ... etc
}

public static class ChartThemes
{
    public static ChartTheme Light { get; }
    public static ChartTheme Dark { get; }
}
```

### 2. Histogram Chart Control

A WPF `UserControl` that displays:
- Frequency histogram with configurable bin count (auto or manual)
- Optional kernel density estimate (KDE) curve overlay
- Vertical dashed lines for P10, P50, P90 with floating labels
- Optional target line (red dashed) with probability annotation: "P(X > target) = 37%"
- X-axis with intelligent tick spacing and formatted values
- Y-axis with frequency or relative frequency
- Clean gridlines at low opacity (0.15)
- Responsive to container width changes

**Input data**: `double[] samples` + optional `OutputStatistics` for pre-computed percentiles + optional `double? targetValue`.

**Rendering approach**: Use **LiveCharts2** with a `CartesianChart`. Column series for the histogram bars, line series for the KDE curve, visual elements for the percentile/target lines.

### 3. Tornado Chart Control

A custom WPF `UserControl` rendered with **SkiaSharp** (for pixel-perfect control):
- Horizontal bars extending left/right from a center baseline
- Left bars (decrease output) in orange, right bars (increase output) in blue
- Sorted by absolute impact (largest at top)
- Input labels on the left Y-axis with distribution shortcode in muted smaller text
- Value labels at bar tips showing the output range
- Center baseline labeled with the base case output value
- Responsive to container width changes

**Input data**: `List<SensitivityResult>` from the sensitivity analysis engine + `double baseCase` (the output value when all inputs are at their means).

**Rendering approach**: Custom `SKCanvasView` or `SKElement` in WPF. Direct SkiaSharp drawing gives us full control over anti-aliasing, text positioning, and bar geometry.

### 4. Distribution Preview Control

A small (`~120x40px`) inline chart that shows the shape of a distribution. Used in the Setup view next to each input.

- Renders the PDF curve filled with low-opacity primary color
- No axes, no labels — just the shape
- Lightweight enough to render 10+ simultaneously in the task pane

### 5. Unit / Visual Tests

- Verify histogram bin calculation with known data
- Verify tornado chart renders without exceptions for edge cases (single input, all sensitivities near zero)
- Snapshot test (render to image, compare against baseline) if feasible

---

## Acceptance Criteria

- [ ] `ChartTheme` system exists with Light and Dark themes
- [ ] Histogram control renders with bars, KDE overlay, percentile lines, and optional target
- [ ] Tornado control renders with sorted horizontal bars, color coding, labels
- [ ] Distribution preview renders inline PDF shape
- [ ] All controls respect the theme (colors, fonts, spacing from ChartTheme)
- [ ] Controls are responsive to width changes
- [ ] No Excel or VSTO dependencies in MonteCarlo.Charts
