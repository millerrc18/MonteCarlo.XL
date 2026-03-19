# TASK-009: Results Dashboard — Histogram, CDF, and Stats Panel

## Context

Read `ROADMAP.md` for full project context. Specifically read **Section 4 (Visual Design Language)** — the chart specs, color palette, and design principles are requirements, not suggestions.

## Dependencies

- TASK-004 (SummaryStatistics, HistogramData)
- TASK-007 (Task pane shell, GlobalStyles)

## Objective

Build the Results view that appears after a simulation completes. This is the centerpiece of the product — the visualization that tells the story. It includes a histogram, a CDF chart (togglable), a summary stats panel, and a target line feature.

## Design

### Results View Layout

```
┌───────────────────────────────────┐
│ RESULTS          [Output ▾ dropdown]│  ← Switch between output cells
├───────────────────────────────────┤
│                                   │
│ ┌───────────────────────────────┐ │
│ │   $2.4M                      │ │  ← Headline stat (P50) — big, bold
│ │   Median Net Profit           │ │
│ │   95% CI: $1.1M — $3.8M      │ │  ← P5–P95 range
│ └───────────────────────────────┘ │
│                                   │
│ ┌───────────────────────────────┐ │
│ │                               │ │
│ │      [HISTOGRAM CHART]        │ │  ← ~200px tall
│ │                               │ │
│ │   P10 ╎     P50 ╎      P90 ╎ │ │  ← Percentile markers
│ │   ─ ─ ─ Target ─ ─ ─ ─ ─ ─ ─│ │  ← Optional target line
│ │                               │ │
│ │    [PDF ○] [CDF ●]           │ │  ← Toggle between views
│ └───────────────────────────────┘ │
│                                   │
│ ┌─── Stats─────────────────────┐  │
│ │ Mean     $2.45M   StdDev $0.8M│ │
│ │ P5       $1.1M    P95    $3.8M│ │
│ │ Min      $0.2M    Max    $5.1M│ │
│ │ Skew     0.34     Kurt   -0.12│ │
│ └───────────────────────────────┘ │
│                                   │
│ ┌───Target──────────────────────┐ │
│ │ Target: [$2,000,000    ]      │ │
│ │ P(above target): 62.3%       │ │
│ │ P(below target): 37.7%       │ │
│ └───────────────────────────────┘ │
│                                   │
│ [Export to Sheet]  [📋 Copy Stats]│
└───────────────────────────────────┘
```

## Chart Controls

### HistogramChart (MonteCarlo.Charts)

A LiveCharts2 `CartesianChart` configured for histograms.

**Visual Requirements (from ROADMAP.md Section 4):**
- Bars: fill `#3B82F6` at 0.85 opacity, no stroke, 2px border radius (if LiveCharts2 supports it; otherwise no radius is fine)
- Grid lines: `#E2E8F0` at 0.15 opacity, horizontal only
- X-axis: formatted values (auto-detect currency, %, or plain number based on magnitude), `Segoe UI Variable` 11px, color `#64748B`
- Y-axis: relative frequency (0.00–0.05 range typical), same font/color
- Background: transparent (inherits from the task pane)
- Padding: generous — at least 16px on all sides inside the chart area

**Percentile markers:**
- P10, P50, P90 as vertical dashed lines
- Color: `#F59E0B` (amber)
- Line style: dashed (4px dash, 4px gap)
- Floating label above each line showing the value (e.g., "$1.8M")
- Use LiveCharts2 `Sections` or custom paint to draw these

**Optional KDE overlay:**
- Smooth kernel density curve overlaid on the histogram
- Color: `#1E40AF` (darker blue) as a line, no fill
- Only show if user toggles it on (default off — keep the histogram clean)

**Target line:**
- User-defined threshold value
- Color: `#EF4444` (red), dashed
- Annotation text: "P(X > $2M) = 62.3%" positioned near the line
- Updated when the user changes the target value in the input field below the chart

### CDFChart (MonteCarlo.Charts)

A LiveCharts2 `CartesianChart` configured as a cumulative distribution.

**Visual Requirements:**
- Smooth S-curve using the sorted output values
- Line: `#3B82F6`, 2px weight
- Area fill below the curve: `#3B82F6` at 0.08 opacity
- Horizontal reference lines at P10, P50, P90 — dashed `#F59E0B` with value labels on the right axis
- Y-axis: 0% to 100% (probability)
- X-axis: same as histogram

**Toggle:** The histogram and CDF share the same chart area — user toggles between them with a segmented control (two small buttons: "PDF" and "CDF"). The chart animates between views if LiveCharts2 supports transitions; otherwise a clean swap is fine.

### Summary Stats Panel

A card (using the `Card` style from GlobalStyles) with a 2-column grid of stats:

| Left | Right |
|------|-------|
| Mean: $2.45M | Std Dev: $0.8M |
| P5: $1.1M | P95: $3.8M |
| Min: $0.2M | Max: $5.1M |
| Skewness: 0.34 | Kurtosis: -0.12 |

**Formatting rules:**
- Values > 1,000,000: format as "$X.XM"
- Values > 1,000: format as "$X.XK" or use comma separators
- Percentages: "XX.X%"
- Small decimals: show 2–4 significant figures
- Use the `MonoFont` (Cascadia Code) for all numeric values — they should be right-aligned and visually consistent
- Labels in `TextSecondary` color, values in `TextPrimary`

Create a smart number formatter utility that handles all of these cases based on the data range.

### Headline Stat Card

The big number at the top — this is the most important visual element:
- The P50 (median) value, displayed in `HeadlineLarge` style (24px, semibold)
- Output label below it ("Median Net Profit") in `BodyText` style
- 95% credible interval below that ("$1.1M — $3.8M") in `TextTertiary`
- Centered in a subtle card with extra vertical padding

### Target Input

A simple text field + live calculation:
- User types a target value (e.g., "2000000")
- On change (debounced 300ms): compute `ProbabilityAbove(target)` and `ProbabilityBelow(target)` from SummaryStatistics
- Display both probabilities
- Update the target line on the histogram chart

## ViewModels

### ResultsViewModel

```csharp
public partial class ResultsViewModel : ObservableObject
{
    [ObservableProperty] private SimulationResult _simulationResult;
    [ObservableProperty] private string _selectedOutputId;
    [ObservableProperty] private SummaryStatistics _currentStats;
    [ObservableProperty] private HistogramData _histogramData;
    [ObservableProperty] private bool _showCDF;            // false = histogram, true = CDF
    [ObservableProperty] private double? _targetValue;
    [ObservableProperty] private double? _probabilityAboveTarget;
    [ObservableProperty] private double? _probabilityBelowTarget;

    [ObservableProperty] private ObservableCollection<string> _availableOutputs;

    // Headline stat
    [ObservableProperty] private string _headlineValue;     // Formatted P50
    [ObservableProperty] private string _headlineLabel;     // Output label
    [ObservableProperty] private string _headlineRange;     // Formatted P5–P95

    // Stats panel values (all pre-formatted strings)
    [ObservableProperty] private string _meanFormatted;
    [ObservableProperty] private string _stdDevFormatted;
    // ... etc

    partial void OnSelectedOutputIdChanged(string value)
    {
        // Recompute stats and rebuild chart data for the new output
    }

    partial void OnTargetValueChanged(double? value)
    {
        // Recompute probability annotations
    }

    [RelayCommand] private void ToggleChartMode();
    [RelayCommand] private void ExportToSheet();
    [RelayCommand] private void CopyStatsToClipboard();
}
```

## Implementation Notes

### LiveCharts2 Setup in a Task Pane

LiveCharts2 with SkiaSharp renders via `CartesianChart` control. In the task pane:

```xml
<lvc:CartesianChart
    Series="{Binding HistogramSeries}"
    XAxes="{Binding XAxes}"
    YAxes="{Binding YAxes}"
    Sections="{Binding Sections}"
    Height="200"
    Background="Transparent"/>
```

You'll need to configure `LiveChartsSettings` once at startup:
```csharp
LiveCharts.Configure(config =>
    config.AddSkiaSharp()
          .AddDefaultMappers());
```

### Number Formatting Utility

Create `MonteCarlo.UI/Converters/NumberFormatter.cs`:

```csharp
public static class NumberFormatter
{
    public static string Format(double value, string? hint = null)
    {
        var abs = Math.Abs(value);
        if (abs >= 1_000_000) return $"${value / 1_000_000:F1}M";
        if (abs >= 1_000) return $"${value / 1_000:F1}K";
        if (abs >= 1) return $"{value:F2}";
        if (abs >= 0.01) return $"{value:F3}";
        return $"{value:G4}";
    }
}
```

Expand this to handle percentages, custom format hints, and currency detection based on context.

### Copy Stats to Clipboard

Format as a clean text table and use `Clipboard.SetText()`:
```
MonteCarlo.XL Simulation Results — Net Profit
Iterations: 5,000
──────────────────────────
Mean:        $2,450,000
Median:      $2,380,000
Std Dev:     $812,000
P5:          $1,100,000
P95:         $3,800,000
Min:         $180,000
Max:         $5,120,000
Skewness:    0.34
Kurtosis:    -0.12
```

## File Structure

```
MonteCarlo.Charts/
├── Themes/
│   └── ChartTheme.cs                   # LiveCharts2 color/font config
├── Controls/
│   ├── HistogramChart.xaml/.cs          # LiveCharts2 histogram
│   ├── CDFChart.xaml/.cs                # LiveCharts2 CDF
│   ├── DistributionPreview.xaml/.cs     # Mini sparkline (used in SetupView too)
│   └── SparklineBar.xaml/.cs            # Mini bar for stats panel (optional)

MonteCarlo.UI/
├── Views/
│   ├── ResultsView.xaml/.cs             # Main results dashboard
│   ├── HeadlineStatCard.xaml/.cs        # The big P50 number at top
│   ├── StatsPanelControl.xaml/.cs       # 2-column stats grid
│   └── TargetLineControl.xaml/.cs       # Target input + probability display
├── ViewModels/
│   └── ResultsViewModel.cs
└── Converters/
    └── NumberFormatter.cs               # Smart number formatting
```

## Commit Strategy

```
feat(charts): add ChartTheme with color palette and LiveCharts2 config
feat(charts): implement HistogramChart with percentile markers and target line
feat(charts): implement CDFChart with area fill and reference lines
feat(charts): add DistributionPreview sparkline control
feat(ui): add ResultsView with headline stat, chart, stats panel, and target input
feat(ui): add NumberFormatter for intelligent value display
feat(ui): add PDF/CDF toggle and output selector
```

## Done When

- [ ] Histogram renders with correct colors, bar styling, and transparent background
- [ ] Percentile markers (P10/P50/P90) display as dashed amber lines with value labels
- [ ] CDF chart renders with smooth curve, area fill, and reference lines
- [ ] PDF/CDF toggle switches between views
- [ ] Headline stat shows P50 in large format with credible interval
- [ ] Stats panel shows all key metrics in a clean 2-column grid
- [ ] Target line feature: user enters value → chart updates → probability displayed
- [ ] Number formatting works across orders of magnitude ($, K, M, %, raw)
- [ ] Output dropdown allows switching between multiple outputs
- [ ] Copy Stats copies a formatted text table to clipboard
- [ ] All visuals follow the design system from ROADMAP.md Section 4
- [ ] `dotnet build` clean
