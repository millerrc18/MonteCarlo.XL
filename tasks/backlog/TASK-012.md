# TASK-012: Results Export to Excel

## Context

Read `ROADMAP.md` for full project context. After a simulation, users need to get the results out of the task pane and into the workbook for sharing in presentations and reports.

## Dependencies

- TASK-004 (SummaryStatistics)
- TASK-005 (SensitivityAnalysis)
- TASK-006 (WorkbookManager)
- TASK-009 (Histogram and CDF chart controls — for image export)
- TASK-010 (Tornado chart control — for image export)

## Objective

Build an "Export to Sheet" feature that writes a formatted results summary to a new Excel sheet, including statistics tables, embedded chart images, and raw data.

## Design

### Export Content

The export creates a new sheet named `MC Results — {OutputLabel}` (e.g., "MC Results — Net Profit") with this layout:

```
┌─────────────────────────────────────────────────────────────────┐
│ A                    B                C              D          │
├─────────────────────────────────────────────────────────────────┤
│ Row 1:  MonteCarlo.XL Simulation Results                       │ ← Bold title
│ Row 2:  Output: Net Profit (Sheet1!D10)                        │
│ Row 3:  Iterations: 5,000 | Date: 2026-03-19                  │
│ Row 4:  (blank)                                                │
│ Row 5:  SUMMARY STATISTICS                                     │ ← Section header
│ Row 6:  Statistic          Value                               │
│ Row 7:  Mean               $2,450,000                          │
│ Row 8:  Median             $2,380,000                          │
│ Row 9:  Std Dev            $812,000                            │
│ Row 10: Minimum            $180,000                            │
│ Row 11: Maximum            $5,120,000                          │
│ Row 12: Skewness           0.34                                │
│ Row 13: Kurtosis           -0.12                               │
│ Row 14: (blank)                                                │
│ Row 15: PERCENTILES                                            │
│ Row 16: Percentile         Value                               │
│ Row 17: P1                 $420,000                            │
│ Row 18: P5                 $1,100,000                          │
│ ...     ...                ...                                 │
│ Row 25: P99                $4,800,000                          │
│ Row 26: (blank)                                                │
│ Row 27: SENSITIVITY (TOP 10 INPUTS)                            │
│ Row 28: Input              Rank Corr     Swing                 │
│ Row 29: Material Cost      0.82          $1,400,000            │
│ Row 30: Labor Hours        0.61          $900,000              │
│ ...                                                            │
│ Row 39: (blank)                                                │
│ Row 40: INPUT ASSUMPTIONS                                      │
│ Row 41: Input              Cell     Distribution               │
│ Row 42: Material Cost      B4       Normal(μ=100, σ=10)        │
│ ...                                                            │
│ Row 50: (blank)                                                │
│ Row 51: [Embedded Histogram Image]            (spans ~E1:L20)  │
│ Row 52: [Embedded Tornado Image]              (spans ~E22:L40) │
└─────────────────────────────────────────────────────────────────┘
```

### Chart Image Export

Render the histogram and tornado chart as high-DPI PNG images and embed them in the Excel sheet:

```csharp
// Render a WPF control to a bitmap
public static BitmapSource RenderControlToBitmap(FrameworkElement control, int width, int height, double dpi = 192)
{
    control.Measure(new Size(width, height));
    control.Arrange(new Rect(0, 0, width, height));
    control.UpdateLayout();

    var renderBitmap = new RenderTargetBitmap(
        (int)(width * dpi / 96), (int)(height * dpi / 96),
        dpi, dpi, PixelFormats.Pbgra32);
    renderBitmap.Render(control);
    return renderBitmap;
}

// Save to PNG bytes for embedding
var encoder = new PngBitmapEncoder();
encoder.Frames.Add(BitmapFrame.Create(renderBitmap));
using var stream = new MemoryStream();
encoder.Save(stream);
byte[] pngBytes = stream.ToArray();
```

For the SkiaSharp-based tornado chart, render to an `SKBitmap` and export directly:
```csharp
var bitmap = new SKBitmap(width * 2, height * 2);  // 2x for retina
var canvas = new SKCanvas(bitmap);
// ... render tornado chart ...
var image = SKImage.FromBitmap(bitmap);
var data = image.Encode(SKEncodedImageFormat.Png, 100);
byte[] pngBytes = data.ToArray();
```

Embed in Excel using:
```csharp
var sheet = (Worksheet)workbook.Sheets["MC Results — Net Profit"];
var tempPath = Path.GetTempFileName() + ".png";
File.WriteAllBytes(tempPath, pngBytes);
sheet.Shapes.AddPicture(tempPath, MsoTriState.msoFalse, MsoTriState.msoCTrue,
    left, top, width, height);
File.Delete(tempPath);
```

### Formatting

Apply professional formatting to the export sheet:
- Section headers: bold, larger font (14pt), bottom border
- Stat labels: regular weight, left-aligned
- Stat values: number-formatted, right-aligned, `Cascadia Code` or `Consolas` font
- Alternating row shading on tables (`#F8FAFC` on even rows)
- Column widths auto-fitted
- Print area set to the content region
- Sheet tab color: `#3B82F6` (blue)

### Raw Data Sheet (Optional)

Optionally create a second sheet `MC Raw Data — {OutputLabel}` with:
- Column A: Iteration number (1 to N)
- Columns B+: Input sample values (one column per input, labeled)
- Last column: Output values

This lets users run their own analysis on the raw simulation data. Only create this sheet if the user opts in (checkbox in the export dialog or a separate "Export Raw Data" button).

## ResultsExporter Service

```csharp
public class ResultsExporter
{
    private readonly IWorkbookManager _workbook;

    /// Export summary + charts to a formatted Excel sheet.
    public void ExportSummary(
        SimulationResult result,
        SummaryStatistics stats,
        SensitivityResult sensitivity,
        SimulationProfile profile,
        string outputId,
        byte[] histogramImage,
        byte[] tornadoImage
    );

    /// Export raw simulation data to a data sheet.
    public void ExportRawData(SimulationResult result, string outputId);
}
```

## File Structure

```
MonteCarlo.Addin/
└── Export/
    ├── ResultsExporter.cs          # Writes formatted results to Excel sheets
    └── ChartImageRenderer.cs       # Renders WPF/SkiaSharp controls to PNG bytes

MonteCarlo.UI/
└── Views/
    └── (Update ResultsView to wire the Export button)
```

## Commit Strategy

```
feat(addin): add ChartImageRenderer — WPF and SkiaSharp controls to PNG
feat(addin): add ResultsExporter — formatted summary sheet with stats tables
feat(addin): add chart image embedding in export sheet
feat(addin): add raw data export option
feat(ui): wire Export button in ResultsView to ResultsExporter
```

## Done When

- [ ] "Export to Sheet" creates a professionally formatted summary sheet
- [ ] Stats table includes all summary statistics and percentiles
- [ ] Sensitivity table shows top inputs with rank correlation and swing values
- [ ] Input assumptions table documents the distribution config
- [ ] Histogram and tornado chart embedded as high-DPI PNG images
- [ ] Sheet formatting is clean and presentation-ready (fonts, borders, number formats)
- [ ] Optional raw data export works
- [ ] Existing sheets with the same name are cleared and overwritten (with confirmation)
- [ ] `dotnet build` clean
