# TASK-018: Dark Theme, Performance Optimization, and UX Polish

## Context

Read `ROADMAP.md` for full project context. This is the final polish task before the product is "v1 ready." It covers dark theme support, performance optimization for large simulations, and UX improvements across the entire task pane.

## Dependencies

- All prior tasks (this task polishes everything that's been built)

## Objective

Three workstreams in one task:
1. **Dark theme** — full dark mode for the task pane and all charts
2. **Performance** — optimize the simulation loop for 10,000+ iteration runs
3. **UX polish** — animations, validation, error states, keyboard shortcuts, empty states

---

## Part A: Dark Theme

### Color Palette (Dark)

Add to `GlobalStyles.xaml` or create `DarkTheme.xaml`:

```
Background:         #1E293B   (slate-800)
Surface:            #334155   (slate-700)
SurfaceHover:       #475569   (slate-600)
Border:             #475569   (slate-600)
TextPrimary:        #F1F5F9   (slate-100)
TextSecondary:      #94A3B8   (slate-400)
TextTertiary:       #64748B   (slate-500)

// Data colors stay the same — they need to pop on both backgrounds
Blue500:            #3B82F6   (unchanged)
Orange500:          #F97316   (unchanged)
Emerald500:         #10B981   (unchanged)
Red500:             #EF4444   (unchanged)
Amber500:           #F59E0B   (unchanged)
```

### Implementation Strategy

Use WPF `DynamicResource` instead of `StaticResource` for all theme-sensitive colors. This allows runtime switching:

```xml
<!-- In control templates, use DynamicResource -->
<Border Background="{DynamicResource BackgroundBrush}">
    <TextBlock Foreground="{DynamicResource TextPrimaryBrush}" ... />
</Border>
```

Create a `ThemeManager` service:

```csharp
public class ThemeManager
{
    public enum Theme { Light, Dark, System }

    public void ApplyTheme(Theme theme)
    {
        var dict = theme switch
        {
            Theme.Dark => new ResourceDictionary { Source = "DarkTheme.xaml" },
            Theme.Light => new ResourceDictionary { Source = "LightTheme.xaml" },
            Theme.System => DetectSystemTheme() ? DarkDict : LightDict,
            _ => LightDict
        };

        // Replace the theme dictionary in MergedDictionaries
        var app = Application.Current.Resources.MergedDictionaries;
        // Remove old theme dict, add new one
    }

    private bool DetectSystemTheme()
    {
        // Read Windows registry:
        // HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme
        // 0 = dark, 1 = light
    }
}
```

### Chart Theme Updates

- **LiveCharts2 charts:** Update axis label colors, gridline colors, and tooltip background based on theme
- **SkiaSharp tornado chart:** Pass theme colors from WPF to the SkiaSharp paint objects
- **Distribution previews:** Update stroke and fill colors

### Settings View

Add a theme selector to the SettingsView:
- Light / Dark / System (auto-detect)
- Save preference in user settings (app.config or registry, NOT in the workbook)

---

## Part B: Performance Optimization

### Target

10,000 iterations × 20 inputs should complete in < 5 seconds on a mid-range machine (for fast mode with formula recalculation).

### Optimizations

**1. Parallel sampling (already in place):**
Input sample generation should already be parallelized. Verify.

**2. Batch COM calls:**
In recalc mode, minimize COM round-trips:
```csharp
// Instead of writing one cell at a time:
foreach (var input in inputs)
    sheet.Range[input.Address].Value2 = value;  // N COM calls

// Write all inputs as a batch:
// Group by sheet, build a value array, write in one call
var range = sheet.Range[firstCell, lastCell];
range.Value2 = valueArray;  // 1 COM call
```

If inputs are non-contiguous, use `Application.Union()` to create a multi-area range, or write sheet-by-sheet.

**3. Disable Excel features during simulation:**
```csharp
app.ScreenUpdating = false;
app.EnableEvents = false;
app.Calculation = XlCalculation.xlCalculationManual;
// ... run simulation, calling app.Calculate() manually per iteration ...
app.Calculation = XlCalculation.xlCalculationAutomatic;
app.EnableEvents = true;
app.ScreenUpdating = true;
```

**4. Progress reporting throttle:**
Already specified in TASK-003, but verify: fire progress events at most every 100ms or every 100 iterations. Do NOT fire on every iteration.

**5. Memory pre-allocation:**
Verify that the input matrix `double[iterations, inputCount]` and output matrix `double[iterations, outputCount]` are allocated once upfront, not grown dynamically.

**6. Stats computation deferral:**
Don't compute SummaryStatistics or SensitivityAnalysis until the simulation is complete and the user is on the ResultsView. If the user cancels before completion, don't compute stats at all.

### Performance Benchmarking

Add a simple benchmark mode (accessible from Settings or a debug menu):
- Runs a synthetic simulation (N inputs, all Normal, trivial evaluator)
- Measures iterations/second
- Reports total time and per-iteration overhead
- Helps users understand how their model complexity affects simulation speed

---

## Part C: UX Polish

### Animations

- **View transitions:** Fade or slide when navigating between Setup → Run → Results (WPF `Storyboard` with `DoubleAnimation` on `Opacity`, 200ms duration)
- **Progress bar:** Smooth width transition (not jumpy discrete steps)
- **Chart loading:** Fade-in when chart data populates (150ms)
- **Card additions:** Slide-in when a new input/output card is added to the list

Keep all animations under 300ms. Respect `SystemParameters.MinimizeAnimation` — if the user has disabled Windows animations, skip them.

### Input Validation

Enhance all input fields across the task pane:
- **Real-time validation** as the user types (debounced 300ms)
- **Visual states:** Default border → Red border + error message on invalid → Green border on valid
- **Error messages** below the field, in red, specific: "Standard deviation must be positive" not just "Invalid input"
- **Disable submit buttons** until all fields are valid

### Error States

When something goes wrong during simulation:
- Show a clean error card in the RunView (not a raw exception message)
- Include the error type, a human-readable message, and a "Try Again" button
- Common errors to handle gracefully:
  - Cell reference no longer valid (sheet renamed, cell deleted)
  - Formula error in an output cell (#REF!, #VALUE!, #DIV/0!)
  - Out of memory (too many iterations)
  - Excel busy (another add-in or macro is running)

### Empty States

When views have no content:
- **SetupView with no inputs:** Show a friendly illustration or icon + "Add your first uncertain input to get started" + an "Add Input" button
- **ResultsView with no simulation run:** "Run a simulation to see results here"
- Don't show blank white space — always guide the user to the next action

### Keyboard Shortcuts

- `Ctrl+Shift+R` — Run simulation (from any view)
- `Ctrl+Shift+S` — Stop simulation (during run)
- `Escape` — Cancel current action (close input editor, exit cell selection mode)
- `Ctrl+Shift+T` — Toggle task pane visibility

Register these as Excel keyboard shortcuts via `Application.OnKey()` in the add-in.

### Loading Skeleton

While stats are computing after a simulation completes, show skeleton loading placeholders:
- Gray animated rectangles where the headline stat, chart, and stats panel will appear
- Replaced by real content as each piece computes

### Number Input Enhancement

For all numeric input fields (distribution parameters, iteration count, target value):
- Allow paste
- Allow scientific notation (1e6 → 1,000,000)
- Strip commas and currency symbols on input
- Show formatted value on blur (e.g., "1000000" → "1,000,000")

---

## File Structure

```
MonteCarlo.UI/
├── Styles/
│   ├── GlobalStyles.xaml            # Updated: use DynamicResource throughout
│   ├── LightTheme.xaml              # Light color definitions
│   └── DarkTheme.xaml               # Dark color definitions
├── Services/
│   └── ThemeManager.cs
├── Controls/
│   ├── SkeletonLoader.xaml/.cs      # Animated loading placeholder
│   └── ValidatedTextBox.xaml/.cs    # Text input with built-in validation states
└── Views/
    └── SettingsView.xaml/.cs        # Updated: theme selector

MonteCarlo.Charts/
└── Themes/
    └── ChartTheme.cs                # Updated: accepts light/dark colors
```

## Commit Strategy

```
feat(ui): add DarkTheme.xaml and ThemeManager with system detection
feat(ui): update all views to use DynamicResource for theme colors
feat(charts): update chart controls to accept theme-aware colors
feat(ui): add view transition animations (fade, slide)
feat(ui): add ValidatedTextBox control with real-time validation states
feat(ui): add empty states for Setup and Results views
feat(ui): add skeleton loading placeholders
feat(addin): add keyboard shortcuts for Run, Stop, and Task Pane toggle
perf(addin): optimize batch COM calls and disable Excel features during simulation
```

## Done When

- [ ] Dark theme renders correctly across all views and chart controls
- [ ] System theme auto-detection works
- [ ] Theme preference persists between sessions
- [ ] View transitions are smooth and under 300ms
- [ ] All input fields show real-time validation with error messages
- [ ] Error states during simulation display clean, actionable messages
- [ ] Empty states guide the user with clear CTAs
- [ ] Keyboard shortcuts registered and working
- [ ] 10,000 iterations × 20 inputs completes in < 5 seconds (fast mode benchmark)
- [ ] `dotnet build` clean
