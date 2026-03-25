# MonteCarlo.XL — Project Roadmap

> A powerful, free Excel add-in for Monte Carlo simulation, built with VSTO (C#) and WPF. Inspired by Palisade's @RISK — delivering the core 20% of features that cover 80% of real-world use, with clean, modern visualizations that make the story jump off the screen.

---

## 1. Why This Exists

Palisade's @RISK is the gold standard for Monte Carlo simulation in Excel, but it costs $1,500–$5,000+ per seat. Most analysts use fewer than 20% of its features: distributions in cells, a simulation engine, histograms, tornado charts, and basic correlation. MonteCarlo.XL delivers that core feature set as a free, Windows-native Excel add-in — with a modern visual design language that @RISK's dated UI can't match.

---

## 2. Tech Stack

| Layer | Choice | Rationale |
|-------|--------|-----------|
| **Platform** | VSTO Add-in (.NET Framework / .NET 8+) | Full COM interop with Excel; native Windows performance; access to WPF for visuals |
| **Language** | C# | Strong typing for math-heavy code; rich ecosystem; first-class VSTO support |
| **UI Framework** | WPF (Custom Task Pane via ElementHost) | Hardware-accelerated rendering; full control over layout, animation, and styling; modern flat design achievable out of the box |
| **Charting** | LiveCharts2 (primary) + custom WPF controls | LiveCharts2 is open-source, GPU-accelerated, MVVM-native, and produces beautiful modern charts. Custom WPF/SkiaSharp controls for tornado diagrams and specialized visuals |
| **Distributions** | Math.NET Numerics | Industry-standard .NET library; 40+ distributions with PDF/CDF/quantile/sampling; thoroughly tested |
| **Correlation** | Iman-Conover (custom implementation using Math.NET for matrix ops) | Industry standard for rank correlation in Monte Carlo; same approach @RISK uses |
| **State / MVVM** | CommunityToolkit.Mvvm | Lightweight, source-generated MVVM toolkit from Microsoft; clean separation of engine and UI |
| **Testing** | xUnit + FluentAssertions + Moq | Standard .NET test stack; fast engine unit tests, mockable Excel interop layer |
| **Build / CI** | MSBuild + GitHub Actions | Automated build, test, and artifact publishing |

### Why VSTO over Office.js?

We're going Windows-only, which unlocks massive advantages:

- **Performance**: Direct COM interop means we can drive Excel's recalculation engine natively. Reading/writing cells is orders of magnitude faster than Office.js REST calls.
- **WPF visuals**: We get hardware-accelerated, fully customizable UI — gradients, animations, custom chart controls, proper typography. Office.js task panes are limited to what a browser can render in a narrow iframe.
- **Threading**: We can run simulations on background threads with full progress reporting, keeping Excel responsive. Office.js is single-threaded.
- **Excel integration depth**: Ribbon customization, right-click context menus, custom cell formatting, worksheet event hooks — all available through VSTO.

The tradeoff is no Excel Online support, which we've accepted.

---

## 3. Architecture

```
MonteCarlo.XL/
├── MonteCarlo.XL.sln
│
├── src/
│   ├── MonteCarlo.Engine/                  # Pure C# class library — ZERO Excel dependency
│   │   ├── Distributions/
│   │   │   ├── IDistribution.cs            # Interface: Sample(), Mean, Percentile(), PDF(), CDF()
│   │   │   ├── NormalDistribution.cs        # Wraps Math.NET Normal
│   │   │   ├── TriangularDistribution.cs
│   │   │   ├── PERTDistribution.cs
│   │   │   ├── LognormalDistribution.cs
│   │   │   ├── UniformDistribution.cs
│   │   │   ├── DiscreteDistribution.cs
│   │   │   └── DistributionFactory.cs       # Create distribution by name + params
│   │   ├── Simulation/
│   │   │   ├── SimulationEngine.cs          # Core Monte Carlo loop
│   │   │   ├── SimulationConfig.cs          # Inputs, outputs, iteration count, seed
│   │   │   ├── SimulationResult.cs          # Raw results matrix + metadata
│   │   │   └── ConvergenceChecker.cs        # Monitors stat stability across iterations
│   │   ├── Correlation/
│   │   │   ├── CorrelationMatrix.cs         # Defines input dependencies
│   │   │   └── ImanConover.cs               # Iman-Conover rank correlation algorithm
│   │   ├── Analysis/
│   │   │   ├── SummaryStatistics.cs         # Mean, median, std dev, percentiles, CI
│   │   │   ├── SensitivityAnalysis.cs       # Regression coefficients + rank correlation
│   │   │   └── ScenarioAnalysis.cs          # P(X > target), conditional stats
│   │   └── MonteCarlo.Engine.csproj
│   │
│   ├── MonteCarlo.Charts/                   # WPF chart controls — reusable, no Excel dependency
│   │   ├── Themes/
│   │   │   ├── ChartTheme.cs                # Color palette, fonts, spacing constants
│   │   │   ├── DarkTheme.cs
│   │   │   └── LightTheme.cs
│   │   ├── Controls/
│   │   │   ├── HistogramChart.xaml/.cs       # Frequency histogram + optional PDF overlay
│   │   │   ├── CDFChart.xaml/.cs             # Cumulative distribution with percentile markers
│   │   │   ├── TornadoChart.xaml/.cs         # Custom WPF control — horizontal sensitivity bars
│   │   │   ├── DistributionPreview.xaml/.cs  # Small inline preview of a distribution shape
│   │   │   └── SparklineBar.xaml/.cs         # Mini stat bars for the summary panel
│   │   └── MonteCarlo.Charts.csproj
│   │
│   ├── MonteCarlo.Addin/                    # VSTO Add-in project — the Excel integration layer
│   │   ├── ThisAddIn.cs                     # Add-in entry point, lifecycle
│   │   ├── Ribbon/
│   │   │   ├── MonteCarloRibbon.xml         # Ribbon XML (custom tab with simulation controls)
│   │   │   └── MonteCarloRibbon.cs          # Ribbon callbacks
│   │   ├── TaskPane/
│   │   │   ├── SimulationTaskPane.cs        # WPF task pane host (ElementHost)
│   │   │   └── TaskPaneController.cs        # Show/hide, state management
│   │   ├── Excel/
│   │   │   ├── WorkbookManager.cs           # Read/write cells, named ranges, events
│   │   │   ├── InputTagManager.cs           # Track which cells are simulation inputs
│   │   │   ├── OutputTagManager.cs          # Track which cells are simulation outputs
│   │   │   ├── ConfigPersistence.cs         # Save/load sim config in CustomXMLParts
│   │   │   └── CellHighlighter.cs           # Color-code input/output cells in the sheet
│   │   ├── UDF/
│   │   │   └── MonteCarloFunctions.cs       # (Phase 3) Excel UDFs: =MC.Normal(100, 10)
│   │   └── MonteCarlo.Addin.csproj
│   │
│   └── MonteCarlo.UI/                       # WPF views for the task pane
│       ├── Views/
│       │   ├── SetupView.xaml/.cs            # Configure inputs, distributions, outputs
│       │   ├── RunView.xaml/.cs              # Progress bar, iteration counter, live stats
│       │   ├── ResultsView.xaml/.cs          # Dashboard: histogram + stats + tornado
│       │   ├── CorrelationView.xaml/.cs      # (Phase 3) Correlation matrix editor
│       │   └── SettingsView.xaml/.cs         # Iteration count, seed, theme, export options
│       ├── ViewModels/
│       │   ├── SetupViewModel.cs
│       │   ├── RunViewModel.cs
│       │   ├── ResultsViewModel.cs
│       │   └── CorrelationViewModel.cs
│       ├── Converters/                       # WPF value converters
│       ├── Styles/
│       │   ├── GlobalStyles.xaml             # Shared control styles, typography, colors
│       │   └── ChartStyles.xaml              # Chart-specific theming
│       └── MonteCarlo.UI.csproj
│
├── tests/
│   ├── MonteCarlo.Engine.Tests/             # Unit tests for simulation engine
│   │   ├── Distributions/                   # Validate sampling, PDF, CDF against known values
│   │   ├── Simulation/                      # Engine correctness, convergence
│   │   ├── Correlation/                     # Iman-Conover validation against published examples
│   │   └── Analysis/                        # Summary stats, sensitivity coefficients
│   └── MonteCarlo.Charts.Tests/             # Snapshot tests for chart rendering
│
├── samples/
│   ├── SimpleFinancialModel.xlsx            # Basic revenue/cost model for testing
│   ├── ProjectScheduleRisk.xlsx             # Duration estimation with PERT distributions
│   └── PortfolioReturns.xlsx                # Correlated asset returns (Phase 3)
│
├── docs/
│   ├── USER_GUIDE.md
│   ├── ARCHITECTURE.md
│   └── DESIGN_SYSTEM.md                     # Chart design language, color palette, typography
│
├── .github/
│   └── workflows/
│       └── build.yml                        # CI: build + test on push/PR
│
├── ROADMAP.md                               # ← You are here
├── LICENSE
└── README.md
```

### Key Architectural Decisions

**1. Engine is a pure C# class library with zero Excel dependency.**
`MonteCarlo.Engine` has no reference to Excel interop or VSTO assemblies. This means we can unit test the entire simulation pipeline at full speed with xUnit, and the engine is portable — could drive a console app, web API, or WPF desktop app tomorrow without changes.

**2. Charts are a separate WPF control library.**
`MonteCarlo.Charts` depends only on `MonteCarlo.Engine` (for data types). This makes the chart controls reusable and independently testable. The design system (colors, fonts, spacing) is defined in theme classes, not scattered across XAML files.

**3. Excel is an I/O boundary with two modes.**

- **Fast mode (default):** Read input cell values once, run the full simulation in-memory, write results to a dedicated output sheet. This is fast and works for models where inputs are simple values.
- **Recalc mode (Phase 2):** For models where output cells contain formulas that depend on input cells, we swap input values → trigger `Application.Calculate()` → read outputs → repeat. This is slower but handles complex formula chains. The user chooses the mode.

**4. Configuration persists inside the workbook.**
Simulation configs (which cells, which distributions, correlation matrix) are stored as a `CustomXMLPart` inside the .xlsx file. This means the config travels with the workbook — share the file, and the recipient has the simulation setup ready to go.

---

## 4. Visual Design Language

This is a first-class concern, not an afterthought. The charts need to look like they belong in a modern analytics product — think Linear, Stripe Dashboard, or Observable — not like default Excel charts.

### Design Principles

1. **Generous whitespace.** Let the data breathe. No cramped labels or cluttered legends.
2. **Muted base palette, vivid accents.** Grays and soft backgrounds for structure; one or two saturated colors for the data itself.
3. **Typography matters.** Use Segoe UI Variable (ships with Windows 11) or Segoe UI as the fallback. Numeric values in tabular/monospace alignment.
4. **Smooth interactions.** Animated transitions when switching between views or updating charts. Subtle hover states on chart elements.
5. **Information hierarchy.** The single most important number (e.g., P50 estimate) should be visually dominant. Supporting stats are secondary.

### Color Palette

```
Primary data:       #3B82F6  (blue-500)     — histograms, primary series
Secondary data:     #8B5CF6  (violet-500)    — CDF overlay, secondary series
Positive accent:    #10B981  (emerald-500)   — "good" outcomes, upside
Negative accent:    #EF4444  (red-500)       — "bad" outcomes, downside
Neutral:            #64748B  (slate-500)     — axes, gridlines, labels

Background:         #FFFFFF  (light theme)   / #1E293B (dark theme)
Surface:            #F8FAFC  (light)         / #334155 (dark)
Border:             #E2E8F0  (light)         / #475569 (dark)

Tornado positive:   #3B82F6  (blue)          — inputs that increase output
Tornado negative:   #F97316  (orange-500)    — inputs that decrease output

Percentile markers: #F59E0B  (amber-500)     — P10/P90 lines on histogram
Target line:        #EF4444  (red, dashed)   — user-defined target threshold
```

### Chart Specifications

**Histogram:**
- Soft-radius bars (2px border radius) with slight transparency (0.85 opacity)
- Optional smooth kernel density curve overlaid in a darker shade
- Percentile markers (P10, P50, P90) as vertical dashed lines with floating labels
- User-defined target line in red with probability annotation ("P(x > target) = 37%")
- X-axis: formatted values with intelligent tick spacing
- Y-axis: frequency or relative frequency, clean gridlines at 0.15 opacity

**Tornado Chart:**
- Horizontal bars extending left (decrease) and right (increase) from a center baseline
- Bars sorted by absolute impact (largest at top)
- Left bars in orange, right bars in blue — visually distinct without being garish
- Input labels on the left axis with distribution info in a smaller, muted font
- Value labels at bar tips showing the output range when that input swings

**CDF Chart:**
- Smooth S-curve with area fill at 0.1 opacity below the curve
- Interactive: hover to see "P(X ≤ value) = %" tooltip
- Horizontal reference lines at P10, P50, P90 with labeled values

**Summary Stats Panel:**
- Card-based layout with the headline stat (P50 or Mean) as a large number
- Supporting stats (P10, P90, Std Dev, Min, Max) in a clean grid below
- Mini sparkline-style bars showing the distribution shape inline

### Charting Technology

**LiveCharts2** for histogram and CDF charts — it's WPF-native, GPU-accelerated, supports MVVM binding, and produces clean output with proper theming. We'll use custom `WPF UserControls` with **SkiaSharp** rendering for the tornado chart, since tornado diagrams are a specialized layout that general charting libraries handle poorly. SkiaSharp gives us pixel-perfect control with hardware acceleration.

---

## 5. Phased Roadmap

### Phase 1 — Walking Skeleton (Weeks 1–3)

**Goal:** End-to-end simulation working. User tags inputs, runs a sim, sees a beautiful histogram in the task pane.

| Task | Details | Est. |
|------|---------|------|
| Project scaffolding | VSTO solution structure, NuGet packages (Math.NET, LiveCharts2, CommunityToolkit.Mvvm), CI pipeline | 2 days |
| Distribution module | Implement 6 core distributions wrapping Math.NET: Normal, Triangular, PERT, Lognormal, Uniform, Discrete. IDistribution interface with Sample(), Mean, PDF(), CDF(), Percentile(). | 3 days |
| Simulation engine | Core loop: generate sample matrix → (optionally recalc workbook) → store results. Support for seeded RNG, configurable iteration count. Background thread with progress callback. | 4 days |
| Summary statistics | Mean, median, std dev, min, max, percentiles (P1/P5/P10/P25/P50/P75/P90/P95/P99), skewness, kurtosis. | 2 days |
| Excel I/O layer | WorkbookManager: read/write cell values. InputTagManager: track which cells are simulation inputs. OutputTagManager: track output cells. CellHighlighter: color-code tagged cells. | 3 days |
| Ribbon + Task Pane shell | Custom ribbon tab ("MonteCarlo.XL") with Run/Stop/Settings buttons. WPF task pane hosted via ElementHost. Navigation between Setup → Run → Results views. | 3 days |
| Setup view | WPF form: add input cells (click-to-select from sheet), choose distribution + params, add output cells, set iteration count. Distribution preview sparkline. | 3 days |
| Results view — histogram | LiveCharts2 histogram with percentile markers (P10/P50/P90), kernel density overlay, and the summary stats card panel. Apply the full design system. | 4 days |
| Config persistence | Save/load simulation configuration as CustomXMLPart in the workbook. | 2 days |
| Sample workbook + manual testing | SimpleFinancialModel.xlsx — revenue model with 3 uncertain inputs and 1 output. | 1 day |
| Engine unit tests | Distribution sampling validation, engine correctness, summary stats accuracy. | 2 days |

**Deliverable:** User opens a workbook, clicks the ribbon to open the task pane, defines 3 uncertain inputs with distributions, runs 5,000 iterations, and sees a polished histogram with percentile markers and a summary statistics panel.

---

### Phase 2 — Storytelling Layer (Weeks 4–8)

**Goal:** Tornado charts, CDF view, multi-output support, and Excel export — everything needed to build a risk narrative.

| Task | Details | Est. |
|------|---------|------|
| Sensitivity analysis engine | Regression-based and rank-correlation-based sensitivity coefficients. Measures how much each input contributes to output variance. | 4 days |
| Tornado chart control | Custom WPF + SkiaSharp control. Horizontal bars sorted by impact, dual-color (increase/decrease), labeled axes, bar-tip annotations. Full design system compliance. | 5 days |
| CDF chart | LiveCharts2 cumulative distribution curve with area fill, percentile reference lines, and hover tooltip. Toggle between PDF and CDF views on the histogram. | 3 days |
| Multiple outputs | Support N output cells. Results dashboard gets a dropdown/tab to switch between outputs. Each output retains its own histogram, tornado, and stats. | 3 days |
| Target line & probability | User sets a target value; histogram shows a dashed red line with "P(X > target) = %" annotation. Stats panel shows probability of exceeding/falling below target. | 2 days |
| Recalc mode | For formula-dependent outputs: swap input values → Application.Calculate() → read outputs → repeat. User toggles between fast mode and recalc mode. Performance benchmarking. | 4 days |
| Results export to Excel | Write a formatted summary sheet: stats table, histogram as embedded image, tornado as embedded image. Styled for direct copy-paste into PowerPoint. | 4 days |
| Convergence monitoring | Running display of key stats (mean, P50, P90) as iterations progress. Visual indicator when stats stabilize (< 0.5% change over last 500 iterations). | 2 days |
| Additional distributions | Beta, Weibull, Exponential, Poisson — expanding the distribution picker. | 2 days |

**Deliverable:** Full analytical storytelling. User can answer "what are the key risk drivers?" (tornado), "what's the probability we hit our target?" (CDF + target line), and export a presentation-ready summary sheet directly from the add-in.

---

### Phase 3 — Correlation & Polish (Weeks 9–14)

**Goal:** Correlated inputs for realistic models, custom Excel functions, and a production-quality UX.

| Task | Details | Est. |
|------|---------|------|
| Iman-Conover correlation engine | Applies rank correlation to the sample matrix using a user-defined correlation matrix. Uses Cholesky decomposition (Math.NET). Validates matrix is positive semi-definite. | 6 days |
| Correlation matrix UI | Editable grid in the task pane. Color-coded cells (blue = positive, red = negative). Auto-validates PSD constraint and highlights violations. | 5 days |
| Excel UDF functions | =MC.Normal(mean, stddev), =MC.Triangular(min, mode, max), =MC.PERT(min, mode, max), etc. Registered via ExcelDna or COM automation. Cells display the mean/mode in non-simulation mode; during a sim run, they return the current iteration's sample. Auto-detected as inputs. | 6 days |
| Simulation profiles | Save named configurations ("Base Case", "Pessimistic", "Optimistic"). Quick-switch between profiles. Stored in CustomXMLParts. | 3 days |
| Dark theme | Full dark mode for the task pane and all chart controls. Follows Windows system theme or manual toggle. | 3 days |
| Performance optimization | Parallel simulation batches (Parallel.For on the iteration loop). Batch COM calls for recalc mode. Memory pooling for large iteration counts. Target: 10,000 iterations on a 20-input model in < 5 seconds. | 4 days |
| UX polish | Input validation with inline error messages. Keyboard shortcuts. Smooth view transitions (WPF animations). Loading skeletons. Empty states with helpful prompts. | 4 days |
| Documentation | User guide with screenshots, architecture doc, design system doc, README with GIF demos. | 3 days |
| Installer | MSI/ClickOnce installer for easy distribution to team members. | 2 days |

**Deliverable:** A robust, professional-grade tool that handles correlated inputs, offers in-cell formula syntax, looks beautiful in light and dark modes, and installs cleanly on team machines.

---

### Phase 4 — Future (Uncommitted)

Stretch goals to consider after Phase 3 ships:

- **Distribution fitting** — paste historical data, auto-fit best distribution (KS test, Anderson-Darling, AIC/BIC comparison)
- **Scatter matrix** — input vs. output scatter plots for exploring non-linear relationships
- **Scenario comparison** — overlay histograms from two profiles ("Base vs. Pessimistic") on the same chart
- **Goal seek under uncertainty** — "what input mean gives me 90% chance of hitting target?"
- **Report generation** — one-click multi-page report (cover + histogram + tornado + stats + assumptions) as a formatted Excel sheet or exported PDF
- **Copula-based correlation** — Gaussian, Clayton, Gumbel copulas for tail-dependent inputs
- **Time-series simulation** — correlated random walks for multi-period forecasting models
- **Sensitivity spider chart** — one-at-a-time sensitivity plot showing how each input affects the output across its range

---

## 6. Distribution Library

### Phase 1 — Core 6

| Distribution | Typical Use Case | Parameters | Math.NET Class |
|-------------|-----------------|------------|----------------|
| **Normal** | Symmetric uncertainty (costs, rates) | μ (mean), σ (std dev) | `MathNet.Numerics.Distributions.Normal` |
| **Triangular** | Expert estimates: min/likely/max | min, mode, max | `Triangular` |
| **PERT** | Smoother expert estimates (less extreme tails) | min, mode, max | Custom (Beta transform) |
| **Lognormal** | Right-skewed positives (revenue, prices) | μ, σ (of underlying normal) | `LogNormal` |
| **Uniform** | "Equally likely in range" | min, max | `ContinuousUniform` |
| **Discrete** | Scenario probabilities | values[], probs[] | `Categorical` + value map |

### Phase 2 — Add 4

| Distribution | Typical Use Case | Parameters |
|-------------|-----------------|------------|
| **Beta** | Percentages, conversion rates, yields | α, β |
| **Weibull** | Failure/reliability, time-to-event | shape (k), scale (λ) |
| **Exponential** | Time between events | rate (λ) |
| **Poisson** | Count of events in a fixed interval | rate (λ) |

### Phase 3+

Beta-PERT (4-param), Gamma, Pareto, Binomial, custom empirical from data.

---

## 7. Input Tagging: How Users Define Uncertainty

### Phase 1–2: Task Pane Setup

1. User clicks **"Add Input"** in the task pane
2. Clicks a cell in the workbook (e.g., B4) — the add-in captures the reference
3. Selects a distribution from a dropdown, enters parameters
4. A small distribution preview renders inline showing the shape
5. The cell gets a subtle colored border/background to indicate it's a simulation input
6. Config is stored as a CustomXMLPart inside the workbook

### Phase 3: In-Cell UDF Syntax

Users can also type distribution functions directly in cells:

```excel
=MC.Normal(100, 10)
=MC.Triangular(50, 100, 150)
=MC.PERT(80, 100, 130)
=MC.Lognormal(4.6, 0.3)
```

**In normal mode:** These functions return the distribution's expected value (mean or mode), so the spreadsheet behaves normally.

**During simulation:** The engine detects cells containing MC functions and treats them as inputs, sampling from the specified distribution on each iteration.

This gives users two workflows — GUI-driven (accessible) and formula-driven (fast for power users) — both producing the same simulation behavior.

---

## 8. Key Risks & Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| **VSTO deployment friction** | VSTO add-ins require .NET Framework + VSTO runtime on each machine; IT policies may block installs | Provide MSI installer + ClickOnce alternative; document prerequisites clearly; consider ExcelDna as a lighter deployment option if VSTO proves painful |
| **COM interop performance in recalc mode** | Recalculating a complex workbook 10,000 times through COM can be slow | Default to fast mode (in-memory); batch COM calls; disable screen updating + events during sim; offer iteration count guidance based on model complexity |
| **LiveCharts2 rendering edge cases** | LiveCharts2 is mature but can have layout quirks with dynamic resizing in a task pane | Fall back to SkiaSharp direct rendering for any chart that proves unreliable; keep chart controls behind an interface for easy swap |
| **Iman-Conover correctness** | Correlation implementation is mathematically non-trivial; subtle bugs produce plausible-looking but wrong results | Validate against published test cases (Iman & Conover 1982 paper examples); cross-validate against @RISK output on a trial license for specific test workbooks |
| **CustomXMLPart compatibility** | Some Excel versions or third-party tools may strip custom XML parts | Implement a backup persistence strategy (hidden sheet with JSON); detect and recover gracefully |
| **.NET version fragmentation** | VSTO traditionally targets .NET Framework 4.x; newer .NET 8+ is preferred for new code | Target .NET Framework 4.8 for maximum compatibility; architect engine + charts as .NET Standard 2.0 so they can migrate to .NET 8+ later |

---

## 9. Dev Environment Setup

### Prerequisites

- Visual Studio 2022 (Community is free) with:
  - .NET desktop development workload
  - Office/SharePoint development workload (includes VSTO project templates)
- Microsoft Excel (desktop, Windows)
- Git

### First-Time Setup

```bash
# Clone the repo
git clone https://github.com/<your-org>/MonteCarlo.XL.git
cd MonteCarlo.XL

# Open the solution in Visual Studio
start MonteCarlo.XL.sln

# NuGet packages will auto-restore on build:
#   - MathNet.Numerics
#   - LiveChartsCore.SkiaSharpView.WPF
#   - CommunityToolkit.Mvvm
#   - SkiaSharp.Views.WPF
#   - xunit, FluentAssertions, Moq (test projects)

# Build and run (F5)
# → Visual Studio will launch Excel with the add-in sideloaded
```

### Project Structure in Visual Studio

```
Solution 'MonteCarlo.XL'
├── src/
│   ├── MonteCarlo.Engine        (Class Library — .NET Standard 2.0)
│   ├── MonteCarlo.Charts        (WPF Class Library — .NET Framework 4.8)
│   ├── MonteCarlo.UI            (WPF Class Library — .NET Framework 4.8)
│   └── MonteCarlo.Addin         (VSTO Excel Add-in — .NET Framework 4.8)
└── tests/
    └── MonteCarlo.Engine.Tests  (xUnit — .NET Standard 2.0)
```

---

## 10. Definition of Done (Per Phase)

Each phase is "done" when:

1. All tasks in the phase table are complete and merged to `main`
2. Engine unit tests pass with >90% coverage on `MonteCarlo.Engine`
3. Manual testing on Excel 2019+ (desktop Windows) with sample workbooks
4. Sample workbooks updated to demonstrate new features
5. README updated with current feature list and screenshots/GIFs
6. No known crashes or data-loss bugs in the simulation workflow

---

## 11. Open Questions

- [ ] **VSTO vs. ExcelDna**: VSTO is the plan, but ExcelDna (open-source) is lighter to deploy (single .xll file, no VSTO runtime needed) and supports UDFs natively. Worth a spike in Week 1 to compare. If ExcelDna wins on deployment simplicity, we pivot early — the engine/charts architecture is identical either way.
- [ ] **LiveCharts2 vs. ScottPlot vs. OxyPlot**: LiveCharts2 is the current pick for its modern aesthetic, but ScottPlot is simpler and OxyPlot is battle-tested. Build one histogram in each during Week 1 to compare visual quality and task-pane behavior.
- [ ] **.NET Framework 4.8 vs. .NET 8+**: VSTO requires .NET Framework. If we pivot to ExcelDna, we could target .NET 8+ for the whole stack (better performance, modern C# features). Decision depends on the VSTO vs. ExcelDna spike.
- [ ] **Chart export format**: For the "export to Excel sheet" feature, should we render charts as high-DPI PNG images embedded in the sheet, or attempt to generate native Excel charts programmatically? PNG is easier and preserves our design language; native charts are editable but won't look as good.
