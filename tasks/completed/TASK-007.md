# TASK-007: Ribbon & WPF Task Pane Shell

## Context

Read `ROADMAP.md` for full project context. The ExcelDna add-in scaffold (TASK-001) has a basic ribbon and entry point. This task turns it into a real UI shell with a WPF task pane, navigation between views, and a styled foundation for all subsequent UI work.

## Objective

Build the ribbon with working buttons, host a WPF task pane in Excel's Custom Task Pane, and implement view navigation (Setup → Run → Results). Establish the global WPF style system (colors, typography, spacing) that all future views will use.

## Dependencies

- TASK-001 (ExcelDna scaffold with ribbon placeholder)

## Design

### Ribbon

The ribbon tab "MonteCarlo.XL" should have these groups and buttons:

**Simulation group:**
- **Run** (large button, play icon) — starts the simulation. Disabled when no inputs are configured.
- **Stop** (large button, stop icon) — cancels a running simulation. Only enabled during a run.

**View group:**
- **Task Pane** (toggle button) — shows/hides the task pane

**Settings group:**
- **Settings** (small button, gear icon) — opens the Settings view in the task pane

Update the existing `MonteCarloRibbon.cs` and ribbon XML from TASK-001. Use ExcelDna's `ExcelRibbon` base class and `GetCustomUI` override.

### WPF Task Pane

ExcelDna supports Custom Task Panes via `CustomTaskPaneFactory`. Host a WPF `UserControl` inside a `WindowsFormsHost`/`ElementHost`:

```csharp
// In AddIn.cs or a TaskPaneController
var wpfControl = new MainTaskPaneControl();  // WPF UserControl
var host = new ElementHost { Child = wpfControl, Dock = DockStyle.Fill };
var ctp = CustomTaskPaneFactory.CreateCustomTaskPane(host, "MonteCarlo.XL");
ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
ctp.Width = 380;
ctp.Visible = true;
```

### Navigation

The task pane contains a single `MainTaskPaneControl` that swaps between views:

```
┌──────────────────────────┐
│  MonteCarlo.XL    [gear] │  ← Header bar with title + settings icon
├──────────────────────────┤
│                          │
│   [Current View]         │  ← Content area: swaps between views
│                          │
├──────────────────────────┤
│  Setup │ Results         │  ← Tab bar at bottom (or top)
└──────────────────────────┘
```

Views to navigate between:
1. **SetupView** — configure inputs, outputs, distributions (TASK-008)
2. **RunView** — progress bar during simulation (TASK-008)
3. **ResultsView** — histogram, stats, tornado (TASK-009+)
4. **SettingsView** — iteration count, seed, theme

For now, create **placeholder UserControls** for each view with just a label ("Setup View — Coming Soon", etc.). The real implementations come in subsequent tasks.

Use a `MainViewModel` with a `CurrentView` property and `ICommand`s for navigation. The `MainTaskPaneControl` binds its content area to `CurrentView`.

```csharp
public partial class MainViewModel : ObservableObject
{
    [ObservableProperty]
    private object _currentView;

    [RelayCommand]
    private void NavigateToSetup() => CurrentView = new SetupView();

    [RelayCommand]
    private void NavigateToResults() => CurrentView = new ResultsView();

    [RelayCommand]
    private void NavigateToSettings() => CurrentView = new SettingsView();
}
```

### Global Style System

Create `MonteCarlo.UI/Styles/GlobalStyles.xaml` as a `ResourceDictionary` that all views reference. This is critical — it establishes the visual language for the entire product.

**Color palette (as WPF resources):**
```xml
<!-- Primary -->
<Color x:Key="Blue500">#3B82F6</Color>
<Color x:Key="Violet500">#8B5CF6</Color>
<Color x:Key="Emerald500">#10B981</Color>
<Color x:Key="Red500">#EF4444</Color>
<Color x:Key="Orange500">#F97316</Color>
<Color x:Key="Amber500">#F59E0B</Color>

<!-- Neutrals (Light Theme) -->
<Color x:Key="Background">#FFFFFF</Color>
<Color x:Key="Surface">#F8FAFC</Color>
<Color x:Key="SurfaceHover">#F1F5F9</Color>
<Color x:Key="Border">#E2E8F0</Color>
<Color x:Key="TextPrimary">#0F172A</Color>
<Color x:Key="TextSecondary">#64748B</Color>
<Color x:Key="TextTertiary">#94A3B8</Color>

<!-- Create SolidColorBrush versions for binding -->
<SolidColorBrush x:Key="Blue500Brush" Color="{StaticResource Blue500}"/>
<!-- ... etc for all colors ... -->
```

**Typography:**
```xml
<FontFamily x:Key="DefaultFont">Segoe UI Variable, Segoe UI, sans-serif</FontFamily>
<FontFamily x:Key="MonoFont">Cascadia Code, Consolas, monospace</FontFamily>

<Style x:Key="HeadlineLarge" TargetType="TextBlock">
    <Setter Property="FontFamily" Value="{StaticResource DefaultFont}"/>
    <Setter Property="FontSize" Value="24"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
</Style>

<Style x:Key="HeadlineMedium" TargetType="TextBlock">
    <Setter Property="FontSize" Value="18"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <!-- ... -->
</Style>

<Style x:Key="BodyText" TargetType="TextBlock">
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="FontWeight" Value="Normal"/>
    <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
</Style>

<Style x:Key="StatValue" TargetType="TextBlock">
    <Setter Property="FontFamily" Value="{StaticResource MonoFont}"/>
    <Setter Property="FontSize" Value="20"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
</Style>
```

**Base control styles (buttons, cards, input fields):**
```xml
<!-- Primary button -->
<Style x:Key="PrimaryButton" TargetType="Button">
    <Setter Property="Background" Value="{StaticResource Blue500Brush}"/>
    <Setter Property="Foreground" Value="White"/>
    <Setter Property="Padding" Value="16,8"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}"
                        CornerRadius="6" Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <!-- Hover/press triggers -->
</Style>

<!-- Card container -->
<Style x:Key="Card" TargetType="Border">
    <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="CornerRadius" Value="8"/>
    <Setter Property="Padding" Value="16"/>
</Style>
```

**Spacing constants:**
```xml
<sys:Double x:Key="SpacingXs">4</sys:Double>
<sys:Double x:Key="SpacingSm">8</sys:Double>
<sys:Double x:Key="SpacingMd">12</sys:Double>
<sys:Double x:Key="SpacingLg">16</sys:Double>
<sys:Double x:Key="SpacingXl">24</sys:Double>
<sys:Double x:Key="SpacingXxl">32</sys:Double>
```

## Implementation Notes

### Task Pane Width

The task pane should default to **380px wide**. This is wider than the default (~300px) but still comfortable alongside a typical spreadsheet. Charts and content should be designed for this width. All views should be responsive down to 320px.

### Resource Dictionary Loading

The `GlobalStyles.xaml` resource dictionary needs to be merged into `App.xaml` or loaded explicitly in the task pane control's resources:

```csharp
// In MainTaskPaneControl constructor or XAML
Resources.MergedDictionaries.Add(new ResourceDictionary
{
    Source = new Uri("pack://application:,,,/MonteCarlo.UI;component/Styles/GlobalStyles.xaml")
});
```

Since we're running inside Excel (not a standalone WPF app), there's no `App.xaml`. Load resource dictionaries in the root task pane control.

### ExcelDna Ribbon Icons

For ribbon button images, use `imageMso` attribute with built-in Office icons:
- Run: `imageMso="PlayMacro"` or `imageMso="AnimationPlay"`
- Stop: `imageMso="RecordStop"` or `imageMso="AnimationStop"`
- Task Pane: `imageMso="ViewSideBySide"` or `imageMso="TaskPane"`
- Settings: `imageMso="AdvancedFileProperties"` or `imageMso="PropertySheet"`

## File Structure

```
MonteCarlo.Addin/
├── AddIn.cs                          # Updated: initialize task pane on startup
├── Ribbon/
│   ├── MonteCarloRibbon.xml          # Updated: full ribbon layout
│   └── MonteCarloRibbon.cs           # Updated: button callbacks
└── TaskPane/
    ├── TaskPaneController.cs         # Show/hide logic, ExcelDna CTP management
    └── TaskPaneHost.cs               # WinForms host with ElementHost → WPF

MonteCarlo.UI/
├── Views/
│   ├── MainTaskPaneControl.xaml/.cs  # Root control with navigation + content area
│   ├── SetupView.xaml/.cs            # Placeholder
│   ├── RunView.xaml/.cs              # Placeholder
│   ├── ResultsView.xaml/.cs          # Placeholder
│   └── SettingsView.xaml/.cs         # Placeholder
├── ViewModels/
│   └── MainViewModel.cs             # Navigation state, current view
└── Styles/
    └── GlobalStyles.xaml             # Colors, typography, buttons, cards, spacing
```

## Commit Strategy

```
feat(ui): add GlobalStyles.xaml — color palette, typography, buttons, cards
feat(ui): add MainTaskPaneControl with view navigation
feat(ui): add placeholder views — Setup, Run, Results, Settings
feat(addin): implement TaskPaneController with ExcelDna CTP hosting
feat(addin): update ribbon with Run/Stop/TaskPane/Settings buttons
```

## Done When

- [ ] Ribbon tab "MonteCarlo.XL" visible in Excel with all buttons
- [ ] Task Pane toggle button shows/hides a WPF task pane on the right side
- [ ] Navigation between Setup / Results / Settings views works
- [ ] GlobalStyles.xaml contains the full color palette, typography scale, button styles, and card styles
- [ ] All placeholder views render correctly in the task pane at 380px width
- [ ] `dotnet build` clean
