# TASK-001: Solution Scaffolding + VSTO vs ExcelDna Spike

**Priority**: 🔴 Critical (blocks everything)
**Phase**: 1 — Walking Skeleton
**Estimated Effort**: ~2 hours

---

## Objective

Create the Visual Studio solution structure and make a decision on VSTO vs ExcelDna. This is the foundation everything else builds on.

---

## Background

We need to decide between two approaches for the Excel add-in host:

**Option A — VSTO (Visual Studio Tools for Office)**
- Microsoft's official approach
- Requires .NET Framework 4.8
- Requires VSTO runtime installed on target machines
- Rich designer support in Visual Studio
- Custom Task Pane via `UserControl` + `ElementHost` (for WPF)

**Option B — ExcelDna**
- Open-source, community standard
- Can target .NET 8+ (modern C#, better performance)
- Deploys as a single `.xll` file (no runtime install)
- Native UDF support (easier path to `=MC.Normal()` functions)
- Custom Task Pane via `ExcelDna.Integration.CustomUI`
- More lightweight, less ceremony

**The architect's recommendation is to go with ExcelDna** if the spike confirms it can:
1. Host a WPF task pane with proper lifecycle management
2. Register custom ribbon buttons
3. Deploy without requiring admin install on target machines

---

## Deliverables

### 1. Solution Structure

Create the following solution and projects. Even if some projects are mostly empty, establish them now so the dependency graph is correct from day one.

```
MonteCarlo.XL/
├── MonteCarlo.XL.sln
├── src/
│   ├── MonteCarlo.Engine/
│   │   ├── MonteCarlo.Engine.csproj          # Class Library, .NET Standard 2.0 (or net8.0 if ExcelDna)
│   │   └── [placeholder class]
│   ├── MonteCarlo.Charts/
│   │   ├── MonteCarlo.Charts.csproj          # WPF Class Library
│   │   └── [placeholder class]
│   ├── MonteCarlo.UI/
│   │   ├── MonteCarlo.UI.csproj              # WPF Class Library
│   │   └── [placeholder class]
│   └── MonteCarlo.Addin/
│       ├── MonteCarlo.Addin.csproj           # VSTO or ExcelDna project
│       └── [add-in entry point]
├── tests/
│   └── MonteCarlo.Engine.Tests/
│       ├── MonteCarlo.Engine.Tests.csproj    # xUnit test project
│       └── [placeholder test]
├── samples/
│   └── (empty, for now)
├── .gitignore                                # Visual Studio + .NET gitignore
├── CLAUDE.md
├── ROADMAP.md
├── README.md
└── tasks/
```

### 2. Project Dependencies (References)

```
MonteCarlo.Engine    →  (no project refs — standalone)
MonteCarlo.Charts    →  MonteCarlo.Engine
MonteCarlo.UI        →  MonteCarlo.Engine, MonteCarlo.Charts
MonteCarlo.Addin     →  MonteCarlo.Engine, MonteCarlo.UI
MonteCarlo.Engine.Tests → MonteCarlo.Engine
```

### 3. NuGet Packages

| Project | Packages |
|---------|----------|
| Engine | `MathNet.Numerics` |
| Charts | `LiveChartsCore.SkiaSharpView.WPF`, `SkiaSharp.Views.WPF` |
| UI | `CommunityToolkit.Mvvm` |
| Addin | `ExcelDna.AddIn` (if ExcelDna) or VSTO template refs |
| Tests | `xunit`, `xunit.runner.visualstudio`, `FluentAssertions`, `Moq`, `Microsoft.NET.Test.Sdk` |

### 4. Verify It Builds

The solution should build cleanly with `dotnet build` (or MSBuild for VSTO). No errors, no warnings treated as errors.

### 5. Smoke Test

If using ExcelDna: confirm the .xll loads in Excel and a placeholder ribbon tab appears.
If using VSTO: confirm F5 launches Excel with the add-in loaded and a placeholder ribbon tab appears.

(If Excel is not available in the dev environment, just confirm the build succeeds and note that the Excel smoke test needs to be done manually.)

### 6. .gitignore

Use the standard Visual Studio .gitignore. Ensure it covers:
- `bin/`, `obj/`
- `.vs/`
- `*.user`
- `*.suo`
- NuGet `packages/` (if not using PackageReference)

---

## Acceptance Criteria

- [ ] Solution file exists and all 5 projects are present
- [ ] Project references match the dependency graph above
- [ ] NuGet packages are referenced (not necessarily restored if offline)
- [ ] `dotnet build` (or equivalent) succeeds with no errors
- [ ] .gitignore is in place
- [ ] A brief note in the log explaining the VSTO vs ExcelDna decision and why

---

## Notes from the Architect

- If you're working in an environment without Visual Studio or Excel, that's fine — use `dotnet new` to create the projects via CLI and wire up the references manually. The important thing is the structure and dependency graph, not the IDE.
- For ExcelDna, the entry point is a class implementing `IExcelAddIn` with `AutoOpen()` and `AutoClose()` methods.
- For the WPF projects (Charts, UI), they need to target a framework that supports WPF: `net8.0-windows` (if ExcelDna/.NET 8) or `net48` (if VSTO).
- Don't overthink this task — it's scaffolding. Get the structure right, get it building, log your decision, move on.
