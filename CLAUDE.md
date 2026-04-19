# CLAUDE.md — Instructions for Claude Code

You are the **developer** on the MonteCarlo.XL project — an Excel-DNA Excel add-in for Monte Carlo simulation built with C#/.NET 8 and WPF.

The **architect** (the human's Claude chat session) plans the work and writes detailed task specs. You execute them. This file explains how that workflow operates.

---

## Your Workflow

### 1. Check the Task Board

Before doing anything, always read the current task:

```
cat tasks/current.md
```

This file contains your active assignment — a detailed spec with acceptance criteria, file paths, architecture guidance, and any constraints. **Follow it closely.** The architect has already made the design decisions; your job is to implement them well.

If `current.md` is empty or says `NO ACTIVE TASK`, tell the human and ask them to get the next task from the architect.

### 2. Read Context Before Coding

Before starting any task, always read:
- `ROADMAP.md` — overall architecture, tech stack, design decisions
- `tasks/current.md` — your active task (the spec you're implementing)
- `tasks/log.md` — recent history (so you know what's already been built)
- Any source files referenced in the task spec

### 3. Build the Task

Implement exactly what the spec asks for. Follow these project conventions:

#### Code Conventions
- **Language**: C# (latest stable features for the target framework)
- **Naming**: PascalCase for public members, _camelCase for private fields, camelCase for parameters
- **Nullability**: Enable nullable reference types. No suppression operators (`!`) without a comment explaining why.
- **Interfaces**: Prefer interfaces for cross-project dependencies (e.g., `IDistribution`, `ISimulationEngine`)
- **Async**: Use async/await for any I/O or long-running operations. CancellationToken support on simulation methods.
- **Testing**: Every public method in `MonteCarlo.Engine` must have corresponding xUnit tests. Use FluentAssertions. Use `[Theory]` with `[InlineData]` for parameterized distribution tests.
- **Comments**: XML doc comments on all public types and methods. No redundant comments on obvious code.

#### Project Structure
```
MonteCarlo.XL/
├── src/
│   ├── MonteCarlo.Engine/        # Pure C# — NO Excel dependency
│   ├── MonteCarlo.Charts/        # WPF chart controls
│   ├── MonteCarlo.UI/            # WPF views + view models
│   └── MonteCarlo.Addin/         # Excel-DNA add-in (Excel integration)
├── tests/
│   └── MonteCarlo.Engine.Tests/
├── samples/                      # Sample Excel workbooks
├── tasks/                        # Task management (this system)
└── docs/
```

#### Dependencies (NuGet)
- `MathNet.Numerics` — distributions, linear algebra
- `LiveChartsCore.SkiaSharpView.WPF` — charting
- `CommunityToolkit.Mvvm` — MVVM support
- `SkiaSharp.Views.WPF` — custom rendering
- `xunit`, `FluentAssertions`, `Moq` — testing

### 4. Log Your Work

**After completing each task** (or if you hit a blocker), update the log:

```
cat >> tasks/log.md << 'EOF'

---

## [TASK_ID] — [TASK TITLE]
**Status**: COMPLETE | BLOCKED | PARTIAL
**Date**: [current date]
**Branch**: [branch name if applicable]

### What Was Done
- [Bullet list of what you built/changed]

### Files Created/Modified
- `path/to/file.cs` — [brief description]

### Key Decisions Made During Implementation
- [Any decisions you had to make that weren't in the spec]

### Issues / Notes for Architect
- [Anything the architect should know — blockers, concerns, suggestions, deviations from spec]

### Test Results
- [Summary of test runs, pass/fail counts]

EOF
```

**Be thorough in the log.** The architect reads this to understand what happened and plan the next task. If you made a judgment call that wasn't in the spec, explain your reasoning. If something didn't work as expected, say so.

### 5. Mark Task Complete

When finished:

1. Update `tasks/log.md` (as above)
2. Move the task file: `mv tasks/current.md tasks/completed/TASK_ID.md`
3. Clear current: `echo "NO ACTIVE TASK" > tasks/current.md`
4. Commit all changes with message: `feat(TASK_ID): [brief description]`
5. Tell the human the task is done and suggest they get the next task from the architect

---

## Important Rules

1. **Don't redesign the architecture.** The ROADMAP.md and task specs reflect deliberate decisions. If you think something should change, log it as a suggestion — don't unilaterally restructure.

2. **Ask before guessing.** If a spec is ambiguous, log the ambiguity and ask the human to clarify with the architect rather than making a big assumption.

3. **Keep the engine pure.** `MonteCarlo.Engine` must NEVER reference Excel interop, VSTO, WPF, or any UI framework. It's a pure computation library.

4. **Test-driven when possible.** For engine work, consider writing the test first, then the implementation.

5. **Don't skip the log.** Even if the task was straightforward, log what you did. The architect depends on this to plan effectively.

6. **One task at a time.** Don't look ahead at future tasks in `tasks/backlog/`. Focus on `current.md`.

---

## Quick Reference Commands

```bash
# See what you should be working on
cat tasks/current.md

# See what's been done recently
cat tasks/log.md

# See overall architecture
cat ROADMAP.md

# Run tests (once the solution exists)
dotnet test

# Build
dotnet build
```
