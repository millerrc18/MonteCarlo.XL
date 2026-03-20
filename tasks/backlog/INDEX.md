# Task Backlog — Ordered Pipeline

> Tasks are promoted from here to `tasks/current.md` one at a time by the **architect** (Claude chat). The **developer** (Claude Code) only works on `current.md`.

## Phase 1 — Walking Skeleton

| ID | Title | Depends On | Status |
|----|-------|------------|--------|
| TASK-001 | Solution Scaffolding + ExcelDna Spike | — | ✅ COMPLETE |
| TASK-002 | Distribution Module — Core 6 Distributions | 001 | ✅ COMPLETE |
| TASK-003 | Simulation Engine — Core Monte Carlo Loop | 002 | ✅ COMPLETE |
| TASK-004 | Summary Statistics | — | ✅ COMPLETE |
| TASK-005 | Sensitivity Analysis Engine | 003, 004 | ✅ COMPLETE |
| TASK-006 | Excel I/O Layer (COM Interop) | — | ✅ COMPLETE |
| TASK-007 | Ribbon + WPF Task Pane Shell + Style System | — | ✅ COMPLETE |
| TASK-008 | Setup View (Input/Output Config UI) | 002, 006, 007 | ✅ COMPLETE |
| TASK-009 | Results Dashboard — Histogram, CDF, Stats Panel | 004, 007 | ⬜ QUEUED |
| TASK-010 | Tornado Chart (Custom SkiaSharp) | 005, 007 | ⬜ QUEUED |
| TASK-011 | Config Persistence + Run View + Convergence | 003, 006, 007 | ⬜ QUEUED |
| TASK-012 | Results Export to Excel Sheet | 004, 005, 006, 009, 010 | ⬜ QUEUED |
| TASK-013 | Simulation Orchestrator (End-to-End Integration) | 002–012 | ⬜ QUEUED |

## Phase 2 — Advanced Features

| ID | Title | Depends On | Status |
|----|-------|------------|--------|
| TASK-014 | Additional Distributions (Beta, Weibull, Exp, Poisson) | 002 | ⬜ QUEUED |

## Phase 3 — Correlation, UDFs, & Polish

| ID | Title | Depends On | Status |
|----|-------|------------|--------|
| TASK-015 | Iman-Conover Correlation Engine | 002, 003 | ⬜ QUEUED |
| TASK-016 | Correlation Matrix UI | 015, 007, 008 | ⬜ QUEUED |
| TASK-017 | Custom Excel Functions (=MC.Normal() UDFs) | 002, 013 | ⬜ QUEUED |
| TASK-018 | Dark Theme + Performance + UX Polish | ALL | ⬜ QUEUED |

---

## Critical Path (fastest route to first working simulation)

```
002 ✅ → 003 ✅ → 004 ✅ → 005 ✅ → 006 ✅ → 007 ✅ → 008 ✅ → 009 → 011 → 013
```

Tasks 006 and 007 can run in parallel (no interdependencies).
Tasks 009, 010, and 011 can run in parallel once 007 is done.

---

### Statuses
- ✅ COMPLETE — in `tasks/completed/`
- 🔵 CURRENT — in `tasks/current.md`, being worked on
- ⬜ QUEUED — spec written in `tasks/backlog/`, ready to promote
