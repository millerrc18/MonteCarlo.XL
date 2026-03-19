# Task Backlog — Ordered Pipeline

> Tasks are promoted from here to `tasks/current.md` one at a time by the **architect** (Claude chat). The **developer** (Claude Code) only works on `current.md`.

## Phase 1 — Walking Skeleton

| ID | Title | Depends On | Status |
|----|-------|------------|--------|
| TASK-001 | Solution Scaffolding + VSTO vs ExcelDna Spike | — | 🔵 CURRENT |
| TASK-002 | Distribution Engine — Core 6 Distributions | 001 | ⬜ QUEUED |
| TASK-003 | Simulation Engine — Core Monte Carlo Loop | 002 | ⬜ QUEUED |
| TASK-005 | Chart Controls — Histogram + Tornado | 003, 004 | ⬜ QUEUED |
| TASK-006 | Excel I/O Layer + Cell Tagging | 001 | ⬜ NOT WRITTEN |
| TASK-007 | Task Pane UI — Setup + Run + Results Views | 005, 006 | ⬜ NOT WRITTEN |
| TASK-008 | Config Persistence (CustomXMLPart) | 007 | ⬜ NOT WRITTEN |
| TASK-009 | Sample Workbook + End-to-End Integration Test | ALL | ⬜ NOT WRITTEN |

## Phase 2 — Storytelling Layer

| ID | Title | Depends On | Status |
|----|-------|------------|--------|
| TASK-004 | Sensitivity Analysis Engine | 003 | ⬜ QUEUED |
| TASK-010 | CDF Chart + Multi-Output Support | 005 | ⬜ NOT WRITTEN |
| TASK-011 | Target Line + Probability Annotations | 005 | ⬜ NOT WRITTEN |
| TASK-012 | Recalc Mode (Formula-Dependent Outputs) | 006 | ⬜ NOT WRITTEN |
| TASK-013 | Results Export to Excel Sheet | 007 | ⬜ NOT WRITTEN |
| TASK-014 | Convergence Monitoring | 003 | ⬜ NOT WRITTEN |

## Phase 3 — Correlation & Polish

| ID | Title | Depends On | Status |
|----|-------|------------|--------|
| TASK-015 | Iman-Conover Correlation Engine | 003 | ⬜ NOT WRITTEN |
| TASK-016 | Correlation Matrix UI | 015 | ⬜ NOT WRITTEN |
| TASK-017 | Custom Excel UDFs (=MC.Normal) | 001 | ⬜ NOT WRITTEN |
| TASK-018 | Simulation Profiles | 008 | ⬜ NOT WRITTEN |
| TASK-019 | Dark Theme | 005 | ⬜ NOT WRITTEN |
| TASK-020 | Performance Optimization | 003 | ⬜ NOT WRITTEN |
| TASK-021 | UX Polish + Installer | ALL | ⬜ NOT WRITTEN |

---

### Statuses
- 🔵 CURRENT — in `tasks/current.md`, being worked on
- ⬜ QUEUED — spec written in `tasks/backlog/`, ready to promote
- ⬜ NOT WRITTEN — planned but spec not yet written by the architect
- ✅ COMPLETE — in `tasks/completed/`
