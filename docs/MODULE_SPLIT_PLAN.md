# Capacity Tracker – Module Split Plan

Goal: split the single, very large `modCapacityPlanner.bas` into two modules to improve maintainability and make it easier to paste/import in the VBE.

Target modules

- modCapacityCore
  - Core bootstrap and setup (sheets, tables, names)
  - Dashboard and Availability features
  - Roster/Config helpers
  - Logging and generic utilities
  - Metrics skeleton
  - Shared helpers used by Flow/Jira (exposed as Public)

- modFlowJira
  - Flow Metrics (all `Flow_*` procedures and helpers)
  - WIP CSV Sanitizer (`WIP_ImportCSV` and helpers)
  - Jira normalization and insights (Jira_* procs + stats helpers)
  - Raw data sample builders
  - Orchestration (`SanitizeRawAndBuildInsights`, `RefreshSamples`)

Notes

- Keep procedure names unchanged to preserve button `OnAction` wiring.
- Keep constants `Private` in each module to avoid name collisions.
- Expose these helpers as `Public` in `modCapacityCore` so `modFlowJira` can call them:
  - `EnsureSheet`, `EnsureTable`, `HasTable`, `GetNameValueOr`
  - `ColumnsSummary`, `WorkdaysBetween`, `ParseDurationDays`, `Norm`, `FindByContains`
  - `QuarterStartDate`, `FormatSprintName`, `FormatSprintTag`

Suggested cut list

- Move the following sections from `modCapacityPlanner.bas` to `modCapacityCore`:
  - Bootstrap + step wrappers: `Bootstrap`, `Step_EnsureSheets|Tables|Seed*`, `HealthCheck`
  - Core setup: `EnsureSheets`, `RemoveLegacySampleSheets`, `EnsureTables`
  - Settings/samples: `SeedNamedValues`, `SeedSamplesIfPresent`
  - Helpers and logging listed above
  - Dashboard + Availability block
  - Roster/config block
  - Metrics skeleton block

- Move the following sections from `modCapacityPlanner.bas` to `modFlowJira`:
  - "WIP CSV Sanitizer" block
  - "Flow Metrics" block (all `Flow_*`)
  - "Jira integration" block
  - "Jira issues analysis" block
  - Raw data/sample builders and normalization
  - Orchestration (`SanitizeRawAndBuildInsights`, `FindSheetLoose`, `RefreshSamples`)

Button wiring (no changes needed)

- Dashboard assigns:
  - `CreateOrAdvanceAvailability` (Core)
  - `SanitizeRawAndBuildInsights` (Flow/Jira)
  - `RefreshSamples` (Flow/Jira)

Step‑by‑step (VBE)

1) Insert Module → name it `modCapacityCore`.
2) Cut/paste the Core sections listed above into `modCapacityCore`.
3) Ensure the shared helpers are `Public` in `modCapacityCore`.
4) Insert Module → name it `modFlowJira`.
5) Cut/paste Flow, WIP CSV, Jira, and orchestration sections into `modFlowJira`.
6) Back in `modCapacityPlanner.bas`, delete the moved code (or leave a comment header and remove the module from the workbook to avoid duplicates).
7) Debug → Compile VBAProject. Fix any missing `Public` exposure or references.

Repository strategy

- Branch: `feature/fix-duplicate-dim-flow-scatter` (already created).
- Commit in small chunks to avoid tooling/path limits:
  1. Add `modCapacityCore.bas` with Core bootstrap + utilities only.
  2. Add `modFlowJira.bas` with Flow entrypoints first (e.g., `Flow_BuildCharts`) that call into existing helpers in Core.
  3. Incrementally migrate remaining procedures in follow‑up commits.
- After both modules are complete, reduce `modCapacityPlanner.bas` to a short shim (or remove the file from the repo).

Post‑split validation

- Run `Diagnostics_RunBootstrap`.
- Run `SanitizeRawAndBuildInsights` against `Raw_Data` and verify:
  - WIP Aging chart + percentile bands render.
  - Throughput and Cycle Scatter render.
  - Jira Insights build (when source looks Jira‑like).
- Dashboard buttons still work.

Changelog entry (suggested)

- 2025‑10‑29: Fixed duplicate Dim in cycle scatter (`tag` → `tagKey`).
- 2025‑10‑29: Split single module into `modCapacityCore` and `modFlowJira` (no functional changes), updated docs.

