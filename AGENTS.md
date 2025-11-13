# Capacity Tracker – Agent Guide

Ground rules
- Always create a branch. Never commit to `main`.
- Ask before pushing or merging. PRs only with owner approval.
- Keep changes minimal and focused. No unrequested features.
- Update docs as you go (README, SCHEMA, CHANGELOG, this log).

Working model
- Paste-once single VBA module for a minimal, working baseline.
- No classes or forms initially to avoid VBE import issues.
- CSV-first for PTO and Holidays. Outlook/Jira later.

Current status
- Fresh restart. Minimal baseline to be added on `feature/restart`.

Next actions
- Add `modCapacityPlanner.bas` (single module) with:
  - `Bootstrap` to create sheets/tables/names and seed sample CSVs (if present).
  - `ImportPTO_CSV` to append PTO from a selected CSV.
  - `HealthCheck` to validate required pieces.
- Provide simple docs and sample CSVs.

Running work log
- [2025-10-19] Repo reset to empty. Planning minimal baseline on `feature/restart`.
- [2025-10-19] Added minimal single-module baseline, Getting_Started sheet, and docs (README, SCHEMA, GETTING_STARTED, ROADMAP, PTO_OPTIONS). No PTO automation yet.
- [2025-10-27] On `feature/sanitize-only-button-and-sp-cycle`: removed "Build Jira Insights" button (use "Sanitize Raw + Build Insights"), and replaced the dashboard insights summary with average cycle time by story points (1,2,3,5,8,13). Updated JIRA_INTEGRATION docs.
- [2025-10-28] On `feature/remove-epic-chart-and-status-table`: removed Epic chart ("Epic: Sum SP and Count") and the Status table (Done/In Progress/To Do) from Insights build.
 - [2025-10-28] On `feature/flow-metrics-charts`: added pasteable `Flow_BuildCharts` to generate CFD (WIP), Throughput, and Cycle Time scatter from sanitized facts; added docs/FLOW_METRICS.md.
 - [2025-10-28] On `feature/wip-csv-sanitizer`: added `WIP_ImportCSV` to sanitize WIP CSV (time-in-status + dates) into `WIP_Facts!tblWIPFacts`; added `data/wip_example.csv` and docs/WIP_IMPORT.md.
 - [2025-10-28] On `feature/remove-epic-chart-and-cfd-table`: removed Epic pivot (Sum SP + Count) from Jira_Insights and suppressed the CFD source table (To Do/In Progress/Done) in Flow_Metrics; added header mapping for "Resolved date" in WIP CSV import.
 - [2025-10-28] On `feature/wip-aging-first-with-data`: WIP Aging now builds first on Flow_Metrics and writes a visible "WIP Aging – Data" table (IssueKey, Stage, AgeDays) before the chart; fallbacks retained so the chart always renders.
 - [2025-10-28] On `feature/wip-aging-always`: WIP Aging now always renders — if no unresolved items, shows historical-at-resolution ages; if time-in-status columns are missing, derives durations from Created/Start Progress/Resolved. Updated docs/FLOW_METRICS.md.
- [2025-10-28] On `feature/wip-aging-backdrop-fix`: fixed compile error on AddShape by drawing WIP Aging backdrop/labels as worksheet shapes positioned via PlotArea coordinates; added DoEvents/Refresh before reading PlotArea; improved cleanup to remove shapes by AlternativeText so re-runs don’t stack.
 - [2025-10-29] On `feature/fix-duplicate-dim-flow-scatter`: fixed duplicate declaration bug in Cycle Scatter (renamed inner `tag` to `tagKey`). Abandoned split experiment; removed `modCapacityCore.bas` and `modFlowJira.bas`, restored single-module public macros (removed `Option Private Module`), and deleted split docs.
 - [2025-10-29] Date-first with tag confirmation: `CreateTeamAvailability` first asks for a start date (MM/DD/YYYY), then prompts to confirm the Sprint tag (default suggested) — no tag assumptions from date. If the date prompt is canceled, it falls back to Year/Quarter/Sprint. `CreateOrAdvanceAvailability` still advances by +14 days for dates, and names the new sheet by incrementing the last sheet’s tag.
- [2025-10-29] Add guard before creating Availability: `CreateTeamAvailabilityAtDate` now validates that `Config!tblRoster` has at least one member before adding a new sheet; otherwise shows a friendly message and exits without creating an empty sheet.
- [2025-11-02] On `feature/fix-dashboard-onaction`: fixed Bootstrap failure when wiring Dashboard buttons by qualifying `OnAction` with the workbook name and falling back to `Shape.OnAction` if needed. Keeps button behavior the same and avoids the "Unable to set the OnAction property of the Button class" error during `SeedNamedValues`.
 - [2025-11-02] On `feature/consolidate-flow-into-jira-insights`: consolidated Flow Metrics (WIP Aging, Sprint Spans, Throughput, Cycle Scatter) into `Jira_Insights` instead of a separate `Flow_Metrics` tab. Added `Flow_AppendChartsToSheet` and wired it into `Jira_CreatePivotsAndCharts` and `SanitizeRawAndBuildInsights` (WIP branch). No sheet clears on Jira_Insights.
