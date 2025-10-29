# Flow Metrics Charts (Pasteable)


## Aging Work in Progress (AA-aligned)

What it shows
- Active items only by default (unresolved). If there are no active items, it shows a historical view at the time of resolution.
- X-axis = current workflow step (1=To Do, 2=In Progress, 3=Testing, 4=Review). A lane caption is shown below the chart and a “WIP: n” label appears per lane.
- Y-axis = Work Item Age in calendar days: Now − StartProgress. When `Time In ...` columns exist, age = `In Progress + Testing + Review` (excludes `To Do`). If `StartProgress` is missing, the chart falls back to Now − Created.
- Background bands = Service Level Expectation (SLE) from completed items’ `CycleCalDays` with percentiles P50 (green), P70 (yellow), P85 (orange), and red for anything above P85. Horizontal guide lines at P50/P70/P85 are overlaid.

How to read it
- Items high in earlier lanes indicate flow risk — more likely to miss the team’s SLE (orange/red). The same age in later lanes is typically less risky.
- The SLE bands come from your historical cycle times, so they adapt as your flow changes.
- Aim to keep most points green/yellow. Items drifting into orange/red — especially in earlier lanes — deserve attention.
Overview
- Builds Scrum flow metrics directly in Excel from a sanitized facts table.
- Creates a new sheet `Flow_Metrics` with three charts:
  - Cumulative Flow Diagram (stacked To Do / In Progress / Done)
  - Throughput Run Chart (completed items per day)
  - Cycle Time Scatter (Resolved date vs cycle time in days)
  - WIP Aging bubble chart (items by stage vs age)
  - WIP Aging bubble chart (items by stage vs age)
- Defensive by design: if a needed column is missing, the affected chart is skipped.

Data assumptions
- Preferred source: sheet `Jira_Facts` with table `tblJiraFacts`.
- Auto-detection: if not found, the macro scans all tables and picks one that has:
  - Required: `Created`
  - Plus one of: `Resolved` or `CycleCalDays` (resolved still required for charts)
  - Optional: `StartProgress` to compute the In Progress band for the CFD

Column usage
- `Created` (date): work item creation date
- `StartProgress` (date, optional): first time work started
- `Resolved` (date, optional): completion date
- `CycleCalDays` (number, optional): calendar-day cycle time; falls back to Resolved - Created

How to run
- Paste or import the module `src/modules/modCapacityPlanner.bas` into your .xlsm.
- Run the macro `Flow_BuildCharts`.
- The sheet `Flow_Metrics` will be created or refreshed with data blocks and charts.
- WIP Aging is built first and includes a visible "WIP Aging – Data" table (IssueKey, Stage, AgeDays, Created, Resolved) before the chart.

Notes
- CFD (WIP) uses dates from `Created` to max(`Resolved`) and rolls daily counts:
  - To Do: Created ≤ d and (StartProgress > d or StartProgress missing)
  - In Progress: StartProgress ≤ d and (Resolved > d or missing)
  - Done: Resolved ≤ d
- If the range exceeds ~120 days, it limits the CFD window to ~90 recent days for readability.
- Throughput is a simple count of items by `Resolved` date.
- Cycle Time scatter prefers `CycleCalDays`; otherwise uses `Resolved - Created`.

WIP Aging
- Primary: plots unresolved items using `Time In Todo/Progress/Testing/Review` (days). Age = sum of those columns; lane = last non-zero stage.
- No unresolved fallback: when all items are resolved, it still renders using historical ages at resolution so the chart always appears.
- No time-in-status fallback: if `Time In ...` columns are missing, it derives durations from dates when possible (`To Do` ≈ StartProgress - Created; `In Progress` ≈ Resolved/Today - StartProgress).
 - Data table: the sheet writes a "WIP Aging – Data" block to make source rows explicit before plotting.
