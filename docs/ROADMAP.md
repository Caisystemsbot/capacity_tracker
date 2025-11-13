# Roadmap (Minimal → Full)

Phase 0 — Baseline (done)
- Single module (`modCapacityPlanner`): Bootstrap, HealthCheck, ImportPTO_CSV.
- Tables: tblRoster, tblHolidays, tblPTO, tblLogs.
- Getting_Started sheet and named defaults (`RolesWithVelocity` = "Dev,QA").

Phase 1 — Capacity (next)
- `CreateNextSprint_Simple(team, startDate, endDate)`
  - Build working days (Mon–Fri), exclude `tblHolidays`.
  - Join roster and PTO to compute per‑member per‑day:
    - `NetHours = DefaultHoursPerDay × AllocationPct − PTODeduct`
    - `Points = NetHours ÷ DefaultHoursPerPoint`
  - Output table `tblSprintCapacity` on sheet `Sprint_Capacity_[Team]` (Date, Team, Member, HoursPerDay, AllocationPct, PTOHours, NetHours, Points).
- Respect `ContributesToVelocity` (per person) and `RolesWithVelocity` default list.

Phase 2 — Sprint Ops
- `CreateNextSprint(team)` wrapper to resolve dates from `Config_Sprints`.
- Snapshot roster to keep historical allocations stable.
- Minimal spillover handling placeholder.

Phase 3 — Velocity + Forecast
- Add `Jira_Raw` PQ connector (parameterized by team), build `tblVelocity` rollups.
- `RunMonteCarlo(team, trials)` using last N `CompletedPts`.
- Named outputs for p50/p85/p95.

Phase 4 — Automation + UI
- `Workbook_Open`: refresh PQ, optional PTO import, rebuild velocity, and create sprint if due.
- Buttons for: Refresh Jira, Import PTO, Recalc Capacity, Run Monte Carlo, Create Next Sprint.
- Health Check enhancements + Logs.

Phase 5 — Outlook PTO (optional)
- Outlook Object Model import (shared PTO calendar or member calendars filtered by category/subject).
- Graph integration (advanced, requires tenant app and consent).

Principles
- Keep everything config‑driven in tables and named ranges.
- Only Dev/QA contribute to velocity by default; make it configurable via `RolesWithVelocity`.
- Ship minimal, working increments; style later.
