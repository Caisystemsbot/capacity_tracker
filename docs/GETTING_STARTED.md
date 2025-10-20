# Getting Started

This minimal baseline lets you set up your workbook fast and safely.

Steps
- Open a blank `.xlsm` workbook in Excel.
- Alt+F11 → Insert → Module → paste the contents of `src/modules/modCapacityPlanner.bas`.
- Run `Bootstrap`.
- Open sheet `Getting_Started` and follow the checklist:
- Fill `Config_Teams` → `tblRoster` with your team. Columns:
  - `Member` (person’s name)
  - `Role` (dropdown: QA, Developer, Analyst, Squad Leader, Project Manager)
  - `ContributesToVelocity` (Yes/No). Developer/QA = Yes; Analyst/SL/PM = No.
- Only roles in `RolesWithVelocity` contribute to velocity by default (Developer, QA). `ContributesToVelocity` lets you override per person if needed.
  - Holidays/PTO are deferred in this minimal baseline and will be added later.
- Run `HealthCheck` to validate structure.

Named values (Config_Sprints H2:H8)
- `ActiveTeam` — the team you’re working with.
- `TemplateVersion` — informational.
- `SprintLengthDays` — default sprint length in working days.
- `DefaultHoursPerDay` — baseline daily hours per person.
- `DefaultAllocationPct` — fraction of daily hours allocated to sprint work.
- `DefaultHoursPerPoint` — hours-to-points conversion.
- `RolesWithVelocity` — comma-separated roles that count toward velocity (defaults to Developer, QA).

Next
- Keep working from the `Getting_Started` checklist until your roster and calendars are ready.
- When you’re ready for capacity planning, we’ll add `CreateNextSprint_Simple` that builds a per-day capacity sheet using your roster, holidays, and PTO.
