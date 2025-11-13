Bug Metrics

Overview
- Adds sprint- and quarter-level bug metrics to `Jira_Insights`.
- Default basis shows both counts: Created and Resolved.
- Sprint view also graphs running bug backlog (Created minus Resolved) on a secondary axis.

How it works
- Classification: A work item is a “bug” if its Issue Type matches any value in the name `BugIssueTypes` (case‑insensitive, comma‑separated). Default: `Bug,Defect`.
- Sprint grouping: Dates are mapped to sprint tags using `SprintNamePattern` via `FormatSprintName(date)`.
  - Created count uses Created date; Resolved count uses Resolved date (skips unresolved).
- Quarter grouping: Labels are `YYYY Q#` computed from Created and Resolved dates respectively.
- Backlog line: Running sum of (Created − Resolved) across the sprints present in the facts table (relative trend; not an absolute product backlog).

Configuration (optional)
- `BugCountBasis` (H11 on `Config`): `Both` (default), `Created`, or `Resolved`.
- `BugIssueTypes` (H12 on `Config`): Comma list of bug types, e.g. `Bug,Defect,Production Issue`.

Where it appears
- After Story Point and Quarter pivots in `Jira_Insights`.
  - Tables: “Bug Metrics – Sprint” and “Bug Metrics – Quarter”.
  - Charts: “Bugs per Sprint (Created vs Resolved)” with backlog line; “Bugs per Quarter (Created vs Resolved)”.

Edge cases
- If there are no bug rows, the sections are skipped.
- If only Created or only Resolved exists (per `BugCountBasis`), the other series is omitted.
- Sorting is by label text to avoid assumptions about `SprintNamePattern`; results remain readable and stable.

