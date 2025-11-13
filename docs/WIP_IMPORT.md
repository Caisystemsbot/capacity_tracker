# WIP CSV Import (Time-in-Status)

Purpose
- Import and sanitize a CSV that contains time-in-status durations per item for WIP analysis.
- Writes to sheet `WIP_Facts` as table `tblWIPFacts`.

Expected columns (flexible names)
- Created: e.g., `09/09/2025 12:15`
- Resolved: e.g., `09/12/2025 18:30` (optional)
- Time In Todo: days (decimal), e.g., `0.05`, `1.21`
- Time In Progress: days (decimal)
- Time In Testing: days (decimal)
- Time In Review: days (decimal)
- Optional passthrough if present: `Issue key`, `Summary`

Output columns
- IssueKey, Summary, Created, Resolved
- TimeInTodo, TimeInProgress, TimeInTesting, TimeInReview
- WIPTotalDays (sum of the time-in-status columns)
- CycleCalDays (Resolved - Created, in calendar days, when both exist)

How to run
- Macro: `WIP_ImportCSV`
- You will be prompted to pick the CSV path.
- Data is written to `WIP_Facts!tblWIPFacts`. Existing rows are cleared.

Sample CSV
- Included at `data/wip_example.csv`.

Notes
- Column names are matched case-insensitively with forgiving variants (e.g., `Time In To Do` vs `Time In Todo`).
- Nonexistent columns are treated as zero duration; the import never crashes due to missing optional columns.
- If `Resolved` is blank, `CycleCalDays` remains blank for that row.
