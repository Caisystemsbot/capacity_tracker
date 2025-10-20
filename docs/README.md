# Capacity Tracker (Minimal Baseline)

This restart provides a robust, minimal baseline you can run immediately:
- One paste-once module: no classes/forms; avoids VBE import quirks.
- CSV-first PTO and Holidays; Outlook/Jira added later.

Quick start
- Open a blank .xlsm workbook in Excel.
- Alt+F11 → Insert → Module.
- Paste the contents of `src/modules/modCapacityPlanner.bas`.
- Run `Bootstrap` to create sheets/tables/named ranges.
- Run `ImportPTO_CSV` to append PTO rows from a CSV.

Sheets and tables
- Config_Teams: `tblRoster` (Team, Member, Role, HoursPerDay, AllocationPct)
- Calendars: `tblHolidays` (Date, Region, Name), `tblPTO` (Team, Member, Date, Hours, Source)
- Config_Sprints: future use; holds named values
- Logs: `tblLogs` (Timestamp, User, Action, Outcome, Details)

Named values
- `ActiveTeam`, `TemplateVersion`, `SprintLengthDays`, `DefaultHoursPerDay`, `DefaultAllocationPct`, `DefaultHoursPerPoint`

Sample data
- `data/holidays.csv`, `data/pto_example.csv`, `data/roster_example.csv`

Troubleshooting
- If macros don’t appear, press Alt+F11 and confirm the module is present.
- If FileDialog isn’t available, the PTO import prompts for a path.

