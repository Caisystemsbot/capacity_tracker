# Schema (Minimal Baseline)

Sheets
- Config_Teams: table `tblRoster`
- Calendars: tables `tblHolidays`, `tblPTO`
- Config_Sprints: config + named values
- Logs: table `tblLogs`

Tables
- tblRoster: Team, Member, Role, HoursPerDay, AllocationPct
- tblHolidays: Date, Region, Name
- tblPTO: Team, Member, Date, Hours, Source
- tblLogs: Timestamp, User, Action, Outcome, Details

Named values
- ActiveTeam, TemplateVersion, SprintLengthDays, DefaultHoursPerDay, DefaultAllocationPct, DefaultHoursPerPoint
