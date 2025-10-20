# Schema (Minimal Baseline)

Sheets
- Config_Teams: table `tblRoster`
- Config_Sprints: config + named values (labels on sheet; can be hidden later)
- Logs: table `tblLogs`

Tables
- tblRoster: Member, Role, ContributesToVelocity
- tblLogs: Timestamp, User, Action, Outcome, Details

Named values
- ActiveTeam, TemplateVersion, SprintLengthDays, DefaultHoursPerDay, DefaultAllocationPct, DefaultHoursPerPoint, RolesWithVelocity

Planned (later)
- Calendars sheet with `tblHolidays` and `tblPTO`
- Sprint_Capacity_[Team]: Date, Team, Member, HoursPerDay, NetHours, Points
