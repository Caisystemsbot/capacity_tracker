# Capacity Planner – Agent Guide (AGENTS.md)

This file provides repo-wide instructions for any agent or contributor working on the Capacity Planner (Excel + VBA + Jira + Outlook) project. Read this fully before making changes. Its scope is the entire repository.

## 1) Ground Rules (Must-Follow)

- Always create a branch for any work. Never commit directly to `main`.
- Ask before pushing: open a PR and request review from the repo owner; do not push or merge without explicit approval.
- New features require explicit permission: do not add features beyond those requested by the owner.
- Keep changes minimal and focused; avoid unrelated refactors.
- Prefer text-form VBA (`.bas/.cls/.frm`) in `src/` as the source of truth. The `.xlsm/.xlam` files are built artifacts.
- Be cautious with destructive actions (deleting sheets, modifying Power Query, or changing connections). Confirm with the owner first.
- Share concise preambles before running commands and provide short progress updates during multi-step work.

## 2) Current Status (Continuity)

- Strategy, schema, metrics, and VBA skeletons have been defined by the owner in planning notes (this conversation). The repository itself is currently minimal.
- This AGENTS.md has been added to establish collaboration rules and shared context.
- Pending owner approval, the next step is to scaffold repo structure, import/export scripts, and initial VBA module shells.

## 3) Objectives (What We Are Building)

- One‑click capacity planning per sprint for multiple teams (CraicForce & PsiForce initially).
- PTO/holiday-aware per‑member capacity, velocity-based forecasting, Monte Carlo, and Jira-backed progress stats.
- Clean dashboard with status flair, sparklines, and automated `Workbook_Open` refresh/build behaviors.

## 4) Planned Repository Layout (Intended)

```
capacity-planner/
├─ template/
│  ├─ CapacityPlanner_template.xlsm      # primary workbook (signed for release)
│  └─ CapacityPlanner_addin.xlam         # optional add-in
├─ src/                                  # VBA in text form (source of truth)
│  ├─ modules/
│  │  ├─ modBootstrap.bas
│  │  ├─ modJira.bas
│  │  ├─ modOutlook.bas
│  │  ├─ modCalendar.bas
│  │  ├─ modSprint.bas
│  │  └─ modUI.bas
│  ├─ classes/
│  │  └─ clsTeam.cls
│  └─ forms/
│     └─ frmCreateTeam.frm
├─ data/                                 # seed/sample data
│  ├─ holidays.csv
│  ├─ roster_example.csv
│  └─ pto_example.csv
├─ config/
│  ├─ pq_parameters.json                 # Power Query parameter names/ids
│  └─ defaults.json                      # defaults (hours/point, allocation, etc.)
├─ scripts/
│  ├─ export_modules.ps1                 # workbook → src (text)
│  ├─ import_modules.ps1                 # src (text) → workbook
│  └─ bump_version.ps1
├─ docs/
│  └─ README.md
└─ CHANGELOG.md
```

Notes:
- Keep `.xlsm/.xlam` out of frequent diffs; Git diffs are tracked on text modules under `src/`.
- Use `scripts/import_modules.ps1` and `scripts/export_modules.ps1` to sync code with the workbook. Excel must have “Trust access to the VBA project object model” enabled (Trust Center).

## 5) Branch & PR Workflow

- Create a branch per task: `feature/<short-desc>` or `chore/<short-desc>` or `fix/<short-desc>`.
- Open a PR early and convert to “Ready for review” only after local validation.
- Always request explicit approval from the owner before merging or pushing artifacts.
- Never merge to `main` without the owner’s approval.

## 6) Work Phases (from the plan)

1. Build workbook skeleton + tables + two teams’ rosters.
2. Wire Jira Power Query + test refresh.
3. Implement PTO import (start with CSV; then Outlook OOM).
4. Implement CreateNextSprint + capacity math.
5. Add velocity calc + Monte Carlo + dashboard.
6. Add Workbook_Open automation + buttons.
7. Polish flair, locks, and error handling.

## 7) Immediate Next Tasks (Backlog)

- [ ] Scaffold repo layout as in Section 4.
- [ ] Add PowerShell scripts: `import_modules.ps1`, `export_modules.ps1`, `bump_version.ps1`.
- [ ] Create initial empty VBA module shells in `src/modules` and `src/classes`.
- [ ] Add `docs/README.md` with quick start and trust-center instructions.
- [ ] Add `data/holidays.csv` seed and example CSVs for PTO and roster.
- [ ] Prepare `template/CapacityPlanner_template.xlsm` placeholder; wire import script.
- [ ] Confirm Jira Power Query parameterization approach (board/filter per team).
- [ ] Add “Health Check” macro skeleton and `Logs` sheet plan.

All of the above require owner approval before pushing to the remote or publishing artifacts/releases.

## 8) Coding & Data Conventions

- VBA style: explicit `Option Explicit`, clear procedure names, no single-letter vars, no hard-coded cell addresses—use tables and named ranges.
- Data tables: `tblRoster`, `tblHolidays`, `tblPTO`, `tblVelocity`, plus `tblSprintDays_[Team]`.
- Named constants: `HoursPerPoint_[Team]`, `SprintLengthDays`, `RetroDay`, `PlanningDay`.
- Security boundaries: limit Jira queries to team board/filter; PTO via shared calendar or CSV; do not access personal calendars without authorization.

## 9) Excel/Outlook/Jira Integration Notes

- Excel Trust Center: enable “Trust access to the VBA project object model” for import/export scripts.
- Outlook OOM: requires Outlook installed and user authenticated; fallback path is CSV import.
- Power Query: parameterize by team; refresh via button and `ActiveWorkbook.RefreshAll`.

## 10) Validation & Testing

- Validate capacity numbers against a hand-calculated sprint for one team before broadening scope.
- Keep a `Logs` sheet for actions/outcomes/errors; ensure graceful failure messages to users.
- Avoid running long tasks silently; use progress updates and clear completion toasts.

## 11) Communication Protocol

- Before starting multi-step work, outline a short plan and get owner confirmation where scope may expand.
- Share concise progress updates and ask for permission before any significant change (new features, schema changes, or migration of data).
- If unsure about data boundaries or access, pause and request guidance.

## 12) Running Work Log (to be appended by contributors)

- [YYYY-MM-DD] Added AGENTS.md with collaboration rules, status, and backlog. Awaiting owner approval to scaffold repo and add scripts/modules.

---

Owner approvals are required for merges to `main`, pushes of binary artifacts, and any features outside the requested scope.

