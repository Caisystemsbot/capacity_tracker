# Capacity Planner (Excel + VBA + Jira + Outlook)

This repo hosts the source for a sprint capacity planning workbook. VBA code lives as text in `src/` and is imported into the Excel template under `template/` via PowerShell scripts.

## Quick Start (Local)

- Prereqs: Windows + Excel desktop, PowerShell, Git.
- Excel Trust Center: enable "Trust access to the VBA project object model".

### 1) Structure

- `src/` holds `.bas/.cls/.frm` files (source of truth).
- `template/` contains the `.xlsm` workbook (built artifact).
- `scripts/` contains PowerShell import/export helpers.
- `data/` sample CSVs for holidays, roster, PTO.
- `config/` JSON for defaults and Power Query parameter names.

### 2) Import modules into the workbook

- Place or create your workbook at `template/CapacityPlanner_template.xlsm`.
- Run: `powershell -ExecutionPolicy Bypass -File scripts/import_modules.ps1`

### 3) Export modules from the workbook

- After editing inside the VBA editor, sync back to text:
- Run: `powershell -ExecutionPolicy Bypass -File scripts/export_modules.ps1`

### Notes

- Work in a feature branch; never commit to `main`.
- Ask before pushing or merging.
- Do not add unapproved features.

## Where to start

- See `AGENTS.md` for collaboration rules and backlog.
- Edit VBA under `src/` and use the scripts to sync with Excel.

## Paste-once Installer (no bulk pasting)

If you don’t want to paste every module, use the single installer module:

1) Open your target workbook in Excel, press `Alt+F11`.
2) Insert → Module, then paste the contents of `src/modules/modInstaller.bas`.
3) In the VBE, run `InstallFromFolder` and choose your repo root (the folder that contains `src`).

The installer will remove existing standard modules/classes/forms (keeps Sheets/ThisWorkbook and itself) and import everything from `src/modules`, `src/classes`, and `src/forms`.

Note: In Excel Trust Center, enable “Trust access to the VBA project object model”.
