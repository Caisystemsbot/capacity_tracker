# Capacity Tracker – Agent Guide

Ground rules
- Always create a branch. Never commit to `main`.
- Ask before pushing or merging. PRs only with owner approval.
- Keep changes minimal and focused. No unrequested features.
- Update docs as you go (README, SCHEMA, CHANGELOG, this log).

Working model
- Paste-once single VBA module for a minimal, working baseline.
- No classes or forms initially to avoid VBE import issues.
- CSV-first for PTO and Holidays. Outlook/Jira later.

Current status
- Fresh restart. Minimal baseline to be added on `feature/restart`.

Next actions
- Add `modCapacityPlanner.bas` (single module) with:
  - `Bootstrap` to create sheets/tables/names and seed sample CSVs (if present).
  - `ImportPTO_CSV` to append PTO from a selected CSV.
  - `HealthCheck` to validate required pieces.
- Provide simple docs and sample CSVs.

Running work log
- [YYYY-MM-DD] Repo reset to empty. Planning minimal baseline on `feature/restart`.

