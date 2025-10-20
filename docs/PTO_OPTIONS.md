# PTO Integration Options

Option A — Manual/CSV (baseline)
- Enter PTO rows directly into `Calendars` → `tblPTO`.
- Or use `ImportPTO_CSV` to append rows from a CSV with columns:
  - Team, Member, Date, Hours, Source

Option B — Outlook Object Model (desktop Outlook)
- Reads events from a shared PTO calendar or members’ calendars filtered by Category/Subject (e.g., PTO/Vacation/OOO).
- Requires:
  - Outlook installed and signed in.
  - Excel Trust Center: “Trust access to the VBA project object model”.
  - Consent to read the selected calendar.
- Behavior:
  - For each event matching rules within a date window: write Team, Member, Date, Hours, Source="Outlook" into `tblPTO`.

Option C — Microsoft Graph (advanced)
- Pull calendar events via Graph API; robust and tenant-friendly.
- Requires Azure AD app registration, scopes, admin/consented permissions, token flow.
- Excel VBA can call Graph via MSAL or Power Query connector.

Recommendation
- Start with CSV/manual (`Option A`).
- Add Outlook OOM if your org uses a shared PTO calendar and accepts desktop automation.
- Consider Graph only if you need cross-tenant scale or service automation.
