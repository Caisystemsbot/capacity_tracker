## Jira Integration (No Tokens)

Overview
- Uses Excel Power Query (Get Data from Web) — not API tokens — to pull Jira sprint data.
- Macros auto-create Power Query queries, load a `Jira_Metrics` sheet, then copy Committed/Completed into the `Metrics` table.

Queries created by `Jira_CreateQueries`
- `JiraSprints` (table): GET `{JiraBaseUrl}/rest/agile/1.0/board/{JiraBoardId}/sprint?state=active,closed&maxResults=50`
- `JiraSprintReport` (function): GET `{JiraBaseUrl}/rest/greenhopper/1.0/rapid/charts/sprintreport?rapidViewId={JiraBoardId}&sprintId={id}` and returns Committed/Completed
- `JiraSprintMetrics` (table): Joins `JiraSprints` + per-sprint metrics
- Loads to sheet `Jira_Metrics` as table `tblJiraMetrics`

Config (sheet `Config`)
- `JiraBaseUrl` (H10): e.g., `https://your-domain.atlassian.net`
- `JiraBoardId` (H13): Jira board/rapidView id
- Notes: `JiraEmail` and `JiraApiToken` are unused in this Power Query flow and can be left blank.

How to run
1) Set `JiraBaseUrl` and `JiraBoardId` on `Config`.
2) Run macro `Jira_PopulateMetrics`.
   - First refresh: Excel prompts you to sign in (Organizational/Work account). Approve and store credentials.
   - The macro refreshes, then copies `Committed`/`Completed` into `Metrics` columns E/D by sprint.

Matching behavior
- Metrics row is matched by the `Sprint` text. Ensure your Metrics `Sprint` values match Jira sprint names, or we can switch to a date-range based match in a follow-up.

Limitations & Notes
- Some tenants may block the (legacy) Sprint Report endpoint. If so, we can pivot to Board Reports or updated endpoints.
- Large boards: increase `maxResults` in the `JiraSprints` query or add paging.
- Security: Credentials are stored by Excel’s Data Source settings, not in the workbook.

Offline testing (sample data)
- Import `data/jira_metrics_sample.csv` to a sheet named `Jira_Metrics` as table `tblJiraMetrics`.
- Run `Jira_ApplyMetricsFromQuery` to copy the sample Committed/Completed into your `Metrics` sheet.

Troubleshooting
- If `Jira_Metrics` is empty after refresh, open Data → Queries & Connections and refresh `JiraSprints` to complete auth.
- If sprint name mismatches prevent copying, temporarily type the Jira sprint `name` into the Metrics `Sprint` column for the rows you want to update.

