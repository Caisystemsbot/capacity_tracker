# Python Flow Metrics

Run flow metrics from a CSV (Jira export or the Excel Raw_Data you already use) and generate nicer charts.

## Install

```
pip install -r requirements.txt
```

## Usage

```
# With explicit CSV
python tools/flow_metrics.py --input path/to/raw.csv --out out

# Or just run (auto-uses data/raw_data_example.csv if present)
python tools/flow_metrics.py --out out
```

Outputs:
- `out/jira_facts.csv` – normalized facts with columns:
  `IssueKey, Summary, IssueType, Status, Epic, Created, StartProgress, Resolved,
   StoryPoints, CycleDays, SprintSpan, IsCrossSprint, QuarterTag, YearTag,
   FixVersion, CreatedMonth, CycleCalDays, LeadCalDays`.
- `out/cycle_time_scatter.png` – Cycle Time scatter with dashed P50/P70/P85/P95 guides and P85+ points in red.
- `out/throughput.png` – throughput per day (bars) with a 7‑day rolling average line.
- `out/cfd.png` – Cumulative Flow Diagram (To Do / In Progress / Done).

## Input columns (flexible)
The script maps common Jira headers case‑insensitively. Useful names:
- Created, Start Progress, Resolved
- Issue key, Summary, Issue Type, Status, Fix Version/s, Epic Link
- Story Points
- Optional: Time In Todo / Progress / Testing / Review (days) – used to infer Start when missing.

If `Start Progress` is missing but `Time In Todo` exists, the script sets `StartProgress = Created + TimeInTodo`.

If `--input` isn’t provided, the tool looks for (in order):
- `data/raw_data_example.csv`
- `data/jira_metrics_sample.csv`
- `data/wip_example.csv`

## Notes
- `CycleDays` uses business days (Mon‑Fri). `CycleCalDays` uses calendar days.
- CFD limits the time window to ~90 days if the range exceeds ~120 days for readability.
- Use the same CSV you would paste into `Raw_Data` in Excel; the script is defensive about header variations.
