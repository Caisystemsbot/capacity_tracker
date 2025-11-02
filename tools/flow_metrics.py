#!/usr/bin/env python3
"""
Flow Metrics from CSV (Python)

Reads a Jira-like CSV export (or the Excel Raw_Data CSV), normalizes it to a
Jira Facts schema, and produces flow metrics charts:
  - Cycle Time Scatter (with percentile lines and P85 outliers in red)
  - Throughput (daily bars + 7d rolling average)
  - Cumulative Flow Diagram (To Do / In Progress / Done)

Usage:
  python tools/flow_metrics.py --input path/to/raw.csv --out out

Requirements: pandas, numpy, matplotlib, seaborn
Install: pip install -r requirements.txt
"""
from __future__ import annotations

import argparse
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


# ---------------------- Header mapping ----------------------

HeaderSynonyms: Dict[str, List[str]] = {
    "IssueKey": ["issue key", "key", "issuekey", "id", "ticket"],
    "Summary": ["summary", "title"],
    "IssueType": ["issue type", "type"],
    "Status": ["status"],
    "Epic": ["parent", "parent link", "parent id", "parent key", "epic link", "epic"],
    "Created": ["created", "created date", "created on", "created_date"],
    "Resolved": ["resolved", "resolved date", "done date", "resolution date", "closed"],
    "StartProgress": ["start progress", "startprogress", "started", "in progress date", "start date"],
    "StoryPoints": ["story points", "story point", "story point estimate", "custom field (story points)"],
    "FixVersion": ["fix version/s", "fix version"],
    # time in status (days)
    "TimeInTodo": ["time in todo", "time in to do", "in todo", "todo days", "timeintodo"],
    "TimeInProgress": ["time in progress", "in progress", "in progress days", "timeinprogress"],
    "TimeInTesting": ["time in testing", "in testing", "testing days", "timeintesting"],
    "TimeInReview": ["time in review", "in review", "review days", "timeinreview"],
}


def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    for ch in ["_", "-", "/", "(", ")", ":"]:
        s = s.replace(ch, " ")
    s = " ".join(s.split())
    return s


def build_header_map(cols: List[str]) -> Dict[str, str]:
    inv = { _norm(c): c for c in cols }
    mapping: Dict[str, str] = {}
    for std, syns in HeaderSynonyms.items():
        for s in syns:
            k = _norm(s)
            # exact then contains
            if k in inv:
                mapping[std] = inv[k]
                break
            # contains search
            found = None
            for nk, orig in inv.items():
                if k in nk or nk in k:
                    found = orig
                    break
            if found is not None:
                mapping[std] = found
                break
    return mapping


def to_date(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce", utc=False).dt.tz_localize(None)


def workdays_between(s: pd.Series, e: pd.Series) -> pd.Series:
    # naive business days counting (Mon-Fri). If NaT -> NaN
    s = pd.to_datetime(s, errors="coerce")
    e = pd.to_datetime(e, errors="coerce")
    mask = s.notna() & e.notna()
    out = pd.Series(np.nan, index=s.index, dtype=float)
    if mask.any():
        sd = s[mask].dt.normalize()
        ed = e[mask].dt.normalize()
        # vectorized: count business days inclusive
        b = np.busday_count(sd.values.astype('datetime64[D]'), (ed + pd.Timedelta(days=1)).values.astype('datetime64[D]'))
        out.loc[mask] = b
    return out


@dataclass
class Facts:
    df: pd.DataFrame


def build_facts(df_raw: pd.DataFrame) -> Facts:
    cols = list(df_raw.columns)
    hmap = build_header_map(cols)
    g = lambda k: hmap.get(k)

    def col(k: str) -> pd.Series:
        if g(k) in df_raw.columns:
            return df_raw[g(k)]
        return pd.Series([np.nan] * len(df_raw))

    created = to_date(col("Created"))
    start = to_date(col("StartProgress"))
    resolved = to_date(col("Resolved"))

    # Derive StartProgress from TimeInTodo if missing
    if start.isna().all() and g("TimeInTodo"):
        t_todo = pd.to_numeric(df_raw[g("TimeInTodo")], errors="coerce")
        start = created + pd.to_timedelta(t_todo.fillna(0), unit="D")

    story_points = pd.to_numeric(col("StoryPoints"), errors="coerce")

    cycle_cal_days = (resolved - created).dt.days
    cycle_days = workdays_between(created, resolved)

    sprint_len = 10.0
    span = np.ceil(cycle_days.fillna(0) / sprint_len)

    ref_date = resolved.fillna(created)
    quarter = ref_date.dt.to_period("Q").astype(str).str.replace("Q", " Q", regex=False)
    year = ref_date.dt.year
    created_month = created.dt.to_period("M").dt.to_timestamp()

    facts = pd.DataFrame({
        "IssueKey": col("IssueKey"),
        "Summary": col("Summary"),
        "IssueType": col("IssueType"),
        "Status": col("Status"),
        "Epic": col("Epic"),
        "Created": created,
        "StartProgress": start,
        "Resolved": resolved,
        "StoryPoints": story_points,
        "CycleDays": cycle_days,
        "SprintSpan": span,
        "IsCrossSprint": span > 1,
        "QuarterTag": year.astype(str) + " " + quarter.str[-2:],
        "YearTag": year,
        "FixVersion": col("FixVersion"),
        "CreatedMonth": created_month,
        "CycleCalDays": cycle_cal_days,
        "LeadCalDays": cycle_cal_days,  # simple placeholder
    })

    return Facts(facts)


# ---------------------- Charts ----------------------

def _nice_ceiling(v: float) -> float:
    if v <= 0:
        return 1
    if v > 20:
        step = 5
    elif v > 10:
        step = 2
    else:
        step = 1
    return step * np.ceil(v / step)


def _quantiles(vals: np.ndarray, ps: List[float]) -> List[float]:
    if len(vals) == 0:
        return [np.nan] * len(ps)
    qs = []
    for p in ps:
        qs.append(np.quantile(vals, p, method="linear"))
    return qs


def plot_cycle_scatter(facts: Facts, out_dir: str) -> None:
    df = facts.df.dropna(subset=["Resolved", "CycleCalDays"])
    if df.empty:
        return
    y = df["CycleCalDays"].astype(float).values
    x = pd.to_datetime(df["Resolved"]).values
    p50, p70, p85, p95 = _quantiles(y[~np.isnan(y)], [0.5, 0.7, 0.85, 0.95])
    mask_red = y > (p85 if not np.isnan(p85) else np.inf)

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.scatter(x[~mask_red], y[~mask_red], s=32, c="#2D62A3", label="≤P85")
    if mask_red.any():
        ax.scatter(x[mask_red], y[mask_red], s=32, c="#C00000", label=">P85")

    # Percentile lines
    for val, label, color in [
        (p50, "50%", "#808080"),
        (p70, "70%", "#BF9000"),
        (p85, "85%", "#C00000"),
        (p95, "95%", "#800000"),
    ]:
        if not np.isnan(val):
            ax.axhline(val, ls="--", lw=1.25, c=color, alpha=0.9)

    ax.set_title("Cycle Time Scatter")
    ax.set_xlabel("Completion Date")
    ax.set_ylabel("Cycle Time (days)")
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(ax.xaxis.get_major_locator()))

    ymax = _nice_ceiling(max(float(np.nanmax(y)), float(p95) if not np.isnan(p95) else 0.0))
    ax.set_ylim(0, ymax)
    ax.grid(True, axis="y", alpha=0.3)
    # Right-side percentile labels
    x_left, x_right = ax.get_xlim()
    if len(x) > 0:
        xr = np.max(mdates.date2num(pd.to_datetime(df["Resolved"])) )
    else:
        xr = x_right
    # pad right by ~3% of range for labels
    xpad = (x_right - x_left) * 0.03
    ax.set_xlim(x_left, x_right + xpad)
    for val, label, color in [
        (p50, "50%", "#808080"),
        (p70, "70%", "#BF9000"),
        (p85, "85%", "#C00000"),
        (p95, "95%", "#800000"),
    ]:
        if not np.isnan(val):
            ax.text(x_right + xpad * 0.2, val, label, va="center", ha="left", color=color, fontsize=8, clip_on=False)
    ax.legend()

    os.makedirs(out_dir, exist_ok=True)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "cycle_time_scatter.png"), dpi=150)
    plt.close(fig)


def plot_throughput(facts: Facts, out_dir: str) -> None:
    df = facts.df.copy()
    df = df.dropna(subset=["Resolved"])  # completed only
    if df.empty:
        return
    res = pd.to_datetime(df["Resolved"]) 
    # Week-of-Sunday label
    week_start = res.dt.to_period("W-SUN").apply(lambda p: p.start_time)
    weekly = week_start.value_counts().sort_index()
    weeks = weekly.index.to_pydatetime()
    vals = weekly.values.astype(float)
    # 4-week rolling average (align to week)
    ma = pd.Series(vals).rolling(4, min_periods=1).mean().values

    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.plot(weeks, vals, "-o", color="#2D2D2D", markerfacecolor="#2D62A3", markeredgecolor="#2D62A3", label="Weekly Completed")
    ax.plot(weeks, ma, linestyle="--", color="#00B050", linewidth=2, label="4w Avg")
    ax.set_title("Throughput Run Chart")
    ax.set_xlabel("Week of")
    ax.set_ylabel("Weekly Throughput")
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(ax.xaxis.get_major_locator()))
    ax.grid(True, axis="y", alpha=0.3)
    ax.legend()

    os.makedirs(out_dir, exist_ok=True)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "throughput.png"), dpi=150)
    plt.close(fig)


def plot_cfd(facts: Facts, out_dir: str) -> None:
    df = facts.df.copy()
    if df["Created"].isna().all():
        return
    df["Created_d"] = pd.to_datetime(df["Created"]).dt.normalize()
    df["Resolved_d"] = pd.to_datetime(df["Resolved"]).dt.normalize()
    df["Start_d"] = pd.to_datetime(df["StartProgress"]).dt.normalize()

    # Estimate Start from Created + TimeInTodo when missing
    if df["Start_d"].isna().all():
        # try to infer if time-in-todo exists in the raw frame (kept via Facts?)
        pass  # already handled during build_facts

    min_d = df["Created_d"].min()
    max_d = pd.concat([df["Created_d"], df["Resolved_d"]]).max()
    if pd.isna(min_d) or pd.isna(max_d):
        return
    # limit to last ~120 days
    if (max_d - min_d).days > 120:
        min_d = max_d - pd.Timedelta(days=90)
    days = pd.date_range(min_d, max_d, freq="D")

    todo = []
    inprog = []
    done = []
    for d in days:
        # counts at end of day d
        r_le = (df["Resolved_d"].notna()) & (df["Resolved_d"] <= d)
        s_le = (df["Start_d"].notna()) & (df["Start_d"] <= d)
        c_le = (df["Created_d"].notna()) & (df["Created_d"] <= d)

        done.append(int(r_le.sum()))
        inprog.append(int((s_le & ~r_le).sum()))
        todo.append(int((c_le & ~s_le & ~r_le).sum()))

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.stackplot(days, [todo, inprog, done], labels=["To Do", "In Progress", "Done"],
                 colors=["#BDD7EE", "#9DC3E6", "#4472C4"], alpha=0.9)
    ax.set_title("Cumulative Flow Diagram (WIP)")
    ax.set_xlabel("Date")
    ax.set_ylabel("Count")
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(ax.xaxis.get_major_locator()))
    ax.legend(loc="upper left")
    ax.grid(True, axis="y", alpha=0.3)

    os.makedirs(out_dir, exist_ok=True)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "cfd.png"), dpi=150)
    plt.close(fig)


# ---------------------- CLI ----------------------

def run_cli() -> None:
    ap = argparse.ArgumentParser(description="Flow metrics from CSV")
    ap.add_argument("--input", required=False, help="Path to input CSV (Jira export / Raw_Data)")
    ap.add_argument("--out", default="out", help="Output directory for charts and facts")
    ap.add_argument("--facts-csv", default="jira_facts.csv", help="Output facts CSV filename")
    args = ap.parse_args()

    input_path = args.input
    if not input_path:
        # Try sensible defaults so users can run without flags
        for cand in [
            os.path.join("data", "raw_data_example.csv"),
            os.path.join("data", "jira_metrics_sample.csv"),
            os.path.join("data", "wip_example.csv"),
        ]:
            if os.path.exists(cand):
                input_path = cand
                break
    if not input_path or not os.path.exists(input_path):
        raise SystemExit("No --input provided and no default CSV found (data/raw_data_example.csv, data/jira_metrics_sample.csv, data/wip_example.csv). Provide --input path to a CSV.")

    df_raw = pd.read_csv(input_path)
    facts = build_facts(df_raw)

    os.makedirs(args.out, exist_ok=True)
    facts_path = os.path.join(args.out, args.facts_csv)
    # Write a tidy facts CSV
    facts.df.to_csv(facts_path, index=False)

    plot_cycle_scatter(facts, args.out)
    plot_throughput(facts, args.out)
    plot_cfd(facts, args.out)

    print(f"Input: {input_path}")
    print(f"Wrote facts: {facts_path}")
    print(f"Charts: {os.path.join(args.out, 'cycle_time_scatter.png')}, {os.path.join(args.out, 'throughput.png')}, {os.path.join(args.out, 'cfd.png')}")


if __name__ == "__main__":
    run_cli()
