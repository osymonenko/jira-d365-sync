"""Recurring standard-task schedule and per-week row generation.

Pure logic: no openpyxl, no network, no system clock (today is passed in).
"""

from datetime import date, timedelta

from month_filter import clamp_week_to_month

# Excel day columns (Mon=D .. Fri=H). The first hours column is Monday.
DAY_COL = {"Mon": 4, "Tue": 5, "Wed": 6, "Thu": 7, "Fri": 8}
# Offset in days from the Sunday week_start stored in Excel.
DAY_OFFSET = {"Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5}

SCHEDULE = [
    {"name": "Internal Daily meeting",    "hours": 0.5, "days": ["Mon", "Tue", "Wed", "Thu", "Fri"], "freq": "weekly"},
    {"name": "Internal bug triage",       "hours": 0.5, "days": ["Tue"],          "freq": "weekly"},
    {"name": "External customer meeting", "hours": 1.0, "days": ["Tue", "Wed"],   "freq": "weekly"},
    {"name": "Weekly project report",     "hours": 1.0, "days": ["Fri"],          "freq": "weekly"},
    {"name": "Internal sprint review",    "hours": 0.5, "days": ["Fri"],          "freq": "sprint-end"},
    {"name": "Summary report creation",   "hours": 2.0, "days": ["Fri"],          "freq": "sprint-end"},
]

QA_PLACEHOLDERS = [
    "Bug verification",
    "Functional testing",
    "Automation test maintenance",
    "Investigation issue",
]


def _friday(week_start: date) -> date:
    return week_start + timedelta(days=DAY_OFFSET["Fri"])


def is_sprint_end_week(week_start: date, sprint_anchor: date) -> bool:
    """A week is a sprint end iff its Friday is a whole number of 2-week cycles
    away from the anchor Friday."""
    return (_friday(week_start) - sprint_anchor).days % 14 == 0


def rows_for_week(week_start, week_end, month, sprint_anchor, today):
    """Standard task rows for one week block.

    Returns list of {"name": str, "hours_by_col": {col: hours}}.
    QA placeholders have an empty hours_by_col. Returns [] if the week is
    entirely outside the selected month.
    """
    if month:
        clamped = clamp_week_to_month(week_start, week_end, month)
        if clamped is None:
            return []
        win_start, win_end = clamped
    else:
        win_start, win_end = week_start, week_end

    rows = []
    for task in SCHEDULE:
        if task["freq"] == "sprint-end":
            if sprint_anchor is None or not is_sprint_end_week(week_start, sprint_anchor):
                continue
        hours_by_col = {}
        for day in task["days"]:
            d = week_start + timedelta(days=DAY_OFFSET[day])
            if win_start <= d <= win_end:
                hours_by_col[DAY_COL[day]] = task["hours"]
        if hours_by_col:
            rows.append({"name": task["name"], "hours_by_col": hours_by_col})

    # QA placeholders only when the week still has a future work day.
    if _friday(week_start) > today:
        for name in QA_PLACEHOLDERS:
            rows.append({"name": name, "hours_by_col": {}})

    return rows
