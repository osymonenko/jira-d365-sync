"""Pure date-window helpers for month-scoped Jira sync. No I/O, no network."""

from datetime import date, timedelta


def month_bounds(month: str) -> tuple[date, date]:
    """First and last calendar day of a YYYY-MM month. Raises ValueError otherwise."""
    try:
        year_s, mon_s = month.split("-")
        year, mon = int(year_s), int(mon_s)
        first = date(year, mon, 1)
    except (ValueError, AttributeError):
        raise ValueError(f"--month must be YYYY-MM, got: {month!r}")
    next_month = date(year + 1, 1, 1) if mon == 12 else date(year, mon + 1, 1)
    return first, next_month - timedelta(days=1)


def clamp_week_to_month(week_start: date, week_end: date, month: str):
    """Intersect [week_start, week_end] with the month.

    Returns the clamped (start, end) tuple, or None if the week does not
    overlap the month at all.
    """
    first, last = month_bounds(month)
    if week_end < first or week_start > last:
        return None
    return (max(week_start, first), min(week_end, last))
