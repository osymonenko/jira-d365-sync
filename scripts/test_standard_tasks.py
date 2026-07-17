from datetime import date
from standard_tasks import rows_for_week, is_sprint_end_week


def check(name, got, exp):
    assert got == exp, f"{name}: got {got}, expected {exp}"
    print(f"  OK {name}")


def by_name(rows):
    return {r["name"]: r["hours_by_col"] for r in rows}


ANCHOR = date(2026, 7, 17)           # a Friday, sprint end
PAST = date(2026, 7, 10)             # QA trigger: friday(2026-07-17) > 2026-07-10 -> True

# --- weekly tasks present every week (week 2026-07-12 Sun .. 2026-07-18) ---
rows = by_name(rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, ANCHOR, PAST))
check("daily cols", rows["Internal Daily meeting"], {4: 0.5, 5: 0.5, 6: 0.5, 7: 0.5, 8: 0.5})
check("bug triage Tue", rows["Internal bug triage"], {5: 0.5})
check("external Tue+Wed", rows["External customer meeting"], {5: 1.0, 6: 1.0})
check("weekly report Fri", rows["Weekly project report"], {8: 1.0})

# --- sprint-end tasks: present on anchor week ---
check("sprint review present", "Internal sprint review" in rows, True)
check("summary present", "Summary report creation" in rows, True)
check("summary hours", rows["Summary report creation"], {8: 2.0})

# --- sprint-end absent on the off-sprint week (2026-07-19, friday 2026-07-24) ---
off = by_name(rows_for_week(date(2026, 7, 19), date(2026, 7, 25), None, ANCHOR, PAST))
check("sprint review absent off-week", "Internal sprint review" in off, False)
check("summary absent off-week", "Summary report creation" in off, False)

# --- sprint-end present again two weeks later (2026-07-26, friday 2026-07-31) ---
nxt = by_name(rows_for_week(date(2026, 7, 26), date(2026, 8, 1), None, ANCHOR, PAST))
check("sprint review +14", "Internal sprint review" in nxt, True)

# --- no anchor -> sprint-end tasks skipped, weekly ones stay ---
noanchor = by_name(rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, None, PAST))
check("no-anchor skips sprint", "Internal sprint review" in noanchor, False)
check("no-anchor keeps daily", "Internal Daily meeting" in noanchor, True)

# --- QA placeholders appear when week friday > today ---
qa = rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, ANCHOR, PAST)
qa_names = {r["name"] for r in qa if not r["hours_by_col"]}
check("QA present when future", qa_names,
      {"Bug verification", "Functional testing", "Automation test maintenance", "Investigation issue"})

# --- QA absent when friday <= today (today = the friday itself) ---
qa_none = rows_for_week(date(2026, 7, 12), date(2026, 7, 18), None, ANCHOR, date(2026, 7, 17))
check("QA absent when not future", any(not r["hours_by_col"] for r in qa_none), False)

# --- month clamp drops out-of-month work days (week 2026-06-28 .. 07-04, July) ---
clamped = by_name(rows_for_week(date(2026, 6, 28), date(2026, 7, 4), "2026-07", ANCHOR, PAST))
check("daily clamped to July", clamped["Internal Daily meeting"], {6: 0.5, 7: 0.5, 8: 0.5})

# --- week fully outside month -> no rows ---
outside = rows_for_week(date(2026, 6, 1), date(2026, 6, 7), "2026-07", ANCHOR, PAST)
check("outside month empty", outside, [])

# --- is_sprint_end_week direct ---
check("anchor is sprint end", is_sprint_end_week(date(2026, 7, 12), ANCHOR), True)
check("off week not sprint end", is_sprint_end_week(date(2026, 7, 19), ANCHOR), False)

print("ALL PASS")
