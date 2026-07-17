from datetime import date
from month_filter import clamp_week_to_month, month_bounds


def check(name, got, exp):
    assert got == exp, f"{name}: got {got}, expected {exp}"
    print(f"  OK {name}")


# month_bounds
check("bounds July", month_bounds("2026-07"), (date(2026, 7, 1), date(2026, 7, 31)))
check("bounds Feb leap", month_bounds("2024-02"), (date(2024, 2, 1), date(2024, 2, 29)))

# whole week inside the month -> unchanged
check("inside", clamp_week_to_month(date(2026, 7, 6), date(2026, 7, 12), "2026-07"),
      (date(2026, 7, 6), date(2026, 7, 12)))

# boundary week at start of month (Sun 29 Jun - 5 Jul) -> 1-5 Jul
check("start boundary", clamp_week_to_month(date(2026, 6, 29), date(2026, 7, 5), "2026-07"),
      (date(2026, 7, 1), date(2026, 7, 5)))

# boundary week at end of month (27 Jul - 2 Aug) -> 27-31 Jul
check("end boundary", clamp_week_to_month(date(2026, 7, 27), date(2026, 8, 2), "2026-07"),
      (date(2026, 7, 27), date(2026, 7, 31)))

# week fully outside the month -> None
check("outside", clamp_week_to_month(date(2026, 6, 1), date(2026, 6, 7), "2026-07"), None)

# invalid month string -> ValueError
try:
    month_bounds("July")
    raise SystemExit("FAIL: expected ValueError for 'July'")
except ValueError:
    print("  OK invalid month raises")

print("ALL PASS")
