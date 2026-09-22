"""Автономные тесты balance_hours (без pytest). Запускать из scripts/."""

from balance_hours import (
    allocate, spread, split_counts, plan_week, hours_to_q, q_to_hours,
    is_spread_task, is_automation_task, task_weight,
    AUTOMATION_WEIGHT, DEFAULT_WEIGHT,
)

MON, TUE, WED, THU, FRI = 4, 5, 6, 7, 8
ALL_DAYS = [MON, TUE, WED, THU, FRI]

failures = []


def check(name, cond, detail=""):
    if cond:
        print(f"  PASS  {name}")
    else:
        print(f"  FAIL  {name}  {detail}")
        failures.append(name)


def totals_by_day(plan):
    out = {}
    for row in plan:
        for col, h in row.items():
            out[col] = round(out.get(col, 0) + h, 2)
    return out


def names(n, prefix="Task"):
    return [f"{prefix} {i}" for i in range(n)]


def ones(n):
    return [1.0] * n


# ── конверсия ────────────────────────────────────────────────────────────────
check("hours_to_q(0.5)==2", hours_to_q(0.5) == 2)
check("hours_to_q float noise", hours_to_q(0.4999999999) == 2)
check("q_to_hours(3)==0.75", q_to_hours(3) == 0.75)

# ── классификация задач ──────────────────────────────────────────────────────
check("spread: Bug verification", is_spread_task("Bug verification P1-2, P2-4"))
check("spread: Investigation issue", is_spread_task("Investigation issue 4"))
check("spread: case-insensitive", is_spread_task("BUG VERIFICATION"))
check("spread: not functional testing", not is_spread_task("Functional testing of story ID GT2-1662"))
check("spread: not empty", not is_spread_task(""))

check("automation: maintenance", is_automation_task("Automation test maintenance"))
check("automation: mid-string", is_automation_task("Setup of automation environment"))
check("automation: not functional testing", not is_automation_task("Functional testing of story ID GT2-1662"))
check("weight: automation is 2x", task_weight("Automation test update") == AUTOMATION_WEIGHT)
check("weight: default is 1", task_weight("Functional testing") == DEFAULT_WEIGHT)

# ── allocate ─────────────────────────────────────────────────────────────────
a = allocate(96, ones(18))
check("allocate sums exactly", sum(a) == 96, f"got {sum(a)}")
check("allocate is even (spread <= 0.5h)", max(a) - min(a) <= 2, f"{a}")
check("allocate min chunk honoured", min(a) >= 2, f"{a}")
check("allocate multiples of 0.5h", all(x % 2 == 0 for x in a), f"{a}")

a = allocate(121, ones(7))  # 30.25 ч — хвост в одну четверть не делится пополам
check("allocate odd tail sums exactly", sum(a) == 121, f"{a}")
check("allocate odd tail: only one odd value", sum(1 for x in a if x % 2) == 1, f"{a}")

a = allocate(122, ones(7))
check("allocate even total stays on 0.5 grid", all(x % 2 == 0 for x in a), f"{a}")

a = allocate(6, ones(10))  # 1.5 ч на 10 задач — на всех не хватит
check("allocate scarce sums exactly", sum(a) == 6, f"{a}")
check("allocate scarce leaves zeros", a.count(0) == 7, f"{a}")
check("allocate scarce keeps min 0.5", all(x == 0 or x >= 2 for x in a), f"{a}")

check("allocate no tasks", allocate(10, []) == [])
check("allocate no capacity", allocate(0, ones(3)) == [0, 0, 0])

# веса: автоматизация тяжелее
a = allocate(32, [1.0, 1.0, 2.0])  # 8 ч на три задачи
check("weighted sums exactly", sum(a) == 32, f"{a}")
check("weighted: automation is the biggest", a[2] > a[0] and a[2] > a[1], f"{a}")
check("weighted: equal peers stay equal", a[0] == a[1], f"{a}")

a = allocate(64, [1.0, 2.0])  # 16 ч на две задачи
# По 0.5 ч минимума каждой, остальные 15 ч делятся 1:2 -> 5.5 ч и 10.5 ч
check("weighted 1:2 surplus split", a == [22, 42], f"{a}")

# ── spread ───────────────────────────────────────────────────────────────────
days = [[c, 8] for c in ALL_DAYS]
out = spread(8, days)  # 2 ч
check("spread covers 4 days by 0.5h", out == {MON: 2, TUE: 2, WED: 2, THU: 2}, f"{out}")

days = [[c, 8] for c in ALL_DAYS]
out = spread(20, days)  # 5 ч на 5 дней -> по 1 ч
check("spread second lap evens out", out == {c: 4 for c in ALL_DAYS}, f"{out}")

days = [[MON, 2], [TUE, 8]]
out = spread(8, days)
check("spread respects day capacity", out == {MON: 2, TUE: 6}, f"{out}")
check("spread debits capacity", days == [[MON, 0], [TUE, 2]], f"{days}")

days = [[MON, 2]]
check("spread stops when week is full", spread(10, days) == {MON: 2})

# ── split_counts ─────────────────────────────────────────────────────────────
c = split_counts(10, [8, 8, 8, 8, 8])
check("split_counts sums to n", sum(c) == 10, f"{c}")
check("split_counts is even", max(c) - min(c) <= 1, f"{c}")

c = split_counts(3, [8, 8, 8, 8, 8])
check("split_counts seeds every day first", sorted(c, reverse=True)[:3] == [1, 1, 1], f"{c}")

c = split_counts(10, [2, 16])  # в дне на 0.5ч помещается 1 задача
check("split_counts respects day capacity", c[0] <= 1, f"{c}")

c = split_counts(5, [0, 8])
check("split_counts skips empty days", c[0] == 0, f"{c}")

# ── plan_week ────────────────────────────────────────────────────────────────
plan, free = plan_week({}, ALL_DAYS, names(10))
check("plan_week free = 160q", free == 160, f"{free}")
total = sum(sum(r.values()) for r in plan)
check("plan_week distributes all 40h", abs(total - 40) < 1e-9, f"{total}")
per_day = totals_by_day(plan)
check("plan_week every day hits 8", all(abs(v - 8) < 1e-9 for v in per_day.values()), f"{per_day}")
check("plan_week covers 5 days", sorted(per_day) == ALL_DAYS, f"{per_day}")
vals = [h for r in plan for h in r.values()]
check("plan_week: no 0.25/0.75 values", all(abs(v * 2 - round(v * 2)) < 1e-9 for v in vals), f"{vals}")

# ГЛАВНОЕ: обычная задача целиком в одном дне
check("plan_week: no carry-over to next day", all(len(r) <= 1 for r in plan),
      f"{[r for r in plan if len(r) > 1]}")

# Регулярка уже заняла часть дня
used = {MON: 0.5, TUE: 3.0, WED: 2.5, THU: 0.5, FRI: 5.0}
plan, free = plan_week(used, ALL_DAYS, names(6))
per_day = totals_by_day(plan)
totals = {c: used.get(c, 0) + per_day.get(c, 0) for c in ALL_DAYS}
check("plan_week tops up to 8 with prefilled", all(abs(v - 8) < 1e-9 for v in totals.values()), f"{totals}")
check("plan_week prefilled: no carry-over", all(len(r) <= 1 for r in plan), f"{plan}")
vals = [h for r in plan for h in r.values()]
check("plan_week prefilled: half-hour grid", all(abs(v * 2 - round(v * 2)) < 1e-9 for v in vals), f"{vals}")

# Размазываемые задачи — единственное исключение из «один день на задачу»
task_names = ["Bug verification P1-2", "Investigation issue 3"] + names(8)
plan, _ = plan_week({}, ALL_DAYS, task_names)
check("plan_week spreads bug verification", len(plan[0]) >= 4, f"{plan[0]}")
check("plan_week spreads investigation", len(plan[1]) >= 4, f"{plan[1]}")
check("plan_week spread chunks on 0.5 grid",
      all(abs(v * 2 - round(v * 2)) < 1e-9 for v in plan[0].values()), f"{plan[0]}")
check("plan_week spread chunks stay small", max(plan[0].values()) <= 1.0, f"{plan[0]}")
check("plan_week: blocks still single-day", all(len(r) <= 1 for r in plan[2:]), f"{plan[2:]}")
per_day = totals_by_day(plan)
check("plan_week still hits 8 with spread", all(abs(v - 8) < 1e-9 for v in per_day.values()), f"{per_day}")

# Регрессия: блочных задач меньше, чем дней. Раньше последний день оставался
# недозаполненным (8/8/8/8/3) — теперь хвост подбирают размазываемые задачи.
task_names = ["Bug verification", "Investigation issue", "Investigation issue 2",
              "Functional testing", "Automation test maintenance",
              "Other project documentation work", "[REVIEW] Update report script"]
plan, _ = plan_week({MON: 0.5, TUE: 2.0, WED: 1.5, THU: 0.5, FRI: 1.5}, ALL_DAYS, task_names)
per_day = totals_by_day(plan)
used = {MON: 0.5, TUE: 2.0, WED: 1.5, THU: 0.5, FRI: 1.5}
totals = {c: round(used[c] + per_day.get(c, 0), 2) for c in ALL_DAYS}
check("plan_week: few block tasks still fill every day",
      all(abs(v - 8) < 1e-9 for v in totals.values()), f"{totals}")
check("plan_week: few block tasks, no carry-over",
      all(len(plan[i]) <= 1 for i in range(3, len(task_names))), f"{plan[3:]}")

# Автоматизация тяжелее соседей в том же дне
task_names = ["Functional testing A", "Automation test maintenance", "Functional testing B"]
plan, _ = plan_week({}, [MON], task_names)
hours = [sum(r.values()) for r in plan]
check("plan_week: automation outweighs peers", hours[1] > hours[0] and hours[1] > hours[2], f"{hours}")
check("plan_week: automation day still 8", abs(sum(hours) - 8) < 1e-9, f"{hours}")
check("plan_week: automation within 1..8", 1 < hours[1] <= 8, f"{hours}")

# Обрезка по месяцу: доступны только Вт..Чт
plan, free = plan_week({}, [TUE, WED, THU], names(5))
touched = {c for r in plan for c in r}
check("plan_week respects active days", touched == {TUE, WED, THU}, f"{touched}")
check("plan_week free for 3 days", free == 96, f"{free}")

plan, _ = plan_week({}, [TUE, WED], ["Bug verification", "Other"])
check("spread stays inside active days", set(plan[0]) <= {TUE, WED}, f"{plan[0]}")

# День уже переполнен — в минус не уходим
plan, free = plan_week({MON: 9.0}, [MON], names(3))
check("plan_week overfull day gives nothing", free == 0 and all(r == {} for r in plan), f"{plan} {free}")

# Нет задач под раскладку
plan, free = plan_week({}, ALL_DAYS, [])
check("plan_week no tasks", plan == [] and free == 160, f"{plan} {free}")

print()
if failures:
    print(f"{len(failures)} FAILURE(S): {', '.join(failures)}")
    raise SystemExit(1)
print("ALL PASS")
