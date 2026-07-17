# Standard Tasks Generator — Design

**Date:** 2026-07-17
**Status:** Approved (design), pending implementation plan

## Problem

`Sync Jira → Excel` fills the timesheet from **real, already-happened** Jira
activity (bugs created, issues moved to Done by the user, etc.). For **future**
weeks that activity does not exist yet, so the sync returns nothing (verified:
running it for 2026-07-19…08-01 produced `0 found` on every query).

The user needs a way to pre-fill a timesheet with the recurring, predictable
work that happens every week/sprint, plus placeholder rows for the QA activities
that will only be quantifiable later.

## Goal

Given a **month** selected in the GUI and a set of **checked weeks**, insert the
standard recurring tasks (with hours) and — for weeks that still have future
work days — QA placeholder rows (names only) into the existing week blocks of
`timesheet.xlsx`. Everything is driven from the GUI; no CLI use by the user.

## Non-goals (YAGNI)

- Not creating week blocks that don't exist — only filling existing ones.
- No JSON/editable-config schedule (a Python dict is enough; it changes rarely).
- No new date-picker widgets — reuse the existing Month dropdown + week checkboxes.
- Not touching the D365 automation or the Jira sync queries.

## Schedule

Days map to Excel columns: **Mon=D(4), Tue=E(5), Wed=F(6), Thu=G(7), Fri=H(8)**.
The first hours column is Monday. Saturday/Sunday are never work days.

### Base tasks (always inserted, with hours)

| Task | Frequency | Days | Hours/day |
|------|-----------|------|-----------|
| Internal Daily meeting | weekly | Mon–Fri | 0.5 |
| Internal bug triage | weekly | Tue | 0.5 |
| External customer meeting | weekly | Tue, Wed | 1.0 |
| Weekly project report | weekly | Fri | 1.0 |
| Internal sprint review | sprint-end (biweekly) | Fri | 0.5 |
| Summary report creation | sprint-end (biweekly) | Fri | 2.0 |

Task-name casing is irrelevant to D365 (its lookup is case-insensitive); names
are written to Excel exactly as in the schedule.

### QA placeholder tasks (names only, no hours)

`Bug verification`, `Functional testing`, `Automation test maintenance`,
`Investigation issue`.

Added to a week **only if that week has a future work day** (see rules). They
carry no hours and no day mapping, so D365 will not submit them — they exist in
Excel for the user to fill in later. For past/current-done weeks these instead
come from the real Jira sync.

## Rules

- **Weeks processed:** existing week blocks (from `find_weeks`) that intersect
  the selected range and are checked in the GUI. Missing blocks → `[SKIP]`.
- **Month clamping:** the "range" is the **month selected in the dropdown**. For
  each week, `clamp_week_to_month(week_start, week_end, month)` restricts the
  window; base-task hours are written **only on Mon–Fri days that fall inside
  the selected month**. (Reuses the existing `month_filter` helper.) When
  `Month = All` (no month chosen) there is no clamping — whole weeks are filled.
- **Sprint-end weeks:** `SPRINT_ANCHOR` (settings) is any known sprint-end
  Friday. A week is a sprint-end week iff its Friday `f` satisfies
  `(f − anchor).days % 14 == 0`. The week's Friday = `week_start` (Sunday) + 5.
  Sprint-end is judged on the **real** Friday, not the clamped window.
- **QA trigger:** add QA placeholders iff `week_friday > today` (Friday is the
  last work day, so this is exactly "has a future work day"). Judged on the real
  Friday. `today` is injected for testability.
- **Dedup:** skip a task whose name already exists in the week block,
  **case-insensitively** (so "Internal daily meeting" ≠ a second insert of
  "Internal Daily meeting").
- **Hours:** written as numbers (0.5 / 1.0 / 2.0) into the day columns; all are
  valid D365 durations (15-min grid, ≤ 8h).

### Worked example (today = Fri 2026-07-17)

- Week 2026-07-13…07-19: last work day Fri 07-17 = today, not after → **no QA**.
- Week 2026-07-20…07-24: entirely future → base tasks + **QA placeholders**.
- Week 2026-06-29…07-05 with July selected: "Internal Daily meeting" hours only
  on Wed 07-01 / Thu 07-02 / Fri 07-03 (Mon 06-29, Tue 06-30 clamped out).

## Architecture (Approach A)

Separation: `standard_tasks.py` knows **what/when**; `jira-sync.py` knows **how
to write Excel**.

### `scripts/standard_tasks.py` (new — pure logic, no openpyxl/network)

- `SCHEDULE` — list of dicts `{name, hours, days, freq}` (`freq` ∈
  `weekly | sprint-end`).
- `QA_PLACEHOLDERS` — list of names.
- `rows_for_week(week_start, week_end, month, sprint_anchor, today) -> list[TaskRow]`
  where `TaskRow = {name, hours_by_day}` and `hours_by_day` maps a day column
  index → hours (empty dict for QA placeholders). Encapsulates frequency,
  sprint-end, month-clamp, and future-day logic. Fully unit-testable.

### `scripts/jira-sync.py` (extended)

- Refactor row insertion so both Jira rows (name + hyperlink) and standard rows
  (name + per-day hours) share one insert mechanism + dedup. The shared helper
  keeps `_insert_rows_preserving_hyperlinks` and the check-sum-anchor logic.
- Standard-row writing additionally sets day-column values (4–8) and uses a
  plain font (no hyperlink).
- New command `--command fill-standard --weeks <mondays> --month <YYYY-MM>
  --file <path>`; `SPRINT_ANCHOR` read from `.env`. Warn + skip sprint-end tasks
  if the anchor is missing/invalid; if the anchor is not a Friday, snap to that
  week's Friday and warn.

### GUI (`jira-sync.ps1`)

- New button **"Fill Standard → Excel"** next to "Sync Jira → Excel". Uses the
  already-selected month + checked weeks; calls the new command. No new date
  inputs.
- Settings: add **"Sprint end (last Fri):"** field (hint: "any sprint-end
  Friday — 2-week cycles counted from here"), persisted to `.env` as
  `SPRINT_ANCHOR`.

## Error handling

Soft-fail; never abort the whole run:

- `SPRINT_ANCHOR` empty/invalid → `[WARN]`, skip the two sprint-end tasks; other
  tasks proceed.
- Checked week with no block in Excel → `[SKIP] <week>`.
- All rows already present → `[SKIP] <week>: all rows already present`.
- Excel open/locked → existing `PermissionError` handling (ask user to close).
- No weeks in the selected month → informational message, no changes.

## Testing

`scripts/test_standard_tasks.py`, unit tests over `rows_for_week()` with `today`
and `sprint_anchor` passed in (deterministic — no system clock):

- Weekly tasks present in every processed week.
- Sprint-end tasks only on `anchor ± 14·n` weeks; absent in between.
- QA placeholders appear when `week_friday > today`, absent when `≤ today`.
- Month clamp drops out-of-month work days on boundary weeks.
- Per-day hours match the schedule (correct columns, correct values).

Pattern follows the existing `scripts/test_month_filter.py`.
