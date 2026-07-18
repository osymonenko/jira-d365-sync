# Settings Dialog Follow-up — Connection Reorg, Sprint Length, JQL Row Unification — Design

**Date:** 2026-07-18
**Status:** Approved (design)
**Builds on:**
- `docs/superpowers/specs/2026-07-18-gui-redesign-editable-config-design.md` (Phase 1)
- `docs/superpowers/specs/2026-07-18-gui-redesign-custom-jql-queries-design.md` (custom JQL slots)

## Problem

After the custom-JQL-slots feature shipped, live use of the Settings dialog
surfaced four issues:

1. The Connection tab's field order doesn't match how the user actually
   fills it in, "Project" is ambiguous (vs. account/board), Sprint end
   doesn't belong next to connection credentials, and Test Connection reads
   as a dialog-wide action even though it only tests the Connection tab's
   fields.
2. The sprint-end cadence is hardcoded to 14 days in Python
   (`is_sprint_end_week`) — there's no way to tell the tool a different
   sprint length without editing code.
3. The six built-in JQL query rows and user-added custom rows are visibly
   different (different control sizes/fonts) because they're built by two
   separate code paths, and the six built-in rows can't be removed at all.
4. "+ Add query" sits above the row list, and a newly added row isn't
   scrolled into view or focused, so on a long list it's easy to miss.

## Goals

1. Reorder and relabel the Connection tab; move Sprint end off it; scope
   Test Connection (button + status label) to the Connection tab instead of
   the dialog-wide footer.
2. Add a **Sprint length (weeks)** field (Standard tasks tab, alongside the
   relocated Sprint end) that actually changes the sprint-end cycle Python
   computes — not just a label.
3. Unify the six built-in and user-added JQL rows into one visual/behavioral
   shape, and let the built-in six be deleted too (GUI-only, see Non-goals).
4. Move "+ Add query" below the row list; scroll and focus a newly added
   row into view.

## Non-goals (YAGNI)

- Deleting a built-in JQL row does **not** disable the corresponding
  automated query in `Sync Jira -> Excel`. `generate_week_rows` always calls
  all six `build(key)` sites regardless of GUI state; deleting a built-in
  row only removes its custom JQL text override (config falls back to
  `DEFAULT_JQL` for that key on next load, same as if it were never
  customized) and its row from the Settings UI. This was explicitly
  confirmed as the accepted behavior, not something this plan works around.
- No "restore a deleted built-in row" button. If a user wants a hidden
  built-in row back, they'd need to edit `config/jql_queries.json` by hand
  (remove its key from `hidden_builtin`) — out of scope for this pass.
- No validation UI for Sprint length beyond a warning + default fallback
  (matches every other config value in this codebase — missing/invalid →
  `[WARN]`, never crash).
- No change to `generate_week_rows`, `search_jira`, or any Excel-writing
  logic — this plan only touches the Settings dialog and the standard-task
  sprint-cycle calculation.

## Connection tab: field order and scope

New top-to-bottom order: **Jira URL → API Token (+ Show token checkbox
directly under it) → Project key → Email → Account ID (+ hint) → Excel file
(+ browse) → Test Connection (+ status label)**.

- "Project" label text becomes "Project key:" (cosmetic only — the
  underlying `.env` key `JIRA_PROJECT` and the PowerShell variable `$tProj`
  are unchanged).
- Sprint end (the `$tAnchor` field + its hint) is removed from this tab —
  it moves to Standard tasks (below).
- Test Connection's button and status label move from `$dlg`'s global
  footer into `$tabConn` itself, using tab-local coordinates. They keep
  their exact existing behavior (`Save-JiraEnv` call, Python `--command
  test` subprocess, connection-status-dot update) — only their parent
  container and position change.
- Save & Close remains the dialog's one global, footer-level button (still
  persists Connection fields, Standard-tasks grid + sprint fields, and JQL
  config in one action).

## Standard tasks tab: Sprint length + Sprint end

Two new fields at the top of the tab, above the existing grid (which shifts
down to make room):

- **Sprint length (weeks):** a plain text field (small width), default
  display `'2'` if unset — matches today's hardcoded behavior for a user
  who never touches it.
- **Sprint end:** the field relocated from Connection, same hint text as
  before ("Any sprint-end Friday (YYYY-MM-DD) - N-week cycles counted from
  here" — the hint's cycle count reflects the current Sprint-length value
  so it stays accurate after a user changes it away from 2).

Both persist through `Save-JiraEnv`, which gains a `$sprintLengthWeeks`
parameter and writes a new `.env` key `SPRINT_LENGTH_WEEKS` alongside the
existing `SPRINT_ANCHOR`. Both call sites of `Save-JiraEnv` (Test Connection
and Save & Close) pass the current values of these fields — they're plain
PowerShell variables in the same `Show-Settings` scope regardless of which
tab their controls visually live on, so no cross-tab plumbing is needed
beyond referencing the right variable names.

## Sprint length: real effect on Python

`scripts/standard_tasks.py`:
- `is_sprint_end_week(week_start, sprint_anchor, cycle_days=14)` gains an
  optional `cycle_days` parameter (default `14` preserves every existing
  2-arg call site and test).
- `rows_for_week(..., schedule=None, placeholders=None, cycle_days=14)`
  gains the same optional parameter, passed straight through to
  `is_sprint_end_week`.

`scripts/jira-sync.py::cmd_fill_standard`:
- Reads `SPRINT_LENGTH_WEEKS` from `.env` (same `load_env` already used for
  `SPRINT_ANCHOR`). Missing → default `2`. Present but not a positive
  integer → `[WARN]` + default `2` (never crash, matches every other
  config-reading path in this codebase).
- Computes `cycle_days = sprint_length_weeks * 7`, passes it to
  `rows_for_week(..., cycle_days=cycle_days)`.

## Jira queries tab: unify fixed and custom rows

Replace `New-JqlFixedRow`/`New-JqlCustomRow` (two separate functions
building visibly different row shapes) with **one** function,
`New-JqlRow($key, $titleText, $jqlText, $titleEditable)`:

- Same row `Panel` size, same JQL `TextBox` size/font (Consolas 8) for every
  row — this is what fixes the reported visual mismatch.
- Title is a read-only `Label` when `$titleEditable` is `$false` (the six
  built-in slots — their key/label identity stays fixed, only their JQL
  text and existence are user-controlled) or an editable `TextBox` when
  `$true` (custom slots, as already shipped).
- Every row — built-in or custom — gets both **Copy** (unchanged
  substitution logic) and **Delete** (unchanged
  `$this.Parent.Parent.Controls.Remove($this.Parent)` pattern, already
  proven safe by the custom-slots feature's empirical WinForms tests).
  Deleting is immediate, no confirmation, for every row — consistent with
  the existing custom-row behavior and the standard-tasks grid's Delete
  column.

**Persistence:** `config/jql_queries.json` gains a new top-level array,
`hidden_builtin`: the list of built-in keys whose row was deleted and not
re-added. On dialog open, a built-in slot is skipped (not rendered) if its
key is in `hidden_builtin`. On Save, `hidden_builtin` is recomputed as "the
six fixed keys minus whichever fixed-tagged rows are still present in
`$panelJqlFlow.Controls`". This is a GUI-only bookkeeping field —
`jira-sync.py::load_jql()` is not modified and does not read it; it
continues to ignore any JSON key it doesn't recognize (proven by the
existing test in `scripts/test_jql_config.py`).

**Save logic simplifies to one loop:** since every row (built-in or custom)
now has the same control layout (`Controls[0]` = title, `Controls[1]` =
JQL, `.Tag` = key), `$script:SaveJqlConfig` iterates
`$panelJqlFlow.Controls` once, reading `Controls[0].Text`/`Controls[1].Text`
uniformly. The per-slot `$script:jqlBoxes` dictionary and the two-loop
structure (fixed via `$jqlSlots`, custom via `$panelJqlFlow.Controls`) are
removed — there is no longer a distinction in *how* a row's current text is
read, only in whether it came from `$jqlSlots` (built-in key set, for
computing `hidden_builtin`) or not.

**"+ Add query" moves below the row list:** the button strip's `Dock`
changes from `'Top'` to `'Bottom'`; it must still be added to
`$tabJql.Controls` **after** `$panelJqlFlow` (same "most-recently-added
docked control wins its edge" rule as before, just the opposite edge).

**Focus/scroll on add:** `New-JqlRow` returns the row `Panel` it built. The
"+ Add query" click handler captures that return value and calls
`$panelJqlFlow.ScrollControlIntoView($newRow)` then focuses its title
control (`$newRow.Controls[0].Focus()`) — both run synchronously within the
same click-handler invocation that created the row, so there's no
scope-lifetime risk (unlike the earlier closure bug, which involved reusing
a *different* function's locals *after* it had already returned).

## Error handling

- Sprint length missing/invalid → `[WARN]` + default `2` weeks (14 days),
  never crash — mirrors every other config-reading path in this codebase.
- `hidden_builtin` missing/invalid in `config/jql_queries.json` → treated as
  empty (show all six built-in rows) — mirrors `schedule`/`placeholders`
  handling in the standard-tasks config.
- Deleting a built-in row and Saving does not stop `Sync Jira -> Excel`
  from running that query with default text (see Non-goals) — this is
  intentional, not an error condition.

## Testing

- **Python:** extend `scripts/test_standard_tasks.py` (or a focused new
  test) with a case that passes a non-default `cycle_days` (e.g. `21` for a
  3-week sprint) and asserts a week that would NOT be sprint-end at 14 days
  IS flagged sprint-end at 21, and vice versa — this is genuinely new
  behavior, not a lock-in test, so it follows normal TDD (write failing
  test, implement, confirm passing).
- **GUI:** PowerShell parse check + manual eyeball (this project's
  standing pattern for WinForms layout) — reordered/renamed Connection
  fields, Sprint length/Sprint end on Standard tasks, Test Connection
  scoped to Connection tab and still functional, all six + any custom JQL
  rows rendering identically in shape/font, Delete works on a built-in row
  and persists across reopen (row skipped on reload, `hidden_builtin`
  written), "+ Add query" below the list, new row scrolls into view and
  receives focus.
- Given the prior custom-JQL-slots bug was only caught by *empirical*
  headless `PerformClick()` testing (not diff-reading), the implementation
  plan requires a headless verification script for `New-JqlRow`'s Copy and
  Delete handlers — for both `$titleEditable = $false` (built-in-shaped) and
  `$true` (custom-shaped) rows — before any task touching this function is
  considered done, not just a parser-check pass.
