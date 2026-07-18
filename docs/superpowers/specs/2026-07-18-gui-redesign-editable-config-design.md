# GUI Redesign + Editable Config (Standard schedule & JQL) — Design

**Date:** 2026-07-18
**Status:** Approved (design), to be implemented in two phases
**Platform decision:** Stay on WinForms PowerShell (`jira-sync.ps1`). Web design
skills (impeccable/ui-ux-pro-max) are out of scope — they target HTML/React and
cannot apply to WinForms. UX is improved by hand within WinForms.

## Problem

The Settings dialog is a flat stack of fields and the main window's controls
aren't grouped, so the app is "не совсем удобно". Two configs the user needs to
see/tune are locked in code: the standard-task **schedule** (`SCHEDULE` /
`QA_PLACEHOLDERS` in `standard_tasks.py`) and the six **JQL queries** (inline
f-strings in `generate_week_rows`). The user's real JQL is more elaborate than
the code's, so they need to edit queries themselves and copy a ready-to-run JQL
into a browser.

## Goals

1. Externalize the standard schedule and the JQL queries to editable JSON
   configs that the Python modules read (falling back to built-in defaults).
2. Reorganize the **Settings** dialog into a tabbed layout with three tabs:
   Connection, Standard tasks (editable grid), Jira queries (editable text +
   per-query Copy button that copies a fully-substituted JQL).
3. Tidy the **main window** UX: logical button grouping, spacing, clearer
   labels, cleaner status/log area.

## Non-goals (YAGNI)

- No web rewrite; no Electron; no HTML.
- Do not change the sync/query *result-processing* logic in `generate_week_rows`
  (priority bucketing, name selection) — only the JQL *text* becomes editable.
- Do not change `fill-standard` behavior or the Excel output (done separately).
- No reordering/renaming of query slots — slot keys are fixed (they bind to the
  Python post-processing); only each slot's JQL text is user-editable.

## Config files

Both live under `config/` (gitignored as user data — they hold user edits;
Python ships built-in defaults so a fresh clone works without them).

### `config/standard_tasks.json`
```json
{
  "schedule": [
    {"name": "Internal Daily meeting",    "hours": 0.5, "days": ["Mon","Tue","Wed","Thu","Fri"], "freq": "weekly"},
    {"name": "Internal bug triage",       "hours": 0.5, "days": ["Tue"],        "freq": "weekly"},
    {"name": "External customer meeting", "hours": 1.0, "days": ["Tue","Wed"],  "freq": "weekly"},
    {"name": "Weekly project report",     "hours": 1.0, "days": ["Fri"],        "freq": "weekly"},
    {"name": "Internal sprint review",    "hours": 0.5, "days": ["Fri"],        "freq": "sprint-end"},
    {"name": "External sprint review",    "hours": 1.0, "days": ["Fri"],        "freq": "sprint-end"},
    {"name": "Summary report creation",   "hours": 2.0, "days": ["Fri"],        "freq": "sprint-end"}
  ],
  "placeholders": ["Bug verification", "Functional testing", "Automation test maintenance", "Investigation issue"]
}
```
`standard_tasks.py` gains `load_schedule()` that reads this file (from the repo
root's `config/`), validates shape, and falls back to the current built-in
`SCHEDULE`/`QA_PLACEHOLDERS` on missing/invalid. `rows_for_week` consumes the
loaded values. `freq` ∈ `weekly | sprint-end`; placeholders are names-only.

### `config/jql_queries.json`
```json
{
  "queries": [
    {"key": "investigation", "label": "Investigation issues (bugs you created/reported)",
     "jql": "project = {project} AND issuetype = Bug AND created >= \"{ws}\" AND created <= \"{we}\" AND (creator = {account_id} OR reporter = {account_id}) ORDER BY priority DESC, issuetype ASC, key ASC"},
    ... one entry per existing query slot, keys fixed ...
  ]
}
```
`jira-sync.py` gains `load_jql()` returning `{key: template}` with fallback to
the templates currently inlined in `generate_week_rows`. `generate_week_rows`
looks up each slot's template by key, substitutes `{project} {account_id} {ws}
{we}`, runs it, and keeps its existing per-slot result-processing. Keys are the
stable contract between config and code.

## GUI: Settings (tabbed)

`Show-Settings` is rebuilt around a `System.Windows.Forms.TabControl` with three
`TabPage`s:

- **Connection** — existing fields (Jira URL, Email, API Token + show toggle,
  Project, Account ID, Sprint end, Excel file + browse), grouped with consistent
  label column and spacing. Test Connection + Save & Close stay at the dialog
  footer (outside the tabs).
- **Standard tasks** — a `DataGridView` bound to the schedule: columns *Task,
  Mon, Tue, Wed, Thu, Fri, Frequency*. Day columns hold the per-day hours (blank
  = not that day). Frequency is a combo cell: `weekly | sprint-end |
  placeholder` (placeholder rows ignore hours). Add row / Remove row buttons.
  Save writes `config/standard_tasks.json`.
- **Jira queries** — a scrollable panel; one row per query slot: a bold label, a
  multiline read/write `TextBox` (the JQL template), and a **Copy** button on the
  right. Copy substitutes `{project}` and `{account_id}` from the Connection
  fields and `{ws}`/`{we}` from the first checked week in the main window (fall
  back to the current week if none checked), then puts the resolved JQL on the
  clipboard. Save writes `config/jql_queries.json`.

Save & Close persists `.env` (Connection) **and** both config JSONs in one action.

## GUI: main window

Regroup the button strip into labelled clusters with separators — **Excel**
(Read File, Open File, Standard → Excel), **Jira** (Test, Sync Jira → Excel),
**D365** (Submit Tasks → D365, Fill Times → D365), plus Stop. Increase padding,
align the Weeks/Month strip, and give the status line + log clearer separation.
No functional change to what the buttons do.

## Phasing

- **Phase 1 — Editable config + tabbed Settings.** Externalize both configs
  (Python read + fallback + tests), then rebuild the Settings dialog as the
  three-tab layout with the grid, the JQL editors + Copy, and combined Save.
  Delivers requirements 3 and 4 and the Settings half of 5.
- **Phase 2 — Main-window UX.** Restructure the main window layout only
  (grouping, spacing, labels, status/log). Delivers the rest of 5.

## Error handling

- Missing/invalid config JSON → log a warning, use built-in defaults; never crash.
- Copy with no week selected → use the current week and note it in the status line.
- Saving configs while a file is locked/unwritable → message box, no partial write.
- DataGridView rows with an empty task name or no day hours (non-placeholder) →
  skipped on save with a warning.

## Testing

- Python: `load_schedule()` / `load_jql()` — valid file parsed, missing file →
  defaults, malformed file → defaults + warning. `rows_for_week` still passes
  with a config-provided schedule. `generate_week_rows` substitutes a
  config-provided template (mock the HTTP call). Plain-script tests, no pytest.
- GUI: PowerShell parse check + manual eyeball (tabs render, grid edits persist,
  Copy places resolved JQL on the clipboard). WinForms can't be asserted headless.
