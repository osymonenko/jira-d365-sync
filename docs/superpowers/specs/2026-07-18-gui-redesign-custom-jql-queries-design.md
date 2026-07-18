# JQL Queries Tab — Custom Query Slots (Add/Delete) — Design

**Date:** 2026-07-18
**Status:** Approved (design)
**Builds on:** `docs/superpowers/specs/2026-07-18-gui-redesign-editable-config-design.md` (Phase 1, shipped)

## Problem

The Jira queries tab (Phase 1) only holds the six fixed, automated query slots
(`investigation` … `other_qa`). The user has additional ad-hoc JQL queries they
want to store alongside them and copy to the clipboard on demand, without
wiring them into `Sync Jira -> Excel` automation.

## Goals

1. Let the user add any number of custom query slots: blank title + blank JQL
   text, a Copy button with the same `{project}`/`{account_id}`/`{ws}`/`{we}`
   substitution as the fixed six.
2. Let the user delete a custom slot.
3. Both persist through the existing Save & Close / `config/jql_queries.json`.
4. Zero change to Python — custom slots must never affect `generate_week_rows`
   or `Sync Jira -> Excel`.

## Non-goals (YAGNI)

- No automation/Excel-row generation for custom slots — Copy-only, per
  explicit user choice.
- No limit on the number of custom slots.
- No reordering — custom rows render in creation order, after the fixed six.
- The six fixed slots stay exactly as Phase 1 left them: not deletable, keys
  and order unchanged, only their JQL text stays user-editable. This spec
  does not reopen that constraint — custom slots are a separate, additional
  concept layered alongside them.

## Data model

`config/jql_queries.json` keeps its existing shape:
```json
{ "queries": [ {"key": "...", "label": "...", "jql": "..."}, ... ] }
```
Custom slots are additional entries whose `key` is not one of the six fixed
keys. The key is generated once, at slot creation, as `extra_<8 lowercase
hex>` via `[guid]::NewGuid().ToString('N').Substring(0,8)` — invisible to the
user, stable for the slot's lifetime (never regenerated while the slot
exists).

`jira-sync.py::load_jql()` needs **no changes**. It seeds
`templates = dict(DEFAULT_JQL)` (the six fixed keys) and only overlays an
entry when `key in templates`; any custom-slot entry is silently skipped.
This is already the shipped, tested behavior — the implementation plan adds
one explicit test locking in the contract this feature depends on: an
unrecognized-key entry in `queries` must not raise and must not appear in
`load_jql()`'s returned key set.

## UI / Interaction (`jira-sync.ps1`, `Show-Settings`, Jira queries tab)

Replace the current fixed-Y-coordinate row layout (built directly on
`$panelJql`, a plain `Panel`) with:

- A top strip (`Dock = 'Top'`, never scrolls) holding one **"+ Add query"**
  button.
- A `FlowLayoutPanel` below it (`Dock = 'Fill'`, `FlowDirection = 'TopDown'`,
  `WrapContents = $false`, `AutoScroll = $true`) holding one row-panel per
  query slot: the six fixed rows first (existing order), then custom rows in
  creation order.

This replaces Phase 1's manual Y-increment layout, which cannot survive rows
being added or removed at runtime without re-deriving every row below the
change — the `FlowLayoutPanel` re-flows automatically.

Each row is a single child `Panel`:

- **Fixed slot row** (six of these): unchanged from Phase 1 — bold static
  `Label` title, multiline `TextBox` (JQL), `Copy` button. No Delete button.
- **Custom slot row**: an editable `TextBox` for the title (blank on
  creation), a multiline `TextBox` for JQL (blank on creation), a `Copy`
  button (identical substitution logic to the fixed six), and a small
  **Delete** button that removes this row-panel from the `FlowLayoutPanel`
  immediately — no confirmation, mirroring the standard-tasks grid's Delete
  button.

Clicking **"+ Add query"** appends a new blank custom row to the bottom of
the flow, ready to type — same placeholder syntax as the fixed six
(`{project}`, `{account_id}`, `{ws}`, `{we}`).

Clicking a custom row's **Delete** removes that row from the flow (and from
whatever the Save handler iterates next). If the slot was never saved, it
simply never reaches the config file. If it existed from a prior save, the
next Save & Close is what actually removes it from
`config/jql_queries.json`.

`$script:SaveJqlConfig` is extended to append one entry per surviving custom
row (its stored key, current title-textbox text as `label`, current
JQL-textbox text as `jql`) after the six fixed entries. A custom row with
both title and JQL blank at save time is dropped silently — the same
convention Phase 1 already uses for blank standard-task rows.

On dialog open, after building the six fixed rows, `Show-Settings` scans
`$jqlCfg.queries` for any entry whose key is not one of the six fixed keys
and creates one custom row per such entry, seeded with its stored
key/label/jql.

## Error handling

- Missing/malformed `config/jql_queries.json` → identical to Phase 1 (silent
  defaults / `[WARN]` + defaults for the six fixed slots); no custom rows
  render (nothing to seed them from).
- A custom row's Copy with a blank Account ID field → same behavior as the
  fixed six (a Phase 1 limitation, unchanged, not addressed here).

## Testing

- **Python:** extend `scripts/test_jql_config.py` with a case that seeds a
  `queries` entry using an unrecognized key and asserts `load_jql()`'s
  returned key set is still exactly the six fixed keys — pins the "Python
  needs zero changes" contract this design relies on.
- **GUI:** PowerShell parse check (`ParseFile`) + manual eyeball — add a
  custom row, type a title + JQL, confirm Copy resolves all placeholders,
  Save & Close persists it, reopen Settings shows it still there; delete it,
  Save & Close, reopen confirms it's gone; confirm the six fixed rows are
  unaffected throughout.
