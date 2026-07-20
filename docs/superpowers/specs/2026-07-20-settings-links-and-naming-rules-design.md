# Settings: Links tab + configurable Jira-query naming rules

Date: 2026-07-20
Status: approved by user, ready for implementation plan

## Context

`jira-sync.ps1` is the active GUI (see project memory: "Main GUI file is jira-sync.ps1, not gui.ps1"). Its `Show-Settings` dialog (jira-sync.ps1:343-736) has three tabs: Connection, Standard tasks, Jira queries.

`scripts/jira-sync.py` runs 6 fixed JQL queries per week and inserts one or more Excel rows per query, with the row *name* computed by hardcoded Python logic per query (`generate_week_rows`, jira-sync.py:219-304):

1. `investigation` — total count: `Investigation issue {N}`
2. `bug_verification` — priority-bucketed count: `Bug verification P1-N, P2-N, ...` (zero buckets omitted)
3. `story_creation` — total count: `User story creation {N}`
4. `functional_testing` — one row per issue: `Functional testing of story ID {key}`
5. `regression_testing` — one row per issue, last number extracted from summary: `Regression testing {N} test items`, or `Regression testing [REVIEW] {key}` if no number found
6. `other_qa` — one row per issue, classified by keyword rules in `_AUTOMATION_RULES` (~12 rules) into names like "Automation test creation", "Checklist update", etc., with a `[REVIEW] {summary} ({key})` fallback

The Settings "Jira queries" tab currently shows a static ordinal label per query ("2. Bug verification") that has no relationship to the name actually written to Excel. `config/jql_queries.json` already exists on disk with these ordinal labels persisted under a `label` field — this field is **not** currently read by Python for anything.

## Goals

1. Move the Excel file path into its own Settings tab, alongside a link to the QAE reporting rules doc.
2. Make the row-naming logic for 5 of the 6 built-in queries user-configurable (mode + template) instead of hardcoded, driven directly from the Settings "Jira queries" tab, without changing default behavior for existing users.

`other_qa` naming stays hardcoded (per user decision) — only its Settings row gets a descriptive (non-editable) label explaining the current rules, with the full rule list in a tooltip. Custom queries (`+ Add query`, `extra_*` keys) are unaffected — they remain clipboard-copy-only and are not executed by the sync.

## Part 1 — Links tab

Add a new `TabPage` "Links" to the Settings `TabControl`, after "Jira queries".

- Move the existing Excel-file row (label, textbox, Browse button — jira-sync.ps1:664-682) from the Connection tab to the Links tab. No behavior change: `$tExcel` is referenced by `Test Connection` and `Save & Close` regardless of which tab parents the control.
- Add a `LinkLabel` "QAE Reporting Rules" that opens `https://sitrusllc.sharepoint.com/sites/amcwiki/CompanyRulesandPolicies/Pages/QADPOReportingRules.aspx` via `Start-Process` on click.
- No "Open file" button (explicitly declined).

## Part 2 — Configurable naming rules

### Data model

New top-level field in `config/jql_queries.json`: `name_rules`, a map of the 5 templatable keys to `{ mode, template }`:

```json
"name_rules": {
  "investigation":      { "mode": "count",       "template": "Investigation issue {count}" },
  "bug_verification":   { "mode": "priority",     "template": "Bug verification {buckets}" },
  "story_creation":     { "mode": "count",       "template": "User story creation {count}" },
  "functional_testing": { "mode": "per_issue",    "template": "Functional testing of story ID {key}" },
  "regression_testing": { "mode": "per_issue",    "template": "Regression testing {number} test items" }
}
```

This is independent of the existing `queries[].label` field (left untouched, still written, never read by Python) — this avoids the migration hazard where a saved ordinal label like `"1. Investigation issues"` would silently become a literal (placeholder-free) "template" if `label` were reinterpreted.

Modes and their placeholders (all computed from the same `search_jira` result set, which already fetches `key`, `summary`, `priority` for every query):

| Mode | Placeholders | Aggregation |
|---|---|---|
| `count` | `{count}` | `len(issues)` |
| `priority` | `{P1} {P2} {P3} {count} {buckets}` | bucket each issue's priority into P1 (Highest/High) / P2 (Medium) / P3 (else); `{buckets}` = comma-joined `"P1-N"` parts, zero buckets omitted (matches current behavior) |
| `per_issue` | `{key} {summary} {number}` | one row per issue; `{number}` = last number found in a trailing parenthetical in `summary` (existing `extract_last_number`), or `""` if none |

Special case preserved: for `regression_testing` specifically, if `{number}` resolves empty, the row name falls back to the hardcoded `Regression testing [REVIEW] {key}` regardless of the configured template — this matches current behavior exactly and isn't user-configurable.

Template rendering uses Python `str.format(**ctx)`. If a template references an unknown placeholder (`KeyError`) or fails to render, `jira-sync.py` logs `[WARN] Invalid name_template for <key>: <error> — using default` and falls back to that key's hardcoded default template for the run (never aborts the sync).

### Settings UI (Jira queries tab)

For the 5 templatable rows (not `other_qa`), `New-JqlRow` gains:
- A **mode dropdown** (`ComboBox`, display strings "Total count" / "Priority breakdown (P1-P3)" / "One row per issue" mapped to `count`/`priority`/`per_issue`), placed top-right of the row.
- The existing title control becomes an **editable textbox** holding the naming template (replacing the static ordinal label). Title width narrows from 420 to ~290px to make room for the mode dropdown alongside it; the JQL box below stays full width (420px) as today.
- A `ToolTip` on the template textbox listing that row's current mode's placeholders.

Row order and default JQL text are unchanged. Defaults for mode+template match the table above, so upgrading with an existing `jql_queries.json` (which has no `name_rules` key yet) is a no-op until the user edits something.

`other_qa`'s row keeps its current static `Label` title (not editable), text changed from "6. Other QA activities" to a short description (e.g. "Other QA activities — auto-named by keyword rules"), with `AutoEllipsis` on and a `ToolTip` listing the full `_AUTOMATION_RULES` outcomes (Automation test creation/update/maintenance, Checklist creation/update, Backlog refinement, Debugging, Maintenance, Other project documentation work, `[REVIEW]` fallback).

### Save/Load

- `Show-Settings` loads `name_rules` from `jql_queries.json` at open, overlaying onto PS-side hardcoded defaults (mirrors the existing `$jqlValues` pattern for JQL bodies).
- `$script:SaveJqlConfig` additionally collects `{mode, template}` per templatable row into `name_rules` and writes it into `jql_queries.json` alongside the existing `queries`/`hidden_builtin` keys.

### Python changes (`scripts/jira-sync.py`)

- `DEFAULT_NAME_RULES`: dict matching the table above.
- `load_name_rules()`: analogous to `load_jql()` — overlay `config/jql_queries.json`'s `name_rules` onto defaults per key; an invalid/missing `mode` for a key falls back to that key's default mode.
- Replace the 5 hardcoded naming blocks in `generate_week_rows` with one generic per-mode row-name builder used by all 5 keys; `other_qa`'s block is untouched.
- Existing `[N/6] ... → "name"` log lines are preserved (just sourced from the generic builder's output).

## Out of scope

- `other_qa` keyword-rule table stays hardcoded (not exposed as an editable rules grid).
- Custom (`extra_*`) queries stay clipboard-copy-only; sync does not execute them.
- No "Open file" button for the Excel path.
