# Main Window Redesign — Two-Row Toolbar + Right-Side Weeks Panel — Design

**Date:** 2026-07-18
**Status:** Approved (design)

## Problem

The main window's top area wastes horizontal space: `pnlTop` (Settings button +
connection status dot/text) leaves most of its 1060px width empty, while the
action row below it (`pnlBtns`) crowds eight buttons (Test, Read File, Open
File, Sync, Submit, Fill, Standard, Stop) into one row. Separately, the Weeks
strip (`pnlWeeks`) sits as a full-width horizontal band above the log,
requiring the whole window to grow vertically (and every other panel below it
to shift down) whenever the week list wraps to more rows — and "select
all"/"none" are plain text links rather than a single visible toggle that
reflects the current selection state.

## Goals

1. Merge the two near-empty/over-crowded top rows into one deliberate
   two-row icon toolbar: connection-status/setup actions on row 1,
   work-triggering actions on row 2 — same buttons, same `Add_Click`
   bindings, new grouping/icons/order.
2. Move the Weeks/Month controls, the "file loaded: N weeks found" summary,
   and the Copy Log button into a fixed-width sidebar on the right that spans
   the full remaining window height, so the week list scrolls internally
   instead of resizing the whole window.
3. Replace the "All"/"None" text links with a single tri-state checkbox that
   both reflects and drives the current week-selection state.

## Non-goals (YAGNI)

- No change to what any button *does* — `Add_Click` bodies, Python command
  invocations, and `$script:weekCheckboxes`/`$script:monthCodes` semantics are
  untouched. This is layout + one small piece of new UI-state logic (the
  tri-state checkbox), not a behavior change to Sync/Submit/Fill/Standard.
- No change to the status bar's dot/text semantics (`Set-Status`) — only its
  container loses the Copy Log button.
- No window-width change — stays 1060px (matches this project's existing
  Phase 2 constraint for `Populate-Weeks`' column math, which this design
  replaces anyway, see below).
- No persistence of sidebar/toolbar layout preferences — fixed pixel layout,
  same as every other panel in this app today.

## Toolbar: two rows, same buttons, new grouping

Row 1 (setup/connection) and row 2 (actions that do work) replace the current
`pnlTop` + `pnlBtns` pair. Both rows live in one merged toolbar block at the
top of the window (same total footprint region, not two disconnected areas):

- **Row 1:** connection-status dot + text (unchanged control, `$lblDot`/
  `$lblConnStatus`) · `❓ Test` (`$btnTest`) · `⚙️ Settings` (`$btnSettings`) ·
  `👀 Read File` (`$btnRead`) · `✏️ Open File` (`$btnOpen`)
- **Row 2:** `Standard ⬇️` (`$btnFillStd`) · `Jira ⬇️` (`$btnSync`) ·
  `Submit tasks ➡️` (`$btnSubmit`) · `Fill days ⬆️⬆️⬆️⬆️⬆️` (`$btnFill`) ·
  `⛔ Stop` (`$btnStop`)

Every button keeps its existing PowerShell variable name (`$btnTest`,
`$btnRead`, `$btnOpen`, `$btnSync`, `$btnSubmit`, `$btnFill`, `$btnFillStd`,
`$btnStop`) — only `.Text` (icon + label), `.Location`, and parent panel
change, exactly like the existing (unimplemented) Phase 2 button-cluster plan
already established as safe for this codebase. Row 2's left-to-right order is
Standard → Jira → Submit → Fill → Stop, which reorders `$btnFillStd` before
`$btnSync` relative to today's code (today: Sync, Submit, Fill, Standard,
Stop) — this reorder is intentional, per the approved mockup, not an
oversight.

Icon meaning (confirmed): `👀` = Read File (parses the workbook and prints
week/task info to the log — already existing behavior, icon-only rename);
`✏️` = Open File (launches the Excel file for editing — already existing
`Start-Process` behavior, icon-only rename). `⬇️` on Standard/Jira marks
"writes into the Excel file"; `➡️` on Submit marks "sends onward to D365";
`⬆️⬆️⬆️⬆️⬆️` (five arrows) on Fill Days marks the five weekday columns it
fills; `⛔` on Stop and `❓`/`⚙️` on Test/Settings are literal icon choices
from the approved mockup, not open to reinterpretation.

## Status bar: unchanged except losing Copy Log

`pnlStatus`/`$lblStatusDot`/`$lblStatus` (`Set-Status`) stay exactly as they
are today — same full-width strip above the log, same dot/text semantics
updated by every long-running action. Only `$btnCopyLog` moves out of this
panel into the new sidebar (see below); its own `Add_Click` body (clipboard
copy + "Copied!" timer flash) is unchanged, only its parent panel and
position change.

## Right-side Weeks sidebar (fixed width, full remaining height)

A new panel, `$pnlWeeksSidebar`, replaces `$pnlWeeks` and absorbs
`$btnCopyLog` from the status bar. It sits to the right of the log, spanning
from just below the status bar down to the bottom of the window (fixed
width, `Dock='Right'` on itself with the log `Dock='Fill'` to its left, so
window resize behavior matches the log's existing `Add_Resize` growth
pattern — just horizontally partitioned now instead of the log alone owning
the full width).

Top to bottom inside the sidebar:

1. **Month:** label + `$cmbMonth` (existing control, existing population
   logic in `Populate-Weeks` — unchanged semantics, just relocated and
   docked `Top` in the sidebar instead of positioned in the old horizontal
   strip).
2. **Master checkbox** (`$chkWeeksAll`, new) — tri-state (`ThreeState=$true`),
   labeled "Select all". Reflects and drives the week checkboxes below it
   (see Data Flow).
3. **Week checkboxes** (`$script:weekCheckboxes`, same array/contract as
   today — same `.Tag` = week-start date, same `.Checked` reads by
   Sync/Submit/Fill) — now laid out as a **vertical, single-column,
   scrollable list** (a `Panel` with `AutoScroll=$true`, `Dock='Fill'`
   inside the sidebar between the master checkbox and the bottom section)
   instead of `Populate-Weeks`' current column-of-5 grid. This removes the
   need for `Populate-Weeks` to compute panel height / push `$divider`/
   `$pnlStatus`/`$txtLog` down — the sidebar's fixed-width, `Dock='Fill'`,
   `AutoScroll` list absorbs any number of weeks without resizing anything
   else.
4. **"file loaded: N weeks found"** — a persistent label (`$lblWeeksSummary`,
   new), updated by `$btnRead`'s success handler in addition to (not instead
   of) the existing `Append-Log` line — the log keeps its full history of
   reads; the sidebar shows only the latest one at a glance.
5. **Copy Log button** (`$btnCopyLog`, moved here, `Add_Click` body
   unchanged).

The existing `$lnkSelectAll`/`$lnkNone` `LinkLabel`s are removed entirely —
the master checkbox in position 2 fully replaces their function.

## Data flow: the one new piece of logic (tri-state master checkbox)

This is the only genuinely new behavior in this design (everything else is
layout):

- **Click on `$chkWeeksAll` while `Indeterminate` or `Unchecked`** → check
  every week checkbox (`$script:weekCheckboxes | % { $_.Checked = $true }`),
  master ends up `Checked`.
- **Click on `$chkWeeksAll` while `Checked`** → uncheck every week checkbox,
  master ends up `Unchecked`.
- **Any individual week checkbox's `CheckedChanged`** recomputes
  `$chkWeeksAll.CheckState` from the current set: all checked → `Checked`;
  none checked → `Unchecked`; otherwise → `Indeterminate`. This recompute
  must not itself re-trigger the master's own click-to-toggle-all logic (a
  program-driven `CheckState` assignment must not be mistaken for a user
  click) — the implementation plan will need an explicit reentrancy guard
  (e.g. a script-scope flag checked at the top of each handler) around
  this, since naive `Add_CheckedChanged` wiring on both the master and every
  child checkbox is the standard way this class of bug ships.
- `Populate-Weeks` (rebuilt for the vertical list) wires each new week
  checkbox's `CheckedChanged` to the recompute step above, and sets the
  master's initial `CheckState` once all checkboxes exist (default: all
  checked today, so master starts `Checked`).

Everything downstream of `$script:weekCheckboxes` — `$btnSync`, `$btnSubmit`,
`$btnFillStd`'s click handlers, all of which read `.Checked`/`.Tag` off this
same array — is unchanged. The master checkbox is purely a convenience
control layered on top of the existing array; it is not itself read by any
action handler.

## Error handling

No new failure modes are introduced. The tri-state recompute is pure,
in-memory, deterministic UI-state logic with no I/O — there is nothing to
validate or fail. All existing error handling (missing Excel file, Python
process exit codes, etc.) is untouched by this design.

## Testing

- **Parse check** (same command used for every prior GUI-only task in this
  project): `powershell -NoProfile -Command "$e=$null;$null=
  [System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',
  [ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"` → `OK`.
- **Headless verification of the tri-state logic** (`PerformClick()`/direct
  `.Checked` assignment on real `System.Windows.Forms.CheckBox` instances,
  no visible window needed) — required before this task is considered done,
  following this codebase's own precedent (the JQL row Copy/Delete bug,
  commit `d6a6fa9`, was only caught this way, not by parse-check or diff
  review). Minimum scenarios: check all three individually → master becomes
  `Checked`; uncheck one → master becomes `Indeterminate`; uncheck all →
  master becomes `Unchecked`; click master while `Indeterminate` → all
  become checked; click master while `Checked` → all become unchecked; a
  programmatic recompute of the master's `CheckState` must not itself fire
  the master's click-to-toggle-all handler (reentrancy guard proof).
- **Manual/visual verification** (deferred to a human, as with every prior
  WinForms layout task in this project): toolbar rows don't overlap or
  crowd at 1060px width; sidebar scrolls correctly with a large week count;
  log still resizes correctly on window resize with the sidebar present;
  "file loaded: N weeks found" updates on each Read File click; Copy Log
  still copies the full log text from its new location.
