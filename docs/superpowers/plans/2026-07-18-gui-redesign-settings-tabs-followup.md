# Settings Dialog Follow-up — Connection Reorg, Sprint Length, JQL Row Unification Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reorganize the Settings dialog's Connection/Standard-tasks tabs, make sprint length a real (Python-honored) config value instead of a hardcoded 14-day cycle, and unify the Jira-queries tab's six built-in and user-added rows into one visual/behavioral shape that can all be deleted.

**Architecture:** Three independent slices, done in dependency order: (1) a pure-Python `cycle_days` parameter threaded through `standard_tasks.py`'s sprint-end calculation, tested without any GUI or Excel file; (2) a WinForms layout change moving fields between the Connection and Standard-tasks tabs and relocating the Test Connection button, wired to Python only via the existing `.env`-write/read path (`Save-JiraEnv` / `load_env`); (3) a WinForms refactor collapsing two near-duplicate row-builder functions (`New-JqlFixedRow`, `New-JqlCustomRow`) into one (`New-JqlRow`), with a headless `PerformClick()` script proving its Copy/Delete handlers work for both row shapes — required because the same class of bug (a handler silently resolving to `$null`/no-op) already shipped once in this exact area and was invisible to diff review.

**Tech Stack:** Python 3 (`scripts/standard_tasks.py`, `scripts/jira-sync.py`, plain-script tests), PowerShell WinForms (`jira-sync.ps1`).

**Spec:** `docs/superpowers/specs/2026-07-18-gui-redesign-settings-tabs-followup-design.md`

## Global Constraints

- Deleting a built-in JQL row does **not** disable its automated query in `Sync Jira -> Excel` — `hidden_builtin` only affects the Settings UI and the JQL-text override fallback (confirmed accepted behavior, not a gap to work around).
- No "restore a deleted built-in row" button — restoring one requires manually editing `config/jql_queries.json` (out of scope).
- Sprint length validation is `[WARN]` + default fallback only (default `2` weeks / 14 days), never crash — matches every other config-reading path in this codebase.
- Zero changes to `generate_week_rows`, `search_jira`, or any Excel-writing logic.
- Save & Close remains the dialog's one global, footer-level button. Test Connection's button + status label move into the Connection tab itself (tab-local coordinates) but keep their exact existing behavior.
- Python tests in this repo are plain scripts run as `python scripts/<test>.py`, using a `check()`/assert helper and printing `ALL PASS` — no pytest. Run from the repo root (`scripts/` is `sys.path[0]`, so `from standard_tasks import ...` resolves).
- PowerShell WinForms `Add_Click`/`Add_*` handlers must never close over a builder/helper function's own locals (they're gone by click time) — only `$this`, script-scope (`$script:`) variables, or locals of a function provably still blocked on the call stack (e.g. `Show-Settings` itself, blocked inside `ShowDialog`) for the handler's entire lifetime. See `d6a6fa9` for the shipped bug this rule prevents.
- Platform is WinForms PowerShell + Python 3 — no web/HTML.

## File Structure

- `scripts/standard_tasks.py` — gains `cycle_days` param on `is_sprint_end_week`/`rows_for_week` and a new pure function `parse_sprint_length_weeks`.
- `scripts/jira-sync.py` — `cmd_fill_standard` reads `SPRINT_LENGTH_WEEKS`, computes `cycle_days`, passes it through.
- `scripts/test_standard_tasks.py` — new cases for non-default `cycle_days` and `parse_sprint_length_weeks`.
- `jira-sync.ps1` — `Save-JiraEnv` gains a `SPRINT_LENGTH_WEEKS` param; Connection tab reordered/relabeled with Test Connection relocated into it; Standard-tasks tab gains Sprint length + Sprint end fields above its grid; Jira-queries tab's `New-JqlFixedRow`/`New-JqlCustomRow` collapse into `New-JqlRow`, `hidden_builtin` persistence added, "+ Add query" moves to the bottom with focus/scroll on add.
- `scripts/test_jql_row_gui.ps1` — new headless PowerShell script proving `New-JqlRow`'s Copy/Delete handlers work for both row shapes via `PerformClick()`, no visible window needed.

---

### Task 1: Sprint length as a real Python parameter (`cycle_days`)

**Files:**
- Modify: `scripts/standard_tasks.py`
- Modify: `scripts/jira-sync.py`
- Test: `scripts/test_standard_tasks.py`

**Interfaces:**
- Consumes: nothing new — pure extension of existing `is_sprint_end_week(week_start, sprint_anchor)`, `rows_for_week(week_start, week_end, month, sprint_anchor, today, schedule=None, placeholders=None)`.
- Produces: `is_sprint_end_week(week_start, sprint_anchor, cycle_days=14)`, `rows_for_week(..., cycle_days=14)`, `parse_sprint_length_weeks(raw: str) -> tuple[int, str | None]` (returns `(weeks, warning_or_None)`). Task 2's GUI writes the `.env` key `SPRINT_LENGTH_WEEKS` that `cmd_fill_standard` reads here — the two tasks share only that key name, no code dependency.

- [ ] **Step 1: Write the failing tests**

Open `scripts/test_standard_tasks.py`. Find:

```python
from datetime import date
from standard_tasks import rows_for_week, is_sprint_end_week
```

Replace with:

```python
from datetime import date
from standard_tasks import rows_for_week, is_sprint_end_week, parse_sprint_length_weeks
```

Find the block:

```python
# --- sprint-end present again two weeks later (2026-07-26, friday 2026-07-31) ---
nxt = by_name(rows_for_week(date(2026, 7, 26), date(2026, 8, 1), None, ANCHOR, PAST))
check("sprint review +14", "Internal sprint review" in nxt, True)

# --- no anchor -> sprint-end tasks skipped, weekly ones stay ---
```

Replace with (adds the new cases right after the existing `+14` check, before the no-anchor case):

```python
# --- sprint-end present again two weeks later (2026-07-26, friday 2026-07-31) ---
nxt = by_name(rows_for_week(date(2026, 7, 26), date(2026, 8, 1), None, ANCHOR, PAST))
check("sprint review +14", "Internal sprint review" in nxt, True)

# --- cycle_days is a real parameter: a 3-week (21-day) sprint disagrees with
# the default 14-day cycle on the +14 week, and agrees again on the +21 week ---
check("14-day cycle flags +14 week", is_sprint_end_week(date(2026, 7, 26), ANCHOR), True)
check("21-day cycle does NOT flag +14 week", is_sprint_end_week(date(2026, 7, 26), ANCHOR, cycle_days=21), False)
check("21-day cycle flags +21 week", is_sprint_end_week(date(2026, 8, 2), ANCHOR, cycle_days=21), True)

rows_21_plus14 = by_name(rows_for_week(date(2026, 7, 26), date(2026, 8, 1), None, ANCHOR, PAST, cycle_days=21))
check("rows_for_week cycle_days=21: sprint review absent on +14 week",
      "Internal sprint review" in rows_21_plus14, False)
rows_21_plus21 = by_name(rows_for_week(date(2026, 8, 2), date(2026, 8, 8), None, ANCHOR, PAST, cycle_days=21))
check("rows_for_week cycle_days=21: sprint review present on +21 week",
      "Internal sprint review" in rows_21_plus21, True)

# --- parse_sprint_length_weeks: WARN + default 2, never crash ---
check("empty string -> default 2, no warning", parse_sprint_length_weeks(""), (2, None))
check("whitespace -> default 2, no warning", parse_sprint_length_weeks("   "), (2, None))
check("valid '3' -> (3, None)", parse_sprint_length_weeks("3"), (3, None))
w_bad, warn_bad = parse_sprint_length_weeks("abc")
check("invalid 'abc' -> weeks defaults to 2", w_bad, 2)
check("invalid 'abc' -> warning present", warn_bad is not None, True)
w_zero, warn_zero = parse_sprint_length_weeks("0")
check("zero -> weeks defaults to 2", w_zero, 2)
check("zero -> warning present", warn_zero is not None, True)
w_neg, warn_neg = parse_sprint_length_weeks("-1")
check("negative -> weeks defaults to 2", w_neg, 2)
check("negative -> warning present", warn_neg is not None, True)

# --- no anchor -> sprint-end tasks skipped, weekly ones stay ---
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python scripts/test_standard_tasks.py`
Expected: `TypeError: is_sprint_end_week() got an unexpected keyword argument 'cycle_days'` (or an `ImportError: cannot import name 'parse_sprint_length_weeks'` if Python reaches that import failure first).

- [ ] **Step 3: Implement `cycle_days` in `standard_tasks.py`**

Find:

```python
def is_sprint_end_week(week_start: date, sprint_anchor: date) -> bool:
    """A week is a sprint end iff its Friday is a whole number of 2-week cycles
    away from the anchor Friday."""
    return (_friday(week_start) - sprint_anchor).days % 14 == 0


def rows_for_week(week_start, week_end, month, sprint_anchor, today,
                  schedule=None, placeholders=None):
```

Replace with:

```python
def is_sprint_end_week(week_start: date, sprint_anchor: date, cycle_days: int = 14) -> bool:
    """A week is a sprint end iff its Friday is a whole number of `cycle_days`-day
    cycles away from the anchor Friday. Default cycle_days=14 (2-week sprint)."""
    return (_friday(week_start) - sprint_anchor).days % cycle_days == 0


def parse_sprint_length_weeks(raw: str):
    """Parse the SPRINT_LENGTH_WEEKS .env value. Empty -> (2, None). Present
    but not a positive integer -> (2, warning_message) — never raises."""
    raw = (raw or "").strip()
    if not raw:
        return 2, None
    try:
        weeks = int(raw)
        if weeks <= 0:
            raise ValueError("must be positive")
        return weeks, None
    except ValueError:
        return 2, f"[WARN] SPRINT_LENGTH_WEEKS {raw!r} invalid (want a positive integer); using default 2"


def rows_for_week(week_start, week_end, month, sprint_anchor, today,
                  schedule=None, placeholders=None, cycle_days=14):
```

Find (inside `rows_for_week`):

```python
    for task in schedule:
        if task["freq"] == "sprint-end":
            if sprint_anchor is None or not is_sprint_end_week(week_start, sprint_anchor):
                continue
```

Replace with:

```python
    for task in schedule:
        if task["freq"] == "sprint-end":
            if sprint_anchor is None or not is_sprint_end_week(week_start, sprint_anchor, cycle_days):
                continue
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python scripts/test_standard_tasks.py`
Expected: `ALL PASS`, including all the new lines added in Step 1.

- [ ] **Step 5: Wire `cycle_days` into `cmd_fill_standard`**

Open `scripts/jira-sync.py`. Find the import line:

```python
from standard_tasks import rows_for_week, DAY_COL, load_schedule
```

Replace with:

```python
from standard_tasks import rows_for_week, DAY_COL, load_schedule, parse_sprint_length_weeks
```

Find (inside `cmd_fill_standard`):

```python
    else:
        print("[WARN] SPRINT_ANCHOR not set; sprint-end tasks (sprint review, summary report) skipped", flush=True)

    wb_read = openpyxl.load_workbook(args.file, data_only=True)
```

Replace with:

```python
    else:
        print("[WARN] SPRINT_ANCHOR not set; sprint-end tasks (sprint review, summary report) skipped", flush=True)

    sprint_length_weeks, sprint_len_warning = parse_sprint_length_weeks(env.get("SPRINT_LENGTH_WEEKS", ""))
    if sprint_len_warning:
        print(sprint_len_warning, flush=True)
    cycle_days = sprint_length_weeks * 7

    wb_read = openpyxl.load_workbook(args.file, data_only=True)
```

Find:

```python
        rows = rows_for_week(week["week_start"], we, args.month, sprint_anchor, today,
                             schedule=schedule, placeholders=placeholders)
```

Replace with:

```python
        rows = rows_for_week(week["week_start"], we, args.month, sprint_anchor, today,
                             schedule=schedule, placeholders=placeholders, cycle_days=cycle_days)
```

- [ ] **Step 6: Re-run the Python test suite to confirm no regressions**

Run: `python scripts/test_standard_tasks.py`
Expected: `ALL PASS`.

Run: `python scripts/test_month_filter.py`
Run: `python scripts/test_standard_config.py`
Run: `python scripts/test_standard_insert.py`
Run: `python scripts/test_excel_insert.py`
Run: `python scripts/test_jql_config.py`
Expected for all: `ALL PASS` (this task doesn't touch any of these files — a regression here means something unexpected happened).

- [ ] **Step 7: Commit**

```bash
git add scripts/standard_tasks.py scripts/jira-sync.py scripts/test_standard_tasks.py
git commit -m "feat: sprint length becomes a real cycle_days parameter, not hardcoded 14"
```

---

### Task 2: Connection tab reorder + Sprint length/Sprint end on Standard tasks tab

**Files:**
- Modify: `jira-sync.ps1` (`Save-JiraEnv` at lines 19-34; Connection tab block at lines 605-699; Standard-tasks tab grid setup starting at line 497; Save & Close handler at lines 707-712)

**Interfaces:**
- Consumes: `Add-Row` (existing helper, lines 592-603, unchanged), `Read-EnvFile`, `Write-JsonConfig`, `$scriptDir`, `$lblDot`/`$lblConnStatus` (main-window controls, already proven readable from inside `Show-Settings`'s click handlers today).
- Produces: `Save-JiraEnv($url, $email, $token, $project, $accountId, $excelFile, $sprintAnchor, $sprintLengthWeeks)` (one new trailing param); `$tSprintLen` (Standard-tasks tab TextBox, new); `$tAnchor` (Sprint end TextBox, now built inside the Standard-tasks tab block instead of Connection). Task 3 does not touch any of these.

- [ ] **Step 1: Add `SPRINT_LENGTH_WEEKS` to `Save-JiraEnv`**

Find (jira-sync.ps1:19-34):

```powershell
function Save-JiraEnv($url, $email, $token, $project, $accountId, $excelFile, $sprintAnchor) {
    $keys = @('JIRA_URL','JIRA_EMAIL','JIRA_API_TOKEN','JIRA_PROJECT','JIRA_ACCOUNT_ID','EXCEL_FILE','SPRINT_ANCHOR')
    $vals = @{ JIRA_URL=$url; JIRA_EMAIL=$email; JIRA_API_TOKEN=$token; JIRA_PROJECT=$project; JIRA_ACCOUNT_ID=$accountId; EXCEL_FILE=$excelFile; SPRINT_ANCHOR=$sprintAnchor }
```

Replace with:

```powershell
function Save-JiraEnv($url, $email, $token, $project, $accountId, $excelFile, $sprintAnchor, $sprintLengthWeeks) {
    $keys = @('JIRA_URL','JIRA_EMAIL','JIRA_API_TOKEN','JIRA_PROJECT','JIRA_ACCOUNT_ID','EXCEL_FILE','SPRINT_ANCHOR','SPRINT_LENGTH_WEEKS')
    $vals = @{ JIRA_URL=$url; JIRA_EMAIL=$email; JIRA_API_TOKEN=$token; JIRA_PROJECT=$project; JIRA_ACCOUNT_ID=$accountId; EXCEL_FILE=$excelFile; SPRINT_ANCHOR=$sprintAnchor; SPRINT_LENGTH_WEEKS=$sprintLengthWeeks }
```

(The rest of the function body — the `$lines`/`$written` merge loop — is unchanged; it already iterates `$keys` generically.)

- [ ] **Step 2: Reorder the Connection tab and relocate Test Connection into it**

Find the entire block from `$d = Read-EnvFile` (jira-sync.ps1:605) through the closing `[void]$dlg.Controls.Add($btnT)` (jira-sync.ps1:699). Replace it in full with:

```powershell
    $d = Read-EnvFile
    $tUrl  = Add-Row $tabConn 'Jira URL:'    16
    $tTok  = Add-Row $tabConn 'API Token:'   54 $true

    $chkSh = New-Object System.Windows.Forms.CheckBox
    $chkSh.Text = 'Show token'; $chkSh.Location = New-Object System.Drawing.Point(114,80)
    $chkSh.Size = New-Object System.Drawing.Size(100,22)
    $chkSh.Add_CheckedChanged({ $tTok.UseSystemPasswordChar = -not $chkSh.Checked })
    [void]$tabConn.Controls.Add($chkSh)

    $tProj = Add-Row $tabConn 'Project key:' 110
    $tMail = Add-Row $tabConn 'Email:'       148
    $tAcct = Add-Row $tabConn 'Account ID:'  186

    $tUrl.Text  = if ($d['JIRA_URL'])        { $d['JIRA_URL'] }        else { 'https://amcbridge.atlassian.net' }
    $tTok.Text  = if ($d['JIRA_API_TOKEN'])   { $d['JIRA_API_TOKEN'] }   else { '' }
    $tProj.Text = if ($d['JIRA_PROJECT'])     { $d['JIRA_PROJECT'] }     else { 'GT2' }
    $tMail.Text = if ($d['JIRA_EMAIL'])       { $d['JIRA_EMAIL'] }       else { '' }
    $tAcct.Text = if ($d['JIRA_ACCOUNT_ID'])  { $d['JIRA_ACCOUNT_ID'] }  else { '' }

    # Hint: which Jira user the activity queries (2,3,4,5,6) filter on. Empty = token owner.
    $lblAcctHint = New-Object System.Windows.Forms.Label
    $lblAcctHint.Text = 'Leave empty to use the API-token owner (/myself)'
    $lblAcctHint.Location = New-Object System.Drawing.Point(114,211)
    $lblAcctHint.Size = New-Object System.Drawing.Size(390,16)
    $lblAcctHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAcctHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$tabConn.Controls.Add($lblAcctHint)

    # ---- Excel file row ----
    $lblEx = New-Object System.Windows.Forms.Label
    $lblEx.Text = 'Excel file:'; $lblEx.Location = New-Object System.Drawing.Point(16,237)
    $lblEx.Size = New-Object System.Drawing.Size(90,20); $lblEx.TextAlign = 'MiddleRight'
    [void]$tabConn.Controls.Add($lblEx)
    $tExcel = New-Object System.Windows.Forms.TextBox
    $tExcel.Location = New-Object System.Drawing.Point(114,234)
    $tExcel.Size = New-Object System.Drawing.Size(354,24)
    $tExcel.Text = $txtFile.Text
    [void]$tabConn.Controls.Add($tExcel)
    $btnExBrowse = New-Object System.Windows.Forms.Button
    $btnExBrowse.Text = '...'; $btnExBrowse.Location = New-Object System.Drawing.Point(474,234)
    $btnExBrowse.Size = New-Object System.Drawing.Size(30,24); $btnExBrowse.FlatStyle = 'Flat'
    $btnExBrowse.Add_Click({
        $fd = New-Object System.Windows.Forms.OpenFileDialog
        $fd.Filter = 'Excel files (*.xlsx)|*.xlsx'
        $fd.InitialDirectory = Split-Path $tExcel.Text -Parent
        if ($fd.ShowDialog() -eq 'OK') { $tExcel.Text = $fd.FileName }
    })
    [void]$tabConn.Controls.Add($btnExBrowse)

    $lblTest = New-Object System.Windows.Forms.Label
    $lblTest.Location = New-Object System.Drawing.Point(16,272)
    $lblTest.Size = New-Object System.Drawing.Size(460,20)
    $lblTest.ForeColor = [System.Drawing.Color]::Gray
    [void]$tabConn.Controls.Add($lblTest)

    $btnT = New-Object System.Windows.Forms.Button
    $btnT.Text = 'Test Connection'
    $btnT.Location = New-Object System.Drawing.Point(16,296)
    $btnT.Size = New-Object System.Drawing.Size(140,32)
    $btnT.Add_Click({
        $lblTest.Text = 'Testing...'; $lblTest.ForeColor = [System.Drawing.Color]::DodgerBlue
        $dlg.Refresh()
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text $tSprintLen.Text
        $psi2 = New-Object System.Diagnostics.ProcessStartInfo
        $psi2.FileName = 'python'; $psi2.Arguments = "scripts\jira-sync.py --command test --file `"$($tExcel.Text)`""
        $psi2.WorkingDirectory = $scriptDir; $psi2.UseShellExecute = $false
        $psi2.RedirectStandardOutput = $true; $psi2.RedirectStandardError = $true
        $psi2.CreateNoWindow = $true
        $psi2.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi2.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
        $p2 = [System.Diagnostics.Process]::Start($psi2)
        $out = $p2.StandardOutput.ReadToEnd(); $p2.WaitForExit()
        if ($p2.ExitCode -eq 0) {
            $lblTest.Text = 'Connected OK'; $lblTest.ForeColor = [System.Drawing.Color]::DarkGreen
            $lblDot.ForeColor = [System.Drawing.Color]::LimeGreen
            $lblConnStatus.Text = 'Connected'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,100,220,100)
        } else {
            $lblTest.Text = 'Connection failed - check credentials'; $lblTest.ForeColor = [System.Drawing.Color]::Red
            $lblDot.ForeColor = [System.Drawing.Color]::OrangeRed
            $lblConnStatus.Text = 'Not connected'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::OrangeRed
        }
    })
    [void]$tabConn.Controls.Add($btnT)
```

`$tAnchor` and `$tSprintLen` are referenced inside `$btnT.Add_Click` here but are not defined until Step 3 (Standard-tasks tab). This is safe by the same rule already proven in this codebase: `Show-Settings` is still executing — blocked inside `$dlg.ShowDialog($form)` — for the entire time a user could click this button, so by the time a click actually happens, every variable `Show-Settings` has ever assigned (regardless of source-line order) is resolvable on the scope chain. Do not reorder Steps 2/3 relative to each other based on this dependency — order in the *file* doesn't matter here, only that both blocks run before `ShowDialog` is called, which they do.

- [ ] **Step 3: Add Sprint length + Sprint end to the Standard-tasks tab, shift the grid down**

Find (jira-sync.ps1, starts at line 497):

```powershell
    $gridStd = New-Object System.Windows.Forms.DataGridView
    $gridStd.Location = New-Object System.Drawing.Point(8,8)
    $gridStd.Size = New-Object System.Drawing.Size(572,350)
```

Replace with:

```powershell
    $d = Read-EnvFile
    $lblSprintLen = New-Object System.Windows.Forms.Label
    $lblSprintLen.Text = 'Sprint length (weeks):'
    $lblSprintLen.Location = New-Object System.Drawing.Point(8,10)
    $lblSprintLen.Size = New-Object System.Drawing.Size(140,20)
    [void]$tabStd.Controls.Add($lblSprintLen)
    $tSprintLen = New-Object System.Windows.Forms.TextBox
    $tSprintLen.Location = New-Object System.Drawing.Point(152,8)
    $tSprintLen.Size = New-Object System.Drawing.Size(40,24)
    [void]$tabStd.Controls.Add($tSprintLen)

    $lblAnchorRow = New-Object System.Windows.Forms.Label
    $lblAnchorRow.Text = 'Sprint end:'
    $lblAnchorRow.Location = New-Object System.Drawing.Point(210,10)
    $lblAnchorRow.Size = New-Object System.Drawing.Size(70,20)
    [void]$tabStd.Controls.Add($lblAnchorRow)
    $tAnchor = New-Object System.Windows.Forms.TextBox
    $tAnchor.Location = New-Object System.Drawing.Point(284,8)
    $tAnchor.Size = New-Object System.Drawing.Size(110,24)
    [void]$tabStd.Controls.Add($tAnchor)

    $lblAnchorHint = New-Object System.Windows.Forms.Label
    $lblAnchorHint.Location = New-Object System.Drawing.Point(8,36)
    $lblAnchorHint.Size = New-Object System.Drawing.Size(572,16)
    $lblAnchorHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAnchorHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$tabStd.Controls.Add($lblAnchorHint)

    function Update-AnchorHint {
        $weeks = 2
        $parsed = 0
        if ([int]::TryParse($tSprintLen.Text, [ref]$parsed) -and $parsed -gt 0) { $weeks = $parsed }
        $lblAnchorHint.Text = "Any sprint-end Friday (YYYY-MM-DD) - $weeks-week cycles counted from here"
    }
    $tSprintLen.Add_TextChanged({ Update-AnchorHint })

    $tAnchor.Text = if ($d['SPRINT_ANCHOR']) { $d['SPRINT_ANCHOR'] } else { '' }
    $tSprintLen.Text = if ($d['SPRINT_LENGTH_WEEKS']) { $d['SPRINT_LENGTH_WEEKS'] } else { '2' }
    Update-AnchorHint

    $gridStd = New-Object System.Windows.Forms.DataGridView
    $gridStd.Location = New-Object System.Drawing.Point(8,60)
    $gridStd.Size = New-Object System.Drawing.Size(572,304)
```

`Update-AnchorHint` (a function defined inside `Show-Settings`) is called from `$tSprintLen.Add_TextChanged({ Update-AnchorHint })`. This is safe for the same reason as Step 2's forward reference: the scriptblock resolves `Update-AnchorHint` against the live scope chain when the event actually fires, and `Show-Settings` — where the function is defined — is still on the call stack (blocked in `ShowDialog`) for the whole time the dialog can raise `TextChanged` events.

Now find (a few lines below, originally at line 526):

```powershell
    $lblStdHint = New-Object System.Windows.Forms.Label
    $lblStdHint.Text = 'One hours value per task (same each day it occurs); blank = not that day. Placeholders: name only.'
    $lblStdHint.Location = New-Object System.Drawing.Point(8,362); $lblStdHint.Size = New-Object System.Drawing.Size(572,18)
```

Replace with:

```powershell
    $lblStdHint = New-Object System.Windows.Forms.Label
    $lblStdHint.Text = 'One hours value per task (same each day it occurs); blank = not that day. Placeholders: name only.'
    $lblStdHint.Location = New-Object System.Drawing.Point(8,368); $lblStdHint.Size = New-Object System.Drawing.Size(572,18)
```

Everything below this (grid columns, `Add_CellContentClick`, `$stdCfg`/`$schedule`/`$placeholders` loading, `$script:SaveStandardFromGrid`) is unchanged — none of it references `$tAnchor`/`$tSprintLen` or the grid's Y-position.

- [ ] **Step 4: Pass the new field into Save & Close**

Find (jira-sync.ps1:707-712):

```powershell
    $btnSv.Add_Click({
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text
        & $script:SaveStandardFromGrid
        & $script:SaveJqlConfig
        $txtFile.Text = $tExcel.Text
    })
```

Replace with:

```powershell
    $btnSv.Add_Click({
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text $tSprintLen.Text
        & $script:SaveStandardFromGrid
        & $script:SaveJqlConfig
        $txtFile.Text = $tExcel.Text
    })
```

- [ ] **Step 5: Verify parse**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

- [ ] **Step 6: Manual verification (defer to human — no live GUI in this environment)**

Note in your report that this step is pending: launch `jira-sync.ps1`, open Settings.
- Connection tab shows, top to bottom: Jira URL, API Token (+ Show token directly under it), Project key, Email, Account ID (+ hint), Excel file (+ browse), Test Connection button + status label — no Sprint end field on this tab.
- Standard tasks tab shows Sprint length (weeks) and Sprint end side by side above the grid; the hint text below them reads e.g. "2-week cycles" by default and updates live to "3-week cycles" (etc.) as you type into Sprint length.
- Click Test Connection: still calls `python scripts\jira-sync.py --command test`, updates the same status label/dot/top-bar indicator as before.
- Change Sprint length to `3`, set an Excel file, Save & Close, reopen Settings: Sprint length still shows `3`, Sprint end unchanged, hint reflects `3-week cycles`.
- Inspect `.env` after Save & Close: contains a `SPRINT_LENGTH_WEEKS=3` line.

- [ ] **Step 7: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: reorder/relabel Connection tab, move Sprint fields to Standard tasks tab"
```

---

### Task 3: Unify Jira-queries rows into `New-JqlRow`, add `hidden_builtin` delete-and-persist, move "+ Add query" below the list

**Files:**
- Modify: `jira-sync.ps1` (`$script:jqlBoxes`/`New-JqlFixedRow`/`New-JqlCustomRow`/loading loops at lines 376-495; `$script:SaveJqlConfig` at lines 576-590)
- Test: `scripts/test_jql_row_gui.ps1` (new)

**Interfaces:**
- Consumes: `$jqlSlots`, `$jqlValues`, `$jqlCfg` (all already parsed earlier in `Show-Settings`, untouched by this task), `$tProj`, `$tAcct`, `$script:weekCheckboxes`, `$panelJqlFlow` (from the existing `FlowLayoutPanel` refactor, untouched).
- Produces: `New-JqlRow($key, $titleText, $jqlText, $titleEditable)` — builds one row (title is a read-only `Label` when `$titleEditable=$false`, an editable `TextBox` when `$true`; both cases get Copy + Delete), adds it to `$panelJqlFlow`, and **returns the row `Panel`**. `$script:jqlBoxes` and the two old builder functions are removed entirely — nothing outside this task's own code referenced them (confirmed: `$script:SaveJqlConfig`, the only other reader, is rewritten in this same task).

- [ ] **Step 1: Replace `New-JqlFixedRow`/`New-JqlCustomRow` with `New-JqlRow`**

Find the entire block from `$script:jqlBoxes = @{}` (jira-sync.ps1:376) through the closing `[void]$tabJql.Controls.Add($panelJqlAddStrip)` (jira-sync.ps1:495). Replace it in full with:

```powershell
    function New-JqlRow($key, $titleText, $jqlText, $titleEditable) {
        $row = New-Object System.Windows.Forms.Panel
        $row.Size = New-Object System.Drawing.Size(566,74)
        $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
        $row.Tag = $key

        if ($titleEditable) {
            $titleCtl = New-Object System.Windows.Forms.TextBox
        } else {
            $titleCtl = New-Object System.Windows.Forms.Label
        }
        $titleCtl.Text = $titleText; $titleCtl.Location = New-Object System.Drawing.Point(0,0)
        $titleCtl.Size = New-Object System.Drawing.Size(420,18)
        $titleCtl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
        [void]$row.Controls.Add($titleCtl)

        $box = New-Object System.Windows.Forms.TextBox
        $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
        $box.Location = New-Object System.Drawing.Point(0,20); $box.Size = New-Object System.Drawing.Size(420,44)
        $box.Font = New-Object System.Drawing.Font('Consolas',8)
        $box.Text = $jqlText
        [void]$row.Controls.Add($box)

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(424,20); $btnCopy.Size = New-Object System.Drawing.Size(60,21)
        $btnCopy.FlatStyle = 'Flat'
        $btnCopy.Add_Click({
            $tpl = [string]$this.Parent.Controls[1].Text
            $proj = if ($tProj.Text) { $tProj.Text } else { 'GT2' }
            $acct = $tAcct.Text
            $checked = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
            if ($checked.Count -gt 0) {
                $ws = [string]$checked[0]
                $we = ([datetime]::ParseExact($ws,'yyyy-MM-dd',$null).AddDays(6)).ToString('yyyy-MM-dd')
            } else {
                $today = Get-Date
                $sunday = $today.AddDays(-[int]$today.DayOfWeek)
                $ws = $sunday.ToString('yyyy-MM-dd'); $we = $sunday.AddDays(6).ToString('yyyy-MM-dd')
            }
            $resolved = $tpl.Replace('{project}',$proj).Replace('{account_id}',$acct).Replace('{ws}',$ws).Replace('{we}',$we)
            [System.Windows.Forms.Clipboard]::SetText($resolved)
        })
        [void]$row.Controls.Add($btnCopy)

        $btnDel = New-Object System.Windows.Forms.Button
        $btnDel.Text = 'Delete'; $btnDel.Location = New-Object System.Drawing.Point(424,44); $btnDel.Size = New-Object System.Drawing.Size(60,20)
        $btnDel.FlatStyle = 'Flat'; $btnDel.ForeColor = [System.Drawing.Color]::FromArgb(255,180,40,40)
        $btnDel.Add_Click({
            $this.Parent.Parent.Controls.Remove($this.Parent)
        })
        [void]$row.Controls.Add($btnDel)

        [void]$panelJqlFlow.Controls.Add($row)
        return $row
    }

    $hiddenBuiltin = @{}
    if ($jqlCfg -and $jqlCfg.PSObject.Properties.Match('hidden_builtin').Count -and $jqlCfg.hidden_builtin) {
        foreach ($hk in $jqlCfg.hidden_builtin) { $hiddenBuiltin[$hk] = $true }
    }

    foreach ($slot in $jqlSlots) {
        if ($hiddenBuiltin.ContainsKey($slot.key)) { continue }
        New-JqlRow $slot.key $slot.label ([string]$jqlValues[$slot.key]) $false
    }

    if ($jqlCfg -and $jqlCfg.queries) {
        $fixedKeySet = @{}
        foreach ($s in $jqlSlots) { $fixedKeySet[$s.key] = $true }
        foreach ($q in $jqlCfg.queries) {
            if (-not $fixedKeySet.ContainsKey($q.key)) {
                New-JqlRow $q.key $q.label $q.jql $true
            }
        }
    }

    $panelJqlAddStrip = New-Object System.Windows.Forms.Panel
    $panelJqlAddStrip.Dock = 'Bottom'; $panelJqlAddStrip.Height = 32
    $btnAddJql = New-Object System.Windows.Forms.Button
    $btnAddJql.Text = '+ Add query'; $btnAddJql.Location = New-Object System.Drawing.Point(0,2)
    $btnAddJql.Size = New-Object System.Drawing.Size(110,26); $btnAddJql.FlatStyle = 'Flat'
    $btnAddJql.Add_Click({
        $newKey = 'extra_' + ([guid]::NewGuid().ToString('N').Substring(0,8))
        $newRow = New-JqlRow $newKey '' '' $true
        $panelJqlFlow.ScrollControlIntoView($newRow)
        $newRow.Controls[0].Focus()
    })
    [void]$panelJqlAddStrip.Controls.Add($btnAddJql)
    [void]$tabJql.Controls.Add($panelJqlAddStrip)
```

Notes on this replacement:
- `$panelJqlAddStrip.Dock` changed from `'Top'` to `'Bottom'` — no reordering of the two `.Controls.Add` calls is needed: `$panelJqlFlow` (`Dock='Fill'`) was already added to `$tabJql` *before* this block runs (in the untouched code above line 376), and `$panelJqlAddStrip` is still added *after* it here, which is what makes the add-strip claim its docked edge (now Bottom instead of Top) rather than being covered by the fill panel.
- Both Copy and Delete use `$this.Parent...` control-tree navigation for every row, fixed-shaped or custom-shaped alike — this is the exact pattern proven safe by commit `d6a6fa9` (the shipped bug's fix), applied uniformly instead of only to custom rows.
- `$box.Text = $jqlText` for a fixed-shaped row now takes `[string]$jqlValues[$slot.key]` directly as a constructor argument instead of being looked up later through `$script:jqlBoxes` — there is no longer a `$script:jqlBoxes` dictionary at all.

- [ ] **Step 2: Rewrite `$script:SaveJqlConfig` as one uniform loop**

Find (jira-sync.ps1:576-590):

```powershell
    $script:SaveJqlConfig = {
        $queries = @()
        foreach ($slot in $jqlSlots) {
            $queries += @{ key=$slot.key; label=$slot.label; jql=[string]$script:jqlBoxes[$slot.key].Text }
        }
        $fixedKeys = @($jqlSlots | ForEach-Object { $_.key })
        foreach ($row in $panelJqlFlow.Controls) {
            if ($fixedKeys -contains $row.Tag) { continue }
            $title = [string]$row.Controls[0].Text
            $jql   = [string]$row.Controls[1].Text
            if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($jql)) { continue }
            $queries += @{ key=$row.Tag; label=$title; jql=$jql }
        }
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries }
    }
```

Replace with:

```powershell
    $script:SaveJqlConfig = {
        $queries = @()
        $fixedKeys = @($jqlSlots | ForEach-Object { $_.key })
        $presentFixed = @{}
        foreach ($row in $panelJqlFlow.Controls) {
            $title = [string]$row.Controls[0].Text
            $jql   = [string]$row.Controls[1].Text
            if ($fixedKeys -contains $row.Tag) { $presentFixed[$row.Tag] = $true }
            if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($jql)) { continue }
            $queries += @{ key=$row.Tag; label=$title; jql=$jql }
        }
        $hidden = @($fixedKeys | Where-Object { -not $presentFixed.ContainsKey($_) })
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries; hidden_builtin=$hidden }
    }
```

A fixed-shaped row's title is a read-only `Label` whose text is always its slot's fixed label (never blank), so the blank-title-and-blank-jql skip only ever drops custom rows the user left empty — a present fixed row is never accidentally dropped from `$queries` or marked hidden.

- [ ] **Step 3: Verify parse**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

- [ ] **Step 4: Write the headless `New-JqlRow` verification script**

Create `scripts/test_jql_row_gui.ps1`:

```powershell
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:weekCheckboxes = @()
$tProj = New-Object System.Windows.Forms.TextBox; $tProj.Text = 'GT2'
$tAcct = New-Object System.Windows.Forms.TextBox; $tAcct.Text = 'ACC123'
$panelJqlFlow = New-Object System.Windows.Forms.FlowLayoutPanel

function New-JqlRow($key, $titleText, $jqlText, $titleEditable) {
    $row = New-Object System.Windows.Forms.Panel
    $row.Size = New-Object System.Drawing.Size(566,74)
    $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
    $row.Tag = $key

    if ($titleEditable) {
        $titleCtl = New-Object System.Windows.Forms.TextBox
    } else {
        $titleCtl = New-Object System.Windows.Forms.Label
    }
    $titleCtl.Text = $titleText; $titleCtl.Location = New-Object System.Drawing.Point(0,0)
    $titleCtl.Size = New-Object System.Drawing.Size(420,18)
    $titleCtl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    [void]$row.Controls.Add($titleCtl)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
    $box.Location = New-Object System.Drawing.Point(0,20); $box.Size = New-Object System.Drawing.Size(420,44)
    $box.Font = New-Object System.Drawing.Font('Consolas',8)
    $box.Text = $jqlText
    [void]$row.Controls.Add($box)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(424,20); $btnCopy.Size = New-Object System.Drawing.Size(60,21)
    $btnCopy.FlatStyle = 'Flat'
    $btnCopy.Add_Click({
        $tpl = [string]$this.Parent.Controls[1].Text
        $proj = if ($tProj.Text) { $tProj.Text } else { 'GT2' }
        $acct = $tAcct.Text
        $checked = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
        if ($checked.Count -gt 0) {
            $ws = [string]$checked[0]
            $we = ([datetime]::ParseExact($ws,'yyyy-MM-dd',$null).AddDays(6)).ToString('yyyy-MM-dd')
        } else {
            $today = Get-Date
            $sunday = $today.AddDays(-[int]$today.DayOfWeek)
            $ws = $sunday.ToString('yyyy-MM-dd'); $we = $sunday.AddDays(6).ToString('yyyy-MM-dd')
        }
        $resolved = $tpl.Replace('{project}',$proj).Replace('{account_id}',$acct).Replace('{ws}',$ws).Replace('{we}',$we)
        [System.Windows.Forms.Clipboard]::SetText($resolved)
    })
    [void]$row.Controls.Add($btnCopy)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Delete'; $btnDel.Location = New-Object System.Drawing.Point(424,44); $btnDel.Size = New-Object System.Drawing.Size(60,20)
    $btnDel.FlatStyle = 'Flat'; $btnDel.ForeColor = [System.Drawing.Color]::FromArgb(255,180,40,40)
    $btnDel.Add_Click({
        $this.Parent.Parent.Controls.Remove($this.Parent)
    })
    [void]$row.Controls.Add($btnDel)

    [void]$panelJqlFlow.Controls.Add($row)
    return $row
}

$script:failures = @()
function Check($name, $cond) {
    if ($cond) { Write-Host "  OK $name" } else { $script:failures += $name; Write-Host "  FAIL $name" }
}

# --- fixed-shaped row (titleEditable=$false), no {ws}/{we}/{account_id} needed for a stable assert ---
$fixedRow = New-JqlRow 'investigation' '1. Investigation issues' 'project = {project} AND creator = {account_id}' $false
[System.Windows.Forms.Clipboard]::SetText('')
$fixedRow.Controls[2].PerformClick()   # Copy
$clipFixed = [System.Windows.Forms.Clipboard]::GetText()
Check "fixed row Copy substitutes project/account" ($clipFixed -eq 'project = GT2 AND creator = ACC123')

# --- custom-shaped row (titleEditable=$true) ---
$customRow = New-JqlRow 'extra_ab12cd34' 'My query' 'project = {project} custom' $true
[System.Windows.Forms.Clipboard]::SetText('')
$customRow.Controls[2].PerformClick()  # Copy
$clipCustom = [System.Windows.Forms.Clipboard]::GetText()
Check "custom row Copy substitutes project" ($clipCustom -eq 'project = GT2 custom')

# --- Delete removes the row from the panel, for both shapes ---
$countBeforeFixed = $panelJqlFlow.Controls.Count
$fixedRow.Controls[3].PerformClick()   # Delete
Check "fixed row Delete removes it from the panel" (
    $panelJqlFlow.Controls.Count -eq ($countBeforeFixed - 1) -and -not $panelJqlFlow.Controls.Contains($fixedRow)
)

$countBeforeCustom = $panelJqlFlow.Controls.Count
$customRow.Controls[3].PerformClick()  # Delete
Check "custom row Delete removes it from the panel" (
    $panelJqlFlow.Controls.Count -eq ($countBeforeCustom - 1) -and -not $panelJqlFlow.Controls.Contains($customRow)
)

if ($script:failures.Count -gt 0) {
    Write-Host "FAILURES: $($script:failures -join ', ')"
    exit 1
} else {
    Write-Host "ALL PASS"
    exit 0
}
```

This script defines a standalone copy of `New-JqlRow` rather than dot-sourcing `jira-sync.ps1` — the real file builds and shows the main window and blocks on `[System.Windows.Forms.Application]::Run($form)` at the bottom, so sourcing it would open a visible window and hang instead of returning control to the test. Whenever `New-JqlRow`'s body changes in `jira-sync.ps1` (Step 1), update this copy to match, byte-for-byte in the part being tested (the two `Add_Click` bodies and the control-tree shape) — a drift here would let this test pass while the real function silently regresses, defeating its purpose.

- [ ] **Step 5: Run the headless verification script**

Run: `powershell -NoProfile -File scripts/test_jql_row_gui.ps1`
Expected: five `OK` lines, then `ALL PASS`, exit code `0`.

If any line prints `FAIL`, stop and fix `New-JqlRow` (or this test, if the test itself has a mistake) before proceeding — do not move on with a failing headless check, per this area's history of appearing correct on visual/diff review while actually being a no-op.

- [ ] **Step 6: Manual verification (defer to human — no live GUI in this environment)**

Note in your report that this step is pending: launch `jira-sync.ps1`, open Settings → Jira queries.
- All six built-in rows and any existing custom rows render with identical panel size, JQL box size, and font — no visual difference between a built-in and a custom row except Delete's *effect* (both have the button; both work).
- "+ Add query" sits below the row list, not above it.
- Click "+ Add query": a new blank editable row appears, the view auto-scrolls to reveal it, and its title field has keyboard focus immediately (no need to click into it first).
- Click Delete on one of the six built-in rows, then Save & Close, then reopen Settings → Jira queries: that row is gone, the other five built-in rows and any custom rows are unaffected.
- Open `config/jql_queries.json` and confirm it now has a `hidden_builtin` array containing that row's key.
- Re-add nothing (no restore button expected — see Non-goals); confirm running `Sync Jira -> Excel` still produces that query's data using the default JQL (per spec, deleting the row does not disable the query — this is intentional, not a bug to chase).

- [ ] **Step 7: Commit**

```bash
git add jira-sync.ps1 scripts/test_jql_row_gui.ps1
git commit -m "refactor: unify built-in/custom JQL rows into New-JqlRow, add hidden_builtin delete, move Add-query below list"
```

---

## Notes for the implementer

- Task 1 is fully automatable and testable without touching the GUI or any Excel file — do not skip its tests waiting for Task 2/3.
- Tasks 2 and 3 are WinForms layout/behavior changes. The parser check (Step "Verify parse" in each) only catches syntax errors, not logic errors — this is exactly the gap that let the `d6a6fa9` bug ship. Task 3's headless `PerformClick()` script is not optional polish; per the design's Testing section, no task touching `New-JqlRow` is done until it passes.
- Manual/visual verification steps in Tasks 2 and 3 are explicitly deferred to a human — report them as pending in your task summary, don't attempt to launch or screenshot the live GUI yourself.
- Keep the Copy/Delete handler bodies (in `New-JqlRow` and its test-script twin) using `$this.Parent...` control-tree navigation — never rewrite them to close over a builder function's own locals, even if it looks "safe" for a specific case. See the Global Constraints closure-safety rule.
