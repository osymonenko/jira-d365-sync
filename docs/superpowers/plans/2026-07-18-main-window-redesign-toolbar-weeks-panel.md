# Main Window Redesign — Two-Row Toolbar + Right-Side Weeks Panel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Split the main window's cramped single-row toolbar into a deliberate two-row icon toolbar, and move the Weeks/Month controls, the "file loaded" summary, and the Copy Log button into a fixed-width sidebar on the right that spans the full remaining window height — with a tri-state "Select all" checkbox replacing the old "All"/"None" text links.

**Architecture:** Edit `jira-sync.ps1` only, no Python changes. Three independent regions of the file, done in dependency order: (1) the toolbar rows (`$pnlTop`/`$pnlBtns`) — pure relabel/reorder, zero coordinate impact on anything below; (2) the weeks/status/log region — replaces the dynamic-height horizontal weeks strip with a fixed-width, `Dock`-based vertical sidebar (log narrows to make room, same resize behavior as today just applied to two columns instead of one); (3) the one genuinely new piece of logic — a tri-state master checkbox — built on top of (2)'s sidebar, verified with a headless `PerformClick()` script per this codebase's own precedent (the JQL row Copy/Delete bug, commit `d6a6fa9`, was only caught this way).

**Tech Stack:** PowerShell WinForms (`jira-sync.ps1`).

**Spec:** `docs/superpowers/specs/2026-07-18-main-window-redesign-toolbar-weeks-panel-design.md`

## Global Constraints

- Zero change to what any button *does* — every `Add_Click` body, every Python command invocation, and `$script:weekCheckboxes`/`$script:monthCodes` semantics are untouched. This is layout + one small new piece of UI-state logic, not a behavior change.
- Existing button variable names (`$btnTest`, `$btnRead`, `$btnOpen`, `$btnSync`, `$btnSubmit`, `$btnFill`, `$btnFillStd`, `$btnStop`, `$btnSettings`) are never renamed — their `Add_Click` handlers are defined elsewhere in the file and bind by variable.
- Window width stays 1060px; `MinimumSize`/`MaximumSize` are unchanged.
- Icon/order mapping is fixed, confirmed with the user, not open to reinterpretation: Row 1 = connection-status + `❓ Test` + `⚙️ Settings` + `👀 Read File` + `✏️ Open File`. Row 2 = `Standard ⬇️` + `Jira ⬇️` + `Submit tasks ➡️` + `Fill days ⬆️⬆️⬆️⬆️⬆️` + `⛔` (Stop).
- The status bar (`$lblStatusDot`/`$lblStatus`, driven by `Set-Status`) stays a full-width strip above the log — it does NOT move into the sidebar. Only `$btnCopyLog` moves out of it.
- `jira-sync.ps1` is already saved as UTF-8 with a BOM (confirmed) — Windows PowerShell 5.1 needs the BOM to read the literal emoji correctly; do not re-save the file in a way that strips it (standard Edit-tool usage on an already-open file preserves it; a full-file rewrite through a non-UTF-8-aware tool would not).
- A ThreeState `CheckBox`'s `Click` event fires AFTER WinForms' own built-in `CheckState` auto-cycle (`Unchecked→Checked→Indeterminate→Unchecked`) has already been applied — any handler on `Click` reads the already-cycled state, it does not need to compute the "previous" state itself.
- Master-checkbox toggle-all logic is wired to `Add_Click`, never `Add_CheckedChanged`/`Add_CheckStateChanged` — programmatic `.CheckState` assignment (used by the recompute function) never fires `Click`, so there is no reentrancy risk and no manual guard flag is needed. This is a deliberate, more robust choice than the design doc's suggested "guard flag" — document this in the task-3 report since a reviewer comparing literally against the design doc's wording could otherwise flag a "missing guard" that isn't actually needed.

## File Structure

- `jira-sync.ps1` — toolbar rows rebuilt (Task 1); weeks/status/log region rebuilt into log+sidebar split (Task 2); tri-state master-checkbox logic added on top (Task 3).
- `scripts/test_weeks_master_checkbox_gui.ps1` — new headless PowerShell script proving the tri-state master-checkbox logic via `PerformClick()`/direct `CheckState` assertions, no visible window needed (Task 3).

---

### Task 1: Two-row icon toolbar (pure relabel/reorder)

**Files:**
- Modify: `jira-sync.ps1:78-137` (the block from `# ---- Top bar: Settings + connection dot + Excel picker` through `$btnStop.Enabled = $false`)

**Interfaces:**
- Consumes: `New-Btn($parent,$text,$x,$y,$w,$h,$r,$g,$b)` (existing helper, unchanged), `$form`.
- Produces: the same button/label variables (`$lblDot`, `$lblConnStatus`, `$txtFile`, `$btnTest`, `$btnSettings`, `$btnRead`, `$btnOpen`, `$btnFillStd`, `$btnSync`, `$btnSubmit`, `$btnFill`, `$btnStop`), redistributed across `$pnlTop` (row 1) and `$pnlBtns` (row 2), same total footprint (`$pnlTop` still 1060×46 at (0,0), `$pnlBtns` still 1060×52 at (0,46)) — so nothing below the toolbar needs to move in this task.

- [ ] **Step 1: Replace the toolbar block**

Find the block starting at `# ---- Top bar: Settings + connection dot + Excel picker --------------------` (jira-sync.ps1:78) and ending at `$btnStop.Enabled = $false` (jira-sync.ps1:137). Replace the ENTIRE block with:

```powershell
# ---- Toolbar row 1: connection status + setup actions ---------------------
$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Location = New-Object System.Drawing.Point(0,0)
$pnlTop.Size = New-Object System.Drawing.Size(1060,46)
$pnlTop.BackColor = [System.Drawing.Color]::FromArgb(255,50,50,60)
[void]$form.Controls.Add($pnlTop)

$lblDot = New-Object System.Windows.Forms.Label
$lblDot.Text = 'l'; $lblDot.ForeColor = [System.Drawing.Color]::Gray
$lblDot.Font = New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold)
$lblDot.Location = New-Object System.Drawing.Point(8,6)
$lblDot.Size = New-Object System.Drawing.Size(22,30)
$lblDot.TextAlign = 'MiddleCenter'
[void]$pnlTop.Controls.Add($lblDot)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = 'not tested'
$lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,180,180,180)
$lblConnStatus.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblConnStatus.Location = New-Object System.Drawing.Point(32,15)
$lblConnStatus.Size = New-Object System.Drawing.Size(110,18)
[void]$pnlTop.Controls.Add($lblConnStatus)

$btnTest     = New-Btn $pnlTop '❓ Test'      150 6 74  34 70  70  85
$btnSettings = New-Btn $pnlTop '⚙️ Settings'  230 6 100 34 70  70  85
$btnRead     = New-Btn $pnlTop '👀 Read File' 336 6 120 34 40  100 140
$btnOpen     = New-Btn $pnlTop '✏️ Open File' 462 6 120 34 40  110 60

# Excel path lives in Settings only (persisted to .env as EXCEL_FILE). We keep
# $txtFile as an off-screen holder so the rest of the UI (Read/Submit/Fill) can
# still read $txtFile.Text without threading the path through every handler.
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Text = if ($envData['EXCEL_FILE']) { $envData['EXCEL_FILE'] } else { Join-Path $scriptDir 'data\timesheet.xlsx' }

# ---- Toolbar row 2: actions that do work -----------------------------------
$pnlBtns = New-Object System.Windows.Forms.Panel
$pnlBtns.Location = New-Object System.Drawing.Point(0,46)
$pnlBtns.Size = New-Object System.Drawing.Size(1060,52)
$pnlBtns.BackColor = [System.Drawing.Color]::FromArgb(255,60,60,72)
[void]$form.Controls.Add($pnlBtns)

$btnFillStd = New-Btn $pnlBtns 'Standard ⬇️'         8   8 150 36 120 80  160
$btnSync    = New-Btn $pnlBtns 'Jira ⬇️'             166 8 150 36 0   122 200
$btnSubmit  = New-Btn $pnlBtns 'Submit tasks ➡️'      324 8 185 36 180 80  0
$btnFill    = New-Btn $pnlBtns 'Fill days ⬆️⬆️⬆️⬆️⬆️' 517 8 220 36 80  80  80
$btnFill.Enabled = $false
$btnFill.ForeColor = [System.Drawing.Color]::FromArgb(255,140,140,140)

$btnStop = New-Btn $pnlBtns '⛔' 1024 8 22 36 140 30 30
$btnStop.Enabled = $false
```

Notes on this replacement:
- `$btnTest`/`$btnRead`/`$btnOpen` move from `$pnlBtns` to `$pnlTop`; `$btnSettings` stays in `$pnlTop` (it already lived there — only its `.Text`/`.Location`/font override change, and its old `Font = New-Object ... 'Segoe UI',9` non-bold override is dropped so it matches the other row-1 buttons' bold style, a deliberate small cosmetic simplification, not an oversight).
- `$btnFillStd` (Standard) moves before `$btnSync` (Jira) in row 2's left-to-right order, per the approved design — this is an intentional reorder versus today's code (today: Sync, Submit, Fill, Standard, Stop), not an oversight.
- The old `$sep` cluster-separator label is removed entirely — row 2 is a flat list of 5 action buttons now, no cluster grouping needed since Test/Read/Open moved out to row 1.
- `$btnStop.Text = [char]9632` (the old stop-square glyph) is replaced by the literal `'⛔'` passed directly to `New-Btn`, so the old separate `$btnStop.Text = [char]9632` assignment line is dropped (superseded, would just be overwritten).

- [ ] **Step 2: Verify parse**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

- [ ] **Step 3: Manual verification (defer to human — no live GUI in this environment)**

Note in your report that this step is pending: launch `jira-sync.ps1`.
- Row 1 (dark navy, y=0-46) reads left to right: connection dot+status, `❓ Test`, `⚙️ Settings`, `👀 Read File`, `✏️ Open File` — no overlap, all emoji render as color glyphs (not boxes/mojibake — if they render as `□`/`?` boxes, the file's UTF-8 BOM was lost somewhere and needs restoring, escalate rather than guessing a fix).
- Row 2 (slightly lighter dark, y=46-98) reads: `Standard ⬇️`, `Jira ⬇️`, `Submit tasks ➡️`, `Fill days ⬆️⬆️⬆️⬆️⬆️`, then `⛔` isolated at the far right — no overlap.
- Click each button and confirm it still triggers its existing action unchanged (Read File loads weeks, Open File launches Excel, Test tests the connection, Settings opens the dialog, Sync/Submit/Fill/Standard run their existing Python/D365 commands, Stop is disabled until a process is running).
- Everything below the toolbar (weeks strip, divider, status, log) is untouched by this task and should look exactly as it did before.

- [ ] **Step 4: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: split main-window toolbar into two icon rows (setup / actions)"
```

---

### Task 2: Right-side Weeks sidebar (layout only, no new behavior)

**Files:**
- Modify: `jira-sync.ps1:139-312` (the block from `# ---- Weeks strip` through `[void]$form.Controls.Add($txtLog)`)
- Modify: `jira-sync.ps1:314-317` (`$form.Add_Resize`)
- Modify: `jira-sync.ps1` — the `$btnRead.Add_Click` success handler (currently around line 855-858, locate by content, this file's line numbers have shifted since Task 1)

**Interfaces:**
- Consumes: `$form`, `$txtFile`, `Set-Status` (unchanged).
- Produces: `$pnlWeeksSidebar` (the sidebar container), `$pnlWeeksList` (the scrollable vertical week-checkbox host, replaces the old horizontal grid inside `$pnlWeeks`), `$chkWeeksAll` (the master checkbox control — created here, but its `Add_Click`/tri-state behavior is Task 3's job; this task only creates it, sets `ThreeState=$true`, and starts it `Visible=$false`), `$lblWeeksSummary` (persistent "file loaded" label), `$btnCopyLog` (moved here, `Add_Click` body byte-for-byte unchanged), rewritten `Populate-Weeks($weeksJson)` (same signature, same `$script:weekCheckboxes`/`$script:monthCodes` contract, vertical layout instead of column-of-5 grid, no more dynamic panel-height/divider/status/log repositioning). `$divider`, `$pnlStatus`, `$lblStatusDot`, `$lblStatus`, `$txtLog` all keep their existing names, `Set-Status` is untouched.

- [ ] **Step 1: Replace the weeks/divider/status/log block**

Find the block starting at `# ---- Weeks strip (compact, checkbox-only) ----------------------------------` (jira-sync.ps1:139) and ending at `[void]$form.Controls.Add($txtLog)` (jira-sync.ps1:312) — this spans the old `$pnlWeeks` setup, `Populate-Weeks`, the old `$lnkSelectAll`/`$lnkNone` click handlers, `$divider`, `$pnlStatus` (including the old `$btnCopyLog`), and `$txtLog`. Replace the ENTIRE block with:

```powershell
# ---- Right-side Weeks sidebar (fixed width, full remaining height) ---------
$pnlWeeksSidebar = New-Object System.Windows.Forms.Panel
$pnlWeeksSidebar.Location = New-Object System.Drawing.Point(800,128)
$pnlWeeksSidebar.Size = New-Object System.Drawing.Size(260,524)
$pnlWeeksSidebar.BackColor = [System.Drawing.Color]::White
[void]$form.Controls.Add($pnlWeeksSidebar)

$pnlWeeksList = New-Object System.Windows.Forms.Panel
$pnlWeeksList.Dock = 'Fill'
$pnlWeeksList.AutoScroll = $true
$pnlWeeksList.BackColor = [System.Drawing.Color]::White

$pnlWeeksTop = New-Object System.Windows.Forms.Panel
$pnlWeeksTop.Dock = 'Top'; $pnlWeeksTop.Height = 64
$pnlWeeksTop.BackColor = [System.Drawing.Color]::White

$lblMonth = New-Object System.Windows.Forms.Label
$lblMonth.Text = 'Month:'
$lblMonth.Location = New-Object System.Drawing.Point(6,8)
$lblMonth.Size = New-Object System.Drawing.Size(44,18)
$lblMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblMonth.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
[void]$pnlWeeksTop.Controls.Add($lblMonth)

$cmbMonth = New-Object System.Windows.Forms.ComboBox
$cmbMonth.Location = New-Object System.Drawing.Point(54,6)
$cmbMonth.Size = New-Object System.Drawing.Size(190,22)
$cmbMonth.DropDownStyle = 'DropDownList'
$cmbMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
[void]$cmbMonth.Items.Add('All')
$cmbMonth.SelectedIndex = 0
[void]$pnlWeeksTop.Controls.Add($cmbMonth)

$chkWeeksAll = New-Object System.Windows.Forms.CheckBox
$chkWeeksAll.Text = 'Select all'
$chkWeeksAll.ThreeState = $true
$chkWeeksAll.Location = New-Object System.Drawing.Point(6,34)
$chkWeeksAll.Size = New-Object System.Drawing.Size(244,20)
$chkWeeksAll.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
$chkWeeksAll.Visible = $false
[void]$pnlWeeksTop.Controls.Add($chkWeeksAll)

$pnlWeeksBottom = New-Object System.Windows.Forms.Panel
$pnlWeeksBottom.Dock = 'Bottom'; $pnlWeeksBottom.Height = 64
$pnlWeeksBottom.BackColor = [System.Drawing.Color]::White

$lblWeeksSummary = New-Object System.Windows.Forms.Label
$lblWeeksSummary.Text = ''
$lblWeeksSummary.Location = New-Object System.Drawing.Point(6,6)
$lblWeeksSummary.Size = New-Object System.Drawing.Size(244,20)
$lblWeeksSummary.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$lblWeeksSummary.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
[void]$pnlWeeksBottom.Controls.Add($lblWeeksSummary)

$btnCopyLog = New-Object System.Windows.Forms.Button
$btnCopyLog.Text = 'Copy log'
$btnCopyLog.Location = New-Object System.Drawing.Point(6,32)
$btnCopyLog.Size = New-Object System.Drawing.Size(244,24)
$btnCopyLog.FlatStyle = 'Flat'
$btnCopyLog.FlatAppearance.BorderSize = 1
$btnCopyLog.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255,200,200,210)
$btnCopyLog.BackColor = [System.Drawing.Color]::FromArgb(255,245,245,250)
$btnCopyLog.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
$btnCopyLog.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$btnCopyLog.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCopyLog.Add_Click({
    if ([string]::IsNullOrEmpty($txtLog.Text)) { return }
    [System.Windows.Forms.Clipboard]::SetText($txtLog.Text)
    $prevText = $btnCopyLog.Text
    $btnCopyLog.Text = 'Copied!'
    $btnCopyLog.ForeColor = [System.Drawing.Color]::DarkGreen
    $script:copyLogTimer = New-Object System.Windows.Forms.Timer
    $script:copyLogTimer.Interval = 1500
    $script:copyLogTimer.Add_Tick({
        $btnCopyLog.Text = $prevText
        $btnCopyLog.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
        $script:copyLogTimer.Stop(); $script:copyLogTimer.Dispose()
    })
    $script:copyLogTimer.Start()
})
[void]$pnlWeeksBottom.Controls.Add($btnCopyLog)

# Add order matters: Fill panel first, then Top/Bottom strips after — each
# claims its own edge without covering the Fill panel (same pattern already
# proven safe for the Jira-queries tab's FlowLayoutPanel + add-strip).
[void]$pnlWeeksSidebar.Controls.Add($pnlWeeksList)
[void]$pnlWeeksSidebar.Controls.Add($pnlWeeksTop)
[void]$pnlWeeksSidebar.Controls.Add($pnlWeeksBottom)

$script:monthCodes = @{}
$script:weekCheckboxes = @()

function Populate-Weeks($weeksJson) {
    $toRemove = @($pnlWeeksList.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] })
    foreach ($c in $toRemove) { $pnlWeeksList.Controls.Remove($c) }
    $script:weekCheckboxes = @()
    $chkWeeksAll.Visible = $false
    if (-not $weeksJson -or $weeksJson.Count -eq 0) { return }

    # Rebuild month dropdown from the months each week touches (start + end).
    $cmbMonth.Items.Clear()
    [void]$cmbMonth.Items.Add('All')
    $script:monthCodes = @{}
    $codes = New-Object 'System.Collections.Generic.SortedSet[string]'
    foreach ($w in $weeksJson) {
        [void]$codes.Add($w.start.Substring(0,7))
        [void]$codes.Add($w.end.Substring(0,7))
    }
    foreach ($code in $codes) {
        $display = [datetime]::ParseExact("$code-01",'yyyy-MM-dd',$null).ToString('MMMM yyyy',[System.Globalization.CultureInfo]::InvariantCulture)
        $script:monthCodes[$display] = $code
        [void]$cmbMonth.Items.Add($display)
    }
    $cmbMonth.SelectedIndex = 0

    $i = 0
    foreach ($w in $weeksJson) {
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = $w.start.Substring(5) + ' – ' + $w.end.Substring(5)
        $chk.Location = New-Object System.Drawing.Point(6, (4 + $i * 24))
        $chk.Size = New-Object System.Drawing.Size(230, 20)
        $chk.Checked = $true
        $chk.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
        $chk.Tag = $w.start
        [void]$pnlWeeksList.Controls.Add($chk)
        $script:weekCheckboxes += $chk
        $i++
    }
    $chkWeeksAll.Visible = $true
}

# ---- Divider ---------------------------------------------------------------
$divider = New-Object System.Windows.Forms.Label
$divider.Location = New-Object System.Drawing.Point(0,98)
$divider.Size = New-Object System.Drawing.Size(1060,1)
$divider.BackColor = [System.Drawing.Color]::FromArgb(255,200,200,210)
[void]$form.Controls.Add($divider)

# ---- Status bar ------------------------------------------------------------
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Location = New-Object System.Drawing.Point(0,99)
$pnlStatus.Size = New-Object System.Drawing.Size(1060,28)
$pnlStatus.BackColor = [System.Drawing.Color]::FromArgb(255,235,235,240)
[void]$form.Controls.Add($pnlStatus)

$lblStatusDot = New-Object System.Windows.Forms.Label
$lblStatusDot.Text = 'l'
$lblStatusDot.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
$lblStatusDot.Location = New-Object System.Drawing.Point(8,-4)
$lblStatusDot.Size = New-Object System.Drawing.Size(20,30)
$lblStatusDot.ForeColor = [System.Drawing.Color]::LimeGreen
[void]$pnlStatus.Controls.Add($lblStatusDot)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(30,6)
$lblStatus.Size = New-Object System.Drawing.Size(1020,18)
$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI',9)
[void]$pnlStatus.Controls.Add($lblStatus)

# ---- Log -------------------------------------------------------------------
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(0,128)
$txtLog.Size = New-Object System.Drawing.Size(800,524)
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(255,20,20,26)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(255,200,200,200)
$txtLog.Font = New-Object System.Drawing.Font('Consolas',9)
$txtLog.BorderStyle = 'None'
$txtLog.ScrollBars = 'Vertical'
[void]$form.Controls.Add($txtLog)
```

Notes on this replacement:
- `$pnlWeeks`, `$lblWeeksTitle`, `$lnkSelectAll`, `$lnkNone` are gone entirely — the master checkbox (created here, wired in Task 3) fully replaces the old text links' function.
- `Populate-Weeks` no longer computes `$panelH`/pushes `$divider.Top`/`$pnlStatus.Top`/`$txtLog.Top` around — the sidebar's `Dock='Fill'` + `AutoScroll` on `$pnlWeeksList` absorbs any number of weeks by scrolling internally, so nothing else needs to move when the week count changes.
- `$divider` moves from y=138 to y=98 (right after the 98px-tall toolbar, since the old 40px `$pnlWeeks` band is gone); `$pnlStatus` moves from y=139 to y=99; `$txtLog` moves from y=167/width=1060 to y=128/width=800 (narrower, sharing the row with the sidebar).
- `$lblStatus.Size` widens from `(710,18)` to `(1020,18)` since it no longer needs to leave room for `$btnCopyLog` in the same panel.
- `$pnlWeeksSidebar` sits at x=800 (right after `$txtLog`'s new 800px width), width 260, so `800+260=1060` — exactly fills the remaining width, no gap or overlap.

- [ ] **Step 2: Update the resize handler**

Find (jira-sync.ps1:314-317):

```powershell
$form.Add_Resize({
    $h = $form.ClientSize.Height
    $txtLog.Height = [Math]::Max($h - $txtLog.Top, 50)
})
```

Replace with:

```powershell
$form.Add_Resize({
    $h = $form.ClientSize.Height
    $newHeight = [Math]::Max($h - $txtLog.Top, 50)
    $txtLog.Height = $newHeight
    $pnlWeeksSidebar.Height = $newHeight
})
```

- [ ] **Step 3: Update the "file loaded" summary in `$btnRead`'s success handler**

Find (inside `$btnRead.Add_Click`, locate by content — line numbers have shifted):

```powershell
            $weeks = $out | ConvertFrom-Json
            Populate-Weeks $weeks
            Set-Status ("File loaded: " + $weeks.Count + " week(s) found") ([System.Drawing.Color]::LimeGreen)
            Append-Log ('[INFO] Loaded ' + $weeks.Count + ' week(s) from: ' + $txtFile.Text)
```

Replace with:

```powershell
            $weeks = $out | ConvertFrom-Json
            Populate-Weeks $weeks
            Set-Status ("File loaded: " + $weeks.Count + " week(s) found") ([System.Drawing.Color]::LimeGreen)
            $lblWeeksSummary.Text = "File loaded: " + $weeks.Count + " week(s) found"
            Append-Log ('[INFO] Loaded ' + $weeks.Count + ' week(s) from: ' + $txtFile.Text)
```

This adds the persistent sidebar summary alongside (not instead of) the existing status-bar text and log line — all three keep showing the same information in their respective places.

- [ ] **Step 4: Verify parse**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

- [ ] **Step 5: Manual verification (defer to human — no live GUI in this environment)**

Note in your report that this step is pending: launch `jira-sync.ps1`, click Read File on a real timesheet.
- The right sidebar spans the full height from just below the status bar to the bottom of the window, at a fixed ~260px width; the log to its left is correspondingly narrower.
- Top of the sidebar: `Month:` + dropdown, then a "Select all" checkbox (unchecked/disabled-looking is fine before a file is read — it should be invisible until then, matching the old links' behavior).
- After Read File: week checkboxes appear as a vertical list (not the old horizontal grid); if there are enough weeks to overflow the visible area, the list scrolls internally without moving anything else.
- Bottom of the sidebar: "File loaded: N week(s) found" text, then the Copy Log button below it — Copy Log still copies the full log text (same behavior as before, just relocated).
- Resize the window taller/shorter: both the log and the sidebar grow/shrink together, staying full height, exactly like the log alone did before.

- [ ] **Step 6: Commit**

```bash
git add jira-sync.ps1
git commit -m "refactor: move Weeks/Month/summary/Copy-log into a fixed right sidebar"
```

---

### Task 3: Tri-state "Select all" master checkbox

**Files:**
- Modify: `jira-sync.ps1` (the `Populate-Weeks` function and its call site from Task 2; add `Update-MasterCheckboxState` and `$chkWeeksAll.Add_Click` nearby)
- Test: `scripts/test_weeks_master_checkbox_gui.ps1` (new)

**Interfaces:**
- Consumes: `$chkWeeksAll`, `$pnlWeeksList`, `$script:weekCheckboxes` (all from Task 2, unchanged contract).
- Produces: `Update-MasterCheckboxState` (recomputes `$chkWeeksAll.CheckState` from the current `$script:weekCheckboxes` — `Checked` if all are checked, `Unchecked` if none are, `Indeterminate` otherwise); `$chkWeeksAll.Add_Click` (toggle-all: selects all unless the master was fully `Checked` before this click, in which case it deselects all — see Global Constraints for why `Indeterminate` is never the *result* of a real click, only of the recompute).

- [ ] **Step 1: Add `Update-MasterCheckboxState` and the master's `Add_Click`, and wire the early-return branch**

Find (as landed by Task 2 — the start of `Populate-Weeks`):

```powershell
function Populate-Weeks($weeksJson) {
    $toRemove = @($pnlWeeksList.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] })
    foreach ($c in $toRemove) { $pnlWeeksList.Controls.Remove($c) }
    $script:weekCheckboxes = @()
    $chkWeeksAll.Visible = $false
    if (-not $weeksJson -or $weeksJson.Count -eq 0) { return }
```

Replace with:

```powershell
function Update-MasterCheckboxState {
    if ($script:weekCheckboxes.Count -eq 0) { $chkWeeksAll.CheckState = 'Unchecked'; return }
    $checkedCount = @($script:weekCheckboxes | Where-Object { $_.Checked }).Count
    if ($checkedCount -eq 0) { $chkWeeksAll.CheckState = 'Unchecked' }
    elseif ($checkedCount -eq $script:weekCheckboxes.Count) { $chkWeeksAll.CheckState = 'Checked' }
    else { $chkWeeksAll.CheckState = 'Indeterminate' }
}

$chkWeeksAll.Add_Click({
    # WinForms auto-cycles a ThreeState CheckBox's CheckState on click, in order
    # Unchecked -> Checked -> Indeterminate -> Unchecked, BEFORE this handler runs.
    # So by the time we read $chkWeeksAll.CheckState here, it already reflects the
    # post-click auto-cycled value:
    #   prior Unchecked    -> now Checked        -> already means "select all", keep it
    #   prior Checked      -> now Indeterminate  -> override to mean "deselect all"
    #   prior Indeterminate -> now Unchecked     -> override to mean "select all"
    #                          (a real click never leaves the box on Indeterminate;
    #                          only Update-MasterCheckboxState's recompute does that)
    if ($chkWeeksAll.CheckState -eq 'Indeterminate') {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $false }
        $chkWeeksAll.CheckState = 'Unchecked'
    } else {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $true }
        $chkWeeksAll.CheckState = 'Checked'
    }
})

function Populate-Weeks($weeksJson) {
    $toRemove = @($pnlWeeksList.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] })
    foreach ($c in $toRemove) { $pnlWeeksList.Controls.Remove($c) }
    $script:weekCheckboxes = @()
    $chkWeeksAll.Visible = $false
    if (-not $weeksJson -or $weeksJson.Count -eq 0) {
        $chkWeeksAll.CheckState = 'Unchecked'
        return
    }
```

- [ ] **Step 2: Wire each week checkbox's `CheckedChanged` to the recompute, and recompute once after populating**

Find (as landed by Task 2 — inside `Populate-Weeks`'s loop and its tail):

```powershell
        $chk.Tag = $w.start
        [void]$pnlWeeksList.Controls.Add($chk)
        $script:weekCheckboxes += $chk
        $i++
    }
    $chkWeeksAll.Visible = $true
}
```

Replace with:

```powershell
        $chk.Tag = $w.start
        $chk.Add_CheckedChanged({ Update-MasterCheckboxState })
        [void]$pnlWeeksList.Controls.Add($chk)
        $script:weekCheckboxes += $chk
        $i++
    }
    $chkWeeksAll.Visible = $true
    Update-MasterCheckboxState
}
```

The handler `{ Update-MasterCheckboxState }` doesn't reference `$chk`/`$w`/any loop-scoped local at all — it just calls a script-scope function that re-reads the current `$script:weekCheckboxes` array fresh each time. This sidesteps the classic "closure captures the wrong loop iteration's variable" bug entirely, and the "helper-function's-own-locals-are-gone-by-click-time" bug from this codebase's history (commit `d6a6fa9`) doesn't apply here either: `Update-MasterCheckboxState` is a script-scope (top-level) function, not a locals of some builder function that already returned — it stays resolvable for the entire lifetime of the running app (`Application.Run` keeps the process alive), same as `Set-Status`/`Append-Log` already are.

- [ ] **Step 3: Verify parse**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

- [ ] **Step 4: Write the headless tri-state verification script**

Create `scripts/test_weeks_master_checkbox_gui.ps1`:

```powershell
Add-Type -AssemblyName System.Windows.Forms

$script:weekCheckboxes = @()

$chkWeeksAll = New-Object System.Windows.Forms.CheckBox
$chkWeeksAll.ThreeState = $true

function Update-MasterCheckboxState {
    if ($script:weekCheckboxes.Count -eq 0) { $chkWeeksAll.CheckState = 'Unchecked'; return }
    $checkedCount = @($script:weekCheckboxes | Where-Object { $_.Checked }).Count
    if ($checkedCount -eq 0) { $chkWeeksAll.CheckState = 'Unchecked' }
    elseif ($checkedCount -eq $script:weekCheckboxes.Count) { $chkWeeksAll.CheckState = 'Checked' }
    else { $chkWeeksAll.CheckState = 'Indeterminate' }
}

$chkWeeksAll.Add_Click({
    if ($chkWeeksAll.CheckState -eq 'Indeterminate') {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $false }
        $chkWeeksAll.CheckState = 'Unchecked'
    } else {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $true }
        $chkWeeksAll.CheckState = 'Checked'
    }
})

$week1 = New-Object System.Windows.Forms.CheckBox; $week1.Checked = $true
$week2 = New-Object System.Windows.Forms.CheckBox; $week2.Checked = $true
$week3 = New-Object System.Windows.Forms.CheckBox; $week3.Checked = $true
foreach ($w in @($week1,$week2,$week3)) { $w.Add_CheckedChanged({ Update-MasterCheckboxState }) }
$script:weekCheckboxes = @($week1,$week2,$week3)
Update-MasterCheckboxState

$script:failures = @()
function Check($name, $cond) {
    if ($cond) { Write-Host "  OK $name" } else { $script:failures += $name; Write-Host "  FAIL $name" }
}

Check "all three checked -> master Checked" ($chkWeeksAll.CheckState -eq 'Checked')

$week2.Checked = $false
Check "one unchecked -> master Indeterminate" ($chkWeeksAll.CheckState -eq 'Indeterminate')

$week1.Checked = $false
$week3.Checked = $false
Check "all unchecked -> master Unchecked" ($chkWeeksAll.CheckState -eq 'Unchecked')

$chkWeeksAll.PerformClick()   # was Unchecked -> auto-cycles to Checked -> handler keeps it, selects all
Check "click while Unchecked selects all" (
    $week1.Checked -and $week2.Checked -and $week3.Checked -and $chkWeeksAll.CheckState -eq 'Checked'
)

$chkWeeksAll.PerformClick()   # was Checked -> auto-cycles to Indeterminate -> handler deselects all
Check "click while Checked deselects all" (
    (-not $week1.Checked) -and (-not $week2.Checked) -and (-not $week3.Checked) -and $chkWeeksAll.CheckState -eq 'Unchecked'
)

$week1.Checked = $true
$week2.Checked = $false
$week3.Checked = $true
Check "mixed selection -> master Indeterminate (again)" ($chkWeeksAll.CheckState -eq 'Indeterminate')

$chkWeeksAll.PerformClick()   # was Indeterminate -> auto-cycles to Unchecked -> handler selects all
Check "click while Indeterminate selects all (never gets stuck)" (
    $week1.Checked -and $week2.Checked -and $week3.Checked -and $chkWeeksAll.CheckState -eq 'Checked'
)

$week1.Checked = $false
$beforeWeek2 = $week2.Checked
$beforeWeek3 = $week3.Checked
$chkWeeksAll.CheckState = 'Indeterminate'   # programmatic set, NOT a click
Check "programmatic CheckState set does not trigger toggle-all (no Click fired)" (
    $week1.Checked -eq $false -and $week2.Checked -eq $beforeWeek2 -and $week3.Checked -eq $beforeWeek3
)

if ($script:failures.Count -gt 0) {
    Write-Host "FAILURES: $($script:failures -join ', ')"
    exit 1
} else {
    Write-Host "ALL PASS"
    exit 0
}
```

This script defines its own standalone copy of `Update-MasterCheckboxState`/`$chkWeeksAll.Add_Click` rather than dot-sourcing `jira-sync.ps1` — the real file builds and shows the main window and blocks on `[System.Windows.Forms.Application]::Run($form)` at the bottom, so sourcing it would open a visible window and hang instead of returning control to the test (same reasoning already established for `scripts/test_jql_row_gui.ps1`). Whenever this logic changes in `jira-sync.ps1` (Step 1/2 above), update this copy to match, byte-for-byte in the part being tested — a drift here would let this test pass while the real function silently regresses.

- [ ] **Step 5: Run the headless verification script**

Run: `powershell -NoProfile -File scripts/test_weeks_master_checkbox_gui.ps1`
Expected: eight `OK` lines, then `ALL PASS`, exit code `0`.

If any line prints `FAIL`, stop and fix the tri-state logic (or this test, if the test itself has a mistake) before proceeding — do not move on with a failing headless check, per this area's history of appearing correct on visual/diff review while actually being wrong.

- [ ] **Step 6: Manual verification (defer to human — no live GUI in this environment)**

Note in your report that this step is pending: launch `jira-sync.ps1`, click Read File on a real timesheet with multiple weeks.
- The "Select all" checkbox starts fully checked (matches today's default of all weeks pre-checked).
- Uncheck one week checkbox: the master visibly shows the indeterminate (dashed/grayed) box state, not fully checked or fully unchecked.
- Click the master while indeterminate: all weeks become checked, master becomes fully checked.
- Click the master again while fully checked: all weeks become unchecked, master becomes fully unchecked.
- Sync/Submit/Fill/Standard still only act on the weeks whose individual checkboxes are checked, exactly as before (this task doesn't change that logic, only how weeks get checked/unchecked).

- [ ] **Step 7: Commit**

```bash
git add jira-sync.ps1 scripts/test_weeks_master_checkbox_gui.ps1
git commit -m "feat: tri-state Select-all checkbox for the Weeks sidebar"
```

---

## Notes for the implementer

- Task 1 has zero coordinate impact on anything below the toolbar — if you find yourself needing to touch `$pnlWeeks`/`$divider`/`$pnlStatus`/`$txtLog` while doing Task 1, stop: that means the toolbar's total height changed, which the design does not call for (both rows keep their existing 46px/52px heights).
- Task 2 is layout-only — no new behavior, so (like every prior WinForms layout task in this project) the parser check catches syntax but not visual overlap; defer to the human eyeball step, don't attempt to launch the GUI yourself.
- Task 3 is the one task with genuinely new logic, and per this codebase's own history (the JQL row Copy/Delete bug, commit `d6a6fa9`) that class of bug is invisible to both parser checks and careful diff reading — no task touching the tri-state logic is done until `scripts/test_weeks_master_checkbox_gui.ps1` passes.
- All three tasks are independent enough in *which lines they touch* to review separately, but they must land in order 1→2→3: Task 3 edits code that Task 2 creates, and both edit regions Task 1 leaves untouched.
