# GUI Redesign Phase 2 — Main Window UX

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the main window read clearly — group the action buttons into labelled clusters (Excel / Jira / D365) with consistent per-cluster colors and separators, and tidy the Weeks/Month strip and the status/log area. No behavior change.

**Architecture:** Edit `jira-sync.ps1` only. Rewrite the button-strip creation block into ordered clusters keeping the existing button variable names (so all `Add_Click` handlers stay bound); adjust spacing on the Weeks strip and status/log. No panel-height change, so the dynamic layout offsets in `Populate-Weeks` stay valid.

**Tech Stack:** PowerShell WinForms.

**Spec:** `docs/superpowers/specs/2026-07-18-gui-redesign-editable-config-design.md`

## Global Constraints

- Platform is WinForms — no web/HTML.
- Do NOT rename `$btnTest`, `$btnRead`, `$btnOpen`, `$btnSync`, `$btnSubmit`, `$btnFill`, `$btnFillStd`, `$btnStop` — their `Add_Click` handlers are defined elsewhere and bind by variable.
- Do NOT change what any button does; this is layout/labels/color only.
- Keep the panel/window width at 1060 and the button-strip height (52) unchanged so `Populate-Weeks`' offset math (base 98) stays correct.
- Cluster colors (R,G,B): Excel = `40,140,90` (green), Jira = `0,122,200` (blue), D365 = `200,110,0` (orange), Stop = `140,30,30` (red).

---

### Task 1: Cluster and recolor the button strip

**Files:**
- Modify: `jira-sync.ps1` (the button-strip creation block, currently the lines creating `$btnTest`…`$btnStop` plus the `$sep` label)

**Interfaces:**
- Consumes: `New-Btn($parent,$text,$x,$y,$w,$h,$r,$g,$b)`, `$pnlBtns`.
- Produces: the same eight button variables, re-ordered/re-colored, plus three cluster separators `$sep1`/`$sep2`/`$sep3`.

- [ ] **Step 1: Replace the button-strip block**

Find the block that starts at `$btnTest   = New-Btn $pnlBtns 'Test' ...` and ends at the `$btnStop.Enabled = $false` line (it includes the old `$sep` separator label). Replace the ENTIRE block with:

```powershell
# ---- Cluster: Excel (read/open workbook, fill standard rows) ----
$btnRead    = New-Btn $pnlBtns 'Read File'         8   8 100 36 40 140 90
$btnOpen    = New-Btn $pnlBtns 'Open File'         112 8 100 36 40 140 90
$btnFillStd = New-Btn $pnlBtns 'Standard -> Excel' 216 8 140 36 40 140 90

function New-Sep($x) {
    $s = New-Object System.Windows.Forms.Label
    $s.Location = New-Object System.Drawing.Point($x,10); $s.Size = New-Object System.Drawing.Size(2,30)
    $s.BackColor = [System.Drawing.Color]::FromArgb(255,100,100,115)
    [void]$pnlBtns.Controls.Add($s); return $s
}
$sep1 = New-Sep 366

# ---- Cluster: Jira (test connection, sync issues into Excel) ----
$btnTest = New-Btn $pnlBtns 'Test'              376 8 70  36 0 122 200
$btnSync = New-Btn $pnlBtns 'Sync Jira -> Excel' 450 8 150 36 0 122 200

$sep2 = New-Sep 610

# ---- Cluster: D365 (push the timesheet into Dynamics) ----
$btnSubmit = New-Btn $pnlBtns 'Submit -> D365'     620 8 150 36 200 110 0
$btnFill   = New-Btn $pnlBtns 'Fill Times -> D365' 774 8 150 36 200 110 0
$btnFill.Enabled = $false
$btnFill.ForeColor = [System.Drawing.Color]::FromArgb(255,220,200,170)

# ---- Stop ----
$btnStop = New-Btn $pnlBtns 'Stop' 1024 8 22 36 140 30 30
$btnStop.Text = [char]9632  # stop square
$btnStop.Enabled = $false
```

(If a later `$btnFillStd.Add_Click` handler exists in the file it stays as-is and still binds to this `$btnFillStd`.)

- [ ] **Step 2: Verify parse + render**

Run: `powershell -NoProfile -Command "$e=$null;$null=[System.Management.Automation.Language.Parser]::ParseFile('jira-sync.ps1',[ref]$null,[ref]$e);if($e){$e|%{$_.Message}}else{'OK'}"`
Expected: `OK`

Launch `jira-sync.ps1`: the strip reads left-to-right as Excel (green: Read/Open/Standard) | Jira (blue: Test/Sync) | D365 (orange: Submit/Fill) | Stop (red), with separators between clusters and no overlap. Click each button and confirm it still triggers its action (Read File loads weeks, Standard → Excel fills, Sync runs, Test connects).

- [ ] **Step 3: Commit**

```bash
git add jira-sync.ps1
git commit -m "feat: group main-window buttons into Excel/Jira/D365 clusters"
```

---

### Task 2: Tidy the Weeks/Month strip and status/log

**Files:**
- Modify: `jira-sync.ps1` (the `$pnlWeeks` labels and the `$pnlStatus`/`$lblStatus`/`$txtLog` styling)

**Interfaces:**
- Consumes: `$pnlWeeks`, `$lblWeeksTitle`, `$lblMonth`, `$lnkSelectAll`, `$lnkNone`, `$pnlStatus`, `$lblStatus`, `$txtLog`.
- Produces: adjusted spacing/fonts only; no new controls with handlers.

- [ ] **Step 1: Give the Weeks strip a clearer left header**

Find `$lblWeeksTitle.Text = 'Weeks:'` and its font line; change the label to be bolder and add a subtle instruction. Replace the `$lblWeeksTitle.Font` line with:

```powershell
$lblWeeksTitle.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
```

- [ ] **Step 2: Pad the status bar text**

Find the `$lblStatus.Location = New-Object System.Drawing.Point(30,6)` line and change the X to align with a little more breathing room and widen it:

```powershell
$lblStatus.Location = New-Object System.Drawing.Point(34,6)
$lblStatus.Size = New-Object System.Drawing.Size(1000,18)
```

- [ ] **Step 3: Give the log a top margin and monospace clarity**

Find the `$txtLog.Font = New-Object System.Drawing.Font('Consolas',9)` line; leave the font, but immediately after `$txtLog.BorderStyle = 'None'` add a small left indent via padding-by-location. Change the `$txtLog.Location = New-Object System.Drawing.Point(0,167)` line to:

```powershell
$txtLog.Location = New-Object System.Drawing.Point(6,167)
```

and the width line `$txtLog.Size = New-Object System.Drawing.Size(1060,472)` to:

```powershell
$txtLog.Size = New-Object System.Drawing.Size(1048,472)
```

- [ ] **Step 4: Verify parse + render**

Parser check (same command as Task 1 Step 2) → `OK`.
Launch `jira-sync.ps1`: the Weeks header is bold, the status text isn't crammed against the dot, and the log has a small left margin. Nothing overlaps; resizing the window still lays out the log correctly.

- [ ] **Step 5: Commit**

```bash
git add jira-sync.ps1
git commit -m "polish: tidy weeks header, status padding, log margin"
```

---

## Notes for the implementer

- This phase is layout-only; the parser check catches syntax but not visual overlap — eyeball the window after each task.
- Do Phase 1 first: it restructures the Settings dialog and the config the app reads. Phase 2 only touches the main window and is independent, but running it second keeps the review surface small.
- If any coordinate change causes visual crowding on your display, nudge the X/width values — the exact pixels are a starting point, the clusters and order are the requirement.
